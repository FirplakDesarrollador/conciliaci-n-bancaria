import ExcelJS from "exceljs";
import { ACCOUNT_MAP, BOGOTA_CARD_DATE_TOLERANCE_DAYS, FILL_AMBIGUOUS, FILL_MATCHED, FILL_NO_MATCH } from "./config";
import { setComment, setFill } from "./excel-helpers";
import {
  applyCountMatching,
  applyGroupTotalMatch,
  applyReverseCombo,
  bestByTercero,
  findCandidates,
  tryMultiMatch,
  tryMultiMatchNoCode,
} from "./matching";
import { fixCompensacionDates, findHeaderRowAndCols, READERS, REQUIRED_HEADERS_HINT } from "./readers";
import { loadSapDocs } from "./sap";
import type { AccountStats, BankMove, SapDoc, SummaryRow } from "./types";
import { isBankFee, norm, normAccount } from "./utils";

export interface ReconcileInput {
  sapBuffer: Buffer;
  /** Buffer de cada archivo de banco descargado, indexado por la llave (cuenta SAP) de ACCOUNT_MAP. */
  bankBuffers: Map<string, Buffer>;
}

export interface AccountResult {
  cuentaKey: string;
  archivo: string;
  outputFileName: string;
  workbookBuffer: Buffer;
  stats: AccountStats;
}

export interface ReconcileOutput {
  results: AccountResult[];
  summaryRows: SummaryRow[];
  unusedDocs: SapDoc[];
  resumenBuffer: Buffer;
  totalSapDocs: number;
}

function isEmptyDoc(v: unknown): boolean {
  return v === null || v === undefined || v === "";
}

export async function reconcile({ sapBuffer, bankBuffers }: ReconcileInput): Promise<ReconcileOutput> {
  const pool = await loadSapDocs(sapBuffer);
  return reconcileDocs(pool, bankBuffers);
}

export async function reconcileDocs(pool: SapDoc[], bankBuffers: Map<string, Buffer>): Promise<ReconcileOutput> {
  if (pool.length === 0) {
    throw new Error("No hay documentos SAP para conciliar (revisa fechas/valores).");
  }

  const results: AccountResult[] = [];
  const summaryRows: SummaryRow[] = [];

  for (const [accountKey, { file: fname, format: fmt }] of Object.entries(ACCOUNT_MAP)) {
    const buffer = bankBuffers.get(accountKey);
    if (!buffer) continue; // no se encontró/descargó el archivo, se omite

    const wb = new ExcelJS.Workbook();
    // exceljs augmenta el tipo global `Buffer` de forma incompatible con la
    // versión actual de @types/node; el cast evita ese choque de tipos.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await wb.xlsx.load(buffer as any);

    if (fmt === "compensacion") {
      fixCompensacionDates(wb);
    }

    const reader = READERS[fmt];
    const hint = REQUIRED_HEADERS_HINT[fmt];

    const stats: AccountStats = {
      cuentaKey: accountKey,
      archivo: fname,
      matchExacto: 0,
      matchTolerancia: 0,
      matchValorUnico: 0,
      matchPar: 0,
      matchConteo: 0,
      matchInverso: 0,
      matchGrupo: 0,
      ambiguos: 0,
      sinDocumento: 0,
      comisionesIgnoradas: 0,
      yaTeniaDocumento: 0,
      fueraDeRango: 0,
    };

    for (const ws of wb.worksheets) {
      const found = findHeaderRowAndCols(ws, hint);
      if (!found) continue;
      const { headerRow, headers } = found;

      const moves = reader(ws, headers, headerRow);
      const eligibleGroup = moves.filter((mv) => isEmptyDoc(mv.docValue));
      const { resolvedRows, nAssigned: nGrp0 } = applyGroupTotalMatch(eligibleGroup, pool, ws, fname, ws.name, accountKey);
      stats.matchGrupo += nGrp0;
      const nomatchMoves: BankMove[] = [];

      for (const mv of moves) {
        if (resolvedRows.has(mv.row)) continue;
        if (!isEmptyDoc(mv.docValue)) {
          stats.yaTeniaDocumento++;
          continue;
        }
        if (isBankFee(mv.refText)) {
          stats.comisionesIgnoradas++;
          continue;
        }

        const cuentaForMatch = mv.cuentaOverride ?? normAccount(accountKey);
        let poolForMatch = pool;
        if (mv.terceroFilter) {
          const tf = norm(mv.terceroFilter);
          poolForMatch = pool.filter((d) => d.tercero.includes(tf));
        }

        // En Banco de Bogotá los pagos con tarjeta débito/crédito son
        // normales: SAP registra el recibo el día de la venta, pero el banco
        // solo refleja el depósito varios días después. Se usa una
        // tolerancia de fecha más amplia solo para esta cuenta.
        const toleranceDays =
          accountKey === "BANCO DE BOGOTA # 406007252" ? BOGOTA_CARD_DATE_TOLERANCE_DAYS : undefined;

        // Una suma de recibos EXACTAMENTE del mismo día es más confiable que
        // un documento individual de otro día, así que se intenta primero.
        const sameDayCombo = tryMultiMatch(poolForMatch, mv.tipo, cuentaForMatch, mv.value, mv.date, 0);
        const [candidates, how] = findCandidates(poolForMatch, mv.tipo, cuentaForMatch, mv.value, mv.date, toleranceDays);
        const cell = ws.getRow(mv.row).getCell(mv.docCol);

        if (sameDayCombo && !(candidates.length === 1 && how === "exacta")) {
          for (const d of sameDayCombo) {
            d.used = true;
            d.usedBy = `${fname}!${ws.name}!R${mv.row}`;
          }
          const nums = sameDayCombo.map((d) => d.docNum).join(", ");
          const suma = sameDayCombo.reduce((s, d) => s + d.value, 0);
          cell.value = nums;
          setFill(cell, FILL_MATCHED);
          setComment(
            cell,
            `Suma de ${sameDayCombo.length} recibos del mismo dia con la misma info detallada: ${nums} = ${suma.toFixed(2)}`
          );
          stats.matchPar++;
          continue;
        }

        if (candidates.length === 0) {
          const combo1 = tryMultiMatch(poolForMatch, mv.tipo, cuentaForMatch, mv.value, mv.date);
          if (combo1) {
            for (const d of combo1) {
              d.used = true;
              d.usedBy = `${fname}!${ws.name}!R${mv.row}`;
            }
            const nums = combo1.map((d) => d.docNum).join(", ");
            const suma = combo1.reduce((s, d) => s + d.value, 0);
            cell.value = nums;
            setFill(cell, FILL_MATCHED);
            setComment(cell, `Suma de ${combo1.length} recibos con la misma info detallada: ${nums} = ${suma.toFixed(2)}`);
            stats.matchPar++;
            continue;
          }

          const combo2 = tryMultiMatchNoCode(poolForMatch, mv.tipo, cuentaForMatch, mv.value, mv.date, 4, toleranceDays);
          if (combo2) {
            for (const d of combo2) {
              d.used = true;
              d.usedBy = `${fname}!${ws.name}!R${mv.row}`;
            }
            const nums = combo2.map((d) => d.docNum).join(", ");
            const suma = combo2.reduce((s, d) => s + d.value, 0);
            cell.value = nums;
            setFill(cell, FILL_MATCHED);
            setComment(
              cell,
              `Suma de ${combo2.length} documentos del mismo dia (sin codigo en comun) que cuadra con el valor: ${nums} = ${suma.toFixed(2)}. Revisar.`
            );
            stats.matchPar++;
            continue;
          }

          nomatchMoves.push(mv);
          continue;
        }

        if (candidates.length === 1) {
          const doc = candidates[0];
          doc.used = true;
          doc.usedBy = `${fname}!${ws.name}!R${mv.row}`;
          cell.value = doc.docNum;
          setFill(cell, FILL_MATCHED);
          if (how === "exacta") stats.matchExacto++;
          else if (how === "tolerancia") stats.matchTolerancia++;
          else stats.matchValorUnico++;
          continue;
        }

        const [bestDoc] = bestByTercero(candidates, mv.refText);
        if (bestDoc) {
          bestDoc.used = true;
          bestDoc.usedBy = `${fname}!${ws.name}!R${mv.row}`;
          cell.value = bestDoc.docNum;
          setFill(cell, FILL_MATCHED);
          if (how === "exacta") stats.matchExacto++;
          else stats.matchTolerancia++;
          continue;
        }

        // Candidatos empatados en valor y fecha, sin que el tercero los
        // diferencie: si TODOS corresponden al mismo tercero, son recibos
        // duplicados/interconmutables y cualquiera es una asignación válida.
        const tercerosUnicos = new Set(candidates.map((d) => d.tercero));
        if (how !== "ambiguo_lejano" && tercerosUnicos.size === 1) {
          const doc = [...candidates].sort((a, b) => a.docNum.localeCompare(b.docNum))[0];
          doc.used = true;
          doc.usedBy = `${fname}!${ws.name}!R${mv.row}`;
          cell.value = doc.docNum;
          setFill(cell, FILL_MATCHED);
          setComment(
            cell,
            `Varios documentos identicos en valor, fecha y tercero (${doc.tercero}): ${candidates
              .map((d) => d.docNum)
              .join(", ")}. Se asigno el #${doc.docNum}; cualquiera de ellos es una asignacion valida ya que son interconmutables.`
          );
          if (how === "exacta") stats.matchExacto++;
          else stats.matchTolerancia++;
          continue;
        }

        const candidatosTxt = candidates
          .slice(0, 6)
          .map((d) => `#${d.docNum} (${d.tercero})`)
          .join(", ");
        
        // Ya no pintamos la celda de amarillo
        stats.ambiguos++;
        summaryRows.push({
          cuentaSap: accountKey,
          hoja: ws.name,
          fila: mv.row,
          tipo: mv.tipo,
          fecha: mv.date,
          valor: mv.value,
          estado: "AMBIGUO",
          candidatos: candidatosTxt,
        });
      }

      // Segunda pasada: puede que lo que quedó sin documento ahora sí
      // cuadre en cantidad contra los documentos que sobraron en el pool.
      const { resolvedRows: resolvedRows2, nAssigned: nPost } = applyCountMatching(
        nomatchMoves,
        pool,
        ws,
        fname,
        ws.name,
        accountKey
      );
      stats.matchConteo += nPost;

      // Tercera pasada: un solo documento SAP dividido en varios movimientos.
      const remainingMoves = nomatchMoves.filter((mv) => !resolvedRows2.has(mv.row));
      const { resolvedRows: resolvedRows3, nAssigned: nRev } = applyReverseCombo(
        remainingMoves,
        pool,
        ws,
        fname,
        ws.name,
        accountKey
      );
      stats.matchInverso += nRev;

      // Cuarta pasada: categorías conocidas sin correspondencia 1 a 1.
      const remainingMoves2 = remainingMoves.filter((mv) => !resolvedRows3.has(mv.row));
      const { resolvedRows: resolvedRows4, nAssigned: nGrp } = applyGroupTotalMatch(
        remainingMoves2,
        pool,
        ws,
        fname,
        ws.name,
        accountKey
      );
      stats.matchGrupo += nGrp;

      const allResolved = new Set<number>([...resolvedRows2, ...resolvedRows3, ...resolvedRows4]);
      for (const mv of nomatchMoves) {
        if (allResolved.has(mv.row)) continue;
        
        stats.sinDocumento++;
        summaryRows.push({
          cuentaSap: accountKey,
          hoja: ws.name,
          fila: mv.row,
          tipo: mv.tipo,
          fecha: mv.date,
          valor: mv.value,
          estado: "SIN DOCUMENTO",
          candidatos: "",
        });
      }
    }

    const outputFileName = fname.replace(/\.xlsx$/i, "_CONCILIADO.xlsx");
    const outBuffer = Buffer.from(await wb.xlsx.writeBuffer());
    results.push({ cuentaKey: accountKey, archivo: fname, outputFileName, workbookBuffer: outBuffer, stats });
  }

  const unusedDocs = pool.filter((d) => !d.used);
  const resumenBuffer = await writeSummary(summaryRows, unusedDocs);

  return { results, summaryRows, unusedDocs, resumenBuffer, totalSapDocs: pool.length };
}

async function writeSummary(summaryRows: SummaryRow[], unused: SapDoc[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();

  const ws1 = wb.addWorksheet("Sin doc o ambiguos");
  ws1.addRow(["Cuenta SAP", "Hoja", "Fila", "Tipo", "Fecha", "Valor", "Estado", "Candidatos (si es ambiguo)"]);
  ws1.getRow(1).eachCell((c) => {
    c.font = { bold: true };
  });
  for (const row of summaryRows) {
    ws1.addRow([
      row.cuentaSap,
      row.hoja,
      row.fila,
      row.tipo,
      row.fecha.toISOString().slice(0, 10),
      row.valor,
      row.estado,
      row.candidatos,
    ]);
  }
  [32, 12, 6, 6, 12, 14, 14, 60].forEach((w, i) => {
    ws1.getColumn(i + 1).width = w;
  });

  const ws2 = wb.addWorksheet("Documentos SAP sin usar");
  ws2.addRow(["# Documento", "Fecha", "Valor", "Tercero", "Cuenta SAP", "Tipo"]);
  ws2.getRow(1).eachCell((c) => {
    c.font = { bold: true };
  });
  for (const d of unused) {
    ws2.addRow([d.docNum, d.date.toISOString().slice(0, 10), d.value, d.tercero, d.cuenta, d.tipo]);
  }
  [14, 12, 14, 40, 32, 6].forEach((w, i) => {
    ws2.getColumn(i + 1).width = w;
  });

  return Buffer.from(await wb.xlsx.writeBuffer());
}
