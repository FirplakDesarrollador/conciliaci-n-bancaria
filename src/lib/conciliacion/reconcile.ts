import ExcelJS from "exceljs";
import { ACCOUNT_MAP, BOGOTA_CARD_DATE_TOLERANCE_DAYS, FILL_AMBIGUOUS, FILL_MATCHED, FILL_NO_MATCH, VALUE_TOLERANCE } from "./config";
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
import { isBankFee, norm, normAccount, sameMonth } from "./utils";

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

// Una celda con documento ya puede traer un solo # o una combinación
// ("81073, 81074", "81216-81217"); se extraen todos los números que
// contenga para poder marcar esos documentos como usados.
function extractDocNums(v: unknown): string[] {
  if (isEmptyDoc(v)) return [];
  return String(v).match(/\d+/g) ?? [];
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

  const poolByDocNum = new Map<string, SapDoc[]>();
  // Indice por valor redondeado, para calcular rapido (sin recorrer todo el
  // pool por cada movimiento) que tan cerca esta el mejor documento
  // candidato de cada movimiento bancario.
  const poolByValue = new Map<number, SapDoc[]>();
  for (const d of pool) {
    if (!poolByDocNum.has(d.docNum)) poolByDocNum.set(d.docNum, []);
    poolByDocNum.get(d.docNum)!.push(d);
    const k = Math.round(d.value);
    if (!poolByValue.has(k)) poolByValue.set(k, []);
    poolByValue.get(k)!.push(d);
  }

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

    // Pre-pasada: leer TODAS las hojas y reservar, antes de cruzar nada, los
    // documentos que ya estan escritos en alguna fila de este archivo. Si la
    // reserva se hiciera fila por fila durante el cruce, una fila vacia que
    // se procesa ANTES de otra fila ya conciliada (mismo valor, dias
    // distintos) tomaria el documento todavia "libre" y quedaria relacionado
    // dos veces. La reserva se limita a las patas del documento que
    // pertenecen a ESTA cuenta: un documento entre dos cuentas propias (ej.
    // venta de dolares Miami -> Bancolombia) tiene una pata por archivo, y
    // reservar la del otro banco lo dejaria sin conciliar alla.
    const sheetsData: { ws: ExcelJS.Worksheet; moves: BankMove[] }[] = [];
    for (const ws of wb.worksheets) {
      const found = findHeaderRowAndCols(ws, hint);
      if (!found) continue;
      const moves = reader(ws, found.headers, found.headerRow);
      sheetsData.push({ ws, moves });

      for (const mv of moves) {
        if (isEmptyDoc(mv.docValue)) continue;
        const allowed = new Set(
          [mv.cuentaOverride ?? normAccount(accountKey)].flat()
        );
        for (const num of extractDocNums(mv.docValue)) {
          const legs = poolByDocNum.get(num);
          if (!legs) continue;
          const legsHere = legs.filter((d) => allowed.has(d.cuenta));
          for (const d of legsHere.length > 0 ? legsHere : legs) {
            if (!d.used) {
              d.used = true;
              d.usedBy = `${fname}!${ws.name}!R${mv.row}`;
            }
          }
        }
      }
    }

    for (const { ws, moves } of sheetsData) {
      const eligibleGroup = moves.filter((mv) => isEmptyDoc(mv.docValue));
      const { resolvedRows, nAssigned: nGrp0 } = applyGroupTotalMatch(eligibleGroup, pool, ws, fname, ws.name, accountKey);
      stats.matchGrupo += nGrp0;
      const nomatchMoves: BankMove[] = [];

      // Se procesa primero el movimiento cuyo mejor documento candidato esta
      // mas cerca en fecha (no por orden de fila). Si dos movimientos tienen
      // el mismo valor y solo hay un documento, se lo queda el de mejor
      // coincidencia en vez del primero que aparezca en la hoja.
      const nearestCandidateDays = (mv: BankMove): number => {
        const allowed = new Set([mv.cuentaOverride ?? normAccount(accountKey)].flat());
        const tf = mv.terceroFilter ? norm(mv.terceroFilter) : null;
        const k = Math.round(mv.value);
        let best = Infinity;
        for (let dk = -1; dk <= 1; dk++) {
          for (const d of poolByValue.get(k + dk) ?? []) {
            if (d.used || d.tipo !== mv.tipo || !allowed.has(d.cuenta)) continue;
            if (Math.abs(d.value - mv.value) > VALUE_TOLERANCE) continue;
            if (!sameMonth(d.date, mv.date)) continue;
            if (tf && !d.tercero.includes(tf)) continue;
            best = Math.min(best, Math.abs(d.date.getTime() - mv.date.getTime()) / 86_400_000);
          }
        }
        return best;
      };
      const orderedMoves = moves
        .map((mv) => ({
          mv,
          dist:
            !resolvedRows.has(mv.row) && isEmptyDoc(mv.docValue) && !isBankFee(mv.refText)
              ? nearestCandidateDays(mv)
              : Infinity,
        }))
        .sort((a, b) => (a.dist === b.dist ? 0 : a.dist < b.dist ? -1 : 1))
        .map((x) => x.mv);

      // Pasada previa: una suma EXACTA (al centavo) de documentos del mismo
      // tercero es una evidencia mucho mas fuerte que un documento individual
      // con diferencia de centavos. Sin esta prioridad, el documento
      // individual se asigna primero a otra fila y el combo correcto queda
      // sin uno de sus miembros (82064: fila de $120.000 vs. combo
      // 82064+82066 = $452.200,00 exactos). Solo aplica si NO existe un
      // documento individual con valor exacto para esa fila.
      for (const mv of orderedMoves) {
        if (resolvedRows.has(mv.row) || !isEmptyDoc(mv.docValue) || isBankFee(mv.refText)) continue;
        const cuentaEx = mv.cuentaOverride ?? normAccount(accountKey);
        const allowedEx = new Set([cuentaEx].flat());
        const tfEx = mv.terceroFilter ? norm(mv.terceroFilter) : null;
        const poolEx = tfEx ? pool.filter((d) => d.tercero.includes(tfEx)) : pool;
        const hasExactSingle = poolEx.some(
          (d) =>
            !d.used &&
            d.tipo === mv.tipo &&
            allowedEx.has(d.cuenta) &&
            Math.abs(d.value - mv.value) < 0.005 &&
            sameMonth(d.date, mv.date)
        );
        if (hasExactSingle) continue;
        // Se agrupan los documentos disponibles por tercero y se busca la
        // suma exacta DENTRO de cada tercero (el codigo comun que usa
        // tryMultiMatch no distingue clientes en pagos recibidos).
        const byTercero = new Map<string, SapDoc[]>();
        for (const d of poolEx) {
          if (d.used || d.tipo !== mv.tipo || !allowedEx.has(d.cuenta) || !d.tercero) continue;
          if (!sameMonth(d.date, mv.date)) continue;
          if (Math.abs(d.date.getTime() - mv.date.getTime()) > 10 * 86_400_000) continue;
          if (!byTercero.has(d.tercero)) byTercero.set(d.tercero, []);
          byTercero.get(d.tercero)!.push(d);
        }
        const exactCombos: SapDoc[][] = [];
        for (const docsT of byTercero.values()) {
          if (docsT.length < 2) continue;
          if (docsT.reduce((sum, d) => sum + d.value, 0) < mv.value - 0.005) continue;
          const c = tryMultiMatch(docsT, mv.tipo, cuentaEx, mv.value, mv.date, undefined, undefined, 8, 0.005);
          if (c) exactCombos.push(c);
        }
        if (exactCombos.length !== 1) continue;
        const exactCombo = exactCombos[0];
        for (const d of exactCombo) {
          d.used = true;
          d.usedBy = `${fname}!${ws.name}!R${mv.row}`;
        }
        const nums = exactCombo.map((d) => d.docNum).join(", ");
        const cellEx = ws.getRow(mv.row).getCell(mv.docCol);
        cellEx.value = nums;
        setFill(cellEx, FILL_MATCHED);
        setComment(
          cellEx,
          `Suma exacta de ${exactCombo.length} documentos del mismo tercero: ${nums} = ${exactCombo.reduce((s, d) => s + d.value, 0).toFixed(2)}`
        );
        stats.matchPar++;
        resolvedRows.add(mv.row);
      }

      for (const mv of orderedMoves) {
        if (resolvedRows.has(mv.row)) continue;
        if (!isEmptyDoc(mv.docValue)) {
          // Los documentos de esta fila ya fueron reservados en la
          // pre-pasada de arriba.
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
