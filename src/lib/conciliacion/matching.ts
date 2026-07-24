import type ExcelJS from "exceljs";
import { DATE_TOLERANCE_DAYS, FILL_MATCHED, GROUP_TOTAL_CATEGORIES, MULTI_VALUE_TOLERANCE, VALUE_TOLERANCE } from "./config";
import { setComment, setFill } from "./excel-helpers";
import type { BankMove, MovTipo, SapDoc } from "./types";
import { combinations, daysBetween, infoCode, norm, normAccount, rangeInclusive, sequenceRatio } from "./utils";

export type MatchHow =
  | "exacta"
  | "tolerancia"
  | "valor_unico"
  | "mas_cercano"
  | "ambiguo_fecha"
  | "ambiguo_lejano";

/**
 * Busca TODOS los documentos que coincidan en cuenta + tipo + valor, sin
 * importar la fecha. Si hay un solo documento, ese es el match (aunque la
 * fecha esté lejos). Si hay varios, se prefiere el de fecha MÁS CERCANA,
 * siempre que sea único; si hay empate, se devuelve el grupo empatado para
 * que el desempate por tercero lo resuelva.
 */
export function findCandidates(
  pool: SapDoc[],
  tipo: MovTipo,
  cuenta: string | string[],
  valor: number,
  fecha: Date
): [SapDoc[], MatchHow | null] {
  const allowed = new Set(typeof cuenta === "string" ? [cuenta] : cuenta);
  const sameValue = pool.filter(
    (d) => !d.used && d.tipo === tipo && allowed.has(d.cuenta) && Math.abs(d.value - valor) <= VALUE_TOLERANCE
  );
  if (sameValue.length === 0) return [[], null];

  if (sameValue.length === 1) {
    const diff = Math.abs(daysBetween(sameValue[0].date, fecha));
    const how: MatchHow = diff === 0 ? "exacta" : diff <= DATE_TOLERANCE_DAYS ? "tolerancia" : "valor_unico";
    return [sameValue, how];
  }

  sameValue.sort((a, b) => Math.abs(daysBetween(a.date, fecha)) - Math.abs(daysBetween(b.date, fecha)));
  const bestDiff = Math.abs(daysBetween(sameValue[0].date, fecha));
  const secondDiff = Math.abs(daysBetween(sameValue[1].date, fecha));
  if (bestDiff < secondDiff) {
    const how: MatchHow = bestDiff === 0 ? "exacta" : bestDiff <= DATE_TOLERANCE_DAYS ? "tolerancia" : "mas_cercano";
    return [[sameValue[0]], how];
  }

  const tied = sameValue.filter((d) => Math.abs(daysBetween(d.date, fecha)) === bestDiff);
  const how: MatchHow = bestDiff <= DATE_TOLERANCE_DAYS ? "ambiguo_fecha" : "ambiguo_lejano";
  return [tied, how];
}

/**
 * Busca combinaciones de documentos SAP (mismo tipo, misma cuenta, mismo
 * código de 'Info detallada'/Proveedor) cuya suma sea igual al valor del
 * movimiento bancario.
 */
export function tryMultiMatch(
  pool: SapDoc[],
  tipo: MovTipo,
  cuenta: string | string[],
  valor: number,
  fecha: Date,
  dateWindow?: number,
  sizes?: number[],
  maxComboSize = 8
): SapDoc[] | null {
  const window = dateWindow !== undefined ? dateWindow : Math.max(DATE_TOLERANCE_DAYS, 10);
  const allowed = new Set(typeof cuenta === "string" ? [cuenta] : cuenta);
  const candidates = pool.filter(
    (d) =>
      !d.used &&
      d.tipo === tipo &&
      allowed.has(d.cuenta) &&
      d.infoDetallada &&
      Math.abs(daysBetween(d.date, fecha)) <= window
  );
  if (candidates.length < 2) return null;

  const groups = new Map<string, SapDoc[]>();
  for (const d of candidates) {
    const code = infoCode(d.infoDetallada);
    if (code) {
      if (!groups.has(code)) groups.set(code, []);
      groups.get(code)!.push(d);
    }
  }

  const found: SapDoc[][] = [];
  for (const docs of groups.values()) {
    if (docs.length < 2) continue;
    const effectiveMax = docs.length <= 10 ? maxComboSize : 3;
    const groupSizes = sizes ?? rangeInclusive(2, Math.min(docs.length, effectiveMax));
    for (const size of groupSizes) {
      if (docs.length < size) continue;
      for (const combo of combinations(docs, size)) {
        if (Math.abs(combo.reduce((s, x) => s + x.value, 0) - valor) <= MULTI_VALUE_TOLERANCE) {
          found.push(combo);
        }
      }
    }
  }

  if (found.length === 0) return null;
  const minSize = Math.min(...found.map((c) => c.length));
  const best = found.filter((c) => c.length === minSize);
  return best.length === 1 ? best[0] : null;
}

/**
 * Como tryMultiMatch, pero sin exigir un código de 'Info detallada' en
 * común: solo exige EXACTA la misma fecha, cuenta y tipo. Menos selectivo,
 * así que se limita a combinaciones pequeñas.
 */
export function tryMultiMatchNoCode(
  pool: SapDoc[],
  tipo: MovTipo,
  cuenta: string | string[],
  valor: number,
  fecha: Date,
  maxComboSize = 4
): SapDoc[] | null {
  const candidates = pool.filter(
    (d) => !d.used && d.tipo === tipo && d.cuenta === cuenta && Math.abs(daysBetween(d.date, fecha)) <= DATE_TOLERANCE_DAYS
  );
  if (candidates.length < 2) return null;

  const found: SapDoc[][] = [];
  for (const size of rangeInclusive(2, Math.min(candidates.length, maxComboSize))) {
    for (const combo of combinations(candidates, size)) {
      if (Math.abs(combo.reduce((s, x) => s + x.value, 0) - valor) <= MULTI_VALUE_TOLERANCE) {
        found.push(combo);
      }
    }
  }
  if (found.length === 0) return null;
  const minSize = Math.min(...found.map((c) => c.length));
  const best = found.filter((c) => c.length === minSize);
  return best.length === 1 ? best[0] : null;
}

export function bestByTercero(candidates: SapDoc[], refText: string): [SapDoc | null, number] {
  if (!refText) return [null, 0];
  const scored: Array<[number, SapDoc]> = [];
  for (const d of candidates) {
    if (!d.tercero) continue;
    let score = sequenceRatio(d.tercero, refText);
    if (d.tercero && refText.includes(d.tercero)) score += 0.3;
    scored.push([score, d]);
  }
  if (scored.length === 0) return [null, 0];
  scored.sort((a, b) => b[0] - a[0]);
  const [bestScore, bestDoc] = scored[0];
  const secondScore = scored.length > 1 ? scored[1][0] : 0;
  if (bestScore >= 0.35 && bestScore - secondScore >= 0.1) return [bestDoc, bestScore];
  return [null, bestScore];
}

/**
 * Agrupa movimientos por (tipo, cuenta, valor, filtro de tercero). Si la
 * cantidad de movimientos coincide (o casi) con la cantidad de documentos
 * SAP disponibles de ese mismo valor en la cuenta, los empareja por fecha
 * más cercana y asigna.
 */
export function applyCountMatching(
  mvsPool: BankMove[],
  pool: SapDoc[],
  ws: ExcelJS.Worksheet,
  fname: string,
  sheetName: string,
  accountKey: string,
  maxImbalance = 1
): { resolvedRows: Set<number>; nAssigned: number } {
  const groups = new Map<string, BankMove[]>();
  for (const mv of mvsPool) {
    const cuentaMv = mv.cuentaOverride ?? normAccount(accountKey);
    const key = JSON.stringify([mv.tipo, cuentaMv, Math.round(mv.value * 100) / 100, mv.terceroFilter ?? null]);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(mv);
  }

  const resolvedRows = new Set<number>();
  let nAssigned = 0;

  for (const mvs of groups.values()) {
    if (mvs.length < 2) continue;
    const first = mvs[0];
    const tipoK = first.tipo;
    const cuentaK = first.cuentaOverride ?? normAccount(accountKey);
    const valK = Math.round(first.value * 100) / 100;
    const tfK = first.terceroFilter;

    let cand = pool.filter(
      (d) => !d.used && d.tipo === tipoK && d.cuenta === cuentaK && Math.abs(d.value - valK) <= VALUE_TOLERANCE
    );
    if (tfK) {
      const tf = norm(tfK);
      cand = cand.filter((d) => d.tercero.includes(tf));
    }
    if (cand.length === 0) continue;
    const diff = Math.abs(cand.length - mvs.length);
    if (diff > maxImbalance) continue;

    const remainingMvs = [...mvs];
    const remainingDocs = [...cand];
    const pairs: Array<[BankMove, SapDoc]> = [];
    while (remainingMvs.length > 0 && remainingDocs.length > 0) {
      let best: [BankMove, SapDoc] | null = null;
      let bestDist: number | null = null;
      for (const mv of remainingMvs) {
        for (const d of remainingDocs) {
          const dist = Math.abs(daysBetween(d.date, mv.date));
          if (bestDist === null || dist < bestDist) {
            bestDist = dist;
            best = [mv, d];
          }
        }
      }
      pairs.push(best!);
      remainingMvs.splice(remainingMvs.indexOf(best![0]), 1);
      remainingDocs.splice(remainingDocs.indexOf(best![1]), 1);
    }

    for (const [mv, doc] of pairs) {
      doc.used = true;
      doc.usedBy = `${fname}!${sheetName}!R${mv.row}`;
      const cell = ws.getRow(mv.row).getCell(mv.docCol);
      cell.value = doc.docNum;
      setFill(cell, FILL_MATCHED);
      const notaDesbalance = diff
        ? ` ADVERTENCIA: la cantidad de movimientos (${mvs.length}) y de documentos (${cand.length}) no era igual; revisar el/los sobrante(s).`
        : "";
      setComment(
        cell,
        `Asignado por coincidencia de cantidad: ${mvs.length} movimientos y ${cand.length} documentos con el mismo valor (${valK.toFixed(2)}) en esta cuenta durante el mes; se emparejo por fecha mas cercana (doc ${doc.date.toISOString().slice(0, 10)} vs movimiento ${mv.date.toISOString().slice(0, 10)}).${notaDesbalance}`
      );
      resolvedRows.add(mv.row);
      nAssigned++;
    }
  }

  return { resolvedRows, nAssigned };
}

/**
 * Caso inverso a la suma de recibos: un solo documento SAP cuyo valor es la
 * SUMA de varios movimientos bancarios sin documento, con descripción
 * similar (mismas primeras palabras) y fecha cercana entre sí.
 */
export function applyReverseCombo(
  mvsPool: BankMove[],
  pool: SapDoc[],
  ws: ExcelJS.Worksheet,
  fname: string,
  sheetName: string,
  accountKey: string,
  sizes?: number[],
  dateWindow?: number,
  maxComboSize = 8
): { resolvedRows: Set<number>; nAssigned: number } {
  const window = dateWindow !== undefined ? dateWindow : DATE_TOLERANCE_DAYS;
  const resolvedRows = new Set<number>();
  let nAssigned = 0;

  const groups = new Map<string, { tipoK: MovTipo; cuentaK: string | string[]; prefix: string; mvs: BankMove[] }>();
  for (const mv of mvsPool) {
    const cuentaMv = mv.cuentaOverride ?? normAccount(accountKey);
    const prefix = mv.refText.split(" ").slice(0, 3).join(" ");
    if (!prefix) continue;
    const key = JSON.stringify([mv.tipo, cuentaMv, prefix]);
    if (!groups.has(key)) groups.set(key, { tipoK: mv.tipo, cuentaK: cuentaMv, prefix, mvs: [] });
    groups.get(key)!.mvs.push(mv);
  }

  for (const { tipoK, cuentaK, prefix, mvs } of groups.values()) {
    if (mvs.length < 2) continue;
    const docsCand = pool.filter((d) => !d.used && d.tipo === tipoK && d.cuenta === cuentaK);
    if (docsCand.length === 0) continue;

    const effectiveMax = mvs.length <= 10 ? maxComboSize : 3;
    const groupSizes = sizes ?? rangeInclusive(2, Math.min(mvs.length, effectiveMax));
    const found: Array<[SapDoc, BankMove[]]> = [];

    for (const size of groupSizes) {
      if (mvs.length < size) continue;
      for (const combo of combinations(mvs, size)) {
        const times = combo.map((m) => m.date.getTime());
        const maxT = Math.max(...times);
        const minT = Math.min(...times);
        if (Math.round((maxT - minT) / 86_400_000) > window) continue;
        const total = combo.reduce((s, m) => s + m.value, 0);
        for (const d of docsCand) {
          if (Math.abs(daysBetween(d.date, combo[0].date)) > window) continue;
          if (Math.abs(d.value - total) <= MULTI_VALUE_TOLERANCE) {
            found.push([d, combo]);
          }
        }
      }
    }

    const docCounts = new Map<SapDoc, number>();
    for (const [d] of found) docCounts.set(d, (docCounts.get(d) ?? 0) + 1);
    const candidatesF = found.filter(([d]) => docCounts.get(d) === 1);

    const rowOwner = new Map<number, BankMove[]>();
    const rowConflict = new Set<number>();
    for (const [, combo] of candidatesF) {
      for (const mvr of combo) {
        if (rowOwner.has(mvr.row) && rowOwner.get(mvr.row) !== combo) {
          rowConflict.add(mvr.row);
        }
        rowOwner.set(mvr.row, combo);
      }
    }

    const finalCombos = candidatesF.filter(([, combo]) => !combo.some((mvr) => rowConflict.has(mvr.row)));

    for (const [doc, combo] of finalCombos) {
      doc.used = true;
      const valoresTxt = combo.map((m) => m.value.toFixed(2)).join(", ");
      for (const mv of combo) {
        doc.usedBy = `${fname}!${sheetName}!R${mv.row}`;
        const cell = ws.getRow(mv.row).getCell(mv.docCol);
        cell.value = doc.docNum;
        setFill(cell, FILL_MATCHED);
        setComment(
          cell,
          `Documento SAP #${doc.docNum} (valor ${doc.value.toFixed(2)}) dividido en ${combo.length} movimientos bancarios con descripcion similar ('${prefix}'): ${valoresTxt}.`
        );
        resolvedRows.add(mv.row);
      }
      nAssigned += combo.length;
    }
  }

  return { resolvedRows, nAssigned };
}

/**
 * Para categorías conocidas (GROUP_TOTAL_CATEGORIES): agrupa documentos SAP
 * y movimientos bancarios sin documento por (cuenta, fecha); si el TOTAL de
 * ambos grupos coincide, asigna TODOS los números de documento a CADA
 * movimiento bancario del grupo (no hay correspondencia uno-a-uno clara).
 */
export function applyGroupTotalMatch(
  nomatchMoves: BankMove[],
  pool: SapDoc[],
  ws: ExcelJS.Worksheet,
  fname: string,
  sheetName: string,
  accountKey: string
): { resolvedRows: Set<number>; nAssigned: number } {
  const resolvedRows = new Set<number>();
  let nAssigned = 0;

  for (const [categoria, keywords] of Object.entries(GROUP_TOTAL_CATEGORIES)) {
    const docsByKey = new Map<string, SapDoc[]>();
    for (const d of pool) {
      if (!d.used && d.categoria === categoria) {
        const key = JSON.stringify([d.cuenta, d.date.getTime()]);
        if (!docsByKey.has(key)) docsByKey.set(key, []);
        docsByKey.get(key)!.push(d);
      }
    }

    const movesByKey = new Map<string, BankMove[]>();
    for (const mv of nomatchMoves) {
      const cuentaMv = mv.cuentaOverride ?? normAccount(accountKey);
      const ref = mv.refText;
      if (keywords.some((kw) => ref.includes(kw))) {
        const key = JSON.stringify([cuentaMv, mv.date.getTime()]);
        if (!movesByKey.has(key)) movesByKey.set(key, []);
        movesByKey.get(key)!.push(mv);
      }
    }

    for (const [key, docs] of docsByKey.entries()) {
      const mvs = movesByKey.get(key);
      if (!mvs || mvs.length === 0) continue;
      const docTotal = docs.reduce((s, d) => s + d.value, 0);
      const mvTotal = mvs.reduce((s, m) => s + m.value, 0);
      if (Math.abs(docTotal - mvTotal) > MULTI_VALUE_TOLERANCE) continue;

      const docNums = docs.map((d) => d.docNum).join(", ");
      for (const d of docs) d.used = true;
      for (const mv of mvs) {
        docs[0].usedBy = `${fname}!${sheetName}!R${mv.row}`;
        const cell = ws.getRow(mv.row).getCell(mv.docCol);
        cell.value = docNums;
        setFill(cell, FILL_MATCHED);
        setComment(
          cell,
          `Categoria '${categoria}': ${docs.length} documentos SAP (${docNums}, total ${docTotal.toFixed(2)}) financian en conjunto ${mvs.length} movimientos bancarios de este mismo dia (total ${mvTotal.toFixed(2)}). No hay correspondencia 1 a 1 clara, se listan todos los documentos en cada movimiento.`
        );
        resolvedRows.add(mv.row);
        nAssigned++;
      }
    }
  }

  return { resolvedRows, nAssigned };
}
