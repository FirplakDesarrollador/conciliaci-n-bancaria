import type ExcelJS from "exceljs";
import type { BankFormat } from "./config";
import type { BankMove, SpecialRule } from "./types";
import { cellValue, norm, parseDate, parseMoney } from "./utils";

// exceljs a veces reporta la dimensión nominal de la hoja (hasta 16384) en
// vez del rango realmente usado; se limita el barrido de columnas a un
// margen generoso (los archivos reales usan como mucho ~13 columnas).
const MAX_SCAN_COLUMNS = 60;

function usedColumnCount(ws: ExcelJS.Worksheet): number {
  return Math.min(ws.columnCount, MAX_SCAN_COLUMNS);
}

const DOC_COL_NAMES = ["RC", "DOC SAP", "DOCUMENTO SAP"];
const DATE_COL_NAMES = ["FECHA", "FECHA INGRESO"];

function getDocCol(headers: Map<string, number>, order: string[] = DOC_COL_NAMES): number | undefined {
  for (const name of order) {
    const c = headers.get(name);
    if (c !== undefined) return c;
  }
  return undefined;
}

function getFechaCol(headers: Map<string, number>): number | undefined {
  for (const name of DATE_COL_NAMES) {
    const c = headers.get(name);
    if (c !== undefined) return c;
  }
  return undefined;
}

export function findHeaderRowAndCols(
  ws: ExcelJS.Worksheet,
  required: readonly string[]
): { headerRow: number; headers: Map<string, number> } | null {
  const lastRow = Math.min(ws.rowCount, 5);
  for (let r = 1; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const headers = new Map<string, number>();
    for (let c = 1; c <= usedColumnCount(ws); c++) {
      const v = norm(row.getCell(c).value);
      if (v) headers.set(v, c);
    }
    if (getFechaCol(headers) !== undefined && required.some((h) => headers.has(h))) {
      return { headerRow: r, headers };
    }
  }
  return null;
}

function refPartsFrom(row: ExcelJS.Row, headers: Map<string, number>, names: string[]): string[] {
  const parts: string[] = [];
  for (const name of names) {
    const c = headers.get(name);
    if (c === undefined) continue;
    const v = cellValue(row.getCell(c).value);
    if (v) parts.push(String(v));
  }
  return parts;
}

/**
 * Para bancos con columnas Débito/Crédito separadas (Bogotá, Compensación).
 *
 * specialRules: reglas para movimientos que, por su descripción, en realidad
 * deben cruzarse contra un tipo de documento y/o cuenta SAP distintos a los
 * que indicaría el signo débito/crédito (p. ej. traslados entre la cuenta
 * corriente y la fiducia, que el banco registra como cargo pero SAP como
 * "Pago recibido" bajo la cuenta FIDECOMISO).
 */
function readGenericTwoColumn(
  ws: ExcelJS.Worksheet,
  headers: Map<string, number>,
  headerRow: number,
  debitH: string,
  creditH: string,
  specialRules: SpecialRule[] = []
): BankMove[] {
  const moves: BankMove[] = [];
  const docCol = getDocCol(headers);
  if (docCol === undefined) return moves;
  const fechaCol = getFechaCol(headers)!;

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fecha = parseDate(row.getCell(fechaCol).value);
    if (!fecha) continue;

    const debitoCol = headers.get(debitH);
    const creditoCol = headers.get(creditH);
    const debito = debitoCol !== undefined ? parseMoney(row.getCell(debitoCol).value) : null;
    const credito = creditoCol !== undefined ? parseMoney(row.getCell(creditoCol).value) : null;
    const docValue = cellValue(row.getCell(docCol).value);
    const refText = norm(refPartsFrom(row, headers, ["DESCRIPCIÓN", "TRANSACCION", "OBSERVACIONES", "NOMBRE"]).join(" "));

    const matchedRule = specialRules.find((rule) => rule.pattern.test(refText));
    if (matchedRule) {
      const valor = credito || debito;
      if (valor) {
        moves.push({
          sheet: ws.name,
          row: r,
          date: fecha,
          value: Math.abs(valor),
          tipo: matchedRule.tipo,
          docCol,
          docValue,
          refText,
          cuentaOverride: matchedRule.cuenta,
          terceroFilter: matchedRule.tercero,
        });
      }
      continue;
    }

    if (credito) {
      moves.push({ sheet: ws.name, row: r, date: fecha, value: Math.abs(credito), tipo: "IN", docCol, docValue, refText });
    }
    if (debito) {
      moves.push({ sheet: ws.name, row: r, date: fecha, value: Math.abs(debito), tipo: "OUT", docCol, docValue, refText });
    }
  }

  return moves;
}

function readSudameris(ws: ExcelJS.Worksheet, headers: Map<string, number>, headerRow: number): BankMove[] {
  const moves: BankMove[] = [];
  const docCol = getDocCol(headers);
  if (docCol === undefined) return moves;
  const fechaCol = getFechaCol(headers)!;

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fecha = parseDate(row.getCell(fechaCol).value);
    if (!fecha) continue;

    const dcCol = headers.get("D/C");
    const dc = dcCol !== undefined ? norm(row.getCell(dcCol).value) : "";
    const valorCol = headers.get("VALOR");
    const valor = valorCol !== undefined ? parseMoney(row.getCell(valorCol).value) : null;
    if (!valor) continue;

    const docValue = cellValue(row.getCell(docCol).value);
    const refParts = refPartsFrom(row, headers, ["DESCRIPCIÓN", "OFICINA"]);
    const extraCol = docCol + 1;
    if (extraCol <= usedColumnCount(ws)) {
      const extraV = cellValue(row.getCell(extraCol).value);
      if (extraV) refParts.push(String(extraV));
    }
    const refText = norm(refParts.join(" "));

    const tipo = dc === "C" ? "IN" : dc === "D" ? "OUT" : null;
    if (tipo) {
      moves.push({ sheet: ws.name, row: r, date: fecha, value: Math.abs(valor), tipo, docCol, docValue, refText });
    }
  }

  return moves;
}

function readDavivienda(ws: ExcelJS.Worksheet, headers: Map<string, number>, headerRow: number): BankMove[] {
  const moves: BankMove[] = [];
  const docCol = getDocCol(headers);
  if (docCol === undefined) return moves;
  const fechaCol = getFechaCol(headers)!;

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fecha = parseDate(row.getCell(fechaCol).value);
    if (!fecha) continue;

    const transacCol = headers.get("TRANSACCIÓN");
    const transac = transacCol !== undefined ? norm(row.getCell(transacCol).value) : "";
    const valorCol = headers.get("VALOR TOTAL");
    const valor = valorCol !== undefined ? parseMoney(row.getCell(valorCol).value) : null;
    if (!valor) continue;

    const docValue = cellValue(row.getCell(docCol).value);
    const refText = norm(refPartsFrom(row, headers, ["DESCRIPCIÓN", "OFICINA"]).join(" "));

    let tipo: "IN" | "OUT" | null = null;
    if (transac.includes("CRÉDITO")) tipo = "IN";
    else if (transac.includes("DÉBITO")) tipo = "OUT";
    else continue;

    moves.push({ sheet: ws.name, row: r, date: fecha, value: Math.abs(valor), tipo, docCol, docValue, refText });
  }

  return moves;
}

function readBancolombia(
  ws: ExcelJS.Worksheet,
  headers: Map<string, number>,
  headerRow: number,
  specialRules: SpecialRule[] = []
): BankMove[] {
  const moves: BankMove[] = [];
  const docCol = getDocCol(headers, ["DOCUMENTO SAP", "DOC SAP", "RC"]);
  if (docCol === undefined) return moves;
  const fechaCol = getFechaCol(headers)!;
  const valorCol = headers.get("VALOR");

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fecha = parseDate(row.getCell(fechaCol).value);
    if (!fecha) continue;

    const valor = valorCol !== undefined ? parseMoney(row.getCell(valorCol).value) : null;
    if (!valor) continue;

    const docValue = cellValue(row.getCell(docCol).value);
    const refText = norm(refPartsFrom(row, headers, ["DETALLE", "OBSERVACIONES"]).join(" "));

    const matchedRule = specialRules.find((rule) => rule.pattern.test(refText));
    if (matchedRule) {
      moves.push({
        sheet: ws.name,
        row: r,
        date: fecha,
        value: Math.abs(valor),
        tipo: matchedRule.tipo,
        docCol,
        docValue,
        refText,
        cuentaOverride: matchedRule.cuenta,
        terceroFilter: matchedRule.tercero,
      });
      continue;
    }

    const tipo = valor > 0 ? "IN" : "OUT";
    moves.push({ sheet: ws.name, row: r, date: fecha, value: Math.abs(valor), tipo, docCol, docValue, refText });
  }

  return moves;
}

// Traslados hacia el fondo de inversión (fiducia): Bancolombia los registra
// como un cargo (salida) con la descripción "TRASLADO A FONDO DE INVERSION",
// pero al ser un movimiento interno entre la cuenta y la fiducia, SAP lo
// registra como un "Pago recibido" (IN) bajo la cuenta FIDECOMISO, con
// tercero "FONDO DE INVERSION COLECTIVA ABIERTO FIDUCUENTA". Los traslados
// de regreso ("TRASLADO DE FONDO DE INVERS...") ya llegan como créditos
// normales y SAP los registra directamente en la cuenta de Bancolombia, así
// que no necesitan una regla especial.
export const BANCOLOMBIA_SPECIAL_RULES: SpecialRule[] = [
  {
    pattern: /TRASLADO A FONDO DE INVERSION/,
    tipo: "IN",
    cuenta: ["BANCOL.CTE # 008-927404-01", "FIDECOMISO"],
    tercero: "FONDO DE INVERSION COLECTIVA ABIERTO FIDUCUENTA",
  },
];

// Traslados entre la cuenta corriente y la fiducia: el banco de Bogotá los
// registra como cargo (débito) con la descripción "DB Inversion No ...". SAP
// a veces registra el documento bajo la cuenta FIDECOMISO y a veces
// directamente bajo la cuenta del banco; se buscan ambas y se usa la fecha
// más cercana para decidir. Se restringe al tercero "FIDUCIARIA BOGOTA S.A",
// que es el que corresponde a estos traslados de tesorería; "FONDO DE
// INVERSION COLECTIVA ABIERTO FIDUCUENTA" es un flujo distinto (compra/
// redención de unidades del fondo) que no debe cruzarse aquí.
export const BOGOTA_SPECIAL_RULES: SpecialRule[] = [
  {
    pattern: /DB INVERSION/,
    tipo: "IN",
    cuenta: ["BANCO DE BOGOTA # 406007252", "FIDECOMISO"],
    tercero: "FIDUCIARIA BOGOTA S.A",
  },
];

type Reader = (ws: ExcelJS.Worksheet, headers: Map<string, number>, headerRow: number) => BankMove[];

export const READERS: Record<BankFormat, Reader> = {
  bogota: (ws, h, hr) => readGenericTwoColumn(ws, h, hr, "DEBITO", "CREDITO", BOGOTA_SPECIAL_RULES),
  compensacion: (ws, h, hr) => readGenericTwoColumn(ws, h, hr, "DÉBITO", "CRÉDITO"),
  sudameris: readSudameris,
  davivienda: readDavivienda,
  bancolombia: (ws, h, hr) => readBancolombia(ws, h, hr, BANCOLOMBIA_SPECIAL_RULES),
};

export const REQUIRED_HEADERS_HINT: Record<BankFormat, readonly string[]> = {
  bogota: ["DEBITO", "CREDITO"],
  compensacion: ["DÉBITO", "CRÉDITO"],
  sudameris: ["D/C", "VALOR"],
  davivienda: ["VALOR TOTAL"],
  bancolombia: ["VALOR"],
};

/**
 * La cuenta de compensación en Miami trae un bloque de fechas exportadas
 * como objetos de fecha con día y mes invertidos (ej. aparece "6 de mayo"
 * cuando en realidad es "5 de junio"). Se corrige intercambiando mes y día
 * solo quando el intercambio produce una fecha válida (día original <= 12),
 * igual que hacía el script original con openpyxl.
 */
export function fixCompensacionDates(wb: ExcelJS.Workbook): number {
  let nFixed = 0;
  wb.eachSheet((ws) => {
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 2; c++) {
        const cell = row.getCell(c);
        const v = cell.value;
        if (v instanceof Date) {
          const year = v.getUTCFullYear();
          const month = v.getUTCMonth() + 1;
          const day = v.getUTCDate();
          if (day <= 12) {
            const nuevo = new Date(Date.UTC(year, day - 1, month));
            if (nuevo.getTime() !== v.getTime()) {
              cell.value = nuevo;
              nFixed++;
            }
          }
        }
      }
    }
  });
  return nFixed;
}
