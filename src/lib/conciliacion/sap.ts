import ExcelJS from "exceljs";
import { MANUAL_CUENTA_OVERRIDES, MIAMI_ACCOUNT } from "./config";
import type { SapDoc } from "./types";
import { cellValue, norm, normAccount, parseDate, parseMoney } from "./utils";

function buildHeaderCols(ws: ExcelJS.Worksheet): Map<string, number> {
  const col = new Map<string, number>();
  const headerRow = ws.getRow(1);
  for (let c = 1; c <= ws.columnCount; c++) {
    const h = norm(headerRow.getCell(c).value);
    if (h) col.set(h, c);
  }
  return col;
}

function requireCol(col: Map<string, number>, name: string, sheetName: string): number {
  const c = col.get(name);
  if (c === undefined) {
    throw new Error(`No se encontró la columna "${name}" en la hoja "${sheetName}" del informe SAP.`);
  }
  return c;
}

export async function loadSapDocs(buffer: Buffer): Promise<SapDoc[]> {
  const wb = new ExcelJS.Workbook();
  // exceljs augmenta el tipo global `Buffer` de forma incompatible con la
  // versión actual de @types/node; el cast evita ese choque de tipos.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await wb.xlsx.load(buffer as any);
  const docs: SapDoc[] = [];

  // --- Pagos recibidos (ingresos) ---
  {
    const sheetName = "Pagos recibidos";
    const ws = wb.getWorksheet(sheetName);
    if (!ws) throw new Error(`No se encontró la hoja "${sheetName}" en el informe SAP.`);
    const col = buildHeaderCols(ws);

    const cNum = requireCol(col, "#", sheetName);
    const cFecha = requireCol(col, "FECHA", sheetName);
    const cTotal = requireCol(col, "TOTAL DEL DOCUMENTO", sheetName);
    const cComentarios = requireCol(col, "COMENTARIOS", sheetName);
    const cCliente = requireCol(col, "CLIENTE", sheetName);
    const cBanco = requireCol(col, "BANCO", sheetName);
    const cTercero = col.get("TERCERO");
    const cInfoDet = col.get("INFO.DETALLADA");
    const cValorUsd = col.get("VALOR EN USD");

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const docNum = cellValue(row.getCell(cNum).value);
      if (docNum === null || docNum === undefined) continue;

      const fecha = parseDate(row.getCell(cFecha).value);
      let valor = parseMoney(row.getCell(cTotal).value);
      const comentario = cellValue(row.getCell(cComentarios).value);
      if (comentario && norm(comentario) === "CANCELADO") continue;

      const cliente = cellValue(row.getCell(cCliente).value);
      const tercero2 = cTercero !== undefined ? cellValue(row.getCell(cTercero).value) : null;
      let cuenta = cellValue(row.getCell(cBanco).value);
      const override = MANUAL_CUENTA_OVERRIDES[String(docNum)];
      if (override) cuenta = override;
      const infoDet = cInfoDet !== undefined ? cellValue(row.getCell(cInfoDet).value) : null;

      if (normAccount(cuenta) === normAccount(MIAMI_ACCOUNT) && cValorUsd !== undefined) {
        const valorUsd = parseMoney(row.getCell(cValorUsd).value);
        if (valorUsd) valor = valorUsd;
      }

      if (fecha === null || valor === null) continue;

      docs.push({
        docNum: String(docNum),
        date: fecha,
        value: valor,
        tercero: norm((tercero2 as string) || (cliente as string)),
        cuenta: normAccount(cuenta),
        comentario: String(comentario ?? ""),
        tipo: "IN",
        infoDetallada: norm(infoDet),
        categoria: "",
        used: false,
        usedBy: "",
      });
    }
  }

  // --- Pagos efectuados (egresos) ---
  {
    const sheetName = "Pagos efectuados";
    const ws = wb.getWorksheet(sheetName);
    if (!ws) throw new Error(`No se encontró la hoja "${sheetName}" en el informe SAP.`);
    const col = buildHeaderCols(ws);

    const cNum = requireCol(col, "#", sheetName);
    const cFecha = requireCol(col, "FECHA", sheetName);
    const cImporte = requireCol(col, "IMPORTE DE TRANSFERENCIA", sheetName);
    const cComentarios = requireCol(col, "COMENTARIOS", sheetName);
    const cAcreedor = requireCol(col, "NOMBRE DE DEUDOR/ACREEDOR", sheetName);
    const cCuentaTransf = requireCol(col, "CUENTA DE TRANSFERENCIA", sheetName);
    const cProveedor = col.get("PROVEEDOR");
    const cInfoDet = col.get("INFO.DETALLADA");
    const cValorUsd = col.get("VALOR EN USD");

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const docNum = cellValue(row.getCell(cNum).value);
      if (docNum === null || docNum === undefined) continue;

      const fecha = parseDate(row.getCell(cFecha).value);
      let valor = parseMoney(row.getCell(cImporte).value);
      const comentario = cellValue(row.getCell(cComentarios).value);
      if (comentario && norm(comentario) === "CANCELADO") continue;

      const acreedor = cellValue(row.getCell(cAcreedor).value);
      const proveedor = cProveedor !== undefined ? cellValue(row.getCell(cProveedor).value) : null;
      let cuenta = cellValue(row.getCell(cCuentaTransf).value);
      const override = MANUAL_CUENTA_OVERRIDES[String(docNum)];
      if (override) cuenta = override;
      const infoDet = cInfoDet !== undefined ? cellValue(row.getCell(cInfoDet).value) : null;

      if (normAccount(cuenta) === normAccount(MIAMI_ACCOUNT) && cValorUsd !== undefined) {
        const valorUsd = parseMoney(row.getCell(cValorUsd).value);
        if (valorUsd) valor = valorUsd;
      }

      if (fecha === null || valor === null) continue;

      docs.push({
        docNum: String(docNum),
        date: fecha,
        value: valor,
        tercero: norm((acreedor as string) || (proveedor as string)),
        cuenta: normAccount(cuenta),
        comentario: String(comentario ?? ""),
        tipo: "OUT",
        // Para Pagos efectuados, el código que identifica documentos
        // "hermanos" (ej. AC70076116-01 / -02) vive en 'Proveedor', no en
        // 'Info.detallada' (que suele venir vacío).
        infoDetallada: norm(proveedor),
        // 'Info.detallada' sí trae una categoría útil en otros casos, como
        // 'CAJA MENOR', que agrupa documentos sin código común entre sí
        // pero que en conjunto financian los mismos retiros.
        categoria: norm(infoDet),
        used: false,
        usedBy: "",
      });
    }
  }

  return docs;
}

export function convertApiToSapDocs(incomingPayments: any[], vendorPayments: any[]): SapDoc[] {
  const docs: SapDoc[] = [];

  for (const p of incomingPayments) {
    const valor = p.TransferSum || p.CashSum || 0;
    if (!p.DocNum || !p.DocDate || valor === 0) continue;
    
    if (p.JournalRemarks && norm(p.JournalRemarks) === "CANCELADO") continue;
    if (p.Remarks && norm(p.Remarks) === "CANCELADO") continue;

    let cuenta = p.Reference1 || p.CardCode;
    const override = MANUAL_CUENTA_OVERRIDES[String(p.DocNum)];
    if (override) cuenta = override;

    docs.push({
      docNum: String(p.DocNum),
      date: new Date(p.DocDate),
      value: valor,
      tercero: norm(p.U_DescripTercero || p.CardName),
      cuenta: normAccount(cuenta),
      comentario: String(p.JournalRemarks || p.Remarks || ""),
      tipo: "IN",
      infoDetallada: norm(p.JournalRemarks || p.Remarks),
      categoria: "",
      used: false,
      usedBy: "",
    });
  }

  for (const p of vendorPayments) {
    const valor = p.TransferSum || p.CashSum || 0;
    if (!p.DocNum || !p.DocDate || valor === 0) continue;
    
    if (p.JournalRemarks && norm(p.JournalRemarks) === "CANCELADO") continue;
    if (p.Remarks && norm(p.Remarks) === "CANCELADO") continue;

    let cuenta = p.TransferAccount;
    const override = MANUAL_CUENTA_OVERRIDES[String(p.DocNum)];
    if (override) cuenta = override;


    docs.push({
      docNum: String(p.DocNum),
      date: new Date(p.DocDate),
      value: valor,
      tercero: norm(p.U_DescripTercero || p.CardName || p.CardCode),
      cuenta: normAccount(cuenta),
      comentario: String(p.JournalRemarks || p.Remarks || ""),
      tipo: "OUT",
      infoDetallada: norm(p.U_CodTecero || p.CardCode), 
      categoria: norm(p.Remarks || p.Reference1),
      used: false,
      usedBy: "",
    });
  }

  return docs;
}
