import "server-only";
import ExcelJS from "exceljs";
import { sapClient } from "@/lib/sap/service-layer";

// A partir de esta fecha se vigila el consecutivo (decision del usuario,
// 2026-10-01: "solo octubre en adelante", setiembre ya quedo cerrado sin
// este control).
export const CONTROL_START_DATE = "2026-10-01";

const FIELDS = "DocNum,DocEntry,Series,DocDate,Cancelled,CardName,TransferSum,CashSum,Remarks,JournalRemarks";

interface RawDoc {
  DocNum: number;
  Series: number;
  DocDate: string;
  Cancelled: string;
  CardName?: string;
  TransferSum?: number;
  CashSum?: number;
  Remarks?: string;
  JournalRemarks?: string;
}

async function fetchAllSince(entity: "IncomingPayments" | "VendorPayments", fromDate: string): Promise<RawDoc[]> {
  let all: RawDoc[] = [];
  let next: string | null = `/${entity}?$filter=DocDate ge '${fromDate}'&$select=${FIELDS}&$orderby=DocNum`;
  while (next) {
    const res = await sapClient.request(next);
    if (!res.ok) break;
    const data = await res.json();
    all = all.concat(data.value || []);
    const link = data["odata.nextLink"];
    next = link ? (String(link).startsWith("/") ? link : "/" + link) : null;
  }
  return all;
}

interface ControlRow {
  docNum: number;
  estado: "OK" | "CANCELADO" | "FALTANTE";
  fecha?: string;
  tercero?: string;
  valor?: number;
}

function buildControlRows(docs: RawDoc[]): ControlRow[] {
  const byDocNum = new Map<number, RawDoc>();
  for (const d of docs) byDocNum.set(d.DocNum, d);
  const nums = [...byDocNum.keys()];
  if (nums.length === 0) return [];
  const min = Math.min(...nums);
  const max = Math.max(...nums);

  const rows: ControlRow[] = [];
  for (let n = min; n <= max; n++) {
    const d = byDocNum.get(n);
    if (!d) {
      rows.push({ docNum: n, estado: "FALTANTE" });
      continue;
    }
    rows.push({
      docNum: n,
      estado: d.Cancelled === "tYES" ? "CANCELADO" : "OK",
      fecha: d.DocDate.slice(0, 10),
      tercero: d.CardName || "",
      valor: d.TransferSum || d.CashSum || 0,
    });
  }
  return rows;
}

function monthNameEs(monthIndex0: number): string {
  const names = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  return names[monthIndex0];
}

function writeSheet(wb: ExcelJS.Workbook, name: string, rows: ControlRow[]) {
  const ws = wb.addWorksheet(name.slice(0, 31));
  ws.addRow(["# Documento", "Estado", "Fecha", "Tercero", "Valor"]);
  ws.getRow(1).font = { bold: true };
  for (const r of rows) {
    const row = ws.addRow([r.docNum, r.estado, r.fecha ?? "", r.tercero ?? "", r.valor ?? ""]);
    if (r.estado === "FALTANTE") {
      row.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } }; });
    } else if (r.estado === "CANCELADO") {
      row.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFEB9C" } }; });
    }
  }
  [14, 12, 12, 40, 16].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

/**
 * Genera el archivo de control de consecutivo para pagos recibidos
 * (IncomingPayments, serie unica 151) y pagos efectuados (VendorPayments,
 * puede tener varias series -- ej. 70 para ACH electronico, -1 para
 * pagos manuales -- cada una con su propia numeracion independiente, asi
 * que el consecutivo se revisa POR SERIE, no por tipo de documento).
 * Una hoja por mes y por serie; dentro de cada hoja, cada numero entre el
 * minimo y el maximo visto queda marcado OK / CANCELADO / FALTANTE.
 */
export async function buildConsecutivoWorkbook(): Promise<Buffer> {
  const [incoming, vendor] = await Promise.all([
    fetchAllSince("IncomingPayments", CONTROL_START_DATE),
    fetchAllSince("VendorPayments", CONTROL_START_DATE),
  ]);

  const wb = new ExcelJS.Workbook();

  const byMonth = (docs: RawDoc[]) => {
    const m = new Map<string, RawDoc[]>();
    for (const d of docs) {
      const dt = new Date(d.DocDate);
      const key = `${dt.getUTCFullYear()}-${dt.getUTCMonth()}`;
      if (!m.has(key)) m.set(key, []);
      m.get(key)!.push(d);
    }
    return m;
  };

  const incomingByMonth = byMonth(incoming);
  for (const [key, docs] of [...incomingByMonth.entries()].sort()) {
    const [, monthIdx] = key.split("-").map(Number);
    const rows = buildControlRows(docs);
    writeSheet(wb, `Recibidos ${monthNameEs(monthIdx)}`, rows);
  }

  const vendorBySeries = new Map<number, RawDoc[]>();
  for (const d of vendor) {
    if (!vendorBySeries.has(d.Series)) vendorBySeries.set(d.Series, []);
    vendorBySeries.get(d.Series)!.push(d);
  }
  for (const [series, docsOfSeries] of vendorBySeries) {
    const byMonthForSeries = byMonth(docsOfSeries);
    for (const [key, docs] of [...byMonthForSeries.entries()].sort()) {
      const [, monthIdx] = key.split("-").map(Number);
      const rows = buildControlRows(docs);
      writeSheet(wb, `Efectuados S${series} ${monthNameEs(monthIdx)}`, rows);
    }
  }

  const resumen = wb.addWorksheet("Resumen", { views: [{ state: "frozen", ySplit: 1 }] });
  resumen.addRow(["Hoja", "OK", "Cancelados", "Faltantes"]);
  resumen.getRow(1).font = { bold: true };
  for (const ws of wb.worksheets) {
    if (ws.name === "Resumen") continue;
    let ok = 0, canc = 0, falt = 0;
    ws.eachRow((row, i) => {
      if (i === 1) return;
      const estado = row.getCell(2).value;
      if (estado === "OK") ok++;
      else if (estado === "CANCELADO") canc++;
      else if (estado === "FALTANTE") falt++;
    });
    resumen.addRow([ws.name, ok, canc, falt]);
  }
  wb.worksheets.sort((a, b) => (a.name === "Resumen" ? -1 : b.name === "Resumen" ? 1 : 0));
  [40, 10, 14, 14].forEach((w, i) => { resumen.getColumn(i + 1).width = w; });

  return Buffer.from(await wb.xlsx.writeBuffer());
}
