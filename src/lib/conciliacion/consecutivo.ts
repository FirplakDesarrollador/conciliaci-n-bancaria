import "server-only";
import ExcelJS from "exceljs";
import { sapClient } from "@/lib/sap/service-layer";
import { downloadDriveFile, folderPathForYear } from "@/lib/graph/sharepoint";
import { ACCOUNT_MAP, MANUAL_CUENTA_OVERRIDES, MIAMI_ACCOUNT, TRANSFER_ACCOUNT_NAMES } from "./config";
import { findHeaderRowAndCols, READERS, REQUIRED_HEADERS_HINT } from "./readers";

// En 2026 el control arranca en octubre (decision del usuario, 2026-10-01:
// setiembre ya quedo cerrado sin este control). Desde 2027 se genera un
// archivo por ano, que cubre el ano completo y se guarda en la carpeta
// "FIRPLAK <ano>" correspondiente.
export function controlRangeForYear(year: number): { from: string; to: string } {
  return { from: year === 2026 ? "2026-10-01" : `${year}-01-01`, to: `${year}-12-31` };
}

export function controlFileName(year: number): string {
  return `CONTROL_CONSECUTIVO_PAGOS_${year}.xlsx`;
}

const FIELDS = "DocNum,DocEntry,Series,DocDate,Cancelled,CardCode,CardName,TransferAccount,CashAccount,TransferSum,CashSum,DocCurrency,DocRate,Remarks,JournalRemarks";
const FIDUCIA_ACCOUNT = "12450505";

interface RawDoc {
  DocNum: number;
  Series: number;
  DocDate: string;
  Cancelled: string;
  CardCode?: string;
  CardName?: string;
  DocCurrency?: string;
  DocRate?: number;
  TransferAccount?: string | null;
  CashAccount?: string | null;
  TransferSum?: number;
  CashSum?: number;
}

async function fetchAllBetween(
  entity: "IncomingPayments" | "VendorPayments",
  fromDate: string,
  toDate: string
): Promise<RawDoc[]> {
  let all: RawDoc[] = [];
  let next: string | null = `/${entity}?$filter=DocDate ge '${fromDate}' and DocDate le '${toDate}'&$select=${FIELDS}&$orderby=DocNum`;
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

// Un documento puede crearse hoy con fecha anterior (ej. cierre del mes: DocNum
// 82319 creado en octubre con DocDate 30-sep). Por fecha quedaria fuera de la
// ventana y su numero se reportaria como FALTANTE. Para cada serie se
// completa la corrida de numeracion con los documentos de esa serie cuyo
// DocNum es >= al primero del periodo aunque su fecha sea anterior.
async function withBackdatedDocs(
  entity: "IncomingPayments" | "VendorPayments",
  docs: RawDoc[],
  fromDate: string
): Promise<RawDoc[]> {
  const minBySeries = new Map<number, number>();
  const maxBySeries = new Map<number, number>();
  for (const d of docs) {
    const cur = minBySeries.get(d.Series);
    if (cur === undefined || d.DocNum < cur) minBySeries.set(d.Series, d.DocNum);
    const mx = maxBySeries.get(d.Series);
    if (mx === undefined || d.DocNum > mx) maxBySeries.set(d.Series, d.DocNum);
  }
  const byNum = new Map<string, RawDoc>();
  for (const d of docs) byNum.set(`${d.Series}:${d.DocNum}`, d);
  for (const [series, min] of minBySeries) {
    const max = maxBySeries.get(series)!;
    let next: string | null = `/${entity}?$filter=Series eq ${series} and DocNum ge ${min} and DocNum le ${max} and DocDate lt '${fromDate}'&$select=${FIELDS}&$orderby=DocNum`;
    while (next) {
      const res = await sapClient.request(next);
      if (!res.ok) break;
      const data = await res.json();
      for (const d of (data.value || []) as RawDoc[]) byNum.set(`${d.Series}:${d.DocNum}`, d);
      const link = data["odata.nextLink"];
      next = link ? (String(link).startsWith("/") ? link : "/" + link) : null;
    }
  }
  return [...byNum.values()];
}

const normalizeStr = (s: string) =>
  s.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Z0-9]/gi, "").toUpperCase();

// ---------------------------------------------------------------------------
// Validacion contra los archivos de banco
// ---------------------------------------------------------------------------

interface Loc {
  bankKey: string;
  sheet: string;
  row: number;
  /** la celda del banco lista varios documentos (asignacion por grupo/combo) */
  grupo: boolean;
}

interface BankIndex {
  /** numero de documento -> filas de banco donde esta escrito */
  byDoc: Map<string, Loc[]>;
  /** llaves de ACCOUNT_MAP cuyo archivo del ano se pudo leer */
  available: Set<string>;
}

function bankFileForYear(file: string, year: number): string {
  return file.replace(/\d{4}\.xlsx$/, `${year}.xlsx`);
}

async function buildBankIndex(year: number): Promise<BankIndex> {
  const byDoc = new Map<string, Loc[]>();
  const available = new Set<string>();
  const folder = folderPathForYear(year);

  await Promise.all(
    Object.entries(ACCOUNT_MAP).map(async ([bankKey, cfg]) => {
      let buffer: Buffer;
      try {
        buffer = await downloadDriveFile(bankFileForYear(cfg.file, year), folder);
      } catch (e) {
        console.error(`Control consecutivo: no se pudo descargar ${bankKey}:`, e);
        return;
      }
      const wb = new ExcelJS.Workbook();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      await wb.xlsx.load(buffer as any);
      available.add(bankKey);

      for (const ws of wb.worksheets) {
        const found = findHeaderRowAndCols(ws, REQUIRED_HEADERS_HINT[cfg.format]);
        if (!found) continue;
        const seenRows = new Set<number>();
        for (const mv of READERS[cfg.format](ws, found.headers, found.headerRow)) {
          if (seenRows.has(mv.row)) continue;
          seenRows.add(mv.row);
          if (mv.docValue === null || mv.docValue === undefined || mv.docValue === "") continue;
          const nums = (String(mv.docValue).match(/\d+/g) ?? []).filter((n) => n.length >= 5);
          for (const num of nums) {
            if (!byDoc.has(num)) byDoc.set(num, []);
            byDoc.get(num)!.push({ bankKey, sheet: ws.name, row: mv.row, grupo: nums.length > 1 });
          }
        }
      }
    })
  );

  return { byDoc, available };
}

function accountKeyOf(code?: string | null): string | undefined {
  if (!code) return undefined;
  const name = TRANSFER_ACCOUNT_NAMES[code];
  if (!name) return undefined;
  const n = normalizeStr(name);
  return Object.keys(ACCOUNT_MAP).find(
    (k) => normalizeStr(k) === n || n.includes(normalizeStr(k)) || normalizeStr(k).includes(n)
  );
}

/**
 * Bancos (llaves de ACCOUNT_MAP) donde SAP indica que debe quedar el
 * documento: la cuenta de transferencia/caja y, en traslados entre cuentas
 * propias (Fiducia, venta de dolares), el banco que viaja en CardCode.
 */
function expectedBanks(d: RawDoc): string[] {
  const acc = d.TransferAccount || d.CashAccount || "";
  const keys = new Set<string>();
  if (acc !== FIDUCIA_ACCOUNT) {
    const k = accountKeyOf(acc);
    if (k) keys.add(k);
  }
  const ck = accountKeyOf(d.CardCode);
  if (ck) keys.add(ck);
  return [...keys];
}

export const PROBLEMAS = new Set([
  "NO DESCARGADO",
  "DESCARGA PARCIAL",
  "BANCO DISTINTO",
  "VARIAS FILAS",
  "CANCELADO AUN DESCARGADO",
]);

interface Validacion {
  conciliacion: string;
  ubicacion: string;
  observacion: string;
}

function validar(d: RawDoc, index: BankIndex, bancoSap: string): Validacion {
  const num = String(d.DocNum);
  const locs = index.byDoc.get(num) ?? [];
  const ubicacion = locs.map((l) => `${l.bankKey} ${l.sheet} fila ${l.row}`).join(" | ");

  if (d.Cancelled === "tYES") {
    return locs.length > 0
      ? { conciliacion: "CANCELADO AUN DESCARGADO", ubicacion, observacion: "El documento esta cancelado en SAP pero sigue escrito en el banco." }
      : { conciliacion: "No aplica (cancelado)", ubicacion: "", observacion: "" };
  }

  const expected = expectedBanks(d);
  const override = MANUAL_CUENTA_OVERRIDES[num];
  const expectedWithFile = expected.filter((k) => index.available.has(k));

  if (expectedWithFile.length === 0) {
    if (locs.length > 0) {
      return {
        conciliacion: "BANCO DISTINTO",
        ubicacion,
        observacion: `SAP indica ${bancoSap} (sin archivo de banco) pero esta descargado en un banco.${override ? " Correccion manual conocida." : ""}`,
      };
    }
    return { conciliacion: "No aplica (sin archivo de banco)", ubicacion: "", observacion: "" };
  }

  const inExpected = locs.filter((l) => expectedWithFile.includes(l.bankKey));
  const inOther = locs.filter((l) => !expectedWithFile.includes(l.bankKey));

  if (inOther.length > 0) {
    return {
      conciliacion: "BANCO DISTINTO",
      ubicacion,
      observacion: `SAP indica ${expectedWithFile.join(" / ")} pero esta en ${[...new Set(inOther.map((l) => l.bankKey))].join(" / ")}.${override ? " Correccion manual conocida." : " Posible error de banco asignado en SAP."}`,
    };
  }

  const banksCovered = new Set(inExpected.map((l) => l.bankKey));
  const missing = expectedWithFile.filter((k) => !banksCovered.has(k));
  if (missing.length === expectedWithFile.length) {
    return { conciliacion: "NO DESCARGADO", ubicacion: "", observacion: `No aparece en ${missing.join(" / ")}.` };
  }
  if (missing.length > 0) {
    return { conciliacion: "DESCARGA PARCIAL", ubicacion, observacion: `Falta en ${missing.join(" / ")}.` };
  }

  for (const k of banksCovered) {
    const enBanco = inExpected.filter((l) => l.bankKey === k);
    const rows = new Set(enBanco.map((l) => `${l.sheet}!${l.row}`));
    // Varias filas con el documento listado junto a otros = asignacion por grupo
    // (ej. reembolsos de caja menor contra varios retiros de cajero): es el
    // diseno, no un duplicado.
    if (rows.size > 1 && enBanco.every((l) => l.grupo)) {
      return { conciliacion: "Descargado", ubicacion, observacion: `Descargado en grupo: el documento comparte ${rows.size} filas de ${k} con otros documentos (categoria con varios movimientos).` };
    }
    if (rows.size > 1) {
      return { conciliacion: "VARIAS FILAS", ubicacion, observacion: `El mismo documento esta en ${rows.size} filas de ${k}.` };
    }
  }
  return { conciliacion: "Descargado", ubicacion, observacion: "" };
}

// ---------------------------------------------------------------------------
// Consecutivo + nombres de banco
// ---------------------------------------------------------------------------

export interface ControlRow {
  docNum: number;
  estado: "VIGENTE" | "CANCELADO" | "FALTANTE";
  fecha?: string;
  banco?: string;
  tercero?: string;
  valor?: number;
  conciliacion?: string;
  ubicacion?: string;
  observacion?: string;
}

// Banco del documento: la cuenta de transferencia/caja de SAP traducida al
// nombre del banco. En los traslados a Fiducia la cuenta de transferencia es
// la de la Fiducia y el banco real viaja en CardCode (ver convertApiToSapDocs).
function bancoDe(d: RawDoc, extraNames: Map<string, string>): string {
  const nameOf = (code: string) => TRANSFER_ACCOUNT_NAMES[code] ?? extraNames.get(code) ?? code;
  const acc = d.TransferAccount || d.CashAccount || "";
  if (acc === FIDUCIA_ACCOUNT) {
    return `${d.CardCode ? nameOf(d.CardCode) : ""} (traslado Fiducia)`.trim();
  }
  return nameOf(acc);
}

// Cuentas que no son bancos de TRANSFER_ACCOUNT_NAMES (ej. "CAJA CONFIRMING",
// "CAJA TARJETA CREDITO") se traducen con el nombre del plan de cuentas.
async function fetchAccountNames(codes: string[]): Promise<Map<string, string>> {
  const names = new Map<string, string>();
  const missing = [...new Set(codes)].filter((c) => c && !TRANSFER_ACCOUNT_NAMES[c] && c !== FIDUCIA_ACCOUNT);
  for (let i = 0; i < missing.length; i += 15) {
    const chunk = missing.slice(i, i + 15);
    const filter = chunk.map((c) => `Code eq '${c}'`).join(" or ");
    const res = await sapClient.request(`/ChartOfAccounts?$filter=${filter}&$select=Code,Name`);
    if (!res.ok) continue;
    for (const a of (await res.json()).value || []) names.set(a.Code, a.Name);
  }
  return names;
}

interface KeyedRow {
  monthKey: string;
  row: ControlRow;
}

function monthKeyOf(isoDate: string): string {
  const dt = new Date(isoDate);
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth()).padStart(2, "0")}`;
}

// El consecutivo se revisa sobre TODA la corrida de numeracion de la serie
// (de su menor a su mayor numero), no mes por mes; cada fila va luego a la hoja
// del mes de su fecha. Un numero que no existe en SAP en ninguna fecha
// (FALTANTE) se ubica en el mes del documento anterior de la secuencia.
function buildSeriesRows(docs: RawDoc[], extraNames: Map<string, string>, index: BankIndex, fromDate: string): KeyedRow[] {
  const byDocNum = new Map<number, RawDoc>();
  for (const d of docs) byDocNum.set(d.DocNum, d);
  const nums = [...byDocNum.keys()];
  if (nums.length === 0) return [];
  const min = Math.min(...nums);
  const max = Math.max(...nums);
  // Proteccion: una serie con numeracion saltada de millones no es un consecutivo
  // revisable (y generaria millones de filas); se omite en vez de agotar memoria.
  if (max - min > 20000) {
    console.error(`Control consecutivo: serie omitida, rango ${min}-${max} demasiado amplio`);
    return [];
  }

  const rows: KeyedRow[] = [];
  let lastKey = monthKeyOf(byDocNum.get(min)!.DocDate);
  for (let n = min; n <= max; n++) {
    const d = byDocNum.get(n);
    if (!d) {
      rows.push({ monthKey: lastKey, row: { docNum: n, estado: "FALTANTE" } });
      continue;
    }
    // Un documento con fecha anterior al periodo (creado hoy con fecha del mes
    // pasado) queda en la hoja que le corresponde por su numero, para que el
    // consecutivo de cada hoja este completo, y se anota donde se descargo.
    const fechaSap = d.DocDate.slice(0, 10);
    const retroactivo = fechaSap < fromDate;
    if (!retroactivo) lastKey = monthKeyOf(d.DocDate);
    // La cuenta de compensacion Miami opera en USD: sus movimientos de banco
    // estan en dolares, asi que el documento se muestra en esa moneda
    // (TransferSum de SAP viene en pesos a la tasa del documento).
    const acc = d.TransferAccount || d.CashAccount || "";
    const esMiami = accountKeyOf(acc) === MIAMI_ACCOUNT;
    const totalCop = d.TransferSum || d.CashSum || 0;
    const usd = d.DocCurrency === "USD" && (d.DocRate ?? 0) > 0 ? Math.round((totalCop / d.DocRate!) * 100) / 100 : null;
    let banco = bancoDe(d, extraNames);
    if (esMiami && usd !== null) banco = `${banco} (USD)`;
    const v = validar(d, index, banco);
    // Traslado entre cuentas propias cuya otra pata es Miami (venta de dolares)
    if (!esMiami && usd !== null && accountKeyOf(d.CardCode) === MIAMI_ACCOUNT) {
      const nota = `Pata Miami: USD ${usd.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}.`;
      v.observacion = v.observacion ? `${v.observacion} ${nota}` : nota;
    }
    if (retroactivo) {
      const hojas = [...new Set((index.byDoc.get(String(n)) ?? []).map((l) => l.sheet))];
      const nota =
        `Documento con fecha SAP ${fechaSap} (anterior al periodo de control), creado dentro del periodo.` +
        (hojas.length > 0 ? ` Descargado en ${hojas.join(" / ")}.` : "");
      v.observacion = v.observacion ? `${nota} ${v.observacion}` : nota;
    }
    rows.push({
      monthKey: lastKey,
      row: {
        docNum: n,
        estado: d.Cancelled === "tYES" ? "CANCELADO" : "VIGENTE",
        fecha: d.DocDate.slice(0, 10),
        banco,
        tercero: d.CardName || "",
        valor: esMiami && usd !== null ? usd : totalCop,
        conciliacion: v.conciliacion,
        ubicacion: v.ubicacion,
        observacion: v.observacion,
      },
    });
  }
  return rows;
}

function monthNameEs(monthIndex0: number): string {
  const names = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  return names[monthIndex0];
}

export const COLUMNS = ["# Documento", "Estado", "Fecha", "Banco SAP", "Tercero", "Valor", "Conciliacion", "Ubicacion en bancos", "Observacion"];
export const WIDTHS = [14, 12, 12, 42, 40, 16, 30, 50, 60];

/**
 * Datos del control: una hoja por mes y por serie (recibidos serie 151,
 * efectuados series 70, 71, -1...; el consecutivo se revisa POR SERIE). Cada
 * numero entre el minimo y el maximo visto queda VIGENTE / CANCELADO /
 * FALTANTE, y cada documento vigente se valida contra los archivos de banco
 * del ano (Descargado, NO DESCARGADO, DESCARGA PARCIAL, BANCO DISTINTO,
 * VARIAS FILAS, CANCELADO AUN DESCARGADO; las cajas son No aplica).
 */
export async function buildConsecutivoSheets(year: number): Promise<{ name: string; rows: ControlRow[] }[]> {
  const { from, to } = controlRangeForYear(year);
  const [incomingBase, vendorBase, index] = await Promise.all([
    fetchAllBetween("IncomingPayments", from, to),
    fetchAllBetween("VendorPayments", from, to),
    buildBankIndex(year),
  ]);
  const [incoming, vendor] = await Promise.all([
    withBackdatedDocs("IncomingPayments", incomingBase, from),
    withBackdatedDocs("VendorPayments", vendorBase, from),
  ]);

  const extraNames = await fetchAccountNames(
    [...incoming, ...vendor].map((d) => {
      const acc = d.TransferAccount || d.CashAccount || "";
      return acc === FIDUCIA_ACCOUNT ? d.CardCode || "" : acc;
    })
  );

  const sheets: { name: string; rows: ControlRow[] }[] = [];
  const pushSheets = (prefix: string, keyed: KeyedRow[]) => {
    const byKey = new Map<string, ControlRow[]>();
    for (const k of keyed) {
      if (!byKey.has(k.monthKey)) byKey.set(k.monthKey, []);
      byKey.get(k.monthKey)!.push(k.row);
    }
    for (const [key, rows] of [...byKey.entries()].sort()) {
      sheets.push({ name: `${prefix} ${monthNameEs(Number(key.split("-")[1]))}`, rows });
    }
  };

  pushSheets("Recibidos", buildSeriesRows(incoming, extraNames, index, from));

  const vendorBySeries = new Map<number, RawDoc[]>();
  for (const d of vendor) {
    if (!vendorBySeries.has(d.Series)) vendorBySeries.set(d.Series, []);
    vendorBySeries.get(d.Series)!.push(d);
  }
  for (const [series, docsOfSeries] of [...vendorBySeries.entries()].sort((a, b) => a[0] - b[0])) {
    pushSheets(`Efectuados S${series}`, buildSeriesRows(docsOfSeries, extraNames, index, from));
  }
  return sheets;
}

