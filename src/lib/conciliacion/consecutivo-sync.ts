import "server-only";
import ExcelJS from "exceljs";
import {
  getDriveItemId,
  folderPathForYear,
  uploadDriveFile,
  workbookCall,
  workbookPatchBatch,
} from "@/lib/graph/sharepoint";
import {
  buildConsecutivoSheets,
  COLUMNS,
  controlFileName,
  PROBLEMAS,
  WIDTHS,
  type ControlRow,
} from "./consecutivo";

type Cell = string | number;
type SheetData = { name: string; rows: ControlRow[] };

function rowValues(r: ControlRow): Cell[] {
  return [
    r.docNum, r.estado, r.fecha ?? "", r.banco ?? "", r.tercero ?? "", r.valor ?? "",
    r.conciliacion ?? "", r.ubicacion ?? "", r.observacion ?? "",
  ];
}

const isProblem = (r: ControlRow) =>
  r.estado === "FALTANTE" || (r.conciliacion !== undefined && PROBLEMAS.has(r.conciliacion));

/** Tabla Resumen y filas de Inconsistencias, calculadas desde las hojas de detalle. */
function buildSummaries(sheets: SheetData[]) {
  const resumen: Cell[][] = [["Hoja", "Vigentes", "Cancelados", "Faltantes", "Inconsistencias de conciliacion"]];
  const inconsistencias: { key: string; values: Cell[] }[] = [];
  const porTipo = new Map<string, number>();
  for (const sh of sheets) {
    let vig = 0, canc = 0, falt = 0, inc = 0;
    for (const r of sh.rows) {
      if (r.estado === "VIGENTE") vig++;
      else if (r.estado === "CANCELADO") canc++;
      else falt++;
      if (r.conciliacion && PROBLEMAS.has(r.conciliacion)) {
        inc++;
        porTipo.set(r.conciliacion, (porTipo.get(r.conciliacion) ?? 0) + 1);
      }
      if (r.estado === "FALTANTE") {
        porTipo.set("NUMERO FALTANTE EN EL CONSECUTIVO", (porTipo.get("NUMERO FALTANTE EN EL CONSECUTIVO") ?? 0) + 1);
      }
      if (isProblem(r)) {
        const v = rowValues(r);
        v[6] = r.estado === "FALTANTE" ? "NUMERO FALTANTE" : r.conciliacion ?? "";
        inconsistencias.push({ key: `${sh.name}|${r.docNum}`, values: [sh.name, ...v] });
      }
    }
    resumen.push([sh.name, vig, canc, falt, inc]);
  }
  resumen.push([]);
  resumen.push(["Tipo de inconsistencia", "Cantidad"]);
  for (const [tipo, n] of porTipo) resumen.push([tipo, n]);
  if (porTipo.size === 0) resumen.push(["Sin inconsistencias", 0]);
  return { resumen, inconsistencias };
}

/**
 * Crea el archivo de control completo. Solo se usa cuando el archivo todavia
 * no existe; para uno existente se usa syncConsecutivoFile, que nunca lo
 * reemplaza.
 */
async function buildNewWorkbook(sheets: SheetData[]): Promise<Buffer> {
  const { resumen, inconsistencias } = buildSummaries(sheets);
  const wb = new ExcelJS.Workbook();

  const rs = wb.addWorksheet("Resumen");
  for (const r of resumen) rs.addRow(r);
  rs.getRow(1).font = { bold: true };
  [44, 12, 12, 12, 30].forEach((w, i) => { rs.getColumn(i + 1).width = w; });

  const inc = wb.addWorksheet("Inconsistencias");
  inc.addRow(["Hoja", ...COLUMNS]);
  inc.getRow(1).font = { bold: true };
  for (const i of inconsistencias) {
    const row = inc.addRow(i.values);
    row.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } }; });
  }
  [30, ...WIDTHS].forEach((w, i) => { inc.getColumn(i + 1).width = w; });
  inc.views = [{ state: "frozen", ySplit: 1 }];

  for (const sh of sheets) {
    const ws = wb.addWorksheet(sh.name.slice(0, 31));
    ws.addRow(COLUMNS);
    ws.getRow(1).font = { bold: true };
    for (const r of sh.rows) ws.addRow(rowValues(r));
    WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.views = [{ state: "frozen", ySplit: 1 }];
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

const colLetter = (n: number) => String.fromCharCode(64 + n);

// La fecha va como texto (igual que en las filas ya existentes): con el
// apostrofo inicial Excel no la convierte a fecha de calendario.
const asText = (v: Cell[], fechaIdx: number): Cell[] => v.map((x, i) => (i === fechaIdx && x !== "" ? `'${x}` : x));
const sheetUrl = (name: string) => `/worksheets('${encodeURIComponent(name)}')`;

function sameValue(a: unknown, b: Cell): boolean {
  const na = Number(a === "" || a === null || a === undefined ? NaN : a);
  const nb = Number(b === "" ? NaN : b);
  if ((typeof a === "number" || typeof b === "number") && !Number.isNaN(na) && !Number.isNaN(nb)) {
    return Math.abs(na - nb) < 0.005;
  }
  return String(a ?? "") === String(b ?? "");
}

async function readUsedValues(itemId: string, sheet: string): Promise<Cell[][]> {
  const res = await workbookCall(itemId, "GET", `${sheetUrl(sheet)}/usedRange?$select=values`);
  if (!res.ok) return [];
  return ((await res.json()).values ?? []) as Cell[][];
}

/**
 * Actualiza el archivo de control YA EXISTENTE sin reemplazarlo. A las filas
 * que ya estan solo se les actualizan los valores de las columnas B..I (estado
 * de descarga, banco, etc.); los documentos nuevos se agregan al final de su
 * hoja. No se cambia ningun color, formato ni columna adicional, asi que las
 * marcas manuales de verificacion (filas en verde, notas) quedan intactas.
 * Escribe con la API de Excel (no con una subida completa), que ademas no
 * bloquea el archivo si alguien lo tiene abierto. Si el archivo no existe, lo
 * crea completo.
 */
export async function syncConsecutivoFile(
  year: number
): Promise<{ mode: "creado" | "actualizado"; filasNuevas: number; filasActualizadas: number }> {
  const folder = folderPathForYear(year);
  const fileName = controlFileName(year);
  const sheets = await buildConsecutivoSheets(year);

  const itemId = await getDriveItemId(fileName, folder);
  if (!itemId) {
    await uploadDriveFile(fileName, await buildNewWorkbook(sheets), folder);
    return { mode: "creado", filasNuevas: sheets.reduce((n, s) => n + s.rows.length, 0), filasActualizadas: 0 };
  }

  const { resumen, inconsistencias } = buildSummaries(sheets);
  const existingRes = await workbookCall(itemId, "GET", "/worksheets?$select=name");
  const existing = new Set<string>(((await existingRes.json()).value ?? []).map((w: { name: string }) => w.name));

  const patches: { relUrl: string; body: unknown }[] = [];
  let nuevas = 0, actualizadas = 0;

  const createSheet = async (name: string, header: string[], widths: number[]) => {
    const r = await workbookCall(itemId, "POST", "/worksheets/add", { name: name.slice(0, 31) });
    if (!r.ok) throw new Error(`No se pudo crear la hoja ${name}: ${await r.text()}`);
    existing.add(name);
    patches.push({ relUrl: `${sheetUrl(name)}/range(address='A1:${colLetter(header.length)}1')`, body: { values: [header] } });
    patches.push({ relUrl: `${sheetUrl(name)}/range(address='A1:${colLetter(header.length)}1')/format/font`, body: { bold: true } });
    widths.forEach((w, i) => {
      patches.push({ relUrl: `${sheetUrl(name)}/range(address='${colLetter(i + 1)}:${colLetter(i + 1)}')/format`, body: { columnWidth: w * 6 } });
    });
  };

  // Hojas de detalle
  for (const sh of sheets) {
    if (!existing.has(sh.name)) await createSheet(sh.name, COLUMNS, WIDTHS);
    const values = await readUsedValues(itemId, sh.name);
    const rowOf = new Map<number, number>(); // # documento -> fila (1-based)
    values.forEach((row, i) => { if (i > 0 && typeof row[0] === "number") rowOf.set(row[0], i + 1); });
    let next = Math.max(values.length, 1) + 1;
    const append: Cell[][] = [];
    for (const r of sh.rows) {
      const v = rowValues(r);
      const at = rowOf.get(r.docNum);
      if (at === undefined) { append.push(v); continue; }
      const cur = values[at - 1] ?? [];
      // la fecha guardada como numero (fecha de calendario) se reescribe como texto
      const mismaFecha = typeof cur[2] !== "number";
      if (mismaFecha && v.slice(1).every((x, i) => sameValue(cur[i + 1], x))) continue;
      patches.push({ relUrl: `${sheetUrl(sh.name)}/range(address='B${at}:I${at}')`, body: { values: [asText(v, 2).slice(1)] } });
      actualizadas++;
    }
    for (let i = 0; i < append.length; i += 100) {
      const chunk = append.slice(i, i + 100).map((v) => asText(v, 2));
      patches.push({ relUrl: `${sheetUrl(sh.name)}/range(address='A${next}:I${next + chunk.length - 1}')`, body: { values: chunk } });
      next += chunk.length;
    }
    nuevas += append.length;
  }

  // Inconsistencias: se actualizan las existentes (hoja + documento) con su
  // estado actual y se agregan las nuevas; no se borra ninguna fila.
  if (!existing.has("Inconsistencias")) await createSheet("Inconsistencias", ["Hoja", ...COLUMNS], [30, ...WIDTHS]);
  const incValues = await readUsedValues(itemId, "Inconsistencias");
  const incRow = new Map<string, number>();
  incValues.forEach((row, i) => { if (i > 0) incRow.set(`${row[0]}|${row[1]}`, i + 1); });
  const detalle = new Map<string, ControlRow>();
  for (const sh of sheets) for (const r of sh.rows) detalle.set(`${sh.name}|${r.docNum}`, r);
  for (const [key, at] of incRow) {
    const r = detalle.get(key);
    if (!r) continue;
    const v = rowValues(r);
    if (r.estado === "FALTANTE") v[6] = "NUMERO FALTANTE";
    const cur = incValues[at - 1] ?? [];
    if (typeof cur[3] !== "number" && v.slice(1).every((x, i) => sameValue(cur[i + 2], x))) continue;
    patches.push({ relUrl: `${sheetUrl("Inconsistencias")}/range(address='C${at}:J${at}')`, body: { values: [asText(v, 2).slice(1)] } });
  }
  let nextInc = Math.max(incValues.length, 1) + 1;
  const incNew = inconsistencias.filter((i) => !incRow.has(i.key));
  if (incNew.length > 0) {
    const first = nextInc;
    for (let i = 0; i < incNew.length; i += 100) {
      const chunk = incNew.slice(i, i + 100).map((x) => asText(x.values, 3));
      patches.push({ relUrl: `${sheetUrl("Inconsistencias")}/range(address='A${nextInc}:J${nextInc + chunk.length - 1}')`, body: { values: chunk } });
      nextInc += chunk.length;
    }
    patches.push({ relUrl: `${sheetUrl("Inconsistencias")}/range(address='A${first}:J${nextInc - 1}')/format/fill`, body: { color: "#FFC7CE" } });
  }

  // Resumen: tabla calculada, se reescribe completa (solo valores).
  if (!existing.has("Resumen")) {
    await createSheet("Resumen", ["Hoja", "Vigentes", "Cancelados", "Faltantes", "Inconsistencias de conciliacion"], [44, 12, 12, 12, 30]);
  }
  const oldResumen = await readUsedValues(itemId, "Resumen");
  const rowsOut = Math.max(oldResumen.length, resumen.length);
  const block: Cell[][] = [];
  for (let i = 0; i < rowsOut; i++) {
    const r = resumen[i] ?? [];
    block.push([0, 1, 2, 3, 4].map((c) => (r[c] ?? "") as Cell));
  }
  patches.push({ relUrl: `${sheetUrl("Resumen")}/range(address='A1:E${rowsOut}')`, body: { values: block } });

  const errors = await workbookPatchBatch(itemId, patches);
  if (errors.length > 0) throw new Error(`Errores al actualizar el control: ${errors.slice(0, 3).join(" | ")}`);
  return { mode: "actualizado", filasNuevas: nuevas, filasActualizadas: actualizadas };
}
