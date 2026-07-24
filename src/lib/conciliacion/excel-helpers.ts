import type ExcelJS from "exceljs";

export function setFill(cell: ExcelJS.Cell, argb: string): void {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}

export function setComment(cell: ExcelJS.Cell, text: string): void {
  cell.note = text;
}

/** Formatea una fecha (medianoche UTC) como YYYY-MM-DD, igual que str(date) en Python. */
export function fmtDate(d: Date): string {
  return d.toISOString().slice(0, 10);
}
