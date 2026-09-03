import { IGNORE_KEYWORDS } from "./config";

/** Extrae el valor "plano" de una celda de exceljs: desenreda fórmulas,
 * richText e hipervínculos, que Python/openpyxl ya entrega como texto/num. */
export function cellValue(v: unknown): unknown {
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    const obj = v as Record<string, unknown>;
    if ("result" in obj) return cellValue(obj.result);
    if ("richText" in obj && Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text: string }>).map((t) => t.text).join("");
    }
    if ("text" in obj && "hyperlink" in obj) return obj.text as string;
  }
  return v;
}

export function norm(s: unknown): string {
  const v = cellValue(s);
  if (v === null || v === undefined) return "";
  return String(v).replace(/\s+/g, " ").trim().toUpperCase();
}

export function normAccount(s: unknown): string {
  return norm(s);
}

export function parseMoney(v: unknown): number | null {
  const value = cellValue(v);
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;
  const s = String(value).replace(/,/g, "").replace(/\$/g, "").trim();
  const n = parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

/** Fecha "calendario" representada como medianoche UTC, para poder comparar
 * y restar en días sin corrimientos de huso horario. */
export function parseDate(v: unknown): Date | null {
  const value = cellValue(v);
  if (value === null || value === undefined || value === "") return null;

  if (value instanceof Date) {
    return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
  }

  if (typeof value === "number") {
    const s = String(Math.trunc(value));
    if (s.length === 8) {
      const y = Number(s.slice(0, 4));
      const m = Number(s.slice(4, 6));
      const d = Number(s.slice(6, 8));
      const date = new Date(Date.UTC(y, m - 1, d));
      return Number.isNaN(date.getTime()) ? null : date;
    }
    return null;
  }

  if (typeof value === "string") {
    const s = value.trim();

    let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (m) return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));

    m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
    if (m) return new Date(Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1])));

    m = /^(\d{8})$/.exec(s);
    if (m) {
      return new Date(Date.UTC(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, Number(s.slice(6, 8))));
    }
  }

  return null;
}

export function daysBetween(a: Date, b: Date): number {
  return Math.round((a.getTime() - b.getTime()) / 86_400_000);
}

// Un documento de un mes ya cerrado no debe cruzarse contra un movimiento
// bancario de otro mes solo porque cae dentro de la tolerancia de días.
export function sameMonth(a: Date, b: Date): boolean {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
}

export function isBankFee(refText: string): boolean {
  const t = norm(refText);
  return IGNORE_KEYWORDS.some((kw) => t.includes(kw));
}

/**
 * Extrae el código que identifica documentos "hermanos" del mismo cobro/pago:
 * - Pagos recibidos: '11100535 // LINK DE PAGO-' o '11100535 EPAYCO' -> '11100535'
 * - Pagos efectuados: 'AC70076116-01' / 'AC70076116-02' -> 'AC70076116'
 */
export function infoCode(s: string): string {
  if (!s) return "";
  let token = s.split(" ")[0];
  token = token.replace(/[/-]+$/, "").trim();
  const m = /^(.*)-\d+$/.exec(token);
  if (m) token = m[1];
  return token;
}

export function rangeInclusive(start: number, end: number): number[] {
  const result: number[] = [];
  for (let i = start; i <= end; i++) result.push(i);
  return result;
}

export function combinations<T>(items: T[], size: number): T[][] {
  const result: T[][] = [];
  if (size > items.length || size <= 0) return result;

  const combo: T[] = [];
  function backtrack(start: number) {
    if (combo.length === size) {
      result.push([...combo]);
      return;
    }
    for (let i = start; i < items.length; i++) {
      combo.push(items[i]);
      backtrack(i + 1);
      combo.pop();
    }
  }
  backtrack(0);
  return result;
}

// --- Ratio de similitud tipo difflib.SequenceMatcher (Ratcliff/Obershelp) ---

function longestMatch(
  a: string,
  b: string,
  aLo: number,
  aHi: number,
  bLo: number,
  bHi: number
): [number, number, number] {
  let besti = aLo;
  let bestj = bLo;
  let bestSize = 0;
  let j2len = new Map<number, number>();

  for (let i = aLo; i < aHi; i++) {
    const newJ2Len = new Map<number, number>();
    for (let j = bLo; j < bHi; j++) {
      if (a[i] === b[j]) {
        const k = (j2len.get(j - 1) ?? 0) + 1;
        newJ2Len.set(j, k);
        if (k > bestSize) {
          besti = i - k + 1;
          bestj = j - k + 1;
          bestSize = k;
        }
      }
    }
    j2len = newJ2Len;
  }

  return [besti, bestj, bestSize];
}

function matchingBlocksTotal(a: string, b: string): number {
  let total = 0;
  const queue: Array<[number, number, number, number]> = [[0, a.length, 0, b.length]];

  while (queue.length > 0) {
    const [aLo, aHi, bLo, bHi] = queue.pop()!;
    const [i, j, k] = longestMatch(a, b, aLo, aHi, bLo, bHi);
    if (k > 0) {
      total += k;
      if (aLo < i && bLo < j) queue.push([aLo, i, bLo, j]);
      if (i + k < aHi && j + k < bHi) queue.push([i + k, aHi, j + k, bHi]);
    }
  }

  return total;
}

export function sequenceRatio(a: string, b: string): number {
  if (a.length === 0 && b.length === 0) return 1;
  const matches = matchingBlocksTotal(a, b);
  return (2 * matches) / (a.length + b.length);
}
