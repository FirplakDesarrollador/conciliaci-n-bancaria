export type MovTipo = "IN" | "OUT";

export interface SapDoc {
  docNum: string;
  date: Date;
  value: number;
  tercero: string;
  cuenta: string;
  comentario: string;
  tipo: MovTipo;
  infoDetallada: string;
  categoria: string;
  used: boolean;
  usedBy: string;
}

export interface BankMove {
  sheet: string;
  row: number;
  date: Date;
  value: number;
  tipo: MovTipo;
  docCol: number;
  docValue: unknown;
  refText: string;
  /** Cuenta(s) SAP a las que debe cruzarse este movimiento en vez de la
   * cuenta por defecto de la cuenta bancaria (para traslados a fiducia/fondo). */
  cuentaOverride?: string | string[];
  /** Restringe el pool de documentos SAP candidatos a un tercero específico. */
  terceroFilter?: string;
}

export interface SpecialRule {
  pattern: RegExp;
  tipo: MovTipo;
  cuenta: string | string[];
  tercero?: string;
}

export interface SummaryRow {
  cuentaSap: string;
  hoja: string;
  fila: number;
  tipo: MovTipo;
  fecha: Date;
  valor: number;
  estado: "AMBIGUO" | "SIN DOCUMENTO";
  candidatos: string;
}

export interface AccountStats {
  cuentaKey: string;
  archivo: string;
  matchExacto: number;
  matchTolerancia: number;
  matchValorUnico: number;
  matchPar: number;
  matchConteo: number;
  matchInverso: number;
  matchGrupo: number;
  ambiguos: number;
  sinDocumento: number;
  comisionesIgnoradas: number;
  yaTeniaDocumento: number;
  fueraDeRango: number;
}
