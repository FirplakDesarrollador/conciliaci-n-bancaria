// Tolerancia de fecha para el cruce (días de diferencia permitidos entre la
// fecha del movimiento bancario y la fecha del documento SAP). Se intenta
// primero fecha EXACTA; si no hay match, se amplía a esta tolerancia.
export const DATE_TOLERANCE_DAYS = 3;

// Tolerancia de valor (cuentas en pesos, 1 peso).
export const VALUE_TOLERANCE = 1.0;

// Tolerancia de valor para sumas de varios recibos (redondeos acumulados de
// centavos entre 2 o 3 documentos pueden dejar una diferencia de 1-2 pesos
// frente al valor consignado en el banco).
export const MULTI_VALUE_TOLERANCE = 3.0;

export type BankFormat = "bogota" | "bancolombia" | "sudameris" | "davivienda" | "compensacion";

// Mapeo entre el texto de "Banco"/"Cuenta de transferencia" del informe SAP
// y el archivo real en SharePoint (carpeta FIRPLAK 2026) + el formato de
// lectura. Los nombres de archivo se ajustaron a como están hoy en
// SharePoint (no los que asumía el script original en disco local).
//
// NOTA: "BANCO DE BOGOTA MIAMI CTA#79038" se asume que corresponde a la
// cuenta de compensación en USD (Compensación Bogotá). Las cuentas
// "CAJA..." y "FIDECOMISO" no tienen archivo bancario asociado, por lo que
// se omiten del cruce. El archivo "FIDUBOGOTÁ MAYO 2026.xlsx" tampoco tiene
// cuenta SAP mapeada y se deja fuera del cruce automático.
export const ACCOUNT_MAP: Record<string, { file: string; format: BankFormat }> = {
  "BANCO DE BOGOTA # 406007252": { file: "CTA CTE BOGOTA 406007252 2026.xlsx", format: "bogota" },
  "BANCOL.CTE # 008-927404-01": {
    file: "CTA 07 CTE 00892740401 BANCOLOMBIA 2026.xlsx",
    format: "bancolombia",
  },
  SUDAMERIS: { file: "CTA CTE SUDAMERIS 61406765 2026.xlsx", format: "sudameris" },
  "BANCO DAVIVIENDA CTE #3979-6999-8061": {
    file: "CTA CTE DAVIVIENDA 397969998061 2026.xlsx",
    format: "davivienda",
  },
  "BANCO DE BOGOTA MIAMI CTA#79038": {
    file: "CTA COMPENSACION BOGOTA 2026.xlsx",
    format: "compensacion",
  },
};

// Conceptos que el banco genera automáticamente (comisiones, impuestos,
// gravámenes) y que por naturaleza NUNCA tienen un documento SAP asociado.
// Estos movimientos se dejan sin resaltar (no se marcan como "sin
// documento"), aunque tampoco se les asigna un número de documento.
export const IGNORE_KEYWORDS = [
  "RETEICA",
  "RETEFUENTE",
  "RETEIVA",
  "GRAVAMEN MOVIMIENTOS FINANCIER",
  "GMF GRAVAMEN",
  "COBRO 4X1.000",
  "4X1000",
  "CARGO IVA",
  "COBRO IVA",
  "IVA COMISION",
  "COMISION COBRADA A ESTABLECIMIENTO",
  "COMISION REC CORRESPONSAL",
  "COMISION RECAUDO CAJA",
  "COMISION RECAUDOS CON COMPROBANTE",
  "COMISION DISFON",
  "COBRO COMISION",
  "COBRO DE COMISION",
  "COBRO SERVICIO MANEJO PORTAL",
  "SERVICIO PAGO A PROVEEDORES",
  "SERVICIO PAGO DE NOMINA",
  "INTERNET TFR FEE",
];

// Correcciones manuales para documentos SAP donde el campo "Banco"/"Cuenta
// de transferencia" quedó mal digitado y no corresponde al banco real donde
// se registra el movimiento. Se identifican caso a caso.
export const MANUAL_CUENTA_OVERRIDES: Record<string, string> = {
  // El 20042394 quedó marcado como SUDAMERIS en SAP, pero el pago (PSE a
  // Enlace Operativo S.A., $2.647.400, 5-jun) realmente se hizo desde
  // Bancolombia.
  "20042394": "BANCOL.CTE # 008-927404-01",
  // El 80816 (18-ago) es un traslado de Bancolombia hacia la Fiducia: SAP
  // registra en TransferAccount la cuenta contable de la Fiducia (12450505,
  // que ni siquiera está en TRANSFER_ACCOUNT_NAMES), pero el movimiento real
  // a conciliar es el débito que sale de Bancolombia.
  "80816": "BANCOL.CTE # 008-927404-01",
};

// Cuenta de compensación en Miami: opera en USD, a diferencia de las demás
// cuentas que operan en pesos colombianos.
export const MIAMI_ACCOUNT = "BANCO DE BOGOTA MIAMI CTA#79038";

// Palabras clave que en el campo 'Info.detallada' de SAP agrupan documentos
// que en conjunto financian varios movimientos bancarios del mismo tipo (ej.
// varios reembolsos de "CAJA MENOR" financian varios "retiros por cajero" el
// mismo día, sin correspondencia uno-a-uno).
export const GROUP_TOTAL_CATEGORIES: Record<string, string[]> = {
  "CAJA MENOR": ["CAJERO", "RETIRO"],
};

export const FILL_NO_MATCH = "FFFFC7CE"; // rojo suave
export const FILL_AMBIGUOUS = "FFFFEB9C"; // amarillo
export const FILL_MATCHED = "FFC6EFCE"; // verde suave (recién conciliado)

export const COMMENT_AUTHOR = "Conciliacion automatica";

/**
 * Mapa de cuenta contable SAP (TransferAccount) → nombre legible del banco.
 * Extraído del catálogo de cuentas (ChartOfAccounts) de SAP Business One.
 */
export const TRANSFER_ACCOUNT_NAMES: Record<string, string> = {
  "11100505": "BANCOL. AHORROS (SENA)",
  "11100510": "BANCOL. CTE #008-927404-01",
  "11100515": "BANCO SANTANDER CT#160-08814-9",
  "11100520": "BANCOL. AHOR CT #008-927404-02",
  "11100525": "BBVA COLOMBIA AHORROS",
  "11100530": "SUDAMERIS",
  "11100535": "BANCO DE BOGOTÁ #406007252",
  "11100540": "BBVA #10016129",
  "11100545": "AV VILLAS AHORROS #477-008106",
  "11100550": "AV VILLAS AHORROS #477-008889-0",
  "11100555": "BANCO CAJA SOCIAL CTE#21003425410",
  "11100560": "BANCO DAVIVIENDA CTE #3979-6999-8061",
  "11100565": "BANCO SANTANDER",
  "11101005": "BANCOLOMBIA CAYMAN",
  "11101010": "BANCO DE BOGOTÁ MIAMI CTA#79038",
  "11101015": "BANCO DE BOGOTÁ PANAMÁ",
};
