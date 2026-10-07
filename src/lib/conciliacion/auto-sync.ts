import "server-only";
import { convertApiToSapDocs } from "./sap";
import { reconcileDocs } from "./reconcile";
import { listDriveFiles, downloadDriveFile, updateExcelCellsBatch } from "@/lib/graph/sharepoint";
import { ACCOUNT_MAP, TRANSFER_ACCOUNT_NAMES } from "./config";

const normalizeStr = (s: string) =>
  s.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Z0-9]/gi, "").toUpperCase();

function colToLetter(columnNumber: number): string {
  let temp: number;
  let letter = "";
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  return letter;
}

/**
 * Motor completo de sincronizacion SAP -> Excel de bancos: descarga los
 * archivos de SharePoint, corre reconcileDocs (todas las pasadas de
 * matching) y sube de vuelta los que cambiaron. No valida sesion de
 * usuario: quien llama (la Server Action del boton "Sincronizar" o el cron
 * de auto-sync) es responsable de esa verificacion antes de invocarla.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function runFullSync(sapPayments: any[], vendorPayments: any[]) {
  try {
    // 1. Convertir a SapDoc[]
    const pool = convertApiToSapDocs(sapPayments, vendorPayments);

    // 2. Determinar que archivos necesitamos segun los bancos presentes en los pagos
    const bankMap = new Set<string>();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const processPayment = (payment: any) => {
      const acc = payment.TransferAccount || payment.CashAccount;
      if (!acc) return;
      const bankName = TRANSFER_ACCOUNT_NAMES[acc] || acc;
      bankMap.add(bankName);
    };

    sapPayments.forEach(processPayment);
    vendorPayments.forEach(processPayment);

    let filesInSharepoint: string[] = [];
    try {
      filesInSharepoint = await listDriveFiles();
    } catch (error) {
      console.error("Error fetching sharepoint files:", error);
      throw new Error("No se pudieron obtener los archivos de SharePoint.");
    }

    const filesToDownload = new Set<string>();
    const activeAccountKeys = new Set<string>();

    Array.from(bankMap).forEach((bank) => {
      const accountEntry = Object.entries(ACCOUNT_MAP).find(([key]) =>
        normalizeStr(key) === normalizeStr(bank) ||
        normalizeStr(bank).includes(normalizeStr(key)) ||
        normalizeStr(key).includes(normalizeStr(bank))
      );
      if (accountEntry) {
        const [accountKey, configInfo] = accountEntry;
        if (filesInSharepoint.includes(configInfo.file)) {
          filesToDownload.add(configInfo.file);
          activeAccountKeys.add(accountKey);
        }
      }
    });

    if (filesToDownload.size === 0) {
      return { success: false, error: `No se encontraron archivos de banco. banks=${Array.from(bankMap).join(', ')}. files=${filesInSharepoint.join(', ')}. sapLen=${sapPayments.length}, vendorLen=${vendorPayments.length}` };
    }

    // 3. Descargar buffers
    const bankBuffers = new Map<string, Buffer>();
    const downloadErrors: Record<string, string> = {};
    await Promise.allSettled(Array.from(activeAccountKeys).map(async (accountKey) => {
      const configInfo = ACCOUNT_MAP[accountKey];
      // SharePoint responde 503/429 de forma transitoria: se reintenta antes
      // de dar el banco por perdido.
      for (let attempt = 1; attempt <= 4; attempt++) {
        try {
          bankBuffers.set(accountKey, await downloadDriveFile(configInfo.file));
          return;
        } catch (e) {
          console.error(`Error downloading ${accountKey} (intento ${attempt}):`, e);
          downloadErrors[accountKey] = e instanceof Error ? e.message : String(e);
          if (attempt < 4) await new Promise((r) => setTimeout(r, 3000 * attempt));
        }
      }
    }));

    if (bankBuffers.size === 0) {
      return { success: false, error: "No se pudo descargar ningun archivo de SharePoint." };
    }

    // 4. Ejecutar motor de conciliacion
    const { results } = await reconcileDocs(pool, bankBuffers);

    // 5. Aplicar SOLO las celdas que cambiaron, con PATCH quirurgico
    // (updateExcelCellsBatch), en vez de subir el workbook completo. Subir
    // el buffer entero pisaria cualquier cambio hecho al archivo real entre
    // la descarga (paso 3) y este punto -- una correccion manual de alguien
    // en Excel, u otra corrida de sincronizacion concurrente -- sin dejar
    // rastro de que se perdio. Ver commit del 2026-09-29 (81399 -> 81993).
    const updateResults: Record<string, { status: string, error?: string }> = {};

    await Promise.allSettled(results.map(async (res) => {
      if (res.writes.length === 0) {
        updateResults[res.cuentaKey] = { status: "SIN CAMBIOS" };
        return;
      }
      try {
        const updates = res.writes.map((w) => ({
          sheetName: w.sheet,
          cellAddress: `${colToLetter(w.col)}${w.row}`,
          value: w.value,
          color: "#C6EFCE",
        }));
        const result = await updateExcelCellsBatch(res.archivo, updates);
        if (result.success) {
          updateResults[res.cuentaKey] = { status: "ACTUALIZADO SATISFACTORIAMENTE" };
        } else {
          updateResults[res.cuentaKey] = { status: "ERROR", error: result.errors.join(" | ") };
        }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (e: any) {
        console.error(`Error actualizando celdas de ${res.archivo}:`, e);
        updateResults[res.cuentaKey] = { status: "ERROR", error: e.message || String(e) };
      }
    }));

    // Un banco cuyo archivo no se pudo bajar NO se procesó: se informa en vez
    // de dejarlo pasar como un exito silencioso.
    for (const [accountKey, msg] of Object.entries(downloadErrors)) {
      if (!bankBuffers.has(accountKey)) updateResults[accountKey] = { status: "ERROR", error: `No se pudo descargar el archivo del banco: ${msg}` };
    }
    const failed = Object.values(updateResults).some((r) => r.status === "ERROR");
    return { success: !failed, updateResults };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    console.error("Error en runFullSync:", error);
    return { success: false, error: error.message || String(error) };
  }
}
