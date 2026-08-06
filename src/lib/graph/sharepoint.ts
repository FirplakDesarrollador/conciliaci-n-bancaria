import "server-only";
import { getGraphAccessToken } from "./token";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

// Sitio "FPK Tesoreria y Cartera", carpeta "FIRPLAK 2026" (ya resueltos vía
// el share link que compartió el usuario). Ver plan en
// C:\Users\claudia.duque\.claude\plans\cheeky-mapping-lecun.md
export const CONCILIACION_DRIVE_ID =
  "b!jk-4z9x8cku_nOTExW_ZfrUM6FX1P4VItm0n1QZDHWm81Nc17Wx_Tro_F7bRAfBV";
export const CONCILIACION_FOLDER_PATH =
  "/TESORERIA/INFORMACION BANCARIA/BANCOS FIRPLAK S.A/MOVIMIENTOS Y CUADRES FIRPLAK/FIRPLAK 2026";

function encodeDrivePath(path: string): string {
  return path
    .split("/")
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

export async function downloadDriveFile(fileName: string): Promise<Buffer> {
  const token = await getGraphAccessToken();
  const path = `${CONCILIACION_FOLDER_PATH}/${fileName}`;
  const res = await fetch(
    `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/root:${encodeDrivePath(path)}:/content`,
    { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
  );

  if (!res.ok) {
    throw new Error(`No se pudo descargar "${fileName}" de SharePoint (HTTP ${res.status})`);
  }

  return Buffer.from(await res.arrayBuffer());
}

export async function uploadDriveFile(fileName: string, content: Buffer): Promise<{ webUrl: string }> {
  const token = await getGraphAccessToken();
  const path = `${CONCILIACION_FOLDER_PATH}/${fileName}`;
  const res = await fetch(
    `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/root:${encodeDrivePath(path)}:/content`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      // exceljs augmenta el tipo global `Buffer` de forma incompatible con
      // los tipos de `fetch`/@types/node actuales; el cast evita ese choque.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      body: content as any,
    }
  );

  if (!res.ok) {
    const body = await res.json().catch(() => null);
    throw new Error(
      `No se pudo subir "${fileName}" a SharePoint (HTTP ${res.status}): ${body?.error?.message ?? ""}`
    );
  }

  const data = (await res.json()) as { webUrl: string };
  return { webUrl: data.webUrl };
}

export async function listDriveFiles(): Promise<string[]> {
  const token = await getGraphAccessToken();
  const path = encodeDrivePath(CONCILIACION_FOLDER_PATH);
  const res = await fetch(
    `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/root:${path}:/children`,
    { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
  );

  if (!res.ok) {
    console.error("No se pudo listar los archivos de SharePoint.");
    return [];
  }

  const data = await res.json();
  return data.value.map((item: any) => item.name);
}

export async function updateExcelCell(
  fileName: string,
  sheetName: string,
  cellAddress: string,
  value: string | number
): Promise<void> {
  const token = await getGraphAccessToken();
  const path = `${CONCILIACION_FOLDER_PATH}/${fileName}`;
  
  // 1. Obtener el ID del archivo
  const fileRes = await fetch(
    `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/root:${encodeDrivePath(path)}`,
    { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
  );
  if (!fileRes.ok) {
    throw new Error(`No se pudo obtener el ID del archivo ${fileName} (HTTP ${fileRes.status})`);
  }
  const fileData = await fileRes.json();
  const fileId = fileData.id;

  // 2. Actualizar la celda vía Excel Graph API
  // sheetName is wrapped in single quotes for the Graph API syntax
  const updateRes = await fetch(
    `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/items/${fileId}/workbook/worksheets('${sheetName}')/range(address='${cellAddress}')`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        values: [[value]],
      }),
    }
  );

  if (!updateRes.ok) {
    const errorBody = await updateRes.text();
    throw new Error(`Error actualizando celda en SharePoint: ${updateRes.status} - ${errorBody}`);
  }
}

export async function updateExcelCellsBatch(
  fileName: string,
  updates: { sheetName: string; cellAddress: string; value: string | number; color: string }[]
): Promise<{ success: boolean; errors: string[] }> {
  const token = await getGraphAccessToken();
  const path = `${CONCILIACION_FOLDER_PATH}/${fileName}`;
  
  const fileRes = await fetch(
    `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/root:${encodeDrivePath(path)}`,
    { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
  );
  if (!fileRes.ok) {
    throw new Error(`No se pudo obtener el ID del archivo ${fileName} (HTTP ${fileRes.status})`);
  }
  const fileData = await fileRes.json();
  const fileId = fileData.id;

  // 1. Create a Workbook Session to prevent lock errors
  let sessionId = null;
  try {
    const sessionRes = await fetch(
      `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/items/${fileId}/workbook/createSession`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ persistChanges: true })
      }
    );
    if (sessionRes.ok) {
      const sessionData = await sessionRes.json();
      sessionId = sessionData.id;
    } else {
      console.warn(`No se pudo crear sesión de Workbook (HTTP ${sessionRes.status}). Continuará sin sesión exclusiva.`);
    }
  } catch (e) {
    console.warn("Fallo en creación de sesión de Workbook", e);
  }

  const requests: any[] = [];
  let reqId = 1;

  for (const update of updates) {
    const rangeUrl = `/drives/${CONCILIACION_DRIVE_ID}/items/${fileId}/workbook/worksheets('${update.sheetName}')/range(address='${update.cellAddress}')`;
    
    // Request to update the value
    requests.push({
      id: String(reqId++),
      method: "PATCH",
      url: rangeUrl,
      body: { values: [[update.value]] },
      headers: { 
        "Content-Type": "application/json",
        ...(sessionId ? { "workbook-session-id": sessionId } : {})
      }
    });
    
    // Request to update the background color
    requests.push({
      id: String(reqId++),
      method: "PATCH",
      url: `${rangeUrl}/format/fill`,
      body: { color: update.color },
      headers: { 
        "Content-Type": "application/json",
        ...(sessionId ? { "workbook-session-id": sessionId } : {})
      }
    });
  }

  const batchErrors: string[] = [];

  // Microsoft Graph $batch API has a limit of 20 requests per batch.
  for (let i = 0; i < requests.length; i += 20) {
    const batch = requests.slice(i, i + 20);
    
    let attempt = 0;
    const maxAttempts = 3;
    let batchSuccess = false;

    while (attempt < maxAttempts && !batchSuccess) {
      try {
        const batchRes = await fetch(`${GRAPH_BASE_URL}/$batch`, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ requests: batch })
        });
        
        if (!batchRes.ok) {
          if (batchRes.status === 429 || batchRes.status >= 500) {
            throw new Error(`Throttling/Server Error: ${batchRes.status}`);
          }
          batchErrors.push(`Error en batch update a SharePoint: ${batchRes.status} - ${await batchRes.text()}`);
          break; // Don't retry on 400 Bad Request, etc.
        }
        
        const batchData = await batchRes.json();
        let hasThrottledItem = false;

        for (const res of batchData.responses || []) {
          if (res.status === 429 || res.status >= 500) {
             hasThrottledItem = true;
          } else if (res.status >= 400) {
            // Find which original request this was to give better error info
            const req = batch.find(r => r.id === res.id);
            batchErrors.push(`Fallo al actualizar celda (URL: ${req?.url || 'desconocida'}, Status ${res.status}): ${JSON.stringify(res.body)}`);
          }
        }

        if (hasThrottledItem) {
           throw new Error("Throttling in individual batch response items");
        }

        batchSuccess = true;
      } catch (e: any) {
        attempt++;
        if (attempt >= maxAttempts) {
          batchErrors.push(`Fallo persistente tras ${maxAttempts} intentos al procesar lote: ${e.message}`);
        } else {
          // Exponential backoff: 2s, 4s, 8s
          const backoff = Math.pow(2, attempt) * 1000;
          await sleep(backoff);
        }
      }
    }
    
    // Add a small delay between successful batches to prevent locking issues
    if (batchSuccess && i + 20 < requests.length) {
      await sleep(1000); 
    }
  }

  // 3. Close the Workbook Session
  if (sessionId) {
    try {
      await fetch(
        `${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/items/${fileId}/workbook/closeSession`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "workbook-session-id": sessionId
          }
        }
      );
    } catch (e) {
      console.warn("Fallo cerrando la sesión de Workbook", e);
    }
  }
  
  if (batchErrors.length > 0) {
    return { success: false, errors: batchErrors };
  }
  
  return { success: true, errors: [] };
}

