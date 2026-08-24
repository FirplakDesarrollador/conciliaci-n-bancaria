import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });

import { getGraphAccessToken } from './src/lib/graph/token';
import { updateExcelCellsBatch } from './src/lib/graph/sharepoint';
import { listDriveFiles } from './src/lib/graph/sharepoint';

async function main() {
  const token = await getGraphAccessToken();
  const CONCILIACION_DRIVE_ID = "b!jk-4z9x8cku_nOTExW_ZfrUM6FX1P4VItm0n1QZDHWm81Nc17Wx_Tro_F7bRAfBV";
  const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
  const fileName = "CTA CTE BOGOTA 406007252 2026.xlsx";
  const CONCILIACION_FOLDER_PATH = "/TESORERIA/INFORMACION BANCARIA/BANCOS FIRPLAK S.A/MOVIMIENTOS Y CUADRES FIRPLAK/FIRPLAK 2026";
  
  const path = `${CONCILIACION_FOLDER_PATH}/${fileName}`;
  const encodedPath = path.split("/").map(encodeURIComponent).join("/");

  console.log("Buscando el archivo:", fileName);
  const fileRes = await fetch(`${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/root:${encodedPath}`, {
      headers: { Authorization: `Bearer ${token}` }
  });
  const fileData = await fileRes.json();
  const fileId = fileData.id;

  console.log("Obteniendo datos de la hoja AGOSTO...");
  const response = await fetch(`${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/items/${fileId}/workbook/worksheets/AGOSTO/usedRange`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const data = await response.json();
  const values = data.values;

  // Let's find the header row. Banco de Bogota header is usually row 1-5.
  let headerRowIdx = -1;
  let headers: any[] = [];
  for (let r = 0; r < 5; r++) {
      if (values[r] && values[r].some((c: any) => String(c).toUpperCase() === 'CREDITO' || String(c).toUpperCase() === 'CRÉDITO')) {
          headerRowIdx = r;
          headers = values[r];
          break;
      }
  }

  if (headerRowIdx === -1) {
      console.error("No se encontraron encabezados (CREDITO) en las primeras 5 filas.");
      return;
  }

  console.log(`Encabezados en fila ${headerRowIdx + 1}:`, headers);

  let docColIdx = headers.findIndex(h => {
      const val = String(h).toUpperCase();
      return val === "RC" || val === "DOC SAP" || val === "DOCUMENTO SAP";
  });

  if (docColIdx === -1) {
      console.log("No se encontró columna DOC SAP. Mostrando todos los encabezados...");
      headers.forEach((h, i) => console.log(`[${i}] ${h}`));
      // In Bogota format, the doc is often called "DOCUMENTO SAP" but could be named slightly different. 
      // We will fallback to searching. Let's wait.
  } else {
      console.log(`Columna SAP encontrada en índice ${docColIdx} (${String.fromCharCode(65 + docColIdx)})`);
  }

  let targetRowIndex = -1;
  let targetRow = null;

  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    // Check if the row has the value 8174925
    const hasValue = row.some((cell: any) => String(cell).includes("8174925"));

    if (hasValue) {
      console.log(`Encontrado en fila ${r + 1}:`);
      console.log(row);
      targetRowIndex = r + 1; // 1-based index for Excel
      targetRow = row;
      break;
    }
  }

  if (targetRowIndex > -1 && targetRow) {
      if (docColIdx === -1) {
          console.log("No doc col explicitly found. Using column A as fallback if it is empty, or wait for manual inspection.");
          return;
      }

      const docColLetter = String.fromCharCode(65 + docColIdx);
      const cellAddress = `${docColLetter}${targetRowIndex}`;
      
      console.log(`Actualizando celda ${cellAddress} con 80636...`);
      const res = await updateExcelCellsBatch(fileName, [
        {
          sheetName: "AGOSTO",
          cellAddress: cellAddress,
          value: "80636",
          color: "#C6EFCE"
        }
      ]);
      console.log("Resultado de actualizacion:", res);
  } else {
      console.log("No se encontro la fila con el valor 8174925.");
  }
}

main().catch(console.error);
