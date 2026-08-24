import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });

import { getGraphAccessToken } from './src/lib/graph/token';
import { updateExcelCellsBatch } from './src/lib/graph/sharepoint';

async function main() {
  const token = await getGraphAccessToken();
  const fileId = "01I4CEXKHUEDQU3QGI5BG2ZXNCNZGH3BV2";
  const CONCILIACION_DRIVE_ID = "b!jk-4z9x8cku_nOTExW_ZfrUM6FX1P4VItm0n1QZDHWm81Nc17Wx_Tro_F7bRAfBV";
  const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

  const response = await fetch(`${GRAPH_BASE_URL}/drives/${CONCILIACION_DRIVE_ID}/items/${fileId}/workbook/worksheets/AGOSTO/usedRange`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const data = await response.json();
  const values = data.values;

  let docColIdx = -1;
  
  for (let r = 0; r < 10; r++) {
      const row = values[r];
      if (!row) continue;
      const idx = row.findIndex((c: any) => {
          const val = String(c).toUpperCase();
          return val === "RC" || val === "DOC SAP" || val === "DOCUMENTO SAP";
      });
      if (idx !== -1) {
          docColIdx = idx;
          console.log(`Cabecera encontrada en fila ${r+1}, columna ${idx} (${String.fromCharCode(65 + idx)})`);
          break;
      }
  }

  // Restore the balance value
  const resRestore = await updateExcelCellsBatch("CTA 07 CTE 00892740401 BANCOLOMBIA 2026.xlsx", [
    {
      sheetName: "AGOSTO",
      cellAddress: "G148",
      value: "64062388.81",
      color: "#FFFFFF"
    }
  ]);
  console.log("Restoration of G148 result:", resRestore);

  if (docColIdx !== -1) {
      const docColLetter = String.fromCharCode(65 + docColIdx);
      console.log(`Writing 80637 to ${docColLetter}148...`);
      const resUpdate = await updateExcelCellsBatch("CTA 07 CTE 00892740401 BANCOLOMBIA 2026.xlsx", [
        {
          sheetName: "AGOSTO",
          cellAddress: `${docColLetter}148`,
          value: "80637",
          color: "#C6EFCE"
        }
      ]);
      console.log("Update result:", resUpdate);
  }
}

main().catch(console.error);
