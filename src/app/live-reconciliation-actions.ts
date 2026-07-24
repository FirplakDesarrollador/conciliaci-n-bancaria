'use server';

import { createClient } from '@/lib/supabase/server';
import { convertApiToSapDocs } from '@/lib/conciliacion/sap';
import { reconcileDocs } from '@/lib/conciliacion/reconcile';
import { listDriveFiles, downloadDriveFile, uploadDriveFile } from '@/lib/graph/sharepoint';
import { ACCOUNT_MAP, TRANSFER_ACCOUNT_NAMES } from '@/lib/conciliacion/config';

const normalizeStr = (s: string) => 
  s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]/gi, "").toUpperCase();

export async function syncToSharepoint(sapPayments: any[], vendorPayments: any[]) {
  const supabase = await createClient();
  const { data: { user } } = await supabase.auth.getUser();

  if (!user) {
    throw new Error('Unauthorized');
  }

  try {
    // 1. Convertir a SapDoc[]
    const pool = convertApiToSapDocs(sapPayments, vendorPayments);

    // 2. Determinar qué archivos necesitamos según los bancos presentes en los pagos
    const bankMap = new Set<string>();
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
    await Promise.allSettled(Array.from(activeAccountKeys).map(async (accountKey) => {
      try {
        const configInfo = ACCOUNT_MAP[accountKey];
        const buffer = await downloadDriveFile(configInfo.file);
        bankBuffers.set(accountKey, buffer);
      } catch (e) {
        console.error(`Error downloading ${accountKey}:`, e);
      }
    }));

    if (bankBuffers.size === 0) {
      return { success: false, error: "No se pudo descargar ningún archivo de SharePoint." };
    }

    // 4. Ejecutar motor de conciliación
    const { results } = await reconcileDocs(pool, bankBuffers);

    // 5. Subir los archivos actualizados
    const updateResults: Record<string, { status: string, error?: string }> = {};

    await Promise.allSettled(results.map(async (res) => {
      try {
        await uploadDriveFile(res.archivo, res.workbookBuffer);
        updateResults[res.cuentaKey] = { status: "ACTUALIZADO SATISFACTORIAMENTE" };
      } catch (e: any) {
        console.error(`Error subiendo ${res.archivo}:`, e);
        updateResults[res.cuentaKey] = { status: "ERROR", error: e.message || String(e) };
      }
    }));

    return { success: true, updateResults };
  } catch (error: any) {
    console.error("Error en syncToSharepoint:", error);
    return { success: false, error: error.message || String(error) };
  }
}

function colToLetter(columnNumber: number): string {
  let temp, letter = '';
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  return letter;
}

export async function syncSingleToSharepointGraph(
  fileName: string,
  sheetName: string,
  row: number,
  col: number,
  docNumStr: string
) {
  const supabase = await createClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) {
    throw new Error('Unauthorized');
  }

  try {
    const { updateExcelCell } = await import('@/lib/graph/sharepoint');
    const cellAddress = `${colToLetter(col)}${row}`;
    await updateExcelCell(fileName, sheetName, cellAddress, docNumStr);
    return { success: true };
  } catch (error: any) {
    console.error("Error en syncSingleToSharepointGraph:", error);
    return { success: false, error: error.message || String(error) };
  }
}

export async function syncBulkToSharepointGraph(
  items: {
    fileName: string;
    sheetName: string;
    row: number;
    col: number;
    docNumStr: string;
  }[]
) {
  const supabase = await createClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) {
    throw new Error('Unauthorized');
  }

  const { updateExcelCell } = await import('@/lib/graph/sharepoint');
  const errors: string[] = [];

  for (const item of items) {
    try {
      const cellAddress = `${colToLetter(item.col)}${item.row}`;
      await updateExcelCell(item.fileName, item.sheetName, cellAddress, item.docNumStr);
    } catch (e: any) {
      errors.push(`Error en ${item.docNumStr}: ${e.message}`);
    }
  }

  if (errors.length > 0) {
    return { success: false, error: errors.join(" | ") };
  }
  return { success: true };
}

export async function generateExcelReport(
  sapPayments: any[],
  vendorPayments: any[],
  matchedDocNums: string[]
) {
  try {
    const ExcelJS = (await import('exceljs')).default;
    const wb = new ExcelJS.Workbook();
    
    const supabase = await createClient();
    const { data: history, error } = await supabase.from('reconciliation_history').select('*');
    if (error) throw error;

    const historySap = (history || []).filter(h => h.tipo === 'Recibido');
    const historyVendor = (history || []).filter(h => h.tipo === 'Efectuado');
    
    const matchedSet = new Set(matchedDocNums);
    (history || []).forEach(h => {
        h.doc_num.split('-').forEach((d: string) => matchedSet.add(d));
    });

    const setupSheet = (name: string, historyData: any[], uiData: any[]) => {
      const ws = wb.addWorksheet(name);
      
      ws.columns = [
        { header: 'Número', key: 'docNum', width: 20 },
        { header: 'Valor', key: 'valor', width: 20 },
        { header: 'Banco / Tercero', key: 'banco', width: 40 },
        { header: 'Fecha', key: 'fecha', width: 20 },
        { header: 'Descripción / Info', key: 'info', width: 50 },
      ];

      ws.getRow(1).font = { bold: true };

      // Add historical matched items
      historyData.forEach(h => {
        ws.addRow({
          docNum: h.doc_num,
          valor: h.valor,
          banco: h.banco,
          fecha: h.fecha,
          info: h.info
        });
      });

      // Add unmatched UI items (Yellow)
      uiData.forEach(p => {
        if (!matchedSet.has(String(p.DocNum))) {
          const val = p.TransferSum || p.CashSum || p.DocTotal || 0;
          const isUSD = p.DocCurrency === 'USD' && p.DocRate > 0;
          const finalVal = isUSD ? (val / p.DocRate) : val;

          const row = ws.addRow({
            docNum: p.DocNum || '',
            valor: finalVal,
            banco: p.TransferAccount || p.CashAccount || p.AcctName || p.CardName || '',
            fecha: p.DocDate ? new Date(p.DocDate).toISOString().split('T')[0] : '',
            info: p.Comments || p.Address || ''
          });

          row.eachCell({ includeEmpty: false }, (cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFF00' } // Yellow
            };
          });
        }
      });
    };

    setupSheet('Pagos Recibidos', historySap, sapPayments);
    setupSheet('Pagos Efectuados', historyVendor, vendorPayments);

    const buffer = await wb.xlsx.writeBuffer();
    return { 
      success: true, 
      base64: Buffer.from(buffer).toString('base64') 
    };
  } catch (error: any) {
    console.error("Error al generar reporte Excel:", error);
    return { success: false, error: error.message || String(error) };
  }
}

export async function logToHistory(items: any[]) {
  const supabase = await createClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return { success: false, error: 'Unauthorized' };

  try {
    const { error } = await supabase.from('reconciliation_history').insert(items);
    if (error) throw error;
    return { success: true };
  } catch (e: any) {
    console.error("Error logging history:", e);
    return { success: false, error: e.message };
  }
}
