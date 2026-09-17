'use server';

import { createClient } from '@/lib/supabase/server';
import { runFullSync } from '@/lib/conciliacion/auto-sync';

export async function syncToSharepoint(sapPayments: any[], vendorPayments: any[]) {
  const supabase = await createClient();
  const { data: { user } } = await supabase.auth.getUser();

  if (!user) {
    throw new Error('Unauthorized');
  }

  return runFullSync(sapPayments, vendorPayments);
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
  return syncBulkToSharepointGraph([{ fileName, sheetName, row, col, docNumStr }]);
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

  const { updateExcelCellsBatch } = await import('@/lib/graph/sharepoint');

  const errors: string[] = [];
  const filesMap = new Map<string, typeof items>();
  for (const item of items) {
    if (!filesMap.has(item.fileName)) {
      filesMap.set(item.fileName, []);
    }
    filesMap.get(item.fileName)!.push(item);
  }

  for (const [fileName, fileItems] of filesMap.entries()) {
    try {
      const updates = fileItems.map(item => ({
        sheetName: item.sheetName,
        cellAddress: `${colToLetter(item.col)}${item.row}`,
        value: item.docNumStr,
        color: "#C6EFCE" // FILL_MATCHED (verde suave)
      }));
      const res = await updateExcelCellsBatch(fileName, updates);
      if (!res.success && res.errors) {
        errors.push(`Errores en archivo ${fileName}: ${res.errors.join(' | ')}`);
      }
    } catch (e: any) {
      errors.push(`Excepción procesando archivo ${fileName}: ${e.message}`);
    }
  }

  if (errors.length > 0) {
    return { success: false, error: errors.join("\n\n") };
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

    const now = new Date();
    const yesterday = new Date(now);
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);

    const { data: history, error } = await supabase
      .from('reconciliation_history')
      .select('*')
      .gte('created_at', yesterday.toISOString());

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
