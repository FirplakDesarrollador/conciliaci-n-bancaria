'use client';

import { useState } from 'react';
import { syncToSharepoint, syncSingleToSharepointGraph } from './live-reconciliation-actions';

interface BankDocs {
  isCombo: boolean;
  docs: any[];
  value: number;
  isMatch: boolean;
  date: Date;
  isUSD: boolean;
  isFilled: boolean;
  excelFile?: string;
  sheetName?: string;
  excelRow?: number;
  excelCol?: number;
}

interface BankItem {
  bank: string;
  docs: BankDocs[];
  status: string;
  statusColor: string;
}

interface LiveReconciliationClientProps {
  initialBankList: BankItem[];
  sapPayments: any[];
  vendorPayments: any[];
}

export default function LiveReconciliationClient({
  initialBankList,
  sapPayments,
  vendorPayments
}: LiveReconciliationClientProps) {
  const [bankList, setBankList] = useState<BankItem[]>(initialBankList);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncingItems, setSyncingItems] = useState<Set<string>>(new Set());
  const [syncResult, setSyncResult] = useState<{ status: 'idle' | 'syncing' | 'success' | 'error', message?: string }>({ status: 'idle' });
  const [isDownloading, setIsDownloading] = useState(false);

  const handleSyncSingle = async (bankItem: BankItem, itemDocs: BankDocs) => {
    const docNumsStr = itemDocs.docs.map((d: any) => d.docNum).join("-");
    const itemKey = `${bankItem.bank}-${docNumsStr}`;
    
    setSyncingItems(prev => {
      const next = new Set(prev);
      next.add(itemKey);
      return next;
    });

    try {
      if (!itemDocs.excelFile || !itemDocs.sheetName || !itemDocs.excelRow || !itemDocs.excelCol) {
        alert("Faltan datos de ubicación en SharePoint para este registro.");
        return;
      }
      
      const result = await syncSingleToSharepointGraph(
        itemDocs.excelFile,
        itemDocs.sheetName,
        itemDocs.excelRow,
        itemDocs.excelCol,
        docNumsStr
      );
      
      if (!result.success) {
        alert(`Error al sincronizar: ${result.error}`);
      } else {
        // Log to history
        const historyItems: any[] = [];
        itemDocs.docs.forEach((d: any) => {
           const sapP = sapPayments.find(p => String(p.DocNum) === String(d.docNum));
           const venP = vendorPayments.find(p => String(p.DocNum) === String(d.docNum));
           const p = sapP || venP;
           if (p) {
              const val = p.TransferSum || p.CashSum || p.DocTotal || 0;
              const isUSD = p.DocCurrency === 'USD' && p.DocRate > 0;
              const finalVal = isUSD ? (val / p.DocRate) : val;
              historyItems.push({
                 doc_num: String(p.DocNum),
                 valor: finalVal,
                 banco: p.TransferAccount || p.CashAccount || p.AcctName || p.CardName || bankItem.bank,
                 fecha: p.DocDate ? new Date(p.DocDate).toISOString().split('T')[0] : '',
                 info: p.Comments || p.Address || '',
                 tipo: sapP ? 'Recibido' : 'Efectuado'
              });
           }
        });
        if (historyItems.length > 0) {
           const { logToHistory } = await import('./live-reconciliation-actions');
           await logToHistory(historyItems);
        }

        setBankList(prev => prev.map(item => {
          if (item.bank === bankItem.bank) {
            return {
              ...item,
              docs: item.docs.map(d => 
                (d.docs.map((x: any) => x.docNum).join("-") === docNumsStr) ? { ...d, isFilled: true } : d
              )
            };
          }
          return item;
        }));
      }
    } catch (e: any) {
      alert(`Error al sincronizar: ${e.message || String(e)}`);
    } finally {
      setSyncingItems(prev => {
        const next = new Set(prev);
        next.delete(itemKey);
        return next;
      });
    }
  };

  const handleDownloadReport = async () => {
    setIsDownloading(true);
    setSyncResult({ status: 'idle' });
    try {
      // Recolectar todos los docNums que tienen match
      const matchedSet = new Set<string>();
      bankList.forEach(bankItem => {
        bankItem.docs.forEach(docGroup => {
          if (docGroup.isMatch) {
            docGroup.docs.forEach((d: any) => matchedSet.add(String(d.docNum)));
          }
        });
      });

      const { generateExcelReport } = await import('./live-reconciliation-actions');
      const result = await generateExcelReport(sapPayments, vendorPayments, Array.from(matchedSet));
      
      if (!result.success || !result.base64) {
        setSyncError(result.error || "Ocurrió un error al generar el reporte.");
        return;
      }

      // Descargar el archivo
      const byteCharacters = atob(result.base64);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Reporte_Conciliacion_${new Date().toISOString().split('T')[0]}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

    } catch (error: any) {
      setSyncError(error.message || String(error));
    } finally {
      setIsDownloading(false);
    }
  };

  const handleSync = async () => {
    setIsSyncing(true);
    setSyncResult({ status: 'syncing' });

    const itemsToSync: {
       fileName: string;
       sheetName: string;
       row: number;
       col: number;
       docNumStr: string;
       bank: string;
    }[] = [];

    bankList.forEach(bankItem => {
       bankItem.docs.forEach(docGroup => {
          if (docGroup.isMatch && !docGroup.isFilled && docGroup.excelFile && docGroup.sheetName && docGroup.excelRow && docGroup.excelCol) {
             itemsToSync.push({
                fileName: docGroup.excelFile,
                sheetName: docGroup.sheetName,
                row: docGroup.excelRow,
                col: docGroup.excelCol,
                docNumStr: docGroup.docs.map((d: any) => d.docNum).join("-"),
                bank: bankItem.bank
             });
          }
       });
    });

    if (itemsToSync.length === 0) {
       setIsSyncing(false);
       setSyncResult({ status: 'idle' });
       alert("No hay documentos válidos pendientes para sincronizar.");
       return;
    }

    setBankList(prev => prev.map(item => ({
       ...item,
       status: itemsToSync.some(i => i.bank === item.bank) ? "SINCRONIZANDO..." : item.status,
       statusColor: itemsToSync.some(i => i.bank === item.bank) ? "bg-blue-100 text-blue-800 border-blue-200" : item.statusColor
    })));

    try {
      const { syncBulkToSharepointGraph, logToHistory } = await import('./live-reconciliation-actions');
      const result = await syncBulkToSharepointGraph(itemsToSync);
      
      if (!result.success) {
        setSyncResult({ status: 'error', message: result.error || "Ocurrió un error al sincronizar algunos documentos." });
      } else {
        setSyncResult({ status: 'success', message: '¡Sincronización completada exitosamente!' });
        // Log to history for all successfully matched items
        const historyItems: any[] = [];
        itemsToSync.forEach(syncItem => {
           const docNums = syncItem.docNumStr.split('-');
           docNums.forEach(num => {
              const sapP = sapPayments.find(p => String(p.DocNum) === num);
              const venP = vendorPayments.find(p => String(p.DocNum) === num);
              const p = sapP || venP;
              if (p) {
                 const val = p.TransferSum || p.CashSum || p.DocTotal || 0;
                 const isUSD = p.DocCurrency === 'USD' && p.DocRate > 0;
                 const finalVal = isUSD ? (val / p.DocRate) : val;
                 historyItems.push({
                    doc_num: String(p.DocNum),
                    valor: finalVal,
                    banco: p.TransferAccount || p.CashAccount || p.AcctName || p.CardName || syncItem.bank,
                    fecha: p.DocDate ? new Date(p.DocDate).toISOString().split('T')[0] : '',
                    info: p.Comments || p.Address || '',
                    tipo: sapP ? 'Recibido' : 'Efectuado'
                 });
              }
           });
        });
        if (historyItems.length > 0) {
           await logToHistory(historyItems);
        }
      }
      
      setBankList(prev => prev.map(item => {
        const hadUpdates = itemsToSync.some(i => i.bank === item.bank);
        if (!hadUpdates) return item;

        return {
          ...item,
          status: result.success ? "ACTUALIZADO SATISFACTORIAMENTE" : "ERROR PARCIAL",
          statusColor: result.success ? "bg-emerald-600 text-white border-emerald-700" : "bg-amber-100 text-amber-800 border-amber-200",
          docs: item.docs.map(docGroup => {
            const docNumStr = docGroup.docs.map((d: any) => d.docNum).join("-");
            const wasSynced = itemsToSync.some(i => i.bank === item.bank && i.docNumStr === docNumStr);
            if (wasSynced && result.success) {
               return { ...docGroup, isFilled: true };
            }
            return docGroup;
          })
        };
      }));

    } catch (e: any) {
      setSyncError(e.message || String(e));
      setBankList(initialBankList);
    } finally {
      setIsSyncing(false);
    }
  };

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm overflow-hidden flex flex-col mt-6">
      <div className="mb-4 flex items-center justify-between">
        <div>
          <h3 className="flex items-center gap-2 text-lg font-bold text-slate-900">
            Resumen de Bancos y Documentos
          </h3>
          <p className="text-sm text-slate-500">
            Validación de valores en SAP vs SharePoint por fecha y monto.
          </p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={handleDownloadReport}
            disabled={isDownloading || (sapPayments.length === 0 && vendorPayments.length === 0)}
            className="flex items-center gap-2 rounded-lg bg-white border border-slate-300 px-4 py-2 text-sm font-medium text-slate-700 transition-colors hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {isDownloading ? (
              <svg className="animate-spin h-4 w-4 text-slate-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
            ) : (
              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            )}
            Descargar Informe
          </button>
          
          <button 
            onClick={handleSync}
            disabled={isSyncing || bankList.length === 0}
            className="flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-2 text-sm font-medium text-white transition-colors hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {isSyncing ? (
              <>
                <svg className="animate-spin h-4 w-4 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
                Enviando a SharePoint...
              </>
            ) : (
              <>
                <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>
                Enviar a SharePoint
              </>
            )}
          </button>
        </div>
      </div>

      {syncResult.status === 'syncing' && (
        <div className="mb-4 rounded-xl border border-blue-200 bg-blue-50 p-5 shadow-sm">
          <div className="flex justify-between items-center mb-2">
            <span className="text-sm font-bold text-blue-800 flex items-center gap-2">
              <svg className="animate-spin h-4 w-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg>
              Sincronizando con SharePoint...
            </span>
            <span className="text-xs font-medium text-blue-600 bg-blue-100 px-2 py-1 rounded-full">Procesando en segundo plano</span>
          </div>
          <p className="text-xs text-blue-700 mb-3 opacity-80">Puedes continuar navegando o realizando otras tareas. Te avisaremos cuando finalice.</p>
          <div className="w-full bg-blue-200/50 rounded-full h-2 overflow-hidden">
            <div className="bg-blue-600 h-2 rounded-full animate-pulse w-full origin-left transform scale-x-100 transition-transform duration-1000"></div>
          </div>
        </div>
      )}

      {syncResult.status === 'error' && (
        <div className="mb-4 rounded-xl border border-red-200 bg-red-50 p-5 shadow-sm">
          <h3 className="text-sm font-bold text-red-800 flex items-center gap-2 mb-2">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" /></svg>
            Sincronización finalizada con errores
          </h3>
          <div className="text-xs text-red-700 max-h-40 overflow-y-auto pr-2 custom-scrollbar">
            <ul className="list-disc pl-4 space-y-1">
              {syncResult.message?.split('\n\n').map((msg, i) => (
                <li key={i} className="whitespace-pre-wrap">{msg}</li>
              ))}
            </ul>
          </div>
        </div>
      )}

      {syncResult.status === 'success' && (
        <div className="mb-4 rounded-xl border border-emerald-200 bg-emerald-50 p-4 shadow-sm flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-8 w-8 items-center justify-center rounded-full bg-emerald-100 text-emerald-600">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" /></svg>
            </div>
            <div>
              <h3 className="text-sm font-bold text-emerald-800">¡Sincronización Exitosa!</h3>
              <p className="text-xs text-emerald-600">{syncResult.message}</p>
            </div>
          </div>
          <button onClick={() => setSyncResult({ status: 'idle' })} className="text-emerald-500 hover:text-emerald-700 p-1">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" /></svg>
          </button>
        </div>
      )}
      
      {bankList.length > 0 ? (
        <div className="grid gap-4">
          {bankList.map(({ bank, docs, status, statusColor }) => (
            <div key={bank} className="rounded-xl border border-slate-100 bg-slate-50 p-4">
              <div className="flex items-center gap-3 mb-3">
                <h4 className="font-semibold text-slate-800">{bank}</h4>
                <span className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-semibold border ${statusColor}`}>
                  {status}
                </span>
              </div>
              <div className="flex flex-col gap-2">
                {docs.map((item, i) => {
                  const isUsd = item.isUSD;
                  const formattedValue = new Intl.NumberFormat(isUsd ? 'en-US' : 'es-CO', { 
                    style: 'currency', 
                    currency: isUsd ? 'USD' : 'COP',
                    minimumFractionDigits: isUsd ? 2 : 0,
                    maximumFractionDigits: isUsd ? 2 : 0
                  }).format(item.value);

                  const docNums = item.docs.map((d: any) => d.docNum).join(" + ");
                  const displayDate = item.docs.map((d: any) => new Date(d.date).toISOString().split('T')[0]).filter((v: string, idx: number, a: string[]) => a.indexOf(v) === idx).join(" | ");

                  return (
                    <div key={i} className={`flex items-center justify-between rounded-md px-3 py-2 text-sm border ${item.isMatch ? 'bg-emerald-50 border-emerald-200' : 'bg-white border-slate-200 shadow-sm'}`}>
                      <div className="font-medium text-slate-700">
                        # {docNums} <span className="mx-1.5 text-slate-300">|</span> <span className="text-slate-600">{formattedValue}</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="text-xs text-slate-400">{displayDate}</span>
                        {(status === "ENCONTRADO EN SHAREPOINT" || status === "ACTUALIZADO SATISFACTORIAMENTE" || status === "SINCRONIZANDO...") && (
                          item.isMatch ? (
                            <span className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-700 bg-emerald-100 px-2 py-0.5 rounded-full">
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-3.5 w-3.5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" /></svg>
                              Coincide
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 text-xs font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full border border-slate-200">
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-3.5 w-3.5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" /></svg>
                              Sin match individual
                            </span>
                          )
                        )}
                        {item.isFilled ? (
                          <span className="ml-2 inline-flex items-center gap-1 text-xs font-medium text-slate-500 bg-slate-100 px-2 py-1 rounded-md border border-slate-200">
                            <svg xmlns="http://www.w3.org/2000/svg" className="h-3.5 w-3.5 text-slate-400" viewBox="0 0 20 20" fill="currentColor">
                              <path fillRule="evenodd" d="M5 9V7a5 5 0 0110 0v2a2 2 0 012 2v5a2 2 0 01-2 2H5a2 2 0 01-2-2v-5a2 2 0 012-2zm8-2v2H7V7a3 3 0 016 0z" clipRule="evenodd" />
                            </svg>
                            Ya lleno
                          </span>
                        ) : (
                          <button
                            onClick={() => handleSyncSingle({bank, docs, status, statusColor}, item)}
                            disabled={syncingItems.has(`${bank}-${item.docs.map((d: any) => d.docNum).join("-")}`) || isSyncing}
                            className="ml-2 flex items-center gap-1 rounded-md bg-emerald-50 px-2 py-1 text-xs font-medium text-emerald-700 transition-colors hover:bg-emerald-100 border border-emerald-200 disabled:opacity-50 disabled:cursor-not-allowed"
                            title="Subir este dato a SharePoint"
                          >
                            {syncingItems.has(`${bank}-${item.docs.map((d: any) => d.docNum).join("-")}`) ? (
                               <svg className="animate-spin h-3.5 w-3.5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg>
                            ) : (
                               <svg xmlns="http://www.w3.org/2000/svg" className="h-3.5 w-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>
                            )}
                            Subir a SharePoint
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      ) : (
        <div className="text-center py-8 text-slate-500">
          No se encontraron bancos en las listas de SAP.
        </div>
      )}
    </div>
  );
}
