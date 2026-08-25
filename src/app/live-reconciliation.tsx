import { TRANSFER_ACCOUNT_NAMES, ACCOUNT_MAP, MANUAL_CUENTA_OVERRIDES } from "@/lib/conciliacion/config";
import { listDriveFiles, downloadDriveFile } from "@/lib/graph/sharepoint";
import ExcelJS from "exceljs";
import { findHeaderRowAndCols, READERS, REQUIRED_HEADERS_HINT, fixCompensacionDates } from "@/lib/conciliacion/readers";
import LiveReconciliationClient from "./live-reconciliation-client";

export default async function LiveReconciliation({
  sapPayments,
  vendorPayments,
}: {
  sapPayments: any[];
  vendorPayments: any[];
}) {
  // Almacenamos el DocNum, el valor, fecha, si es en USD, y tipo (IN/OUT)
  const bankMap = new Map<string, { docNum: number, value: number, isUSD: boolean, date: Date, tipo: "IN" | "OUT" }[]>();

  const processPayment = (payment: any, isIncoming: boolean) => {
    if (!payment.DocNum) return;

    const isGlobalUSD = payment.DocCurrency === 'USD' && payment.DocRate > 0;

    // Función auxiliar para agregar el leg (pata) del movimiento
    const addLeg = (accountCode: string, localVal: number, fcVal: number | null, legTipo: "IN" | "OUT", bankNameOverride?: string) => {
      const bankName = bankNameOverride || TRANSFER_ACCOUNT_NAMES[accountCode] || accountCode;
      if (!bankMap.has(bankName)) {
        bankMap.set(bankName, []);
      }
      
      const isMiami = bankName.includes("MIAMI");
      
      let finalVal = localVal;
      let finalIsUSD = false;
      
      if (isMiami) {
        if (fcVal && fcVal > 0) {
           finalVal = fcVal;
           finalIsUSD = true;
        } else if (isGlobalUSD) {
           finalVal = localVal / payment.DocRate;
           finalIsUSD = true;
        }
      }
      
      if (finalVal > 0) {
        bankMap.get(bankName)!.push({
          docNum: payment.DocNum,
          value: finalVal,
          isUSD: finalIsUSD,
          date: new Date(payment.DocDate),
          tipo: legTipo
        });
      }
    };

    // 1. Procesar la cuenta principal
    const mainAcc = payment.TransferAccount || payment.CashAccount;
    if (mainAcc) {
      const mainVal = payment.TransferSum || payment.CashSum || 0;
      const hasPaymentAccounts = payment.PaymentAccounts && payment.PaymentAccounts.length > 0;
      // Si SAP ya trae las líneas de PaymentAccounts (ej. traslados a
      // Fiducia), esas líneas son la pata real contra la cual concilia el
      // banco; forzar el override aquí también duplicaría el documento con
      // un leg fantasma (IN) que nunca va a cruzar contra ningún movimiento.
      const bankNameOverride = hasPaymentAccounts ? undefined : MANUAL_CUENTA_OVERRIDES[String(payment.DocNum)];
      addLeg(mainAcc, mainVal, null, isIncoming ? "IN" : "OUT", bankNameOverride);
    }

    // 2. Procesar las líneas de pago adicionales
    if (payment.PaymentAccounts && payment.PaymentAccounts.length > 0) {
      for (const pa of payment.PaymentAccounts) {
        if (pa.AccountCode && pa.SumPaid) {
           addLeg(pa.AccountCode, pa.SumPaid, pa.SumPaidFC, isIncoming ? "OUT" : "IN");
        }
      }
    }
  };

  sapPayments.forEach(p => processPayment(p, true));
  vendorPayments.forEach(p => processPayment(p, false));

  let filesInSharepoint: string[] = [];
  try {
    filesInSharepoint = await listDriveFiles();
  } catch (error) {
    console.error("Error fetching sharepoint files:", error);
  }

  const normalizeStr = (s: string) => 
    s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]/gi, "").toUpperCase();

  // 1. Identificar qué archivos necesitamos descargar y procesar
  const filesToDownload = new Set<string>();
  const bankAccountEntries = new Map<string, any>();

  Array.from(bankMap.keys()).forEach((bank) => {
    const accountEntry = Object.entries(ACCOUNT_MAP).find(([key]) => 
      normalizeStr(key) === normalizeStr(bank) || 
      normalizeStr(bank).includes(normalizeStr(key)) || 
      normalizeStr(key).includes(normalizeStr(bank))
    );
    if (accountEntry) {
      bankAccountEntries.set(bank, accountEntry);
      const configInfo = accountEntry[1];
      if (filesInSharepoint.includes(configInfo.file)) {
        filesToDownload.add(configInfo.file);
      }
    }
  });

  // 2. Descargar y parsear los Excels
  const bankMovements = new Map<string, any[]>(); // archivo -> BankMove[]
  
  await Promise.allSettled(Array.from(filesToDownload).map(async (fileName) => {
    try {
      const buffer = await downloadDriveFile(fileName);
      const wb = new ExcelJS.Workbook();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      await wb.xlsx.load(buffer as any);

      // Usamos el format desde ACCOUNT_MAP para encontrar el formato correcto
      const configEntry = Object.values(ACCOUNT_MAP).find(v => v.file === fileName);
      if (!configEntry) return;

      const format = configEntry.format;
      if (format === "compensacion") {
        fixCompensacionDates(wb);
      }

      const reader = READERS[format];
      const hint = REQUIRED_HEADERS_HINT[format];
      const allMoves: any[] = [];

      for (const ws of wb.worksheets) {
        const found = findHeaderRowAndCols(ws, hint);
        if (!found) continue;
        
        const moves = reader(ws, found.headers, found.headerRow);
        allMoves.push(...moves);
      }
      
      bankMovements.set(fileName, allMoves);
    } catch (e) {
      console.error(`Error downloading/parsing ${fileName}:`, e);
    }
  }));

  const bankList = Array.from(bankMap.entries()).map(([bank, docs]) => {
    const accountEntry = bankAccountEntries.get(bank);
    let status = "NO MAPEADO";
    let statusColor = "bg-slate-100 text-slate-800";
    let fileMoves: any[] = [];
    const configInfo = accountEntry ? accountEntry[1] : undefined;
    
    if (accountEntry) {
      if (filesInSharepoint.includes(configInfo.file)) {
        status = "ENCONTRADO EN SHAREPOINT";
        statusColor = "bg-emerald-100 text-emerald-800 border-emerald-200";
        // Solo se ofrecen como candidatas las filas que AÚN no tienen un
        // número de documento escrito en el Excel. Si no se excluyen aquí,
        // una fila ya conciliada en una corrida anterior puede volver a
        // "matchear" contra un documento SAP distinto de otro día que
        // coincida en valor (ej. anticipos redondos que se repiten), dejando
        // el mismo documento SAP asociado a dos movimientos con fechas
        // distintas.
        const allMoves = bankMovements.get(configInfo.file) || [];
        fileMoves = allMoves.filter((m) => m.docValue === null || m.docValue === undefined || m.docValue === "");
      } else {
        status = "FALTA EXCEL EN SHAREPOINT";
        statusColor = "bg-amber-100 text-amber-800 border-amber-200";
      }
    } else {
      status = "NO MAPEADO";
      statusColor = "bg-slate-100 text-slate-600 border-slate-200";
    }

    const processedDocs: any[] = [];
    const usedDocs = new Set<number>();
    const usedBankMoves = new Set<any>();

    // Pass 1: Individual matches
    for (const doc of docs) {
      if (fileMoves.length > 0) {
        // Encontrar candidatos con el mismo valor y una diferencia de fecha de max 10 días
        const candidates = fileMoves.filter(m => 
          !usedBankMoves.has(m) && 
          m.tipo === doc.tipo &&
          Math.abs(m.value - doc.value) < 1 &&
          Math.abs(m.date.getTime() - doc.date.getTime()) <= 10 * 24 * 60 * 60 * 1000
        );

        if (candidates.length > 0) {
          // Ordenar por cercanía de fecha
          candidates.sort((a, b) => 
            Math.abs(a.date.getTime() - doc.date.getTime()) - Math.abs(b.date.getTime() - doc.date.getTime())
          );
          const match = candidates[0];
          usedDocs.add(doc.docNum);
          usedBankMoves.add(match);
          const isFilled = match.docValue !== null && match.docValue !== undefined && match.docValue !== "";
          processedDocs.push({ 
              isCombo: false, docs: [doc], value: doc.value, isMatch: true, 
              date: doc.date, isUSD: doc.isUSD, isFilled,
              excelFile: configInfo ? configInfo.file : undefined,
              sheetName: match.sheet,
              excelRow: match.row,
              excelCol: match.docCol
          });
        }
      }
    }

    // Pass 2: Combo matches (size 2 to 3)
    const unusedDocs = docs.filter(d => !usedDocs.has(d.docNum));
    const unusedBankMoves = fileMoves.filter(m => !usedBankMoves.has(m));
    
    if (unusedDocs.length > 0 && unusedDocs.length <= 50 && unusedBankMoves.length > 0) {
      const getCombinations = function* (elements: any[], length: number): IterableIterator<any[]> {
        if (length === 1) {
          for (const el of elements) yield [el];
          return;
        }
        for (let i = 0; i <= elements.length - length; i++) {
          const head = elements.slice(i, i + 1);
          const tailCombs = getCombinations(elements.slice(i + 1), length - 1);
          for (const tail of tailCombs) {
            yield head.concat(tail);
          }
        }
      };

      const findComboMatch = () => {
         const allCombos = [];
         for (let size = 2; size <= Math.min(3, unusedDocs.length); size++) {
            for (const combo of getCombinations(unusedDocs, size)) {
                allCombos.push({
                    combo,
                    sum: combo.reduce((acc: number, d: any) => acc + d.value, 0)
                });
            }
         }

         for (const m of unusedBankMoves) {
            for (const { combo, sum } of allCombos) {
                // Verificar que todos los docs del combo tienen el mismo tipo que el movimiento bancario
                const sameTipo = combo.every((d: any) => d.tipo === m.tipo);
                if (sameTipo && Math.abs(sum - m.value) < 1 && Math.abs(combo[0].date.getTime() - m.date.getTime()) <= 10 * 24 * 60 * 60 * 1000) {
                    return { combo, m };
                }
            }
         }
         return null;
      };

      let comboMatch = findComboMatch();
      while (comboMatch) {
         const { combo, m } = comboMatch;
         combo.forEach(d => usedDocs.add(d.docNum));
         usedBankMoves.add(m);
         
         combo.forEach(d => {
           const idx = unusedDocs.findIndex(ud => ud.docNum === d.docNum);
           if (idx !== -1) unusedDocs.splice(idx, 1);
         });
         const idxM = unusedBankMoves.indexOf(m);
         if (idxM !== -1) unusedBankMoves.splice(idxM, 1);

         const isFilled = m.docValue !== null && m.docValue !== undefined && m.docValue !== "";
         processedDocs.push({
             isCombo: true,
             docs: combo,
             value: m.value, // sum value
             isMatch: true,
             date: combo[0].date,
             isUSD: combo[0].isUSD,
             isFilled,
             excelFile: configInfo ? configInfo.file : undefined,
             sheetName: m.sheet,
             excelRow: m.row,
             excelCol: m.docCol
         });

         comboMatch = findComboMatch();
      }
    }

    // Pass 3: Unmatched remaining
    for (const doc of unusedDocs) {
        processedDocs.push({ isCombo: false, docs: [doc], value: doc.value, isMatch: false, date: doc.date, isUSD: doc.isUSD, isFilled: false });
    }

    return { bank, docs: processedDocs, status, statusColor };
  });

  return (
    <LiveReconciliationClient 
      initialBankList={bankList} 
      sapPayments={sapPayments} 
      vendorPayments={vendorPayments} 
    />
  );
}
