"use server";

import { ACCOUNT_MAP } from "@/lib/conciliacion/config";
import { reconcileDocs } from "@/lib/conciliacion/reconcile";
import type { AccountStats } from "@/lib/conciliacion/types";
import { downloadDriveFile, uploadDriveFile } from "@/lib/graph/sharepoint";

export interface ReconciliationAccountSummary {
  cuentaKey: string;
  archivo: string;
  outputFileName: string;
  webUrl: string;
  stats: AccountStats;
}

export interface ReconciliationState {
  status: "idle" | "success" | "error";
  error?: string;
  summary?: {
    totalSapDocs: number;
    accounts: ReconciliationAccountSummary[];
    resumenWebUrl: string;
  };
}

export async function runReconciliation(
  _prevState: ReconciliationState,
  formData: FormData
): Promise<ReconciliationState> {
  try {
    const startDateStr = formData.get("startDate") as string;
    const endDateStr = formData.get("endDate") as string;
    if (!startDateStr || !endDateStr) {
      return { status: "error", error: "Selecciona el rango de fechas de inicio y fin." };
    }

    const { sapClient } = await import("@/lib/sap/service-layer");
    const { convertApiToSapDocs } = await import("@/lib/conciliacion/sap");

    let sapPayments: any[] = [];
    let vendorPayments: any[] = [];
    
    const filterStr = `DocDate ge '${startDateStr}' and DocDate le '${endDateStr}'`;
    
    // 1. Fetch Pagos recibidos
    let nextLink: string | null = `/IncomingPayments?$filter=${filterStr}`;
    while (nextLink) {
      const res = await sapClient.request(nextLink);
      if (res.ok) {
        const data = await res.json();
        const validPayments = (data.value || []).filter((p: any) => {
          const jr = p.JournalRemarks ? String(p.JournalRemarks).trim().toUpperCase() : "";
          const rm = p.Remarks ? String(p.Remarks).trim().toUpperCase() : "";
          return jr !== "CANCELADO" && rm !== "CANCELADO" && p.Cancelled !== "tYES";
        });
        sapPayments = sapPayments.concat(validPayments);
        if (data['odata.nextLink']) {
          nextLink = data['odata.nextLink'] as string;
          if (!nextLink.startsWith('/')) nextLink = '/' + nextLink;
        } else {
          nextLink = null;
        }
      } else {
        const errorMessage = await res.text();
        throw new Error(`Error conectando a SAP (IncomingPayments): ${errorMessage}`);
      }
    }

    // 2. Fetch Pagos efectuados
    let vendorNextLink: string | null = `/VendorPayments?$filter=${filterStr}`;
    while (vendorNextLink) {
      const res = await sapClient.request(vendorNextLink);
      if (res.ok) {
        const data = await res.json();
        const validVendorPayments = (data.value || []).filter((p: any) => {
          const jr = p.JournalRemarks ? String(p.JournalRemarks).trim().toUpperCase() : "";
          const rm = p.Remarks ? String(p.Remarks).trim().toUpperCase() : "";
          return jr !== "CANCELADO" && rm !== "CANCELADO" && p.Cancelled !== "tYES";
        });
        vendorPayments = vendorPayments.concat(validVendorPayments);
        if (data['odata.nextLink']) {
          vendorNextLink = data['odata.nextLink'] as string;
          if (!vendorNextLink.startsWith('/')) vendorNextLink = '/' + vendorNextLink;
        } else {
          vendorNextLink = null;
        }
      } else {
        const errorMessage = await res.text();
        throw new Error(`Error conectando a SAP (VendorPayments): ${errorMessage}`);
      }
    }

    const pool = convertApiToSapDocs(sapPayments, vendorPayments);

    const bankBuffers = new Map<string, Buffer>();
    for (const [accountKey, { file: fileName }] of Object.entries(ACCOUNT_MAP)) {
      const buffer = await downloadDriveFile(fileName);
      bankBuffers.set(accountKey, buffer);
    }

    const result = await reconcileDocs(pool, bankBuffers);

    const accounts: ReconciliationAccountSummary[] = [];
    for (const r of result.results) {
      const { webUrl } = await uploadDriveFile(r.outputFileName, r.workbookBuffer);
      accounts.push({
        cuentaKey: r.cuentaKey,
        archivo: r.archivo,
        outputFileName: r.outputFileName,
        webUrl,
        stats: r.stats,
      });
    }

    const { webUrl: resumenWebUrl } = await uploadDriveFile("RESUMEN_CONCILIACION.xlsx", result.resumenBuffer);

    return {
      status: "success",
      summary: { totalSapDocs: result.totalSapDocs, accounts, resumenWebUrl },
    };
  } catch (error) {
    return { status: "error", error: (error as Error).message };
  }
}
