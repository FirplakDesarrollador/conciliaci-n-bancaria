"use server";

import { ACCOUNT_MAP } from "@/lib/conciliacion/config";
import { reconcile } from "@/lib/conciliacion/reconcile";
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
    const file = formData.get("sapFile");
    if (!(file instanceof File) || file.size === 0) {
      return { status: "error", error: "Selecciona el informe SAP (Informe_de_recaudos_y_pagos.xlsx)." };
    }

    const sapBuffer = Buffer.from(await file.arrayBuffer());

    const bankBuffers = new Map<string, Buffer>();
    for (const [accountKey, { file: fileName }] of Object.entries(ACCOUNT_MAP)) {
      const buffer = await downloadDriveFile(fileName);
      bankBuffers.set(accountKey, buffer);
    }

    const result = await reconcile({ sapBuffer, bankBuffers });

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
