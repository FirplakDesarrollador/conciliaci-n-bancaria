"use client";

import { useActionState } from "react";
import { runReconciliation, type ReconciliationState } from "./actions";

const initialState: ReconciliationState = { status: "idle" };

export default function ConciliacionPage() {
  const [state, formAction, isPending] = useActionState(runReconciliation, initialState);

  return (
    <div className="mx-auto flex min-h-screen max-w-3xl flex-col gap-6 p-8 font-sans">
      <div>
        <h1 className="text-2xl font-semibold text-black dark:text-zinc-50">Conciliación diaria bancos</h1>
        <p className="mt-1 text-sm text-zinc-600 dark:text-zinc-400">
          Cruza el informe SAP contra los archivos de banco en SharePoint (carpeta FIRPLAK 2026) y sube los
          resultados conciliados de vuelta a la misma carpeta.
        </p>
      </div>

      <form
        action={formAction}
        className="flex flex-col gap-3 rounded-xl border border-black/[.08] p-6 dark:border-white/[.145]"
      >
        <label className="flex flex-col gap-1 text-sm text-zinc-700 dark:text-zinc-300">
          Informe SAP (Informe_de_recaudos_y_pagos.xlsx)
          <input
            type="file"
            name="sapFile"
            accept=".xlsx"
            required
            className="rounded-md border border-black/[.08] p-2 text-black dark:border-white/[.145] dark:text-zinc-50"
          />
        </label>
        <button
          type="submit"
          disabled={isPending}
          className="h-11 self-start rounded-full bg-foreground px-6 text-sm font-medium text-background transition-colors hover:bg-[#383838] disabled:opacity-50 dark:hover:bg-[#ccc]"
        >
          {isPending ? "Conciliando..." : "Ejecutar conciliación"}
        </button>
      </form>

      {state.status === "error" && (
        <p className="rounded-md bg-red-50 px-3 py-2 text-sm text-red-700 dark:bg-red-950 dark:text-red-400">
          {state.error}
        </p>
      )}

      {state.status === "success" && state.summary && (
        <div className="flex flex-col gap-4">
          <p className="text-sm text-zinc-600 dark:text-zinc-400">
            {state.summary.totalSapDocs} documentos SAP cargados.{" "}
            <a
              href={state.summary.resumenWebUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="font-medium underline"
            >
              Ver RESUMEN_CONCILIACION.xlsx
            </a>
          </p>

          <div className="overflow-x-auto">
            <table className="w-full min-w-[720px] border-collapse text-sm">
              <thead>
                <tr className="border-b border-black/[.08] text-left dark:border-white/[.145]">
                  <th className="py-2 pr-2">Cuenta</th>
                  <th className="py-2 pr-2">Exacto</th>
                  <th className="py-2 pr-2">Toler.</th>
                  <th className="py-2 pr-2">Val. único</th>
                  <th className="py-2 pr-2">Sumas</th>
                  <th className="py-2 pr-2">Conteo</th>
                  <th className="py-2 pr-2">Inverso</th>
                  <th className="py-2 pr-2">Grupo</th>
                  <th className="py-2 pr-2">Ambiguos</th>
                  <th className="py-2 pr-2">Sin doc.</th>
                  <th className="py-2 pr-2">Archivo</th>
                </tr>
              </thead>
              <tbody>
                {state.summary.accounts.map((a) => (
                  <tr key={a.cuentaKey} className="border-b border-black/[.05] dark:border-white/[.08]">
                    <td className="py-2 pr-2">{a.cuentaKey}</td>
                    <td className="py-2 pr-2">{a.stats.matchExacto}</td>
                    <td className="py-2 pr-2">{a.stats.matchTolerancia}</td>
                    <td className="py-2 pr-2">{a.stats.matchValorUnico}</td>
                    <td className="py-2 pr-2">{a.stats.matchPar}</td>
                    <td className="py-2 pr-2">{a.stats.matchConteo}</td>
                    <td className="py-2 pr-2">{a.stats.matchInverso}</td>
                    <td className="py-2 pr-2">{a.stats.matchGrupo}</td>
                    <td className="py-2 pr-2">{a.stats.ambiguos}</td>
                    <td className="py-2 pr-2">{a.stats.sinDocumento}</td>
                    <td className="py-2 pr-2">
                      <a href={a.webUrl} target="_blank" rel="noopener noreferrer" className="underline">
                        {a.outputFileName}
                      </a>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
