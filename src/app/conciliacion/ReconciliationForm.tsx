"use client";

import { useActionState } from "react";
import { runReconciliation, type ReconciliationState } from "./actions";

const initialState: ReconciliationState = { status: "idle" };

export default function ReconciliationForm() {
  const [state, formAction, isPending] = useActionState(runReconciliation, initialState);

  return (
    <div className="flex flex-col gap-6">
      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
        <h3 className="mb-2 text-xl font-bold text-slate-900">Conciliación diaria bancos</h3>
        <p className="mb-6 text-sm text-slate-500">
          Consulta la API de SAP en el rango de fechas seleccionado y cruza la información contra los archivos de banco en SharePoint (carpeta FIRPLAK 2026), subiendo los resultados conciliados a la misma carpeta.
        </p>

        <form
          action={formAction}
          className="flex flex-col gap-4"
        >
          <div className="flex gap-4">
            <label className="flex flex-1 flex-col gap-1 text-sm font-medium text-slate-700">
              Fecha de inicio
              <input
                type="date"
                name="startDate"
                required
                defaultValue={new Date(new Date().setDate(1)).toISOString().split('T')[0]}
                className="mt-1 block w-full rounded-md border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </label>
            <label className="flex flex-1 flex-col gap-1 text-sm font-medium text-slate-700">
              Fecha fin
              <input
                type="date"
                name="endDate"
                required
                defaultValue={new Date().toISOString().split('T')[0]}
                className="mt-1 block w-full rounded-md border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </label>
          </div>
          <button
            type="submit"
            disabled={isPending}
            className="h-10 self-start rounded-lg bg-blue-600 px-6 text-sm font-medium text-white transition-colors hover:bg-blue-700 disabled:opacity-50"
          >
            {isPending ? "Conciliando..." : "Ejecutar conciliación"}
          </button>
        </form>

        {state.status === "error" && (
          <p className="mt-4 rounded-md bg-red-50 px-4 py-3 text-sm text-red-700 border border-red-200">
            {state.error}
          </p>
        )}
      </div>

      {state.status === "success" && state.summary && (
        <div className="flex flex-col gap-4 rounded-2xl border border-emerald-200 bg-white p-6 shadow-sm">
          <h3 className="text-lg font-bold text-slate-900">Resultados de conciliación</h3>
          <p className="text-sm font-medium text-emerald-900">
            {state.summary.totalSapDocs} documentos SAP cargados.{" "}
            <a
              href={state.summary.resumenWebUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="font-bold underline text-blue-700 hover:text-blue-800 ml-2"
            >
              Ver RESUMEN_CONCILIACION.xlsx
            </a>
          </p>

          <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
            <table className="w-full text-left text-sm whitespace-nowrap">
              <thead className="bg-slate-50 text-xs uppercase tracking-wider text-slate-500 border-b border-slate-200">
                <tr>
                  <th className="px-4 py-3 font-medium">Cuenta</th>
                  <th className="px-4 py-3 font-medium">Exacto</th>
                  <th className="px-4 py-3 font-medium">Toler.</th>
                  <th className="px-4 py-3 font-medium">Val. único</th>
                  <th className="px-4 py-3 font-medium">Sumas</th>
                  <th className="px-4 py-3 font-medium">Conteo</th>
                  <th className="px-4 py-3 font-medium">Inverso</th>
                  <th className="px-4 py-3 font-medium">Grupo</th>
                  <th className="px-4 py-3 font-medium">Ambiguos</th>
                  <th className="px-4 py-3 font-medium">Sin doc.</th>
                  <th className="px-4 py-3 font-medium">Archivo</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {state.summary.accounts.map((a) => (
                  <tr key={a.cuentaKey} className="hover:bg-slate-50 transition-colors">
                    <td className="px-4 py-2 font-medium text-slate-900">{a.cuentaKey}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.matchExacto}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.matchTolerancia}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.matchValorUnico}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.matchPar}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.matchConteo}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.matchInverso}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.matchGrupo}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.ambiguos}</td>
                    <td className="px-4 py-2 text-slate-600">{a.stats.sinDocumento}</td>
                    <td className="px-4 py-2">
                      <a href={a.webUrl} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline">
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
