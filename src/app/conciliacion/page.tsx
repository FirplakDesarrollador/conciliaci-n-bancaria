import Link from "next/link";
import { redirect } from "next/navigation";
import { createClient } from "@/lib/supabase/server";
import { logout } from "../login/actions";
import { sapClient } from "@/lib/sap/service-layer";
import { TRANSFER_ACCOUNT_NAMES } from "@/lib/conciliacion/config";
import ReconciliationForm from "./ReconciliationForm";

function getTargetDates() {
  const now = new Date();
  const getBogotaDate = (daysOffset: number) => {
    const d = new Date(now.getTime() + daysOffset * 24 * 60 * 60 * 1000);
    return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Bogota' }).format(d);
  };
  
  const todayBogotaStr = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Bogota', weekday: 'short' }).format(now);
  const yesterdayBogotaStr = getBogotaDate(-1);

  // Días festivos para Colombia (YYYY-MM-DD)
  const HOLIDAYS = [
    '2026-01-01', '2026-01-12', '2026-03-23', '2026-04-02', '2026-04-03',
    '2026-05-01', '2026-05-18', '2026-06-08', '2026-06-15', '2026-06-29',
    '2026-07-13', '2026-07-20', '2026-08-07', '2026-08-17', '2026-10-12',
    '2026-11-02', '2026-11-16', '2026-12-08', '2026-12-25'
  ];
  
  const dates = [];
  if (todayBogotaStr === 'Mon') {
    dates.push(getBogotaDate(-3));
    dates.push(getBogotaDate(-2));
    dates.push(getBogotaDate(-1));
  } else if (todayBogotaStr === 'Tue' && HOLIDAYS.includes(yesterdayBogotaStr)) {
    // Si hoy es martes y ayer fue lunes festivo, traemos viernes, sábado, domingo y lunes festivo
    dates.push(getBogotaDate(-4)); // viernes
    dates.push(getBogotaDate(-3)); // sábado
    dates.push(getBogotaDate(-2)); // domingo
    dates.push(getBogotaDate(-1)); // lunes
  } else {
    dates.push(getBogotaDate(-1));
  }
  return dates;
}

export default async function ConciliacionPage() {
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();

  if (!user) {
    redirect("/login");
  }

  const dates = getTargetDates();
  
  let sapPayments: any[] = [];
  let errorMessage: string | null = null;
  
  let vendorPayments: any[] = [];
  let vendorErrorMessage: string | null = null;
  
  try {
    const filterStr = dates.map(d => `DocDate eq '${d}'`).join(' or ');
    
    // 1. Fetch Pagos recibidos (IncomingPayments)
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
        errorMessage = await res.text();
        console.error("Error fetching SAP incoming payments:", errorMessage);
        break;
      }
    }

    // 2. Fetch Pagos efectuados (VendorPayments)
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
        vendorErrorMessage = await res.text();
        console.error("Error fetching SAP vendor payments:", vendorErrorMessage);
        break;
      }
    }

  } catch (error: any) {
    const errStr = error.message || String(error);
    errorMessage = errorMessage || errStr;
    vendorErrorMessage = vendorErrorMessage || errStr;
    console.error("Exception fetching SAP payments:", error);
  }

  return (
    <div className="flex h-screen w-full bg-slate-50 selection:bg-blue-600/20">
      {/* Sidebar (Slide bar) */}
      <aside className="flex w-64 flex-col border-r border-slate-200 bg-white shadow-sm shrink-0">
        <div className="flex h-16 items-center gap-3 border-b border-slate-200 px-6">
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-blue-600 shadow-sm shadow-blue-600/20">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M22 12h-4l-3 9L9 3l-3 9H2" />
            </svg>
          </div>
          <span className="font-bold text-slate-900">Firplak S.A.</span>
        </div>

        <nav className="flex-1 space-y-1 p-4">
          <Link href="/" className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-slate-600 transition-colors hover:bg-slate-100 hover:text-slate-900">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path><polyline points="9 22 9 12 15 12 15 22"></polyline></svg>
            Inicio
          </Link>
          <Link href="/conciliacion" className="flex items-center gap-3 rounded-lg bg-blue-50 px-3 py-2.5 text-sm font-medium text-blue-700 transition-colors">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
            SAP
          </Link>
          <Link href="/comparacion" className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-slate-600 transition-colors hover:bg-slate-100 hover:text-slate-900">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line></svg>
            Informe
          </Link>
        </nav>

        <div className="border-t border-slate-200 p-4">
          <form action={logout}>
            <button
              type="submit"
              className="flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-slate-600 transition-colors hover:bg-slate-100 hover:text-slate-900"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path>
                <polyline points="16 17 21 12 16 7"></polyline>
                <line x1="21" y1="12" x2="9" y2="12"></line>
              </svg>
              Cerrar sesión
            </button>
          </form>
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="flex flex-1 flex-col h-screen overflow-hidden">
        {/* Header Superior */}
        <header className="flex h-16 shrink-0 items-center justify-between border-b border-slate-200 bg-white px-8">
          <h2 className="text-lg font-semibold text-slate-800">Conciliación SAP</h2>
          
          <div className="flex items-center gap-4">
            <div className="hidden items-center gap-2 text-sm text-slate-500 md:flex">
              <span className="relative flex h-2 w-2">
                <span className="absolute inline-flex h-full w-full animate-ping rounded-full bg-emerald-400 opacity-75"></span>
                <span className="relative inline-flex h-2 w-2 rounded-full bg-emerald-500"></span>
              </span>
              Conectado a SAP
            </div>
            <div className="h-8 w-px bg-slate-200"></div>
            <div className="flex items-center gap-2">
              <div className="flex h-8 w-8 items-center justify-center rounded-full bg-blue-100 text-sm font-bold text-blue-700">
                {user.email?.charAt(0).toUpperCase()}
              </div>
              <span className="text-sm font-medium text-slate-700">{user.email}</span>
            </div>
          </div>
        </header>

        {/* Scrollable Content */}
        <div className="relative flex-1 overflow-auto p-8">
          <div className="pointer-events-none absolute inset-0 z-0 bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:24px_24px]"></div>
          
          <div className="relative z-10 mx-auto max-w-5xl space-y-6">
            <ReconciliationForm />
            
            <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm overflow-hidden flex flex-col">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <h3 className="flex items-center gap-2 text-lg font-bold text-slate-900">
                    Lista de Pagos recibidos (SAP)
                    <span className="rounded-full bg-blue-100 px-2.5 py-0.5 text-xs font-semibold text-blue-700">
                      {sapPayments.length} registros
                    </span>
                  </h3>
                  <p className="text-sm text-slate-500">
                    Fechas consultadas: {dates.join(', ')}
                  </p>
                </div>
              </div>
              
              <div className="overflow-x-auto border border-slate-200 rounded-lg">
                <table className="w-full text-left text-sm text-slate-600 whitespace-nowrap">
                  <thead className="border-b border-slate-200 bg-slate-50 text-xs text-slate-500 uppercase tracking-wider">
                    <tr>
                      <th className="px-4 py-3 font-medium">#</th>
                      <th className="px-4 py-3 font-medium">Fecha</th>
                      <th className="px-4 py-3 font-medium">Cliente (CardName)</th>
                      <th className="px-4 py-3 font-medium">Tercero (U_DescripTercero)</th>
                      <th className="px-4 py-3 font-medium">Info.detallada</th>
                      <th className="px-4 py-3 font-medium">Comentarios</th>
                      <th className="px-4 py-3 font-medium">Asesor (U_asesor)</th>
                      <th className="px-4 py-3 font-medium">Tipo RC</th>
                      <th className="px-4 py-3 font-medium text-right">Total del documento</th>
                      <th className="px-4 py-3 font-medium">Banco</th>
                      <th className="px-4 py-3 font-medium text-right">Total doc. (ME)</th>
                      <th className="px-4 py-3 font-medium">Cuenta Bancaria</th>
                      <th className="px-4 py-3 font-medium">RC Prov. (U_NumRC...)</th>
                      <th className="px-4 py-3 font-medium">Cod.Tercero</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 bg-white">
                    {errorMessage ? (
                      <tr>
                        <td colSpan={14} className="px-4 py-8 text-center text-red-500 font-medium">
                          Error conectando a SAP: {errorMessage}
                        </td>
                      </tr>
                    ) : sapPayments.length > 0 ? (
                      sapPayments.map((payment: any, idx: number) => (
                        <tr key={payment.DocNum || idx} className="hover:bg-slate-50 transition-colors">
                          <td className="px-4 py-2">{payment.DocNum}</td>
                          <td className="px-4 py-2">{payment.DocDate ? payment.DocDate.split('T')[0] : ''}</td>
                          <td className="px-4 py-2 font-medium text-slate-900 truncate max-w-[150px]" title={payment.CardName}>{payment.CardName}</td>
                          <td className="px-4 py-2 truncate max-w-[150px]" title={payment.U_DescripTercero}>{payment.U_DescripTercero || ''}</td>
                          <td className="px-4 py-2 truncate max-w-[150px]" title={payment.JournalRemarks}>{payment.JournalRemarks || ''}</td>
                          <td className="px-4 py-2 truncate max-w-[150px]" title={payment.Remarks}>{payment.Remarks || ''}</td>
                          <td className="px-4 py-2">{payment.U_asesor || payment.SalesPersonCode || ''}</td>
                          <td className="px-4 py-2">{payment.U_TipoRC || payment.DocType || ''}</td>
                          <td className="px-4 py-2 text-right font-medium">
                            {new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP' }).format(payment.TransferSum || payment.CashSum || 0)}
                          </td>
                          <td className="px-4 py-2 font-medium text-slate-700">
                            {TRANSFER_ACCOUNT_NAMES[payment.TransferAccount] || TRANSFER_ACCOUNT_NAMES[payment.CashAccount] || payment.TransferAccount || payment.CashAccount || '—'}
                          </td>
                          <td className="px-4 py-2 text-right font-medium text-emerald-700">
                            {payment.DocCurrency === 'USD' && payment.DocRate > 0
                              ? new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format((payment.TransferSum || payment.CashSum || 0) / payment.DocRate)
                              : '—'}
                          </td>
                          <td className="px-4 py-2">{payment.TransferAccount || payment.CashAccount || '—'}</td>
                          <td className="px-4 py-2">{payment.U_NumRCProvisional || ''}</td>
                          <td className="px-4 py-2">{payment.U_CodTecero || payment.CardCode || ''}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={14} className="px-4 py-8 text-center text-slate-500">
                          No se encontraron pagos recibidos en SAP para las fechas seleccionadas.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Vendor Payments Table */}
            <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm overflow-hidden flex flex-col">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <h3 className="flex items-center gap-2 text-lg font-bold text-slate-900">
                    Lista de Pagos efectuados (SAP)
                    <span className="rounded-full bg-indigo-100 px-2.5 py-0.5 text-xs font-semibold text-indigo-700">
                      {vendorPayments.length} registros
                    </span>
                  </h3>
                  <p className="text-sm text-slate-500">
                    Fechas consultadas: {dates.join(', ')}
                  </p>
                </div>
              </div>
              
              <div className="overflow-x-auto border border-slate-200 rounded-lg">
                <table className="w-full text-left text-sm text-slate-600 whitespace-nowrap">
                  <thead className="border-b border-slate-200 bg-slate-50 text-xs text-slate-500 uppercase tracking-wider">
                    <tr>
                      <th className="px-4 py-3 font-medium">#</th>
                      <th className="px-4 py-3 font-medium">Fecha</th>
                      <th className="px-4 py-3 font-medium">Prov. (U_CodTecero)</th>
                      <th className="px-4 py-3 font-medium">Acreedor/Nombre</th>
                      <th className="px-4 py-3 font-medium">Info.detallada</th>
                      <th className="px-4 py-3 font-medium">Comentarios</th>
                      <th className="px-4 py-3 font-medium text-right">Importe</th>
                      <th className="px-4 py-3 font-medium">Banco</th>
                      <th className="px-4 py-3 font-medium text-right">Valor en USD</th>
                      <th className="px-4 py-3 font-medium">Cuenta Transf.</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 bg-white">
                    {vendorErrorMessage ? (
                      <tr>
                        <td colSpan={10} className="px-4 py-8 text-center text-red-500 font-medium">
                          Error conectando a SAP: {vendorErrorMessage}
                        </td>
                      </tr>
                    ) : vendorPayments.length > 0 ? (
                      vendorPayments.map((payment: any, idx: number) => (
                        <tr key={payment.DocNum || idx} className="hover:bg-slate-50 transition-colors">
                          <td className="px-4 py-2">{payment.DocNum}</td>
                          <td className="px-4 py-2">{payment.DocDate ? payment.DocDate.split('T')[0] : ''}</td>
                          <td className="px-4 py-2 font-medium text-slate-900">{payment.U_CodTecero || payment.CardCode || ''}</td>
                          <td className="px-4 py-2 truncate max-w-[150px]" title={payment.U_DescripTercero || payment.CardName}>{payment.U_DescripTercero || payment.CardName}</td>
                          <td className="px-4 py-2 truncate max-w-[150px]" title={payment.JournalRemarks}>{payment.JournalRemarks || ''}</td>
                          <td className="px-4 py-2 truncate max-w-[150px]" title={payment.Remarks}>{payment.Remarks || ''}</td>
                          <td className="px-4 py-2 text-right font-medium">
                            {new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP' }).format(payment.TransferSum || payment.CashSum || 0)}
                          </td>
                          <td className="px-4 py-2 font-medium text-slate-700">
                            {TRANSFER_ACCOUNT_NAMES[payment.TransferAccount] || TRANSFER_ACCOUNT_NAMES[payment.CashAccount] || payment.TransferAccount || payment.CashAccount || '—'}
                          </td>
                          <td className="px-4 py-2 text-right font-medium text-emerald-700">
                            {payment.DocCurrency === 'USD' && payment.DocRate > 0
                              ? new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format((payment.TransferSum || payment.CashSum || 0) / payment.DocRate)
                              : '—'}
                          </td>
                          <td className="px-4 py-2">{payment.TransferAccount || payment.CashAccount || '—'}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={10} className="px-4 py-8 text-center text-slate-500">
                          No se encontraron pagos efectuados en SAP para las fechas seleccionadas.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}
