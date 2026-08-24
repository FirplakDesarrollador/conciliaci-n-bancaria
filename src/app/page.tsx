import Link from "next/link";
import { redirect } from "next/navigation";
import { createClient } from "@/lib/supabase/server";
import { logout } from "./login/actions";
import { sapClient } from "@/lib/sap/service-layer";
import { Suspense } from "react";
import LiveReconciliation from "./live-reconciliation";
import { TRANSFER_ACCOUNT_NAMES } from "@/lib/conciliacion/config";

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
  let minOffset = -1;

  if (todayBogotaStr === 'Mon') {
    minOffset = -3;
  } else if (todayBogotaStr === 'Tue' && HOLIDAYS.includes(yesterdayBogotaStr)) {
    minOffset = -4;
  }

  // Ampliamos la ventana de búsqueda en SAP:
  // Desde 15 días antes del minOffset, hasta hoy (0)
  for (let offset = minOffset - 15; offset <= 0; offset++) {
    dates.push(getBogotaDate(offset));
  }
  
  return dates;
}

export default async function Home() {
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
    let nextLink: string | null = `/IncomingPayments?$filter=${filterStr}&$orderby=DocNum`;
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
    let vendorNextLink: string | null = `/VendorPayments?$filter=${filterStr}&$orderby=DocNum`;
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
      <aside className="flex w-64 flex-col border-r border-slate-200 bg-white shadow-sm">
        <div className="flex h-16 items-center gap-3 border-b border-slate-200 px-6">
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-blue-600 shadow-sm shadow-blue-600/20">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M22 12h-4l-3 9L9 3l-3 9H2" />
            </svg>
          </div>
          <span className="font-bold text-slate-900">Firplak S.A.</span>
        </div>

        <nav className="flex-1 space-y-1 p-4">
          <Link href="/" className="flex items-center gap-3 rounded-lg bg-blue-50 px-3 py-2.5 text-sm font-medium text-blue-700 transition-colors">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path><polyline points="9 22 9 12 15 12 15 22"></polyline></svg>
            Inicio
          </Link>
          <Link href="/conciliacion" className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-slate-600 transition-colors hover:bg-slate-100 hover:text-slate-900">
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
          <h2 className="text-lg font-semibold text-slate-800">Panel Principal</h2>
          
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
          {/* Textura de fondo sutil */}
          <div className="pointer-events-none absolute inset-0 z-0 bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:24px_24px]"></div>
          
          <div className="relative z-10 mx-auto max-w-5xl space-y-6">
            <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
              <h3 className="mb-1 text-xl font-bold text-slate-900">Bienvenido de nuevo</h3>
              <p className="mb-6 text-sm text-slate-500">Resumen general de conciliación bancaria para el día de hoy.</p>
              
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
                <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
                  <div className="text-sm font-medium text-slate-500">Movimientos SAP</div>
                  <div className="mt-2 text-3xl font-bold text-slate-900">0</div>
                  <div className="mt-1 flex items-center text-xs text-slate-500">
                    Pendientes por conciliar
                  </div>
                </div>
                <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
                  <div className="text-sm font-medium text-slate-500">Extracto Bancario</div>
                  <div className="mt-2 text-3xl font-bold text-slate-900">0</div>
                  <div className="mt-1 flex items-center text-xs text-slate-500">
                    Registros importados
                  </div>
                </div>
                <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
                  <div className="text-sm font-medium text-slate-500">Estado global</div>
                  <div className="mt-2 flex items-center text-xl font-bold text-emerald-600">
                    <svg xmlns="http://www.w3.org/2000/svg" className="mr-2 h-6 w-6" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                    Al día
                  </div>
                </div>
              </div>
            </div>



            <Suspense fallback={
              <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm mt-6 flex justify-center items-center h-48">
                <div className="flex flex-col items-center gap-4">
                  <div className="h-8 w-8 animate-spin rounded-full border-4 border-slate-200 border-t-blue-600"></div>
                  <p className="text-sm font-medium text-slate-500">Realizando cruce de información con bancos...</p>
                </div>
              </div>
            }>
              <LiveReconciliation sapPayments={sapPayments} vendorPayments={vendorPayments} />
            </Suspense>
          </div>
        </div>
      </main>
    </div>
  );
}
