import Link from 'next/link';
import { redirect } from 'next/navigation';
import { createClient } from '@/lib/supabase/server';
import { sapClient } from '@/lib/sap/service-layer';
import { logout } from '../login/actions';
import ReportPreview, { ReportItem } from './ReportPreview';

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
    dates.push(getBogotaDate(-4)); // viernes
    dates.push(getBogotaDate(-3)); // sábado
    dates.push(getBogotaDate(-2)); // domingo
    dates.push(getBogotaDate(-1)); // lunes
  } else {
    dates.push(getBogotaDate(-1));
  }
  return dates;
}

export default async function ComparacionPage() {
  const supabase = await createClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();

  if (!user) {
    redirect("/login");
  }

  const dates = getTargetDates();
  
  let sapPayments: any[] = [];
  let vendorPayments: any[] = [];
  
  try {
    const filterStr = dates.map(d => `DocDate eq '${d}'`).join(' or ');
    
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
        nextLink = data['odata.nextLink'] ? (data['odata.nextLink'].startsWith('/') ? data['odata.nextLink'] : '/' + data['odata.nextLink']) : null;
      } else {
        break;
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
        vendorNextLink = data['odata.nextLink'] ? (data['odata.nextLink'].startsWith('/') ? data['odata.nextLink'] : '/' + data['odata.nextLink']) : null;
      } else {
        break;
      }
    }
  } catch (error) {
    console.error("Exception fetching SAP payments:", error);
  }

  // 3. Obtener Historial de base de datos (solo ayer y hoy)
  const now = new Date();
  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);
  
  const { data: history } = await supabase
    .from('reconciliation_history')
    .select('*')
    .gte('created_at', yesterday.toISOString());
  
  const historySap = (history || []).filter(h => h.tipo === 'Recibido');
  const historyVendor = (history || []).filter(h => h.tipo === 'Efectuado');
  
  const matchedSet = new Set<string>();
  (history || []).forEach(h => {
      h.doc_num.split('-').forEach((d: string) => matchedSet.add(d));
  });

  // 4. Formatear datos para el ReportPreview
  const pagosRecibidos: ReportItem[] = [];
  
  // Agregar conciliados
  historySap.forEach(h => {
    pagosRecibidos.push({
      docNum: h.doc_num,
      valor: h.valor,
      banco: h.banco,
      fecha: h.fecha,
      info: h.info,
      estado: "Conciliado"
    });
  });

  // Agregar pendientes
  sapPayments.forEach(p => {
    if (!matchedSet.has(String(p.DocNum))) {
      const val = p.TransferSum || p.CashSum || p.DocTotal || 0;
      const isUSD = p.DocCurrency === 'USD' && p.DocRate > 0;
      const finalVal = isUSD ? (val / p.DocRate) : val;

      pagosRecibidos.push({
        docNum: p.DocNum ? String(p.DocNum) : '',
        valor: finalVal,
        banco: p.TransferAccount || p.CashAccount || p.AcctName || p.CardName || '',
        fecha: p.DocDate ? new Date(p.DocDate).toISOString().split('T')[0] : '',
        info: p.Comments || p.Address || '',
        estado: "Pendiente"
      });
    }
  });

  const pagosEfectuados: ReportItem[] = [];
  
  // Agregar conciliados
  historyVendor.forEach(h => {
    pagosEfectuados.push({
      docNum: h.doc_num,
      valor: h.valor,
      banco: h.banco,
      fecha: h.fecha,
      info: h.info,
      estado: "Conciliado"
    });
  });

  // Agregar pendientes
  vendorPayments.forEach(p => {
    if (!matchedSet.has(String(p.DocNum))) {
      const val = p.TransferSum || p.CashSum || p.DocTotal || 0;
      const isUSD = p.DocCurrency === 'USD' && p.DocRate > 0;
      const finalVal = isUSD ? (val / p.DocRate) : val;

      pagosEfectuados.push({
        docNum: p.DocNum ? String(p.DocNum) : '',
        valor: finalVal,
        banco: p.TransferAccount || p.CashAccount || p.AcctName || p.CardName || '',
        fecha: p.DocDate ? new Date(p.DocDate).toISOString().split('T')[0] : '',
        info: p.Comments || p.Address || '',
        estado: "Pendiente"
      });
    }
  });

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
          <Link href="/conciliacion" className="flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-slate-600 transition-colors hover:bg-slate-100 hover:text-slate-900">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
            SAP
          </Link>
          <Link href="/comparacion" className="flex items-center gap-3 rounded-lg bg-blue-50 px-3 py-2.5 text-sm font-medium text-blue-700 transition-colors">
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
          <h2 className="text-lg font-semibold text-slate-800">Visualización de Informe</h2>
          <div className="flex items-center gap-4">
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
          
          <div className="relative z-10 mx-auto max-w-6xl h-full">
            <ReportPreview 
              pagosRecibidos={pagosRecibidos} 
              pagosEfectuados={pagosEfectuados} 
            />
          </div>
        </div>
      </main>
    </div>
  );
}
