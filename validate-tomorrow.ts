/**
 * Simula exactamente la logica de page.tsx para el dia de manana (2026-08-21, Viernes).
 * Verifica que:
 *  1. Los documentos que fallaron hoy (80832, 80834, 80835, 80847, 80849, 80828, 80837) SI aparecen.
 *  2. Los VendorPayments de Sudameris (20043385-88) tambien.
 *  3. La paginacion es correcta (ninguna pagina se salta documentos).
 */
import * as dotenv from 'dotenv';
dotenv.config({ path: '.env.local' });
import { sapClient } from './src/lib/sap/service-layer';

function getTargetDatesTomorrow() {
  // Simula el instante del 2026-08-21 en Bogota
  const now = new Date('2026-08-21T09:00:00-05:00');
  const getBogotaDate = (daysOffset: number) => {
    const d = new Date(now.getTime() + daysOffset * 24 * 60 * 60 * 1000);
    return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Bogota' }).format(d);
  };

  const todayBogotaStr = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Bogota', weekday: 'short' }).format(now);
  const yesterdayBogotaStr = getBogotaDate(-1);

  const HOLIDAYS = [
    '2026-01-01', '2026-01-12', '2026-03-23', '2026-04-02', '2026-04-03',
    '2026-05-01', '2026-05-18', '2026-06-08', '2026-06-15', '2026-06-29',
    '2026-07-13', '2026-07-20', '2026-08-07', '2026-08-17', '2026-10-12',
    '2026-11-02', '2026-11-16', '2026-12-08', '2026-12-25',
  ];

  const dates: string[] = [];
  let minOffset = -1;

  if (todayBogotaStr === 'Mon') {
    minOffset = -3;
  } else if (todayBogotaStr === 'Tue' && HOLIDAYS.includes(yesterdayBogotaStr)) {
    minOffset = -4;
  }

  // Con el fix: 15 dias atras del minOffset
  for (let offset = minOffset - 15; offset <= 0; offset++) {
    dates.push(getBogotaDate(offset));
  }
  return { dates, todayBogotaStr, minOffset };
}

async function fetchAll(endpoint: 'IncomingPayments' | 'VendorPayments', filterStr: string) {
  const allDocs: any[] = [];
  let nextLink: string | null = `/${endpoint}?$filter=${filterStr}&$orderby=DocNum`;
  let pageCount = 0;
  const pageDocNums: number[][] = [];

  while (nextLink) {
    const res = await sapClient.request(nextLink);
    if (!res.ok) {
      console.error(`Error en ${endpoint}:`, await res.text());
      break;
    }
    const data = await res.json();
    const pageDocs = (data.value || []).map((p: any) => p.DocNum);
    pageDocNums.push(pageDocs);
    allDocs.push(...(data.value || []));
    pageCount++;

    if (data['odata.nextLink']) {
      nextLink = data['odata.nextLink'] as string;
      if (!nextLink.startsWith('/')) nextLink = '/' + nextLink;
    } else {
      nextLink = null;
    }
  }
  return { allDocs, pageCount, pageDocNums };
}

async function main() {
  const { dates, todayBogotaStr, minOffset } = getTargetDatesTomorrow();
  console.log(`\n=== SIMULACION PARA MANANA 2026-08-21 ===`);
  console.log(`Dia semana: ${todayBogotaStr}, minOffset: ${minOffset}`);
  console.log(`Rango de fechas (${dates.length} dias):`);
  console.log(` Desde: ${dates[0]}  Hasta: ${dates[dates.length - 1]}\n`);

  const filterStr = dates.map(d => `DocDate eq '${d}'`).join(' or ');

  // --- IncomingPayments ---
  console.log("Descargando IncomingPayments...");
  const { allDocs: incoming, pageCount: pInc, pageDocNums: pageNumsInc } = await fetchAll('IncomingPayments', filterStr);
  console.log(`  Total IncomingPayments: ${incoming.length} (en ${pInc} pagina(s))\n`);

  // Verificar duplicados entre paginas (bug de paginacion)
  const allDocNumsInc = incoming.map((p: any) => p.DocNum);
  const uniqueInc = new Set(allDocNumsInc);
  if (uniqueInc.size !== allDocNumsInc.length) {
    console.log(`  ⚠️  ATENCION: Hay DocNums DUPLICADOS entre paginas (posible bug de paginacion)`);
    const counts: Record<number, number> = {};
    allDocNumsInc.forEach((n: number) => counts[n] = (counts[n] || 0) + 1);
    Object.entries(counts).filter(([, c]) => c > 1).forEach(([n, c]) => console.log(`    - DocNum ${n} aparece ${c} veces`));
  } else {
    console.log(`  ✅ Sin duplicados entre paginas (paginacion correcta)\n`);
  }

  // Verificar documentos clave del 20 de agosto
  const expectedIncoming = [80832, 80834, 80835, 80847, 80849, 80828, 80837];
  console.log("Verificando documentos del 20 de agosto (IncomingPayments):");
  for (const d of expectedIncoming) {
    const doc = incoming.find((p: any) => p.DocNum === d);
    if (doc) {
      const tipo = doc.DocType === 'rAccount' ? '🔄 Traslado' : '💰 Cliente';
      const cuenta = doc.TransferAccount || doc.CashAccount || 'N/A';
      const monto = (doc.TransferSum || doc.CashSum || 0).toLocaleString('es-CO');
      console.log(`  ✅ ${d} | DocDate: ${doc.DocDate?.substring(0, 10)} | ${tipo} | Cuenta: ${cuenta} | $${monto}`);
    } else {
      console.log(`  ❌ ${d} NO fue encontrado en la consulta.`);
    }
  }

  // --- VendorPayments ---
  console.log("\nDescargando VendorPayments...");
  const { allDocs: vendor, pageCount: pVen } = await fetchAll('VendorPayments', filterStr);
  console.log(`  Total VendorPayments: ${vendor.length} (en ${pVen} pagina(s))\n`);

  const expectedVendor = [20043385, 20043386, 20043387, 20043388];
  console.log("Verificando pagos efectuados del 20 de agosto (VendorPayments):");
  for (const d of expectedVendor) {
    const doc = vendor.find((p: any) => p.DocNum === d);
    if (doc) {
      const cuenta = doc.TransferAccount || doc.CashAccount || 'N/A';
      const monto = (doc.TransferSum || doc.CashSum || 0).toLocaleString('es-CO');
      console.log(`  ✅ ${d} | DocDate: ${doc.DocDate?.substring(0, 10)} | Cuenta: ${cuenta} | $${monto}`);
    } else {
      console.log(`  ❌ ${d} NO fue encontrado en la consulta.`);
    }
  }

  // --- Resumen de tipo (IN/OUT) para los traslados clave ---
  console.log("\nVerificando logica IN/OUT para traslados 80835 y 80837:");
  for (const docNum of [80835, 80837]) {
    const doc = incoming.find((p: any) => p.DocNum === docNum);
    if (!doc) { console.log(`  ❌ ${docNum} no encontrado`); continue; }
    console.log(`  Doc ${docNum} - DocType: ${doc.DocType}`);
    const mainAcc = doc.TransferAccount || doc.CashAccount;
    const mainTipo = 'IN';
    console.log(`    Cuenta principal: ${mainAcc} → tipo asignado: ${mainTipo}`);
    if (doc.PaymentAccounts?.length > 0) {
      for (const pa of doc.PaymentAccounts) {
        console.log(`    PaymentAccount: ${pa.AccountCode} (SumPaid: ${pa.SumPaid?.toLocaleString('es-CO')}) → tipo asignado: OUT`);
      }
    }
  }

  console.log("\n=== FIN DE LA VALIDACION ===");
}

main().catch(console.error);
