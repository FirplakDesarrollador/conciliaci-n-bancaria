import { NextResponse } from "next/server";
import { fetchSapPayments } from "@/lib/conciliacion/fetch-sap-payments";
import { runFullSync } from "@/lib/conciliacion/auto-sync";

export const maxDuration = 300;

/**
 * Corre el cruce SAP <-> bancos y escribe los matches en SharePoint sin
 * depender de que alguien abra el dashboard. Disparado por Vercel Cron
 * (ver vercel.json); protegido con CRON_SECRET, que Vercel envia como
 * "Authorization: Bearer <CRON_SECRET>" en cada invocacion programada.
 */
export async function GET(request: Request) {
  const authHeader = request.headers.get("authorization");
  if (!process.env.CRON_SECRET || authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  // Idempotente (solo escribe las celdas que faltan): ante una falla
  // transitoria de SAP o SharePoint se reintenta en la misma invocacion.
  let last: unknown = { success: false, error: "sin ejecutar" };
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const { sapPayments, vendorPayments } = await fetchSapPayments();
      const result = await runFullSync(sapPayments, vendorPayments);
      if (result.success) return NextResponse.json({ ...result, attempt });
      last = result;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } catch (error: any) {
      console.error(`Error en cron de auto-sync (intento ${attempt}):`, error);
      last = { success: false, error: error.message || String(error) };
    }
    if (attempt < 3) await new Promise((r) => setTimeout(r, 30_000));
  }
  return NextResponse.json(last, { status: 500 });
}
