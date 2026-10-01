import { NextResponse } from "next/server";
import { buildConsecutivoWorkbook } from "@/lib/conciliacion/consecutivo";
import { uploadDriveFile } from "@/lib/graph/sharepoint";

export const maxDuration = 60;

const OUTPUT_FILE = "CONTROL_CONSECUTIVO_PAGOS.xlsx";

/**
 * Regenera el archivo de control de consecutivo de pagos recibidos y
 * efectuados (ver src/lib/conciliacion/consecutivo.ts) y lo sube a
 * SharePoint, junto a los 5 archivos de banco. Corre por separado del cron
 * de sincronizacion (api/cron/sync) porque ese ya usa casi todo el tiempo
 * disponible de la funcion; sumarle otra consulta paginada a SAP arriesgaria
 * los dos. Protegido con el mismo CRON_SECRET.
 */
export async function GET(request: Request) {
  const authHeader = request.headers.get("authorization");
  if (!process.env.CRON_SECRET || authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  try {
    const buffer = await buildConsecutivoWorkbook();
    await uploadDriveFile(OUTPUT_FILE, buffer);
    return NextResponse.json({ success: true, file: OUTPUT_FILE });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    console.error("Error generando control de consecutivo:", error);
    return NextResponse.json({ success: false, error: error.message || String(error) }, { status: 500 });
  }
}
