import { NextResponse } from "next/server";
import { buildConsecutivoWorkbook, controlFileName } from "@/lib/conciliacion/consecutivo";
import { driveFolderExists, folderPathForYear, uploadDriveFile } from "@/lib/graph/sharepoint";

export const maxDuration = 300;

function bogotaYear(): number {
  return Number(new Intl.DateTimeFormat("en-CA", { timeZone: "America/Bogota", year: "numeric" }).format(new Date()));
}

/**
 * Regenera el archivo de control de consecutivo de pagos recibidos y
 * efectuados del ano en curso (ver src/lib/conciliacion/consecutivo.ts) y lo
 * sube a la carpeta anual de SharePoint ("FIRPLAK <ano>"), junto a los
 * archivos de banco de ese ano. Cada ano se genera su propio archivo en su
 * propia carpeta; si la carpeta del ano nuevo aun no existe no se crea
 * ninguna: se reporta y se reintenta en la siguiente corrida. Corre por
 * separado del cron de sincronizacion (api/cron/sync) porque ese ya usa casi
 * todo el tiempo disponible de la funcion. Protegido con el mismo CRON_SECRET.
 */
export async function GET(request: Request) {
  const authHeader = request.headers.get("authorization");
  if (!process.env.CRON_SECRET || authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  try {
    const year = bogotaYear();
    const folder = folderPathForYear(year);
    if (!(await driveFolderExists(folder))) {
      return NextResponse.json({ success: false, error: `No existe la carpeta de SharePoint del ano ${year}: ${folder}` }, { status: 409 });
    }
    const buffer = await buildConsecutivoWorkbook(year);
    const file = controlFileName(year);
    await uploadDriveFile(file, buffer, folder);
    return NextResponse.json({ success: true, file, folder });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    console.error("Error generando control de consecutivo:", error);
    const msg: string = error.message || String(error);
    if (msg.includes("HTTP 423")) {
      return NextResponse.json(
        { success: false, error: "El archivo de control esta abierto/bloqueado en SharePoint (423); se reintenta en la siguiente corrida." },
        { status: 423 }
      );
    }
    return NextResponse.json({ success: false, error: msg }, { status: 500 });
  }
}
