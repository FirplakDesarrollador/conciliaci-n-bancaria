import { NextResponse } from "next/server";
import { syncConsecutivoFile } from "@/lib/conciliacion/consecutivo-sync";
import { driveFolderExists, folderPathForYear } from "@/lib/graph/sharepoint";

export const maxDuration = 300;

function bogotaYear(): number {
  return Number(new Intl.DateTimeFormat("en-CA", { timeZone: "America/Bogota", year: "numeric" }).format(new Date()));
}

/**
 * Actualiza el archivo de control de consecutivo de pagos recibidos y
 * efectuados del ano en curso, que vive en la carpeta anual de SharePoint
 * ("FIRPLAK <ano>"). El archivo existente NUNCA se reemplaza: solo se
 * actualizan los estados de las filas que ya estan y se agregan los
 * documentos nuevos, para respetar las marcas manuales de verificacion
 * (ver src/lib/conciliacion/consecutivo-sync.ts). Si la carpeta del ano
 * nuevo aun no existe no se crea nada: se reporta y se reintenta en la
 * siguiente corrida. Protegido con CRON_SECRET.
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
    const result = await syncConsecutivoFile(year);
    return NextResponse.json({ success: true, folder, ...result });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    console.error("Error actualizando control de consecutivo:", error);
    return NextResponse.json({ success: false, error: error.message || String(error) }, { status: 500 });
  }
}
