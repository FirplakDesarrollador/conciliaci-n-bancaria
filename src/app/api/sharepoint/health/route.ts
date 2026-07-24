import { NextResponse } from "next/server";
import { graphFetch } from "@/lib/graph/client";

interface SitesSearchResponse {
  value: Array<{ id: string; name: string; displayName: string; webUrl: string }>;
}

export async function GET() {
  try {
    // Sites.Read.All / Sites.ReadWrite.All permiten buscar en todos los sitios;
    // usamos esto como prueba de extremo a extremo del token + permisos.
    const data = await graphFetch<SitesSearchResponse>("/sites?search=*");

    return NextResponse.json({
      status: "ok",
      graphReachable: true,
      sitesFound: data.value?.length ?? 0,
    });
  } catch (error) {
    return NextResponse.json(
      { status: "error", graphReachable: false, message: (error as Error).message },
      { status: 500 }
    );
  }
}
