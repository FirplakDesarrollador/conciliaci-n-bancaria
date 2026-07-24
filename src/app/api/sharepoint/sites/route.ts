import { NextRequest, NextResponse } from "next/server";
import { graphFetch } from "@/lib/graph/client";

interface SitesSearchResponse {
  value: Array<{ id: string; name: string; displayName: string; webUrl: string }>;
}

// Utilidad de exploración: lista los sitios de SharePoint visibles para la app,
// para poder identificar el site id / webUrl que se vaya a usar.
export async function GET(request: NextRequest) {
  const query = request.nextUrl.searchParams.get("q") ?? "*";

  try {
    const data = await graphFetch<SitesSearchResponse>(`/sites?search=${encodeURIComponent(query)}`);

    return NextResponse.json({
      status: "ok",
      sites: data.value.map((site) => ({
        id: site.id,
        name: site.name,
        displayName: site.displayName,
        webUrl: site.webUrl,
      })),
    });
  } catch (error) {
    return NextResponse.json({ status: "error", message: (error as Error).message }, { status: 500 });
  }
}
