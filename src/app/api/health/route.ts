import { NextResponse } from "next/server";

export async function GET() {
  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  if (!supabaseUrl || !supabaseAnonKey) {
    return NextResponse.json(
      { status: "error", message: "Faltan NEXT_PUBLIC_SUPABASE_URL o NEXT_PUBLIC_SUPABASE_ANON_KEY" },
      { status: 500 }
    );
  }

  try {
    // Consulta una tabla que no existe aún: un 404 (PGRST205) confirma que la
    // apikey es válida y que PostgREST respondió; un 401 indica apikey inválida.
    const res = await fetch(`${supabaseUrl}/rest/v1/_healthcheck_?select=*&limit=1`, {
      headers: {
        apikey: supabaseAnonKey,
        Authorization: `Bearer ${supabaseAnonKey}`,
      },
      cache: "no-store",
    });
    const body = await res.json().catch(() => null);

    return NextResponse.json({
      status: "ok",
      supabaseReachable: res.status !== 401,
      supabaseStatus: res.status,
      detail: body,
    });
  } catch (error) {
    return NextResponse.json(
      { status: "error", message: (error as Error).message },
      { status: 500 }
    );
  }
}
