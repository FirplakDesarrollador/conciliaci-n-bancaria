import "server-only";
import { getGraphAccessToken } from "./token";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export async function graphFetch<T = unknown>(path: string, init: RequestInit = {}): Promise<T> {
  const accessToken = await getGraphAccessToken();

  const res = await fetch(`${GRAPH_BASE_URL}${path}`, {
    ...init,
    headers: {
      ...init.headers,
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    cache: "no-store",
  });

  if (!res.ok) {
    const body = await res.json().catch(() => null);
    throw new Error(`Graph API error ${res.status}: ${body?.error?.message ?? res.statusText}`);
  }

  if (res.status === 204) {
    return undefined as T;
  }

  return res.json();
}
