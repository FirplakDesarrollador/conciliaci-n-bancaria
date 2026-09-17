import "server-only";
import { sapClient } from "@/lib/sap/service-layer";
import { getTargetDates } from "./target-dates";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function isValidPayment(p: any): boolean {
  const jr = p.JournalRemarks ? String(p.JournalRemarks).trim().toUpperCase() : "";
  const rm = p.Remarks ? String(p.Remarks).trim().toUpperCase() : "";
  return jr !== "CANCELADO" && rm !== "CANCELADO" && p.Cancelled !== "tYES";
}

function nextLinkOf(data: { "odata.nextLink"?: string }): string | null {
  const link = data["odata.nextLink"];
  if (!link) return null;
  return link.startsWith("/") ? link : `/${link}`;
}

/**
 * Trae IncomingPayments + VendorPayments de SAP para la ventana de fechas de
 * getTargetDates(), filtrando los cancelados. Usado tanto por el dashboard
 * (page.tsx) como por el cron de auto-sync (api/cron/sync), para que ambos
 * vean exactamente los mismos documentos.
 */
export async function fetchSapPayments() {
  const dates = getTargetDates();
  const filterStr = dates.map((d) => `DocDate eq '${d}'`).join(" or ");

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let sapPayments: any[] = [];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let vendorPayments: any[] = [];

  let nextLink: string | null = `/IncomingPayments?$filter=${filterStr}&$orderby=DocNum`;
  while (nextLink) {
    const res = await sapClient.request(nextLink);
    if (!res.ok) {
      console.error("Error fetching SAP incoming payments:", await res.text());
      break;
    }
    const data = await res.json();
    sapPayments = sapPayments.concat((data.value || []).filter(isValidPayment));
    nextLink = nextLinkOf(data);
  }

  let vendorNextLink: string | null = `/VendorPayments?$filter=${filterStr}&$orderby=DocNum`;
  while (vendorNextLink) {
    const res = await sapClient.request(vendorNextLink);
    if (!res.ok) {
      console.error("Error fetching SAP vendor payments:", await res.text());
      break;
    }
    const data = await res.json();
    vendorPayments = vendorPayments.concat((data.value || []).filter(isValidPayment));
    vendorNextLink = nextLinkOf(data);
  }

  return { sapPayments, vendorPayments };
}
