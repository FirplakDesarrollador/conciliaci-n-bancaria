import "server-only";
import { ConfidentialClientApplication } from "@azure/msal-node";

const tenantId = process.env.AZURE_AD_TENANT_ID;
const clientId = process.env.AZURE_AD_CLIENT_ID;
const clientSecret = process.env.AZURE_AD_CLIENT_SECRET;

if (!tenantId || !clientId || !clientSecret) {
  throw new Error(
    "Faltan AZURE_AD_TENANT_ID, AZURE_AD_CLIENT_ID o AZURE_AD_CLIENT_SECRET. Defínelas en .env.local."
  );
}

const msalApp = new ConfidentialClientApplication({
  auth: {
    clientId,
    clientSecret,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
});

export async function getGraphAccessToken() {
  const result = await msalApp.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  if (!result?.accessToken) {
    throw new Error("No se pudo obtener un access token de Microsoft Graph.");
  }

  return result.accessToken;
}
