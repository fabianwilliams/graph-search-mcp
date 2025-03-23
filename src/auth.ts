// Author: Fabian G. Williams
// Purpose: Handles MSAL token acquisition for Microsoft Graph

import { ConfidentialClientApplication } from '@azure/msal-node';

let msalApp: ConfidentialClientApplication;

/**
 * Initialize MSAL app with environment-provided credentials
 */
export function initMsal(
  tenantId: string,
  clientId: string,
  clientSecret: string,
) {
  msalApp = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
  });
}

/**
 * Acquire an access token for Microsoft Graph
 */
export async function getToken(): Promise<string> {
  const token = await msalApp.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });

  return token?.accessToken || '';
}
