// File: src/taskpane/authConfig.ts

import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  SilentRequest,
} from "@azure/msal-browser";

export const msalConfig = {
  auth: {
    clientId:    "a4d1fb6c-f9df-4caf-a091-a2b93b078ddc",
    authority:   "https://login.microsoftonline.com/baac284d-b86c-41d1-9728-bd7b24a4d0eb",
    redirectUri: "https://bdf2-92-64-101-76.ngrok-free.app/taskpane.html",
  },
  cache: {
    cacheLocation:       "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: [
    "User.Read",
    "Contacts.ReadWrite",
    "Sites.ReadWrite.All"       // ← ADD SharePoint scope here
  ],
};

const msalInstance = new PublicClientApplication(msalConfig);

/**
 * Acquires a Graph access token via MSAL.
 * Calls initialize() first, then tries silent, then popup.
 */
export async function getGraphToken(): Promise<string> {
  // Must initialize before anything else
  await msalInstance.initialize();

  // 1) Ensure an account is set
  let account = msalInstance.getActiveAccount();
  if (!account) {
    const loginResp = await msalInstance.loginPopup(loginRequest);
    account = loginResp.account!;
    msalInstance.setActiveAccount(account);
  }

  // 2) Try silent
  const silentRequest: SilentRequest = {
    scopes:  loginRequest.scopes,
    account,
  };

  try {
    const tokenResp = await msalInstance.acquireTokenSilent(silentRequest);
    return tokenResp.accessToken;
  } catch (err: any) {
    if (err instanceof InteractionRequiredAuthError) {
      // 3) Fallback to popup
      const tokenResp = await msalInstance.acquireTokenPopup(loginRequest);
      return tokenResp.accessToken;
    }
    throw err;
  }
}
