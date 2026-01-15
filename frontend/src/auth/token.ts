import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { ensure_msal_initialized, msalInstance, graphScopes } from "./msal";

/**
 * Interactive token acquisition:
 * - May open a popup window if required.
 * - Use this only on a direct user action (button click).
 */
export async function get_access_token(): Promise<string> {
  await ensure_msal_initialized();

  const accounts = msalInstance.getAllAccounts();

  if (!accounts.length) {
    const loginRes = await msalInstance.loginPopup({ scopes: graphScopes });
    const tokenRes = await msalInstance.acquireTokenSilent({
      account: loginRes.account!,
      scopes: graphScopes
    });
    return tokenRes.accessToken;
  }

  const account = accounts[0];

  try {
    const silent = await msalInstance.acquireTokenSilent({ account, scopes: graphScopes });
    return silent.accessToken;
  } catch (e: any) {
    if (e instanceof InteractionRequiredAuthError) {
      const popup = await msalInstance.acquireTokenPopup({ account, scopes: graphScopes });
      return popup.accessToken;
    }
    throw e;
  }
}

/**
 * Silent-only token acquisition:
 * - Never opens a popup.
 * - Returns null if the user must interact (e.g., not signed in).
 */
export async function try_get_access_token_silent(): Promise<string | null> {
  await ensure_msal_initialized();

  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;

  try {
    const silent = await msalInstance.acquireTokenSilent({ account: accounts[0], scopes: graphScopes });
    return silent.accessToken;
  } catch (e: any) {
    if (e instanceof InteractionRequiredAuthError) return null;
    throw e;
  }
}
