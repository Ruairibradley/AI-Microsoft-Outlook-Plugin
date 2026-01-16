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

export type SignedInUser = {
  username?: string; // typically UPN/email
  name?: string;
};

/**
 * Returns the currently cached MSAL account identity (if any).
 * Useful for displaying "Signed in as".
 */
export async function get_signed_in_user(): Promise<SignedInUser | null> {
  await ensure_msal_initialized();
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;

  const a = accounts[0];
  return {
    username: a.username,
    name: a.name
  };
}

/**
 * Signs out the cached account from MSAL.
 * After this, silent token acquisition will fail and the user must sign in again.
 */
export async function sign_out(): Promise<void> {
  await ensure_msal_initialized();

  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return;

  await msalInstance.logoutPopup({
    account: accounts[0]
  });
}

/**
 * Forces an account picker prompt (useful for "Switch account").
 * If no cached account, it behaves like sign-in but with selection enforced.
 */
export async function get_access_token_select_account(): Promise<string> {
  await ensure_msal_initialized();

  const accounts = msalInstance.getAllAccounts();

  if (!accounts.length) {
    const loginRes = await msalInstance.loginPopup({
      scopes: graphScopes,
      prompt: "select_account"
    });

    const tokenRes = await msalInstance.acquireTokenSilent({
      account: loginRes.account!,
      scopes: graphScopes
    });

    return tokenRes.accessToken;
  }

  // If there is a cached account, still force a selection.
  const tokenRes = await msalInstance.acquireTokenPopup({
    account: accounts[0],
    scopes: graphScopes,
    prompt: "select_account"
  });

  return tokenRes.accessToken;
}
