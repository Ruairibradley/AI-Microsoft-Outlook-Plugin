import { InteractionRequiredAuthError, AccountInfo } from "@azure/msal-browser";
import { ensure_msal_initialized, msalInstance, graphScopes } from "./msal";

const LAST_USER_KEY = "msal_last_user";

/**
 * Persist the last-used account so silent login uses the right one in multi-account scenarios.
 */
function set_last_user(username: string | undefined) {
  if (!username) return;
  try {
    localStorage.setItem(LAST_USER_KEY, username);
  } catch {
    // ignore storage failures
  }
}

function get_last_user(): string | null {
  try {
    return localStorage.getItem(LAST_USER_KEY);
  } catch {
    return null;
  }
}

function clear_last_user() {
  try {
    localStorage.removeItem(LAST_USER_KEY);
  } catch {
    // ignore
  }
}

function pick_account(accounts: AccountInfo[]): AccountInfo | null {
  if (!accounts.length) return null;

  // Prefer MSAL's active account if set.
  const active = msalInstance.getActiveAccount();
  if (active) return active;

  // Prefer the last user we recorded.
  const last = get_last_user();
  if (last) {
    const match = accounts.find((a) => a.username?.toLowerCase() === last.toLowerCase());
    if (match) return match;
  }

  // Fallback: first account in cache.
  return accounts[0];
}

/**
 * Interactive sign-in/token acquisition:
 * - Always prompts account selection to avoid "wrong cached account" surprises.
 * - Sets active account and persists "last used user".
 */
export async function sign_in_interactive(): Promise<string> {
  await ensure_msal_initialized();

  const loginRes = await msalInstance.loginPopup({
    scopes: graphScopes,
    prompt: "select_account"
  });

  if (loginRes.account) {
    msalInstance.setActiveAccount(loginRes.account);
    set_last_user(loginRes.account.username);
  }

  const tokenRes = await msalInstance.acquireTokenSilent({
    account: loginRes.account!,
    scopes: graphScopes
  });

  return tokenRes.accessToken;
}

/**
 * Try silent token acquisition:
 * - Uses active/last-used account preference.
 * - Returns null if user interaction is required.
 */
export async function try_get_access_token_silent(): Promise<string | null> {
  await ensure_msal_initialized();

  const accounts = msalInstance.getAllAccounts();
  const account = pick_account(accounts);
  if (!account) return null;

  try {
    const silent = await msalInstance.acquireTokenSilent({ account, scopes: graphScopes });

    // Keep active/last-used in sync.
    msalInstance.setActiveAccount(account);
    set_last_user(account.username);

    return silent.accessToken;
  } catch (e: any) {
    if (e instanceof InteractionRequiredAuthError) return null;
    throw e;
  }
}

export type SignedInUser = {
  username?: string;
  name?: string;
};

/**
 * Returns the currently selected/cached identity for display.
 */
export async function get_signed_in_user(): Promise<SignedInUser | null> {
  await ensure_msal_initialized();

  const accounts = msalInstance.getAllAccounts();
  const account = pick_account(accounts);
  if (!account) return null;

  // Ensure MSAL also considers this the active account.
  msalInstance.setActiveAccount(account);

  return {
    username: account.username,
    name: account.name
  };
}

/**
 * Sign out:
 * - Logs out the selected account from MSAL cache for this app
 * - Clears our last-used marker so the next session does NOT auto-select an account
 */
export async function sign_out(): Promise<void> {
  await ensure_msal_initialized();

  const accounts = msalInstance.getAllAccounts();
  const account = pick_account(accounts);
  if (!account) {
    clear_last_user();
    return;
  }

  clear_last_user();

  await msalInstance.logoutPopup({
    account
  });
}
