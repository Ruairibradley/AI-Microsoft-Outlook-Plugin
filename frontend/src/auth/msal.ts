import { PublicClientApplication } from "@azure/msal-browser";

const clientId = import.meta.env.VITE_MSAL_CLIENT_ID as string;
const authority =
  (import.meta.env.VITE_MSAL_AUTHORITY as string) || "https://login.microsoftonline.com/common";
const redirectUri = import.meta.env.VITE_MSAL_REDIRECT_URI as string;

export const graphScopes = ["User.Read", "Mail.Read"];

export const msalInstance = new PublicClientApplication({
  auth: { clientId, authority, redirectUri },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
});

/**
 * Newer MSAL versions require explicit initialization.
 * Call this once before using any MSAL APIs.
 */
let _initPromise: Promise<void> | null = null;

export function ensure_msal_initialized(): Promise<void> {
  if (!_initPromise) {
    _initPromise = msalInstance.initialize();
  }
  return _initPromise;
}
