import { useEffect, useState } from "react";
import "./styles.css";

import { sign_in_interactive, try_get_access_token_silent, get_signed_in_user, sign_out } from "../auth/token";
import { get_folders, get_index_status, type IndexStatus } from "../api/backend";

import { IndexManager } from "./components/indexmanager";
import { ChatPane } from "./components/chatpane";
import { clear_chat_storage } from "./chat/storage";
import { WelcomeSignIn } from "./components/welcomesignin";

type Screen = "SIGNIN" | "CHAT" | "INDEX";

type Folder = {
  id: string;
  displayName: string;
  totalItemCount?: number;
};

// function format_dt_uk(iso: string | null | undefined): string {
//   if (!iso) return "—";
//   try {
//     const [datePart, timePartRaw] = iso.split("T");
//     if (!datePart) return iso;

//     const [yyyy, mm, dd] = datePart.split("-");
//     const timePart = (timePartRaw || "").slice(0, 5); // HH:MM
//     if (!yyyy || !mm || !dd) return iso;

//     return `${dd}/${mm}/${yyyy} ${timePart || ""}`.trim();
//   } catch {
//     return iso;
//   }
// }

export default function TaskPaneView() {
  const [screen, setScreen] = useState<Screen>("SIGNIN");

  const [busy, setBusy] = useState<string>("");

  const [error, setError] = useState<string>("");
  const [error_details, setErrorDetails] = useState<string>("");

  function set_error(msg: string, details: string = "") {
    setError(msg);
    setErrorDetails(details);
  }

  // ---------- auth ----------
  const [token_ok, setTokenOk] = useState(false);
  const [access_token, setAccessToken] = useState<string>("");
  const [user_label, setUserLabel] = useState<string>("");

  // ---------- backend state ----------
  const [index_status, setIndexStatus] = useState<IndexStatus | null>(null);
  const [folders, setFolders] = useState<Folder[]>([]);

  const index_exists = (index_status?.indexed_count || 0) > 0;

  async function refresh_index_status(
    nextPreferredScreen: Screen | null = null,
    opts: { autoNavigate?: boolean } = { autoNavigate: true }
  ) {
    const st = (await get_index_status()) as IndexStatus;
    setIndexStatus(st);

    const exists = (st?.indexed_count || 0) > 0;

    if (!exists) {
      // If nothing is indexed, always keep user on INDEX
      setScreen("INDEX");
      return;
    }

    // Explicit navigation request wins
    if (nextPreferredScreen) {
      setScreen(nextPreferredScreen);
      return;
    }

    // Only auto-navigate during startup/sign-in scenarios.
    if (opts.autoNavigate) {
      if (screen === "SIGNIN") setScreen("CHAT");
    }
  }

  async function refresh_folders(token: string) {
    const data = await get_folders(token);
    setFolders((data.folders || []) as Folder[]);
  }

  // Called by IndexManager after indexing/cancel/clear etc.
  async function on_index_changed() {
    // do not auto-navigate to CHAT on an index refresh
    await refresh_index_status(null, { autoNavigate: false });
  }

  // ---------- startup ----------
  async function initialize() {
    set_error("");
    setBusy("");

    try {
      const token = await try_get_access_token_silent();
      if (!token) {
        setTokenOk(false);
        setAccessToken("");
        setUserLabel("");
        setIndexStatus(null);
        setFolders([]);
        setScreen("SIGNIN");
        return;
      }

      setTokenOk(true);
      setAccessToken(token);

      const who = await get_signed_in_user();
      setUserLabel(who?.username || who?.name || "");

      await refresh_index_status(null, { autoNavigate: true });
      await refresh_folders(token);
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);
      setScreen("SIGNIN");
      set_error("Initialization failed. Please sign in again.", String(e?.message || e));
    }
  }

  useEffect(() => {
    initialize().catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ---------- auth actions ----------
  async function sign_in_clicked() {
    set_error("");
    setBusy("Signing in…");

    try {
      const token = await sign_in_interactive();
      setTokenOk(true);
      setAccessToken(token);

      const who = await get_signed_in_user();
      setUserLabel(who?.username || who?.name || "");

      await refresh_index_status(null, { autoNavigate: true });
      await refresh_folders(token);
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);
      setScreen("SIGNIN");
      set_error("Sign in failed. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function sign_out_clicked() {
    set_error("");
    setBusy("Signing out…");
    try {
      await sign_out();

      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);

      clear_chat_storage();
      setScreen("SIGNIN");
    } catch (e: any) {
      set_error("Sign out failed.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- UI ----------
  function render_error() {
    if (!error) return null;
    return (
      <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
        <div className="op-bannerTitle">Something went wrong</div>
        <div className="op-bannerText">{error}</div>
        {error_details ? (
          <details style={{ marginTop: 8 }}>
            <summary style={{ cursor: "pointer", fontSize: 12 }}>Show details</summary>
            <pre style={{ whiteSpace: "pre-wrap", marginTop: 8, fontSize: 12 }}>{error_details}</pre>
          </details>
        ) : null}
      </div>
    );
  }

  function render_header() {
    if (screen === "SIGNIN") return null;

    return (
      <div className="op-header">
        <div className="op-subline">
          {token_ok && user_label ? (
            <span>
              <strong>Signed in:</strong> {user_label}
            </span>
          ) : null}

          {busy ? <span>• {busy}</span> : null}
        </div>

        {token_ok ? (
          <div className="op-nav">
            <button className="op-btn op-btnGhost" onClick={() => setScreen("CHAT")} disabled={!index_exists}>
              Chat
            </button>
            <button className="op-btn op-btnGhost" onClick={() => setScreen("INDEX")}>
              Emails
            </button>
            <button className="op-btn op-btnDanger" onClick={() => sign_out_clicked()}>
              Sign out
            </button>
          </div>
        ) : null}

        {render_error()}
      </div>
    );
  }

  function render_signin() {
    return <WelcomeSignIn onSignIn={() => sign_in_clicked()} error={error} error_details={error_details} />;
  }

  function render_index_screen() {
    return (
      <IndexManager
        token_ok={token_ok}
        access_token={access_token}
        folders={folders}
        index_status={index_status}
        onIndexChanged={on_index_changed}
        onNavigate={(to) => setScreen(to)}
      />
    );
  }

  function render_chat_screen() {
    if (!index_exists) return render_index_screen();

    return (
      <div className="op-fit" style={{ height: "100%" }}>
        <ChatPane
          token_ok={token_ok}
          access_token={access_token}
          index_exists={index_exists}
          index_count={index_status?.indexed_count || 0}
          index_last_updated={index_status?.last_updated || null}
        />
      </div>
    );
  }

  return (
    <div className="op-app">
      {render_header()}
      <div className="op-body">
        {screen === "SIGNIN" && render_signin()}
        {screen === "INDEX" && render_index_screen()}
        {screen === "CHAT" && render_chat_screen()}
      </div>
    </div>
  );
}
