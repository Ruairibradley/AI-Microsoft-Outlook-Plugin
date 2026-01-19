import { useEffect, useState } from "react";
import "./styles.css";

import { sign_in_interactive, try_get_access_token_silent, get_signed_in_user, sign_out } from "../auth/token";
import { get_folders, get_index_status, type IndexStatus } from "../api/backend";

import { IndexManager } from "./components/indexmanager";
import { ChatPane } from "./components/chatpane"; 
import { clear_chat_storage } from "./chat/storage";

type Screen = "SIGNIN" | "CHAT" | "INDEX";

type Folder = {
  id: string;
  displayName: string;
  totalItemCount?: number;
};

function fmt_dt(s: string | null | undefined) {
  if (!s) return "—";
  return s.replace("T", " ");
}

export default function TaskPaneView() {
  // ---------- app/screen ----------
  const [screen, setScreen] = useState<Screen>("SIGNIN");

  const [status, setStatus] = useState("Starting…");
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

  async function refresh_index_status(nextPreferredScreen: Screen | null = null) {
    const st = (await get_index_status()) as IndexStatus;
    setIndexStatus(st);

    const exists = (st?.indexed_count || 0) > 0;

    if (!exists) {
      setScreen("INDEX");
      return;
    }

    if (nextPreferredScreen) setScreen(nextPreferredScreen);
    else if (screen === "SIGNIN") setScreen("CHAT");
  }

  async function refresh_folders(token: string) {
    const data = await get_folders(token);
    setFolders((data.folders || []) as Folder[]);
  }

  // Called by IndexManager after ingest/clear operations
  async function on_index_changed() {
    // Re-load index status and keep routing deterministic
    await refresh_index_status(screen === "SIGNIN" ? "CHAT" : null);
    // Also refresh ingestion list inside IndexManager itself (it does that on open), so no need here.
  }

  // ---------- startup ----------
  async function initialize() {
    set_error("");
    setStatus("Checking session…");
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
        setStatus("Not signed in");
        return;
      }

      setTokenOk(true);
      setAccessToken(token);

      const who = await get_signed_in_user();
      setUserLabel(who?.username || who?.name || "");

      setStatus("Loading local index status…");
      await refresh_index_status();

      setStatus("Loading folders…");
      await refresh_folders(token);

      setStatus("Ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);
      setScreen("SIGNIN");
      setStatus("Error");
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

      setStatus("Loading local index status…");
      await refresh_index_status();

      setStatus("Loading folders…");
      await refresh_folders(token);

      setStatus("Ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);
      setScreen("SIGNIN");
      setStatus("Error");
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

      // Reset all local state
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);

      // Prevent cross-account chat restoration confusion
      clear_chat_storage();

      setScreen("SIGNIN");
      setStatus("Not signed in");
    } catch (e: any) {
      setStatus("Error");
      set_error("Sign out failed.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- UI blocks ----------
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
    return (
      <div className="op-header">
        <div className="op-title">
          Outlook Privacy Assistant
          <span className="op-badge">Local</span>
        </div>

        <div className="op-subline">
          <span><strong>Status:</strong> {status}</span>
          {busy ? <span>• {busy}</span> : null}
          {token_ok && user_label ? <span>• <strong>Signed in:</strong> {user_label}</span> : null}
          {token_ok && index_status ? <span>• <strong>Indexed:</strong> {index_status.indexed_count} • <strong>Updated:</strong> {fmt_dt(index_status.last_updated)}</span> : null}
        </div>

        {token_ok ? (
          <div className="op-nav">
            <button className="op-btn op-btnGhost" onClick={() => setScreen("CHAT")} disabled={!index_exists}>
              Chat
            </button>
            <button className="op-btn op-btnGhost" onClick={() => setScreen("INDEX")}>
              Index
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
    return (
      <div className="op-card op-fit">
        <div className="op-cardHeader">
          <div>
            <div className="op-cardTitle">Sign in</div>
            <div className="op-muted">Sign in to fetch emails with Microsoft Graph and index locally.</div>
          </div>
        </div>
        <div className="op-cardBody">
          <button className="op-btn op-btnPrimary" onClick={() => sign_in_clicked()}>
            Sign in
          </button>
          <div className="op-helpNote">
            Tip: If email links fail in your browser, ensure you’re signed into the same mailbox in Outlook on the web.
          </div>
        </div>
      </div>
    );
  }

  function render_index_screen() {
    return (
      <IndexManager
        token_ok={token_ok}
        access_token={access_token}
        folders={folders}
        index_status={index_status}
        onIndexChanged={on_index_changed}
        collapsible={false}
      />
    );
  }

  function render_chat_screen() {
    if (!index_exists) {
      // Deterministic: if index is empty, force index screen
      return (
        <div className="op-fit">
          <div className="op-banner op-bannerStrong">
            <div className="op-bannerTitle">Index required</div>
            <div className="op-bannerText">Index emails first using Index management.</div>
          </div>
          <div className="op-spacer" />
          {render_index_screen()}
        </div>
      );
    }

    return (
      <div className="op-fit" style={{ height: "100%" }}>
        <IndexManager
          token_ok={token_ok}
          access_token={access_token}
          folders={folders}
          index_status={index_status}
          onIndexChanged={on_index_changed}
          collapsible={true}
        />
        <div className="op-spacer" />
        <div style={{ flex: 1, minHeight: 0 }}>
          <ChatPane
            token_ok={token_ok}
            access_token={access_token}
            index_exists={index_exists}
            index_count={index_status?.indexed_count || 0}
          />
        </div>
      </div>
    );
  }

  // ---------- main ----------
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
