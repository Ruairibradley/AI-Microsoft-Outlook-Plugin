import { useEffect, useMemo, useState } from "react";
import { sign_in_interactive, try_get_access_token_silent, get_signed_in_user, sign_out } from "../auth/token";
import {
  ask_question,
  clear_index,
  get_folders,
  get_messages,
  ingest_messages,
  get_index_status,
  type IndexStatus
} from "../api/backend";

type Screen = "SIGNIN" | "CHAT" | "INDEX";

type folder = {
  id: string;
  displayName: string;
  totalItemCount?: number;
};

type graph_message = {
  id: string;
  subject?: string;
  receivedDateTime?: string;
  webLink?: string;
  bodyPreview?: string;
  from?: { emailAddress?: { address?: string } };
};

type source = {
  message_id: string;
  weblink: string;
  subject: string;
  sender: string;
  received_dt: string;
  snippet: string;
  score?: number;
};

function truncate(s: string, n: number) {
  if (!s) return "";
  return s.length <= n ? s : s.slice(0, n) + "…";
}

function fmt_dt(s: string | null | undefined) {
  if (!s) return "—";
  return s.replace("T", " ");
}

export default function TaskPaneView() {
  // ---------- global UI state ----------
  const [screen, setScreen] = useState<Screen>("SIGNIN");

  const [status, setStatus] = useState("starting...");
  const [error, setError] = useState<string>("");
  const [error_details, setErrorDetails] = useState<string>("");

  const [busy, setBusy] = useState<string>("");

  // ---------- auth ----------
  const [token_ok, setTokenOk] = useState(false);
  const [access_token, setAccessToken] = useState<string>("");
  const [user_label, setUserLabel] = useState<string>("");

  // ---------- index status ----------
  const [index_status, setIndexStatus] = useState<IndexStatus | null>(null);
  const [index_panel_open, setIndexPanelOpen] = useState<boolean>(false);

  // ---------- ingestion selection (current working UI; Phase 3 will replace with wizard) ----------
  const [folders, setFolders] = useState<folder[]>([]);
  const [folder_id, setFolderId] = useState("");
  const [messages, setMessages] = useState<graph_message[]>([]);
  const [selected_ids, setSelectedIds] = useState<Set<string>>(new Set());

  const selected_count = selected_ids.size;

  const [show_consent, setShowConsent] = useState(false);
  const [consent_checked, setConsentChecked] = useState(false);

  // ---------- chat (current working query/answer; Phase 4 will add chat history) ----------
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");
  const [sources, setSources] = useState<source[]>([]);

  // ---------- helpers ----------
  function set_error(msg: string, details: string = "") {
    setError(msg);
    setErrorDetails(details);
  }

  async function refresh_index_status() {
    const st = (await get_index_status()) as IndexStatus;
    setIndexStatus(st);

    // If the index is empty, force Index Management screen.
    if ((st?.indexed_count || 0) <= 0) {
      setScreen("INDEX");
      setIndexPanelOpen(true);
    } else {
      // If we are signed in and index exists, default to Chat.
      // Do not forcibly override if the user is already on INDEX intentionally.
      if (screen === "SIGNIN" || screen === "CHAT") {
        setScreen("CHAT");
      }
    }
  }

  async function load_folders_with_token(token: string) {
    setBusy("loading folders...");
    const data = await get_folders(token);
    setFolders((data.folders || []) as folder[]);
    setBusy("");
  }

  // ---------- startup ----------
  async function initialize() {
    set_error("");
    setStatus("checking session...");
    setBusy("");

    try {
      // 1) silent auth attempt
      const token = await try_get_access_token_silent();
      if (!token) {
        setTokenOk(false);
        setAccessToken("");
        setUserLabel("");
        setIndexStatus(null);
        setScreen("SIGNIN");
        setStatus("not signed in");
        return;
      }

      setTokenOk(true);
      setAccessToken(token);

      const who = await get_signed_in_user();
      setUserLabel(who?.username || who?.name || "");

      // 2) load index status
      setStatus("loading local index status...");
      await refresh_index_status();

      // 3) load folders (needed for ingestion screen/panel)
      setStatus("loading folders...");
      await load_folders_with_token(token);

      setStatus("ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setScreen("SIGNIN");
      setStatus("error");
      set_error("Initialization failed. Please try signing in again.", String(e?.message || e));
    }
  }

  useEffect(() => {
    initialize().catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ---------- auth actions ----------
  async function sign_in_clicked() {
    set_error("");
    setBusy("signing in...");
    try {
      const token = await sign_in_interactive();
      setTokenOk(true);
      setAccessToken(token);

      const who = await get_signed_in_user();
      setUserLabel(who?.username || who?.name || "");

      setStatus("loading local index status...");
      await refresh_index_status();

      setStatus("loading folders...");
      await load_folders_with_token(token);

      setStatus("ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setScreen("SIGNIN");
      setStatus("error");
      set_error("Sign in failed. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function sign_out_clicked() {
    set_error("");
    setBusy("signing out...");
    try {
      await sign_out();
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);
      setFolderId("");
      setMessages([]);
      setSelectedIds(new Set());
      setShowConsent(false);
      setConsentChecked(false);
      setQuestion("");
      setAnswer("");
      setSources([]);
      setScreen("SIGNIN");
      setStatus("not signed in");
    } catch (e: any) {
      setStatus("error");
      set_error("Sign out failed.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- ingestion selection ----------
  async function load_messages(fid: string) {
    setFolderId(fid);
    setMessages([]);
    setSelectedIds(new Set());
    setShowConsent(false);
    setConsentChecked(false);

    if (!fid) return;
    if (!token_ok || !access_token) return;

    setBusy("loading messages...");
    try {
      const data = await get_messages(access_token, fid, 25);
      setMessages((data.value || []) as graph_message[]);
    } catch (e: any) {
      set_error("Failed to load messages. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  function toggle_select(id: string) {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  const privacy_text = useMemo(() => {
    return [
      "You are about to store selected email content locally on this device.",
      "This MVP stores the selected email text locally for search and question answering.",
      "No email content is uploaded to a remote server by this tool.",
      "You can clear the local index at any time."
    ].join(" ");
  }, []);

  async function run_ingest_confirmed() {
    if (!token_ok || !access_token) {
      alert("Please sign in first.");
      return;
    }
    if (!folder_id) {
      alert("Select a folder first.");
      return;
    }
    if (!selected_ids.size) {
      alert("Select at least one email.");
      return;
    }

    setBusy("ingesting (storing + indexing locally)...");
    set_error("");

    try {
      await ingest_messages(access_token, folder_id, Array.from(selected_ids));

      // Refresh local status and route user to chat if index now exists.
      await refresh_index_status();

      setAnswer("Ingestion complete. You can now ask questions.");
      setSources([]);
      setShowConsent(false);
      setConsentChecked(false);

      // If index exists now, show chat by default.
      if ((index_status?.indexed_count || 0) > 0) {
        setScreen("CHAT");
      } else {
        // Fallback: show chat anyway; refresh_index_status should normally handle.
        setScreen("CHAT");
      }
    } catch (e: any) {
      set_error("Ingestion failed. If the local index is corrupted, clear it and try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function run_clear_index() {
    if (!token_ok || !access_token) {
      alert("Please sign in first.");
      return;
    }

    setBusy("clearing local index...");
    set_error("");

    try {
      await clear_index(access_token);
      await refresh_index_status();

      setAnswer("Local index cleared.");
      setSources([]);
      setMessages([]);
      setSelectedIds(new Set());
      setFolderId("");

      // Force Index screen when empty.
      setScreen("INDEX");
      setIndexPanelOpen(true);
    } catch (e: any) {
      set_error("Failed to clear local index. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- query ----------
  async function run_query() {
    if (!token_ok || !access_token) {
      alert("Please sign in first.");
      return;
    }

    const q = question.trim();
    if (!q) return;

    setBusy("querying local index...");
    set_error("");

    try {
      const res = await ask_question(access_token, q, 4);
      setAnswer(res.answer || "");
      setSources((res.sources || []) as source[]);
    } catch (e: any) {
      set_error("Query failed. If the local index is unavailable, clear it and re-ingest.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- render blocks ----------
  function render_error() {
    if (!error) return null;
    return (
      <div style={{ marginTop: 10 }}>
        <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5" }}>
          <strong>Error</strong>
          <div style={{ marginTop: 6, fontSize: 12 }}>{error}</div>
          {error_details ? (
            <details style={{ marginTop: 8 }}>
              <summary style={{ cursor: "pointer", fontSize: 12 }}>Show details</summary>
              <pre style={{ whiteSpace: "pre-wrap", marginTop: 8, fontSize: 12 }}>{error_details}</pre>
            </details>
          ) : null}
        </div>
      </div>
    );
  }

  function render_header() {
    return (
      <div style={{ marginBottom: 10 }}>
        <h2 style={{ marginTop: 0, marginBottom: 6 }}>Outlook Privacy Assistant</h2>
        <div style={{ fontSize: 12, opacity: 0.85 }}>
          <strong>Status:</strong> {status} {busy ? `— ${busy}` : ""}
          {token_ok && user_label ? (
            <span> — <strong>Signed in:</strong> {user_label}</span>
          ) : null}
        </div>

        {token_ok ? (
          <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
            <button
              onClick={() => setScreen("CHAT")}
              disabled={(index_status?.indexed_count || 0) <= 0}
              title={(index_status?.indexed_count || 0) <= 0 ? "Index emails first" : "Go to chat"}
            >
              Chat
            </button>
            <button onClick={() => setScreen("INDEX")}>
              Index management
            </button>
            <button onClick={() => sign_out_clicked()}>
              Sign out
            </button>
          </div>
        ) : null}

        {render_error()}
        <hr style={{ marginTop: 12 }} />
      </div>
    );
  }

  function render_signin() {
    return (
      <div>
        <div style={{ fontSize: 12, opacity: 0.85 }}>
          Sign in to access Microsoft Graph and fetch emails for local indexing.
        </div>
        <div style={{ marginTop: 10 }}>
          <button onClick={() => sign_in_clicked()}>Sign in</button>
        </div>
      </div>
    );
  }

  function render_index_management(is_collapsible_panel: boolean) {
    const empty = (index_status?.indexed_count || 0) <= 0;

    const body = (
      <div>
        {empty ? (
          <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5", marginBottom: 10 }}>
            <strong>No emails indexed yet</strong>
            <div style={{ fontSize: 12, marginTop: 6 }}>
              Select a folder and choose emails to index locally.
            </div>
          </div>
        ) : (
          <div style={{ border: "1px solid #ddd", padding: 10, background: "#fafafa", marginBottom: 10 }}>
            <div style={{ fontSize: 12 }}>
              <strong>Indexed:</strong> {index_status?.indexed_count ?? 0} emails
              {" "}
              — <strong>Last updated:</strong> {fmt_dt(index_status?.last_updated)}
            </div>
          </div>
        )}

        <h3 style={{ marginTop: 0 }}>Ingest emails</h3>

        <div style={{ marginBottom: 6 }}>
          <label style={{ display: "block", fontSize: 12, opacity: 0.85 }}>Select folder</label>
          <select
            style={{ width: "100%" }}
            value={folder_id}
            onChange={(e) => load_messages(e.target.value)}
            disabled={!token_ok}
          >
            <option value="">Select…</option>
            {folders.map((f) => (
              <option key={f.id} value={f.id}>
                {f.displayName} {typeof f.totalItemCount === "number" ? `(${f.totalItemCount})` : ""}
              </option>
            ))}
          </select>
        </div>

        <div style={{ fontSize: 12, opacity: 0.8 }}>
          Select individual emails to ingest. The selected emails will be stored locally.
        </div>

        <div style={{ maxHeight: 220, overflow: "auto", border: "1px solid #ddd", marginTop: 8, padding: 6 }}>
          {messages.length ? (
            messages.map((m) => {
              const from = m.from?.emailAddress?.address || "";
              const checked = selected_ids.has(m.id);
              return (
                <label key={m.id} style={{ display: "block", padding: "6px 4px", borderBottom: "1px solid #eee" }}>
                  <input type="checkbox" checked={checked} onChange={() => toggle_select(m.id)} />{" "}
                  <strong>{m.subject || "(no subject)"}</strong>
                  <div style={{ fontSize: 12, opacity: 0.85 }}>
                    {from} — {m.receivedDateTime || ""}
                  </div>
                  <div style={{ fontSize: 12 }}>{truncate(m.bodyPreview || "", 180)}</div>
                </label>
              );
            })
          ) : (
            <div style={{ fontSize: 12, opacity: 0.8, padding: 6 }}>
              {folder_id ? "No messages loaded yet (or folder is empty)." : "Select a folder to view messages."}
            </div>
          )}
        </div>

        <button
          style={{ marginTop: 10 }}
          disabled={!token_ok || !folder_id || selected_count === 0}
          onClick={() => setShowConsent(true)}
        >
          ingest selected ({selected_count})
        </button>

        {show_consent && (
          <div style={{ border: "1px solid #c00", padding: 10, marginTop: 10, background: "#fff5f5" }}>
            <strong>Consent required</strong>
            <p style={{ marginTop: 8, fontSize: 12 }}>
              {privacy_text}
            </p>

            <label style={{ display: "block", marginBottom: 8, fontSize: 12 }}>
              <input
                type="checkbox"
                checked={consent_checked}
                onChange={(e) => setConsentChecked(e.target.checked)}
              />{" "}
              I understand and consent to local storage.
            </label>

            <button disabled={!consent_checked} onClick={() => run_ingest_confirmed()}>
              confirm ingestion
            </button>{" "}
            <button onClick={() => setShowConsent(false)}>
              cancel
            </button>
          </div>
        )}

        <hr />

        <h3>Clear local index</h3>
        <div style={{ fontSize: 12, opacity: 0.8 }}>
          Clears locally stored indexed emails and the local search index.
        </div>
        <button style={{ marginTop: 8 }} disabled={!token_ok} onClick={() => run_clear_index()}>
          clear local index
        </button>
      </div>
    );

    if (!is_collapsible_panel) return body;

    return (
      <div style={{ border: "1px solid #ddd", borderRadius: 6, padding: 10, background: "#fafafa" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <strong>Index management</strong>
          <button onClick={() => setIndexPanelOpen((v) => !v)}>
            {index_panel_open ? "Hide" : "Show"}
          </button>
        </div>

        {index_panel_open ? (
          <div style={{ marginTop: 10 }}>
            {body}
          </div>
        ) : (
          <div style={{ fontSize: 12, opacity: 0.8, marginTop: 8 }}>
            Indexed: {index_status?.indexed_count ?? 0} — Last updated: {fmt_dt(index_status?.last_updated)}
          </div>
        )}
      </div>
    );
  }

  function render_chat() {
    const no_index = (index_status?.indexed_count || 0) <= 0;

    if (no_index) {
      // Should normally not happen because we disable Chat button, but handle defensively.
      return (
        <div>
          <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5" }}>
            <strong>No emails indexed</strong>
            <div style={{ fontSize: 12, marginTop: 6 }}>
              Index emails first using Index management.
            </div>
          </div>
          <div style={{ marginTop: 10 }}>
            {render_index_management(false)}
          </div>
        </div>
      );
    }

    return (
      <div>
        {render_index_management(true)}

        <hr />

        <h3>Chat</h3>
        <div style={{ fontSize: 12, opacity: 0.8 }}>
          Ask questions about your indexed emails. (Chat history will be added in Phase 4.)
        </div>

        <textarea
          value={question}
          onChange={(e) => setQuestion(e.target.value)}
          rows={3}
          style={{ width: "100%", marginTop: 8 }}
          placeholder="Ask a question about your indexed emails..."
        />

        <div style={{ marginTop: 8 }}>
          <button onClick={() => run_query()} disabled={!token_ok}>
            ask
          </button>
        </div>

        <h3 style={{ marginTop: 14 }}>Answer</h3>
        <div style={{ whiteSpace: "pre-wrap", border: "1px solid #ddd", padding: 8, minHeight: 70 }}>
          {answer || "No answer yet."}
        </div>

        <h3 style={{ marginTop: 14 }}>Sources</h3>
        {sources.length ? (
          sources.map((s, idx) => (
            <div key={s.message_id} style={{ border: "1px solid #eee", padding: 8, marginBottom: 8 }}>
              <div>
                <strong>[{idx + 1}] {s.subject || "(no subject)"}</strong>
              </div>
              <div style={{ fontSize: 12, opacity: 0.8 }}>
                {s.sender} — {s.received_dt} {typeof s.score === "number" ? `— score: ${s.score.toFixed(4)}` : ""}
              </div>
              <div style={{ fontSize: 12, marginTop: 6 }}>{s.snippet}</div>
              <button style={{ marginTop: 6 }} onClick={() => window.open(s.weblink, "_blank", "noopener,noreferrer")}>
                open email
              </button>
            </div>
          ))
        ) : (
          <div style={{ fontSize: 12, opacity: 0.8 }}>No sources yet.</div>
        )}
      </div>
    );
  }

  // ---------- main render ----------
  return (
    <div style={{ fontFamily: "Segoe UI, Arial", padding: 12, lineHeight: 1.35 }}>
      {render_header()}

      {screen === "SIGNIN" && render_signin()}
      {screen === "INDEX" && render_index_management(false)}
      {screen === "CHAT" && render_chat()}
    </div>
  );
}
