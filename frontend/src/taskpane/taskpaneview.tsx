import { useEffect, useMemo, useState } from "react";
import { sign_in_interactive, try_get_access_token_silent, get_signed_in_user, sign_out } from "../auth/token";
import { ask_question, clear_index, get_folders, get_messages, ingest_messages } from "../api/backend";

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

export default function TaskPaneView() {
  const [status, setStatus] = useState("checking session...");
  const [error, setError] = useState<string>("");

  const [token_ok, setTokenOk] = useState(false);
  const [access_token, setAccessToken] = useState<string>("");

  const [user_label, setUserLabel] = useState<string>("");

  const [folders, setFolders] = useState<folder[]>([]);
  const [folder_id, setFolderId] = useState("");
  const [messages, setMessages] = useState<graph_message[]>([]);
  const [selected_ids, setSelectedIds] = useState<Set<string>>(new Set());

  const selected_count = selected_ids.size;

  const [show_consent, setShowConsent] = useState(false);
  const [consent_checked, setConsentChecked] = useState(false);

  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");
  const [sources, setSources] = useState<source[]>([]);

  const [busy, setBusy] = useState<string>("");

  async function load_folders_with_token(token: string) {
    setBusy("loading folders...");
    const data = await get_folders(token);
    setFolders((data.folders || []) as folder[]);
    setBusy("");
  }

  async function initialize_session_silent() {
    setError("");
    setStatus("checking session...");

    try {
      const token = await try_get_access_token_silent();
      if (!token) {
        setTokenOk(false);
        setAccessToken("");
        setUserLabel("");
        setStatus("not signed in");
        return;
      }

      setTokenOk(true);
      setAccessToken(token);

      const who = await get_signed_in_user();
      setUserLabel(who?.username || who?.name || "");

      setStatus("signed in");
      await load_folders_with_token(token);
      setStatus("ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setStatus("error");
      setError(String(e?.message || e));
    }
  }

  async function sign_in_clicked() {
    setError("");
    setBusy("signing in...");

    try {
      const token = await sign_in_interactive();
      setTokenOk(true);
      setAccessToken(token);

      const who = await get_signed_in_user();
      setUserLabel(who?.username || who?.name || "");

      setStatus("signed in");
      await load_folders_with_token(token);
      setStatus("ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setStatus("error");
      setError(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function sign_out_clicked() {
    setError("");
    setBusy("signing out...");

    try {
      await sign_out();

      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setFolders([]);
      setFolderId("");
      setMessages([]);
      setSelectedIds(new Set());
      setShowConsent(false);
      setConsentChecked(false);
      setQuestion("");
      setAnswer("");
      setSources([]);
      setStatus("not signed in");
    } catch (e: any) {
      setStatus("error");
      setError(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  useEffect(() => {
    initialize_session_silent().catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function load_messages(fid: string) {
    setFolderId(fid);
    setMessages([]);
    setSelectedIds(new Set());
    setShowConsent(false); // FIX: correct setter name
    setConsentChecked(false);

    if (!fid) return;
    if (!token_ok || !access_token) return;

    setBusy("loading messages...");
    try {
      const data = await get_messages(access_token, fid, 25);
      setMessages((data.value || []) as graph_message[]);
    } catch (e: any) {
      setError(String(e?.message || e));
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
    setError("");
    try {
      await ingest_messages(access_token, folder_id, Array.from(selected_ids));
      setAnswer("Ingestion complete. You can now ask questions.");
      setSources([]);
      setShowConsent(false);
      setConsentChecked(false);
    } catch (e: any) {
      setError(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function run_query() {
    if (!token_ok || !access_token) {
      alert("Please sign in first.");
      return;
    }
    const q = question.trim();
    if (!q) return;

    setBusy("querying local index...");
    setError("");
    try {
      const res = await ask_question(access_token, q, 4);
      setAnswer(res.answer || "");
      setSources((res.sources || []) as source[]);
    } catch (e: any) {
      setError(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function run_clear() {
    if (!token_ok || !access_token) {
      alert("Please sign in first.");
      return;
    }
    setBusy("clearing local index...");
    setError("");
    try {
      await clear_index(access_token);
      setAnswer("Local index cleared.");
      setSources([]);
    } catch (e: any) {
      setError(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  return (
    <div style={{ fontFamily: "Segoe UI, Arial", padding: 12, lineHeight: 1.35 }}>
      <h2 style={{ marginTop: 0 }}>Outlook Privacy Assistant</h2>

      <div style={{ fontSize: 12, opacity: 0.85, marginBottom: 10 }}>
        <strong>Status:</strong> {status} {busy ? `— ${busy}` : ""}
        {token_ok && user_label ? (
          <span> — <strong>Signed in:</strong> {user_label}</span>
        ) : null}
      </div>

      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        {!token_ok ? (
          <button onClick={() => sign_in_clicked()}>Sign in</button>
        ) : (
          <button onClick={() => sign_out_clicked()}>Sign out</button>
        )}
      </div>

      {error && (
        <pre style={{ whiteSpace: "pre-wrap", border: "1px solid #ccc", padding: 8, marginTop: 10 }}>
          {error}
        </pre>
      )}

      <hr />

      <h3>Ingest emails</h3>

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

      <h3>Ask a question</h3>
      <textarea
        value={question}
        onChange={(e) => setQuestion(e.target.value)}
        rows={3}
        style={{ width: "100%" }}
        placeholder="Ask a question about your indexed emails..."
      />
      <div style={{ marginTop: 8 }}>
        <button onClick={() => run_query()} disabled={!token_ok}>
          ask
        </button>{" "}
        <button onClick={() => run_clear()} disabled={!token_ok}>
          clear local index
        </button>
      </div>

      <h3 style={{ marginTop: 14 }}>Answer</h3>
      <div style={{ whiteSpace: "pre-wrap", border: "1px solid #ddd", padding: 8, minHeight: 70 }}>
        {answer || "No answer yet."}
      </div>

      <h3 style={{ marginTop: 14 }}>Citations</h3>
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
        <div style={{ fontSize: 12, opacity: 0.8 }}>No citations.</div>
      )}
    </div>
  );
}
