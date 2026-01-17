import { useEffect, useMemo, useRef, useState } from "react";
import { sign_in_interactive, try_get_access_token_silent, get_signed_in_user, sign_out } from "../auth/token";
import {
  ask_question,
  clear_index,
  get_folders,
  get_messages,
  ingest_messages,
  get_index_status,
  get_messages_page,
  type GraphMessagesPage,
  type IndexStatus
} from "../api/backend";

type Screen = "SIGNIN" | "CHAT" | "INDEX";
type IngestStep = "SELECT" | "PREVIEW" | "RUNNING" | "CANCELLED" | "COMPLETE";
type IngestMode = "FOLDERS" | "EMAILS";
type Phase = "FETCHING" | "STORING" | "INDEXING" | "DONE";

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

function uniq<T>(arr: T[]) {
  return Array.from(new Set(arr));
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// Office.js may not be typed in your TS config; keep it safe.
function try_convert_to_outlook_desktop_url(owaUrl: string): string | null {
  try {
    const w: any = window as any;
    const OfficeObj = w?.Office;
    const mailbox = OfficeObj?.context?.mailbox;
    const fn = mailbox?.convertToLocalClientUrl;
    if (typeof fn === "function") {
      const localUrl = fn.call(mailbox, owaUrl);
      if (typeof localUrl === "string" && localUrl.length > 0) return localUrl;
    }
  } catch {
    // ignore and fall back
  }
  return null;
}

export default function TaskPaneView() {
  // ---------- global UI ----------
  const [screen, setScreen] = useState<Screen>("SIGNIN");

  const [status, setStatus] = useState("starting...");
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

  // ---------- index status ----------
  const [index_status, setIndexStatus] = useState<IndexStatus | null>(null);
  const [index_panel_open, setIndexPanelOpen] = useState<boolean>(false);

  async function refresh_index_status(nextPreferredScreen: Screen | null = null) {
    const st = (await get_index_status()) as IndexStatus;
    setIndexStatus(st);

    if ((st?.indexed_count || 0) <= 0) {
      setScreen("INDEX");
      setIndexPanelOpen(true);
      return;
    }

    if (nextPreferredScreen) setScreen(nextPreferredScreen);
    else if (screen === "SIGNIN") setScreen("CHAT");
  }

  // ---------- graph data ----------
  const [folders, setFolders] = useState<folder[]>([]);
  const [folder_filter, setFolderFilter] = useState("");

  // Messages view for email selection mode
  const [email_folder_id, setEmailFolderId] = useState("");
  const [messages, setMessages] = useState<graph_message[]>([]);
  const [messages_filter, setMessagesFilter] = useState("");
  const [messages_next_link, setMessagesNextLink] = useState<string | null>(null);

  // ---------- ingestion wizard state ----------
  const [ingest_step, setIngestStep] = useState<IngestStep>("SELECT");
  const [ingest_mode, setIngestMode] = useState<IngestMode>("FOLDERS");

  // folder selection
  const [selected_folder_ids, setSelectedFolderIds] = useState<Set<string>>(new Set());
  const [folder_limit, setFolderLimit] = useState<number>(100); // per folder (latest N)

  // email selection
  const [selected_email_ids, setSelectedEmailIds] = useState<Set<string>>(new Set());

  // consent + warnings
  const [consent_checked, setConsentChecked] = useState(false);
  const [large_ack_checked, setLargeAckChecked] = useState(false);

  // run/progress state
  const [run_phase, setRunPhase] = useState<Phase>("FETCHING");
  const [run_total, setRunTotal] = useState<number | null>(null);
  const [fetch_done, setFetchDone] = useState<number>(0);
  const [ingest_done, setIngestDone] = useState<number>(0);

  const [cancel_confirm_open, setCancelConfirmOpen] = useState(false);
  const [cancel_summary, setCancelSummary] = useState<string>("");

  const [complete_summary, setCompleteSummary] = useState<string>("");

  // abort only affects the FETCHING loop
  const abort_ref = useRef<AbortController | null>(null);
  const cancel_requested_ref = useRef<boolean>(false);

  // ---------- chat ----------
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");
  const [sources, setSources] = useState<source[]>([]);

  // ---------- clear index confirmation modal ----------
  const [clear_modal_open, setClearModalOpen] = useState(false);
  const [clear_confirm_checked, setClearConfirmChecked] = useState(false);

  // ---------- consent text ----------
  const privacy_text = useMemo(() => {
    return [
      "You are about to store selected email content locally on this device.",
      "This tool stores selected email text locally for search and question answering.",
      "No email content is uploaded to a remote server by this tool.",
      "You can clear the local index at any time."
    ].join(" ");
  }, []);

  // ---------- startup ----------
  async function load_folders_with_token(token: string) {
    setBusy("loading folders...");
    const data = await get_folders(token);
    setFolders((data.folders || []) as folder[]);
    setBusy("");
  }

  async function initialize() {
    set_error("");
    setStatus("checking session...");
    setBusy("");

    try {
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
      reset_ingest_state();

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

  // ---------- ingestion helpers ----------
  function reset_ingest_state() {
    setIngestStep("SELECT");
    setIngestMode("FOLDERS");
    setSelectedFolderIds(new Set());
    setFolderLimit(100);

    setEmailFolderId("");
    setMessages([]);
    setMessagesFilter("");
    setMessagesNextLink(null);
    setSelectedEmailIds(new Set());

    setConsentChecked(false);
    setLargeAckChecked(false);

    setRunPhase("FETCHING");
    setRunTotal(null);
    setFetchDone(0);
    setIngestDone(0);

    setCancelConfirmOpen(false);
    setCancelSummary("");
    setCompleteSummary("");

    abort_ref.current = null;
    cancel_requested_ref.current = false;
  }

  const index_empty = (index_status?.indexed_count || 0) <= 0;

  const filtered_folders = useMemo(() => {
    const q = folder_filter.trim().toLowerCase();
    if (!q) return folders;
    return folders.filter((f) => (f.displayName || "").toLowerCase().includes(q));
  }, [folders, folder_filter]);

  const filtered_messages = useMemo(() => {
    const q = messages_filter.trim().toLowerCase();
    if (!q) return messages;
    return messages.filter((m) => {
      const subj = (m.subject || "").toLowerCase();
      const from = (m.from?.emailAddress?.address || "").toLowerCase();
      return subj.includes(q) || from.includes(q);
    });
  }, [messages, messages_filter]);

  function toggle_folder(id: string) {
    setSelectedFolderIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function toggle_email(id: string) {
    setSelectedEmailIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function select_all_filtered_emails() {
    setSelectedEmailIds((prev) => {
      const next = new Set(prev);
      for (const m of filtered_messages) next.add(m.id);
      return next;
    });
  }

  function clear_email_selection() {
    setSelectedEmailIds(new Set());
  }

  const selection_summary = useMemo(() => {
    const folderCount = selected_folder_ids.size;
    const emailCount = selected_email_ids.size;

    const approxFolderTotal = (() => {
      if (!folderCount) return 0;
      const selected = folders.filter((f) => selected_folder_ids.has(f.id));
      return selected.reduce((sum, f) => sum + Math.min(folder_limit, f.totalItemCount || folder_limit), 0);
    })();

    const effectiveTotal = ingest_mode === "FOLDERS" ? approxFolderTotal : emailCount;

    return {
      folderCount,
      emailCount,
      approxFolderTotal,
      effectiveTotal,
      large: effectiveTotal >= 200
    };
  }, [selected_folder_ids, selected_email_ids, folders, ingest_mode, folder_limit]);

  // ---------- messages loading for EMAIL mode ----------
  async function load_email_folder(fid: string) {
    setEmailFolderId(fid);
    setMessages([]);
    setMessagesNextLink(null);
    setSelectedEmailIds(new Set());
    setMessagesFilter("");

    if (!fid) return;
    if (!token_ok || !access_token) return;

    setBusy("loading messages...");
    set_error("");

    try {
      const data = await get_messages(access_token, fid, 25);
      setMessages((data.value || []) as graph_message[]);
      setMessagesNextLink((data as any)["@odata.nextLink"] || null);
    } catch (e: any) {
      set_error("Failed to load messages. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function load_more_messages() {
    if (!token_ok || !access_token) return;
    if (!messages_next_link) return;

    setBusy("loading more messages...");
    set_error("");

    try {
      const data = await get_messages_page(access_token, { next_link: messages_next_link, top: 25 });
      const page = (data.value || []) as graph_message[];
      setMessages((prev) => [...prev, ...page]);
      setMessagesNextLink((data as any)["@odata.nextLink"] || null);
    } catch (e: any) {
      set_error("Failed to load more messages. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- wizard transitions ----------
  function can_continue_from_select(): boolean {
    if (ingest_mode === "FOLDERS") return selected_folder_ids.size > 0;
    return selected_email_ids.size > 0;
  }

  function go_preview() {
    set_error("");
    setConsentChecked(false);
    setLargeAckChecked(false);
    setCancelSummary("");
    setIngestStep("PREVIEW");
  }

  function back_to_select() {
    set_error("");
    setCancelSummary("");
    setIngestStep("SELECT");
  }

  // ---------- cancel flow ----------
  function cancel_clicked() {
    setCancelConfirmOpen(true);
  }

  function cancel_continue() {
    setCancelConfirmOpen(false);
  }

  function cancel_confirm_now() {
    setCancelConfirmOpen(false);
    cancel_requested_ref.current = true;

    if (abort_ref.current) abort_ref.current.abort();

    setCancelSummary("Indexing cancelled. Some items may already be indexed.");
    setIngestStep("CANCELLED");
  }

  // ---------- ingestion run ----------
  async function run_ingestion() {
    if (!token_ok || !access_token) {
      alert("Please sign in first.");
      return;
    }

    if (!consent_checked) return;
    if (selection_summary.large && !large_ack_checked) return;

    set_error("");
    setCancelSummary("");
    cancel_requested_ref.current = false;

    setIngestStep("RUNNING");
    setRunPhase("FETCHING");
    setRunTotal(selection_summary.effectiveTotal || null);
    setFetchDone(0);
    setIngestDone(0);

    const MIN_PHASE_MS = 450;
    let phaseStart = Date.now();

    const ac = new AbortController();
    abort_ref.current = ac;

    try {
      // 1) Build IDs (FETCHING)
      let message_ids: string[] = [];
      const PAGE_SIZE = 25;

      if (ingest_mode === "EMAILS") {
        message_ids = Array.from(selected_email_ids);
        setFetchDone(message_ids.length);
      } else {
        const selected = folders.filter((f) => selected_folder_ids.has(f.id));
        const all_ids: string[] = [];
        let done = 0;

        for (const f of selected) {
          const limit = Math.max(1, folder_limit);
          let next_link: string | null = null;
          let collected = 0;

          while (collected < limit) {
            if (ac.signal.aborted || cancel_requested_ref.current) throw new Error("CANCELLED_BY_USER");

            let pageData: GraphMessagesPage;

            if (next_link) pageData = await get_messages_page(access_token, { next_link, top: PAGE_SIZE });
            else pageData = await get_messages_page(access_token, { folder_id: f.id, top: PAGE_SIZE });

            const page = (pageData.value || []) as graph_message[];
            next_link = (pageData as any)["@odata.nextLink"] || null;

            if (!page.length) break;

            for (const m of page) {
              if (collected >= limit) break;
              all_ids.push(m.id);
              collected += 1;
              done += 1;
              setFetchDone(done);
            }

            if (!next_link) break;
          }
        }

        message_ids = uniq(all_ids);
      }

      if (!message_ids.length) throw new Error("No messages selected.");
      if (cancel_requested_ref.current) throw new Error("CANCELLED_BY_USER");

      // Make FETCHING visible
      {
        const elapsed = Date.now() - phaseStart;
        if (elapsed < MIN_PHASE_MS) await sleep(MIN_PHASE_MS - elapsed);
      }

      // 2) STORING: batch ingest calls so progress is visible
      setRunPhase("STORING");
      phaseStart = Date.now();

      const folder_id_to_send = ingest_mode === "EMAILS" ? (email_folder_id || "selected") : "multi";

      const BATCH_SIZE = 10;
      const batches = chunk(message_ids, BATCH_SIZE);

      let ingested = 0;
      for (const b of batches) {
        if (cancel_requested_ref.current) throw new Error("CANCELLED_BY_USER");
        await ingest_messages(access_token, folder_id_to_send, b);
        ingested += b.length;
        setIngestDone(ingested);
      }

      // Make STORING visible
      {
        const elapsed = Date.now() - phaseStart;
        if (elapsed < MIN_PHASE_MS) await sleep(MIN_PHASE_MS - elapsed);
      }

      if (cancel_requested_ref.current) throw new Error("CANCELLED_BY_USER");

      // 3) INDEXING: backend rebuild happens during ingest; keep a visible phase
      setRunPhase("INDEXING");
      phaseStart = Date.now();

      {
        const elapsed = Date.now() - phaseStart;
        if (elapsed < MIN_PHASE_MS) await sleep(MIN_PHASE_MS - elapsed);
      }

      setRunPhase("DONE");

      await refresh_index_status();

      const st = (await get_index_status()) as IndexStatus;
      setIndexStatus(st);

      const summary = ingest_mode === "EMAILS"
        ? `Indexed ${message_ids.length} selected emails.`
        : `Indexed ${message_ids.length} emails from selected folder(s).`;

      setCompleteSummary(`${summary} Index updated at ${fmt_dt(st.last_updated)}.`);
      setIngestStep("COMPLETE");
    } catch (e: any) {  // <-- FIXED HERE
      const msg = String(e?.message || e);

      if (msg === "CANCELLED_BY_USER") {
        await refresh_index_status();
        setCancelSummary("Indexing cancelled. Some items may already be indexed.");
        setIngestStep("CANCELLED");
        return;
      }

      setIngestStep("SELECT");
      set_error(
        "Indexing failed. If the local index is unavailable, clear local index and try again.",
        msg
      );
    } finally {
      abort_ref.current = null;
      setBusy("");
    }
  }

  // ---------- clear index flow ----------
  async function clear_index_confirmed() {
    if (!token_ok || !access_token) return;

    set_error("");
    setBusy("clearing local index...");

    try {
      await clear_index(access_token);
      await refresh_index_status("INDEX");
      reset_ingest_state();
      setClearModalOpen(false);
      setClearConfirmChecked(false);
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

  // ---------- UI render helpers ----------
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
            <button onClick={() => setScreen("INDEX")}>Index management</button>
            <button onClick={() => sign_out_clicked()}>Sign out</button>
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

  function progress_done_now(): number {
    return run_phase === "FETCHING" ? fetch_done : ingest_done;
  }

  function render_phase_bar() {
    const phases: { key: Phase; label: string }[] = [
      { key: "FETCHING", label: "Fetching emails" },
      { key: "STORING", label: "Storing locally" },
      { key: "INDEXING", label: "Indexing for search" },
      { key: "DONE", label: "Done" },
    ];

    const currentIdx = phases.findIndex((p) => p.key === run_phase);
    const doneNow = progress_done_now();

    return (
      <div style={{ border: "1px solid #ddd", padding: 10, background: "#fafafa" }}>
        <div style={{ fontSize: 12, opacity: 0.85, marginBottom: 6 }}>
          <strong>Progress</strong>
        </div>

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", fontSize: 12 }}>
          {phases.map((p, idx) => {
            const active = idx === currentIdx;
            const done = idx < currentIdx;
            return (
              <span
                key={p.key}
                style={{
                  padding: "2px 6px",
                  border: "1px solid #ddd",
                  background: active ? "#fff" : done ? "#f0f0f0" : "#fafafa"
                }}
              >
                {done ? "✓ " : active ? "→ " : ""}{p.label}
              </span>
            );
          })}
        </div>

        <div style={{ marginTop: 10 }}>
          <div style={{ height: 8, background: "#eee", borderRadius: 4, overflow: "hidden" }}>
            <div
              style={{
                height: 8,
                width: run_total
                  ? `${Math.min(100, Math.round((doneNow / Math.max(1, run_total)) * 100))}%`
                  : "25%",
                background: "#bbb"
              }}
            />
          </div>
          <div style={{ fontSize: 12, opacity: 0.8, marginTop: 6 }}>
            {run_total ? (
              <span>Processed {doneNow} / {run_total} emails</span>
            ) : (
              <span>Processed {doneNow} emails</span>
            )}
          </div>
        </div>

        <div style={{ marginTop: 10 }}>
          <button onClick={cancel_clicked}>Cancel</button>
        </div>
      </div>
    );
  }

  function render_cancel_confirm_modal() {
    if (!cancel_confirm_open) return null;
    return (
      <div style={{ border: "2px solid #c00", padding: 12, background: "#fff5f5", marginTop: 10 }}>
        <strong>Cancel indexing now?</strong>
        <div style={{ fontSize: 12, marginTop: 6 }}>
          Some emails may already be indexed.
        </div>
        <div style={{ marginTop: 10 }}>
          <button onClick={cancel_continue}>Continue indexing</button>{" "}
          <button onClick={cancel_confirm_now}>Cancel indexing now</button>
        </div>
      </div>
    );
  }

  function render_clear_modal() {
    if (!clear_modal_open) return null;
    return (
      <div style={{ border: "2px solid #c00", padding: 12, background: "#fff5f5", marginTop: 10 }}>
        <strong>Clear local index?</strong>
        <ul style={{ marginTop: 8, fontSize: 12 }}>
          <li>Deletes locally stored indexed email text</li>
          <li>Deletes local search index</li>
          <li>Cannot be undone</li>
        </ul>

        <label style={{ display: "block", fontSize: 12, marginTop: 8 }}>
          <input
            type="checkbox"
            checked={clear_confirm_checked}
            onChange={(e) => setClearConfirmChecked(e.target.checked)}
          />{" "}
          I understand this cannot be undone.
        </label>

        <div style={{ marginTop: 10 }}>
          <button disabled={!clear_confirm_checked} onClick={clear_index_confirmed}>
            Clear local index
          </button>{" "}
          <button onClick={() => { setClearModalOpen(false); setClearConfirmChecked(false); }}>
            Cancel
          </button>
        </div>
      </div>
    );
  }

  // --- Index management render sections ---
  function render_index_select_scope() {
    return (
      <div>
        {index_empty ? (
          <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5", marginBottom: 10 }}>
            <strong>No emails indexed yet</strong>
            <div style={{ fontSize: 12, marginTop: 6 }}>
              Select folders or individual emails to index locally.
            </div>
          </div>
        ) : null}

        {!index_empty ? (
          <div style={{ border: "1px solid #ddd", padding: 10, background: "#fafafa", marginBottom: 10 }}>
            <div style={{ fontSize: 12 }}>
              <strong>Indexed:</strong> {index_status?.indexed_count ?? 0} emails — <strong>Last updated:</strong> {fmt_dt(index_status?.last_updated)}
            </div>
          </div>
        ) : null}

        <h3 style={{ marginTop: 0 }}>Index management</h3>

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
          <button onClick={() => setIngestMode("FOLDERS")} disabled={ingest_mode === "FOLDERS"}>
            Select folders
          </button>
          <button onClick={() => setIngestMode("EMAILS")} disabled={ingest_mode === "EMAILS"}>
            Select emails
          </button>
        </div>

        {ingest_mode === "FOLDERS" ? (
          <div>
            <div style={{ fontSize: 12, opacity: 0.85, marginBottom: 6 }}>
              Choose one or more folders. The system will index the latest N emails per folder.
            </div>

            <label style={{ display: "block", fontSize: 12, opacity: 0.85 }}>Folder filter</label>
            <input
              value={folder_filter}
              onChange={(e) => setFolderFilter(e.target.value)}
              placeholder="Type to filter folders..."
              style={{ width: "100%", marginBottom: 8 }}
            />

            <label style={{ display: "block", fontSize: 12, opacity: 0.85 }}>Per-folder limit (latest emails)</label>
            <input
              type="number"
              value={folder_limit}
              min={1}
              max={2000}
              onChange={(e) => setFolderLimit(Math.max(1, Math.min(2000, Number(e.target.value || 100))))}
              style={{ width: "100%", marginBottom: 8 }}
            />

            <div style={{ maxHeight: 220, overflow: "auto", border: "1px solid #ddd", padding: 6 }}>
              {filtered_folders.map((f) => {
                const checked = selected_folder_ids.has(f.id);
                return (
                  <label key={f.id} style={{ display: "block", padding: "6px 4px", borderBottom: "1px solid #eee" }}>
                    <input type="checkbox" checked={checked} onChange={() => toggle_folder(f.id)} />{" "}
                    <strong>{f.displayName}</strong>{" "}
                    <span style={{ fontSize: 12, opacity: 0.8 }}>
                      {typeof f.totalItemCount === "number" ? `(${f.totalItemCount})` : ""}
                    </span>
                  </label>
                );
              })}
              {!filtered_folders.length ? (
                <div style={{ fontSize: 12, opacity: 0.8, padding: 6 }}>No folders match your filter.</div>
              ) : null}
            </div>
          </div>
        ) : (
          <div>
            <div style={{ fontSize: 12, opacity: 0.85, marginBottom: 6 }}>
              Select a folder, then choose individual emails.
            </div>

            <label style={{ display: "block", fontSize: 12, opacity: 0.85 }}>Folder</label>
            <select
              style={{ width: "100%", marginBottom: 8 }}
              value={email_folder_id}
              onChange={(e) => load_email_folder(e.target.value)}
              disabled={!token_ok}
            >
              <option value="">Select…</option>
              {folders.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.displayName} {typeof f.totalItemCount === "number" ? `(${f.totalItemCount})` : ""}
                </option>
              ))}
            </select>

            <label style={{ display: "block", fontSize: 12, opacity: 0.85 }}>Email filter</label>
            <input
              value={messages_filter}
              onChange={(e) => setMessagesFilter(e.target.value)}
              placeholder="Filter by subject or sender..."
              style={{ width: "100%", marginBottom: 8 }}
            />

            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
              <button disabled={!filtered_messages.length} onClick={select_all_filtered_emails}>
                Select all (filtered)
              </button>
              <button disabled={!selected_email_ids.size} onClick={clear_email_selection}>
                Clear selection
              </button>
              <button disabled={!messages_next_link} onClick={load_more_messages}>
                Load more
              </button>
            </div>

            <div style={{ maxHeight: 220, overflow: "auto", border: "1px solid #ddd", padding: 6 }}>
              {filtered_messages.length ? (
                filtered_messages.map((m) => {
                  const from = m.from?.emailAddress?.address || "";
                  const checked = selected_email_ids.has(m.id);
                  return (
                    <label key={m.id} style={{ display: "block", padding: "6px 4px", borderBottom: "1px solid #eee" }}>
                      <input type="checkbox" checked={checked} onChange={() => toggle_email(m.id)} />{" "}
                      <strong>{m.subject || "(no subject)"}</strong>
                      <div style={{ fontSize: 12, opacity: 0.85 }}>
                        {from} — {m.receivedDateTime || ""}
                      </div>
                      <div style={{ fontSize: 12 }}>{truncate(m.bodyPreview || "", 160)}</div>
                    </label>
                  );
                })
              ) : (
                <div style={{ fontSize: 12, opacity: 0.8, padding: 6 }}>
                  {email_folder_id ? "No messages loaded yet (or none match filter)." : "Select a folder to view emails."}
                </div>
              )}
            </div>
          </div>
        )}

        <div style={{ border: "1px solid #eee", padding: 10, marginTop: 10 }}>
          <div style={{ fontSize: 12 }}>
            <strong>Selection summary</strong>
          </div>
          <div style={{ fontSize: 12, opacity: 0.85, marginTop: 6 }}>
            Folders selected: {selected_folder_ids.size} — Emails selected: {selected_email_ids.size}
          </div>
          <div style={{ fontSize: 12, opacity: 0.85, marginTop: 2 }}>
            Total to index (approx): {selection_summary.effectiveTotal}
          </div>
          {selection_summary.large ? (
            <div style={{ fontSize: 12, color: "#900", marginTop: 6 }}>
              Large selections take longer to process and may impact responsiveness.
            </div>
          ) : null}
        </div>

        <div style={{ marginTop: 10 }}>
          <button disabled={!can_continue_from_select()} onClick={go_preview}>
            Continue
          </button>
        </div>

        <hr />

        <h3>Clear local index</h3>
        <div style={{ fontSize: 12, opacity: 0.8 }}>
          Clears locally stored indexed emails and the local search index.
        </div>
        <button style={{ marginTop: 8 }} disabled={!token_ok} onClick={() => setClearModalOpen(true)}>
          Clear local index
        </button>

        {render_clear_modal()}
      </div>
    );
  }

  function render_index_preview() {
    return (
      <div>
        <h3 style={{ marginTop: 0 }}>Consent and confirmation</h3>

        <div style={{ border: "1px solid #ddd", padding: 10, background: "#fafafa" }}>
          <div style={{ fontSize: 12 }}>
            <strong>Summary</strong>
          </div>
          <div style={{ fontSize: 12, opacity: 0.85, marginTop: 6 }}>
            Mode: {ingest_mode === "FOLDERS" ? "Folders" : "Emails"}
          </div>
          <div style={{ fontSize: 12, opacity: 0.85 }}>
            Total to index (approx): {selection_summary.effectiveTotal}
          </div>
          {ingest_mode === "FOLDERS" ? (
            <div style={{ fontSize: 12, opacity: 0.85 }}>
              Per-folder limit: {folder_limit}
            </div>
          ) : null}
        </div>

        {selection_summary.large ? (
          <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5", marginTop: 10 }}>
            <strong>Large selection warning</strong>
            <div style={{ fontSize: 12, marginTop: 6 }}>
              Large selections take longer to process and may impact responsiveness.
            </div>
            <label style={{ display: "block", fontSize: 12, marginTop: 8 }}>
              <input
                type="checkbox"
                checked={large_ack_checked}
                onChange={(e) => setLargeAckChecked(e.target.checked)}
              />{" "}
              I understand this may take several minutes.
            </label>
          </div>
        ) : null}

        <div style={{ border: "1px solid #ddd", padding: 10, marginTop: 10 }}>
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
            I consent to local storage and indexing.
          </label>
        </div>

        <div style={{ marginTop: 10 }}>
          <button onClick={back_to_select}>Back</button>{" "}
          <button
            disabled={!consent_checked || (selection_summary.large && !large_ack_checked)}
            onClick={run_ingestion}
          >
            Start indexing
          </button>
        </div>
      </div>
    );
  }

  function render_index_running() {
    return (
      <div>
        {render_phase_bar()}
        {render_cancel_confirm_modal()}
        <div style={{ marginTop: 10, fontSize: 12, opacity: 0.85 }}>
          Keep Outlook open while indexing runs.
        </div>
      </div>
    );
  }

  function render_index_cancelled() {
    return (
      <div>
        <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5" }}>
          <strong>Indexing cancelled</strong>
          <div style={{ fontSize: 12, marginTop: 6 }}>
            {cancel_summary || "Indexing cancelled. Some items may already be indexed."}
          </div>
        </div>

        <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={() => setClearModalOpen(true)}>Clear local index</button>
          <button
            onClick={() => {
              setCancelSummary("");
              cancel_requested_ref.current = false;
              setIngestStep("SELECT");
            }}
          >
            Return to selection
          </button>
        </div>

        {render_clear_modal()}
      </div>
    );
  }

  function render_index_complete() {
    return (
      <div>
        <div style={{ border: "1px solid #0a0", padding: 10, background: "#f5fff5" }}>
          <strong>Indexing complete</strong>
          <div style={{ fontSize: 12, marginTop: 6 }}>
            {complete_summary || "Indexing completed successfully."}
          </div>
        </div>

        <div style={{ marginTop: 10 }}>
          <button onClick={() => setScreen("CHAT")}>Go to chat</button>{" "}
          <button onClick={() => reset_ingest_state()}>Index more emails</button>
        </div>
      </div>
    );
  }

  function render_index_management(is_collapsible_panel: boolean) {
    const body = (
      <div>
        {ingest_step === "SELECT" && render_index_select_scope()}
        {ingest_step === "PREVIEW" && render_index_preview()}
        {ingest_step === "RUNNING" && render_index_running()}
        {ingest_step === "CANCELLED" && render_index_cancelled()}
        {ingest_step === "COMPLETE" && render_index_complete()}
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

  // ---------- chat screen ----------
  function render_chat() {
    const no_index = (index_status?.indexed_count || 0) <= 0;

    if (no_index) {
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
          Ask questions about your indexed emails.
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
              <details style={{ marginTop: 6 }}>
                <summary style={{ cursor: "pointer", fontSize: 12 }}>Show excerpt</summary>
                <div style={{ fontSize: 12, marginTop: 6 }}>{s.snippet}</div>
              </details>

              <button
                style={{ marginTop: 6 }}
                onClick={() => {
                  const local = try_convert_to_outlook_desktop_url(s.weblink);
                  window.open(local || s.weblink, "_blank", "noopener,noreferrer");
                }}
              >
                Open email
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
