import { useEffect, useMemo, useRef, useState } from "react";
import "./taskpaneview.css";
import { sign_in_interactive, try_get_access_token_silent, get_signed_in_user, sign_out } from "../auth/token";
import {
  ask_question,
  clear_index,
  clear_ingestion,
  get_folders,
  get_message_link,
  get_messages,
  ingest_messages,
  get_index_status,
  get_messages_page,
  list_ingestions,
  type GraphMessagesPage,
  type IndexStatus,
  type IngestionInfo
} from "../api/backend";

type Screen = "SIGNIN" | "CHAT" | "INDEX";
type IngestStep = "SELECT" | "PREVIEW" | "RUNNING" | "CANCELLED" | "COMPLETE";
type IngestMode = "FOLDERS" | "EMAILS";
type Phase = "FETCHING" | "STORING" | "INDEXING" | "DONE";

type ChatState = "EMPTY" | "ACTIVE" | "RESTORED";

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

type ChatMsg = {
  id: string;
  role: "user" | "assistant";
  text: string;
  created_at_ms: number;
  sources?: source[];
  sources_open?: boolean; // UI only
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

function make_id(prefix: string): string {
  const c: any = (globalThis as any).crypto;
  if (c?.randomUUID) return `${prefix}_${c.randomUUID()}`;
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function now_iso_local() {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// Office.js may not be typed in TS; keep it safe.
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
    // ignore
  }
  return null;
}

// ---------- Phase 4 chat persistence ----------
const CHAT_TTL_MS = 60 * 60 * 1000;
const LS_CHAT_KEY = "chat_history";
const LS_CHAT_TS_KEY = "chat_history_ts";

function load_chat_from_storage(): { msgs: ChatMsg[]; ts_ms: number } | null {
  try {
    const tsRaw = localStorage.getItem(LS_CHAT_TS_KEY);
    const dataRaw = localStorage.getItem(LS_CHAT_KEY);
    if (!tsRaw || !dataRaw) return null;

    const ts_ms = Number(tsRaw);
    if (!Number.isFinite(ts_ms) || ts_ms <= 0) return null;

    const parsed = JSON.parse(dataRaw);
    if (!Array.isArray(parsed)) return null;

    const msgs: ChatMsg[] = parsed
      .filter((m: any) => m && typeof m === "object")
      .map((m: any) => ({
        id: String(m.id || make_id("msg")),
        role: m.role === "assistant" ? "assistant" : "user",
        text: String(m.text || ""),
        created_at_ms: Number(m.created_at_ms || ts_ms),
        sources: Array.isArray(m.sources) ? m.sources : undefined,
        sources_open: false
      }));

    return { msgs, ts_ms };
  } catch {
    return null;
  }
}

function save_chat_to_storage(msgs: ChatMsg[]) {
  try {
    // Store without UI-only fields
    const trimmed = msgs.map((m) => ({
      id: m.id,
      role: m.role,
      text: m.text,
      created_at_ms: m.created_at_ms,
      sources: m.sources
    }));
    localStorage.setItem(LS_CHAT_KEY, JSON.stringify(trimmed));
    localStorage.setItem(LS_CHAT_TS_KEY, String(Date.now()));
  } catch {
    // ignore
  }
}

function clear_chat_storage() {
  try {
    localStorage.removeItem(LS_CHAT_KEY);
    localStorage.removeItem(LS_CHAT_TS_KEY);
  } catch {
    // ignore
  }
}

export default function TaskPaneView() {
  // ---------- global ----------
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

  // ---------- index ----------
  const [index_status, setIndexStatus] = useState<IndexStatus | null>(null);
  const [index_panel_open, setIndexPanelOpen] = useState<boolean>(false);

  // ---------- ingestion runs ----------
  const [ingestions, setIngestions] = useState<IngestionInfo[]>([]);
  const [selected_clear_mode, setSelectedClearMode] = useState<"ALL" | "ONE">("ALL");
  const [selected_ingestion_id_to_clear, setSelectedIngestionIdToClear] = useState<string>("");

  async function refresh_ingestions() {
    try {
      const res = await list_ingestions(null, 50);
      const arr = (res.ingestions || []) as IngestionInfo[];
      setIngestions(arr);
      if (!selected_ingestion_id_to_clear && arr.length) {
        setSelectedIngestionIdToClear(arr[0].ingestion_id);
      }
    } catch {
      // non-fatal
    }
  }

  // ---------- chat state ----------
  const [chat_state, setChatState] = useState<ChatState>("EMPTY");
  const [chat_msgs, setChatMsgs] = useState<ChatMsg[]>([]);
  const [chat_expired_note, setChatExpiredNote] = useState<boolean>(false);

  // Chat UI
  const [draft, setDraft] = useState<string>("");
  const [sending, setSending] = useState<boolean>(false);

  const transcript_ref = useRef<HTMLDivElement | null>(null);

  function scroll_transcript_to_bottom() {
    try {
      const el = transcript_ref.current;
      if (!el) return;
      el.scrollTop = el.scrollHeight;
    } catch {
      // ignore
    }
  }

  function append_chat(msg: ChatMsg) {
    setChatMsgs((prev) => {
      const next = [...prev, msg];
      save_chat_to_storage(next);
      return next;
    });
    setChatState("ACTIVE");
  }

  function patch_chat_msg(msgId: string, patch: Partial<ChatMsg>) {
    setChatMsgs((prev) => {
      const next = prev.map((m) => (m.id === msgId ? { ...m, ...patch } : m));
      save_chat_to_storage(next);
      return next;
    });
    setChatState("ACTIVE");
  }

  function restore_chat_if_valid(indexExists: boolean) {
    if (!indexExists) {
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(false);
      return;
    }

    const loaded = load_chat_from_storage();
    if (!loaded) {
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(false);
      return;
    }

    const age = Date.now() - loaded.ts_ms;
    if (age <= CHAT_TTL_MS) {
      setChatMsgs(loaded.msgs);
      setChatState(loaded.msgs.length ? "RESTORED" : "EMPTY");
      setChatExpiredNote(false);
    } else {
      clear_chat_storage();
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(true);
    }
  }

  async function refresh_index_status(nextPreferredScreen: Screen | null = null) {
    const st = (await get_index_status()) as IndexStatus;
    setIndexStatus(st);

    const indexExists = (st?.indexed_count || 0) > 0;

    if (!indexExists) {
      setScreen("INDEX");
      setIndexPanelOpen(true);
      restore_chat_if_valid(false);
    } else {
      if (nextPreferredScreen) setScreen(nextPreferredScreen);
      else if (screen === "SIGNIN") setScreen("CHAT");
      restore_chat_if_valid(true);
    }

    await refresh_ingestions();
  }

  // ---------- graph ----------
  const [folders, setFolders] = useState<folder[]>([]);
  const [folder_filter, setFolderFilter] = useState("");

  const [email_folder_id, setEmailFolderId] = useState("");
  const [messages, setMessages] = useState<graph_message[]>([]);
  const [messages_filter, setMessagesFilter] = useState("");
  const [messages_next_link, setMessagesNextLink] = useState<string | null>(null);

  // ---------- ingestion wizard ----------
  const [ingest_step, setIngestStep] = useState<IngestStep>("SELECT");
  const [ingest_mode, setIngestMode] = useState<IngestMode>("FOLDERS");

  const [selected_folder_ids, setSelectedFolderIds] = useState<Set<string>>(new Set());

  // Per-folder limit: use a string input state so it doesn’t “snap back”
  const [folder_limit_input, setFolderLimitInput] = useState<string>("100");
  const folder_limit = useMemo(() => {
    const n = Number(folder_limit_input);
    if (!Number.isFinite(n) || n <= 0) return 100;
    return Math.max(1, Math.min(2000, Math.floor(n)));
  }, [folder_limit_input]);

  const [selected_email_ids, setSelectedEmailIds] = useState<Set<string>>(new Set());

  const [consent_checked, setConsentChecked] = useState(false);
  const [large_ack_checked, setLargeAckChecked] = useState(false);

  const [run_phase, setRunPhase] = useState<Phase>("FETCHING");
  const [run_total, setRunTotal] = useState<number | null>(null);
  const [fetch_done, setFetchDone] = useState<number>(0);
  const [ingest_done, setIngestDone] = useState<number>(0);

  const [cancel_confirm_open, setCancelConfirmOpen] = useState(false);
  const [cancel_summary, setCancelSummary] = useState<string>("");
  const [complete_summary, setCompleteSummary] = useState<string>("");

  const abort_ref = useRef<AbortController | null>(null);
  const cancel_requested_ref = useRef<boolean>(false);

  // pause gate
  const pause_requested_ref = useRef<boolean>(false);
  const pause_resolve_ref = useRef<null | ((v: "continue" | "cancel") => void)>(null);

  function wait_for_cancel_decision(): Promise<"continue" | "cancel"> {
    return new Promise((resolve) => {
      pause_resolve_ref.current = resolve;
    });
  }

  const run_ingestion_id_ref = useRef<string>("");

  // ---------- clear modal ----------
  const [clear_modal_open, setClearModalOpen] = useState(false);
  const [clear_confirm_checked, setClearConfirmChecked] = useState(false);

  // ---------- guidance copy (improved) ----------
  const consent_copy = useMemo(() => {
    return {
      bullets: [
        "Selected email text is stored locally on this device to enable search and question answering.",
        "Your indexed email text is not uploaded by this tool to a remote server.",
        "You can clear the local index at any time."
      ],
      largeWarning:
        "Large selections can take several minutes. Outlook may feel slower while indexing."
    };
  }, []);

  // ---------- startup ----------
  async function load_folders_with_token(token: string) {
    setBusy("Loading folders…");
    const data = await get_folders(token);
    setFolders((data.folders || []) as folder[]);
    setBusy("");
  }

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
      await load_folders_with_token(token);

      setStatus("Ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setScreen("SIGNIN");
      setStatus("Error");
      set_error("Initialization failed. Please sign in again.", String(e?.message || e));
    }
  }

  useEffect(() => {
    initialize().catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Keep transcript pinned to bottom on new messages (when user is near bottom)
  useEffect(() => {
    // Basic: always scroll to bottom after new message; adequate for MVP
    scroll_transcript_to_bottom();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chat_msgs.length]);

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
      await load_folders_with_token(token);

      setStatus("Ready");
    } catch (e: any) {
      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
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

      setTokenOk(false);
      setAccessToken("");
      setUserLabel("");
      setIndexStatus(null);
      setFolders([]);
      reset_ingest_state();

      clear_chat_storage();
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(false);

      setScreen("SIGNIN");
      setStatus("Not signed in");
    } catch (e: any) {
      setStatus("Error");
      set_error("Sign out failed.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  function reset_ingest_state() {
    setIngestStep("SELECT");
    setIngestMode("FOLDERS");
    setSelectedFolderIds(new Set());
    setFolderLimitInput("100");

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
    pause_requested_ref.current = false;
    pause_resolve_ref.current = null;
    run_ingestion_id_ref.current = "";
  }

  // ---------- selection helpers ----------
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

  // ---------- messages loading ----------
  async function load_email_folder(fid: string) {
    setEmailFolderId(fid);
    setMessages([]);
    setMessagesNextLink(null);
    setSelectedEmailIds(new Set());
    setMessagesFilter("");

    if (!fid) return;
    if (!token_ok || !access_token) return;

    setBusy("Loading emails…");
    set_error("");

    try {
      const data = await get_messages(access_token, fid, 25);
      setMessages((data.value || []) as graph_message[]);
      setMessagesNextLink((data as any)["@odata.nextLink"] || null);
    } catch (e: any) {
      set_error("Failed to load emails. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function load_more_messages() {
    if (!token_ok || !access_token) return;
    if (!messages_next_link) return;

    setBusy("Loading more emails…");
    set_error("");

    try {
      const data = await get_messages_page(access_token, { next_link: messages_next_link, top: 25 });
      const page = (data.value || []) as graph_message[];
      setMessages((prev) => [...prev, ...page]);
      setMessagesNextLink((data as any)["@odata.nextLink"] || null);
    } catch (e: any) {
      set_error("Failed to load more emails. Please try again.", String(e?.message || e));
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

  // Cancel pauses at checkpoints
  function cancel_clicked() {
    pause_requested_ref.current = true;
    setCancelConfirmOpen(true);
  }

  function cancel_continue() {
    setCancelConfirmOpen(false);
    pause_requested_ref.current = false;
    pause_resolve_ref.current?.("continue");
    pause_resolve_ref.current = null;
  }

  function cancel_confirm_now() {
    setCancelConfirmOpen(false);
    cancel_requested_ref.current = true;
    pause_requested_ref.current = false;

    pause_resolve_ref.current?.("cancel");
    pause_resolve_ref.current = null;

    if (abort_ref.current) abort_ref.current.abort();

    setCancelSummary("Indexing cancelled. Some emails may already be indexed.");
    setIngestStep("CANCELLED");
  }

  function progress_done_now(): number {
    return run_phase === "FETCHING" ? fetch_done : ingest_done;
  }

  function phase_label(p: Phase): { title: string; desc: string } {
    if (p === "FETCHING") return { title: "Fetching emails", desc: "Collecting selected messages from Microsoft Graph…" };
    if (p === "STORING") return { title: "Storing locally", desc: "Saving selected email text on this device…" };
    if (p === "INDEXING") return { title: "Indexing for search", desc: "Building the local search index…" };
    return { title: "Done", desc: "Index updated." };
  }

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
    pause_requested_ref.current = false;
    pause_resolve_ref.current = null;

    const run_id = make_id("ing");
    run_ingestion_id_ref.current = run_id;
    const run_label = `${ingest_mode} ${now_iso_local()}`;

    setIngestStep("RUNNING");
    setRunPhase("FETCHING");
    setRunTotal(selection_summary.effectiveTotal || null);
    setFetchDone(0);
    setIngestDone(0);

    const MIN_PHASE_MS = 250;
    let phaseStart = Date.now();

    const ac = new AbortController();
    abort_ref.current = ac;

    try {
      // 1) Build IDs
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
            if (pause_requested_ref.current) {
              const decision = await wait_for_cancel_decision();
              if (decision === "cancel") throw new Error("CANCELLED_BY_USER");
            }
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

      if (!message_ids.length) throw new Error("No emails selected.");

      // make FETCHING visible
      {
        const elapsed = Date.now() - phaseStart;
        if (elapsed < MIN_PHASE_MS) await sleep(MIN_PHASE_MS - elapsed);
      }

      // 2) STORING: batched ingestion
      setRunPhase("STORING");
      phaseStart = Date.now();

      const folder_id_to_send = ingest_mode === "EMAILS" ? (email_folder_id || "selected") : "multi";
      const BATCH_SIZE = 5;
      const batches = chunk(message_ids, BATCH_SIZE);

      let ingested = 0;

      for (const b of batches) {
        if (pause_requested_ref.current) {
          const decision = await wait_for_cancel_decision();
          if (decision === "cancel") throw new Error("CANCELLED_BY_USER");
        }
        if (cancel_requested_ref.current) throw new Error("CANCELLED_BY_USER");

        await ingest_messages(access_token, folder_id_to_send, b, {
          ingestion_id: run_id,
          ingestion_label: run_label,
          ingest_mode: ingest_mode
        });

        ingested += b.length;
        setIngestDone(ingested);
      }

      // 3) INDEXING (brief)
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

      setCompleteSummary(`${summary} Updated: ${fmt_dt(st.last_updated)}.`);
      setIngestStep("COMPLETE");
    } catch (e: any) {
      const msg = String(e?.message || e);

      if (msg === "CANCELLED_BY_USER") {
        await refresh_index_status();
        setCancelSummary("Indexing cancelled. Some emails may already be indexed.");
        setIngestStep("CANCELLED");
        return;
      }

      setIngestStep("SELECT");
      set_error(
        "Indexing failed. Try again, or clear the local index and re-ingest.",
        msg
      );
    } finally {
      abort_ref.current = null;
      setBusy("");
      pause_requested_ref.current = false;
      pause_resolve_ref.current = null;
    }
  }

  async function clear_index_confirmed() {
    if (!token_ok || !access_token) return;

    set_error("");
    setBusy("Clearing…");

    try {
      if (selected_clear_mode === "ALL") {
        await clear_index(access_token);
      } else {
        if (!selected_ingestion_id_to_clear) throw new Error("No ingestion selected.");
        await clear_ingestion(null, selected_ingestion_id_to_clear);
      }

      await refresh_index_status("INDEX");
      reset_ingest_state();
      setClearModalOpen(false);
      setClearConfirmChecked(false);

      if (selected_clear_mode === "ALL") {
        clear_chat_storage();
        setChatMsgs([]);
        setChatState("EMPTY");
        setChatExpiredNote(false);
      }
    } catch (e: any) {
      set_error("Failed to clear. Please try again.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- Phase 4 Chat send with better feedback ----------
  async function send_chat() {
    if (!token_ok || !access_token) {
      alert("Please sign in first.");
      return;
    }
    if ((index_status?.indexed_count || 0) <= 0) {
      alert("Index emails first using Index management.");
      return;
    }

    const q = draft.trim();
    if (!q) return;

    setSending(true);
    set_error("");

    const userMsg: ChatMsg = {
      id: make_id("chat"),
      role: "user",
      text: q,
      created_at_ms: Date.now()
    };
    append_chat(userMsg);

    setDraft("");

    const assistantId = make_id("chat");
    const assistantPlaceholder: ChatMsg = {
      id: assistantId,
      role: "assistant",
      text: "Searching indexed emails…",
      created_at_ms: Date.now(),
      sources_open: false
    };
    append_chat(assistantPlaceholder);

    // staged status text (pure UI feedback)
    const t1 = setTimeout(() => patch_chat_msg(assistantId, { text: "Reading top matches…", sources: [] }), 700);
    const t2 = setTimeout(() => patch_chat_msg(assistantId, { text: "Drafting answer…", sources: [] }), 1400);

    try {
      const res = await ask_question(access_token, q, 4);
      const answer = String(res.answer || "");
      const allSources = (res.sources || []) as source[];
      const topSources = allSources.slice(0, 3); // IMPORTANT: only show top 3

      patch_chat_msg(assistantId, {
        text: answer || "No answer.",
        sources: topSources,
        sources_open: false
      });
    } catch (e: any) {
      patch_chat_msg(assistantId, { text: "Query failed. Please try again.", sources: [] });
      set_error("Query failed. If the local index is unavailable, clear it and re-ingest.", String(e?.message || e));
    } finally {
      clearTimeout(t1);
      clearTimeout(t2);
      setSending(false);
    }
  }

  async function open_email(message_id: string, fallback_weblink: string) {
    if (!token_ok || !access_token) {
      window.open(fallback_weblink, "_blank", "noopener,noreferrer");
      return;
    }

    try {
      const res = await get_message_link(access_token, message_id);
      const fresh = (res.weblink || "") as string;
      const urlToUse = fresh || fallback_weblink;

      const local = try_convert_to_outlook_desktop_url(urlToUse);
      window.open(local || urlToUse, "_blank", "noopener,noreferrer");
    } catch {
      const local = try_convert_to_outlook_desktop_url(fallback_weblink);
      window.open(local || fallback_weblink, "_blank", "noopener,noreferrer");
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
        </div>

        {token_ok ? (
          <div className="op-nav">
            <button className="op-btn op-btnGhost" onClick={() => setScreen("CHAT")} disabled={(index_status?.indexed_count || 0) <= 0}>
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

  function render_indexed_panel() {
    if (index_empty) {
      return (
        <div className="op-banner op-bannerStrong" style={{ marginBottom: 10 }}>
          <div className="op-bannerTitle">No emails indexed yet</div>
          <div className="op-bannerText">Select folders or emails to build a local index.</div>
        </div>
      );
    }

    return (
      <div className="op-banner" style={{ marginBottom: 10 }}>
        <div className="op-bannerTitle">Index is ready</div>
        <div className="op-bannerText">
          Indexed <strong>{index_status?.indexed_count ?? 0}</strong> emails • Updated <strong>{fmt_dt(index_status?.last_updated)}</strong>
        </div>
      </div>
    );
  }

  function render_wizard_steps(active: IngestStep) {
    const steps: { key: IngestStep; label: string }[] = [
      { key: "SELECT", label: "Select" },
      { key: "PREVIEW", label: "Confirm" },
      { key: "RUNNING", label: "Progress" },
      { key: "COMPLETE", label: "Done" }
    ];

    return (
      <div className="op-wizardSteps">
        {steps.map((s) => (
          <span key={s.key} className={`op-chip ${s.key === active ? "op-chipActive" : ""}`}>
            {s.label}
          </span>
        ))}
      </div>
    );
  }

  function render_select_step() {
    return (
      <div className="op-fit">
        <div style={{ marginBottom: 10 }}>
          {render_indexed_panel()}
          <div className="op-row" style={{ justifyContent: "space-between" }}>
            <div>
              <div className="op-cardTitle">Index management</div>
              <div className="op-muted">Choose folders or specific emails to index locally.</div>
            </div>

            <div className="op-seg" role="tablist" aria-label="Index mode">
              <button
                className="op-segBtn"
                aria-selected={ingest_mode === "FOLDERS"}
                onClick={() => setIngestMode("FOLDERS")}
              >
                Folders
              </button>
              <button
                className="op-segBtn"
                aria-selected={ingest_mode === "EMAILS"}
                onClick={() => setIngestMode("EMAILS")}
              >
                Emails
              </button>
            </div>
          </div>

          {render_wizard_steps("SELECT")}
        </div>

        <div className="op-fitBody">
          {ingest_mode === "FOLDERS" ? (
            <div className="op-card" style={{ padding: 12 }}>
              <div className="op-row" style={{ alignItems: "flex-end" }}>
                <div style={{ flex: 1, minWidth: 200 }}>
                  <div className="op-label">Folder search</div>
                  <input className="op-input" value={folder_filter} onChange={(e) => setFolderFilter(e.target.value)} placeholder="Filter folders…" />
                </div>

                <div style={{ width: 140 }}>
                  <div className="op-label">Per-folder limit</div>
                  <input
                    className="op-input"
                    value={folder_limit_input}
                    inputMode="numeric"
                    onChange={(e) => setFolderLimitInput(e.target.value)}
                    onBlur={() => {
                      // commit clamp on blur
                      const n = Number(folder_limit_input);
                      if (!Number.isFinite(n) || n <= 0) setFolderLimitInput("100");
                      else setFolderLimitInput(String(Math.max(1, Math.min(2000, Math.floor(n)))));
                    }}
                    placeholder="100"
                  />
                </div>
              </div>

              <div className="op-spacer" />

              <div className="op-list">
                <div className="op-listScroll">
                  {filtered_folders.length ? (
                    filtered_folders.map((f) => {
                      const checked = selected_folder_ids.has(f.id);
                      return (
                        <div key={f.id} className="op-item">
                          <div className="op-itemRow">
                            <input
                              type="checkbox"
                              checked={checked}
                              onChange={() => toggle_folder(f.id)}
                              style={{ marginTop: 2 }}
                            />
                            <div className="op-itemMain">
                              <div className="op-itemTitle">{f.displayName}</div>
                              <div className="op-itemMeta">
                                {typeof f.totalItemCount === "number" ? `${f.totalItemCount} total` : "Count unavailable"} • Indexing up to {folder_limit}
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })
                  ) : (
                    <div className="op-item">
                      <div className="op-muted">No folders match your search.</div>
                    </div>
                  )}
                </div>
              </div>

              <div className="op-spacer" />
              <div className="op-banner">
                <div className="op-bannerTitle">Selection summary</div>
                <div className="op-bannerText">
                  Folders: <strong>{selection_summary.folderCount}</strong> • Approx emails: <strong>{selection_summary.effectiveTotal}</strong>
                </div>
                {selection_summary.large ? (
                  <div className="op-bannerText" style={{ marginTop: 6 }}>
                    <strong style={{ color: "var(--warning)" }}>Large selection:</strong> {consent_copy.largeWarning}
                  </div>
                ) : null}
              </div>
            </div>
          ) : (
            <div className="op-card" style={{ padding: 12 }}>
              <div className="op-row">
                <div style={{ flex: 1, minWidth: 200 }}>
                  <div className="op-label">Folder</div>
                  <select
                    className="op-select"
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
                </div>
                <div style={{ flex: 1, minWidth: 200 }}>
                  <div className="op-label">Search emails</div>
                  <input className="op-input" value={messages_filter} onChange={(e) => setMessagesFilter(e.target.value)} placeholder="Filter by subject or sender…" />
                </div>
              </div>

              <div className="op-spacer" />

              <div className="op-row">
                <button className="op-btn" onClick={select_all_filtered_emails} disabled={!filtered_messages.length}>
                  Select filtered
                </button>
                <button className="op-btn" onClick={clear_email_selection} disabled={!selected_email_ids.size}>
                  Clear selection
                </button>
                <button className="op-btn" onClick={load_more_messages} disabled={!messages_next_link}>
                  Load more
                </button>
                <span className="op-muted">
                  Selected: <strong>{selection_summary.emailCount}</strong>
                </span>
              </div>

              <div className="op-spacer" />

              <div className="op-list">
                <div className="op-listScroll">
                  {filtered_messages.length ? (
                    filtered_messages.map((m) => {
                      const from = m.from?.emailAddress?.address || "";
                      const checked = selected_email_ids.has(m.id);
                      return (
                        <div key={m.id} className="op-item">
                          <div className="op-itemRow">
                            <input type="checkbox" checked={checked} onChange={() => toggle_email(m.id)} style={{ marginTop: 2 }} />
                            <div className="op-itemMain">
                              <div className="op-itemTitle">{m.subject || "(no subject)"}</div>
                              <div className="op-itemMeta">{from} • {m.receivedDateTime || ""}</div>
                              <div className="op-itemPreview">{truncate(m.bodyPreview || "", 140)}</div>
                            </div>
                          </div>
                        </div>
                      );
                    })
                  ) : (
                    <div className="op-item">
                      <div className="op-muted">{email_folder_id ? "No emails loaded or no matches." : "Select a folder to view emails."}</div>
                    </div>
                  )}
                </div>
              </div>

              <div className="op-spacer" />
              <div className="op-banner">
                <div className="op-bannerTitle">Selection summary</div>
                <div className="op-bannerText">
                  Emails selected: <strong>{selection_summary.emailCount}</strong>
                </div>
                {selection_summary.large ? (
                  <div className="op-bannerText" style={{ marginTop: 6 }}>
                    <strong style={{ color: "var(--warning)" }}>Large selection:</strong> {consent_copy.largeWarning}
                  </div>
                ) : null}
              </div>
            </div>
          )}
        </div>

        <div className="op-spacer" />

        <div className="op-row" style={{ justifyContent: "space-between" }}>
          <div className="op-muted">
            {index_empty ? "Index required before chat." : "You can index more at any time."}
          </div>
          <button className="op-btn op-btnPrimary" disabled={!can_continue_from_select()} onClick={go_preview}>
            Continue
          </button>
        </div>

        {/* Danger zone stays compact and within screen */}
        <div className="op-spacer" />
        <details className="op-card" style={{ padding: 12 }}>
          <summary style={{ cursor: "pointer", fontSize: 12, fontWeight: 800, color: "var(--danger)" }}>
            Danger zone
          </summary>
          <div className="op-spacer" />
          <div className="op-muted">Clear the full index or clear a specific ingestion run.</div>
          <div className="op-spacer" />
          <button className="op-btn op-btnDanger" onClick={() => { setClearModalOpen(true); setClearConfirmChecked(false); }}>
            Clear…
          </button>
          {render_clear_modal()}
        </details>
      </div>
    );
  }

  function render_preview_step() {
    return (
      <div className="op-fit">
        <div style={{ marginBottom: 10 }}>
          <div className="op-cardTitle">Confirm and start</div>
          <div className="op-muted">Review your selection and confirm local indexing.</div>
          {render_wizard_steps("PREVIEW")}
        </div>

        <div className="op-fitBody">
          <div className="op-card op-cardBody">
            <div className="op-banner op-bannerStrong">
              <div className="op-bannerTitle">What happens next</div>
              <div className="op-bannerText">
                We will fetch the selected emails via Microsoft Graph, store their text locally on this device, and build a local search index.
              </div>
            </div>

            <div className="op-spacer" />

            <div className="op-banner">
              <div className="op-bannerTitle">Selection</div>
              <div className="op-bannerText">
                Mode: <strong>{ingest_mode === "FOLDERS" ? "Folders" : "Emails"}</strong> • Total (approx): <strong>{selection_summary.effectiveTotal}</strong>
                {ingest_mode === "FOLDERS" ? <> • Per-folder: <strong>{folder_limit}</strong></> : null}
              </div>
            </div>

            {selection_summary.large ? (
              <div className="op-spacer" />
            ) : null}

            {selection_summary.large ? (
              <div className="op-banner op-bannerWarn">
                <div className="op-bannerTitle">Large selection</div>
                <div className="op-bannerText">{consent_copy.largeWarning}</div>
                <div className="op-spacer" />
                <label className="op-muted" style={{ display: "block" }}>
                  <input
                    type="checkbox"
                    checked={large_ack_checked}
                    onChange={(e) => setLargeAckChecked(e.target.checked)}
                    style={{ marginRight: 8 }}
                  />
                  I understand this may take several minutes.
                </label>
              </div>
            ) : null}

            <div className="op-spacer" />

            <div className="op-banner">
              <div className="op-bannerTitle">Consent</div>
              <ul className="op-muted" style={{ margin: "6px 0 0 18px" }}>
                {consent_copy.bullets.map((b) => <li key={b}>{b}</li>)}
              </ul>

              <div className="op-spacer" />
              <label className="op-muted" style={{ display: "block" }}>
                <input
                  type="checkbox"
                  checked={consent_checked}
                  onChange={(e) => setConsentChecked(e.target.checked)}
                  style={{ marginRight: 8 }}
                />
                I consent to local storage and indexing.
              </label>
            </div>
          </div>
        </div>

        <div className="op-spacer" />
        <div className="op-row" style={{ justifyContent: "space-between" }}>
          <button className="op-btn" onClick={back_to_select}>Back</button>
          <button
            className="op-btn op-btnPrimary"
            disabled={!consent_checked || (selection_summary.large && !large_ack_checked)}
            onClick={run_ingestion}
          >
            Start indexing
          </button>
        </div>
      </div>
    );
  }

  function render_running_step() {
    const doneNow = progress_done_now();
    const pct = run_total ? Math.min(100, Math.round((doneNow / Math.max(1, run_total)) * 100)) : 20;
    const p = phase_label(run_phase);

    return (
      <div className="op-fit">
        <div style={{ marginBottom: 10 }}>
          <div className="op-cardTitle">Indexing</div>
          <div className="op-muted">You can cancel at any time.</div>
          {render_wizard_steps("RUNNING")}
        </div>

        <div className="op-fitBody">
          <div className="op-progressWrap">
            <div className="op-row" style={{ justifyContent: "space-between" }}>
              <div>
                <div className="op-cardTitle">{p.title}</div>
                <div className="op-muted">{p.desc}</div>
              </div>
              <div className="op-row">
                <div className="op-spinner" aria-hidden="true" />
                <div className="op-muted">{run_total ? `${doneNow} / ${run_total}` : `${doneNow}`}</div>
              </div>
            </div>

            <div className="op-spacer" />
            <div className="op-progressBar">
              <div className="op-progressFill" style={{ width: `${pct}%` }} />
            </div>

            <div className="op-spacer" />
            <button className="op-btn" onClick={cancel_clicked}>Cancel</button>

            {render_cancel_confirm_modal()}
          </div>
        </div>
      </div>
    );
  }

  function render_cancelled_step() {
    return (
      <div className="op-fit">
        <div className="op-banner op-bannerDanger">
          <div className="op-bannerTitle">Indexing cancelled</div>
          <div className="op-bannerText">{cancel_summary || "Indexing cancelled. Some emails may already be indexed."}</div>
        </div>

        <div className="op-spacer" />
        <div className="op-row">
          <button className="op-btn op-btnDanger" onClick={() => { setClearModalOpen(true); setClearConfirmChecked(false); }}>
            Clear…
          </button>
          <button className="op-btn" onClick={() => { setCancelSummary(""); cancel_requested_ref.current = false; setIngestStep("SELECT"); }}>
            Return to selection
          </button>
        </div>

        {render_clear_modal()}
      </div>
    );
  }

  function render_complete_step() {
    return (
      <div className="op-fit">
        <div className="op-banner op-bannerStrong">
          <div className="op-bannerTitle">Index updated</div>
          <div className="op-bannerText">{complete_summary || "Indexing completed successfully."}</div>
        </div>

        <div className="op-spacer" />
        <div className="op-row">
          <button className="op-btn op-btnPrimary" onClick={() => setScreen("CHAT")}>Go to chat</button>
          <button className="op-btn" onClick={() => reset_ingest_state()}>Index more</button>
        </div>
      </div>
    );
  }

  function render_cancel_confirm_modal() {
    if (!cancel_confirm_open) return null;
    return (
      <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
        <div className="op-bannerTitle">Cancel indexing now?</div>
        <div className="op-bannerText">Indexing is paused. Some emails may already be indexed.</div>
        <div className="op-spacer" />
        <div className="op-row">
          <button className="op-btn" onClick={cancel_continue}>Continue</button>
          <button className="op-btn op-btnDanger" onClick={cancel_confirm_now}>Cancel now</button>
        </div>
      </div>
    );
  }

  function render_clear_modal() {
    if (!clear_modal_open) return null;

    const hasIngestions = ingestions.length > 0;

    return (
      <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
        <div className="op-bannerTitle">Clear local index</div>
        <div className="op-bannerText">Choose what to clear. This cannot be undone.</div>

        <div className="op-spacer" />

        <label className="op-muted" style={{ display: "block", marginBottom: 6 }}>
          <input
            type="radio"
            name="clearMode"
            checked={selected_clear_mode === "ALL"}
            onChange={() => setSelectedClearMode("ALL")}
            style={{ marginRight: 8 }}
          />
          Clear entire index
        </label>

        <label className="op-muted" style={{ display: "block", marginBottom: 6, opacity: hasIngestions ? 1 : 0.5 }}>
          <input
            type="radio"
            name="clearMode"
            checked={selected_clear_mode === "ONE"}
            disabled={!hasIngestions}
            onChange={() => setSelectedClearMode("ONE")}
            style={{ marginRight: 8 }}
          />
          Clear a specific ingestion run
        </label>

        {selected_clear_mode === "ONE" ? (
          <div style={{ marginTop: 8 }}>
            <div className="op-label">Select ingestion</div>
            <select
              className="op-select"
              value={selected_ingestion_id_to_clear}
              onChange={(e) => setSelectedIngestionIdToClear(e.target.value)}
              disabled={!hasIngestions}
            >
              {ingestions.map((ing) => (
                <option key={ing.ingestion_id} value={ing.ingestion_id}>
                  {ing.created_at} — {ing.label} — {ing.email_count} emails
                </option>
              ))}
            </select>
          </div>
        ) : null}

        <div className="op-spacer" />

        <label className="op-muted" style={{ display: "block" }}>
          <input
            type="checkbox"
            checked={clear_confirm_checked}
            onChange={(e) => setClearConfirmChecked(e.target.checked)}
            style={{ marginRight: 8 }}
          />
          I understand this cannot be undone.
        </label>

        <div className="op-spacer" />
        <div className="op-row">
          <button className="op-btn op-btnDanger" disabled={!clear_confirm_checked} onClick={clear_index_confirmed}>
            Clear
          </button>
          <button className="op-btn" onClick={() => { setClearModalOpen(false); setClearConfirmChecked(false); }}>
            Cancel
          </button>
        </div>
      </div>
    );
  }

  function render_index_management(is_collapsible_panel: boolean) {
    const body = (
      <div className="op-fit">
        {ingest_step === "SELECT" && render_select_step()}
        {ingest_step === "PREVIEW" && render_preview_step()}
        {ingest_step === "RUNNING" && render_running_step()}
        {ingest_step === "CANCELLED" && render_cancelled_step()}
        {ingest_step === "COMPLETE" && render_complete_step()}
      </div>
    );

    if (!is_collapsible_panel) return body;

    return (
      <div className="op-card">
        <div className="op-cardHeader">
          <div>
            <div className="op-cardTitle">Index management</div>
            <div className="op-muted">Index more emails or clear local data.</div>
          </div>
          <button className="op-btn" onClick={() => setIndexPanelOpen((v) => !v)}>
            {index_panel_open ? "Hide" : "Show"}
          </button>
        </div>

        <div className="op-cardBody" style={{ display: index_panel_open ? "block" : "none" }}>
          {body}
        </div>

        {!index_panel_open ? (
          <div className="op-cardBody">
            <div className="op-muted">
              Indexed <strong>{index_status?.indexed_count ?? 0}</strong> • Updated <strong>{fmt_dt(index_status?.last_updated)}</strong>
            </div>
          </div>
        ) : null}
      </div>
    );
  }

  // ---------- Chat transcript ----------
  function format_time(ms: number) {
    const d = new Date(ms);
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${hh}:${mm}`;
  }

  function toggle_sources(msgId: string) {
    setChatMsgs((prev) => {
      const next = prev.map((m) => (m.id === msgId ? { ...m, sources_open: !m.sources_open } : m));
      save_chat_to_storage(next);
      return next;
    });
  }

  function render_chat_transcript() {
    if (!chat_msgs.length) {
      return (
        <div className="op-muted" style={{ marginTop: 10 }}>
          {chat_expired_note ? "Previous chat expired after 1 hour." : "Ask a question about your indexed emails to begin."}
          <div className="op-helpNote">
            Example: “Summarise the latest emails from my manager” or “Find invoices from last month.”
          </div>
        </div>
      );
    }

    return (
      <div className="op-chatTranscript" ref={transcript_ref}>
        {chat_state === "RESTORED" ? (
          <div className="op-banner" style={{ marginBottom: 10 }}>
            <div className="op-bannerTitle">Chat restored</div>
            <div className="op-bannerText">Restored messages from the last hour.</div>
          </div>
        ) : null}

        {chat_msgs.map((m) => {
          const isUser = m.role === "user";
          const srcs = (m.sources || []).slice(0, 3); // enforce top 3 at render time too
          return (
            <div key={m.id} className={`op-bubble ${isUser ? "op-bubbleUser" : ""}`}>
              <div className="op-bubbleHeader">
                <div className="op-bubbleRole">{isUser ? "You" : "Assistant"}</div>
                <div className="op-bubbleTime">{format_time(m.created_at_ms)}</div>
              </div>

              <div className="op-bubbleText">{m.text}</div>

              {!isUser && srcs.length ? (
                <>
                  <div
                    className="op-sourcesToggle"
                    onClick={() => toggle_sources(m.id)}
                    role="button"
                    tabIndex={0}
                    onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") toggle_sources(m.id); }}
                  >
                    Sources ({srcs.length}) {m.sources_open ? "▲" : "▼"}
                  </div>

                  {m.sources_open ? (
                    <div>
                      {srcs.map((s) => (
                        <div key={`${m.id}_${s.message_id}`} className="op-sourceRow">
                          <div className="op-sourceTop">
                            <div style={{ minWidth: 0 }}>
                              <div className="op-itemTitle">{truncate(s.subject || "(no subject)", 52)}</div>
                              <div className="op-itemMeta">{truncate(s.sender, 30)} • {truncate(s.received_dt, 26)}</div>
                            </div>
                            <button className="op-sourceLink" onClick={() => open_email(s.message_id, s.weblink)}>
                              Open
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </>
              ) : null}
            </div>
          );
        })}
      </div>
    );
  }

  function render_chat() {
    const no_index = (index_status?.indexed_count || 0) <= 0;

    if (no_index) {
      return (
        <div className="op-fit">
          <div className="op-banner op-bannerStrong">
            <div className="op-bannerTitle">Index required</div>
            <div className="op-bannerText">Index emails first using Index management.</div>
          </div>
          <div className="op-spacer" />
          {render_index_management(false)}
        </div>
      );
    }

    return (
      <div className="op-chatShell">
        {render_index_management(true)}

        <div className="op-spacer" />
        <div className="op-card op-fit" style={{ flex: 1, minHeight: 0 }}>
          <div className="op-cardHeader">
            <div>
              <div className="op-cardTitle">Chat</div>
              <div className="op-muted">Answers are based only on locally indexed emails.</div>
            </div>
            <button
              className="op-btn"
              onClick={() => {
                clear_chat_storage();
                setChatMsgs([]);
                setChatState("EMPTY");
                setChatExpiredNote(false);
              }}
              disabled={sending}
              title="Clear local chat history"
            >
              Clear chat
            </button>
          </div>

          <div className="op-cardBody op-fitBody" style={{ padding: 12 }}>
            {render_chat_transcript()}
          </div>

          <div className="op-chatComposer">
            <textarea
              className="op-textarea"
              value={draft}
              onChange={(e) => setDraft(e.target.value)}
              rows={2}
              placeholder="Ask a question…"
              disabled={sending}
            />
            <div className="op-row" style={{ justifyContent: "space-between", marginTop: 8 }}>
              <div className="op-muted">
                {sending ? "Working…" : "Tip: Keep questions specific for better results."}
              </div>
              <button className="op-btn op-btnPrimary" onClick={send_chat} disabled={!draft.trim() || sending}>
                Send
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ---------- main render ----------
  return (
    <div className="op-app">
      {render_header()}
      <div className="op-body">
        {screen === "SIGNIN" && render_signin()}
        {screen === "INDEX" && render_index_management(false)}
        {screen === "CHAT" && render_chat()}
      </div>
    </div>
  );
}
