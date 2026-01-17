import { useEffect, useMemo, useRef, useState } from "react";
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

function now_iso_local() {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function make_id(prefix: string): string {
  const c: any = (globalThis as any).crypto;
  if (c?.randomUUID) return `${prefix}_${c.randomUUID()}`;
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
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

// ------------------- Phase 4: chat persistence -------------------
const CHAT_TTL_MS = 60 * 60 * 1000; // 1 hour
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
        sources: Array.isArray(m.sources) ? m.sources : undefined
      }));

    return { msgs, ts_ms };
  } catch {
    return null;
  }
}

function save_chat_to_storage(msgs: ChatMsg[]) {
  try {
    localStorage.setItem(LS_CHAT_KEY, JSON.stringify(msgs));
    localStorage.setItem(LS_CHAT_TS_KEY, String(Date.now()));
  } catch {
    // ignore (storage may be blocked)
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

  // ---------- Phase 4 chat state ----------
  const [chat_state, setChatState] = useState<ChatState>("EMPTY");
  const [chat_msgs, setChatMsgs] = useState<ChatMsg[]>([]);
  const [chat_expired_note, setChatExpiredNote] = useState<boolean>(false);

  function append_chat(msg: ChatMsg) {
    setChatMsgs((prev) => {
      const next = [...prev, msg];
      save_chat_to_storage(next);
      return next;
    });
    setChatState("ACTIVE");
  }

  function replace_last_assistant(msgId: string, newText: string, newSources?: source[]) {
    setChatMsgs((prev) => {
      const next = prev.map((m) => {
        if (m.id !== msgId) return m;
        return { ...m, text: newText, sources: newSources };
      });
      save_chat_to_storage(next);
      return next;
    });
    setChatState("ACTIVE");
  }

  function restore_chat_if_valid(indexExists: boolean) {
    if (!indexExists) {
      // If no index, we do not restore chat.
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
      // expired
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
      // do not restore chat when index is empty
      restore_chat_if_valid(false);
    } else {
      if (nextPreferredScreen) setScreen(nextPreferredScreen);
      else if (screen === "SIGNIN") setScreen("CHAT");

      // Phase 4: restore chat once we know index exists
      restore_chat_if_valid(true);
    }

    await refresh_ingestions();
  }

  // ---------- graph data ----------
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
  const [folder_limit, setFolderLimit] = useState<number>(100);

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

  // abort only affects FETCHING loop
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

  // ---------- clear index modal ----------
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

      // Phase 4: clear chat storage on sign-out (prevents cross-account restore confusion)
      clear_chat_storage();
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(false);

      setScreen("SIGNIN");
      setStatus("not signed in");
    } catch (e: any) {
      setStatus("error");
      set_error("Sign out failed.", String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

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
    pause_requested_ref.current = false;
    pause_resolve_ref.current = null;
    run_ingestion_id_ref.current = "";
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

    setCancelSummary("Indexing cancelled. Some items may already be indexed.");
    setIngestStep("CANCELLED");
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

      if (!message_ids.length) throw new Error("No messages selected.");

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

      {
        const elapsed = Date.now() - phaseStart;
        if (elapsed < MIN_PHASE_MS) await sleep(MIN_PHASE_MS - elapsed);
      }

      // 3) INDEXING: show brief phase
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
    } catch (e: any) {
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
      pause_requested_ref.current = false;
      pause_resolve_ref.current = null;
    }
  }

  async function clear_index_confirmed() {
    if (!token_ok || !access_token) return;

    set_error("");
    setBusy("clearing local index...");

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

      // Phase 4: if index is cleared entirely, also clear chat (prevents orphaned transcript)
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

  // ------------- Phase 4: Chat send / persist / restore -------------
  const [draft, setDraft] = useState<string>("");
  const [sending, setSending] = useState<boolean>(false);

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

    const placeholderId = make_id("chat");
    const assistantPlaceholder: ChatMsg = {
      id: placeholderId,
      role: "assistant",
      text: "Working…",
      created_at_ms: Date.now()
    };
    append_chat(assistantPlaceholder);

    try {
      const res = await ask_question(access_token, q, 4);
      const answer = String(res.answer || "");
      const srcs = (res.sources || []) as source[];
      replace_last_assistant(placeholderId, answer || "No answer.", srcs);
    } catch (e: any) {
      replace_last_assistant(placeholderId, "Query failed. Please try again.", []);
      set_error("Query failed. If the local index is unavailable, clear it and re-ingest.", String(e?.message || e));
    } finally {
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
          {token_ok && user_label ? <span> — <strong>Signed in:</strong> {user_label}</span> : null}
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
                width: run_total ? `${Math.min(100, Math.round((doneNow / Math.max(1, run_total)) * 100))}%` : "25%",
                background: "#bbb"
              }}
            />
          </div>
          <div style={{ fontSize: 12, opacity: 0.8, marginTop: 6 }}>
            {run_total ? <span>Processed {doneNow} / {run_total} emails</span> : <span>Processed {doneNow} emails</span>}
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
          Indexing is paused. Some emails may already be indexed.
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

    const hasIngestions = ingestions.length > 0;

    return (
      <div style={{ border: "2px solid #c00", padding: 12, background: "#fff5f5", marginTop: 10 }}>
        <strong>Clear local index?</strong>

        <div style={{ marginTop: 10, fontSize: 12 }}>
          <label style={{ display: "block", marginBottom: 6 }}>
            <input
              type="radio"
              name="clearMode"
              checked={selected_clear_mode === "ALL"}
              onChange={() => setSelectedClearMode("ALL")}
            />{" "}
            Clear entire index (all ingestions)
          </label>

          <label style={{ display: "block", marginBottom: 6, opacity: hasIngestions ? 1 : 0.5 }}>
            <input
              type="radio"
              name="clearMode"
              checked={selected_clear_mode === "ONE"}
              disabled={!hasIngestions}
              onChange={() => setSelectedClearMode("ONE")}
            />{" "}
            Clear a specific ingestion run
          </label>

          {selected_clear_mode === "ONE" ? (
            <div style={{ marginTop: 8 }}>
              <div style={{ fontSize: 12, opacity: 0.85, marginBottom: 4 }}>Select ingestion</div>
              <select
                style={{ width: "100%" }}
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
        </div>

        <ul style={{ marginTop: 10, fontSize: 12 }}>
          <li>Deletes locally stored indexed email text</li>
          <li>Deletes local search index entries for the selected scope</li>
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
            Clear
          </button>{" "}
          <button onClick={() => { setClearModalOpen(false); setClearConfirmChecked(false); }}>
            Cancel
          </button>
        </div>
      </div>
    );
  }

  function render_index_select_scope() {
    return (
      <div>
        {index_empty ? (
          <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5", marginBottom: 10 }}>
            <strong>No emails indexed yet</strong>
            <div style={{ fontSize: 12, marginTop: 6 }}>Select folders or individual emails to index locally.</div>
          </div>
        ) : (
          <div style={{ border: "1px solid #ddd", padding: 10, background: "#fafafa", marginBottom: 10 }}>
            <div style={{ fontSize: 12 }}>
              <strong>Indexed:</strong> {index_status?.indexed_count ?? 0} emails — <strong>Last updated:</strong> {fmt_dt(index_status?.last_updated)}
            </div>
          </div>
        )}

        <h3 style={{ marginTop: 0 }}>Index management</h3>

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
          <button onClick={() => setIngestMode("FOLDERS")} disabled={ingest_mode === "FOLDERS"}>Select folders</button>
          <button onClick={() => setIngestMode("EMAILS")} disabled={ingest_mode === "EMAILS"}>Select emails</button>
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
              {!filtered_folders.length ? <div style={{ fontSize: 12, opacity: 0.8, padding: 6 }}>No folders match your filter.</div> : null}
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
              <button disabled={!filtered_messages.length} onClick={select_all_filtered_emails}>Select all (filtered)</button>
              <button disabled={!selected_email_ids.size} onClick={clear_email_selection}>Clear selection</button>
              <button disabled={!messages_next_link} onClick={load_more_messages}>Load more</button>
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
          <div style={{ fontSize: 12 }}><strong>Selection summary</strong></div>
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
          <button disabled={!can_continue_from_select()} onClick={go_preview}>Continue</button>
        </div>

        <hr />

        <h3>Clear local index</h3>
        <div style={{ fontSize: 12, opacity: 0.8 }}>
          Clear the full index or clear a specific ingestion run.
        </div>
        <button style={{ marginTop: 8 }} disabled={!token_ok} onClick={() => { setClearModalOpen(true); setClearConfirmChecked(false); }}>
          Clear…
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
          <div style={{ fontSize: 12 }}><strong>Summary</strong></div>
          <div style={{ fontSize: 12, opacity: 0.85, marginTop: 6 }}>
            Mode: {ingest_mode === "FOLDERS" ? "Folders" : "Emails"}
          </div>
          <div style={{ fontSize: 12, opacity: 0.85 }}>
            Total to index (approx): {selection_summary.effectiveTotal}
          </div>
          {ingest_mode === "FOLDERS" ? (
            <div style={{ fontSize: 12, opacity: 0.85 }}>Per-folder limit: {folder_limit}</div>
          ) : null}
        </div>

        {selection_summary.large ? (
          <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5", marginTop: 10 }}>
            <strong>Large selection warning</strong>
            <div style={{ fontSize: 12, marginTop: 6 }}>
              Large selections take longer to process and may impact responsiveness.
            </div>
            <label style={{ display: "block", fontSize: 12, marginTop: 8 }}>
              <input type="checkbox" checked={large_ack_checked} onChange={(e) => setLargeAckChecked(e.target.checked)} />{" "}
              I understand this may take several minutes.
            </label>
          </div>
        ) : null}

        <div style={{ border: "1px solid #ddd", padding: 10, marginTop: 10 }}>
          <strong>Consent required</strong>
          <p style={{ marginTop: 8, fontSize: 12 }}>{privacy_text}</p>

          <label style={{ display: "block", marginBottom: 8, fontSize: 12 }}>
            <input type="checkbox" checked={consent_checked} onChange={(e) => setConsentChecked(e.target.checked)} />{" "}
            I consent to local storage and indexing.
          </label>
        </div>

        <div style={{ marginTop: 10 }}>
          <button onClick={back_to_select}>Back</button>{" "}
          <button disabled={!consent_checked || (selection_summary.large && !large_ack_checked)} onClick={run_ingestion}>
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
          <div style={{ fontSize: 12, marginTop: 6 }}>{cancel_summary || "Indexing cancelled. Some items may already be indexed."}</div>
        </div>

        <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={() => { setClearModalOpen(true); setClearConfirmChecked(false); }}>
            Clear…
          </button>
          <button onClick={() => { setCancelSummary(""); cancel_requested_ref.current = false; setIngestStep("SELECT"); }}>
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
          <div style={{ fontSize: 12, marginTop: 6 }}>{complete_summary || "Indexing completed successfully."}</div>
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
          <button onClick={() => setIndexPanelOpen((v) => !v)}>{index_panel_open ? "Hide" : "Show"}</button>
        </div>

        {index_panel_open ? (
          <div style={{ marginTop: 10 }}>{body}</div>
        ) : (
          <div style={{ fontSize: 12, opacity: 0.8, marginTop: 8 }}>
            Indexed: {index_status?.indexed_count ?? 0} — Last updated: {fmt_dt(index_status?.last_updated)}
          </div>
        )}
      </div>
    );
  }

  // ---------------- Phase 4: Chat UI ----------------
  function render_chat_transcript() {
    if (!chat_msgs.length) {
      return (
        <div style={{ fontSize: 12, opacity: 0.8, marginTop: 8 }}>
          {chat_expired_note ? "Previous chat expired after 1 hour." : "No messages in this session yet."}
        </div>
      );
    }

    return (
      <div style={{ marginTop: 10 }}>
        {chat_state === "RESTORED" ? (
          <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 8 }}>
            Restored chat from the last hour.
          </div>
        ) : null}

        {chat_msgs.map((m) => (
          <div key={m.id} style={{ marginBottom: 10 }}>
            <div
              style={{
                border: "1px solid #ddd",
                background: m.role === "user" ? "#f7f7ff" : "#fff",
                padding: 10,
                borderRadius: 6
              }}
            >
              <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 6 }}>
                <strong>{m.role === "user" ? "You" : "Assistant"}</strong>
              </div>
              <div style={{ whiteSpace: "pre-wrap", fontSize: 13 }}>{m.text}</div>
            </div>

            {m.role === "assistant" && m.sources && m.sources.length ? (
              <div style={{ marginTop: 8 }}>
                <div style={{ fontSize: 12, opacity: 0.8, marginBottom: 6 }}>
                  <strong>Sources</strong>
                </div>
                {m.sources.map((s, idx) => (
                  <div key={`${m.id}_${s.message_id}_${idx}`} style={{ border: "1px solid #eee", padding: 8, marginBottom: 8 }}>
                    <div>
                      <strong>{s.subject || "(no subject)"}</strong>
                    </div>
                    <div style={{ fontSize: 12, opacity: 0.8 }}>
                      {s.sender} — {s.received_dt} {typeof s.score === "number" ? `— score: ${s.score.toFixed(4)}` : ""}
                    </div>
                    <details style={{ marginTop: 6 }}>
                      <summary style={{ cursor: "pointer", fontSize: 12 }}>Show excerpt</summary>
                      <div style={{ fontSize: 12, marginTop: 6 }}>{s.snippet}</div>
                    </details>
                    <button style={{ marginTop: 6 }} onClick={() => open_email(s.message_id, s.weblink)}>
                      Open email
                    </button>
                  </div>
                ))}
              </div>
            ) : null}
          </div>
        ))}
      </div>
    );
  }

  function render_chat() {
    const no_index = (index_status?.indexed_count || 0) <= 0;

    if (no_index) {
      return (
        <div>
          <div style={{ border: "1px solid #c00", padding: 10, background: "#fff5f5" }}>
            <strong>No emails indexed</strong>
            <div style={{ fontSize: 12, marginTop: 6 }}>Index emails first using Index management.</div>
          </div>
          <div style={{ marginTop: 10 }}>{render_index_management(false)}</div>
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

        {render_chat_transcript()}

        <div style={{ marginTop: 12 }}>
          <textarea
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            rows={3}
            style={{ width: "100%" }}
            placeholder="Type your message..."
            disabled={sending}
          />
          <div style={{ marginTop: 8, display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button onClick={() => send_chat()} disabled={!token_ok || sending || !draft.trim()}>
              Send
            </button>
            <button
              onClick={() => {
                clear_chat_storage();
                setChatMsgs([]);
                setChatState("EMPTY");
                setChatExpiredNote(false);
              }}
              disabled={sending}
              title="Clears local chat history for this session"
            >
              Clear chat
            </button>
          </div>
        </div>
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
