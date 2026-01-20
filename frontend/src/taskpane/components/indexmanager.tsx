import { useMemo, useRef, useState } from "react";
import { clear_index, clear_ingestion, get_messages, get_messages_page, type IndexStatus } from "../../api/backend";

import { run_ingestion } from "../ingest/runner";
import type { Folder, GraphMessage, IngestMode, IngestStep, Phase } from "../ingest/types";

import { ClearModal } from "./clearmodal";
import { EmailPicker } from "./emailpicker";
import { IngestPreview } from "./ingestpreview";
import { IngestProgress } from "./ingestprogress";
import { IngestResult } from "./ingestresult";

type IndexView = "MAIN" | "CLEAR_STORAGE";

function fmt_dt(s: string | null | undefined) {
  if (!s) return "—";
  return s.replace("T", " ");
}

export function IndexManager(props: {
  token_ok: boolean;
  access_token: string;
  folders: Folder[];
  index_status: IndexStatus | null;
  onIndexChanged: () => Promise<void>;
  onNavigate?: (to: "CHAT" | "INDEX") => void;
}) {
  const [view, setView] = useState<IndexView>("MAIN");

  const [busy, setBusy] = useState<string>("");
  const [err, setErr] = useState<string>("");

  // ---------- wizard state ----------
  const [ingest_step, setIngestStep] = useState<IngestStep>("SELECT");

  // Always EMAILS mode in the simplified UX
  const ingest_mode: IngestMode = "EMAILS";

  // email selection
  const [email_folder_id, setEmailFolderId] = useState<string>("");
  const [messages, setMessages] = useState<GraphMessage[]>([]);
  const [messages_next_link, setMessagesNextLink] = useState<string | null>(null);
  const [messages_filter, setMessagesFilter] = useState<string>("");
  const [selected_email_ids, setSelectedEmailIds] = useState<Set<string>>(new Set());

  // preview
  const [consent_checked, setConsentChecked] = useState(false);
  const [large_ack_checked, setLargeAckChecked] = useState(false);

  // progress
  const [run_phase, setRunPhase] = useState<Phase>("FETCHING");
  const [run_total, setRunTotal] = useState<number | null>(null);
  const [fetch_done, setFetchDone] = useState<number>(0);
  const [ingest_done, setIngestDone] = useState<number>(0);

  // cancel pause gate
  const [cancel_confirm_open, setCancelConfirmOpen] = useState(false);
  const cancel_requested_ref = useRef<boolean>(false);

  const pause_requested_ref = useRef<boolean>(false);
  const pause_resolve_ref = useRef<null | ((v: "continue" | "cancel") => void)>(null);

  function wait_for_decision(): Promise<"continue" | "cancel"> {
    return new Promise((resolve) => {
      pause_resolve_ref.current = resolve;
    });
  }

  // result summaries
  const [cancel_summary, setCancelSummary] = useState<string>("");
  const [complete_summary, setCompleteSummary] = useState<string>("");

  // clear storage sub-screen state
  const [clear_confirm_checked, setClearConfirmChecked] = useState(false);

  const consent_copy = useMemo(() => {
    return {
      bullets: [
        "Selected email text is stored locally on this device to enable search and question answering.",
        "Your indexed email text is not uploaded by this tool to a remote server.",
        "You can clear the local index at any time."
      ],
      largeWarning: "Large selections can take several minutes. Outlook may feel slower while indexing."
    };
  }, []);

  const filtered_messages = useMemo(() => {
    const q = messages_filter.trim().toLowerCase();
    if (!q) return messages;
    return messages.filter((m) => {
      const subj = (m.subject || "").toLowerCase();
      const from = (m.from?.emailAddress?.address || "").toLowerCase();
      return subj.includes(q) || from.includes(q);
    });
  }, [messages, messages_filter]);

  const selection_summary = useMemo(() => {
    const emailCount = selected_email_ids.size;
    const effectiveTotal = emailCount;
    return {
      emailCount,
      effectiveTotal,
      large: effectiveTotal >= 200
    };
  }, [selected_email_ids]);

  function can_continue_from_select(): boolean {
    return selected_email_ids.size > 0;
  }

  function toggle_email(id: string) {
    setSelectedEmailIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function clear_email_selection() {
    setSelectedEmailIds(new Set());
  }

  async function load_email_folder(folderId: string) {
    setEmailFolderId(folderId);
    setMessages([]);
    setMessagesNextLink(null);
    setSelectedEmailIds(new Set());
    setMessagesFilter("");
    setErr("");
    setBusy("");
  }

  async function load_more_messages() {
    if (!props.token_ok || !props.access_token) return;
    if (!messages_next_link) return;

    setBusy("Loading more emails…");
    setErr("");
    try {
      const data = await get_messages_page(props.access_token, { next_link: messages_next_link, top: 25 });
      const page = (data.value || []) as GraphMessage[];
      setMessages((prev) => [...prev, ...page]);
      setMessagesNextLink((data as any)["@odata.nextLink"] || null);
    } catch (e: any) {
      setErr(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function ensure_messages_loaded_for_specific() {
    if (messages.length > 0) return;
    if (!props.token_ok || !props.access_token) return;
    if (!email_folder_id) {
      setErr("Select a folder first.");
      return;
    }

    setBusy("Loading emails…");
    setErr("");
    try {
      const data = await get_messages(props.access_token, email_folder_id, 25);
      setMessages((data.value || []) as GraphMessage[]);
      setMessagesNextLink((data as any)["@odata.nextLink"] || null);
    } catch (e: any) {
      setErr(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function select_most_recent(n: number) {
    setErr("");
    if (!props.token_ok || !props.access_token) return;
    if (!email_folder_id) {
      setErr("Select a folder first.");
      return;
    }

    const target = Math.max(1, Math.min(5000, Math.floor(n || 1)));

    setBusy(`Selecting most recent ${target}…`);
    try {
      let localMessages: GraphMessage[] = [];
      let next: string | null = null;

      const firstTop = Math.min(100, target);
      const first = await get_messages(props.access_token, email_folder_id, firstTop);
      localMessages = (first.value || []) as GraphMessage[];
      next = (first as any)["@odata.nextLink"] || null;

      while (localMessages.length < target && next) {
        const data = await get_messages_page(props.access_token, { next_link: next, top: 100 });
        const page = (data.value || []) as GraphMessage[];
        localMessages = [...localMessages, ...page];
        next = (data as any)["@odata.nextLink"] || null;

        setBusy(`Selecting most recent ${target}… (${Math.min(localMessages.length, target)}/${target})`);
      }

      const UI_LIMIT = 250;
      setMessages(localMessages.slice(0, UI_LIMIT));
      setMessagesNextLink(next);

      const ids = localMessages.slice(0, target).map((m) => m.id);
      setSelectedEmailIds(new Set(ids));
    } catch (e: any) {
      setErr(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  async function select_all_in_folder() {
    setErr("");
    if (!props.token_ok || !props.access_token) return;
    if (!email_folder_id) {
      setErr("Select a folder first.");
      return;
    }

    setBusy("Selecting all emails in folder…");
    try {
      const allIds = new Set<string>();

      const UI_LIMIT = 250;
      let uiMessages: GraphMessage[] = [];

      const first = await get_messages(props.access_token, email_folder_id, 100);
      const firstPage = (first.value || []) as GraphMessage[];
      firstPage.forEach((m) => allIds.add(m.id));
      uiMessages = firstPage.slice(0, UI_LIMIT);

      let next = (first as any)["@odata.nextLink"] || null;
      setBusy(`Selecting all emails in folder… (${allIds.size})`);

      while (next) {
        const data = await get_messages_page(props.access_token, { next_link: next, top: 100 });
        const page = (data.value || []) as GraphMessage[];
        page.forEach((m) => allIds.add(m.id));

        if (uiMessages.length < UI_LIMIT) {
          const remaining = UI_LIMIT - uiMessages.length;
          uiMessages = [...uiMessages, ...page.slice(0, remaining)];
        }

        next = (data as any)["@odata.nextLink"] || null;
        setBusy(`Selecting all emails in folder… (${allIds.size})`);
      }

      setMessages(uiMessages);
      setMessagesNextLink(null);
      setSelectedEmailIds(allIds);
    } catch (e: any) {
      setErr(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  function go_preview() {
    setErr("");
    setCancelSummary("");
    setConsentChecked(false);
    setLargeAckChecked(false);
    setIngestStep("PREVIEW");
  }

  function back_to_select() {
    setErr("");
    setIngestStep("SELECT");
  }

  // ---------- cancel flow ----------
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

    // Resolve the pause gate as "cancel"
    pause_resolve_ref.current?.("cancel");
    pause_resolve_ref.current = null;

    // Terminal cancelled state — MUST persist until user chooses next action
    setBusy("");
    setErr("");
    setCancelSummary("Indexing cancelled. Some emails may already be indexed.");
    setIngestStep("CANCELLED");
  }

  async function start_indexing() {
    if (!props.token_ok || !props.access_token) return;
    if (!consent_checked) return;
    if (selection_summary.large && !large_ack_checked) return;

    setErr("");
    setCancelSummary("");
    setCompleteSummary("");

    cancel_requested_ref.current = false;
    pause_requested_ref.current = false;
    pause_resolve_ref.current = null;

    setIngestStep("RUNNING");
    setRunPhase("FETCHING");
    setRunTotal(selection_summary.effectiveTotal || null);
    setFetchDone(0);
    setIngestDone(0);

    try {
      const res = await run_ingestion({
        access_token: props.access_token,
        ingest_mode,

        email_folder_id,
        selected_email_ids: Array.from(selected_email_ids),

        folders: props.folders,
        selected_folder_ids: [],
        folder_limit: 0,

        onPhase: setRunPhase,
        onTotal: setRunTotal,
        onFetchDone: setFetchDone,
        onIngestDone: setIngestDone,

        isCancelRequested: () => cancel_requested_ref.current,
        isPauseRequested: () => pause_requested_ref.current,
        waitForDecision: wait_for_decision,

        batch_size: 5,
        page_size: 25,
        min_phase_ms: 250
      });

      // If user cancelled during/after ingestion, treat as cancelled terminal state
      if (cancel_requested_ref.current) {
        await props.onIndexChanged();
        setBusy("");
        setErr("");
        setCancelSummary("Indexing cancelled. Some emails may already be indexed.");
        setIngestStep("CANCELLED");
        return; // critical: do not fall through
      }

      await props.onIndexChanged();

      setCompleteSummary(`Indexed ${res.message_ids.length} selected emails.`);
      setIngestStep("COMPLETE");
    } catch (e: any) {
      const msg = String(e?.message || e);

      // IMPORTANT: cancellation must not fall through to SELECT
      if (msg === "CANCELLED_BY_USER" || cancel_requested_ref.current) {
        await props.onIndexChanged();
        setBusy("");
        setErr("");
        setCancelSummary("Indexing cancelled. Some emails may already be indexed.");
        setIngestStep("CANCELLED");
        return; // critical: prevents overwriting CANCELLED with SELECT
      }

      setErr(msg);
      setIngestStep("SELECT");
    } finally {
      // Do NOT touch ingest_step here.
      pause_requested_ref.current = false;
      pause_resolve_ref.current = null;
    }
  }

  async function clear_confirmed(args: { mode: "ALL" | "ONE"; ingestion_id?: string }) {
    if (!props.token_ok || !props.access_token) return;
    setBusy("Clearing…");
    setErr("");

    try {
      if (args.mode === "ALL") await clear_index(props.access_token);
      else {
        if (!args.ingestion_id) throw new Error("No ingestion selected.");
        await clear_ingestion(null, args.ingestion_id);
      }

      setClearConfirmChecked(false);
      setView("MAIN");

      await props.onIndexChanged();
      setIngestStep("SELECT");
    } catch (e: any) {
      setErr(String(e?.message || e));
    } finally {
      setBusy("");
    }
  }

  // ---------- CLEAR STORAGE SUB-SCREEN ----------
  if (view === "CLEAR_STORAGE") {
    return (
      <div className="op-card op-fit">
        <div className="op-cardHeader">
          <div>
            <div className="op-cardTitle">Clear storage</div>
            <div className="op-muted">Remove local indexed email data.</div>
          </div>
          <button
            className="op-btn"
            onClick={() => {
              setClearConfirmChecked(false);
              setView("MAIN");
            }}
          >
            Back
          </button>
        </div>

        <div className="op-cardBody op-fitBody">
          <div className="op-banner op-bannerWarn">
            <div className="op-bannerTitle">Before you clear</div>
            <div className="op-bannerText">
              Clearing storage removes locally indexed email text and search data. This cannot be undone.
            </div>
          </div>

          <div className="op-spacer" />

          <ClearModal
            open={true}
            onClose={() => {
              setClearConfirmChecked(false);
              setView("MAIN");
            }}
            onConfirm={clear_confirmed}
            confirm_checked={clear_confirm_checked}
            set_confirm_checked={setClearConfirmChecked}
            token_ok={props.token_ok}
            access_token={props.access_token}
            busy_text={busy}
            error_text={err}
          />
        </div>
      </div>
    );
  }

  // ---------- MAIN VIEW ----------
  function render_select_step() {
    return (
      <div className="op-fit" style={{ display: "flex", flexDirection: "column", minHeight: 0 }}>
        <div style={{ marginBottom: 8 }}>
          <div className="op-cardTitle">Choose emails to search</div>
          <div className="op-muted">Pick a folder, then choose how to select emails.</div>
          {busy ? <div className="op-helpNote">{busy}</div> : null}

          {err ? (
            <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
              <div className="op-bannerTitle">Error</div>
              <div className="op-bannerText">{err}</div>
            </div>
          ) : null}
        </div>

        <div style={{ flex: "1 1 auto", minHeight: 0, overflow: "auto", paddingBottom: 6 }}>
          <EmailPicker
            token_ok={props.token_ok}
            folders={props.folders}
            email_folder_id={email_folder_id}
            set_email_folder_id={setEmailFolderId}
            on_folder_changed={load_email_folder}
            ensure_loaded_for_specific={ensure_messages_loaded_for_specific}
            messages_filter={messages_filter}
            set_messages_filter={setMessagesFilter}
            messages={messages}
            messages_next_link={messages_next_link}
            on_load_more={load_more_messages}
            selected_email_ids={selected_email_ids}
            toggle_email={toggle_email}
            clear_selection={clear_email_selection}
            filtered_messages={filtered_messages}
            selected_count={selection_summary.emailCount}
            is_large={selection_summary.large}
            large_warning_text={consent_copy.largeWarning}
            on_select_all={select_all_in_folder}
            on_select_most_recent={select_most_recent}
          />
        </div>

        <div className="op-row" style={{ justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
          <button className="op-btn op-btnDanger" onClick={() => setView("CLEAR_STORAGE")}>
            Clear storage
          </button>

          <button className="op-btn op-btnPrimary" disabled={!can_continue_from_select()} onClick={go_preview}>
            Continue
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="op-card op-fit">
      <div className="op-cardBody op-fitBody">
        {ingest_step === "SELECT" && render_select_step()}

        {ingest_step === "PREVIEW" && (
          <IngestPreview
            ingest_mode={ingest_mode}
            approx_total={selection_summary.effectiveTotal}
            folder_limit={0}
            is_large={selection_summary.large}
            large_warning_text={consent_copy.largeWarning}
            consent_bullets={consent_copy.bullets}
            consent_checked={consent_checked}
            set_consent_checked={setConsentChecked}
            large_ack_checked={large_ack_checked}
            set_large_ack_checked={setLargeAckChecked}
            error_text={err}
            onBack={back_to_select}
            onStart={start_indexing}
          />
        )}

        {ingest_step === "RUNNING" && (
          <IngestProgress
            run_phase={run_phase}
            run_total={run_total}
            fetch_done={fetch_done}
            ingest_done={ingest_done}
            cancel_confirm_open={cancel_confirm_open}
            onCancelClick={cancel_clicked}
            onCancelContinue={cancel_continue}
            onCancelNow={cancel_confirm_now}
          />
        )}

        {ingest_step === "CANCELLED" && (
          <IngestResult
            kind="CANCELLED"
            summary={cancel_summary || "Indexing cancelled. Some emails may already be indexed."}
            onOpenClear={() => setView("CLEAR_STORAGE")}
            onReturnToSelect={() => setIngestStep("SELECT")}
            onIndexMore={() => setIngestStep("SELECT")}
          />
        )}

        {ingest_step === "COMPLETE" && (
        <div className="op-fit">
          <IngestResult
            kind="COMPLETE"
            summary={
              (complete_summary || "Indexing completed successfully.") +
              (props.index_status?.last_updated ? ` • Updated: ${fmt_dt(props.index_status.last_updated)}` : "")
            }
            onOpenClear={() => setView("CLEAR_STORAGE")}
            onReturnToSelect={() => setIngestStep("SELECT")}
            onIndexMore={() => setIngestStep("SELECT")}

            // NEW: show "Go to chat" action after indexing
            onGoChat={() => props.onNavigate?.("CHAT")}
          />
        </div>
      )}
      </div>
    </div>
  );
}
