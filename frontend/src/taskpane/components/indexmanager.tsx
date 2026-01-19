import { useMemo, useRef, useState } from "react";
import { clear_index, clear_ingestion, get_messages, get_messages_page, type IndexStatus } from "../../api/backend";

import { run_ingestion } from "../ingest/runner";
import type { Folder, GraphMessage, IngestMode, IngestStep, Phase } from "../ingest/types";

import { ClearModal } from "./clearmodal";
import { FolderPicker } from "./folderpicker";
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
  const [ingest_mode, setIngestMode] = useState<IngestMode>("FOLDERS");

  // folder selection
  const [folder_filter, setFolderFilter] = useState("");
  const [selected_folder_ids, setSelectedFolderIds] = useState<Set<string>>(new Set());

  const [folder_limit_input, setFolderLimitInput] = useState<string>("100");
  const folder_limit = useMemo(() => {
    const n = Number(folder_limit_input);
    if (!Number.isFinite(n) || n <= 0) return 100;
    return Math.max(1, Math.min(2000, Math.floor(n)));
  }, [folder_limit_input]);

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
    const folderCount = selected_folder_ids.size;
    const emailCount = selected_email_ids.size;

    const approxFolderTotal = (() => {
      if (!folderCount) return 0;
      const selected = props.folders.filter((f) => selected_folder_ids.has(f.id));
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
  }, [selected_folder_ids, selected_email_ids, props.folders, ingest_mode, folder_limit]);

  function can_continue_from_select(): boolean {
    if (ingest_mode === "FOLDERS") return selected_folder_ids.size > 0;
    return selected_email_ids.size > 0;
  }

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

    if (!folderId) return;
    if (!props.token_ok || !props.access_token) return;

    setBusy("Loading emails…");
    try {
      const data = await get_messages(props.access_token, folderId, 25);
      setMessages((data.value || []) as GraphMessage[]);
      setMessagesNextLink((data as any)["@odata.nextLink"] || null);
    } catch (e: any) {
      setErr(String(e?.message || e));
    } finally {
      setBusy("");
    }
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

  function go_preview() {
    setErr("");
    setCancelSummary("");
    setConsentChecked(false);
    setLargeAckChecked(false);
    setIngestStep("PREVIEW");
  }

  function back_to_select() {
    setErr("");
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
        selected_folder_ids: Array.from(selected_folder_ids),
        folder_limit,

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

      if (cancel_requested_ref.current) throw new Error("CANCELLED_BY_USER");

      await props.onIndexChanged();

      const summary =
        ingest_mode === "EMAILS"
          ? `Indexed ${res.message_ids.length} selected emails.`
          : `Indexed ${res.message_ids.length} emails from selected folder(s).`;

      setCompleteSummary(summary);
      setIngestStep("COMPLETE");
    } catch (e: any) {
      const msg = String(e?.message || e);
      if (msg === "CANCELLED_BY_USER") {
        await props.onIndexChanged();
        setCancelSummary("Indexing cancelled. Some emails may already be indexed.");
        setIngestStep("CANCELLED");
        return;
      }
      setErr(msg);
      setIngestStep("SELECT");
    } finally {
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
          <button className="op-btn" onClick={() => { setClearConfirmChecked(false); setView("MAIN"); }}>
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
            onClose={() => { setClearConfirmChecked(false); setView("MAIN"); }}
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
      <div className="op-fit">
        <div style={{ marginBottom: 10 }}>
          <div className="op-cardTitle">Choose emails to search</div>
          <div className="op-muted">Pick folders or specific emails to include.</div>

          {busy ? <div className="op-helpNote">{busy}</div> : null}

          <div className="op-spacer" />

          {/* Clear storage always left; segmented always right */}
          <div className="op-row" style={{ justifyContent: "space-between", alignItems: "center" }}>
            <button className="op-btn op-btnDanger" onClick={() => setView("CLEAR_STORAGE")}>
              Clear storage
            </button>

            <div className="op-seg" role="tablist" aria-label="Selection mode">
              <button className="op-segBtn" aria-selected={ingest_mode === "FOLDERS"} onClick={() => setIngestMode("FOLDERS")}>
                Folders
              </button>
              <button className="op-segBtn" aria-selected={ingest_mode === "EMAILS"} onClick={() => setIngestMode("EMAILS")}>
                Emails
              </button>
            </div>
          </div>

          {err ? (
            <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
              <div className="op-bannerTitle">Error</div>
              <div className="op-bannerText">{err}</div>
            </div>
          ) : null}
        </div>

        {ingest_mode === "FOLDERS" ? (
          <FolderPicker
            folders={props.folders}
            folder_filter={folder_filter}
            set_folder_filter={setFolderFilter}
            folder_limit_input={folder_limit_input}
            set_folder_limit_input={setFolderLimitInput}
            folder_limit_value={folder_limit}
            selected_folder_ids={selected_folder_ids}
            toggle_folder={toggle_folder}
            approx_total={selection_summary.effectiveTotal}
            folder_count={selection_summary.folderCount}
            is_large={selection_summary.large}
            large_warning_text={consent_copy.largeWarning}
          />
        ) : (
          <EmailPicker
            token_ok={props.token_ok}
            folders={props.folders}
            email_folder_id={email_folder_id}
            set_email_folder_id={setEmailFolderId}
            on_load_folder={load_email_folder}
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
          />
        )}

        <div className="op-spacer" />

        <div className="op-row" style={{ justifyContent: "flex-end" }}>
          <button className="op-btn op-btnPrimary" disabled={!can_continue_from_select()} onClick={go_preview}>
            Continue
          </button>
        </div>

        <div className="op-helpNote">
          You can’t use Chat until you index at least one email.
        </div>
      </div>
    );
  }

  return (
    <div className="op-card op-fit">
      {/* Removed extra “Emails / Choose what to include…” header entirely */}
      <div className="op-cardBody op-fitBody">
        {ingest_step === "SELECT" && render_select_step()}

        {ingest_step === "PREVIEW" && (
          <IngestPreview
            ingest_mode={ingest_mode}
            approx_total={selection_summary.effectiveTotal}
            folder_limit={folder_limit}
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
            />

            <div className="op-spacer" />

            <div className="op-row">
              <button
                className="op-btn op-btnPrimary"
                onClick={() => props.onNavigate?.("CHAT")}
                disabled={(props.index_status?.indexed_count || 0) <= 0}
                title={(props.index_status?.indexed_count || 0) <= 0 ? "Index emails first" : "Go to chat"}
              >
                Go to chat
              </button>
              <button className="op-btn" onClick={() => setIngestStep("SELECT")}>
                Index more
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
