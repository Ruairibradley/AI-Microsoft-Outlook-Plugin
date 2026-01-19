import type { Phase } from "../ingest/types";

export function IngestProgress(props: {
  run_phase: Phase;
  run_total: number | null;
  fetch_done: number;
  ingest_done: number;

  cancel_confirm_open: boolean;
  onCancelClick: () => void;
  onCancelContinue: () => void;
  onCancelNow: () => void;
}) {
  const doneNow = props.run_phase === "FETCHING" ? props.fetch_done : props.ingest_done;
  const pct = props.run_total ? Math.min(100, Math.round((doneNow / Math.max(1, props.run_total)) * 100)) : 20;

  function phase_label(p: Phase): { title: string; desc: string } {
    if (p === "FETCHING") return { title: "Fetching emails", desc: "Collecting selected messages from Microsoft Graph…" };
    if (p === "STORING") return { title: "Storing locally", desc: "Saving selected email text on this device…" };
    if (p === "INDEXING") return { title: "Indexing for search", desc: "Building the local search index…" };
    return { title: "Done", desc: "Index updated." };
  }

  const p = phase_label(props.run_phase);

  return (
    <div className="op-fit">
      <div style={{ marginBottom: 10 }}>
        <div className="op-cardTitle">Indexing</div>
        <div className="op-muted">You can cancel at any time.</div>
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
              <div className="op-muted">{props.run_total ? `${doneNow} / ${props.run_total}` : `${doneNow}`}</div>
            </div>
          </div>

          <div className="op-spacer" />
          <div className="op-progressBar">
            <div className="op-progressFill" style={{ width: `${pct}%` }} />
          </div>

          <div className="op-spacer" />
          <button className="op-btn" onClick={props.onCancelClick}>Cancel</button>

          {props.cancel_confirm_open ? (
            <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
              <div className="op-bannerTitle">Cancel indexing now?</div>
              <div className="op-bannerText">Indexing is paused. Some emails may already be indexed.</div>
              <div className="op-spacer" />
              <div className="op-row">
                <button className="op-btn" onClick={props.onCancelContinue}>Continue</button>
                <button className="op-btn op-btnDanger" onClick={props.onCancelNow}>Cancel now</button>
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
}
