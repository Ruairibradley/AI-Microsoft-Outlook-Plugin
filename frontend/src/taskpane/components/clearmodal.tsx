import { useEffect, useState } from "react";
import { list_ingestions, type IngestionInfo } from "../../api/backend";

export function ClearModal(props: {
  open: boolean;
  onClose: () => void;

  // When confirmed, parent decides whether to clear ALL or ONE and executes the action
  onConfirm: (args: { mode: "ALL" | "ONE"; ingestion_id?: string }) => Promise<void>;

  confirm_checked: boolean;
  set_confirm_checked: (v: boolean) => void;

  token_ok: boolean;
  access_token: string;

  busy_text?: string;
  error_text?: string;
}) {
  const [mode, setMode] = useState<"ALL" | "ONE">("ALL");
  const [ingestions, setIngestions] = useState<IngestionInfo[]>([]);
  const [selected_ingestion_id, setSelectedIngestionId] = useState<string>("");

  useEffect(() => {
    if (!props.open) return;

    // Load ingestion list when modal opens (best effort)
    (async () => {
      try {
        const res = await list_ingestions(null, 50);
        const arr = (res.ingestions || []) as IngestionInfo[];
        setIngestions(arr);
        if (!selected_ingestion_id && arr.length) setSelectedIngestionId(arr[0].ingestion_id);
      } catch {
        // ignore
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.open]);

  if (!props.open) return null;

  const hasIngestions = ingestions.length > 0;

  return (
    <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
      <div className="op-bannerTitle">Clear local index</div>
      <div className="op-bannerText">Choose what to clear. This cannot be undone.</div>

      {props.busy_text ? <div className="op-helpNote">{props.busy_text}</div> : null}
      {props.error_text ? (
        <div className="op-helpNote" style={{ color: "var(--danger)" }}>
          {props.error_text}
        </div>
      ) : null}

      <div className="op-spacer" />

      <label className="op-muted" style={{ display: "block", marginBottom: 6 }}>
        <input
          type="radio"
          name="clearMode"
          checked={mode === "ALL"}
          onChange={() => setMode("ALL")}
          style={{ marginRight: 8 }}
        />
        Clear entire index
      </label>

      <label className="op-muted" style={{ display: "block", marginBottom: 6, opacity: hasIngestions ? 1 : 0.5 }}>
        <input
          type="radio"
          name="clearMode"
          checked={mode === "ONE"}
          disabled={!hasIngestions}
          onChange={() => setMode("ONE")}
          style={{ marginRight: 8 }}
        />
        Clear a specific ingestion run
      </label>

      {mode === "ONE" ? (
        <div style={{ marginTop: 8 }}>
          <div className="op-label">Select ingestion</div>
          <select
            className="op-select"
            value={selected_ingestion_id}
            onChange={(e) => setSelectedIngestionId(e.target.value)}
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
          checked={props.confirm_checked}
          onChange={(e) => props.set_confirm_checked(e.target.checked)}
          style={{ marginRight: 8 }}
        />
        I understand this cannot be undone.
      </label>

      <div className="op-spacer" />

      <div className="op-row">
        <button
          className="op-btn op-btnDanger"
          disabled={!props.confirm_checked || !props.token_ok || !props.access_token}
          onClick={() => props.onConfirm({ mode, ingestion_id: mode === "ONE" ? selected_ingestion_id : undefined })}
        >
          Clear
        </button>
        <button
          className="op-btn"
          onClick={() => {
            props.set_confirm_checked(false);
            props.onClose();
          }}
        >
          Cancel
        </button>
      </div>
    </div>
  );
}
