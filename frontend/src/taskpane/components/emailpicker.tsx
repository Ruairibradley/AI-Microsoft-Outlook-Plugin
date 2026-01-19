import { useMemo } from "react";
import type { Folder, GraphMessage } from "../ingest/types";

function truncate(s: string, n: number) {
  if (!s) return "";
  return s.length <= n ? s : s.slice(0, n) + "…";
}

export function EmailPicker(props: {
  token_ok: boolean;

  folders: Folder[];

  email_folder_id: string;
  set_email_folder_id: (v: string) => void;
  on_load_folder: (folderId: string) => Promise<void>;

  messages_filter: string;
  set_messages_filter: (v: string) => void;

  messages: GraphMessage[];
  messages_next_link: string | null;
  on_load_more: () => Promise<void>;

  selected_email_ids: Set<string>;
  toggle_email: (id: string) => void;

  clear_selection: () => void;

  filtered_messages: GraphMessage[];

  selected_count: number;
  is_large: boolean;
  large_warning_text: string;
}) {
  const empty_text = useMemo(() => {
    if (!props.email_folder_id) return "Select a folder to view emails.";
    if (!props.filtered_messages.length) return "No emails match your filter.";
    return "";
  }, [props.email_folder_id, props.filtered_messages.length]);

  const smallBtnStyle: React.CSSProperties = {
    padding: "6px 10px",
    borderRadius: 10,
    fontSize: 12,
    fontWeight: 800
  };

  return (
    <div className="op-card" style={{ padding: 12, maxWidth: "100%" }}>
      {/* Controls (stacked, no overflow) */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "1fr",
          gap: 10,
          maxWidth: "100%"
        }}
      >
        <div style={{ minWidth: 0 }}>
          <div className="op-label">Folder</div>
          <select
            className="op-select"
            value={props.email_folder_id}
            onChange={(e) => props.on_load_folder(e.target.value)}
            disabled={!props.token_ok}
            style={{ width: "100%" }}
          >
            <option value="">Select…</option>
            {props.folders.map((f) => (
              <option key={f.id} value={f.id}>
                {f.displayName} {typeof f.totalItemCount === "number" ? `(${f.totalItemCount})` : ""}
              </option>
            ))}
          </select>
        </div>

        <div style={{ minWidth: 0 }}>
          <div className="op-label">Filter</div>
          <input
            className="op-input"
            value={props.messages_filter}
            onChange={(e) => props.set_messages_filter(e.target.value)}
            placeholder="Subject or sender…"
            style={{ width: "100%" }}
          />
        </div>
      </div>

      <div className="op-spacer" />

      {/* Compact actions: only show when relevant */}
      <div className="op-row" style={{ justifyContent: "space-between", alignItems: "center" }}>
        <div className="op-row" style={{ gap: 8 }}>
          {props.messages_next_link ? (
            <button className="op-btn" style={smallBtnStyle} onClick={props.on_load_more}>
              Load more
            </button>
          ) : null}

          {props.selected_count > 0 ? (
            <button className="op-btn" style={smallBtnStyle} onClick={props.clear_selection}>
              Clear
            </button>
          ) : null}
        </div>

        <div className="op-muted" style={{ whiteSpace: "nowrap" }}>
          Selected: <strong>{props.selected_count}</strong>
        </div>
      </div>

      <div className="op-spacer" />

      {/* List */}
      <div className="op-list" style={{ maxWidth: "100%" }}>
        <div className="op-listScroll">
          {empty_text ? (
            <div className="op-item">
              <div className="op-muted">{empty_text}</div>
            </div>
          ) : (
            props.filtered_messages.map((m) => {
              const from = m.from?.emailAddress?.address || "";
              const checked = props.selected_email_ids.has(m.id);

              return (
                <div key={m.id} className="op-item">
                  <div className="op-itemRow">
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => props.toggle_email(m.id)}
                      style={{ marginTop: 2 }}
                    />

                    <div className="op-itemMain">
                      <div className="op-itemTitle">{truncate(m.subject || "(no subject)", 70)}</div>
                      <div className="op-itemMeta">
                        {truncate(from, 32)} • {truncate(m.receivedDateTime || "", 22)}
                      </div>
                      {/* Minimal preview (optional, single short line) */}
                      {m.bodyPreview ? (
                        <div className="op-itemPreview">{truncate(m.bodyPreview, 90)}</div>
                      ) : null}
                    </div>
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>

      <div className="op-spacer" />

      <div className="op-banner" style={{ maxWidth: "100%" }}>
        <div className="op-bannerTitle">Selection summary</div>
        <div className="op-bannerText">
          Emails selected: <strong>{props.selected_count}</strong>
        </div>
        {props.is_large ? (
          <div className="op-bannerText" style={{ marginTop: 6 }}>
            <strong style={{ color: "var(--warning)" }}>Large selection:</strong> {props.large_warning_text}
          </div>
        ) : null}
      </div>
    </div>
  );
}
