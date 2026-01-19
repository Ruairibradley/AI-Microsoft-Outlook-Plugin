import type { Folder, GraphMessage } from "../ingest/types";

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
  select_all_filtered: () => void;
  clear_selection: () => void;

  filtered_messages: GraphMessage[];

  selected_count: number;
  is_large: boolean;
  large_warning_text: string;
}) {
  return (
    <div className="op-card" style={{ padding: 12 }}>
      <div className="op-row">
        <div style={{ flex: 1, minWidth: 200 }}>
          <div className="op-label">Folder</div>
          <select
            className="op-select"
            value={props.email_folder_id}
            onChange={(e) => props.on_load_folder(e.target.value)}
            disabled={!props.token_ok}
          >
            <option value="">Select…</option>
            {props.folders.map((f) => (
              <option key={f.id} value={f.id}>
                {f.displayName} {typeof f.totalItemCount === "number" ? `(${f.totalItemCount})` : ""}
              </option>
            ))}
          </select>
        </div>

        <div style={{ flex: 1, minWidth: 200 }}>
          <div className="op-label">Search emails</div>
          <input
            className="op-input"
            value={props.messages_filter}
            onChange={(e) => props.set_messages_filter(e.target.value)}
            placeholder="Filter by subject or sender…"
          />
        </div>
      </div>

      <div className="op-spacer" />

      <div className="op-row">
        <button className="op-btn" onClick={props.select_all_filtered} disabled={!props.filtered_messages.length}>
          Select filtered
        </button>
        <button className="op-btn" onClick={props.clear_selection} disabled={!props.selected_email_ids.size}>
          Clear selection
        </button>
        <button className="op-btn" onClick={props.on_load_more} disabled={!props.messages_next_link}>
          Load more
        </button>
        <span className="op-muted">
          Selected: <strong>{props.selected_count}</strong>
        </span>
      </div>

      <div className="op-spacer" />

      <div className="op-list">
        <div className="op-listScroll">
          {props.filtered_messages.length ? (
            props.filtered_messages.map((m) => {
              const from = m.from?.emailAddress?.address || "";
              const checked = props.selected_email_ids.has(m.id);
              return (
                <div key={m.id} className="op-item">
                  <div className="op-itemRow">
                    <input type="checkbox" checked={checked} onChange={() => props.toggle_email(m.id)} style={{ marginTop: 2 }} />
                    <div className="op-itemMain">
                      <div className="op-itemTitle">{m.subject || "(no subject)"}</div>
                      <div className="op-itemMeta">{from} • {m.receivedDateTime || ""}</div>
                      <div className="op-itemPreview">{(m.bodyPreview || "").slice(0, 140)}{(m.bodyPreview || "").length > 140 ? "…" : ""}</div>
                    </div>
                  </div>
                </div>
              );
            })
          ) : (
            <div className="op-item">
              <div className="op-muted">{props.email_folder_id ? "No emails loaded or no matches." : "Select a folder to view emails."}</div>
            </div>
          )}
        </div>
      </div>

      <div className="op-spacer" />

      <div className="op-banner">
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
