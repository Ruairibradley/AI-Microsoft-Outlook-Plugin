import { useMemo } from "react";
import type { Folder } from "../ingest/types";

export function FolderPicker(props: {
  folders: Folder[];

  folder_filter: string;
  set_folder_filter: (v: string) => void;

  folder_limit_input: string;
  set_folder_limit_input: (v: string) => void;
  folder_limit_value: number;

  selected_folder_ids: Set<string>;
  toggle_folder: (id: string) => void;

  approx_total: number;
  folder_count: number;
  is_large: boolean;
  large_warning_text: string;
}) {
  const filtered_folders = useMemo(() => {
    const q = props.folder_filter.trim().toLowerCase();
    if (!q) return props.folders;
    return props.folders.filter((f) => (f.displayName || "").toLowerCase().includes(q));
  }, [props.folders, props.folder_filter]);

  return (
    <div className="op-card" style={{ padding: 12, maxWidth: "100%" }}>
      {/* Responsive controls: stack on narrow panes */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "1fr",
          gap: 10,
          maxWidth: "100%"
        }}
      >
        <div style={{ minWidth: 0 }}>
          <div className="op-label">Folder search</div>
          <input
            className="op-input"
            value={props.folder_filter}
            onChange={(e) => props.set_folder_filter(e.target.value)}
            placeholder="Type to filter folders…"
            style={{ width: "100%" }}
          />
        </div>

        <div style={{ minWidth: 0 }}>
          <div className="op-label">Per-folder limit</div>
          <input
            className="op-input"
            value={props.folder_limit_input}
            inputMode="numeric"
            onChange={(e) => props.set_folder_limit_input(e.target.value)}
            onBlur={() => {
              const n = Number(props.folder_limit_input);
              if (!Number.isFinite(n) || n <= 0) props.set_folder_limit_input("100");
              else props.set_folder_limit_input(String(Math.max(1, Math.min(2000, Math.floor(n)))));
            }}
            placeholder="100"
            style={{ width: "100%" }}
          />
          <div className="op-helpNote" style={{ marginTop: 6 }}>
            Index up to <strong>{props.folder_limit_value}</strong> most recent emails per selected folder.
          </div>
        </div>
      </div>

      <div className="op-spacer" />

      <div className="op-list" style={{ maxWidth: "100%" }}>
        <div className="op-listScroll">
          {filtered_folders.length ? (
            filtered_folders.map((f) => {
              const checked = props.selected_folder_ids.has(f.id);

              return (
                <div key={f.id} className="op-item">
                  <div className="op-itemRow">
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => props.toggle_folder(f.id)}
                      style={{ marginTop: 2 }}
                    />
                    <div className="op-itemMain">
                      <div className="op-itemTitle">{f.displayName}</div>
                      <div className="op-itemMeta">
                        {typeof f.totalItemCount === "number" ? `${f.totalItemCount} total` : "Count unavailable"}
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

      <div className="op-banner" style={{ maxWidth: "100%" }}>
        <div className="op-bannerTitle">Selection summary</div>
        <div className="op-bannerText">
          Folders: <strong>{props.folder_count}</strong> • Approx emails: <strong>{props.approx_total}</strong>
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
