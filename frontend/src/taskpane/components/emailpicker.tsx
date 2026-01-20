import { useMemo, useState } from "react";
import type { Folder, GraphMessage } from "../ingest/types";

function truncate(s: string, n: number) {
  if (!s) return "";
  return s.length <= n ? s : s.slice(0, n) + "…";
}

type Mode = "MAIN" | "MOST_RECENT" | "SPECIFIC";

export function EmailPicker(props: {
  token_ok: boolean;

  folders: Folder[];

  email_folder_id: string;
  set_email_folder_id: (v: string) => void;
  on_folder_changed: (folderId: string) => Promise<void>;

  ensure_loaded_for_specific: () => Promise<void>;

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

  on_select_all: () => Promise<void>;
  on_select_most_recent: (n: number) => Promise<void>;
}) {
  const [mode, setMode] = useState<Mode>("MAIN");

  const [recent_n_input, setRecentNInput] = useState<string>("200");
  const recent_n = useMemo(() => {
    const n = Number(recent_n_input);
    if (!Number.isFinite(n) || n <= 0) return 200;
    return Math.max(1, Math.min(5000, Math.floor(n)));
  }, [recent_n_input]);

  const can_select = props.token_ok && !!props.email_folder_id;

  const folder_label = useMemo(() => {
    const f = props.folders.find((x) => x.id === props.email_folder_id);
    return f?.displayName || "";
  }, [props.folders, props.email_folder_id]);

  const selection_line = useMemo(() => {
    if (!props.email_folder_id) return `Selection summary: ${props.selected_count} email(s) selected.`;
    if (!folder_label) return `Selection summary: ${props.selected_count} email(s) selected.`;
    return `Selection summary: ${props.selected_count} email(s) selected from “${folder_label}”.`;
  }, [props.selected_count, props.email_folder_id, folder_label]);

  const empty_text = useMemo(() => {
    if (!props.email_folder_id) return "Select a folder to continue.";
    if (mode !== "SPECIFIC") return "";
    if (!props.filtered_messages.length) return "No emails match your filter.";
    return "";
  }, [props.email_folder_id, mode, props.filtered_messages.length]);

  const smallBtnStyle: React.CSSProperties = {
    padding: "6px 10px",
    borderRadius: 10,
    fontSize: 12,
    fontWeight: 800
  };

  const utilityBtnStyle: React.CSSProperties = {
    ...smallBtnStyle,
    border: "1px solid rgba(255,255,255,0.35)",
    background: "rgba(255,255,255,0.06)"
  };

  async function on_folder_change(folderId: string) {
    props.set_email_folder_id(folderId);
    props.clear_selection();
    setMode("MAIN");
    setRecentNInput("200");
    props.set_messages_filter("");
    await props.on_folder_changed(folderId);
  }

  async function choose_select_all() {
    setMode("MAIN");
    await props.on_select_all();
  }

  function open_most_recent() {
    setMode("MOST_RECENT");
  }

  async function run_most_recent() {
    await props.on_select_most_recent(recent_n);
    setMode("MAIN");
  }

  async function open_specific() {
    setMode("SPECIFIC");
    await props.ensure_loaded_for_specific();
  }

  if (mode === "SPECIFIC") {
    return (
      <div className="op-card" style={{ padding: 12, maxWidth: "100%" }}>
        <div className="op-row" style={{ justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
          <div>
            <div className="op-cardTitle">Select specific emails</div>
            <div className="op-muted">
              Filter and tick individual emails to include.
              {folder_label ? ` Folder: “${folder_label}”.` : ""}
            </div>
          </div>

          <button className="op-btn" style={smallBtnStyle} onClick={() => setMode("MAIN")}>
            Back
          </button>
        </div>

        <div className="op-spacer" />

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

        <div className="op-spacer" />

        <div className="op-row" style={{ justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>

          <div className="op-row" style={{ gap: 8, flexWrap: "wrap" }}>
            {props.messages_next_link ? (
              <button className="op-btn" style={utilityBtnStyle} onClick={props.on_load_more} disabled={!props.token_ok}>
                Load more
              </button>
            ) : null}

            {props.selected_count > 0 ? (
              <button className="op-btn" style={utilityBtnStyle} onClick={props.clear_selection}>
                Clear selection
              </button>
            ) : null}
          </div>
        </div>

        <div className="op-spacer" />

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
                        {m.bodyPreview ? <div className="op-itemPreview">{truncate(m.bodyPreview, 90)}</div> : null}
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
          <div className="op-bannerText">{selection_line}</div>
          {props.is_large ? (
            <div className="op-bannerText" style={{ marginTop: 6 }}>
              <strong style={{ color: "var(--warning)" }}>Large selection:</strong> {props.large_warning_text}
            </div>
          ) : null}
        </div>
      </div>
    );
  }

  return (
    <div className="op-card" style={{ padding: 12, maxWidth: "100%" }}>
      <div style={{ minWidth: 0 }}>
        <div className="op-label">Folder</div>
        <select
          className="op-select"
          value={props.email_folder_id}
          onChange={(e) => on_folder_change(e.target.value)}
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

      <div className="op-spacer" />

      <div className="op-row" style={{ justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <div className="op-row" style={{ gap: 8, flexWrap: "wrap" }}>
          <button className="op-btn op-btnPrimary" style={smallBtnStyle} onClick={choose_select_all} disabled={!can_select}>
            Select all
          </button>

          <button className="op-btn" style={smallBtnStyle} onClick={open_most_recent} disabled={!can_select}>
            Select most recent
          </button>

          <button className="op-btn" style={smallBtnStyle} onClick={open_specific} disabled={!can_select}>
            Select specific
          </button>

          {props.selected_count > 0 ? (
            <button className="op-btn" style={utilityBtnStyle} onClick={props.clear_selection}>
              Clear selection
            </button>
          ) : null}
        </div>

        <div className="op-muted" style={{ maxWidth: "100%" }}>
          {selection_line}
        </div>
      </div>

      {mode === "MOST_RECENT" ? (
        <div style={{ marginTop: 10 }}>
          {/* Neutral banner (no warning/orange) */}
          <div className="op-banner">
            <div className="op-bannerTitle">Select most recent emails</div>
            <div className="op-bannerText" style={{ marginTop: 6 }}>
              Enter how many emails to select from the top of the folder.
            </div>

            <div className="op-row" style={{ gap: 8, alignItems: "center", marginTop: 10, flexWrap: "wrap" }}>
              <input
                className="op-input"
                value={recent_n_input}
                onChange={(e) => setRecentNInput(e.target.value)}
                inputMode="numeric"
                style={{ width: 110, padding: "6px 10px" }}
                disabled={!can_select}
                aria-label="Most recent count"
              />
              <button className="op-btn op-btnPrimary" style={smallBtnStyle} onClick={run_most_recent} disabled={!can_select}>
                Select {recent_n}
              </button>
              <button className="op-btn" style={smallBtnStyle} onClick={() => setMode("MAIN")}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {props.is_large ? (
        <div style={{ marginTop: 10 }}>
          <div className="op-banner op-bannerWarn">
            <div className="op-bannerTitle">Large selection</div>
            <div className="op-bannerText">{props.large_warning_text}</div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
