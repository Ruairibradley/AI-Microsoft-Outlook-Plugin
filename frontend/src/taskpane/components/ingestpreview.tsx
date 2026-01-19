import type { IngestMode } from "../ingest/types";

export function IngestPreview(props: {
  ingest_mode: IngestMode;
  approx_total: number;
  folder_limit: number;

  is_large: boolean;
  large_warning_text: string;

  consent_bullets: string[];

  consent_checked: boolean;
  set_consent_checked: (v: boolean) => void;

  large_ack_checked: boolean;
  set_large_ack_checked: (v: boolean) => void;

  error_text?: string;

  onBack: () => void;
  onStart: () => void;
}) {
  return (
    <div className="op-fit">
      <div style={{ marginBottom: 10 }}>
        <div className="op-cardTitle">Confirm and start</div>
        <div className="op-muted">Review your selection and confirm local indexing.</div>
      </div>

      <div className="op-fitBody">
        <div className="op-card op-cardBody">
          <div className="op-banner op-bannerStrong">
            <div className="op-bannerTitle">What happens next</div>
            <div className="op-bannerText">
              We will fetch selected emails via Microsoft Graph, store their text locally on this device, and build a local search index.
            </div>
          </div>

          <div className="op-spacer" />

          <div className="op-banner">
            <div className="op-bannerTitle">Selection</div>
            <div className="op-bannerText">
              Mode: <strong>{props.ingest_mode === "FOLDERS" ? "Folders" : "Emails"}</strong> • Total (approx):{" "}
              <strong>{props.approx_total}</strong>
              {props.ingest_mode === "FOLDERS" ? (
                <> • Per-folder: <strong>{props.folder_limit}</strong></>
              ) : null}
            </div>
          </div>

          {props.is_large ? (
            <>
              <div className="op-spacer" />
              <div className="op-banner op-bannerWarn">
                <div className="op-bannerTitle">Large selection</div>
                <div className="op-bannerText">{props.large_warning_text}</div>
                <div className="op-spacer" />
                <label className="op-muted" style={{ display: "block" }}>
                  <input
                    type="checkbox"
                    checked={props.large_ack_checked}
                    onChange={(e) => props.set_large_ack_checked(e.target.checked)}
                    style={{ marginRight: 8 }}
                  />
                  I understand this may take several minutes.
                </label>
              </div>
            </>
          ) : null}

          <div className="op-spacer" />

          <div className="op-banner">
            <div className="op-bannerTitle">Consent</div>
            <ul className="op-muted" style={{ margin: "6px 0 0 18px" }}>
              {props.consent_bullets.map((b) => <li key={b}>{b}</li>)}
            </ul>

            <div className="op-spacer" />
            <label className="op-muted" style={{ display: "block" }}>
              <input
                type="checkbox"
                checked={props.consent_checked}
                onChange={(e) => props.set_consent_checked(e.target.checked)}
                style={{ marginRight: 8 }}
              />
              I consent to local storage and indexing.
            </label>
          </div>

          {props.error_text ? (
            <div className="op-banner op-bannerDanger" style={{ marginTop: 10 }}>
              <div className="op-bannerTitle">Error</div>
              <div className="op-bannerText">{props.error_text}</div>
            </div>
          ) : null}
        </div>
      </div>

      <div className="op-spacer" />
      <div className="op-row" style={{ justifyContent: "space-between" }}>
        <button className="op-btn" onClick={props.onBack}>Back</button>
        <button
          className="op-btn op-btnPrimary"
          disabled={!props.consent_checked || (props.is_large && !props.large_ack_checked)}
          onClick={props.onStart}
        >
          Start indexing
        </button>
      </div>
    </div>
  );
}
