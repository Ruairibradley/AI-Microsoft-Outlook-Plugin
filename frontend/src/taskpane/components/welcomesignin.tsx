import { useState } from "react";

type View = "WELCOME" | "INFO";

export function WelcomeSignIn(props: {
  onSignIn: () => void;
  error?: string;
  error_details?: string;
}) {
  const [view, setView] = useState<View>("WELCOME");

  if (view === "INFO") {
    return (
      <div className="op-card op-fit">
        <div className="op-cardHeader">
          <div>
            <div className="op-cardTitle">How this works</div>
          </div>
          <button className="op-btn" onClick={() => setView("WELCOME")}>
            Back
          </button>
        </div>

        <div className="op-cardBody op-fitBody">
          <div className="op-banner">
            <div className="op-bannerTitle">In plain terms</div>
            <div className="op-bannerText">
              <ul style={{ margin: "6px 0 0 18px" }}>
                <li>You choose folders or emails.</li>
                <li>Selected emails are stored locally on this device.</li>
                <li>Ask questions, receive answers based on relevant emails (with source links).</li>
              </ul>
            </div>
          </div>

          <div className="op-spacer" />

          <div className="op-banner op-bannerStrong">
            <div className="op-bannerTitle">Privacy</div>
            <div className="op-bannerText">
              <ul style={{ margin: "6px 0 0 18px" }}>
                <li>No data is uploaded to a remote server, all processing and storage is local to this device.</li>
                <li>You can clear the local index any time from Index management.</li>
              </ul>
            </div>
          </div>

          <div className="op-spacer" />

          <div className="op-banner">
            <div className="op-bannerTitle">Tip</div>
            <div className="op-bannerText">
              If “Open email” ever fails in your browser, ensure you’re signed into the same mailbox in Outlook on the web.
            </div>
          </div>
        </div>
      </div>
    );
  }

  // WELCOME view
  return (
    <div className="op-card op-fit">
      <div className="op-cardHeader">
        <div>
          <div className="op-cardTitle">Welcome</div>
        </div>
      </div>

      <div className="op-cardBody op-fitBody">
        <div className="op-banner op-bannerStrong">
          <div className="op-bannerTitle">Get started in 3 steps</div>
          <div className="op-bannerText">
            <ol style={{ margin: "6px 0 0 18px" }}>
              <li><strong>Sign in to your Outlook account</strong></li>
              <li><strong>Choose emails you want to search</strong></li>
              <li><strong>Ask questions and explore your emails</strong></li>
            </ol>
          </div>
        </div>

        <div className="op-spacer" />

        <div className="op-row" style={{ justifyContent: "space-between", alignItems: "center" }}>
          <div className="op-muted">
            You may be asked to grant permission to read mail so the add-in can index the emails you select.
          </div>
          <button className="op-btn op-btnPrimary" onClick={props.onSignIn}>
            Sign in
          </button>
        </div>

        <div className="op-row" style={{ justifyContent: "space-between", marginTop: 10 }}>
          <button className="op-linkBtn" onClick={() => setView("INFO")} type="button">
            How does this work?
          </button>
        </div>

        {props.error ? (
          <>
            <div className="op-spacer" />
            <div className="op-banner op-bannerDanger">
              <div className="op-bannerTitle">Sign-in issue</div>
              <div className="op-bannerText">{props.error}</div>

              {props.error_details ? (
                <details style={{ marginTop: 8 }}>
                  <summary style={{ cursor: "pointer", fontSize: 12 }}>Show details</summary>
                  <pre style={{ whiteSpace: "pre-wrap", marginTop: 8, fontSize: 12 }}>{props.error_details}</pre>
                </details>
              ) : null}
            </div>
          </>
        ) : null}
      </div>
    </div>
  );
}
