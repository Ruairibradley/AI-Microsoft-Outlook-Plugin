export function WelcomeSignIn(props: {
    onSignIn: () => void;
    error?: string;
    error_details?: string;
  }) {
    return (
      <div className="op-card op-fit">
        <div className="op-cardHeader">
          <div>
            <div className="op-cardTitle">Welcome</div>
            <div className="op-muted">A privacy-first assistant for your Outlook mailbox.</div>
          </div>
        </div>

        <div className="op-cardBody">
          <div className="op-banner op-bannerStrong">
            <div className="op-bannerTitle">What this add-in does</div>
            <div className="op-bannerText">
              It indexes selected email text locally on this device so you can search and ask questions quickly. You remain in control and can clear the local index at any time.
            </div>
          </div>

          <div className="op-spacer" />

          <div className="op-banner">
            <div className="op-bannerTitle">How to get started</div>
            <div className="op-bannerText">
              <ol style={{ margin: "6px 0 0 18px" }}>
                <li>Sign in to Microsoft Graph.</li>
                <li>Choose folders or emails to index locally.</li>
                <li>Ask questions in Chat.</li>
              </ol>
            </div>
          </div>

          <div className="op-spacer" />

          <div className="op-row" style={{ justifyContent: "space-between", alignItems: "center" }}>
            <div className="op-muted">
              You may be asked to grant permission to read mail for indexing.
            </div>
            <button className="op-btn op-btnPrimary" onClick={props.onSignIn}>
              Sign in
            </button>
          </div>

          <div className="op-helpNote">
            Tip: If email links fail in your browser later, ensure you’re signed into the same mailbox in Outlook on the web (cookie/session mismatch is the common cause).
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
