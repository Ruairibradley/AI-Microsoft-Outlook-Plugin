import type React from "react";

export function IngestResult(props: {
  kind: "CANCELLED" | "COMPLETE";
  summary: string;

  onReturnToSelect: () => void;
  onIndexMore: () => void;

  onOpenClear: () => void; // opens ClearModal

  // NEW: lets the parent navigate to Chat
  onGoChat?: () => void;
}) {
  // Utility / outline style for "Clear storage" so it stands out without screaming danger-red
  const utilityBtnStyle: React.CSSProperties = {
    border: "1px solid rgba(255,255,255,0.35)",
    background: "rgba(255,255,255,0.06)"
  };

  const primaryRowStyle: React.CSSProperties = {
    justifyContent: "space-between",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 10
  };

  if (props.kind === "CANCELLED") {
    return (
      <div className="op-fit">
        <div className="op-banner op-bannerDanger">
          <div className="op-bannerTitle">Indexing cancelled</div>
          <div className="op-bannerText">{props.summary}</div>
          <div className="op-bannerText" style={{ marginTop: 6 }}>
            Choose what you want to do next.
          </div>
        </div>

        <div className="op-spacer" />

        <div className="op-row" style={primaryRowStyle}>
          <button className="op-btn" style={utilityBtnStyle} onClick={props.onOpenClear}>
            Clear storage
          </button>

          <button className="op-btn op-btnPrimary" onClick={props.onReturnToSelect}>
            Return to selection
          </button>
        </div>
      </div>
    );
  }

  // COMPLETE
  return (
    <div className="op-fit">
      <div className="op-banner op-bannerStrong">
        <div className="op-bannerTitle">Index updated</div>
        <div className="op-bannerText">{props.summary}</div>
      </div>

      <div className="op-spacer" />

      <div className="op-row" style={{ justifyContent: "flex-end", gap: 10, flexWrap: "wrap" }}>
        {props.onGoChat ? (
          <button className="op-btn" style={utilityBtnStyle} onClick={props.onGoChat}>
            Go to chat
          </button>
        ) : null}

        <button className="op-btn op-btnPrimary" onClick={props.onIndexMore}>
          Index more
        </button>
      </div>
    </div>
  );
}
