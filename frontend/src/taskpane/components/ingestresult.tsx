export function IngestResult(props: {
    kind: "CANCELLED" | "COMPLETE";
    summary: string;

    onReturnToSelect: () => void;
    onIndexMore: () => void;

    onOpenClear: () => void; // opens ClearModal
  }) {
    if (props.kind === "CANCELLED") {
      return (
        <div className="op-fit">
          <div className="op-banner op-bannerDanger">
            <div className="op-bannerTitle">Indexing cancelled</div>
            <div className="op-bannerText">{props.summary}</div>
          </div>

          <div className="op-spacer" />
          <div className="op-row">
            <button className="op-btn op-btnDanger" onClick={props.onOpenClear}>Clear…</button>
            <button className="op-btn" onClick={props.onReturnToSelect}>Return to selection</button>
          </div>
        </div>
      );
    }

    return (
      <div className="op-fit">
        <div className="op-banner op-bannerStrong">
          <div className="op-bannerTitle">Index updated</div>
          <div className="op-bannerText">{props.summary}</div>
        </div>

        <div className="op-spacer" />
        <div className="op-row">
          <button className="op-btn op-btnPrimary" onClick={props.onIndexMore}>Index more</button>
        </div>
      </div>
    );
  }
