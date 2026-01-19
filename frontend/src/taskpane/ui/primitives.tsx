import React from "react";

export function cx(...parts: Array<string | false | null | undefined>) {
  return parts.filter(Boolean).join(" ");
}

export function Card(props: { title?: string; subtitle?: string; right?: React.ReactNode; children: React.ReactNode }) {
  return (
    <div className="op-card op-fit">
      {(props.title || props.subtitle || props.right) ? (
        <div className="op-cardHeader">
          <div>
            {props.title ? <div className="op-cardTitle">{props.title}</div> : null}
            {props.subtitle ? <div className="op-muted">{props.subtitle}</div> : null}
          </div>
          {props.right ? <div>{props.right}</div> : null}
        </div>
      ) : null}
      <div className="op-cardBody">{props.children}</div>
    </div>
  );
}

export function Banner(props: { variant?: "default" | "strong" | "warn" | "danger"; title: string; children?: React.ReactNode }) {
  const v = props.variant || "default";
  return (
    <div
      className={cx(
        "op-banner",
        v === "strong" && "op-bannerStrong",
        v === "warn" && "op-bannerWarn",
        v === "danger" && "op-bannerDanger"
      )}
    >
      <div className="op-bannerTitle">{props.title}</div>
      {props.children ? <div className="op-bannerText">{props.children}</div> : null}
    </div>
  );
}

export function Button(props: React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: "default" | "primary" | "danger" | "ghost" }) {
  const v = props.variant || "default";
  return (
    <button
      {...props}
      className={cx(
        "op-btn",
        v === "primary" && "op-btnPrimary",
        v === "danger" && "op-btnDanger",
        v === "ghost" && "op-btnGhost",
        props.className
      )}
    />
  );
}

export function Input(props: React.InputHTMLAttributes<HTMLInputElement>) {
  return <input {...props} className={cx("op-input", props.className)} />;
}

export function Select(props: React.SelectHTMLAttributes<HTMLSelectElement>) {
  return <select {...props} className={cx("op-select", props.className)} />;
}

export function Textarea(props: React.TextareaHTMLAttributes<HTMLTextAreaElement>) {
  return <textarea {...props} className={cx("op-textarea", props.className)} />;
}

export function Label(props: { children: React.ReactNode }) {
  return <div className="op-label">{props.children}</div>;
}

export function Spinner() {
  return <div className="op-spinner" aria-hidden="true" />;
}

export function ProgressBar(props: { percent: number }) {
  const pct = Math.max(0, Math.min(100, Math.round(props.percent)));
  return (
    <div className="op-progressBar">
      <div className="op-progressFill" style={{ width: `${pct}%` }} />
    </div>
  );
}
