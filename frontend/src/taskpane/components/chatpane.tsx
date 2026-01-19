import { useEffect, useMemo, useRef, useState } from "react";
import { ask_question, get_message_link } from "../../api/backend";
import { ChatMsg, ChatState, CHAT_TTL_MS, clear_chat_storage, load_chat_from_storage, save_chat_to_storage, type Source } from "../chat/storage";

function truncate(s: string, n: number) {
  if (!s) return "";
  return s.length <= n ? s : s.slice(0, n) + "…";
}

// Office.js may not be typed in TS; keep it safe.
function try_convert_to_outlook_desktop_url(owaUrl: string): string | null {
  try {
    const w: any = window as any;
    const OfficeObj = w?.Office;
    const mailbox = OfficeObj?.context?.mailbox;
    const fn = mailbox?.convertToLocalClientUrl;
    if (typeof fn === "function") {
      const localUrl = fn.call(mailbox, owaUrl);
      if (typeof localUrl === "string" && localUrl.length > 0) return localUrl;
    }
  } catch {
    // ignore
  }
  return null;
}

function make_id(prefix: string): string {
  const c: any = (globalThis as any).crypto;
  if (c?.randomUUID) return `${prefix}_${c.randomUUID()}`;
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function format_time(ms: number) {
  const d = new Date(ms);
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${hh}:${mm}`;
}

export function ChatPane(props: {
  token_ok: boolean;
  access_token: string;
  index_exists: boolean;
  index_count: number;
}) {
  const [chat_state, setChatState] = useState<ChatState>("EMPTY");
  const [chat_msgs, setChatMsgs] = useState<ChatMsg[]>([]);
  const [chat_expired_note, setChatExpiredNote] = useState<boolean>(false);

  const [draft, setDraft] = useState<string>("");
  const [sending, setSending] = useState<boolean>(false);

  const transcript_ref = useRef<HTMLDivElement | null>(null);

  function scroll_to_bottom() {
    const el = transcript_ref.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }

  // Restore chat when index exists
  useEffect(() => {
    if (!props.index_exists) {
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(false);
      return;
    }

    const loaded = load_chat_from_storage(make_id);
    if (!loaded) {
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(false);
      return;
    }

    const age = Date.now() - loaded.ts_ms;
    if (age <= CHAT_TTL_MS) {
      setChatMsgs(loaded.msgs);
      setChatState(loaded.msgs.length ? "RESTORED" : "EMPTY");
      setChatExpiredNote(false);
    } else {
      clear_chat_storage();
      setChatMsgs([]);
      setChatState("EMPTY");
      setChatExpiredNote(true);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.index_exists]);

  // keep transcript bottomed out on new messages
  useEffect(() => {
    scroll_to_bottom();
  }, [chat_msgs.length]);

  function append_msg(m: ChatMsg) {
    setChatMsgs((prev) => {
      const next = [...prev, m];
      save_chat_to_storage(next);
      return next;
    });
    setChatState("ACTIVE");
  }

  function patch_msg(id: string, patch: Partial<ChatMsg>) {
    setChatMsgs((prev) => {
      const next = prev.map((m) => (m.id === id ? { ...m, ...patch } : m));
      save_chat_to_storage(next);
      return next;
    });
    setChatState("ACTIVE");
  }

  function toggle_sources(id: string) {
    setChatMsgs((prev) => {
      const next = prev.map((m) => (m.id === id ? { ...m, sources_open: !m.sources_open } : m));
      save_chat_to_storage(next);
      return next;
    });
  }

  async function open_email(message_id: string, fallback_weblink: string) {
    if (!props.token_ok || !props.access_token) {
      window.open(fallback_weblink, "_blank", "noopener,noreferrer");
      return;
    }

    try {
      const res = await get_message_link(props.access_token, message_id);
      const fresh = (res.weblink || "") as string;
      const urlToUse = fresh || fallback_weblink;

      const local = try_convert_to_outlook_desktop_url(urlToUse);
      window.open(local || urlToUse, "_blank", "noopener,noreferrer");
    } catch {
      const local = try_convert_to_outlook_desktop_url(fallback_weblink);
      window.open(local || fallback_weblink, "_blank", "noopener,noreferrer");
    }
  }

  async function send() {
    if (!props.token_ok || !props.access_token) {
      alert("Please sign in first.");
      return;
    }
    if (!props.index_exists) {
      alert("Index emails first using Index management.");
      return;
    }

    const q = draft.trim();
    if (!q) return;

    setSending(true);

    const user: ChatMsg = {
      id: make_id("chat"),
      role: "user",
      text: q,
      created_at_ms: Date.now()
    };
    append_msg(user);

    setDraft("");

    const assistantId = make_id("chat");
    const placeholder: ChatMsg = {
      id: assistantId,
      role: "assistant",
      text: "Searching indexed emails…",
      created_at_ms: Date.now(),
      sources_open: false
    };
    append_msg(placeholder);

    // staged feedback
    const t1 = setTimeout(() => patch_msg(assistantId, { text: "Reading top matches…", sources: [] }), 700);
    const t2 = setTimeout(() => patch_msg(assistantId, { text: "Drafting answer…", sources: [] }), 1400);

    try {
      const res = await ask_question(props.access_token, q, 4);
      const answer = String(res.answer || "");
      const allSources = (res.sources || []) as Source[];
      const topSources = allSources.slice(0, 3);

      patch_msg(assistantId, { text: answer || "No answer.", sources: topSources, sources_open: false });
    } catch {
      patch_msg(assistantId, { text: "Query failed. Please try again.", sources: [] });
    } finally {
      clearTimeout(t1);
      clearTimeout(t2);
      setSending(false);
    }
  }

  const empty_note = useMemo(() => {
    if (chat_expired_note) return "Previous chat expired after 1 hour.";
    return "Ask a question about your indexed emails to begin.";
  }, [chat_expired_note]);

  return (
    <div className="op-chatShell">
      <div className="op-card op-fit" style={{ flex: 1, minHeight: 0 }}>
        <div className="op-cardHeader">
          <div>
            <div className="op-cardTitle">Chat</div>
            <div className="op-muted">Answers are based only on locally indexed emails.</div>
          </div>
          <button
            className="op-btn"
            onClick={() => {
              clear_chat_storage();
              setChatMsgs([]);
              setChatState("EMPTY");
              setChatExpiredNote(false);
            }}
            disabled={sending}
            title="Clear local chat history"
          >
            Clear chat
          </button>
        </div>

        <div className="op-cardBody op-fitBody" style={{ padding: 12 }}>
          {!chat_msgs.length ? (
            <div className="op-muted" style={{ marginTop: 10 }}>
              {empty_note}
              <div className="op-helpNote">
                Example: “Summarise the latest emails from my manager” or “Find invoices from last month.”
              </div>
            </div>
          ) : (
            <div className="op-chatTranscript" ref={transcript_ref}>
              {chat_state === "RESTORED" ? (
                <div className="op-banner" style={{ marginBottom: 10 }}>
                  <div className="op-bannerTitle">Chat restored</div>
                  <div className="op-bannerText">Restored messages from the last hour.</div>
                </div>
              ) : null}

              {chat_msgs.map((m) => {
                const isUser = m.role === "user";
                const srcs = (m.sources || []).slice(0, 3);
                return (
                  <div key={m.id} className={`op-bubble ${isUser ? "op-bubbleUser" : ""}`}>
                    <div className="op-bubbleHeader">
                      <div className="op-bubbleRole">{isUser ? "You" : "Assistant"}</div>
                      <div className="op-bubbleTime">{format_time(m.created_at_ms)}</div>
                    </div>

                    <div className="op-bubbleText">{m.text}</div>

                    {!isUser && srcs.length ? (
                      <>
                        <div
                          className="op-sourcesToggle"
                          onClick={() => toggle_sources(m.id)}
                          role="button"
                          tabIndex={0}
                          onKeyDown={(e) => {
                            if (e.key === "Enter" || e.key === " ") toggle_sources(m.id);
                          }}
                        >
                          Sources ({srcs.length}) {m.sources_open ? "▲" : "▼"}
                        </div>

                        {m.sources_open ? (
                          <div>
                            {srcs.map((s) => (
                              <div key={`${m.id}_${s.message_id}`} className="op-sourceRow">
                                <div className="op-sourceTop">
                                  <div style={{ minWidth: 0 }}>
                                    <div className="op-itemTitle">{truncate(s.subject || "(no subject)", 52)}</div>
                                    <div className="op-itemMeta">{truncate(s.sender, 30)} • {truncate(s.received_dt, 26)}</div>
                                  </div>
                                  <button className="op-sourceLink" onClick={() => open_email(s.message_id, s.weblink)}>
                                    Open
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        ) : null}
                      </>
                    ) : null}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="op-chatComposer">
          <textarea
            className="op-textarea"
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            rows={2}
            placeholder="Ask a question…"
            disabled={sending}
          />
          <div className="op-row" style={{ justifyContent: "space-between", marginTop: 8 }}>
            <div className="op-muted">{sending ? "Working…" : `Indexed emails: ${props.index_count}`}</div>
            <button className="op-btn op-btnPrimary" onClick={send} disabled={!draft.trim() || sending}>
              Send
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
