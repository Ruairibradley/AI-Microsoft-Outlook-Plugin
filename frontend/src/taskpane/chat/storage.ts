export type ChatState = "EMPTY" | "ACTIVE" | "RESTORED";

export type Source = {
  message_id: string;
  weblink: string;
  subject: string;
  sender: string;
  received_dt: string;
  snippet: string;
  score?: number;
};

export type ChatMsg = {
  id: string;
  role: "user" | "assistant";
  text: string;
  created_at_ms: number;
  sources?: Source[];
  sources_open?: boolean; // UI-only
};

export const CHAT_TTL_MS = 60 * 60 * 1000;
export const LS_CHAT_KEY = "chat_history";
export const LS_CHAT_TS_KEY = "chat_history_ts";

export function load_chat_from_storage(make_id: (prefix: string) => string): { msgs: ChatMsg[]; ts_ms: number } | null {
  try {
    const tsRaw = localStorage.getItem(LS_CHAT_TS_KEY);
    const dataRaw = localStorage.getItem(LS_CHAT_KEY);
    if (!tsRaw || !dataRaw) return null;

    const ts_ms = Number(tsRaw);
    if (!Number.isFinite(ts_ms) || ts_ms <= 0) return null;

    const parsed = JSON.parse(dataRaw);
    if (!Array.isArray(parsed)) return null;

    const msgs: ChatMsg[] = parsed
      .filter((m: any) => m && typeof m === "object")
      .map((m: any) => ({
        id: String(m.id || make_id("msg")),
        role: m.role === "assistant" ? "assistant" : "user",
        text: String(m.text || ""),
        created_at_ms: Number(m.created_at_ms || ts_ms),
        sources: Array.isArray(m.sources) ? m.sources : undefined,
        sources_open: false
      }));

    return { msgs, ts_ms };
  } catch {
    return null;
  }
}

export function save_chat_to_storage(msgs: ChatMsg[]) {
  try {
    // store without UI-only fields
    const trimmed = msgs.map((m) => ({
      id: m.id,
      role: m.role,
      text: m.text,
      created_at_ms: m.created_at_ms,
      sources: m.sources
    }));
    localStorage.setItem(LS_CHAT_KEY, JSON.stringify(trimmed));
    localStorage.setItem(LS_CHAT_TS_KEY, String(Date.now()));
  } catch {
    // ignore
  }
}

export function clear_chat_storage() {
  try {
    localStorage.removeItem(LS_CHAT_KEY);
    localStorage.removeItem(LS_CHAT_TS_KEY);
  } catch {
    // ignore
  }
}
