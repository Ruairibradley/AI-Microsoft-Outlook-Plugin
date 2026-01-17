const configuredBase = (import.meta.env.VITE_BACKEND_URL as string) || "";
const backendBase = configuredBase.replace(/\/+$/, "");

async function backend_fetch(path: string, accessToken: string | null, options: RequestInit = {}) {
  const url = `${backendBase}${path}`;

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    ...(options.headers as any || {})
  };

  if (accessToken) {
    headers.Authorization = `Bearer ${accessToken}`;
  }

  const res = await fetch(url, {
    ...options,
    headers
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`backend ${res.status}: ${text}`);
  }

  return res.json();
}

// ---------- index ----------
export type IndexStatus = {
  indexed_count: number;
  last_updated: string | null;
  timestamp: string;
};

export type IngestionInfo = {
  ingestion_id: string;
  created_at: string;
  label: string;
  mode: string;
  email_count: number;
};

export function get_index_status() {
  return backend_fetch("/index/status", null);
}

export function list_ingestions(accessToken: string | null, limit: number = 50) {
  // No auth required for this local endpoint, but we accept token for consistency.
  return backend_fetch(`/index/ingestions?limit=${encodeURIComponent(String(limit))}`, accessToken);
}

export function clear_ingestion(accessToken: string | null, ingestion_id: string) {
  return backend_fetch("/index/clear_ingestion", accessToken, {
    method: "POST",
    body: JSON.stringify({ ingestion_id })
  });
}

// ---------- graph ----------
export function get_folders(accessToken: string) {
  return backend_fetch("/graph/folders", accessToken);
}

export function get_messages(accessToken: string, folder_id: string, top: number = 25) {
  const q = new URLSearchParams({ folder_id, top: String(top) }).toString();
  return backend_fetch(`/graph/messages?${q}`, accessToken);
}

export type GraphMessagesPage = {
  value?: any[];
  "@odata.nextLink"?: string;
};

export function get_messages_page(accessToken: string, args: { folder_id?: string; top?: number; next_link?: string }) {
  const params = new URLSearchParams();
  if (args.folder_id) params.set("folder_id", args.folder_id);
  if (typeof args.top === "number") params.set("top", String(args.top));
  if (args.next_link) params.set("next_link", args.next_link);
  return backend_fetch(`/graph/messages_page?${params.toString()}`, accessToken);
}

export function get_message_link(accessToken: string, message_id: string) {
  const q = new URLSearchParams({ message_id }).toString();
  return backend_fetch(`/graph/message_link?${q}`, accessToken);
}

// ---------- ingest / query ----------
export function ingest_messages(
  accessToken: string,
  folder_id: string,
  message_ids: string[],
  opts?: { ingestion_id?: string; ingestion_label?: string; ingest_mode?: string }
) {
  return backend_fetch("/ingest", accessToken, {
    method: "POST",
    body: JSON.stringify({
      folder_id,
      message_ids,
      ingestion_id: opts?.ingestion_id,
      ingestion_label: opts?.ingestion_label,
      ingest_mode: opts?.ingest_mode
    })
  });
}

export function ask_question(accessToken: string, question: string, n_results: number = 4) {
  return backend_fetch("/query", accessToken, {
    method: "POST",
    body: JSON.stringify({
      question,
      n_results
    })
  });
}

export function clear_index(accessToken: string) {
  return backend_fetch("/clear", accessToken, {
    method: "POST",
    body: JSON.stringify({})
  });
}
