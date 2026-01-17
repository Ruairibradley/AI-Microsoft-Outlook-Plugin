const configuredBase = (import.meta.env.VITE_BACKEND_URL as string) || "";
const backendBase = configuredBase.replace(/\/+$/, "");

async function backend_fetch(path: string, accessToken: string | null, options: RequestInit = {}) {
  const url = `${backendBase}${path}`;

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    ...(options.headers as any || {})
  };

  // Some endpoints do not require auth; most do.
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

export function get_index_status() {
  // index status is local-only; no auth required
  return backend_fetch("/index/status", null);
}

// ---------- graph ----------
export function get_folders(accessToken: string) {
  return backend_fetch("/graph/folders", accessToken);
}

export function get_messages(accessToken: string, folder_id: string, top: number = 25) {
  const q = new URLSearchParams({ folder_id, top: String(top) }).toString();
  return backend_fetch(`/graph/messages?${q}`, accessToken);
}

// ---------- ingest / query ----------
export function ingest_messages(accessToken: string, folder_id: string, message_ids: string[]) {
  return backend_fetch("/ingest", accessToken, {
    method: "POST",
    body: JSON.stringify({
      folder_id,
      message_ids
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
