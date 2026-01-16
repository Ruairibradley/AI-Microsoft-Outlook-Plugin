const configuredBase = (import.meta.env.VITE_BACKEND_URL as string) || "";
const backendBase = configuredBase.replace(/\/+$/, "");

async function backend_fetch(path: string, accessToken: string, options: RequestInit = {}) {
  const url = `${backendBase}${path}`;

  const res = await fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    }
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`backend ${res.status}: ${text}`);
  }

  return res.json();
}

export function get_folders(accessToken: string) {
  return backend_fetch("/graph/folders", accessToken);
}

export function get_messages(accessToken: string, folder_id: string, top: number = 25) {
  const q = new URLSearchParams({ folder_id, top: String(top) }).toString();
  return backend_fetch(`/graph/messages?${q}`, accessToken);
}

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
