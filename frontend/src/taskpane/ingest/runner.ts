import { get_messages_page, ingest_messages, type GraphMessagesPage } from "../../api/backend";
import type { Folder, IngestMode, Phase } from "./types";

function uniq<T>(arr: T[]) {
  return Array.from(new Set(arr));
}

function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function now_iso_local() {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function make_id(prefix: string): string {
  const c: any = (globalThis as any).crypto;
  if (c?.randomUUID) return `${prefix}_${c.randomUUID()}`;
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

export type RunIngestArgs = {
  access_token: string;

  ingest_mode: IngestMode;

  // EMAILS mode
  email_folder_id: string;
  selected_email_ids: string[];

  // FOLDERS mode
  folders: Folder[];
  selected_folder_ids: string[];
  folder_limit: number;

  // callbacks (UI updates)
  onPhase: (p: Phase) => void;
  onTotal: (total: number | null) => void;
  onFetchDone: (n: number) => void;
  onIngestDone: (n: number) => void;

  // cancel/pause gate
  isCancelRequested: () => boolean;
  isPauseRequested: () => boolean;
  waitForDecision: () => Promise<"continue" | "cancel">;

  // tuning
  batch_size?: number;     // default 5
  page_size?: number;      // default 25
  min_phase_ms?: number;   // default 250
};

export type RunIngestResult = {
  run_id: string;
  run_label: string;
  message_ids: string[];
};

export async function run_ingestion(args: RunIngestArgs): Promise<RunIngestResult> {
  const PAGE_SIZE = args.page_size ?? 25;
  const BATCH_SIZE = args.batch_size ?? 5;
  const MIN_PHASE_MS = args.min_phase_ms ?? 250;

  const run_id = make_id("ing");
  const run_label = `${args.ingest_mode} ${now_iso_local()}`;

  args.onPhase("FETCHING");

  // Build ids
  let message_ids: string[] = [];

  if (args.ingest_mode === "EMAILS") {
    message_ids = [...args.selected_email_ids];
    args.onTotal(message_ids.length);
    args.onFetchDone(message_ids.length);
  } else {
    const selected = new Set(args.selected_folder_ids);
    const chosenFolders = args.folders.filter((f) => selected.has(f.id));
    const limit = Math.max(1, args.folder_limit);

    // Approx total for progress denominator
    const approxTotal = chosenFolders.reduce((sum, f) => sum + Math.min(limit, f.totalItemCount || limit), 0);
    args.onTotal(approxTotal || null);

    const all_ids: string[] = [];
    let done = 0;

    for (const f of chosenFolders) {
      let next_link: string | null = null;
      let collected = 0;

      while (collected < limit) {
        if (args.isPauseRequested()) {
          const decision = await args.waitForDecision();
          if (decision === "cancel") throw new Error("CANCELLED_BY_USER");
        }
        if (args.isCancelRequested()) throw new Error("CANCELLED_BY_USER");

        let pageData: GraphMessagesPage;
        if (next_link) pageData = await get_messages_page(args.access_token, { next_link, top: PAGE_SIZE });
        else pageData = await get_messages_page(args.access_token, { folder_id: f.id, top: PAGE_SIZE });

        const page = (pageData.value || []) as Array<{ id: string }>;
        next_link = (pageData as any)["@odata.nextLink"] || null;

        if (!page.length) break;

        for (const m of page) {
          if (collected >= limit) break;
          all_ids.push(m.id);
          collected += 1;
          done += 1;
          args.onFetchDone(done);
        }

        if (!next_link) break;
      }
    }

    message_ids = uniq(all_ids);
  }

  if (!message_ids.length) throw new Error("No emails selected.");

  // Minimum fetch phase
  await sleep(MIN_PHASE_MS);

  // Ingest in batches
  args.onPhase("STORING");

  const folder_id_to_send =
    args.ingest_mode === "EMAILS"
      ? (args.email_folder_id || "selected")
      : "multi";

  const batches = chunk(message_ids, BATCH_SIZE);
  let ingested = 0;

  for (const b of batches) {
    if (args.isPauseRequested()) {
      const decision = await args.waitForDecision();
      if (decision === "cancel") throw new Error("CANCELLED_BY_USER");
    }
    if (args.isCancelRequested()) throw new Error("CANCELLED_BY_USER");

    await ingest_messages(args.access_token, folder_id_to_send, b, {
      ingestion_id: run_id,
      ingestion_label: run_label,
      ingest_mode: args.ingest_mode
    });

    ingested += b.length;
    args.onIngestDone(ingested);
  }

  // Minimum store phase
  await sleep(MIN_PHASE_MS);

  // Indexing (visual phase only; backend already did it during ingest calls)
  args.onPhase("INDEXING");
  await sleep(MIN_PHASE_MS);

  args.onPhase("DONE");

  return { run_id, run_label, message_ids };
}
