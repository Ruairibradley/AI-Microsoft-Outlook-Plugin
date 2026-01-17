import time
import json
import numpy as np
from typing import Dict, Any, List, Optional, Tuple

from .chroma_service import get_client
from .database import (
    get_conn,
    init_db,
    set_meta,
    get_meta,
    upsert_ingestion,
    set_ingestion_count,
    list_ingestions as db_list_ingestions,
)
from .embedding_service import embed_texts
from .timing_utils import timer

COLLECTION_NAME = "emails"


def _now_iso() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


def _normalize_message(m: Dict[str, Any]) -> Dict[str, Any]:
    sender = ""
    try:
        sender = (m.get("from") or {}).get("emailAddress", {}).get("address", "") or ""
    except Exception:
        sender = ""

    subject = m.get("subject") or ""
    received_dt = m.get("receivedDateTime") or ""
    weblink = m.get("webLink") or ""
    body_preview = m.get("bodyPreview") or ""

    content = f"Subject: {subject}\nFrom: {sender}\nReceived: {received_dt}\n\n{body_preview}".strip()

    return {
        "message_id": m["id"],
        "subject": subject,
        "sender": sender,
        "received_dt": received_dt,
        "weblink": weblink,
        "content": content,
    }


def get_index_status() -> Dict[str, Any]:
    conn = get_conn()
    init_db(conn)
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) AS c FROM emails")
    indexed_count = int(cur.fetchone()["c"])

    last_updated = get_meta(conn, "last_updated")

    return {
        "indexed_count": indexed_count,
        "last_updated": last_updated,
        "timestamp": _now_iso(),
    }


def list_ingestions(limit: int = 50) -> Dict[str, Any]:
    conn = get_conn()
    init_db(conn)
    return {"ingestions": db_list_ingestions(conn, limit=limit), "timestamp": _now_iso()}


def ingest_messages(
    messages: List[Dict[str, Any]],
    folder_id: Optional[str] = None,
    ingestion_id: Optional[str] = None,
    ingestion_label: Optional[str] = None,
    ingest_mode: Optional[str] = None,
    log_timings: bool = False
) -> Dict[str, Any]:
    timings: Dict[str, Any] = {}
    conn = get_conn()
    init_db(conn)
    cur = conn.cursor()

    normalized = [_normalize_message(m) for m in messages if m.get("id")]
    if not normalized:
        return {"ingested_count": 0, "timestamp": _now_iso()}

    now = _now_iso()
    ingestion_id = ingestion_id or f"ingest_{int(time.time())}"
    ingest_mode = ingest_mode or "UNKNOWN"
    ingestion_label = ingestion_label or f"{ingest_mode} {now}"

    # Ensure ingestion record exists
    upsert_ingestion(conn, ingestion_id=ingestion_id, created_at=now, label=ingestion_label, mode=ingest_mode)

    with timer(timings, "sqlite_upsert_ms"):
        for m in normalized:
            cur.execute(
                """
                INSERT OR REPLACE INTO emails
                (message_id, folder_id, subject, sender, received_dt, weblink, content, ingestion_id, ingested_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    m["message_id"],
                    folder_id,
                    m["subject"],
                    m["sender"],
                    m["received_dt"],
                    m["weblink"],
                    m["content"],
                    ingestion_id,
                    now,
                ),
            )
        conn.commit()

    # Update embeddings incrementally for the ingested docs only
    ids = [m["message_id"] for m in normalized]
    docs = [m["content"] for m in normalized]

    with timer(timings, "embedding_ms"):
        vectors = np.asarray(embed_texts(docs), dtype=np.float32)

    client = get_client()
    collection = client.get_or_create_collection(name=COLLECTION_NAME)

    with timer(timings, "chroma_upsert_ms"):
        if hasattr(collection, "upsert"):
            collection.upsert(ids=ids, embeddings=vectors)
        else:
            # fallback: recreate collection (rare older chroma versions)
            try:
                client.delete_collection(COLLECTION_NAME)
            except Exception:
                pass
            collection = client.get_or_create_collection(name=COLLECTION_NAME)
            collection.add(ids=ids, embeddings=vectors)

    # Recompute ingestion count deterministically
    cur.execute("SELECT COUNT(*) AS c FROM emails WHERE ingestion_id = ?", (ingestion_id,))
    set_ingestion_count(conn, ingestion_id, int(cur.fetchone()["c"]))

    # Update index metadata
    set_meta(conn, "last_updated", now)

    timings["ingested_count"] = len(normalized)
    timings["ingestion_id"] = ingestion_id
    timings["timestamp"] = now

    if log_timings:
        with open("latency_log.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps({"event": "ingest", **timings}) + "\n")

    return timings


def search_emails(
    query: str,
    n_results: int = 4,
    log_timings: bool = False
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    timings: Dict[str, Any] = {}
    client = get_client()
    collection = client.get_or_create_collection(name=COLLECTION_NAME)

    with timer(timings, "embedding_query_ms"):
        qvec = np.asarray(embed_texts([query])[0], dtype=np.float32)

    with timer(timings, "vector_search_ms"):
        results = collection.query(query_embeddings=[qvec], n_results=n_results, include=["distances"])

    ids = (results.get("ids") or [[]])[0]
    dists = (results.get("distances") or [[]])[0]

    conn = get_conn()
    init_db(conn)
    cur = conn.cursor()

    out: List[Dict[str, Any]] = []
    for i, message_id in enumerate(ids):
        cur.execute(
            "SELECT message_id, subject, sender, received_dt, weblink, content FROM emails WHERE message_id = ?",
            (message_id,),
        )
        row = cur.fetchone()
        if not row:
            continue

        content = row["content"]
        out.append({
            "message_id": row["message_id"],
            "subject": row["subject"] or "",
            "sender": row["sender"] or "",
            "received_dt": row["received_dt"] or "",
            "weblink": row["weblink"] or "",
            "score": float(dists[i]) if i < len(dists) else None,
            "content": content,
            "snippet": content[:500],
        })

    timings["query"] = query
    timings["timestamp"] = _now_iso()

    if log_timings:
        with open("latency_log.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps({"event": "search", **timings}) + "\n")

    return out, timings


def clear_index() -> Dict[str, Any]:
    conn = get_conn()
    init_db(conn)
    cur = conn.cursor()

    cur.execute("DELETE FROM emails")
    cur.execute("DELETE FROM ingestions")
    conn.commit()

    client = get_client()
    try:
        client.delete_collection(COLLECTION_NAME)
    except Exception:
        pass

    set_meta(conn, "last_updated", _now_iso())
    return {"cleared": True, "timestamp": _now_iso()}


def clear_ingestion(ingestion_id: str) -> Dict[str, Any]:
    conn = get_conn()
    init_db(conn)
    cur = conn.cursor()

    # Find message ids in this ingestion for chroma delete
    cur.execute("SELECT message_id FROM emails WHERE ingestion_id = ?", (ingestion_id,))
    ids = [r["message_id"] for r in cur.fetchall()]

    # Delete from sqlite
    cur.execute("DELETE FROM emails WHERE ingestion_id = ?", (ingestion_id,))
    cur.execute("DELETE FROM ingestions WHERE ingestion_id = ?", (ingestion_id,))
    conn.commit()

    # Delete from chroma by ids (best-effort)
    client = get_client()
    collection = client.get_or_create_collection(name=COLLECTION_NAME)
    try:
        if ids and hasattr(collection, "delete"):
            collection.delete(ids=ids)
    except Exception:
        # If delete isn't supported, leave embeddings; search may return stale ids but sqlite filter will drop them.
        pass

    set_meta(conn, "last_updated", _now_iso())
    return {"cleared_ingestion_id": ingestion_id, "deleted_ids": len(ids), "timestamp": _now_iso()}
