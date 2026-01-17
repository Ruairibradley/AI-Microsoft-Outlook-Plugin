import os
from typing import Optional, List, Dict, Any

from fastapi import FastAPI, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .graph_client import list_folders, list_messages, get_messages_by_ids, list_messages_page, get_message_weblink
from .email_processor import ingest_messages, search_emails, clear_index, get_index_status, list_ingestions, clear_ingestion
from .ollama_client import generate

app = FastAPI(title="Outlook Local-Privacy Assistant Backend")

origins = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "https://localhost:8443").split(",") if o.strip()]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _require_token(authorization: Optional[str]) -> str:
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Missing Bearer token")
    return authorization.split(" ", 1)[1].strip()


class IngestRequest(BaseModel):
    folder_id: Optional[str] = None
    message_ids: List[str]
    ingestion_id: Optional[str] = None
    ingestion_label: Optional[str] = None
    ingest_mode: Optional[str] = None


class QueryRequest(BaseModel):
    question: str
    n_results: int = 4


class ClearRequest(BaseModel):
    pass


class ClearIngestionRequest(BaseModel):
    ingestion_id: str


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/index/status")
def index_status():
    return get_index_status()


@app.get("/index/ingestions")
def index_ingestions(limit: int = 50):
    return list_ingestions(limit=limit)


@app.post("/index/clear_ingestion")
def index_clear_ingestion(req: ClearIngestionRequest):
    if not req.ingestion_id:
        raise HTTPException(status_code=400, detail="ingestion_id is required")
    return clear_ingestion(req.ingestion_id)


@app.get("/graph/folders")
def graph_folders(Authorization: Optional[str] = Header(default=None)):
    token = _require_token(Authorization)
    return {"folders": list_folders(token)}


@app.get("/graph/messages")
def graph_messages(folder_id: str, top: int = 25, Authorization: Optional[str] = Header(default=None)):
    token = _require_token(Authorization)
    return list_messages(token, folder_id=folder_id, top=top)


@app.get("/graph/messages_page")
def graph_messages_page(
    folder_id: Optional[str] = None,
    top: int = 25,
    next_link: Optional[str] = None,
    Authorization: Optional[str] = Header(default=None),
):
    token = _require_token(Authorization)
    try:
        return list_messages_page(token, folder_id=folder_id, top=top, next_link=next_link)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/graph/message_link")
def graph_message_link(message_id: str, Authorization: Optional[str] = Header(default=None)):
    token = _require_token(Authorization)
    if not message_id:
        raise HTTPException(status_code=400, detail="message_id is required")
    try:
        link = get_message_weblink(token, message_id)
        return {"message_id": message_id, "weblink": link}
    except Exception as e:
        raise HTTPException(status_code=404, detail=str(e))


@app.post("/ingest")
def ingest(req: IngestRequest, Authorization: Optional[str] = Header(default=None)):
    token = _require_token(Authorization)
    if not req.message_ids:
        raise HTTPException(status_code=400, detail="message_ids is required")

    messages = get_messages_by_ids(token, req.message_ids)

    timings = ingest_messages(
        messages,
        folder_id=req.folder_id,
        ingestion_id=req.ingestion_id,
        ingestion_label=req.ingestion_label,
        ingest_mode=req.ingest_mode,
        log_timings=True
    )
    return {"ok": True, "timings": timings}


def build_prompt(question: str, results: List[Dict[str, Any]]) -> str:
    sources = []
    for i, r in enumerate(results, start=1):
        sources.append(
            f"[{i}] Subject: {r['subject']}\nFrom: {r['sender']}\nReceived: {r['received_dt']}\nLink: {r['weblink']}\n\n{r['content']}"
        )
    return (
        "Answer the question using ONLY the SOURCES below.\n"
        "If the answer is not contained in the sources, say you don't know.\n"
        "Cite sources using [1], [2], etc.\n\n"
        "SOURCES:\n" + "\n\n".join(sources) + "\n\n"
        "QUESTION:\n" + question + "\n\n"
        "ANSWER:\n"
    )


@app.post("/query")
def query(req: QueryRequest):
    results, t_retr = search_emails(req.question, n_results=req.n_results, log_timings=True)
    if not results:
        return {"answer": "I don't know based on the indexed emails.", "sources": [], "timings": t_retr}

    prompt = build_prompt(req.question, results)
    answer, t_llm = generate(prompt, max_tokens=220, log_timings=True)

    sources = [{
        "message_id": r["message_id"],
        "weblink": r["weblink"],
        "subject": r["subject"],
        "sender": r["sender"],
        "received_dt": r["received_dt"],
        "snippet": r["snippet"],
        "score": r.get("score"),
    } for r in results]

    return {"answer": answer, "sources": sources, "timings": {**t_retr, **t_llm}}


@app.post("/clear")
def clear(req: ClearRequest):
    return clear_index()


# ---- Serve frontend build output (must be mounted LAST) ----
WEB_DIR = os.path.join(os.path.dirname(__file__), "web")
if os.path.isdir(WEB_DIR):
    app.mount("/", StaticFiles(directory=WEB_DIR, html=True), name="web")
