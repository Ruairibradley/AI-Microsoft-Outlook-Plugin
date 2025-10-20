# Outlook AI Plugin (Local & Privacy-First)

An Outlook add-in with a React + Office.js frontend and a Python FastAPI backend. Uses Ollama (Mistral) locally for LLM, Chroma for embeddings, and SQLite for structured data.

## Tech Stack
- Frontend: React, Office.js, Fluent UI
- Backend: Python, FastAPI
- AI: Ollama (Mistral)
- Vector DB: Chroma
- DB: SQLite

## Folders
- `frontend/` Outlook add-in (React + Office.js)
- `backend/` FastAPI service, ingestion, retrieval, model calls
- `data/` local databases (SQLite + Chroma)
- `docs/` architecture notes, roadmap

## Getting Started
1) Frontend: build the React Outlook add-in (see `frontend/README.md`).
2) Backend: copy `.env.example` to `.env`, then install Python deps listed in `requirements.txt`.
3) Sideload `frontend/manifest.xml` into Outlook (My Add-ins → Add from file).

> This repo is designed to run **fully local** for privacy.
