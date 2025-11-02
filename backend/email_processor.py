import os
from chroma_service import get_client
from database import get_conn
from chromadb.utils import embedding_functions


def ingest_emails(email_dir: str = "./data/test_emails"):
    chroma = get_client()
    conn = get_conn()
    cursor = conn.cursor()

    # Create SQLite table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            content TEXT
        )
    """)

    # Add a default embedding function - later will be sentenceTransformmers.
    embedding_fn = embedding_functions.DefaultEmbeddingFunction()

    # Create or load collection
    collection = chroma.get_or_create_collection(name="emails")

    # Loop and ingest files
    count = 0
    for filename in os.listdir(email_dir):
        if not filename.endswith(".txt"):
            continue

        file_path = os.path.join(email_dir, filename)
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read().strip()

        if not content:
            print(f" Skipped empty file: {filename}")
            continue

        cursor.execute("INSERT INTO emails (filename, content) VALUES (?, ?)", (filename, content))
        email_id = cursor.lastrowid

        # Add embed to Chroma
        try:
            collection.add(
                documents=[content],
                metadatas=[{"filename": filename}],
                ids=[str(email_id)]
            )
            count += 1
            print(f" Ingested: {filename}")
        except Exception as e:
            print(f" Failed to embed {filename}: {e}")

    conn.commit()
    print(f" Ingestion complete. ({count} files embedded)")

    #  show if items exist in chroma
    print(f" Collection '{collection.name}' now contains {collection.count()} items.")


def search_emails(query: str, n_results: int = 3):
    chroma = get_client()
    embedding_fn = embedding_functions.DefaultEmbeddingFunction()

    collection = chroma.get_or_create_collection(name="emails")


    try:
        results = collection.query(query_texts=[query], n_results=n_results)
    except Exception as e:
        print(f" Query failed: {e}")
        return []

    # Defensive unpack
    if not results:
        print(" No results object returned.")
        return []

    ids = (results.get("ids") or [[]])[0]
    docs_list = (results.get("documents") or [[]])[0]
    metas_list = (results.get("metadatas") or [[]])[0]

    if not ids:
        print(" No results found in Chroma for this query.")
        return []

    docs = []
    for i in range(len(ids)):
        filename = "unknown"
        if metas_list and isinstance(metas_list, list) and i < len(metas_list):
            meta_item = metas_list[i] or {}
            filename = meta_item.get("filename", "unknown")

        content = ""
        if docs_list and isinstance(docs_list, list) and i < len(docs_list):
            content = docs_list[i] or ""

        docs.append({
            "id": ids[i],
            "filename": filename,
            "content": content
        })

    print(f" Found {len(docs)} result(s) for query: '{query}'")
    return docs


if __name__ == "__main__":
    ingest_emails(email_dir="./data/outlook_emails")


