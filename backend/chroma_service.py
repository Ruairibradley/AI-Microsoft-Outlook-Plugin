import os
import chromadb

def get_client():
    db_path = os.getenv("CHROMA_DIR", "./data/chroma")
    os.makedirs(db_path, exist_ok=True)
    return chromadb.PersistentClient(path=db_path)
