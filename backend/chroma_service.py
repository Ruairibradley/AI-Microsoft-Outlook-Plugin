import os
import chromadb

def get_client():
    path = os.getenv("CHROMA_DIR", "./data/chroma")
    os.makedirs(path, exist_ok=True)
    return chromadb.PersistentClient(path=path)

def get_collection(name: str = "emails"):
    client = get_client()
    return client.get_or_create_collection(name=name)
