import os
from typing import List
from sentence_transformers import SentenceTransformer

_MODEL_NAME = os.getenv("EMBED_MODEL", "sentence-transformers/all-MiniLM-L6-v2")
_model = None

def get_embedder() -> SentenceTransformer:
    global _model
    if _model is None:
        _model = SentenceTransformer(_MODEL_NAME, device="cpu")
    return _model

def embed_texts(texts: List[str]) -> List[List[float]]:
    model = get_embedder()
    vecs = model.encode(texts, normalize_embeddings=True)
    return vecs.tolist()
