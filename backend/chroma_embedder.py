from typing import List
from .embedding_service import embed_texts

class LocalEmbeddingFunction:
    """
    Chroma calls this with a list[str] and expects list[list[float]] back.

    Why:
    - Some chromadb versions require embeddings to be supplied via an embedding_function
      rather than passing `embeddings=` into .add()/.upsert().
    - Keeps the system CPU-only and consistent across ingest/query.
    """
    def __call__(self, texts: List[str]) -> List[List[float]]:
        return embed_texts(texts)
