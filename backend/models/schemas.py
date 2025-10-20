from pydantic import BaseModel
from typing import List

class Source(BaseModel):
    id: str
    snippet: str
    score: float

class Answer(BaseModel):
    answer: str
    sources: List[Source] = []
