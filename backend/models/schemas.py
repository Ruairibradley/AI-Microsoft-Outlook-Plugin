from pydantic import BaseModel
from typing import List, Optional

class Source(BaseModel):
    message_id: str
    weblink: str
    subject: str
    sender: str
    received_dt: str
    snippet: str
    score: Optional[float] = None

class Answer(BaseModel):
    answer: str
    sources: List[Source] = []
