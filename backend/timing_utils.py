import time
from contextlib import contextmanager
from typing import Dict, Any

@contextmanager
def timer(record: Dict[str, Any], key: str):
    start = time.perf_counter()
    try:
        yield
    finally:
        record[key] = (time.perf_counter() - start) * 1000.0
