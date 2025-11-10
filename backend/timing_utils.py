import time
from contextlib import contextmanager

@contextmanager
def timer(record, key):
    start = time.perf_counter()
    yield
    record[key] = (time.perf_counter() - start) * 1000