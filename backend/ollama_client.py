import os
import requests
import json
import time
import psutil
import threading
from typing import Dict, Any, Tuple

from .timing_utils import timer

OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5:3b")

def monitor_usage(stop_event, samples):
    process = psutil.Process(os.getpid())
    while not stop_event.is_set():
        cpu = psutil.cpu_percent(interval=0.2)
        mem = process.memory_info().rss / (1024 * 1024)
        samples.append((time.time(), cpu, mem))
        time.sleep(0.2)

def generate(prompt: str, max_tokens: int = 180, log_timings: bool = False) -> Tuple[str, Dict[str, Any]]:
    timings: Dict[str, Any] = {}
    samples = []
    stop_event = threading.Event()
    t = threading.Thread(target=monitor_usage, args=(stop_event, samples))
    t.start()

    output = ""

    with timer(timings, "llm_inference_ms"):
        try:
            with requests.post(
                f"{OLLAMA_HOST}/api/generate",
                json={
                    "model": OLLAMA_MODEL,
                    "prompt": prompt,
                    "stream": True,
                    "options": {"num_predict": max_tokens, "temperature": 0.4, "top_p": 0.9},
                },
                stream=True,
                timeout=120,
            ) as r:
                r.raise_for_status()
                for line in r.iter_lines():
                    if not line:
                        continue
                    try:
                        data = json.loads(line.decode("utf-8"))
                    except json.JSONDecodeError:
                        continue
                    if "response" in data:
                        output += data["response"]
                    if data.get("done", False):
                        break
        except Exception as e:
            output = f"[Error contacting Ollama: {e}]"

    stop_event.set()
    t.join()

    if samples:
        avg_cpu = sum(s[1] for s in samples) / len(samples)
        peak_cpu = max(s[1] for s in samples)
        peak_mem = max(s[2] for s in samples)
        timings["avg_cpu_percent"] = round(avg_cpu, 1)
        timings["peak_cpu_percent"] = round(peak_cpu, 1)
        timings["peak_mem_mb"] = round(peak_mem, 1)

    timings["model"] = OLLAMA_MODEL
    timings["timestamp"] = time.strftime("%Y-%m-%dT%H:%M:%S")

    if log_timings:
        with open("latency_log.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps({"event": "llm", **timings}) + "\n")

    return output.strip(), timings
