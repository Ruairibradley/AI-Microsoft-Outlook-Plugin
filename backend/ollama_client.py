import os
import requests
import json
import time
import psutil
import threading
from timing_utils import timer

OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5:3b")


def monitor_usage(stop_event, samples):
    process = psutil.Process(os.getpid())
    while not stop_event.is_set():
        cpu = psutil.cpu_percent(interval=0.2)
        mem = process.memory_info().rss / (1024 * 1024)  # MB
        samples.append((time.time(), cpu, mem))
        time.sleep(0.2)


def generate(prompt: str, max_tokens: int = 150, log_timings=False):
    """
    Sends a prompt to the local Ollama LLM, measures inference latency,
    and monitors CPU/memory usage in parallel.
    """
    timings = {}
    samples = []
    stop_event = threading.Event()

    # Start monitoring in the background
    monitor_thread = threading.Thread(target=monitor_usage, args=(stop_event, samples))
    monitor_thread.start()

    with timer(timings, "llm_inference_ms"):
        try:
            with requests.post(
                f"{OLLAMA_HOST}/api/generate",
                json={
                    "model": OLLAMA_MODEL,
                    "prompt": prompt,
                    "stream": True,
                    "options": {
                        "num_predict": max_tokens,
                        "temperature": 0.4,
                        "top_p": 0.9
                    },
                },
                stream=True,
                timeout=120,
            ) as response:
                response.raise_for_status()

                output_text = ""

                for line in response.iter_lines():
                    if not line:
                        continue
                    try:
                        data = json.loads(line.decode("utf-8"))
                    except json.JSONDecodeError:
                        continue

                    if "response" in data:
                        token = data["response"]
                        print(token, end="", flush=True)
                        output_text += token

                    if data.get("done", False):
                        break

                print()
                answer = output_text.strip()

        except Exception as e:
            answer = f"[Error contacting Ollama: {e}]"

    # Stop monitoring thread
    stop_event.set()
    monitor_thread.join()

    # Compute CPU/memory stats
    if samples:
        avg_cpu = sum(s[1] for s in samples) / len(samples)
        peak_cpu = max(s[1] for s in samples)
        peak_mem = max(s[2] for s in samples)

        timings["avg_cpu_percent"] = round(avg_cpu, 1)
        timings["peak_cpu_percent"] = round(peak_cpu, 1)
        timings["peak_mem_mb"] = round(peak_mem, 1)

        print(f"Avg CPU: {avg_cpu:.1f}% | Peak CPU: {peak_cpu:.1f}% | Peak Memory: {peak_mem:.1f} MB")

    # Add other metadata
    timings["model"] = OLLAMA_MODEL
    timings["timestamp"] = time.strftime("%Y-%m-%dT%H:%M:%S")

    # Optional logging
    if log_timings:
        with open("latency_log.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps(timings) + "\n")

    return answer, timings
