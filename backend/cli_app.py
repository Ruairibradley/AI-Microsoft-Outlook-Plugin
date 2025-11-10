import time
import json
from email_processor import search_emails
from ollama_client import generate
from timing_utils import timer

def main():
    print("\nOutlook AI CLI – Local Email Assistant")
    print("Type 'exit' to quit.\n")

    while True:
        query = input(" Ask a question: ").strip()
        if query.lower() in {"exit", "quit"}:
            print(" Goodbye!")
            break

        print("\n Searching for relevant emails...\n")

        # Measure total end-to-end latency
        timings = {}
        t0 = time.perf_counter()

        # Step 1 – Retrieval (vector search)
        results, retrieval_timings = search_emails(query, n_results=3, log_timings=False)
        timings.update(retrieval_timings)

        if not results:
            print("No matching emails found.\n")
            continue

        # Step 2 – Prepare prompt
        context = "\n\n".join([f"Email: {r['filename']}\n{r['content']}" for r in results])
        prompt = (
            f"Use the following emails to answer the question.\n\n"
            f"{context}\n\n"
            f"Question: {query}\n"
            f"Answer clearly and cite the source email filenames."
        )

        print(" Thinking...\n")

        # Step 3 – LLM Generation (inference)
        answer, llm_timings = generate(prompt, log_timings=False)
        timings.update(llm_timings)

        # Step 4 – Total timing
        total_time_s = time.perf_counter() - t0
        timings["total_ms"] = round(total_time_s * 1000, 1)
        timings["query"] = query
        timings["timestamp"] = time.strftime("%Y-%m-%dT%H:%M:%S")

        # Step 5 – Output and log
        print("\n Answer:\n", answer)
        print("\n Sources:")
        for r in results:
            print(" -", r["filename"])

        print("\n--- Latency Breakdown ---")
        for k, v in timings.items():
            if k.endswith("_ms"):
                print(f"{k:20s}: {v:.1f} ms")
        print("-" * 60 + "\n")

        # Log to JSONL file
        with open("latency_log.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps(timings) + "\n")


if __name__ == "__main__":
    main()
