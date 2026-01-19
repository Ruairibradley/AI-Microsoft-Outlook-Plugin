# import json
# import time

# from backend.email_processor import search_emails, clear_index
# from backend.ollama_client import generate

# def build_prompt(question: str, results):
#     sources = []
#     for i, r in enumerate(results, start=1):
#         sources.append(
#             f"[{i}] Subject: {r['subject']}\nFrom: {r['sender']}\nReceived: {r['received_dt']}\nLink: {r['weblink']}\n\n{r['content']}"
#         )
#     return (
#         "Answer the question using ONLY the SOURCES below.\n"
#         "If the answer is not contained in the sources, say you don't know.\n"
#         "Cite sources using [1], [2], etc.\n\n"
#         "SOURCES:\n" + "\n\n".join(sources) + "\n\n"
#         "QUESTION:\n" + question + "\n\n"
#         "ANSWER:\n"
#     )

# def main():
#     print("\nOutlook AI CLI – Local Index Query Tool (Encrypted SQLite + Chroma + Ollama)")
#     print("Commands: 'clear' to delete local index, 'exit' to quit.\n")

#     passphrase = input("Passphrase (used to decrypt local index): ").strip()
#     if len(passphrase) < 8:
#         print("Passphrase must be at least 8 characters.")
#         return

#     while True:
#         q = input("\nAsk a question ('clear' / 'exit'): ").strip()
#         if q.lower() in {"exit", "quit"}:
#             print("Goodbye.")
#             break

#         if q.lower() == "clear":
#             res = clear_index()
#             print(f"Cleared: {res}")
#             continue

#         t0 = time.perf_counter()
#         print("\nSearching...\n")
#         results, t_retr = search_emails(q, passphrase=passphrase, n_results=4, log_timings=False)

#         if not results:
#             print("No results found.\n")
#             continue

#         prompt = build_prompt(q, results)

#         print("Thinking...\n")
#         answer, t_llm = generate(prompt, max_tokens=220, log_timings=False)
#         total_ms = (time.perf_counter() - t0) * 1000.0

#         print("\nAnswer:\n")
#         print(answer)

#         print("\n--- Retrieved Sources ---")
#         for i, r in enumerate(results, start=1):
#             score = r.get("score", None)
#             score_str = f"{score:.4f}" if isinstance(score, float) else "n/a"
#             print(f"[{i}] score={score_str} subject={r['subject']}")
#             print(f"    from={r['sender']} received={r['received_dt']}")
#             print(f"    link={r['weblink']}")
#             print(r["snippet"])
#             print()

#         print("\n--- Latency Breakdown ---")
#         for k, v in {**t_retr, **t_llm}.items():
#             if k.endswith("_ms"):
#                 print(f"{k:22s}: {v:.1f} ms")
#         print(f"{'total_ms':22s}: {total_ms:.1f} ms")

#         log = {
#             "event": "cli_query",
#             "query": q,
#             "total_ms": total_ms,
#             **{k: v for k, v in {**t_retr, **t_llm}.items() if isinstance(v, (int, float, str))},
#         }
#         with open("latency_log.jsonl", "a", encoding="utf-8") as f:
#             f.write(json.dumps(log) + "\n")

# if __name__ == "__main__":
#     main()
