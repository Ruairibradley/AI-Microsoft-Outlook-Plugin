from email_processor import search_emails
from ollama_client import generate

def main():
    print("\n Outlook AI CLI – Local Email Assistant")
    print("Type 'exit' to quit.\n")

    while True:
        query = input(" Ask a question: ").strip()
        if query.lower() in {"exit", "quit"}:
            print(" Goodbye!")
            break

        print("\n Searching for relevant emails...\n")
        results = search_emails(query, n_results=3)

        if not results:
            print("No matching emails found.\n")
            continue

        # Build context
        context = "\n\n".join([f"Email: {r['filename']}\n{r['content']}" for r in results])
        prompt = f"Use the following emails to answer the question.\n\n{context}\n\nQuestion: {query}\nAnswer briefly and cite filenames."

        print(" Thinking...\n")
        answer = generate(prompt)

        print(" Answer:\n", answer)
        print("\n Sources:")
        for r in results:
            print(" -", r["filename"])
        print("\n" + "-" * 60 + "\n")

if __name__ == "__main__":
    main()
