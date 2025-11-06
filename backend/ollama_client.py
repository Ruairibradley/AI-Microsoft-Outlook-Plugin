import os
import requests
import json

OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5:3b")


def generate(prompt: str, max_tokens: int = 150) -> str:
    try:
        with requests.post(
            f"{OLLAMA_HOST}/api/generate",
            json={
                "model": OLLAMA_MODEL,
                "prompt": prompt,
                "stream": True,
                "options": {
                "num_predict": 150,
                "temperature": 0.4,
                "top_p": 0.9
            }
            # limit generation
            },
            stream=True,
            timeout=120,
        ) as response:
            response.raise_for_status()

            output_text = ""

            # make text appear line by line rather than all at once - like how llms do
            for line in response.iter_lines():
                if not line:
                    continue


                try:
                    data = json.loads(line.decode("utf-8"))
                except json.JSONDecodeError:
                    continue

                # Print each token as it arrives
                if "response" in data:
                    token = data["response"]
                    print(token, end="", flush=True)
                    output_text += token

                # Stop if Ollama says generation is done
                if data.get("done", False):
                    break

            print()  # newline after streaming output
            return output_text.strip()

    except Exception as e:
        return f"[Error contacting Ollama: {e}]"

