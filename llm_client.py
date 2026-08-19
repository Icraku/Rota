"""Ollama vision-model client for transcribing rota page images."""
import time
from pathlib import Path

from ollama import Client


def get_client(ip_server: str | None) -> Client:
    """Build an Ollama client pointed at ip_server, or the local default if unset."""
    return Client(host=ip_server) if ip_server else Client()


def transcribe_image(client: Client, image_path: Path, model: str, prompt: str, show_progress: bool = True) -> str:
    """Send one image to the vision model and return its raw text response.

    Streams the response instead of waiting for one big blocking reply.
    With stream=False, Ollama sends nothing back at all until generation is
    100% complete — a genuinely slow-but-working call and a silently dead
    connection look identical (total silence). Streaming makes the
    difference visible: if tokens are arriving, it's working.
    """
    image_bytes = Path(image_path).read_bytes()

    stream = client.chat(
        model=model,
        messages=[
            {
                "role": "user",
                "content": prompt,
                "images": [image_bytes],
            }
        ],
        options={"seed": 42},  # add "num_predict": 4096 here if output gets truncated
        stream=True,
    )

    chunks = []
    token_count = 0
    start = time.time()
    first_token_time = None

    for part in stream:
        piece = part["message"]["content"]
        if piece and first_token_time is None:
            first_token_time = time.time()
            if show_progress:
                print(f"  [first token after {first_token_time - start:.1f}s]")
        chunks.append(piece)
        token_count += 1
        if show_progress and token_count % 20 == 0:
            elapsed = time.time() - start
            print(f"  ... {token_count} chunks received ({elapsed:.0f}s elapsed)", end="\r")

    if show_progress:
        print()  # clear the progress line
        if first_token_time is None:
            print("  [WARNING: stream closed with zero tokens received]")

    return "".join(chunks)