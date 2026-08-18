"""Ollama vision-model client for transcribing rota page images."""
from pathlib import Path

from ollama import Client


def get_client(ip_server: str | None) -> Client:
    """Build an Ollama client pointed at ip_server, or the local default"""
    return Client(host=ip_server) if ip_server else Client()


def transcribe_image(client: Client, image_path: Path, model: str, prompt: str) -> str:
    """Send one image to the vision model and return its raw text response."""
    image_bytes = Path(image_path).read_bytes()

    response = client.chat(
        model=model,
        messages=[
            {
                "role": "user",
                "content": prompt,
                "images": [image_bytes],
            }
        ],
        options={"seed": 42},  # add "num_predict": 4096 here if output gets truncated
    )
    return response["message"]["content"]
