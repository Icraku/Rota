"""
Rota transcription pipeline.

Workflow:
    1. Convert each page of a PDF into a PNG image (saved to images/).
    2. Send each image, first page to last, to a vision-capable Ollama model
       along with the active prompt (prompts/current.txt).
    3. Save the raw Markdown transcription for each page via storage.py
       (transcripts/ today; SurrealDB once wired up).

Usage:
    python main.py                      # uses config.PDF_PATH
    python main.py path/to/other.pdf    # override for a different facility/run
"""
import sys
from pathlib import Path

import config
from pdf_utils import pdf_to_images
from llm_client import get_client, transcribe_image
from storage import save_transcript


def load_prompt(prompt_path: Path) -> str:
    return Path(prompt_path).read_text(encoding="utf-8")


def run(pdf_path: Path = config.PDF_PATH) -> None:
    client = get_client(config.IP_SERVER)
    prompt = load_prompt(config.ACTIVE_PROMPT_FILE)

    print(f"Facility: {config.FACILITY}")
    print(f"Model: {config.MODEL} @ {config.IP_SERVER}")
    print(f"Prompt: {config.ACTIVE_PROMPT_FILE.name}")

    image_paths = pdf_to_images(Path(pdf_path), config.IMAGES_DIR, config.DPI)

    for image_path in image_paths:
        print(f"Transcribing {image_path} ...")
        result = transcribe_image(client, image_path, config.MODEL, prompt)
        out = save_transcript(image_path, result, config.TRANSCRIPTS_DIR, config.STORAGE_BACKEND)
        print(f"Saved -> {out}")


if __name__ == "__main__":
    pdf_arg = Path(sys.argv[1]) if len(sys.argv) > 1 else config.PDF_PATH
    run(pdf_arg)
