"""
Rota transcription pipeline.

Workflow:
    - If a PDF is given, convert every page to a PNG in images/ first.
    - Otherwise, process whatever page_*.png files are already sitting in
      images/ — this lets you curate exactly which pages exist (delete
      some, add more over time) without every run silently regenerating
      the full set from the source PDF and undoing that.
    - Send each image, in filename order, to the vision model along with
      the active prompt (prompts/current.txt or whichever config points at).
    - Save the raw Markdown transcription for each page via storage.py.

Usage:
    python main.py                      # process existing images/*.png as-is
    python main.py path/to/file.pdf     # (re)generate images/ from this PDF first
"""
import sys
from pathlib import Path

import config
from pipeline.a_pdf_utils import pdf_to_images
from pipeline.b_llm_client import get_client, transcribe_image
from pipeline.storage import save_transcript


def load_prompt(prompt_path: Path) -> str:
    return Path(prompt_path).read_text(encoding="utf-8")


def run(pdf_path: Path | None = None) -> None:
    client = get_client(config.IP_SERVER)
    prompt = load_prompt(config.ACTIVE_PROMPT_FILE)

    print(f"Facility: {config.FACILITY}")
    print(f"Model: {config.MODEL} @ {config.IP_SERVER}")
    print(f"Prompt: {config.ACTIVE_PROMPT_FILE.name}")

    if pdf_path is not None:
        image_paths = pdf_to_images(Path(pdf_path), config.IMAGES_DIR, config.DPI)
    else:
        image_paths = sorted(config.IMAGES_DIR.glob("page_*.png"))
        if not image_paths:
            print(f"No images found in {config.IMAGES_DIR} and no PDF given.")
            print(f"Either add page images there, or run: python main.py path/to/file.pdf")
            return
        print(f"Processing {len(image_paths)} existing image(s) in {config.IMAGES_DIR}")

    for image_path in image_paths:
        print(f"Transcribing {image_path} ...")
        result = transcribe_image(client, image_path, config.MODEL, prompt)
        out = save_transcript(image_path, result, config.TRANSCRIPTS_DIR, config.STORAGE_BACKEND)
        print(f"Saved -> {out}")


if __name__ == "__main__":
    pdf_arg = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    run(pdf_arg)