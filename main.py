"""
Rota transcription pipeline.

Workflow:
    - If a PDF is given, convert every page to a PNG in the hospital's image
      folder first, named "page_NNN_<hospital>.png".
    - Otherwise, process whatever page_*.png files are already sitting in
      that folder — this lets you curate exactly which pages exist (delete
      some, add more over time) without every run silently regenerating
      the full set from the source PDF and undoing that.
    - Send each image, in filename order, to the vision model along with
      the active prompt.
    - Save the raw Markdown transcription for each page via storage.py, into
      that hospital's own transcripts subfolder.

Everything that varies per run — which hospital, which prompt, which PDF,
where the markdown lands — can be set in config.py or overridden right here
on the command line, so a second (or third) hospital never requires editing
config.py itself.

Usage:
    python main.py                                  # config defaults: config.HOSPITAL, existing images
    python main.py path/to/file.pdf                  # (re)generate images from this PDF first
    python main.py --pdf path/to/file.pdf --facility Kakamega
    python main.py --facility Nakuru --prompt prompts/base2.txt --pdf pdfs/Nakuru_NBU_Rota_2023.pdf
    python main.py --facility Kakamega --out-dir transcripts/kakamega_custom
"""
import argparse
import sys
from pathlib import Path

import config
from pipeline.a_pdf_utils import pdf_to_images
from pipeline.b_llm_client import get_client, transcribe_image
from pipeline.storage import save_transcript


def load_prompt(prompt_path: Path) -> str:
    return Path(prompt_path).read_text(encoding="utf-8")


def resolve_prompt_path(prompt_arg: str | None) -> Path:
    """Accept either a path to a prompt file, or just its name (e.g. "base2"
    for prompts/base2.txt), or nothing (falls back to config.ACTIVE_PROMPT_FILE).
    """
    if not prompt_arg:
        return config.ACTIVE_PROMPT_FILE

    candidate = Path(prompt_arg)
    if candidate.exists():
        return candidate

    named = config.PROMPTS_DIR / f"{prompt_arg}.txt"
    if named.exists():
        return named

    raise FileNotFoundError(
        f"Prompt not found: tried {candidate} and {named}. "
        f"Pass a path to a prompt file, or a bare name matching a file in {config.PROMPTS_DIR}."
    )


def run(
    hospital: str,
    pdf_path: Path | None,
    prompt_path: Path,
    out_dir: Path,
) -> None:
    client = get_client(config.IP_SERVER)
    prompt = load_prompt(prompt_path)
    prompt_name = prompt_path.stem

    images_dir = config.images_dir_for(hospital)

    print(f"Facility: {hospital}")
    print(f"Model: {config.MODEL} @ {config.IP_SERVER}")
    print(f"Prompt: {prompt_path.name}")
    print(f"Images dir: {images_dir}")
    print(f"Output dir: {out_dir}")

    if pdf_path is not None:
        image_paths = pdf_to_images(Path(pdf_path), images_dir, config.DPI, hospital=hospital)
    else:
        image_paths = sorted(images_dir.glob("page_*.png"))
        if not image_paths:
            print(f"No images found in {images_dir} and no PDF given.")
            print(f"Either add page images there, or run: python main.py --pdf path/to/file.pdf --facility {hospital}")
            return
        print(f"Processing {len(image_paths)} existing image(s) in {images_dir}")

    for image_path in image_paths:
        print(f"Transcribing {image_path} ...")
        result = transcribe_image(client, image_path, config.MODEL, prompt)
        out = save_transcript(image_path, result, out_dir, config.STORAGE_BACKEND, prompt_name=prompt_name)
        print(f"Saved -> {out}")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Transcribe one hospital's rota PDF/images with the vision model.")
    parser.add_argument("pdf", type=Path, nargs="?", default=None,
                         help="PDF to (re)convert to images first (positional shortcut for --pdf).")
    parser.add_argument("--pdf", dest="pdf_flag", type=Path, default=None,
                         help="PDF to (re)convert to images first. Omit to process existing images/<hospital>/*.png as-is.")
    parser.add_argument("--facility", "--hospital", dest="hospital", default=config.HOSPITAL,
                         help=f"Hospital tag (default: {config.HOSPITAL!r}). Used for images/<hospital>/, "
                              f"transcripts/<model>/<hospital>/, and xlsx/<hospital>/ — this is also the exact "
                              f"value lookup.py's hospital argument should match.")
    parser.add_argument("--prompt", default=None,
                         help="Prompt file path, or bare name from prompts/ (e.g. 'base2'). Default: config.ACTIVE_PROMPT_FILE.")
    parser.add_argument("--out-dir", dest="out_dir", type=Path, default=None,
                         help="Where to save .md transcripts (default: config.transcripts_dir_for(hospital)).")
    return parser.parse_args(argv)


if __name__ == "__main__":
    args = parse_args(sys.argv[1:])

    pdf_arg = args.pdf_flag or args.pdf
    prompt_path = resolve_prompt_path(args.prompt)
    out_dir = args.out_dir or config.transcripts_dir_for(args.hospital)

    run(args.hospital, pdf_arg, prompt_path, out_dir)