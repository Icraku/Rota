"""
Compare transcription output across multiple prompts and models, side by side.

Does NOT modify main.py / llm_client.py / config.py — imports and reuses
them exactly as they are. Reuses md_to_xlsx.py's conversion logic too,
rather than duplicating a parser here.

For each image, runs every (model, prompt) combination and produces:
  - one .md file per combination in comparisons/
  - one .xlsx file per successful combination in comparisons/
  - one comparisons/<image>__comparison.html report with a column per
    combination, linking to the .xlsx file instead of dumping raw markdown

RESUME: before running a combination, if its .md output already exists and
doesn't start with "ERROR:", it's skipped and the existing result is reused
in the report. This matters a lot once you're sweeping multiple images —
one interruption doesn't mean starting the whole thing over. Use --force to
ignore existing outputs and redo everything.

Each call runs on its own daemon thread with a timeout, so if one model
hangs, it won't block the rest of the sweep.

Usage:
    python compare_prompts.py                                   # one image (first in images/)
    python compare_prompts.py --all-images                      # every .png in images/
    python compare_prompts.py --all-images --images-dir images_gemma
    python compare_prompts.py --models qwen3.6:35b gemma4:31b --prompts current base2
    python compare_prompts.py --force                            # ignore existing outputs, redo everything
"""
import argparse
import html
import threading
import time
from pathlib import Path

import openpyxl

import config
from llm_client import get_client, transcribe_image
from md_to_xlsx import convert_transcript, sheet_name_for

DEFAULT_MODELS = ["qwen3.6:35b", "qwen3.5:35b", "gemma4:31b"]
DEFAULT_PROMPTS = ["current", "currenta", "base", "base2"]
DEFAULT_TIMEOUT = 600  # seconds per (model, prompt) call
OUTPUT_DIR = config.BASE_DIR / "comparisons"


def load_prompt(name: str) -> str:
    path = config.PROMPTS_DIR / f"{name}.txt"
    if not path.exists():
        raise FileNotFoundError(f"No prompt file at {path}")
    return path.read_text(encoding="utf-8")


def pick_default_image() -> Path:
    images = sorted(config.IMAGES_DIR.glob("*.png"))
    if not images:
        raise FileNotFoundError(f"No .png files found in {config.IMAGES_DIR} — pass --image explicitly.")
    return images[0]


def output_path_for(image_path: Path, model: str, prompt_name: str) -> Path:
    safe_model = model.replace(":", "_").replace("/", "_")
    return OUTPUT_DIR / f"{image_path.stem}__{safe_model}__{prompt_name}.md"


def already_succeeded(out_path: Path) -> bool:
    if not out_path.exists():
        return False
    content = out_path.read_text(encoding="utf-8")
    return bool(content) and not content.startswith("ERROR:")


def call_with_timeout(fn, args, timeout):
    """Run fn(*args) on a daemon thread, bounded by timeout.

    Using a daemon thread (rather than e.g. a plain ThreadPoolExecutor)
    means an abandoned, still-hung call won't block this script from
    exiting once everything else is done — it just gets left behind.
    """
    result = {}

    def target():
        try:
            result["value"] = fn(*args)
        except Exception as e:  # noqa: BLE001 — deliberately broad, this is a diagnostic sweep
            result["error"] = str(e)

    t = threading.Thread(target=target, daemon=True)
    t.start()
    t.join(timeout=timeout)

    if t.is_alive():
        return None, f"timed out after {timeout}s — no response (call abandoned, still running in background)"
    if "error" in result:
        return None, result["error"]
    return result.get("value", ""), None


def build_xlsx_for_combo(md_path: Path, xlsx_path: Path) -> Path | None:
    """Convert one combination's saved markdown into its own .xlsx file.

    Reuses md_to_xlsx.py's convert_transcript() rather than re-parsing
    markdown here. Returns None (rather than raising) if conversion fails,
    so one malformed output doesn't take down the whole comparison sweep.
    """
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name_for(md_path)
        convert_transcript(md_path, ws)
        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(xlsx_path)
        return xlsx_path
    except Exception as e:
        print(f"  [WARNING] xlsx conversion failed for {md_path.name}: {e}")
        return None


def run_comparison(image_path: Path, models: list[str], prompt_names: list[str], timeout: int, force: bool) -> list[dict]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    results = []

    for model in models:
        client = None  # created lazily — skip entirely if every prompt for this model is already done
        for prompt_name in prompt_names:
            label = f"{model} / {prompt_name}"
            out_path = output_path_for(image_path, model, prompt_name)
            xlsx_out = OUTPUT_DIR / f"{image_path.stem}__{model.replace(':', '_').replace('/', '_')}__{prompt_name}.xlsx"

            if not force and already_succeeded(out_path):
                print(f"Skipping {label} — already done -> {out_path.name}")
                results.append({
                    "model": model,
                    "prompt": prompt_name,
                    "output": out_path.read_text(encoding="utf-8"),
                    "error": None,
                    "elapsed": None,
                    "path": out_path,
                    "xlsx_path": xlsx_out if xlsx_out.exists() else build_xlsx_for_combo(out_path, xlsx_out),
                })
                continue

            if client is None:
                client = get_client(config.IP_SERVER)

            prompt_text = load_prompt(prompt_name)
            print(f"Running {label} on {image_path.name} ...")

            start = time.time()
            output, error = call_with_timeout(
                transcribe_image, (client, image_path, model, prompt_text), timeout
            )
            elapsed = time.time() - start

            if error:
                print(f"  [ERROR] {label}: {error}")
            else:
                print(f"  OK ({elapsed:.1f}s, {len(output)} chars)")

            out_path.write_text(error and f"ERROR: {error}" or output, encoding="utf-8")

            xlsx_path = None
            if not error and output:
                xlsx_path = build_xlsx_for_combo(out_path, xlsx_out)
                if xlsx_path:
                    print(f"  -> {xlsx_path.name}")

            results.append({
                "model": model,
                "prompt": prompt_name,
                "output": output,
                "error": error,
                "elapsed": elapsed,
                "path": out_path,
                "xlsx_path": xlsx_path,
            })
    return results


def render_col(r: dict) -> str:
    elapsed_label = f"{r['elapsed']:.1f}s" if r["elapsed"] is not None else "cached"
    header = f"""<h3>{html.escape(r['model'])}<br><span class="prompt">{html.escape(r['prompt'])}</span></h3>
      <div class="meta">{elapsed_label}{' — ERROR' if r['error'] else ''}</div>"""

    if r["error"]:
        body = f"<pre>{html.escape(r['error'])}</pre>"
    elif r["xlsx_path"]:
        fname = html.escape(r["xlsx_path"].name)
        body = f'<a class="xlsx-link" href="{fname}">&#128202; Open {fname}</a>'
    else:
        body = "<pre>(xlsx conversion failed — check the terminal output, .md file still saved)</pre>"

    cls = "col error" if r["error"] else "col"
    return f'<div class="{cls}">{header}{body}</div>'


def build_html_report(image_path: Path, results: list[dict]) -> Path:
    cols = "".join(render_col(r) for r in results)

    doc = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Comparison — {html.escape(image_path.name)}</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 20px; background: #fafafa; }}
  h1 {{ font-size: 18px; }}
  .grid {{ display: flex; gap: 16px; overflow-x: auto; align-items: flex-start; flex-wrap: wrap; }}
  .col {{ flex: 0 0 260px; background: white; border: 1px solid #ddd; border-radius: 6px; padding: 12px; }}
  .col.error {{ border-color: #e33; background: #fff5f5; }}
  .col h3 {{ margin: 0 0 4px 0; font-size: 14px; }}
  .col .prompt {{ font-weight: normal; color: #666; font-size: 12px; }}
  .meta {{ color: #888; font-size: 11px; margin-bottom: 10px; }}
  pre {{ white-space: pre-wrap; word-wrap: break-word; font-size: 12px; max-height: 60vh; overflow-y: auto; }}
  .xlsx-link {{
    display: block; text-align: center; padding: 14px 8px;
    background: #1d6f42; color: white; border-radius: 4px;
    text-decoration: none; font-size: 13px; font-weight: bold;
  }}
  .xlsx-link:hover {{ background: #155a34; }}
</style>
</head>
<body>
<h1>Prompt / model comparison — {html.escape(image_path.name)}</h1>
<div class="grid">{cols}
</div>
</body>
</html>"""

    out_path = OUTPUT_DIR / f"{image_path.stem}__comparison.html"
    out_path.write_text(doc, encoding="utf-8")
    return out_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compare prompt/model combinations side by side.")
    parser.add_argument("--image", type=Path, default=None, help="Single image to compare (defaults to first .png in images/)")
    parser.add_argument("--all-images", action="store_true", help="Run the full sweep on every .png in --images-dir instead of just one")
    parser.add_argument("--images-dir", type=Path, default=None, help="Folder to glob .png files from with --all-images (defaults to config.IMAGES_DIR)")
    parser.add_argument("--models", nargs="+", default=DEFAULT_MODELS, help="Model tags to compare")
    parser.add_argument("--prompts", nargs="+", default=DEFAULT_PROMPTS, help="Prompt names (without .txt) from prompts/")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="Seconds to wait per call before giving up on it")
    parser.add_argument("--force", action="store_true", help="Ignore existing outputs and redo every combination")
    args = parser.parse_args()

    if args.all_images:
        images_dir = args.images_dir or config.IMAGES_DIR
        image_paths = sorted(images_dir.glob("*.png"))
        if not image_paths:
            raise FileNotFoundError(f"No .png files found in {images_dir}")
    else:
        image_paths = [args.image or pick_default_image()]

    total_calls = len(image_paths) * len(args.models) * len(args.prompts)
    print(f"{len(image_paths)} image(s) x {len(args.models)} model(s) x {len(args.prompts)} prompt(s) = {total_calls} total combinations")
    print(f"Per-call timeout: {args.timeout}s{'  (--force: redoing everything, resume disabled)' if args.force else '  (resume enabled — already-done combinations will be skipped)'}\n")

    report_paths = []
    for i, image_path in enumerate(image_paths, start=1):
        print(f"\n=== Image {i}/{len(image_paths)}: {image_path.name} ===")
        results = run_comparison(image_path, args.models, args.prompts, args.timeout, args.force)
        report_path = build_html_report(image_path, results)
        report_paths.append(report_path)
        print(f"Report -> {report_path}")

    print(f"\nAll outputs -> {OUTPUT_DIR}/")
    if len(report_paths) > 1:
        print(f"{len(report_paths)} reports generated:")
        for p in report_paths:
            print(f"  {p}")