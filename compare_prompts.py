"""
Compare transcription output across multiple prompts and models, side by side.

Does NOT modify main.py / llm_client.py / config.py — imports and reuses
them exactly as they are.

Runs every (model, prompt) combination against one image and produces:
  - one .md file per combination in comparisons/
  - one combined comparisons/<image>__comparison.html report with a column
    per combination, for quick visual side-by-side review in a browser

Each call runs on its own daemon thread with a timeout, so if one model
hangs (as qwen3.5:35b has done on this server before), it won't block the
rest of the sweep — the script moves on and reports the timeout instead of
freezing.

Usage:
    python compare_prompts.py
    python compare_prompts.py --models qwen3.5:35b gemma4:31b --prompts current base base2
    python compare_prompts.py --image images/page_003.png --models gemma4:31b --timeout 300
"""
import argparse
import html
import threading
import time
from pathlib import Path

import config
from llm_client import get_client, transcribe_image

DEFAULT_MODELS = ["qwen3.6:35b", "gemma4:31b"]
DEFAULT_PROMPTS = ["current", "base", "base2"]
DEFAULT_TIMEOUT = 60000  # seconds per (model, prompt) call
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


def run_comparison(image_path: Path, models: list[str], prompt_names: list[str], timeout: int) -> list[dict]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    results = []

    for model in models:
        client = get_client(config.IP_SERVER)
        for prompt_name in prompt_names:
            prompt_text = load_prompt(prompt_name)
            label = f"{model} / {prompt_name}"
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

            safe_model = model.replace(":", "_").replace("/", "_")
            out_path = OUTPUT_DIR / f"{image_path.stem}__{safe_model}__{prompt_name}.md"
            out_path.write_text(error and f"ERROR: {error}" or output, encoding="utf-8")

            results.append({
                "model": model,
                "prompt": prompt_name,
                "output": output,
                "error": error,
                "elapsed": elapsed,
                "path": out_path,
            })
    return results


def build_html_report(image_path: Path, results: list[dict]) -> Path:
    cols = "".join(
        f"""
        <div class="col{' error' if r['error'] else ''}">
          <h3>{html.escape(r['model'])}<br><span class="prompt">{html.escape(r['prompt'])}</span></h3>
          <div class="meta">{r['elapsed']:.1f}s{' — ERROR' if r['error'] else ''}</div>
          <pre>{html.escape(r['error'] or r['output'] or '(empty)')}</pre>
        </div>"""
        for r in results
    )

    doc = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Comparison — {html.escape(image_path.name)}</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 20px; background: #fafafa; }}
  h1 {{ font-size: 18px; }}
  .grid {{ display: flex; gap: 16px; overflow-x: auto; align-items: flex-start; }}
  .col {{ flex: 0 0 420px; background: white; border: 1px solid #ddd; border-radius: 6px; padding: 12px; }}
  .col.error {{ border-color: #e33; background: #fff5f5; }}
  .col h3 {{ margin: 0 0 4px 0; font-size: 14px; }}
  .col .prompt {{ font-weight: normal; color: #666; font-size: 12px; }}
  .meta {{ color: #888; font-size: 11px; margin-bottom: 8px; }}
  pre {{ white-space: pre-wrap; word-wrap: break-word; font-size: 12px; max-height: 80vh; overflow-y: auto; }}
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
    parser.add_argument("--image", type=Path, default=None, help="Image to transcribe (defaults to first .png in images/)")
    parser.add_argument("--models", nargs="+", default=DEFAULT_MODELS, help="Model tags to compare")
    parser.add_argument("--prompts", nargs="+", default=DEFAULT_PROMPTS, help="Prompt names (without .txt) from prompts/")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="Seconds to wait per call before giving up on it")
    args = parser.parse_args()

    image_path = args.image or pick_default_image()
    print(f"Comparing {len(args.models)} model(s) x {len(args.prompts)} prompt(s) on {image_path}")
    print(f"Per-call timeout: {args.timeout}s\n")

    results = run_comparison(image_path, args.models, args.prompts, args.timeout)
    report_path = build_html_report(image_path, results)

    print(f"\nIndividual outputs -> {OUTPUT_DIR}/")
    print(f"Side-by-side report -> {report_path}")