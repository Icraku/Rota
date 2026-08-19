"""
Diagnostic for a slow run.

Checks:
  1. Is the Ollama server reachable at all?
  2. Is the model currently loaded in memory, or reloading per request?
     (A 35B model reloading cold on every page is the #1 cause of
     transcription that should take an hour taking a week.)
  3. How long does a trivial text-only round trip take? (isolates
     network/model overhead from image processing)
  4. How long does one real page take, start to finish?

Usage:
    python check_server.py                    # times text ping + first image in images/
"""
import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

import config
from llm_client import get_client, transcribe_image


def _get(url: str, timeout: int = 10) -> dict:
    with urllib.request.urlopen(urllib.request.Request(url), timeout=timeout) as resp:
        return json.loads(resp.read().decode())


def check_server_up() -> bool:
    url = config.IP_SERVER.rstrip("/") + "/api/tags"
    try:
        data = _get(url)
        models = [m["name"] for m in data.get("models", [])]
        print(f"[OK]   Server reachable at {config.IP_SERVER}")
        print(f"       Models available: {models}")
        if config.MODEL not in models:
            print(f"       WARNING: {config.MODEL!r} not found exactly in that list — check the tag (e.g. missing ':latest').")
        return True
    except Exception as e:
        print(f"[FAIL] Could not reach {url}: {e}")
        print("       Check the server is running and IP_SERVER/network access from this machine.")
        return False


def check_loaded_models() -> None:
    url = config.IP_SERVER.rstrip("/") + "/api/ps"
    try:
        data = _get(url)
        running = data.get("models", [])
        if not running:
            print("[INFO] No models currently loaded in memory.")
            print("       The next request pays a cold-start cost — for a 35B model this can be minutes —")
            print("       before it even starts responding. If this is still empty mid-run between pages,")
            print("       that reload cost is happening on every single page.")
        else:
            for m in running:
                print(f"[INFO] Loaded: {m.get('name')}  size_vram={m.get('size_vram')}  expires_at={m.get('expires_at')}")
    except Exception as e:
        print(f"[WARN] Could not query /api/ps ({e}) — older Ollama versions may not support this endpoint.")


def time_text_ping() -> float:
    client = get_client(config.IP_SERVER)
    print("Sending minimal text-only prompt...")
    start = time.time()
    client.chat(model=config.MODEL, messages=[{"role": "user", "content": "Reply with exactly: OK"}])
    elapsed = time.time() - start
    print(f"[TIMED] Text-only round trip: {elapsed:.1f}s")
    return elapsed


def time_image_transcribe(image_path: Path) -> float:
    client = get_client(config.IP_SERVER)
    prompt = Path(config.ACTIVE_PROMPT_FILE).read_text(encoding="utf-8")
    print(f"Transcribing {image_path} with the full active prompt...")
    start = time.time()
    result = transcribe_image(client, image_path, config.MODEL, prompt)
    elapsed = time.time() - start
    print(f"[TIMED] Full page transcription: {elapsed:.1f}s  (~{len(result)} chars returned)")
    return elapsed


if __name__ == "__main__":
    print(f"Checking {config.MODEL} @ {config.IP_SERVER}\n")

    if not check_server_up():
        sys.exit(1)
    check_loaded_models()
    print()

    time_text_ping()
    print()

    if len(sys.argv) > 1:
        time_image_transcribe(Path(sys.argv[1]))
    else:
        existing = sorted(config.IMAGES_DIR.glob("*.png"))
        if existing:
            time_image_transcribe(existing[0])
        else:
            print("[INFO] No image argument given and no .png found in images/ — skipping the full-page timing test.")
            print("       Run again as: python check_server.py images/page_001.png")