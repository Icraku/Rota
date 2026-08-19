"""Ollama vision-model client for transcribing rota page images."""
import time
from pathlib import Path

import httpx
from ollama import Client


def get_client(ip_server: str | None, timeout: float | None = 900) -> Client:
    """Build an Ollama client pointed at ip_server, or the local default if unset.

    timeout (seconds): bounds how long we wait for a response before raising
    an error, instead of hanging silently forever with no way to tell
    "still working" from "never coming back". 900s (15 min) is a
    deliberately generous diagnostic bound for now — once you've seen one
    successful transcription and know roughly how long a real page takes,
    adjust this to match.
    """
    kwargs = {"timeout": timeout} if timeout else {}
    return Client(host=ip_server, **kwargs) if ip_server else Client(**kwargs)


def transcribe_image(
    client: Client,
    image_path: Path,
    model: str,
    prompt: str,
    show_progress: bool = True,
    num_ctx: int = 16384,
    think: bool = False,
) -> str:
    """Send one image to the vision model and return its raw text response.

    think=False disables Qwen3-family "thinking mode". Left enabled (the
    Ollama default for these models), the model puts its output in a
    separate internal "thinking" field first — our streaming loop only
    watches "content", so if the model never finishes thinking, content
    stays completely empty even though eval_count shows real generation
    happened. This is very likely what several of our earlier "hangs" with
    qwen3.5/qwen3.6 actually were. gemma4 isn't a thinking-capable model, so
    it never had this problem — which is also probably why switching models
    kept "fixing" things earlier.

    num_ctx caps the context window Ollama allocates for this call, instead
    of Ollama's own VRAM-based default (which can be as high as 256K,
    allocated in full at load time regardless of actual need).

    Streams the response so progress is visible in real time, and raises a
    clear TimeoutError (instead of hanging forever) if nothing comes back
    within the client's configured timeout.
    """
    image_bytes = Path(image_path).read_bytes()

    chunks = []
    token_count = 0
    start = time.time()
    first_token_time = None
    final_stats = None
    thinking_chars = 0

    try:
        stream = client.chat(
            model=model,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                    "images": [image_bytes],
                }
            ],
            think=think,  # top-level, not inside options — see docstring
            options={"seed": 42, "num_ctx": num_ctx},  # add "num_predict": 4096 here if output gets truncated
            stream=True,
        )
        for part in stream:
            piece = part["message"]["content"]
            thinking_piece = part["message"].get("thinking") or ""
            if thinking_piece:
                thinking_chars += len(thinking_piece)
            if piece and first_token_time is None:
                first_token_time = time.time()
                if show_progress:
                    print(f"  [first content token after {first_token_time - start:.1f}s]")
            chunks.append(piece)
            token_count += 1
            if show_progress and token_count % 20 == 0:
                elapsed = time.time() - start
                print(f"  ... {token_count} chunks received ({elapsed:.0f}s elapsed)", end="\r")
            if part.get("done"):
                final_stats = part
    except httpx.TimeoutException as e:
        elapsed = time.time() - start
        raise TimeoutError(
            f"No response from {model} after {elapsed:.0f}s (client timeout hit). "
            f"Nothing was received at all — that points at the server/connection, not just slow "
            f"generation. Check the server directly (ollama ps, ollama logs, nvidia-smi on that "
            f"machine) rather than waiting longer."
        ) from e

    if show_progress:
        print()
        if thinking_chars:
            print(f"  [received {thinking_chars} chars of 'thinking' content — think={think}]")
        if first_token_time is None:
            print("  [WARNING: stream closed with zero content tokens received]")
        if final_stats:
            prompt_tokens = final_stats.get("prompt_eval_count")
            output_tokens = final_stats.get("eval_count")
            print(f"  [prompt_eval_count={prompt_tokens}, eval_count={output_tokens}, num_ctx={num_ctx}]")
            if prompt_tokens and prompt_tokens > num_ctx * 0.9:
                print(f"  [WARNING: prompt_eval_count is close to num_ctx={num_ctx} — output may be getting truncated, consider raising num_ctx]")

    return "".join(chunks)