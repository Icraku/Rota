"""
image -> vision model -> CSV.

Sends one page image to the configured Ollama model using a rota
transcription prompt, saves the model's raw Markdown response for
inspection, and parses every Markdown table found in that response
into its own CSV file.

This exists to iterate quickly on a prompt/model against a single
page — checking whether the extraction looks right in under a
minute — without running the full main.py batch pipeline or waiting
on the separate structuring stage.

Why Markdown-in, CSV-out instead of asking the model for CSV
directly: real rota cells contain commas (e.g. "5PH, 6NO, 2DO"),
which is exactly the character CSV uses as a delimiter. A model
asked to emit CSV directly has to get quoting/escaping right for
every such cell, which vision models are unreliable at. Markdown
tables use "|" as the delimiter instead, which never collides with
real cell content here, so the model's job stays simple and the CSV
conversion (which DOES need to handle commas/quotes correctly) is
done deterministically in Python instead.

Usage:
    python test_image_to_csv.py images/page_001.png
    python test_image_to_csv.py images/page_001.png --prompt prompts/csv_direct.txt
    python test_image_to_csv.py images/page_001.png --out out/page_001
"""
from __future__ import annotations

import argparse
import csv
import re
from pathlib import Path

import config
from pipeline.b_llm_client import get_client, transcribe_image

SEPARATOR_ROW = re.compile(r"\|(\s*:?-+:?\s*\|)+\s*$")


def parse_markdown_tables(markdown_text: str) -> list[list[list[str]]]:
    """Extract every Markdown table in the text.

    Returns a list of tables; each table is a list of rows; each row
    is a list of cell strings. The "| :--- | :--- |" separator row
    that Markdown tables use under the header is detected and
    dropped, not treated as a data row.

    A table is any run of consecutive lines that look like
    "| cell | cell | ...". This still works even if the model wraps
    its output in ``` code fences, since those fence lines don't
    start with "|" and are simply skipped over like any other
    non-table line.
    """
    tables: list[list[list[str]]] = []
    current: list[list[str]] = []

    for raw_line in markdown_text.splitlines():
        line = raw_line.strip()
        looks_like_row = line.startswith("|") and line.endswith("|") and len(line) > 1

        if looks_like_row:
            if SEPARATOR_ROW.fullmatch(line):
                continue  # the "| :--- | :--- |" divider — not a data row
            cells = [c.strip() for c in line.strip("|").split("|")]
            current.append(cells)
        elif current:
            tables.append(current)
            current = []

    if current:
        tables.append(current)

    return tables


def write_csv(table: list[list[str]], path: Path) -> None:
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(table)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("image", type=Path, help="Path to a single page image (e.g. images/page_001.png)")
    ap.add_argument(
        "--prompt",
        type=Path,
        default=Path("prompts/csv_direct.txt"),
        help="Prompt file to use (default: prompts/csv_direct.txt)",
    )
    ap.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Output path prefix (default: same folder/stem as the image)",
    )
    args = ap.parse_args()

    if not args.image.exists():
        raise SystemExit(f"Image not found: {args.image}")
    if not args.prompt.exists():
        raise SystemExit(f"Prompt file not found: {args.prompt}")

    out_prefix = args.out or args.image.with_suffix("")
    out_prefix.parent.mkdir(parents=True, exist_ok=True)

    prompt = args.prompt.read_text(encoding="utf-8")
    client = get_client(config.IP_LOCAL)

    print(f"Model:  {config.MODEL}")
    print(f"Image:  {args.image}")
    print(f"Prompt: {args.prompt}")
    print()

    markdown = transcribe_image(client, args.image, config.MODEL, prompt)

    raw_path = out_prefix.with_name(out_prefix.name + "_raw.md")
    raw_path.write_text(markdown, encoding="utf-8")
    print(f"\nSaved raw response -> {raw_path}")

    tables = parse_markdown_tables(markdown)
    if not tables:
        print("[WARNING] No Markdown tables found in the response.")
        print(f"          Open {raw_path} to see what the model actually returned.")
        return

    for i, table in enumerate(tables, start=1):
        n_cols = len(table[0]) if table else 0
        ragged = any(len(row) != n_cols for row in table)
        csv_path = out_prefix.with_name(f"{out_prefix.name}_table{i}.csv")
        write_csv(table, csv_path)
        flag = "  [WARNING: inconsistent column counts across rows]" if ragged else ""
        print(f"Saved table {i}: {len(table)} rows x {n_cols} cols -> {csv_path}{flag}")


if __name__ == "__main__":
    main()