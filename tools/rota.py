"""
PDF -> Image -> Transcription test script (Ollama backend).

1. Convert each page of a PDF into a PNG image (PyMuPDF).
2. Send each image to a vision-capable model served by Ollama.
3. Save the raw model output for each page as a .txt file next to the image.


Env vars:
    IP_SERVER  -- host[:port] of your Ollama server, e.g. "192.168.1.50:11434"
                  If unset, defaults to Ollama's local default (localhost:11434).
"""

import os
from pathlib import Path

import fitz  # PyMuPDF
from ollama import Client

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

PDF_PATH = "KAKAMEGA NBU ROTAS Jan -Dec 2023.pdf"
OUTPUT_DIR = "rota_transcriptions"
DPI = 200                          # increase for dense/handwritten pages

#IP_SERVER = os.getenv("IP_SERVER")
IP_SERVER="http://172.16.13.68:11434"
MODEL = "qwen3.5:35b"

PROMPT = """
# Optimized Rota Transcription Prompt

## Role
You are an expert document transcription system specializing in handwritten hospital duty rotas. Your task is to transcribe ALL visible content from the provided image into clean, structured Markdown. Do not summarize, interpret, or infer information — transcribe only what is visible, exactly as written.

## General Rules
1. Extract every visible piece of text: printed text, handwritten text, numbers, abbreviations, symbols (checkmarks, circles, arrows, dashes, crossed-out marks), labels, form fields, and marginal/annotation notes.
2. Transcribe literally. Do not expand, translate, or explain any code or symbol (e.g. do not turn "N" into "Night shift").
3. Do not guess missing or illegible text. Mark fully illegible content as "[unclear]" and partially legible content as "word[unclear]".
4. If a cell is visibly blank, write "[blank]" rather than omitting it — this keeps tables rectangular and machine-parseable.
5. If a mark has been crossed out, overwritten, or corrected, transcribe the final/most legible value and add a trailing note "[correction: originally looked like X]" if the earlier mark is still legibly distinguishable. If you cannot tell which is final, list both separated by "/" and add "[unclear correction]".
6. Note ink color only when it is objectively distinguishable and you're confident of it (e.g., "red", "blue", "black", "pencil") — append it in brackets after a cell's value, e.g. "NO [red]". Do not speculate about what the color means.
7. If the image is rotated relative to normal reading orientation, transcribe in the content's true reading order (top-to-bottom, left-to-right as a human would read it), not the raw file's pixel orientation. Note at the top of your output: "Image orientation: rotated [X]° from upright" if applicable.
8. Ignore purely decorative/graphical elements with no text (e.g. the county crest/logo image, ruled lines with no ink).

## Determining Page Type
Before transcribing, classify the image as one of:
- **Tabular rota page**: contains a grid with date/day columns and staff rows.
- **Mixed/multi-page shot**: a single photo capturing two or more physical pages side by side (e.g. two notebook pages). In this case, transcribe each physical page under its own heading and its own table(s), in left-to-right page order.

## Rules for Tables/Grids
1. Treat each distinct grid/week-block as its own table under its own heading (e.g. "## Rota Grid: Dates [X]–[Y]"). Do not merge separate week-blocks into one table, even if they appear on the same page or are visually adjacent/overlapping from a previous/next page bleeding through.
2. Transcribe the full header stack for each table: any title/month-year line (e.g. "NEW BORN UNIT ROTA FOR DECEMBER 2023"), the date-number row, and the day-abbreviation row (MON/TUE/WED/... or single-letter M/T/W/T/F/S/S), as separate header rows in the Markdown table.
3. Include the leftmost identifier columns exactly as labeled — typically "NO", "NAMES", and any row-label column (e.g. "I/C", "DEP I/C", "SUPPORT STAFF") — even when the NAMES cells are blank or covered/redacted (write "[blank]" or "[covered]" as appropriate).
4. Do NOT restrict cell values to a fixed set of codes. Common values in these rotas include E, D, Do/DO, PH, M, N, No, W, but cells may also legitimately contain: numbers (e.g. sequential counts in "DEP I/C" rows), dashes "-", fractions like "4/2/2" or "3/3", single letters that are part of a word spelled across the row (see rule 5), checkmarks, circled numbers, "X" marks, or short words/phrases (e.g. "MENTORSHIP", "owing", "LAST"). Transcribe exactly what is written, whatever it is.
5. **Spanning text across a row**: sometimes a single word or phrase is handwritten with one letter per day-column, spanning most/all of a row (e.g. "S C H O O L" or "M E N T O R S H I P" written across the week). When you detect this pattern, transcribe each cell with its single letter as written (to preserve the literal grid), AND add a one-line note immediately below that table row: "Note: row reads '[reconstructed word]' spanning columns [X]–[Y]."
6. **Multi-value cells**: some rows contain two stacked values in a single cell (e.g. two shift codes on two lines within one box). Transcribe both, separated by " / " in the same cell, top value first.
7. **Marginal/annotation columns**: if there is handwriting outside the main day-columns (e.g. in the "PHONE NUMBERS" column or margins) that isn't a phone number — such as tallies like "5PH", "6NO", "2DO" — transcribe it verbatim in that column/row; do not discard it as decorative.
8. If a header or mark visibly spans multiple sub-columns or sub-rows (a merged header), transcribe it once and repeat it under every column/row it covers, so the resulting table stays rectangular.
9. Each row represents someone therefore, if a blank row visibly spans multiple sun-columns, maintain it as it is when transcribing so the resulting table stays rectangular with the number of rows staying intact.
10. Preserve row order top-to-bottom and column order left-to-right exactly as laid out in the original grid.
11. **Never fragment a single grid into multiple partial tables.** A week-block/date-range is ONE table from its header row to its last ruled line, even if some rows in the middle are sparse, faint, or separated by visual whitespace. If a row has no visible handwriting at all, still include it as a row of all "[blank]" cells — do not skip it, and do not start a "new table" just because content resumed after a gap. Count the ruled/horizontal lines in the image to determine how many rows the table must have, and make sure your output has that many rows.
12. **Resolve each mark's column by vertical alignment to the date header, never by proximity to other marks.** Every value belongs to the date-column whose vertical gridlines it falls between — trace straight down from the "1, 2, 3, 4…" header to place it. Two values that are physically close together on the page but sit under different date columns (e.g. "DO" under column 1 and "PH" under column 2 in the same sparse row) must be transcribed as two separate cells in that row, never concatenated into one string or one row label. When in doubt about which column a mark belongs to, say so with "[column unclear]" rather than merging it with a neighboring value.
13. If a row is a numeric sequence row (e.g. "DEP I/C" running counts), each number still occupies exactly one date-column slot — do not compress the sequence into a shortened list.

## Footer / Metadata Fields
Below or beside the grid, transcribe any signature/metadata fields verbatim under a `## Footer` heading for that page, e.g.:
- PREPARED BY: [name or blank]
- DESIGNATION: [value]
- DATE: [value]
- COUNTERSIGNED BY: [name or blank]
- DESIGNATION: [value]
- DATE: [value]

Include any other handwritten names, numbers, or notes appearing outside the main grid (e.g. a name written near the bottom-left of the page, a phone number on a cover page) under this same heading or under `## Non-Tabular Content` as appropriate.

## Worked Example: Column Alignment for Sparse Rows

Given a row where the date-header columns are 1, 2, 3, 4, 5, 6, 7, 8, and the only visible handwriting in that row is "DO" positioned under column 1 and "PH" positioned under column 2 (columns 3–8 have no visible marks in that row):

**Correct** — one value per column, everything else marked blank:

| 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |
|---|---|---|---|---|---|---|---|
| DO | PH | [blank] | [blank] | [blank] | [blank] | [blank] | [blank] |

**Incorrect — do not do this:**
- Merging both values into a single cell or row label, e.g. a row labeled "DO PH" with no column breakdown.
- Dropping the row from the table because it's mostly empty.
- Shifting "PH" into column 1 and "DO" into column 2 because they were read in the wrong order — always trace each mark straight down to its date-header column, don't infer order from left-to-right reading of the handwriting alone if the marks aren't evenly spaced.

This same logic applies to every sparse row in these documents (there are many).

## Output Format
- Use Markdown headings to separate: page title/header info, each table, and footer/metadata.
- Each table must be a valid Markdown table (consistent column count per row).
- Do not add commentary, summaries, or interpretation anywhere in the output.

---
"""
# ---------------------------------------------------------------------------

def pdf_to_images(pdf_path: str, output_dir: str, dpi: int) -> list[str]:
    """Convert PDF to a PNG file to give a list of image paths."""
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)

    image_paths = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=matrix)
        image_path = os.path.join(output_dir, f"page_{i:03d}.png")
        pix.save(image_path)
        image_paths.append(image_path)
        print(f"Rendered {image_path}")
    return image_paths


def transcribe_image(client: Client, image_path: str, model: str, prompt: str) -> str:
    """Send one image to the vision model and return its raw text response."""
    with open(image_path, "rb") as f:
        image_bytes = f.read()

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


def main():
    client = Client(host=IP_SERVER) if IP_SERVER else Client()

    image_paths = pdf_to_images(PDF_PATH, OUTPUT_DIR, DPI)

    for image_path in image_paths:
        print(f"Transcribing {image_path} with model={MODEL} ...")
        result = transcribe_image(client, image_path, MODEL, PROMPT)

        out_path = image_path.replace(".png", "_transcribed_base3.txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"Saved -> {out_path}")

# MAIN
if __name__ == "__main__":
    main()