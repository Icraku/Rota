# Rota

LLM-based transcription pipeline for handwritten nursing duty rotas, organized per hospital (currently: Kakamega NBU). No classical OCR — a vision-language model reads each page directly.

## Pipeline

```
PDF  -->  page images (images/<hospital>/)  -->  vision-LLM transcription  -->  transcripts/<model>/<hospital>/  -->  xlsx/<hospital>/
```

1. **`pipeline/a_pdf_utils.py`** — renders each PDF page to a PNG, in page order, named `page_NNN_<hospital>.png`.
2. **`pipeline/b_llm_client.py`** — sends one page image + the active prompt to the Ollama vision model.
3. **`pipeline/storage.py`** — saves each page's Markdown transcription; writes to a folder today, with a SurrealDB backend stubbed in for later.
4. **`main.py`** — wires the above together, first page to last.
5. **`pipeline/c_md_to_xlsx.py`** — the Markdown-table parser/renderer (format-tolerant: handles the various column layouts, split headers, and cell annotations real transcripts use). `to_xlsx.py` (root) is its CLI.
6. **`lookup.py`** — looks up one hospital + date's rota data, either straight out of the xlsx workbook (recommended) or the raw Markdown.
7. **`config.py`** — everything that changes per facility or run: paths, model, server, active prompt, hospital tag.

## Per-hospital organization

Every stage keeps one hospital's files under its own subfolder, named after a short, stable **hospital tag** (e.g. `"Kakamega"` — see `config.HOSPITAL`):

- `images/<hospital>/page_NNN_<hospital>.png`
- `transcripts/<model>/<hospital>/page_NNN[_<prompt>].md`
- `xlsx/<hospital>/rota_transcripts.xlsx`

This is what makes `lookup.py`'s hospital argument deterministic: `"Kakamega"` always means "whatever is under `xlsx/Kakamega/`", never a guess based on scanning page text or sheet names for a hospital's name. Some real transcribed pages never mention their hospital anywhere in the text at all — a folder you named yourself always works where text-scanning can't.

Source PDFs aren't organized into a per-hospital folder by convention, but naming them with the hospital in the filename (e.g. `Kakamega_NBU_Rota_2022.pdf`) keeps `pdfs/` self-documenting. Drop them there — `pdfs/` is gitignored, since rota data contains staff names.

## Running

```bash
# Transcribe: uses config.HOSPITAL / config.ACTIVE_PROMPT_FILE / config.PDF_PATH by default
python main.py ####
python main.py path/to/other.pdf                     #### (re)generate images from this PDF first
python main.py --facility Nakuru --prompt base2 --pdf pdfs/Nakuru_NBU_Rota_2023.pdf ####
python main.py --facility Kakamega_NBU --out-dir transcripts/Kakamega # generates the Markdown transcripts from images/Kakamega_NBU into a custom folder transcripts/Kakamega

# Convert that hospital's transcripts from its folder into an xlsx workbook
python to_xlsx.py --facility Kakamega_NBU --transcripts-dir transcripts/Kakamega

# Look up one hospital + date
python lookup.py Kakamega_NBU 2022-01-10
python lookup.py Kakamega_NBU 2022-01-10 --md             # same, but off the raw Markdown instead

# Legacy lookups against a workbook/folder that isn't organized per hospital yet
python lookup.py Kakamega 2022-12-16 xlsx/Kakamega_NBU/rota_transcripts.xlsx   # text-matched
python lookup.py Kakamega 2022-01-10 base2                        # raw Markdown, prompt_name filter
python lookup.py Kakamega 2022-01-10 --ground-truth ROTA_NBU_Nov_2024.xlsx
```

`--facility`/`--hospital` (main.py, to_xlsx.py) and the `hospital` argument (lookup.py) all take the same short tag — `config.HOSPITAL` (default `"Kakamega"`) if omitted. `config.FACILITY` is a separate, longer display label ("Kakamega NBU") used only in printouts.

## Prompts

`prompts/current.txt` is the live prompt used at runtime by default (see `config.ACTIVE_PROMPT_FILE`). `base.txt`, `base2.txt`, `currenta.txt`, `csv_direct.txt` are earlier iterations kept as a changelog — useful for comparing outputs if a change regresses on some page type. Pass `--prompt <name>` (main.py) to use a different one for a single run without touching `config.py`.

## Multi-hospital

Adding a second hospital needs no code changes — pass a different `--facility` tag to `main.py`/`to_xlsx.py`/`lookup.py` (or set `ROTA_HOSPITAL` / `ROTA_FACILITY` env vars to change the defaults), and its images/transcripts/xlsx land in their own subfolders automatically.