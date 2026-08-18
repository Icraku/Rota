# Rota

LLM-based transcription pipeline for handwritten nursing duty rotas (currently: Kakamega NBU). No classical OCR — a vision-language model reads each page directly.

## Pipeline

```
PDF  -->  page images (images/)  -->  vision-LLM transcription  -->  transcripts/ (or SurrealDB)
```

1. **`pdf_utils.py`** — renders each PDF page to a PNG in `images/`, in page order.
2. **`llm_client.py`** — sends one page image + the active prompt to the Ollama vision model.
3. **`storage.py`** — saves each page's Markdown transcription; writes to `transcripts/` today, with a SurrealDB backend stubbed in for later.
4. **`main.py`** — wires the above together, first page to last.
5. **`config.py`** — everything that changes per facility or run: paths, model, server, active prompt.

## Prompts

`prompts/current.txt` is the live prompt used at runtime. `base.txt`, `base2.txt`, `base3.txt` are earlier iterations kept as a changelog — useful for comparing outputs if a change to `current.txt` regresses on some page type.

## Running

```bash
python main.py                      # transcribes config.PDF_PATH
python main.py path/to/other.pdf    # override for a different PDF/facility
```

Drop source PDFs in `pdfs/` (gitignored — rota data contains staff names).

## Multi-facility

Currently tailored to one facility via `config.FACILITY` and a single prompt. To add a second facility, the natural extension is a per-facility subfolder (prompt + PDF dir) selected via a CLI flag or config key, rather than editing `current.txt` in place.
