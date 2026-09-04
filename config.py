"""
Central configuration for the Rota transcription pipeline.

Everything that changes per facility or per run lives here, so the rest of
the pipeline (pdf_utils, llm_client, storage) never has to change.
"""
import os
from pathlib import Path

# ---------------------------------------------------------------------------
# Facility / run label
# ---------------------------------------------------------------------------
FACILITY = "Kakamega NBU"

# ---------------------------------------------------------------------------
# Model / server
# ---------------------------------------------------------------------------
IP_SERVER = os.getenv("IP_SERVER", "http://172.16.13.68:11434")
MODEL = os.getenv("ROTA_MODEL", "qwen3.6:35b")


def _transcripts_dir_name(model: str) -> str:
    """Pick a transcripts folder name based on which model is running.

    While comparing candidates, output lands in a model-specific folder
    (transcripts_qwen, transcripts_gemma, ...) so runs from different
    models never overwrite each other. Once you've settled on a final
    model, either rely on the "transcripts" fallback below (any model
    name that doesn't match a known family lands here automatically), or
    force it explicitly with the ROTA_TRANSCRIPTS_DIR env var.
    """
    lowered = model.lower()
    if "qwen" in lowered:
        return "transcripts_qwen"
    if "gemma" in lowered:
        return "transcripts_gemma"
    return "transcripts"


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

PDF_DIR = BASE_DIR / "pdfs"
IMAGES_DIR = BASE_DIR / "images"  # single shared folder — every model reads from here
TRANSCRIPTS_DIR = BASE_DIR / os.getenv("ROTA_TRANSCRIPTS_DIR", _transcripts_dir_name(MODEL))
PROMPTS_DIR = BASE_DIR / "prompts"
XLSX_DIR = BASE_DIR / "xlsx"

# Override per-run: `python main.py path/to/other.pdf`
PDF_PATH = PDF_DIR / "Kakamega_NBU_Rota_2022.pdf"
ACTIVE_PROMPT_FILE = PROMPTS_DIR / "base.txt"

DPI = 200  # increase for dense/handwritten pages

# ---------------------------------------------------------------------------
# Storage backend
# ---------------------------------------------------------------------------
# "file"      -> write each page's transcription as .md into TRANSCRIPTS_DIR
# "surrealdb" -> for future backend; falls back to "file" until the
#                connection details (host, namespace, database, table) are
#                wired up in storage.py
STORAGE_BACKEND = os.getenv("ROTA_STORAGE_BACKEND", "file")