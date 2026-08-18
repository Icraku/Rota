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
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

PDF_DIR = BASE_DIR / "pdfs"
IMAGES_DIR = BASE_DIR / "images"
TRANSCRIPTS_DIR = BASE_DIR / "transcripts_base"
PROMPTS_DIR = BASE_DIR / "prompts"
XLSX_DIR = BASE_DIR / "xlsx"

# Override per-run: `python main.py path/to/other.pdf`
PDF_PATH = PDF_DIR / "Kakamega_NBU_Rota_2022.pdf"
ACTIVE_PROMPT_FILE = PROMPTS_DIR / "base.txt"

DPI = 200  # increase for dense/handwritten pages

# ---------------------------------------------------------------------------
# Model / server
# ---------------------------------------------------------------------------
IP_SERVER = os.getenv("IP_SERVER", "http://172.16.13.68:11434")
MODEL = os.getenv("ROTA_MODEL", "qwen3.5:35b")

# ---------------------------------------------------------------------------
# Storage backend
# ---------------------------------------------------------------------------
# "file"      -> write each page's transcription as .md into TRANSCRIPTS_DIR
# "surrealdb" -> for future backend; falls back to "file" until the
#                connection details (host, namespace, database, table) are
#                wired up in storage.py
STORAGE_BACKEND = os.getenv("ROTA_STORAGE_BACKEND", "file")
