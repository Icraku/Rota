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
# FACILITY is the full display label (used in printouts/logs).
# HOSPITAL is the short, filename-safe tag used everywhere a deterministic
# per-hospital name is needed: image filenames ("page_001_Kakamega.png"),
# per-hospital subfolders under images/, transcripts/, and xlsx/. Keeping
# these separate means the display label can stay descriptive ("Kakamega
# NBU") while the tag used in paths/filenames stays short and stable.
#
# Both are overridable per run — see main.py's --facility flag, which sets
# both from a single value unless FACILITY is overridden separately.
FACILITY = os.getenv("ROTA_FACILITY", "Kakamega NBU")
HOSPITAL = os.getenv("ROTA_HOSPITAL", "Kakamega")

# ---------------------------------------------------------------------------
# Model / server
# ---------------------------------------------------------------------------
IP_SERVER = os.getenv("IP_SERVER", "172.16.13.68:11434")
IP_LOCAL = os.getenv("IP_LOCAL", "http://127.0.0.1:11434")
MODEL = os.getenv("ROTA_MODEL", "qwen3.6:35b")
MODEL2 = os.getenv("ROTA_MODEL", "qwen3-vl:4b")


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
IMAGES_DIR = BASE_DIR / "images"  # base folder — every model reads from here
TRANSCRIPTS_DIR = BASE_DIR / os.getenv("ROTA_TRANSCRIPTS_DIR", _transcripts_dir_name(MODEL))
PROMPTS_DIR = BASE_DIR / "prompts"
XLSX_DIR = BASE_DIR / "xlsx"

# Override per-run: `python main.py path/to/other.pdf` or `python main.py --pdf ...`
PDF_PATH = PDF_DIR / "Kakamega_NBU_Rota_2022.pdf"
ACTIVE_PROMPT_FILE = PROMPTS_DIR / "csv_direct.txt"

DPI = 200  # increase for dense/handwritten pages


# ---------------------------------------------------------------------------
# Per-hospital path helpers
# ---------------------------------------------------------------------------
# Every stage of the pipeline (images -> transcripts -> xlsx) keeps one
# hospital's files under its own subfolder, named after config.HOSPITAL (or
# whatever --facility override was passed on the command line). This is what
# makes lookup.py's hospital argument deterministic: "Kakamega" always means
# "whatever is under xlsx/Kakamega/", never a guess based on scanning page
# text or sheet names for a hospital's name (some real transcribed pages
# never mention their hospital anywhere in the text at all).
def images_dir_for(hospital: str) -> Path:
    return IMAGES_DIR / hospital


def transcripts_dir_for(hospital: str) -> Path:
    return TRANSCRIPTS_DIR / hospital


def xlsx_dir_for(hospital: str) -> Path:
    return XLSX_DIR / hospital


def xlsx_path_for(hospital: str) -> Path:
    return xlsx_dir_for(hospital) / "rota_transcripts.xlsx"


# ---------------------------------------------------------------------------
# Storage backend
# ---------------------------------------------------------------------------
# "file"      -> write each page's transcription as .md into TRANSCRIPTS_DIR
# "surrealdb" -> for future backend; falls back to "file" until the
#                connection details (host, namespace, database, table) are
#                wired up in storage.py
STORAGE_BACKEND = os.getenv("ROTA_STORAGE_BACKEND", "file")