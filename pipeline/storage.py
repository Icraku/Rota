"""Persist a page's transcription.

Two backends, selected by config.STORAGE_BACKEND:

  - "file"      : write <page>.md into transcripts/ (implemented)
  - "surrealdb" : push into SurrealDB instead/as well (stubbed)

Keeping this behind one function means main.py and everything upstream
of it never needs to know or care where the output actually ends up.

Output is further tagged by which prompt produced it: pass prompt_name
(e.g. "base2", taken from the prompt file's stem) and pages are saved as
transcripts/<page>_<prompt_name>.md instead of plain transcripts/<page>.md.
That way switching prompts to compare results doesn't silently overwrite
the previous prompt's output — e.g. page_001.png run under base.txt and
again under base2.txt produces page_001_base.md and page_001_base2.md
side by side in the same folder. Omit prompt_name (or pass None) to keep
the old plain <page>.md naming.
"""
from pathlib import Path


def save_transcript(
    image_path: Path,
    content: str,
    transcripts_dir: Path,
    backend: str = "file",
    prompt_name: str | None = None,
) -> Path:
    """Save one page's transcription. Returns the path written to (file backend)."""
    if backend == "file":
        transcripts_dir = Path(transcripts_dir)
        transcripts_dir.mkdir(parents=True, exist_ok=True)
        stem = Path(image_path).stem
        if prompt_name:
            stem = f"{stem}_{prompt_name}"
        out_path = transcripts_dir / f"{stem}.md"
        out_path.write_text(content, encoding="utf-8")
        return out_path

    if backend == "surrealdb":
        # TODO: once this project has its own SurrealDB connection details
        # (host, namespace, database, table — e.g. mirroring BridgeProject's
        # setup), replace this block with something like:
        #
        #   from surrealdb import Surreal
        #   db = Surreal(SURREAL_URL)
        #   db.signin({...})
        #   db.use(NAMESPACE, DATABASE)
        #   db.create(TABLE, {
        #       "facility": facility,
        #       "page": Path(image_path).stem,
        #       "prompt": prompt_name,
        #       "content": content,
        #   })
        #
        # Falling back to file storage for now so nothing is lost while
        # that's unwired.
        return save_transcript(image_path, content, transcripts_dir, backend="file", prompt_name=prompt_name)

    raise ValueError(f"Unknown storage backend: {backend!r}")