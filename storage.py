"""Persist a page's transcription.

Two backends, selected by config.STORAGE_BACKEND:
  - "file"      : write <page>.md into transcripts/           (implemented)
  - "surrealdb" : push into SurrealDB instead/as well          (stubbed)

Keeping this behind one function means main.py and everything upstream of it
never needs to know or care where the output actually ends up.
"""
from pathlib import Path


def save_transcript(image_path: Path, content: str, transcripts_dir: Path, backend: str = "file") -> Path:
    """Save one page's transcription. Returns the path written to (file backend)."""
    if backend == "file":
        transcripts_dir = Path(transcripts_dir)
        transcripts_dir.mkdir(parents=True, exist_ok=True)
        out_path = transcripts_dir / f"{Path(image_path).stem}.md"
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
        #       "content": content,
        #   })
        #
        # Falling back to file storage for now so nothing is lost while
        # that's unwired.
        return save_transcript(image_path, content, transcripts_dir, backend="file")

    raise ValueError(f"Unknown storage backend: {backend!r}")
