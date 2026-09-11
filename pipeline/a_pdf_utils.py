"""PDF -> page image conversion."""
from pathlib import Path

import fitz  # PyMuPDF


def pdf_to_images(pdf_path: Path, output_dir: Path, dpi: int, hospital: str | None = None) -> list[Path]:
    """Converts every page of a PDF to a PNG, in page order.

    hospital: when given, every image is named "page_NNN_<hospital>.png"
        instead of plain "page_NNN.png" (e.g. "page_001_Kakamega.png").
        Combined with saving into a per-hospital output_dir (see
        config.images_dir_for), this is what makes a hospital's images
        identifiable two ways at once — by folder and by filename — so
        nothing downstream has to guess which hospital a page belongs to
        by scanning its (possibly incomplete) transcribed text. Omit to
        keep the old plain "page_NNN.png" naming.

    Returns: image paths in the same order as the PDF pages
    """
    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)

    image_paths = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=matrix)
        if hospital:
            image_path = output_dir / f"page_{i:03d}_{hospital}.png"
        else:
            image_path = output_dir / f"page_{i:03d}.png"
        pix.save(image_path)
        image_paths.append(image_path)
        print(f"Rendered {image_path}")
    return image_paths