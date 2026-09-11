""" Convert a Rota transcript (Markdown) into a styled .xlsx sheet.

Real transcript output varies a lot page to page — different column sets, extra
header rows inserted mid-grid, footers that don't all use the same fields.
Rather than fitting every page into one fixed template, this walks the markdown
top to bottom and renders each chunk in place:

  - a pipe-table block  -> a bordered, styled table
  - anything else       -> a plain text row (heading/bullet/paragraph)

Cell annotations from the transcription prompt's conventions are parsed out of
each table cell and turned into real Excel formatting:

  - [red] / [blue]                    -> font color
  - [strikethrough] / [crossed out] / [cross]  -> strikethrough font
  - [blank]                           -> empty cell
  - anything else in brackets (e.g. [covered], [correction: ...],
    [signature illegible]) -> kept as the visible value if the cell would
    otherwise be empty, and always added as a cell comment

This version is written to be easy to read line by line: plain for-loops
and if/else statements are used everywhere, even in a few places where a
shorter one-liner (a list comprehension, a ternary expression) would do
the same thing. If you already know Python, some of this will look more
spelled-out than it needs to be — that's on purpose.
"""
import re
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
SEP_CELL_RE = re.compile(r"^:?-+:?$")
BRACKET_RE = re.compile(r"\[([^\[\]]+)]")
WEEKDAY_RE = re.compile(r"^(MON|TUE|WED|THU|THUR|FRI|SAT|SUN)$", re.IGNORECASE)
BR_TAG_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
RATIO_VALUE_RE = re.compile(r"^\d+(/\d+){1,3}$")

STRIKE_TAGS = {"strikethrough", "crossed out", "cross", "crossed"}
COLOR_TAGS = {"red", "blue"}
RATIO_LABEL_WORDS = {"ratio", "support staff"}


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def split_segments(md_text: str) -> list[tuple[str, list[str]]]:
    """Split markdown into ('table', lines) / ('text', lines) chunks, in order."""
    lines = md_text.splitlines()
    segments = []
    buffer: list[str] = []

    def flush():
        if buffer:
            segments.append(("text", buffer.copy()))
            buffer.clear()

    i = 0
    while i < len(lines):
        if TABLE_ROW_RE.match(lines[i]):
            flush()
            table_lines = []
            while i < len(lines) and TABLE_ROW_RE.match(lines[i]):
                table_lines.append(lines[i])
                i += 1
            segments.append(("table", table_lines))
        else:
            buffer.append(lines[i])
            i += 1
    flush()
    return segments


def _is_separator_row(cells: list[str]) -> bool:
    """True for a Markdown table's divider row, e.g. ['---', ':---', '---']."""
    if not cells:
        return False
    for c in cells:
        if not SEP_CELL_RE.fullmatch(c.strip()):
            return False
    return True


def _pad_row(row: list[str], width: int) -> list[str]:
    """Return a copy of row with empty strings added until it reaches width."""
    padded = row.copy()
    while len(padded) < width:
        padded.append("")
    return padded


def normalize_cell_text(raw: str) -> str:
    """Flatten one raw table cell's markdown formatting before anything else
    looks at it, so header/day matching doesn't have to special-case every
    model's formatting choices.

    Different prompts/model runs write the same "date over weekday" header
    cell in different ways, e.g. "3 MON", "3 / MON", "**3**<br>**MON**", or
    "3<br>MON". This turns "<br>" (however it's written) into a plain space
    and strips "**bold**" markers, so all of those come out the same:
    "3 MON". (Bracket tags like [red] or [blank] are left alone here —
    parse_cell handles those separately.)
    """
    text = BR_TAG_RE.sub(" ", raw)
    text = BOLD_RE.sub(r"\1", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _is_blank_row(row: list[str]) -> bool:
    """True if every cell in the row is empty (after stripping whitespace)."""
    for c in row:
        if c.strip():
            return False
    return True


def _looks_like_ratio_values_row(row: list[str]) -> bool:
    """True if a row is mostly bare ratio values like '3/1/2' with no label.

    Some transcripts split the RATIO row in two: one row carrying the
    "SUPPORT STAFF" / "RATIO" labels, and a separate row right after it
    carrying just the numbers. This spots the second half of that pair so
    the two can be merged back into one logical row (see
    merge_split_data_rows) — otherwise the labels and the numbers end up
    attached to the wrong table row.
    """
    hits = 0
    for c in row:
        if RATIO_VALUE_RE.match(c.strip()):
            hits += 1
    return hits >= 2


def _merge_ratio_cell(a: str, b: str) -> str:
    """Combine one column's two cell values when merging a split RATIO row.

    A plain "RATIO" or "SUPPORT STAFF" label sometimes lands in the very
    same column a real value (e.g. "3/1/2") shows up in on the row right
    after it — the label isn't real per-column data, so it's dropped in
    favor of the real value rather than being glued onto it. If both sides
    are real values (or both are label words), they're joined with a space,
    same as the header merge does.
    """
    a_is_label = a.lower() in RATIO_LABEL_WORDS
    b_is_label = b.lower() in RATIO_LABEL_WORDS

    if a and b:
        if a_is_label and not b_is_label:
            return b
        if b_is_label and not a_is_label:
            return a
        if a == b:
            return a
        return f"{a} {b}"
    if a:
        return a
    return b


def merge_split_data_rows(data_rows: list[list[str]]) -> list[list[str]]:
    """Merge a "labels only" row with the "values only" row right after it,
    when a transcript has split the RATIO row across two physical table
    rows (see _looks_like_ratio_values_row). Every other row passes through
    unchanged.
    """
    merged_rows = []
    i = 0
    while i < len(data_rows):
        row = data_rows[i]
        has_next = i + 1 < len(data_rows)
        row_has_ratio_word = False
        for c in row:
            if "ratio" in c.strip().lower() or "support staff" in c.strip().lower():
                row_has_ratio_word = True
                break
        row_has_values = _looks_like_ratio_values_row(row)

        if has_next and row_has_ratio_word and not row_has_values and _looks_like_ratio_values_row(data_rows[i + 1]):
            next_row = data_rows[i + 1]
            width = max(len(row), len(next_row))
            row_a = _pad_row(row, width)
            row_b = _pad_row(next_row, width)
            combined = []
            for j in range(width):
                a = row_a[j].strip()
                b = row_b[j].strip()
                combined.append(_merge_ratio_cell(a, b))
            merged_rows.append(combined)
            i += 2
        else:
            merged_rows.append(row)
            i += 1
    return merged_rows


def parse_table_rows(table_lines: list[str]) -> list[list[str]]:
    """Turn raw '| a | b |' lines into a rectangular list of cell strings.

    Drops the markdown separator row (the '| :--- | :--- |' line), drops any
    fully-blank row(s) at the very top of the table (some transcripts render
    a stray empty row before the real header — see _is_blank_row), and pads
    every row to the widest row so the result stays rectangular even if the
    model produced a ragged table. Each cell also goes through
    normalize_cell_text so "<br>" tags and "**bold**" markers don't leak
    into header/day matching later on.
    """
    raw_rows = []
    for line in table_lines:
        inner = line.strip()
        if inner.startswith("|"):
            inner = inner[1:]
        if inner.endswith("|"):
            inner = inner[:-1]
        cells = []
        for one_cell in inner.split("|"):
            cells.append(normalize_cell_text(one_cell))
        raw_rows.append(cells)

    rows = []
    for r in raw_rows:
        if not _is_separator_row(r):
            rows.append(r)

    while rows and _is_blank_row(rows[0]):
        rows.pop(0)

    if not rows:
        return []

    max_cols = 0
    for r in rows:
        if len(r) > max_cols:
            max_cols = len(r)

    padded_rows = []
    for r in rows:
        padded_rows.append(_pad_row(r, max_cols))
    return padded_rows


def _looks_like_weekday_row(row: list[str]) -> bool:
    """True if a row is mostly weekday abbreviations (MON, TUE, ...).

    Real transcripts from this project's prompt often split the date header
    across two physical table rows instead of combining "24" and "MON" into
    one cell: one row of bare day numbers ('', '', '', '10', '11', ...) and,
    right after it, a row of column labels ('NO', 'NAMES', 'DATE DAY', 'MON',
    'TUE', ...). This detects the second row of that pair so the two can be
    merged into one logical header — otherwise the weekday-label row gets
    mistaken for the first row of real staff data.
    """
    hits = 0
    for c in row:
        if WEEKDAY_RE.match(c.strip()):
            hits += 1
    return hits >= 3


def build_header_and_data(rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    """Split parsed table rows into (header, data_rows), merging a split
    date/weekday header into one row when present (see _looks_like_weekday_row).

    Falls back to the simple "first row is the header" behavior when the
    second row doesn't look like a weekday-label row, so this is a safe
    drop-in replacement for `rows[0], rows[1:]` either way.

    Also runs the data rows through merge_split_data_rows, so a RATIO row
    that a transcript split into a "labels" row and a "values" row comes
    back out as a single logical row, the same way the split date/weekday
    header does.
    """
    if len(rows) >= 2 and _looks_like_weekday_row(rows[1]):
        row_a, row_b = rows[0], rows[1]
        width = max(len(row_a), len(row_b))
        row_a = _pad_row(row_a, width)
        row_b = _pad_row(row_b, width)

        header = []
        for i in range(width):
            a = row_a[i].strip()
            b = row_b[i].strip()
            if a and b:
                header.append(f"{a} {b}")
            elif a:
                header.append(a)
            else:
                header.append(b)
        return header, merge_split_data_rows(rows[2:])

    return rows[0], merge_split_data_rows(rows[1:])


def parse_cell(raw: str) -> dict:
    """Extract display value + styling from one transcript table cell."""
    tags = []
    for t in BRACKET_RE.findall(raw):
        tags.append(t.strip())
    text = BRACKET_RE.sub("", raw).strip()

    bold = False
    if text.startswith("**") and text.endswith("**") and len(text) > 4:
        text = text[2:-2].strip()
        bold = True

    tags_lower = []
    for t in tags:
        tags_lower.append(t.lower())

    is_red = "red" in tags_lower
    is_blue = "blue" in tags_lower
    is_blank = "blank" in tags_lower

    is_strike = False
    for t in tags_lower:
        if t in STRIKE_TAGS:
            is_strike = True
            break

    comment_tags = []
    for t in tags:
        if t.lower() not in ("red", "blue", "blank"):
            comment_tags.append(t)

    if text:
        value = text
    elif is_blank:
        value = ""
    elif tags:
        value = f"[{tags[0]}]"  # e.g. [covered], [unclear], [signature illegible]
    else:
        value = ""

    if comment_tags:
        comment = "; ".join(comment_tags)
    else:
        comment = None

    return {
        "value": value,
        "bold": bold,
        "red": is_red,
        "blue": is_blue,
        "strike": is_strike,
        "comment": comment,
    }


def parse_text_line(line: str) -> dict:
    """Turn one non-table markdown line into display text + light styling."""
    text = line.strip()
    if not text:
        return {"value": "", "bold": False, "italic": False}

    heading = bool(re.match(r"^#+\s*", text))
    text = re.sub(r"^#+\s*", "", text)

    bullet = False
    if text.startswith("- "):
        text, bullet = text[2:].strip(), True
    elif text.startswith("* ") and not text.startswith("**"):
        text, bullet = text[2:].strip(), True

    bold = heading
    if text.startswith("**") and text.endswith("**") and len(text) > 4 and text.count("**") == 2:
        text, bold = text[2:-2].strip(), True
    else:
        # partial inline bold (e.g. "- **Center:** some text") — strip the
        # markers without forcing the whole line bold
        text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)

    italic = False
    if text.startswith("*") and text.endswith("*") and len(text) > 2 and not text.startswith("**"):
        text, italic = text[1:-1].strip(), True

    if text in ("---", ""):
        return {"value": "", "bold": False, "italic": False}

    if bullet:
        prefix = "• "
    else:
        prefix = ""

    return {"value": prefix + text, "bold": bold, "italic": italic}


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def _font(bold=False, italic=False, red=False, blue=False, strike=False):
    if red:
        color = "FF0000"
    elif blue:
        color = "0000FF"
    else:
        color = "000000"
    return Font(name=FONT_NAME, bold=bold, italic=italic, color=color, strike=strike)


def render_table(ws, start_row: int, rows: list[list[str]]) -> tuple[int, int]:
    """Write one table starting at start_row. Returns (next_free_row, ncols)."""
    if not rows:
        return start_row, 0

    header, data_rows = build_header_and_data(rows)
    ncols = len(header)

    for j, raw in enumerate(header, start=1):
        parsed = parse_cell(raw)
        cell = ws.cell(row=start_row, column=j, value=parsed["value"] or raw.strip())
        cell.font = _font(bold=True, red=parsed["red"], blue=parsed["blue"])
        cell.alignment = CENTER
        cell.border = BORDER

    r = start_row + 1
    for row in data_rows:
        for j, raw in enumerate(row, start=1):
            parsed = parse_cell(raw)
            cell = ws.cell(row=r, column=j, value=parsed["value"])
            cell.font = _font(bold=parsed["bold"], red=parsed["red"], blue=parsed["blue"], strike=parsed["strike"])
            cell.alignment = CENTER
            cell.border = BORDER
            if parsed["comment"]:
                cell.comment = Comment(parsed["comment"], "Rota Transcript")
        r += 1
    return r + 1, ncols  # blank spacer row after the table


def render_text_block(ws, start_row: int, lines: list[str], width: int) -> int:
    r = start_row
    for line in lines:
        parsed = parse_text_line(line)
        if not parsed["value"]:
            r += 1
            continue
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(width, 1))
        cell = ws.cell(row=r, column=1, value=parsed["value"])
        cell.font = _font(bold=parsed["bold"], italic=parsed["italic"])
        cell.alignment = LEFT
        r += 1
    return r


def convert_transcript(md_path: Path, ws) -> None:
    """Render one transcript's markdown into the given worksheet, top to bottom."""
    md_text = Path(md_path).read_text(encoding="utf-8")
    segments = split_segments(md_text)

    # Pre-scan for the widest table so text rows merge across a sensible width.
    max_cols = 1
    for kind, lines in segments:
        if kind == "table":
            rows = parse_table_rows(lines)
            if rows and len(rows[0]) > max_cols:
                max_cols = len(rows[0])

    row = 1
    for kind, lines in segments:
        if kind == "table":
            rows = parse_table_rows(lines)
            if not rows:
                continue
            row, _ = render_table(ws, row, rows)
        else:
            row = render_text_block(ws, row, lines, max_cols)

    for col in range(1, max_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14


def sheet_name_for(md_path: Path) -> str:
    name = Path(md_path).stem
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", name)
    return cleaned[:31]  # Excel sheet names can't be longer than 31 characters


def convert_all(transcripts_dir: Path, output_path: Path) -> Path:
    """Convert every .md file in transcripts_dir into one workbook, one sheet per page."""
    transcripts_dir = Path(transcripts_dir)
    md_files = sorted(transcripts_dir.glob("*.md"))
    if not md_files:
        raise FileNotFoundError(f"No .md files found in {transcripts_dir}")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop the default blank sheet
    for md_path in md_files:
        ws = wb.create_sheet(title=sheet_name_for(md_path))
        convert_transcript(md_path, ws)
        print(f"Converted {md_path.name} -> sheet '{ws.title}'")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    import sys
    import config

    if len(sys.argv) > 1:
        out = Path(sys.argv[1])
    else:
        out = config.XLSX_DIR / "rota_transcripts.xlsx"

    convert_all(config.TRANSCRIPTS_DIR, out)
    print(f"Saved -> {out}")