"""
Look up rota data by hospital + date in the predicted data.

Give this a hospital and a date, it finds the right
page/sheet, finds the column for that date, and hands back every row's
value in that column (plus the RATIO row, pulled out on its own).

Usage as a script:
    python lookup.py Kakamega 2022-01-10
    python lookup.py Kakamega 2022-01-10 --md
    python lookup.py Kakamega 2022-01-10 xlsx/rota_transcripts.xlsx
    python lookup.py Kakamega 2022-01-10 base2
"""
from __future__ import annotations
import sys
from pathlib import Path

import re
from datetime import date as date_cls
from pathlib import Path
from typing import Union

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.comments import Comment
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


import config
from pipeline.c_md_to_xlsx import (
    split_segments,
    parse_table_rows,
    build_header_and_data,
    RATIO_VALUE_RE,
)

# Full month names and common abbreviations ("Jan", "Sept"/"Sep", ...) —
# real transcripts spell this every which way ("JANUARY 2022", "JAN 2022",
# "Jan. 2022"), so all of those are accepted.
MONTH_YEAR_RE = re.compile(
    r"\b(January|February|March|April|May|June|July|August|September|October|November|December"
    r"|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sept|Sep|Oct|Nov|Dec)\b\.?"
    r"\s+(\d{4})",
    re.IGNORECASE,
)
MONTHS = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sept": 9, "sep": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}
LEADING_DAY_RE = re.compile(r"^\D*(\d{1,2})\b")

# How many letters "off" a hospital-name match is still allowed to be, e.g.
# matching the typo "KAKAMEGA" (one letter dropped) against "Kakamega".
HOSPITAL_NAME_MAX_TYPOS = 2

# Recognized shift codes, per the transcription prompt's conventions
# (M/N/No/D/Do/PH/E) plus the special-row designations seen in practice
# (I/C, DEP I/C). Used only to flag a row that looks shifted out of place —
# never to "correct" a value.
KNOWN_SHIFT_CODES = {"M", "E", "N", "D", "DO", "NO", "PH", "I/C", "DEP I/C"}

FONT_NAME = "Arial"
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ---------------------------------------------------------------------------
# Small building blocks used by everything else in this file
# ---------------------------------------------------------------------------

def _as_date(value: Union[str, date_cls]) -> date_cls:
    """Accept either a real date or an ISO string like '2022-01-10'."""
    if isinstance(value, date_cls):
        return value
    return date_cls.fromisoformat(str(value))


def _edit_distance(a: str, b: str) -> int:
    """How many single-letter changes (insert, delete, swap) it takes to
    turn `a` into `b`. Used only to tolerate small transcription typos in
    a hospital name (see _hospital_mentioned) — never to "correct" rota
    data itself.
    """
    rows = len(a) + 1
    cols = len(b) + 1
    table = [[0] * cols for _ in range(rows)]
    for i in range(rows):
        table[i][0] = i
    for j in range(cols):
        table[0][j] = j

    for i in range(1, rows):
        for j in range(1, cols):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            delete_cost = table[i - 1][j] + 1
            insert_cost = table[i][j - 1] + 1
            substitute_cost = table[i - 1][j - 1] + cost
            table[i][j] = min(delete_cost, insert_cost, substitute_cost)

    return table[rows - 1][cols - 1]


def _hospital_mentioned(hospital: str, text: str) -> bool:
    """True if `hospital` shows up in `text`, typo and all.

    First tries a plain case-insensitive substring check (cheap, and right
    almost all the time). If that comes up empty, it slides a
    same-ish-length window across the text and checks whether any window
    is close enough (within HOSPITAL_NAME_MAX_TYPOS single-letter edits)
    to `hospital` — this is what lets "Kakamega" still match a page that
    got transcribed as "KAKAMEA" (one letter dropped).
    """
    hospital_lower = hospital.lower()
    text_lower = text.lower()

    if hospital_lower in text_lower:
        return True

    needle_len = len(hospital_lower)
    shortest_window = max(1, needle_len - HOSPITAL_NAME_MAX_TYPOS)
    longest_window = needle_len + HOSPITAL_NAME_MAX_TYPOS

    for window_len in range(shortest_window, longest_window + 1):
        for start in range(0, len(text_lower) - window_len + 1):
            window = text_lower[start:start + window_len]
            if _edit_distance(window, hospital_lower) <= HOSPITAL_NAME_MAX_TYPOS:
                return True

    return False


def _page_month_years(text: str) -> list[tuple[int, int]]:
    """Every (month, year) pair mentioned anywhere in `text`, e.g.
    "...ROTA FOR JANUARY 2022" -> [(1, 2022)].

    Every month/year mention found is returned, not just the first one.

    An empty list means the page never states a month/year at all — the
    caller checks its date columns anyway
    """
    pairs = []
    for m in MONTH_YEAR_RE.finditer(text):
        pair = (MONTHS[m.group(1).lower()], int(m.group(2)))
        if pair not in pairs:
            pairs.append(pair)
    return pairs


def _matching_day_column(header: list[str], day: int) -> int | None:
    """Which position in `header` is the column for `day`? (e.g. header
    cell "10 MON" matches day=10). Returns None if there's no such column.
    """
    for position, cell in enumerate(header):
        m = LEADING_DAY_RE.match(cell.strip())
        if m and int(m.group(1)) == day:
            return position
    return None


def _is_ratio_row(row_cells: list[str]) -> bool:
    """True for the RATIO/summary row — checked by looking for the word
    "ratio" or "support staff" in the first few cells (not just "ratio"
    alone), since some transcripts end up with "SUPPORT STAFF" sitting in
    the column other transcripts use for "RATIO" itself, once a row that
    got split across two physical table rows is merged back together.
    """
    for cell in row_cells[:3]:
        text = cell.strip().lower()
        if "ratio" in text or "support staff" in text:
            return True
    return False


def _phone_column_index(header: list[str]) -> int | None:
    """Which column is PHONE NO? Found by name rather than assumed to be
    the last column, since short rows get padded with blank trailing
    columns and that padding can land after the real PHONE NO column.
    """
    for position, cell in enumerate(header):
        if "phone" in cell.lower():
            return position
    return None


def _flag_misaligned(header: list[str], row_cells: list[str]) -> bool:
    """A cheap sanity check, Markdown transcripts only: the PHONE NO column
    should always be either blank or an actual phone number. If it instead
    holds something that reads like a shift code (M/E/N/Do/No/PH/...), that
    usually means this row's cells got shifted one column to the right
    somewhere during transcription. This just flags that — it never tries
    to guess the "real" value, since a wrong guess is much harder to spot
    later than an honest "this looks off, please check."
    """
    phone_index = _phone_column_index(header)
    if phone_index is None or phone_index >= len(row_cells):
        return False
    value = row_cells[phone_index].strip()
    return bool(value) and value.upper() in KNOWN_SHIFT_CODES


def _row_label(header: list[str], row_cells: list[str], position: int | None = None) -> str:
    """A human-readable name for one row.

    Normally that's whatever's in the DATE-DAY-style label column (e.g.
    "DEP I/C", "RATIO"), or failing that the NAMES column, or failing that
    just "NO <n>". But some rota pages get split into narrow "week strip"
    tables that have no ID columns at all — every column is a date column
    — which you can tell because the table's very first header cell is
    itself a date number instead of a real label like "NO". There's
    nothing to read a name from in that case, so this falls back to "Row
    <n>" (its plain position in the table) instead.
    """
    looks_like_a_date_table = bool(header) and bool(LEADING_DAY_RE.match(header[0].strip()))
    if looks_like_a_date_table and position is not None:
        return f"Row {position}"

    no = row_cells[0].strip() if len(row_cells) > 0 else ""
    names = row_cells[1].strip() if len(row_cells) > 1 else ""
    date_day = row_cells[2].strip() if len(row_cells) > 2 else ""

    if date_day:
        return date_day
    if names:
        return names
    if no:
        return f"NO {no}"
    return "?"


def _rows_for_day(header: list[str], data_rows: list[list[str]], day: int) -> list[dict] | None:
    """Given one parsed table (a header row + its data rows) and a target
    day-of-month, return one dict per real staff row in that day's column
    — or None if this table doesn't have a column for that day at all.

    Each dict looks like:
        {"label": "DEP I/C", "value": "PH", "is_ratio": False, "row_cells": [...]}

    This is shared by both the Markdown path and the xlsx path (both end
    up with a header + data_rows in the same shape), so the "find the
    column, skip blank rows, read off the values" logic only needs to
    exist once. It's written with a small pandas table because that makes
    "drop every row that's completely blank" and "grab one whole column of
    values" both a single, easy-to-read line, instead of a hand-rolled
    loop with an if-statement inside it.
    """
    column_position = _matching_day_column(header, day)
    if column_position is None:
        return None

    # One row per person. Columns are just numbered 0, 1, 2, ... (their
    # position) rather than named after the header text — header cells are
    # sometimes blank or repeated, and a DataFrame needs a column name it
    # can look up without ambiguity.
    table = pd.DataFrame(data_rows)
    if table.empty or column_position >= table.shape[1]:
        return []

    # A row where every cell is blank is just page padding, not a real
    # person — drop those before reading anything off.
    is_blank_row = table.apply(lambda cells: all(str(c).strip() == "" for c in cells), axis=1)
    table = table[~is_blank_row]

    entries = []
    position = 0
    for row_index, row in table.iterrows():
        row_cells = [str(c) for c in row.tolist()]
        position += 1
        entries.append({
            "label": _row_label(header, row_cells, position=position),
            "value": row_cells[column_position].strip(),
            "is_ratio": _is_ratio_row(row_cells),
            "row_cells": row_cells,
        })

    return entries


def _group_rows_by_source(rows: list[dict], date: Union[str, date_cls]) -> list[dict]:
    """Turn a flat list of row-dicts (as returned by get_hospital_date_rows
    or get_hospital_date_rows_from_xlsx) into one {entries, ratio,
    warnings} result per "source_file" they came from, sorted by that
    name.

    This is the shared last step behind every "give me every matching
    page/sheet, not just one" function below (hospital_date_lookup,
    hospital_date_column_from_xlsx, markdown_hospital_date_lookup).
    Grouping rows by which file/sheet they came from is exactly what
    pandas's groupby is for, so that's used here instead of a hand-rolled
    loop that builds the groups itself.
    """
    if not rows:
        return []

    rows_df = pd.DataFrame(rows)

    results = []
    for source_name, source_rows in rows_df.groupby("source_file", sort=True):
        entries = []
        ratio = None
        warnings = []

        if not source_rows.iloc[0]["month_stated"]:
            warnings.append(
                f"{source_name!r} never states its own month/year anywhere — its date "
                f"columns were matched with no way to confirm they're really for {date!r}. "
                f"The same page/sheet would match just as well for the same day-of-month "
                f"in a different month/year — verify against the source page."
            )

        for _, r in source_rows.iterrows():
            if r["is_ratio"]:
                ratio = r["value"]
                continue
            entries.append({"label": r["label"], "value": r["value"], "misaligned": r["misaligned"]})
            if r["misaligned"]:
                warnings.append(
                    f"Row {r['label']!r} in {source_name!r}: value {r['value']!r} looks shifted "
                    f"out of place (its PHONE NO column held a shift-code-like value) — verify "
                    f"against the source page."
                )

        results.append({
            "source_file": source_name,
            "date_column": source_rows.iloc[0]["date_column"],
            "entries": entries,
            "ratio": ratio,
            "warnings": warnings,
        })

    return results


# ---------------------------------------------------------------------------
# Markdown transcripts
# ---------------------------------------------------------------------------

def get_hospital_date_rows(
    hospital: str,
    date: Union[str, date_cls],
    transcripts_dir: Path,
) -> list[dict]:
    """Every rota row tied to `hospital` + `date`, straight out of the raw
    Markdown transcripts under `transcripts_dir` (including subfolders).

    hospital: matched, case-insensitively and typo-tolerant, against (1)
        the .md file's own path first (its filename or a parent folder
        name — e.g. a file saved as "page_046_kakamega.md" or living
        inside a "kakamega/" folder), then (2) the page's transcribed
        text. The path is checked first because it's something you
        control yourself, where a handful of real pages here never
        mention their hospital anywhere in the transcribed text at all.

    A page whose title states a month/year that isn't `date`'s is skipped
    — so a "10" column on a February page is never confused with January
    10th. A page that never states a month/year at all can't be ruled out
    that way, so it's checked anyway; see the "month_stated" field below.

    Returns a list of dicts, one per matching row:
        source_file  - which .md file this came from
        table_index  - which table on that page (0-based, in page order)
        header       - that table's column header row
        date_column  - the header text of the matched date column
        value        - this row's value in that column
        label        - a human-readable row name (see _row_label)
        is_ratio     - True for the RATIO/summary row
        misaligned   - True if this row's cells look shifted (see _flag_misaligned)
        month_stated - False if the page never states its own month/year
                       anywhere — see the module docstring caveats
    """
    target = _as_date(date)
    transcripts_dir = Path(transcripts_dir)
    results: list[dict] = []

    for md_path in sorted(transcripts_dir.rglob("*.md")):
        md_text = md_path.read_text(encoding="utf-8")

        if not (_hospital_mentioned(hospital, str(md_path)) or _hospital_mentioned(hospital, md_text)):
            continue

        page_month_years = _page_month_years(md_text)
        if page_month_years and (target.month, target.year) not in page_month_years:
            continue
        month_stated = bool(page_month_years)

        table_index = -1
        for kind, lines in split_segments(md_text):
            if kind != "table":
                continue
            table_index += 1

            rows = parse_table_rows(lines)
            if not rows:
                continue
            header, data_rows = build_header_and_data(rows)

            day_entries = _rows_for_day(header, data_rows, target.day)
            if not day_entries:
                continue

            for entry in day_entries:
                results.append({
                    "source_file": md_path.name,
                    "table_index": table_index,
                    "header": header,
                    "date_column": header[_matching_day_column(header, target.day)],
                    "value": entry["value"],
                    "label": entry["label"],
                    "is_ratio": entry["is_ratio"],
                    "misaligned": _flag_misaligned(header, entry["row_cells"]),
                    "month_stated": month_stated,
                })

    return results


def hospital_date_column(
    hospital: str,
    date: Union[str, date_cls],
    transcripts_dir: Path,
    prompt_name: str | None = None,
) -> dict:
    """Convenience wrapper for the common case: exactly one page matches.

    prompt_name: use this when main.py saved more than one prompt's output
        for the same page (e.g. page_001_base.md and page_001_base2.md).
        Pass just the prompt part, e.g. "base2", to only consider files
        whose name ends in "_base2.md".

    Raises ValueError if zero, or more than one, page matches (after the
    prompt_name filter, if given) — call get_hospital_date_rows directly
    if you'd rather see every matching page than have one picked for you.
    """
    rows = get_hospital_date_rows(hospital, date, transcripts_dir)

    if prompt_name:
        rows = [r for r in rows if Path(r["source_file"]).stem.endswith(f"_{prompt_name}")]

    if not rows:
        extra = f", prompt_name={prompt_name!r}" if prompt_name else ""
        raise ValueError(f"No data found for hospital={hospital!r}, date={date!r}{extra} in {transcripts_dir}")

    grouped = _group_rows_by_source(rows, date)
    if len(grouped) > 1:
        matched_files = sorted(r["source_file"] for r in rows)
        raise ValueError(
            f"{len(grouped)} different transcript files matched ({matched_files}) — "
            f"pass prompt_name= to pick one, e.g. prompt_name=\"base2\" for page_001_base2.md."
        )

    only_result = grouped[0]
    return {
        "source_file": only_result["source_file"],
        "date_column": only_result["date_column"],
        "entries": only_result["entries"],
        "ratio": only_result["ratio"],
        "warnings": only_result["warnings"],
    }


def export_column_xlsx(
    hospital: str,
    date: Union[str, date_cls],
    transcripts_dir: Path,
    output_path: Path,
    prompt_name: str | None = None,
) -> Path:
    """Write hospital_date_column(...)'s result out as a small styled
    .xlsx: one row per entry (Label, Value), the RATIO value as its own
    bold row at the bottom, and a cell comment on any row flagged as
    possibly misaligned.

    This writes a *formatted* file (fonts, borders, comments), which is a
    job openpyxl is built for and pandas isn't — pandas' own .to_excel()
    can't add a cell comment or set an individual cell's font, so this one
    function stays on openpyxl rather than pandas.
    """
    result = hospital_date_column(hospital, date, transcripts_dir, prompt_name=prompt_name)

    wb = openpyxl.Workbook()
    ws = wb.active

    raw_title = f"{hospital}_{date}"
    ws.title = re.sub(r"[\[\]:*?/\\]", "_", raw_title)[:31]  # Excel sheet names top out at 31 characters

    ws.merge_cells("A1:B1")
    title_cell = ws.cell(row=1, column=1, value=f"{hospital} — {result['date_column']} ({date})")
    title_cell.font = Font(name=FONT_NAME, bold=True, size=12)
    title_cell.alignment = CENTER

    header_row = 3
    for j, text in enumerate(("Label", "Value"), start=1):
        cell = ws.cell(row=header_row, column=j, value=text)
        cell.font = Font(name=FONT_NAME, bold=True)
        cell.alignment = CENTER
        cell.border = BORDER

    r = header_row + 1
    for entry in result["entries"]:
        label_cell = ws.cell(row=r, column=1, value=entry["label"])
        value_cell = ws.cell(row=r, column=2, value=entry["value"])
        for cell in (label_cell, value_cell):
            cell.font = Font(name=FONT_NAME)
            cell.alignment = CENTER
            cell.border = BORDER
        if entry["misaligned"]:
            value_cell.comment = Comment(
                "Possibly misaligned in the source transcript — verify against the scanned page.",
                "Rota Lookup",
            )
        r += 1

    if result["ratio"] is not None:
        r += 1
        label_cell = ws.cell(row=r, column=1, value="RATIO")
        value_cell = ws.cell(row=r, column=2, value=result["ratio"])
        for cell in (label_cell, value_cell):
            cell.font = Font(name=FONT_NAME, bold=True)
            cell.alignment = CENTER
            cell.border = BORDER

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 14

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Same lookup, but reading an already-built .xlsx instead of raw Markdown
# ---------------------------------------------------------------------------

def _sheet_text(sheet_df: pd.DataFrame) -> str:
    """Every non-blank cell's text on a sheet, joined into one big string
    — used for the same "does this mention the hospital / the month" text
    checks that run against raw Markdown text.
    """
    pieces = []
    for value in sheet_df.to_numpy().flatten():
        if pd.notna(value) and str(value).strip():
            pieces.append(str(value))
    return "\n".join(pieces)


def _cell_has_border(cell) -> bool:
    """True if any of the cell's four sides actually has a border drawn on
    it (has a style, e.g. "thin"). A cell that was never touched by
    to_xlsx.py, or that only ever had a bare value written to it, has no
    border at all — only cells written as part of a rendered table do (see
    pipeline/c_md_to_xlsx.py's render_table).
    """
    border = cell.border
    if border is None:
        return False
    return any(side is not None and side.style for side in (border.left, border.right, border.top, border.bottom))


def _load_workbook_sheets(xlsx_path: Path) -> dict:
    """Read every sheet in a workbook two ways at once and hand back both
    together: its plain values (via pandas — header=None so nothing is
    treated as column titles, since a rota table's real header usually
    isn't row 1 of the sheet) and, for each row, whether ANY cell in it
    has a border drawn around it (via openpyxl, the only one of the two
    that can even see cell formatting).

    The border flag is what _find_table_blocks_in_sheet uses to split a
    sheet into separate tables — see the module docstring for why that
    can't just be done from the values alone.

    Returns {sheet_name: (sheet_df, row_has_border)}, where row_has_border
    is a plain list of one bool per row, in the same top-to-bottom order
    as sheet_df's own rows.
    """
    all_sheets = pd.read_excel(xlsx_path, sheet_name=None, header=None)
    workbook = openpyxl.load_workbook(xlsx_path, data_only=True)

    sheets = {}
    for sheet_name, sheet_df in all_sheets.items():
        ws = workbook[sheet_name]
        row_has_border = [any(_cell_has_border(cell) for cell in ws_row) for ws_row in ws.iter_rows()]
        sheets[sheet_name] = (sheet_df, row_has_border)
    return sheets


def _find_table_blocks_in_sheet(sheet_df: pd.DataFrame, row_has_border: list[bool]) -> list[list[list[str]]]:
    """Split one sheet into separate tables, using cell borders — not
    blank values — to tell where one table ends and the next begins.

    row_has_border (from _load_workbook_sheets) says, for each row in
    sheet_df, whether it's part of a bordered table at all; a row with no
    border anywhere in it is a genuine gap between tables, even if a
    bordered row right next to it happens to have every value blank too
    (a real spacer row inside one table, e.g. right before a RATIO row —
    see the module docstring).

    Returns a list of blocks, one per table found, each block a list of
    rows (each row a list of cell-text strings) in the order they appear
    — block[0] is that table's header row, block[1:] are its data rows.
    """
    text_grid = sheet_df.fillna("").astype(str)

    blocks: list[list[list[str]]] = []
    current_block: list[list[str]] = []
    for i, (_, row) in enumerate(text_grid.iterrows()):
        cells = [cell.strip() for cell in row.tolist()]
        is_table_row = row_has_border[i] if i < len(row_has_border) else False
        if is_table_row:
            current_block.append(cells)
        elif current_block:
            blocks.append(current_block)
            current_block = []
    if current_block:
        blocks.append(current_block)
    return blocks


def _looks_like_day_header_row(row: list[str], min_hits: int = 3) -> bool:
    """True if a row looks like a table's "date numbers" row — at least
    min_hits of its cells start with a short (1-2 digit) number, e.g. a
    plain ['27', '28', '29', ...] row, a combined ['27 MON', '28 TUE', ...]
    row, or a "continued" one like ['32', '33', ...] some transcripts use
    to keep counting past the end of a month instead of resetting to 1.

    A ratio-style value like "2/2/2" also starts with a digit but isn't a
    date, so RATIO_VALUE_RE filters those out first — a RATIO row full of
    such values is never mistaken for a header.

    min_hits defaults to 3, since a normal data row can, purely by luck,
    have a couple of cells that happen to start with a number too — asking
    for several in the same row is what keeps this from false-triggering
    on those. Pass a lower min_hits (see _split_block_into_minitables) for
    a table's very first row specifically: a table can legitimately only
    span one or two date columns, and being row 0 right after a detected
    table boundary already makes "this is the header" a safe bet.
    """
    hits = 0
    for cell in row:
        text = cell.strip()
        if not text or RATIO_VALUE_RE.match(text):
            continue
        if LEADING_DAY_RE.match(text):
            hits += 1
    return hits >= min_hits


def _split_block_into_minitables(block: list[list[str]]) -> list[tuple[list[str], list[list[str]]]]:
    """Split one table block into one or more (header, data_rows)
    mini-tables, by finding every row in it that looks like a date-numbers
    row (see _looks_like_day_header_row) and treating everything up to the
    next one — or the end of the block — as that mini-table's data.

    Most tables only have one such row, right at the top, in the usual "NO
    | NAMES | DATE DAY | 27 MON | 28 TUE | ..." layout. When block[0]
    already looks like a date-numbers row, it's trusted as THE header and
    nothing after it is re-examined — a real staff row can, purely by
    coincidence, also contain several cells that look like an ascending
    run of numbers, and re-scanning an already-correct table for "more
    headers" risks mistaking a row like that for a second header and
    cutting the real one off from most of its own data.

    Only when block[0] does NOT look like a date-numbers row (e.g. a
    weekday-letters row comes first, or a whole month got split into
    several narrower "week strip" tables stacked in this same block, each
    restating its own date-numbers row) does this fall back to scanning
    every row for one. That's a looser search — the same false-positive
    risk applies here too — but it's the only way to find a header that
    isn't at the top.
    """
    if not block:
        return []

    if _looks_like_day_header_row(block[0], min_hits=1):
        return [(block[0], block[1:])]

    header_positions = [i for i, row in enumerate(block) if _looks_like_day_header_row(row)]
    if not header_positions:
        return []

    minitables = []
    for k, start in enumerate(header_positions):
        end = header_positions[k + 1] if k + 1 < len(header_positions) else len(block)
        minitables.append((block[start], block[start + 1:end]))
    return minitables


def get_hospital_date_rows_from_xlsx(
    hospital: str,
    date: Union[str, date_cls],
    xlsx_path: Path,
) -> list[dict]:
    """Same idea as get_hospital_date_rows, but reads an already-built
    workbook (one produced by to_xlsx.py) instead of re-parsing the raw
    Markdown. See the module docstring for how tables are found on each
    sheet, and its trade-offs.

    Returns the same shape as get_hospital_date_rows, except:
      - "source_file" holds the sheet name (e.g. "page_001_base2") instead
        of a .md filename.
      - "misaligned" is always False — the PHONE-NO-column check needs the
        row's original raw text, and by the time a value is sitting in a
        cell here, that raw text is long gone.
    """
    target = _as_date(date)
    xlsx_path = Path(xlsx_path)

    # See _load_workbook_sheets: this reads both the plain values (pandas)
    # and, per row, whether it's part of a bordered table (openpyxl) —
    # table boundaries are found from the borders, everything else from
    # the values.
    sheets = _load_workbook_sheets(xlsx_path)

    results: list[dict] = []
    for sheet_name, (sheet_df, row_has_border) in sheets.items():
        sheet_text = _sheet_text(sheet_df)

        if not (_hospital_mentioned(hospital, sheet_name) or _hospital_mentioned(hospital, sheet_text)):
            continue

        page_month_years = _page_month_years(sheet_text)
        if page_month_years and (target.month, target.year) not in page_month_years:
            continue
        month_stated = bool(page_month_years)

        table_index = -1
        for block in _find_table_blocks_in_sheet(sheet_df, row_has_border):
            for header, data_rows in _split_block_into_minitables(block):
                table_index += 1

                day_entries = _rows_for_day(header, data_rows, target.day)
                if not day_entries:
                    continue

                for entry in day_entries:
                    results.append({
                        "source_file": sheet_name,
                        "table_index": table_index,
                        "header": header,
                        "date_column": header[_matching_day_column(header, target.day)],
                        "value": entry["value"],
                        "label": entry["label"],
                        "is_ratio": entry["is_ratio"],
                        "misaligned": False,
                        "month_stated": month_stated,
                    })

    return results


def hospital_date_column_from_xlsx(
    hospital: str,
    date: Union[str, date_cls],
    xlsx_path: Path,
    sheet_name: str | None = None,
) -> list[dict]:
    """Convenience wrapper around get_hospital_date_rows_from_xlsx: pulls
    out one {entries, ratio, warnings} result per sheet that matches
    hospital + date.

    Unlike hospital_date_column() (the Markdown version), this never
    raises just because more than one sheet matches — real workbooks often
    carry more than one prompt's output for the same page while you're
    still comparing prompts (e.g. "page_046" and "page_046A"), so this
    just returns one result per sheet, in sheet-name order.

    sheet_name: optional filter to see just one specific sheet.

    Raises ValueError only if nothing at all matches (including after the
    sheet_name filter, if given).
    """
    rows = get_hospital_date_rows_from_xlsx(hospital, date, xlsx_path)

    if sheet_name:
        rows = [r for r in rows if r["source_file"] == sheet_name]

    if not rows:
        extra = f", sheet_name={sheet_name!r}" if sheet_name else ""
        raise ValueError(f"No data found for hospital={hospital!r}, date={date!r}{extra} in {xlsx_path}")

    return _group_rows_by_source(rows, date)


# ---------------------------------------------------------------------------
# Deterministic, directory-based hospital resolution (recommended)
# ---------------------------------------------------------------------------

def resolve_hospital_xlsx_files(hospital: str, xlsx_dir: Path | None = None) -> list[Path]:
    """Resolve `hospital` directly to the .xlsx workbook(s) that belong to
    it — no scanning of page text or sheet names involved at all.

    Looks for, in order:
      1. Every *.xlsx file directly inside xlsx_dir/<hospital>/ (default
         xlsx_dir: config.XLSX_DIR — so config.xlsx_dir_for(hospital)) —
         the layout to_xlsx.py writes into by default.
      2. A single legacy file at xlsx_dir/<hospital>.xlsx, for a workbook
         saved directly rather than into its own subfolder.

    Raises FileNotFoundError (naming both paths it checked) if neither
    exists, rather than silently falling back to some unrelated workbook.
    """
    xlsx_dir = Path(xlsx_dir) if xlsx_dir is not None else config.XLSX_DIR
    hospital_dir = xlsx_dir / hospital
    if hospital_dir.is_dir():
        files = sorted(hospital_dir.glob("*.xlsx"))
        if files:
            return files

    legacy_file = xlsx_dir / f"{hospital}.xlsx"
    if legacy_file.exists():
        return [legacy_file]

    raise FileNotFoundError(
        f"No xlsx found for hospital={hospital!r}. Looked for *.xlsx under "
        f"{hospital_dir} and for a single file at {legacy_file}. Run "
        f"`python to_xlsx.py --facility {hospital}` first, or pass an "
        f"explicit xlsx path as lookup.py's third argument to look elsewhere."
    )


def hospital_date_lookup(
    hospital: str,
    date: Union[str, date_cls],
    xlsx_dir: Path | None = None,
) -> list[dict]:
    """Recommended entry point: resolve `hospital` to its own xlsx
    directory (see resolve_hospital_xlsx_files) and return one
    {entries, ratio, warnings} result per matching sheet, across every
    workbook found there.

    No hospital-name text/sheet-name scanning happens here — every sheet
    resolve_hospital_xlsx_files finds already belongs to `hospital`,
    purely by virtue of being saved under its own directory. That's what
    makes this reliable even for pages that never mention their hospital
    anywhere in the transcribed text.

    If more than one workbook turns up in the hospital's directory (e.g.
    output saved separately per model), each result's "source_file" is
    prefixed with that workbook's name ("rota_transcripts::page_046") so
    sheets from different files stay distinguishable; with the normal
    single-workbook layout it's just the plain sheet name.

    Raises FileNotFoundError if `hospital` doesn't resolve to any xlsx at
    all, or ValueError if it resolves to real file(s) but none of their
    sheets have a column for `date`.
    """
    xlsx_paths = resolve_hospital_xlsx_files(hospital, xlsx_dir)

    combined_rows: list[dict] = []
    for xlsx_path in xlsx_paths:
        # hospital="" makes _hospital_mentioned(...) match unconditionally
        # (an empty string is a substring of anything), which turns off
        # that text/sheet-name filtering entirely — correct here, since
        # the directory lookup above already guarantees every sheet in
        # this file belongs to `hospital`.
        rows = get_hospital_date_rows_from_xlsx("", date, xlsx_path)
        if len(xlsx_paths) > 1:
            for r in rows:
                r["source_file"] = f"{xlsx_path.stem}::{r['source_file']}"
        combined_rows.extend(rows)

    if not combined_rows:
        raise ValueError(
            f"hospital={hospital!r} resolved to {[str(p) for p in xlsx_paths]}, but no "
            f"sheet in {'it' if len(xlsx_paths) == 1 else 'them'} has a column for date={date!r}."
        )

    return _group_rows_by_source(combined_rows, date)


def resolve_hospital_transcripts_dir(hospital: str, transcripts_dir: Path | None = None) -> Path:
    """Resolve `hospital` to its own Markdown transcripts folder — the
    Markdown-side counterpart of resolve_hospital_xlsx_files, for looking
    things up before to_xlsx.py has been run yet.

    Default transcripts_dir is config.TRANSCRIPTS_DIR (the active model's
    folder), so the result is config.transcripts_dir_for(hospital) unless
    a different base folder is passed. Raises FileNotFoundError if that
    folder doesn't exist.
    """
    transcripts_dir = Path(transcripts_dir) if transcripts_dir is not None else config.TRANSCRIPTS_DIR
    hospital_dir = transcripts_dir / hospital
    if not hospital_dir.is_dir():
        raise FileNotFoundError(
            f"No transcripts folder for hospital={hospital!r} at {hospital_dir}. "
            f"Run main.py with --facility {hospital} first, or pass an explicit "
            f"transcripts_dir to look elsewhere."
        )
    return hospital_dir


def markdown_hospital_date_lookup(
    hospital: str,
    date: Union[str, date_cls],
    transcripts_dir: Path | None = None,
) -> list[dict]:
    """Markdown-side counterpart of hospital_date_lookup: resolves
    `hospital` directly to its own transcripts folder (see
    resolve_hospital_transcripts_dir) and returns one
    {entries, ratio, warnings} result per matching .md file — no
    hospital-name text/path scanning involved.

    Prefer hospital_date_lookup (the xlsx version) once to_xlsx.py has
    been run for this hospital — it gets the benefit of to_xlsx.py's own
    format-tolerant parsing. Use this one for a quick check straight off
    freshly transcribed Markdown, before converting.
    """
    hospital_dir = resolve_hospital_transcripts_dir(hospital, transcripts_dir)

    # hospital="" turns off the (redundant, once directory-scoped) text/
    # path filtering inside get_hospital_date_rows — see hospital_date_lookup.
    rows = get_hospital_date_rows("", date, hospital_dir)
    if not rows:
        raise ValueError(
            f"hospital={hospital!r} resolved to {hospital_dir}, but no page there "
            f"has a column for date={date!r}."
        )

    return _group_rows_by_source(rows, date)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    cli_parser = argparse.ArgumentParser(
        description="Look up rota data by hospital + date.",
        epilog=(
            "Examples:\n"
            "  python lookup.py Kakamega 2022-01-10\n"
            "      Deterministic lookup in xlsx/Kakamega/ (recommended — see hospital_date_lookup).\n"
            "      No page-text scanning: the hospital argument maps straight to that directory.\n"
            "\n"
            "  python lookup.py Kakamega 2022-01-10 --md\n"
            "      Same, but straight off transcripts/<model>/Kakamega/*.md — useful before\n"
            "      running to_xlsx.py yet.\n"
            "\n"
            "  python lookup.py Kakamega 2022-01-10 xlsx/rota_transcripts.xlsx\n"
            "      Legacy: an explicit workbook not organized per hospital — hospital is\n"
            "      matched by scanning each sheet's name/text instead of by directory.\n"
            "\n"
            "  python lookup.py Kakamega 2022-01-10 base2\n"
            "      Legacy: raw Markdown transcripts (config.TRANSCRIPTS_DIR), hospital matched\n"
            "      by scanning page text/paths, filtered to prompt_name \"base2\" (page_NNN_base2.md)."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    cli_parser.add_argument("hospital", help="Hospital tag, e.g. 'Kakamega' — same value passed to main.py's --facility.")
    cli_parser.add_argument("date", help="Date to look up, YYYY-MM-DD.")
    cli_parser.add_argument("third", nargs="?", default=None,
                             help="Optional: an explicit .xlsx path, or a prompt_name filter for the legacy Markdown path.")
    cli_parser.add_argument("--md", action="store_true",
                             help="Use the deterministic Markdown-directory lookup instead of the default xlsx one.")
    cli_args = cli_parser.parse_args()

    hospital_arg, date_arg, third_arg = cli_args.hospital, cli_args.date, cli_args.third

    def print_one_result(result):
        print(f"{hospital_arg} — {result['date_column']} ({date_arg})  [{result['source_file']}]")
        for entry in result["entries"]:
            flag = "  [!]" if entry["misaligned"] else ""
            print(f"  {entry['label']:<10} {entry['value']!r}{flag}")
        print(f"  {'RATIO':<10} {result['ratio']!r}")
        for w in result["warnings"]:
            print(f"[WARNING] {w}")

    try:
        if third_arg and third_arg.lower().endswith(".xlsx"):
            # Legacy path: an explicit workbook, hospital matched by scanning
            # sheet names/text.
            results = hospital_date_column_from_xlsx(hospital_arg, date_arg, third_arg)
        elif cli_args.md:
            if third_arg:
                raise SystemExit(
                    "--md looks up the hospital's whole transcripts folder directly — it "
                    "doesn't take a prompt_name/xlsx third argument."
                )
            results = markdown_hospital_date_lookup(hospital_arg, date_arg)
        elif third_arg:
            # Legacy path: raw Markdown, hospital matched by scanning page
            # text/paths, filtered down to one prompt's output.
            results = [hospital_date_column(hospital_arg, date_arg, config.TRANSCRIPTS_DIR, prompt_name=third_arg)]
        else:
            # Recommended default: deterministic, directory-based xlsx lookup.
            results = hospital_date_lookup(hospital_arg, date_arg)
    except (ValueError, FileNotFoundError) as e:
        raise SystemExit(str(e))

    first = True
    for result in results:
        if not first:
            print()
        first = False
        print(f"=== {result['source_file']} ===")
        print_one_result(result)