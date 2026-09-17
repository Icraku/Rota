"""
Post-transcription validator + corrector for main.py's pipeline.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable

# Ordinary nurse-shift vocabulary, exactly as defined in
# prompts/csv_direct.txt, compared case-insensitively. An empty/blank
# cell is valid too -- that's how "no mark" actually comes back in
# practice, not the literal string "[blank]".
VALID_SHIFT_CODES = {"E", "M", "N", "NO", "D", "DO", "PH", "", "[BLANK]", "[UNCLEAR]", "[COLUMN UNCLEAR]"}

# Rows whose label cells contain any of these (case-insensitive substring
# match) are special/free-text rows, per the prompt -- never checked
# against the shift-code vocabulary.
SPECIAL_ROW_MARKERS = ("dep i/c", "i/c", "ratio", "support staff", "wc")

# Known handwriting confusions your own prompt already names -- surfaced
# explicitly in the retry prompt instead of a generic "look again".
KNOWN_CONFUSIONS = {
    "Z": "N (a Z-shaped mark in an ordinary shift cell is usually N, per the transcription guide)",
    "3": "M (an upside-down-3-shaped mark in an ordinary shift cell is usually M, per the transcription guide)",
}


@dataclass
class CellIssue:
    table_index: int
    row_index: int
    cell_index: int  # index into the row's raw cell list; -1 for malformed_row issues
    date_column: str
    value: str
    row_label: str
    kind: str  # "invalid_code" or "malformed_row"


@dataclass
class ParsedTable:
    header_cells: list[str]
    date_columns: list[str]       # header cells that look like date columns
    num_label_cols: int           # header cells before the date columns
    rows: list[list[str]]         # raw cell lists, one per data row (separator row excluded)


def _split_row(line: str) -> list[str]:
    """'| a | b | c |' -> ['a', 'b', 'c'], trimming the leading/trailing
    empty strings a leading/trailing '|' produces."""
    parts = [cell.strip() for cell in line.strip().split("|")]
    if parts and parts[0] == "":
        parts = parts[1:]
    if parts and parts[-1] == "":
        parts = parts[:-1]
    return parts


def _is_table_row(line: str) -> bool:
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|")


def _is_separator_row(cells: list[str]) -> bool:
    return bool(cells) and all(re.fullmatch(r":?-+:?", c) for c in cells)


def parse_markdown_tables(markdown: str) -> list[ParsedTable]:
    """Find every Markdown table in the transcription and parse it into
    header/date-columns/rows. Ignores anything that isn't a '|...|' line
    (headings, footer metadata, blank lines).

    num_label_cols is deliberately derived from the DATA rows' most common
    length, not from the header's own cell count -- in real output the
    model is fairly consistent about how many cells its data rows have,
    but not always consistent about keeping the header in sync with that
    (e.g. a header with NO/NAMES/7 dates = 9 columns, while every data row
    actually carries an extra unlabeled column for I/C-style tags = 10).
    Trusting the header here would flag almost every row as malformed
    over something that isn't really a per-row problem.
    """
    tables: list[ParsedTable] = []
    lines = markdown.splitlines()
    i = 0
    while i < len(lines):
        if _is_table_row(lines[i]):
            header_cells = _split_row(lines[i])
            i += 1
            if i < len(lines) and _is_table_row(lines[i]) and _is_separator_row(_split_row(lines[i])):
                i += 1  # skip the "|---|---|" separator row
            rows: list[list[str]] = []
            while i < len(lines) and _is_table_row(lines[i]):
                rows.append(_split_row(lines[i]))
                i += 1

            date_columns = [c for c in header_cells if "date" in c.lower()]
            num_label_cols = len(header_cells) - len(date_columns)  # fallback if there's no data to learn from
            if rows:
                row_lengths = [len(r) for r in rows]
                mode_length = max(set(row_lengths), key=row_lengths.count)
                if mode_length >= len(date_columns):
                    learned_label_cols = mode_length - len(date_columns)
                    if learned_label_cols != num_label_cols:
                        print(
                            f"  [validate_transcript] Table header implies {num_label_cols} label "
                            f"column(s), but {row_lengths.count(mode_length)}/{len(rows)} data rows "
                            f"actually have {learned_label_cols} -- using {learned_label_cols} "
                            f"(the header is very likely missing a column, not the data rows)."
                        )
                    num_label_cols = learned_label_cols

            tables.append(ParsedTable(header_cells, date_columns, num_label_cols, rows))
        else:
            i += 1
    return tables


def is_valid_shift_code(value: str) -> bool:
    return value.strip().upper() in VALID_SHIFT_CODES


def find_issues(tables: list[ParsedTable]) -> list[CellIssue]:
    """Check every ordinary-nurse-shift-row cell against the known
    vocabulary. Rows with the wrong cell count, or whose label cells mark
    them as a special row (DEP I/C, RATIO, ...), are not checked -- see
    the module docstring for why.
    """
    issues: list[CellIssue] = []
    for t_idx, table in enumerate(tables):
        expected_cols = table.num_label_cols + len(table.date_columns)
        for r_idx, row in enumerate(table.rows):
            if len(row) != expected_cols:
                issues.append(CellIssue(
                    t_idx, r_idx, cell_index=-1, date_column="(whole row)",
                    value=" | ".join(row), row_label=f"row {r_idx + 1}", kind="malformed_row",
                ))
                continue

            label_cells = row[:table.num_label_cols]
            date_cells = row[table.num_label_cols:]
            is_special_row = any(
                marker in label_cell.lower()
                for label_cell in label_cells
                for marker in SPECIAL_ROW_MARKERS
            )
            if is_special_row:
                continue

            row_label = " ".join(c for c in label_cells if c) or f"row {r_idx + 1}"
            for offset, (date_col, value) in enumerate(zip(table.date_columns, date_cells)):
                if not is_valid_shift_code(value):
                    issues.append(CellIssue(
                        t_idx, r_idx, cell_index=table.num_label_cols + offset,
                        date_column=date_col, value=value, row_label=row_label, kind="invalid_code",
                    ))
    return issues


def build_batch_retry_prompt(issues: list[CellIssue]) -> str:
    """One follow-up prompt covering every flagged cell on this page."""
    lines = [
        "You previously transcribed this whole rota page. The following cells "
        "did not come out as a valid ordinary nurse-shift code (E, M, N, No, D, "
        "Do, PH, or blank). Look at each one again, carefully, in the image.",
        "",
    ]
    for i, issue in enumerate(issues, start=1):
        hint = KNOWN_CONFUSIONS.get(issue.value.strip().upper())
        hint_text = f" (note: this shape is a known confusable -- it's very likely actually {hint})" if hint else ""
        lines.append(f"{i}. ROW {issue.row_label!r}, COLUMN {issue.date_column!r}: you wrote {issue.value!r}.{hint_text}")
    lines += [
        "",
        "Reply with EXACTLY ONE line per item above, in this exact format and "
        "nothing else -- no other words, no markdown, no explanation:",
        "ROW: <row label> | COLUMN: <column> | ANSWER: <one of E, M, N, No, D, Do, PH, UNCLEAR>",
    ]
    return "\n".join(lines)


_RESPONSE_LINE_RE = re.compile(r"ROW:\s*(.*?)\s*\|\s*COLUMN:\s*(.*?)\s*\|\s*ANSWER:\s*(.*?)\s*$", re.IGNORECASE)


def parse_batch_response(response: str) -> dict[tuple[str, str], str]:
    """Parse the model's batched reply into {(row_label, date_column): answer}.
    Lines that don't match the expected format are silently skipped -- the
    caller treats any issue with no matching line as unresolved.
    """
    parsed: dict[tuple[str, str], str] = {}
    for line in response.splitlines():
        match = _RESPONSE_LINE_RE.search(line)
        if match:
            row_label, date_column, answer = match.groups()
            parsed[(row_label.strip(), date_column.strip())] = answer.strip()
    return parsed


def correct_transcript(
    markdown: str,
    reread_page: Callable[[str], str],
) -> tuple[str, list[CellIssue], list[CellIssue]]:
    """Validate `markdown`'s ordinary-shift cells; if any are invalid, send
    ONE batched follow-up via `reread_page(prompt) -> model's raw reply`
    covering all of them, and apply whatever it resolves. Returns
    (corrected_markdown, resolved_issues, still_flagged_issues).

    Malformed rows (wrong cell count) are never auto-guessed -- they go
    straight into still_flagged_issues with the row's raw text.
    """
    tables = parse_markdown_tables(markdown)
    issues = find_issues(tables)

    malformed = [i for i in issues if i.kind == "malformed_row"]
    invalid_codes = [i for i in issues if i.kind == "invalid_code"]

    resolved: list[CellIssue] = []
    still_flagged: list[CellIssue] = list(malformed)
    replacements: dict[tuple[int, int, int], str] = {}  # (table_idx, row_idx, cell_idx) -> new_value

    if invalid_codes:
        batch_prompt = build_batch_retry_prompt(invalid_codes)
        response = reread_page(batch_prompt)
        parsed = parse_batch_response(response)

        for issue in invalid_codes:
            answer = parsed.get((issue.row_label, issue.date_column))
            if answer is None:
                new_value = f"[NEEDS REVIEW: {issue.value!r}, no answer for this cell in the retry response]"
                still_flagged.append(issue)
            elif answer.upper() == "UNCLEAR":
                new_value = f"[NEEDS REVIEW: model unsure, first read {issue.value!r}]"
                still_flagged.append(issue)
            elif is_valid_shift_code(answer):
                new_value = answer
                resolved.append(issue)
            else:
                new_value = f"[NEEDS REVIEW: {issue.value!r} then {answer!r}, neither is a known code]"
                still_flagged.append(issue)

            replacements[(issue.table_index, issue.row_index, issue.cell_index)] = new_value

    corrected = _apply_replacements(markdown, replacements)
    return corrected, resolved, still_flagged


def _apply_replacements(markdown: str, replacements: dict[tuple[int, int, int], str]) -> str:
    """Rewrite just the flagged cells back into the original markdown text,
    leaving everything else (headings, footer, well-formed cells) byte-
    for-byte untouched.
    """
    if not replacements:
        return markdown

    lines = markdown.splitlines()
    table_row_line_indices: list[list[int]] = []
    t_idx = -1
    i = 0
    while i < len(lines):
        if _is_table_row(lines[i]):
            t_idx += 1
            table_row_line_indices.append([])
            i += 1
            if i < len(lines) and _is_table_row(lines[i]) and _is_separator_row(_split_row(lines[i])):
                i += 1
            while i < len(lines) and _is_table_row(lines[i]):
                table_row_line_indices[t_idx].append(i)
                i += 1
        else:
            i += 1

    for (table_idx, row_idx, cell_idx), new_value in replacements.items():
        line_idx = table_row_line_indices[table_idx][row_idx]
        cells = _split_row(lines[line_idx])
        cells[cell_idx] = new_value
        lines[line_idx] = "| " + " | ".join(cells) + " |"

    return "\n".join(lines)