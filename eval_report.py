"""
Write eval_counts.evaluate()'s results out as a styled, formula-driven .xlsx
report. Kept separate from eval_counts.py so that module stays pure
computation with no openpyxl-rendering concerns mixed in.

Every "match" cell is a real formula comparing the ground-truth cell to the
predicted cell in the same row (not a value computed in Python and pasted
in) so the report keeps recalculating correctly if a reader manually
corrects a cell to investigate a mismatch. The summary block at the bottom
is formulas over the data rows for the same reason.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from eval_counts import METRICS, DateResult

FONT_NAME = "Arial"
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
MISMATCH_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
NA_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Column layout: Date, then (GT, Predicted, Match) for each metric, then
# sheets-used / notes. Kept as a list of (header, kind) so the row-writer
# and the formula-writer agree on exactly where everything landed.
COLUMNS = [("Date", "date")]
for metric in METRICS:
    label = "Total" if metric == "total" else metric
    COLUMNS.append((f"GT {label}", "gt"))
    COLUMNS.append((f"Pred {label}", "pred"))
    COLUMNS.append((f"{label} Match", "match"))
COLUMNS.append(("Exact Match", "exact"))
COLUMNS.append(("Sheets Used", "sheets"))
COLUMNS.append(("Notes", "notes"))


def _col(index: int) -> str:
    return get_column_letter(index)


def write_report(results: list[DateResult], summary: dict, output_path: Path, hospital: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"

    # --- header row ---------------------------------------------------
    for j, (label, _kind) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=j, value=label)
        cell.font = Font(name=FONT_NAME, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    ws.freeze_panes = "A2"

    # Map each metric to the (gt_col, pred_col, match_col) triple, and note
    # where the Exact Match column ended up — the formulas below need both.
    metric_cols = {}
    col_idx = 2
    for metric in METRICS:
        metric_cols[metric] = (col_idx, col_idx + 1, col_idx + 2)
        col_idx += 3
    exact_col = col_idx
    sheets_col = exact_col + 1
    notes_col = sheets_col + 1

    # --- one row per date ----------------------------------------------
    r = 2
    first_data_row = r
    for res in results:
        ws.cell(row=r, column=1, value=res.date.isoformat())

        for metric in METRICS:
            gt_c, pred_c, match_c = metric_cols[metric]
            gt_val = res.gt[metric] if res.gt is not None else None
            pred_val = res.predicted[metric] if res.predicted["found"] else None
            if gt_val is not None:
                ws.cell(row=r, column=gt_c, value=gt_val)
            if pred_val is not None:
                ws.cell(row=r, column=pred_c, value=pred_val)

            gt_ref = f"{_col(gt_c)}{r}"
            pred_ref = f"{_col(pred_c)}{r}"
            match_cell = ws.cell(
                row=r, column=match_c,
                value=f'=IF(OR({gt_ref}="",{pred_ref}=""),"N/A",IF({gt_ref}={pred_ref},TRUE,FALSE))',
            )
            match_cell.alignment = CENTER

        match_cell_refs = [f"{_col(metric_cols[m][2])}{r}" for m in METRICS]
        match_refs = ",".join(match_cell_refs)
        na_checks = ",".join(f'{ref}="N/A"' for ref in match_cell_refs)
        exact_formula = f'=IF(OR({na_checks}),"N/A",AND({match_refs}))'
        exact_cell = ws.cell(row=r, column=exact_col, value=exact_formula)
        exact_cell.alignment = CENTER

        ws.cell(row=r, column=sheets_col, value=", ".join(res.predicted.get("sheets") or []))
        notes_cell = ws.cell(row=r, column=notes_col, value="; ".join(res.notes))
        notes_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        r += 1
    last_data_row = r - 1

    # --- styling pass: font/border everywhere, red fill on FALSE / grey on N/A ---
    for row in ws.iter_rows(min_row=1, max_row=last_data_row, min_col=1, max_col=notes_col):
        for cell in row:
            if cell.font is None or cell.font.name != FONT_NAME:
                cell.font = Font(name=FONT_NAME, bold=(cell.row == 1))
            if cell.row > 1 and cell.column not in (notes_col,):
                cell.alignment = CENTER
            cell.border = BORDER

    # Conditional-style the match/exact columns without formulas re-deriving
    # anything: openpyxl conditional formatting reacts to the *displayed*
    # value, which is exactly the formula result computed above.
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    match_range_cols = [metric_cols[m][2] for m in METRICS] + [exact_col]
    for col in match_range_cols:
        col_letter = _col(col)
        rng = f"{col_letter}{first_data_row}:{col_letter}{last_data_row}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{col_letter}{first_data_row}=FALSE'], fill=MISMATCH_FILL))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"N/A"'], fill=NA_FILL))

    # --- summary block ---------------------------------------------------
    summary_row = last_data_row + 2
    ws.cell(row=summary_row, column=1, value="Summary").font = Font(name=FONT_NAME, bold=True, size=12)

    def add_summary_line(row_offset, label, formula):
        cell_label = ws.cell(row=summary_row + row_offset, column=1, value=label)
        cell_label.font = Font(name=FONT_NAME, bold=True)
        cell_val = ws.cell(row=summary_row + row_offset, column=2, value=formula)
        cell_val.font = Font(name=FONT_NAME)
        return cell_val

    exact_col_letter = _col(exact_col)
    exact_range = f"{exact_col_letter}{first_data_row}:{exact_col_letter}{last_data_row}"

    add_summary_line(1, "Hospital", hospital)
    total_cell = add_summary_line(2, "Total dates evaluated", f"=COUNTA(A{first_data_row}:A{last_data_row})")
    comparable_formula = f'=COUNTIF({exact_range},TRUE)+COUNTIF({exact_range},FALSE)'
    comparable_cell = add_summary_line(3, "Comparable dates (both sides have data)", comparable_formula)
    comparable_ref = f"B{summary_row + 3}"
    add_summary_line(4, "Missing ground truth or generated data", f"=B{summary_row + 2}-{comparable_ref}")
    exact_match_cell = add_summary_line(5, "Exact matches (E, M, N, and Total all agree)", f"=COUNTIF({exact_range},TRUE)")
    exact_match_ref = f"B{summary_row + 5}"
    rate_cell = add_summary_line(
        6, "Exact match rate",
        f'=IF({comparable_ref}=0,"N/A",{exact_match_ref}/{comparable_ref})',
    )
    rate_cell.number_format = "0.0%"

    line = 7
    for metric in METRICS:
        label = "Total" if metric == "total" else metric
        match_col_letter = _col(metric_cols[metric][2])
        match_range = f"{match_col_letter}{first_data_row}:{match_col_letter}{last_data_row}"
        cell = add_summary_line(
            line, f"{label} match rate",
            f'=IF({comparable_ref}=0,"N/A",COUNTIF({match_range},TRUE)/{comparable_ref})',
        )
        cell.number_format = "0.0%"
        line += 1

    # --- column widths ---------------------------------------------------
    ws.column_dimensions["A"].width = 12
    for metric in METRICS:
        gt_c, pred_c, match_c = metric_cols[metric]
        for c in (gt_c, pred_c, match_c):
            ws.column_dimensions[_col(c)].width = 10
    ws.column_dimensions[_col(exact_col)].width = 12
    ws.column_dimensions[_col(sheets_col)].width = 28
    ws.column_dimensions[_col(notes_col)].width = 44

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path