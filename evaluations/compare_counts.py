"""
Compare a generated rota workbook (to_xlsx.py output) against a trusted
ground-truth workbook, by comparing per-date shift-code counts.

For each date, two things are counted and compared:
  - E / M / N: how many staff rows have that exact shift code in that
    date's column.
  - total: how many staff rows have ANY non-blank value in that date's
    column (i.e. how many nurses are listed as on duty that day).

Usage as a script:
    python compare_counts.py ROTA_NBU_Nov_2024.xlsx rota_transcripts.xlsx "Kakamega" --out report.xlsx
    python compare_counts.py ROTA_NBU_Nov_2024.xlsx rota_transcripts.xlsx "Kakamega" --from 2022-01-01 --to 2022-01-31 --out report.xlsx
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from datetime import date as date_cls
from pathlib import Path
from typing import Union

from ground_truth_lookup import GroundTruthColumns, load_ground_truth, METRICS
from lookup import (  # reuse the already-tested table-reading logic
    _as_date,
    _sheet_text,
    _page_month_years,
    _load_workbook_sheets,
    _find_table_blocks_in_sheet,
    _split_block_into_minitables,
    get_hospital_date_rows_from_xlsx,
    LEADING_DAY_RE,
)


# ---------------------------------------------------------------------------
# Generated workbook
# ---------------------------------------------------------------------------

def covered_dates_in_generated_xlsx(xlsx_path: Path) -> list[date_cls]:
    """Every calendar date this workbook has at least one date-column for,
    across every sheet, table, and mini-table.

    A page that states more than one month (e.g. "DECEMBER 2021/ JANUARY
    2022") can produce a date that isn't really meant by that column (the
    same ambiguity get_hospital_date_rows_from_xlsx already lives with via
    its "month_stated" flag) — this is a convenience default for scanning
    a whole workbook without having to name every date by hand, not a
    precise calendar reconstruction. Pass explicit dates to evaluate() to
    sidestep it entirely.
    """
    # Same reasoning as lookup.py's own xlsx reading — see
    # _load_workbook_sheets: table boundaries need the cell borders
    # (openpyxl), everything else just needs the plain values (pandas).
    sheets = _load_workbook_sheets(xlsx_path)
    found: set[date_cls] = set()

    for sheet_df, row_has_border in sheets.values():
        month_years = _page_month_years(_sheet_text(sheet_df))
        if not month_years:
            continue

        for block in _find_table_blocks_in_sheet(sheet_df, row_has_border):
            for header, _data_rows in _split_block_into_minitables(block):
                for cell in header:
                    m = LEADING_DAY_RE.match(cell.strip())
                    if not m:
                        continue
                    day = int(m.group(1))
                    for month, year in month_years:
                        try:
                            found.add(date_cls(year, month, day))
                        except ValueError:
                            continue  # e.g. day 31 paired with a 30-day month

    return sorted(found)


def predicted_counts_for_date(
    generated_xlsx_path: Path,
    date: Union[str, date_cls],
) -> dict[str, object]:
    """Count E / M / N / total in the generated workbook's column(s) for
    `date`, across every sheet/table/mini-table that has one.

    hospital="" is passed to get_hospital_date_rows_from_xlsx deliberately
    — it disables that function's hospital-name text/sheet-name filtering
    entirely (an empty string is a substring of anything), which is
    correct here: the workbook passed in is already scoped to one hospital
    by construction (see the module docstring), so there is nothing left
    to filter by hospital.

    Returns:
        {"E": n, "M": n, "N": n, "total": n, "sheets": [...], "found": bool}
    "found" is False (and every count is None) if no sheet/table anywhere
    in the workbook has a column for this date at all — distinct from a
    date that does have a column but zero E/M/N codes in it.
    """
    rows = get_hospital_date_rows_from_xlsx("", date, generated_xlsx_path)

    if not rows:
        return {"E": None, "M": None, "N": None, "total": None, "sheets": [], "found": False}

    counts = {"E": 0, "M": 0, "N": 0, "total": 0}
    sheets = []
    for r in rows:
        if r["source_file"] not in sheets:
            sheets.append(r["source_file"])
        if r["is_ratio"]:
            continue  # the RATIO/summary row isn't an individual staff entry
        value = r["value"].strip()
        if not value:
            continue
        counts["total"] += 1
        code = value.upper()
        if code in ("E", "M", "N"):
            counts[code] += 1

    counts["sheets"] = sheets
    counts["found"] = True
    return counts


def predicted_counts_from_lookup_results(results: list[dict]) -> dict[str, object]:
    """Aggregate E/M/N/total counts directly from lookup.py's own result
    shape, instead of re-reading a workbook: the same
    {"source_file", "date_column", "entries": [{"label","value",...}, ...],
    "ratio", "warnings"} dicts hospital_date_lookup(),
    markdown_hospital_date_lookup(), hospital_date_column(), and
    hospital_date_column_from_xlsx() all return.

    This is what lets lookup.py's CLI print a ground-truth comparison for
    a single date right after doing the lookup itself, no matter which of
    those four lookup paths (xlsx-directory, markdown-directory, legacy
    xlsx, legacy markdown) produced `results` — every path already
    separates the RATIO row out into "ratio" rather than leaving it in
    "entries", so nothing needs to be filtered out here the way
    predicted_counts_for_date filters out is_ratio rows.

    Counts are summed across every item in `results` (i.e. every matching
    sheet/file), same as predicted_counts_for_date sums across every
    matching sheet in one workbook. "found" is False only if `results`
    itself is empty (the lookup found nothing at all for that date).
    """
    counts = {"E": 0, "M": 0, "N": 0, "total": 0}
    sheets = []
    for result in results:
        source = result.get("source_file")
        if source is not None:
            sheets.append(source)
        for entry in result["entries"]:
            value = (entry["value"] or "").strip()
            if not value:
                continue
            counts["total"] += 1
            code = value.upper()
            if code in ("E", "M", "N"):
                counts[code] += 1

    counts["sheets"] = sheets
    counts["found"] = bool(results)
    return counts


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

@dataclass
class DateResult:
    date: date_cls
    gt: dict[str, int] | None
    predicted: dict[str, object]
    matches: dict[str, bool] = field(default_factory=dict)
    exact_match: bool = False
    notes: list[str] = field(default_factory=list)


def evaluate(
    ground_truth_path: Path,
    generated_xlsx_path: Path,
    hospital: str,
    dates: list[Union[str, date_cls]] | None = None,
    gt_sheet_name: str | None = None,
    gt_columns: GroundTruthColumns | None = None,
) -> list[DateResult]:
    """Compare ground truth vs. generated E/M/N/total counts, one date at
    a time.

    dates: explicit dates to check. Defaults to every date the generated
        workbook has at least one column for (see
        covered_dates_in_generated_xlsx) — the natural default, since
        there's rarely a reason to score a date the generated workbook
        never transcribed in the first place. Ground truth's own date
        range is typically far wider (e.g. years, across every hospital
        it tracks) and isn't a useful default on its own.

    Each DateResult's `notes` explains anything that kept a date from
    being a clean match/mismatch: no ground-truth row for that date, or
    the generated workbook has no column for it at all (never silently
    treated as "0 == 0").
    """
    gt_lookup = load_ground_truth(ground_truth_path, hospital, sheet_name=gt_sheet_name, columns=gt_columns)

    if dates is None:
        target_dates = covered_dates_in_generated_xlsx(generated_xlsx_path)
    else:
        target_dates = [_as_date(d) for d in dates]

    results = []
    for d in target_dates:
        gt = gt_lookup.get(d)
        predicted = predicted_counts_for_date(generated_xlsx_path, d)

        notes = []
        matches = {}
        if gt is None:
            notes.append(f"No ground-truth row for hospital={hospital!r}, date={d!r}.")
        if not predicted["found"]:
            notes.append(f"Generated workbook has no date-column for {d!r} anywhere.")

        if gt is not None and predicted["found"]:
            for metric in METRICS:
                matches[metric] = (gt[metric] == predicted[metric])
            exact_match = all(matches.values())
        else:
            for metric in METRICS:
                matches[metric] = False
            exact_match = False

        results.append(DateResult(date=d, gt=gt, predicted=predicted, matches=matches,
                                   exact_match=exact_match, notes=notes))

    return results


def summarize(results: list[DateResult]) -> dict:
    """Overall accuracy: exact-match rate over dates that had both a
    ground-truth row and generated data (i.e. were actually comparable),
    plus per-metric accuracy the same way. Dates missing one side entirely
    are counted separately (not folded silently into "wrong").
    """
    comparable = [r for r in results if r.gt is not None and r.predicted["found"]]
    missing = [r for r in results if r not in comparable]

    summary = {
        "total_dates": len(results),
        "comparable_dates": len(comparable),
        "missing_dates": len(missing),
        "exact_match_count": sum(1 for r in comparable if r.exact_match),
    }
    summary["exact_match_rate"] = (
        summary["exact_match_count"] / summary["comparable_dates"] if summary["comparable_dates"] else None
    )
    for metric in METRICS:
        matched = sum(1 for r in comparable if r.matches[metric])
        summary[f"{metric}_match_rate"] = (matched / len(comparable)) if comparable else None
    return summary


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _parse_date_arg(s: str) -> date_cls:
    return _as_date(s)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("ground_truth", type=Path, help="Ground-truth workbook, e.g. ROTA_NBU_Nov_2024.xlsx")
    parser.add_argument("generated", type=Path, help="Generated workbook, e.g. rota_transcripts.xlsx")
    parser.add_argument("hospital", help="Hospital name to match in the ground-truth workbook's Hospital column, e.g. 'Kakamega'")
    parser.add_argument("--from", dest="date_from", type=_parse_date_arg, default=None)
    parser.add_argument("--to", dest="date_to", type=_parse_date_arg, default=None)
    parser.add_argument("--gt-sheet", default=None, help="Ground-truth sheet name (default: first sheet)")
    parser.add_argument("--out", type=Path, default=None, help="Write a detailed .xlsx report here")
    args = parser.parse_args()

    dates_arg = None
    if args.date_from or args.date_to:
        covered = covered_dates_in_generated_xlsx(args.generated)
        dates_arg = [d for d in covered
                     if (args.date_from is None or d >= args.date_from)
                     and (args.date_to is None or d <= args.date_to)]

    all_results = evaluate(args.ground_truth, args.generated, args.hospital, dates=dates_arg, gt_sheet_name=args.gt_sheet)

    for r in all_results:
        status = "MATCH" if r.exact_match else "MISMATCH"
        gt_str = r.gt if r.gt is not None else "(no ground truth)"
        pred_str = {k: r.predicted[k] for k in METRICS} if r.predicted["found"] else "(no data)"
        print(f"{r.date}  [{status}]  gt={gt_str}  predicted={pred_str}")
        for note in r.notes:
            print(f"    - {note}")

    overall = summarize(all_results)
    print(f"\n{overall['comparable_dates']}/{overall['total_dates']} dates comparable "
          f"({overall['missing_dates']} missing ground truth or generated data)")
    if overall["exact_match_rate"] is not None:
        print(f"Exact match rate: {overall['exact_match_rate']:.1%}")
        for metric in METRICS:
            print(f"  {metric} match rate: {overall[f'{metric}_match_rate']:.1%}")

    if args.out:
        from eval_report import write_report  # local import: report-writing is openpyxl/formula-heavy, kept separate
        write_report(all_results, overall, args.out, hospital=args.hospital)
        print(f"\nReport -> {args.out}")