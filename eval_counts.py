"""
Look up E/M/N/total shift-code counts by hospital + date

Usage as a script (mirrors lookup.py's hospital/date/path CLI shape):
    python eval_counts.py Kakamega 2022-01-10 ROTA_NBU_Nov_2024.xlsx
    python eval_counts.py Kakamega 2022-01-10 ROTA_NBU_Nov_2024.xlsx --sheet in
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date as date_cls
from pathlib import Path
from typing import Union

import pandas as pd
from openpyxl.utils import column_index_from_string

from lookup import _as_date  # reuse the same "ISO string or date" acceptance lookup.py uses

METRICS = ("E-Evening: ", "M-Afternoon: ", "N-Night: ", "Total nurses: ")


@dataclass
class GroundTruthColumns:
    """Column letters in the ground-truth workbook. Hospital=A,
    Date=H, K="Nurses listed on rota" (total), N="No_Morning (Evening)
    shift" (E), P="No_ Afternoon (Morning off) shift" (M),
    R="No_Night Shift" (N).
    """
    hospital: str = "A"
    date: str = "H"
    total: str = "K"
    E: str = "N"
    M: str = "P"
    N: str = "R"
    header_row: int = 1  # which row (1 = the first row, same as Excel) the column titles are on


def _column_name_at(df: pd.DataFrame, letter: str) -> str:
    """Turn a spreadsheet column letter ("K", "AB", ...) into
    DataFrame column's actual name, e.g. "K" -> "Nurses listed on rota".
    """
    position = column_index_from_string(letter) - 1  # Excel columns start at 1, pandas at 0
    return df.columns[position]


def load_ground_truth(
    path: Path,
    hospital: str,
    sheet_name: str | None = None,
    columns: GroundTruthColumns | None = None,
) -> dict[date_cls, dict[str, int]]:
    """Read every (hospital, date) row from the ground-truth workbook and
    return {date: {"E": n, "M": n, "N": n, "total": n}} for rows whose
    Hospital column matches `hospital` (case-insensitive, exact).

    """
    columns = columns or GroundTruthColumns()

    # Step 1: read the whole sheet into a table. header counts rows from 0,
    # so header_row=1 (the first row, Excel-style) becomes header=0 here.
    df = pd.read_excel(path, sheet_name=sheet_name or 0, header=columns.header_row - 1)

    # Step 2: pull out just the six columns we care about, and give them
    # plain names — everything below reads "E"/"M"/"N"/"total" instead of
    # remembering which spreadsheet letter meant what.
    table = pd.DataFrame({
        "hospital": df[_column_name_at(df, columns.hospital)],
        "date": pd.to_datetime(df[_column_name_at(df, columns.date)], errors="coerce"),
        "E-Evening: ": df[_column_name_at(df, columns.E)],
        "M-Afternoon: ": df[_column_name_at(df, columns.M)],
        "N-Night: ": df[_column_name_at(df, columns.N)],
        "Total nurses: ": df[_column_name_at(df, columns.total)],
    })

    # Step 3: keep only rows for this hospital (case-insensitive, exact
    # match) with a real date — a blank Hospital/Date cell becomes NaN/NaT,
    # which this drops rather than trying to compare against.
    table = table.dropna(subset=["date"])
    is_this_hospital = table["hospital"].astype(str).str.strip().str.lower() == hospital.strip().lower()
    table = table[is_this_hospital]

    # Step 4: blank count cells (NaN) mean "0 of that shift", not "unknown".
    table[["E-Evening: ", "M-Afternoon: ", "N-Night: ", "Total nurses: "]] = table[["E-Evening: ", "M-Afternoon: ", "N-Night: ", "Total nurses: "]].fillna(0).astype(int)

    # Step 5: turn the table into {date: {"E": n, "M": n, "N": n, "total": n}}.
    result: dict[date_cls, dict[str, int]] = {}
    for _, row in table.iterrows():
        result[row["date"].date()] = {
            "E-Evening: ": int(row["E-Evening: "]),
            "M-Afternoon: ": int(row["M-Afternoon: "]),
            "N-Night: ": int(row["N-Night: "]),
            "Total nurses: ": int(row["Total nurses: "]),
        }
    return result


def ground_truth_date_counts(
    hospital: str,
    date: Union[str, date_cls],
    ground_truth_path: Path,
    sheet_name: str | None = None,
    columns: GroundTruthColumns | None = None,
) -> dict[str, int]:
    """The ground-truth-side counterpart of lookup.py's
    hospital_date_lookup(): given a hospital + date, return
    {"E-Evening: ": n, "M-Afternoon: ": n, "N-Night: ": n, "Total nurses: ": n} straight from the ground-truth
    workbook's own row for that (hospital, date)

    Raises ValueError if no row matches (hospital, date) in the workbook —
    same "don't silently return zeros for missing data" stance the rest of
    this project takes elsewhere (see lookup.py's hospital_date_lookup).
    """
    target = _as_date(date)
    rows_by_date = load_ground_truth(ground_truth_path, hospital, sheet_name=sheet_name, columns=columns)

    if target not in rows_by_date:
        raise ValueError(
            f"No ground-truth row for hospital={hospital!r}, date={target!r} in {ground_truth_path}. "
            f"Check the hospital spelling matches the workbook's Hospital column exactly, and that "
            f"the date actually has a row there."
        )

    return rows_by_date[target]


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Look up one hospital + date's E/M/N/total counts in a ground-truth workbook.",
    )
    parser.add_argument("hospital", help="Hospital name to match in the ground-truth workbook's Hospital column, e.g. 'Kakamega'.")
    parser.add_argument("date", help="Date to look up, YYYY-MM-DD.")
    parser.add_argument("ground_truth", type=Path, help="Ground-truth workbook, e.g. ROTA_NBU_Nov_2024.xlsx")
    parser.add_argument("--sheet", default=None, help="Ground-truth sheet name (default: first sheet)")
    args = parser.parse_args()

    try:
        counts = ground_truth_date_counts(args.hospital, args.date, args.ground_truth, sheet_name=args.sheet)
    except ValueError as e:
        raise SystemExit(str(e))

    print(f"{args.hospital} — {args.date}  [ground truth: {args.ground_truth.name}]")
    for metric in METRICS:
        label = "Total nurses" if metric == "Total nurses" else metric
        print(f"  {label:<10} {counts[metric]}")