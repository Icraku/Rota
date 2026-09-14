"""
CLI: convert one hospital's Markdown transcripts into its xlsx workbook.

All the actual parsing/rendering logic lives in pipeline/c_md_to_xlsx.py —
this is just the command-line entry point, plus a re-export of the parsing
helpers so other scripts (lookup.py, older evaluations) can keep doing
`from to_xlsx import split_segments, ...` without caring that the
implementation lives under pipeline/.

Usage:
    python to_xlsx.py                                  # config.HOSPITAL, default dirs
    python to_xlsx.py --facility Kakamega
    python to_xlsx.py --facility Kakamega --transcripts-dir path/to/md --out path/to.xlsx

Output lands at xlsx/<hospital>/rota_transcripts.xlsx by default (see
config.xlsx_path_for), reading from transcripts/<model>/<hospital>/ by
default (see config.transcripts_dir_for) — matching the per-hospital layout
main.py writes into. lookup.py's default (directory-based) hospital lookup
expects exactly this layout, so leave --out/--transcripts-dir unset unless
you have a specific reason to put the workbook somewhere else.
"""
import argparse
from pathlib import Path

import config
from pipeline.c_md_to_xlsx import (  # noqa: F401 — re-exported for other scripts
    split_segments,
    parse_table_rows,
    build_header_and_data,
    normalize_cell_text,
    convert_transcript,
    sheet_name_for,
    convert_all,
    RATIO_VALUE_RE,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert one hospital's Markdown transcripts into an xlsx workbook.")
    parser.add_argument("--facility", "--hospital", dest="hospital", default=config.HOSPITAL,
                         help=f"Hospital tag (default: {config.HOSPITAL!r}) — selects transcripts/<model>/<hospital>/ and xlsx/<hospital>/.")
    parser.add_argument("--transcripts-dir", type=Path, default=None,
                         help="Override the input folder of .md transcripts (default: config.transcripts_dir_for(hospital)).")
    parser.add_argument("--out", type=Path, default=None,
                         help="Override the output .xlsx path (default: config.xlsx_path_for(hospital)).")
    args = parser.parse_args()

    transcripts_dir = args.transcripts_dir or config.transcripts_dir_for(args.hospital)
    out_path = args.out or config.xlsx_path_for(args.hospital)

    convert_all(transcripts_dir, out_path)
    print(f"Saved -> {out_path}")


if __name__ == "__main__":
    main()