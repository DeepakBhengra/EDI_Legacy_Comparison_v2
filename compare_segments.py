#!/usr/bin/env python3
"""Compare GDL and Legacy EDI segment files and write an Excel report."""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict, deque
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

STATUS_MATCH = "Match"
STATUS_MISSING_IN_IMPULSE = "Missing in Impulse"
STATUS_MISSING_IN_GDL = "Missing in GDL"

FILL_YELLOW = "yellow"
FILL_RED = "red"


@dataclass(frozen=True)
class ReportRow:
    gdl_segment: str
    legacy_segment: str
    status: str
    fill: str | None = None


@dataclass(frozen=True)
class ComparisonResult:
    rows: list[ReportRow]
    gdl_line_count: int
    legacy_line_count: int

    @property
    def sequential_matches(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MATCH and row.fill is None)

    @property
    def out_of_order_matches(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MATCH and row.fill == FILL_YELLOW)

    @property
    def missing_in_impulse(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISSING_IN_IMPULSE)

    @property
    def missing_in_gdl(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISSING_IN_GDL)


def read_segment_lines(path: Path) -> list[str]:
    """Read a segment file as individual lines, preserving inner whitespace."""
    text = path.read_text(encoding="utf-8", errors="replace")
    return text.splitlines()


def compare_segments(gdl_lines: Iterable[str], legacy_lines: Iterable[str]) -> ComparisonResult:
    """Walk GDL lines against Legacy lines and build report rows.

    Matching rules:
    - If the current unused Legacy line equals the current GDL line, record Match.
    - Otherwise look ahead through remaining unused Legacy lines.
      If the GDL line is found later, record Match and mark the row yellow.
      If it is not found, record Missing in Impulse and mark the row red.
    - Unused Legacy lines that never matched any GDL line are appended as
      Missing in GDL and marked red.
    """
    gdl = list(gdl_lines)
    legacy = list(legacy_lines)

    remaining: dict[str, deque[int]] = defaultdict(deque)
    for index, line in enumerate(legacy):
        remaining[line].append(index)

    used = [False] * len(legacy)
    next_unused = 0
    assignments: list[tuple[int | None, bool]] = []

    def advance_next_unused() -> None:
        nonlocal next_unused
        while next_unused < len(legacy) and used[next_unused]:
            next_unused += 1

    for gdl_line in gdl:
        advance_next_unused()
        candidates = remaining.get(gdl_line)
        matched_index: int | None = None
        out_of_order = False

        if candidates:
            earliest = candidates[0]
            if earliest == next_unused:
                matched_index = earliest
            elif earliest > next_unused:
                matched_index = earliest
                out_of_order = True

        if matched_index is not None:
            used[matched_index] = True
            candidates.popleft()
            assignments.append((matched_index, out_of_order))
        else:
            assignments.append((None, False))

    rows: list[ReportRow] = []
    for gdl_line, (legacy_index, out_of_order) in zip(gdl, assignments):
        if legacy_index is None:
            rows.append(
                ReportRow(
                    gdl_segment=gdl_line,
                    legacy_segment="",
                    status=STATUS_MISSING_IN_IMPULSE,
                    fill=FILL_RED,
                )
            )
        else:
            rows.append(
                ReportRow(
                    gdl_segment=gdl_line,
                    legacy_segment=legacy[legacy_index],
                    status=STATUS_MATCH,
                    fill=FILL_YELLOW if out_of_order else None,
                )
            )

    for index, line in enumerate(legacy):
        if not used[index]:
            rows.append(
                ReportRow(
                    gdl_segment="",
                    legacy_segment=line,
                    status=STATUS_MISSING_IN_GDL,
                    fill=FILL_RED,
                )
            )

    return ComparisonResult(rows=rows, gdl_line_count=len(gdl), legacy_line_count=len(legacy))


def write_excel_report(result: ComparisonResult, output_path: Path) -> None:
    """Write the comparison rows to an .xlsx workbook."""
    try:
        import xlsxwriter
    except ImportError as exc:  # pragma: no cover - exercised in runtime setups
        raise SystemExit(
            "xlsxwriter is required. Install dependencies with: pip install -r requirements.txt"
        ) from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(output_path))

    header_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#1F4E79",
            "font_color": "#FFFFFF",
            "border": 1,
            "valign": "vcenter",
            "align": "center",
        }
    )
    match_fmt = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
    yellow_fmt = workbook.add_format(
        {"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#FFFF00"}
    )
    red_fmt = workbook.add_format(
        {
            "border": 1,
            "valign": "top",
            "text_wrap": True,
            "bg_color": "#FF0000",
            "font_color": "#FFFFFF",
        }
    )
    label_fmt = workbook.add_format({"bold": True, "border": 1, "bg_color": "#D9E2F3"})
    value_fmt = workbook.add_format({"border": 1})
    yellow_value_fmt = workbook.add_format({"border": 1, "bg_color": "#FFFF00"})
    red_value_fmt = workbook.add_format(
        {"border": 1, "bg_color": "#FF0000", "font_color": "#FFFFFF"}
    )

    detail = workbook.add_worksheet("Comparison")
    detail.write(0, 0, "GDL Segment", header_fmt)
    detail.write(0, 1, "Legacy Segment", header_fmt)
    detail.write(0, 2, "status", header_fmt)
    detail.freeze_panes(1, 0)
    detail.set_column(0, 1, 85)
    detail.set_column(2, 2, 24)
    detail.set_row(0, 22)
    detail.autofilter(0, 0, max(len(result.rows), 1), 2)

    for row_number, row in enumerate(result.rows, start=1):
        if row.fill == FILL_YELLOW:
            cell_fmt = yellow_fmt
        elif row.fill == FILL_RED:
            cell_fmt = red_fmt
        else:
            cell_fmt = match_fmt
        detail.write(row_number, 0, row.gdl_segment, cell_fmt)
        detail.write(row_number, 1, row.legacy_segment, cell_fmt)
        detail.write(row_number, 2, row.status, cell_fmt)

    summary = workbook.add_worksheet("Summary")
    summary.set_column(0, 0, 36)
    summary.set_column(1, 1, 18)
    summary_rows = [
        ("GDL lines", result.gdl_line_count, value_fmt),
        ("Legacy lines", result.legacy_line_count, value_fmt),
        ("Sequential Match", result.sequential_matches, value_fmt),
        ("Out-of-order Match (yellow)", result.out_of_order_matches, yellow_value_fmt),
        ("Missing in Impulse", result.missing_in_impulse, red_value_fmt),
        ("Missing in GDL", result.missing_in_gdl, red_value_fmt),
        ("Total report rows", len(result.rows), value_fmt),
    ]
    summary.write(0, 0, "Metric", header_fmt)
    summary.write(0, 1, "Count", header_fmt)
    for index, (label, value, fmt) in enumerate(summary_rows, start=1):
        summary.write(index, 0, label, label_fmt)
        summary.write(index, 1, value, fmt)

    workbook.close()


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    repo_root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Compare GDL and Legacy segment files and write an Excel report."
    )
    parser.add_argument(
        "--gdl",
        type=Path,
        default=repo_root / "data" / "GDL_Segment.txt",
        help="Path to the GDL segment file.",
    )
    parser.add_argument(
        "--legacy",
        type=Path,
        default=repo_root / "data" / "Legacy_Segment.txt",
        help="Path to the Legacy / Impulse segment file.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=repo_root / "reports" / "segment_comparison.xlsx",
        help="Path to the Excel report to create.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if not args.gdl.is_file():
        print(f"GDL segment file not found: {args.gdl}", file=sys.stderr)
        return 1
    if not args.legacy.is_file():
        print(f"Legacy segment file not found: {args.legacy}", file=sys.stderr)
        return 1

    print(f"Reading GDL segments from {args.gdl}")
    gdl_lines = read_segment_lines(args.gdl)
    print(f"Reading Legacy segments from {args.legacy}")
    legacy_lines = read_segment_lines(args.legacy)

    print(f"Comparing {len(gdl_lines):,} GDL lines with {len(legacy_lines):,} Legacy lines...")
    result = compare_segments(gdl_lines, legacy_lines)
    write_excel_report(result, args.output)

    print()
    print("Comparison complete")
    print(f"  Sequential Match:          {result.sequential_matches:,}")
    print(f"  Out-of-order Match:        {result.out_of_order_matches:,}")
    print(f"  Missing in Impulse:        {result.missing_in_impulse:,}")
    print(f"  Missing in GDL:            {result.missing_in_gdl:,}")
    print(f"  Report rows:               {len(result.rows):,}")
    print(f"  Excel report:              {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
