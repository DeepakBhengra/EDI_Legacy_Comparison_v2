#!/usr/bin/env python3
"""Compare GDL and Legacy EDI segment files and write an Excel report."""

from __future__ import annotations

import argparse
import sys
from collections import Counter, defaultdict, deque
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

STATUS_MATCH = "Match"
STATUS_MISMATCH = "Mismatch"
STATUS_MISSING_IN_IMPULSE = "Missing in Impulse"
STATUS_MISSING_IN_GDL = "Missing in GDL"

FILL_YELLOW = "yellow"
FILL_ORANGE = "orange"
FILL_RED = "red"

# Envelope lines (ISA, GS, ST, BIA) are paired as Mismatch and skipped by matching.
HEADER_MISMATCH_COUNT = 4


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
    def mismatches(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISMATCH)

    @property
    def missing_in_impulse(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISSING_IN_IMPULSE)

    @property
    def missing_in_gdl(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISSING_IN_GDL)


def read_segment_lines(path: Path) -> list[str]:
    """Read a segment file as individual lines, preserving inner whitespace.

    Completely empty lines (common trailing export artifacts) are ignored.
    """
    text = path.read_text(encoding="utf-8", errors="replace")
    return [line for line in text.splitlines() if line != ""]


def compare_segments(gdl_lines: Iterable[str], legacy_lines: Iterable[str]) -> ComparisonResult:
    """Walk GDL lines against Legacy lines and build report rows.

    Matching rules:
    - The first four GDL lines and first four Legacy lines are written on matching
      rows as Mismatch and are not used in later matching.
    - If the current unused Legacy line equals the current GDL line, record Match.
    - Otherwise look ahead through remaining unused Legacy lines.
      If the GDL line is found later, record Match and mark the row yellow.
      If it is not found, record Missing in Impulse and mark the row red.
    - Unused Legacy lines that cannot match any remaining GDL line are written as
      Missing in GDL (red), either when they are skipped or after all GDL lines.
    """
    all_gdl = list(gdl_lines)
    all_legacy = list(legacy_lines)
    rows: list[ReportRow] = []

    header_count = min(HEADER_MISMATCH_COUNT, max(len(all_gdl), len(all_legacy)))
    for index in range(header_count):
        rows.append(
            ReportRow(
                gdl_segment=all_gdl[index] if index < len(all_gdl) else "",
                legacy_segment=all_legacy[index] if index < len(all_legacy) else "",
                status=STATUS_MISMATCH,
                fill=FILL_ORANGE,
            )
        )

    gdl = all_gdl[header_count:]
    legacy = all_legacy[header_count:]

    remaining_legacy: dict[str, deque[int]] = defaultdict(deque)
    for index, line in enumerate(legacy):
        remaining_legacy[line].append(index)

    remaining_gdl_counts = Counter(gdl)
    used = [False] * len(legacy)
    next_unused = 0

    def advance_next_unused() -> None:
        nonlocal next_unused
        while next_unused < len(legacy) and used[next_unused]:
            next_unused += 1

    def flush_unmatchable_legacy() -> None:
        """Emit Legacy lines that cannot pair with any remaining GDL line."""
        nonlocal next_unused
        advance_next_unused()
        while next_unused < len(legacy) and remaining_gdl_counts[legacy[next_unused]] == 0:
            rows.append(
                ReportRow(
                    gdl_segment="",
                    legacy_segment=legacy[next_unused],
                    status=STATUS_MISSING_IN_GDL,
                    fill=FILL_RED,
                )
            )
            used[next_unused] = True
            leftover = remaining_legacy.get(legacy[next_unused])
            if leftover and leftover[0] == next_unused:
                leftover.popleft()
            next_unused += 1
            advance_next_unused()

    def consume_legacy(index: int) -> None:
        used[index] = True
        leftover = remaining_legacy.get(legacy[index])
        if leftover and leftover[0] == index:
            leftover.popleft()
        elif leftover:
            leftover.remove(index)

    for gdl_line in gdl:
        advance_next_unused()
        matched_index: int | None = None
        out_of_order = False
        candidates = remaining_legacy.get(gdl_line)

        if next_unused < len(legacy) and legacy[next_unused] == gdl_line:
            matched_index = next_unused
        elif candidates:
            for candidate in candidates:
                if candidate > next_unused:
                    matched_index = candidate
                    out_of_order = True
                    break

        if matched_index is not None:
            rows.append(
                ReportRow(
                    gdl_segment=gdl_line,
                    legacy_segment=legacy[matched_index],
                    status=STATUS_MATCH,
                    fill=FILL_YELLOW if out_of_order else None,
                )
            )
            consume_legacy(matched_index)
        else:
            rows.append(
                ReportRow(
                    gdl_segment=gdl_line,
                    legacy_segment="",
                    status=STATUS_MISSING_IN_IMPULSE,
                    fill=FILL_RED,
                )
            )

        remaining_gdl_counts[gdl_line] -= 1
        if remaining_gdl_counts[gdl_line] <= 0:
            del remaining_gdl_counts[gdl_line]

        flush_unmatchable_legacy()

    flush_unmatchable_legacy()
    while next_unused < len(legacy):
        if not used[next_unused]:
            rows.append(
                ReportRow(
                    gdl_segment="",
                    legacy_segment=legacy[next_unused],
                    status=STATUS_MISSING_IN_GDL,
                    fill=FILL_RED,
                )
            )
            used[next_unused] = True
        next_unused += 1
        advance_next_unused()

    return ComparisonResult(
        rows=rows,
        gdl_line_count=len(all_gdl),
        legacy_line_count=len(all_legacy),
    )


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
    orange_fmt = workbook.add_format(
        {"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#FFC000"}
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
    orange_value_fmt = workbook.add_format({"border": 1, "bg_color": "#FFC000"})
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
    last_data_row = max(len(result.rows), 1)
    detail.autofilter(0, 0, last_data_row, 2)

    for row_number, row in enumerate(result.rows, start=1):
        if row.fill == FILL_YELLOW:
            cell_fmt = yellow_fmt
        elif row.fill == FILL_ORANGE:
            cell_fmt = orange_fmt
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
        ("Mismatch (first 4 lines)", result.mismatches, orange_value_fmt),
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
    print(f"  Mismatch (first 4 lines):  {result.mismatches:,}")
    print(f"  Missing in Impulse:        {result.missing_in_impulse:,}")
    print(f"  Missing in GDL:            {result.missing_in_gdl:,}")
    print(f"  Report rows:               {len(result.rows):,}")
    print(f"  Excel report:              {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
