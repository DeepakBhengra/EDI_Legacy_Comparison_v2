from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from compare_segments import (
    FILL_ORANGE,
    FILL_RED,
    FILL_YELLOW,
    STATUS_MATCH,
    STATUS_MISMATCH,
    STATUS_MISSING_IN_GDL,
    STATUS_MISSING_IN_IMPULSE,
    compare_segments,
    read_segment_lines,
    timestamped_output_path,
    write_excel_report,
)

HEADERS_GDL = ["ISA-GDL", "GS-GDL", "ST-GDL", "BIA-GDL"]
HEADERS_LEG = ["ISA-LEG", "GS-LEG", "ST-LEG", "BIA-LEG"]
HEADER_MISMATCH_ROWS = [
    ("ISA-GDL", "ISA-LEG", STATUS_MISMATCH, FILL_ORANGE),
    ("GS-GDL", "GS-LEG", STATUS_MISMATCH, FILL_ORANGE),
    ("ST-GDL", "ST-LEG", STATUS_MISMATCH, FILL_ORANGE),
    ("BIA-GDL", "BIA-LEG", STATUS_MISMATCH, FILL_ORANGE),
]


def rows_as_tuples(result) -> list[tuple[str, str, str, str | None]]:
    return [(row.gdl_segment, row.legacy_segment, row.status, row.fill) for row in result.rows]


def test_timestamped_output_path_inserts_local_timestamp() -> None:
    when = datetime(2026, 9, 29, 13, 4, 45)
    path = timestamped_output_path(Path("reports/segment_comparison.xlsx"), when=when)
    assert path == Path("reports/segment_comparison_20260929_130445.xlsx")


def test_timestamped_output_path_adds_microseconds_when_file_exists(tmp_path: Path) -> None:
    when = datetime(2026, 9, 29, 13, 4, 45)
    existing = tmp_path / "segment_comparison_20260929_130445.xlsx"
    existing.write_bytes(b"")
    path = timestamped_output_path(tmp_path / "segment_comparison.xlsx", when=when)
    assert path.name.startswith("segment_comparison_20260929_130445_")
    assert path.suffix == ".xlsx"
    assert path != existing


def test_first_four_lines_are_always_written_as_mismatch() -> None:
    result = compare_segments(
        [*HEADERS_GDL, "N1~WH", "LIN~A"],
        [*HEADERS_LEG, "N1~WH", "LIN~A"],
    )

    assert rows_as_tuples(result) == [
        *HEADER_MISMATCH_ROWS,
        ("N1~WH", "N1~WH", STATUS_MATCH, None),
        ("LIN~A", "LIN~A", STATUS_MATCH, None),
    ]
    assert result.mismatches == 4
    assert result.sequential_matches == 2


def test_first_four_lines_are_excluded_from_later_matching() -> None:
    result = compare_segments(
        ["FOO", "H2", "H3", "H4", "BAR"],
        ["BAZ", "H2", "H3", "H4", "FOO"],
    )

    assert rows_as_tuples(result) == [
        ("FOO", "BAZ", STATUS_MISMATCH, FILL_ORANGE),
        ("H2", "H2", STATUS_MISMATCH, FILL_ORANGE),
        ("H3", "H3", STATUS_MISMATCH, FILL_ORANGE),
        ("H4", "H4", STATUS_MISMATCH, FILL_ORANGE),
        ("BAR", "", STATUS_MISSING_IN_IMPULSE, FILL_RED),
        ("", "FOO", STATUS_MISSING_IN_GDL, FILL_RED),
    ]


def test_files_shorter_than_four_lines_pair_all_as_mismatch() -> None:
    result = compare_segments(["A", "B"], ["C", "D"])

    assert rows_as_tuples(result) == [
        ("A", "C", STATUS_MISMATCH, FILL_ORANGE),
        ("B", "D", STATUS_MISMATCH, FILL_ORANGE),
    ]
    assert result.mismatches == 2
    assert result.sequential_matches == 0


def test_sequential_matches_have_no_fill_after_header_lines() -> None:
    result = compare_segments([*HEADERS_GDL, "B", "C"], [*HEADERS_LEG, "B", "C"])

    assert rows_as_tuples(result) == [
        *HEADER_MISMATCH_ROWS,
        ("B", "B", STATUS_MATCH, None),
        ("C", "C", STATUS_MATCH, None),
    ]
    assert result.sequential_matches == 2
    assert result.out_of_order_matches == 0
    assert result.mismatches == 4
    assert result.missing_in_impulse == 0
    assert result.missing_in_gdl == 0


def test_out_of_order_match_is_yellow() -> None:
    result = compare_segments([*HEADERS_GDL, "A", "C", "B"], [*HEADERS_LEG, "A", "B", "C"])

    assert rows_as_tuples(result) == [
        *HEADER_MISMATCH_ROWS,
        ("A", "A", STATUS_MATCH, None),
        ("C", "C", STATUS_MATCH, FILL_YELLOW),
        ("B", "B", STATUS_MATCH, None),
    ]


def test_gdl_line_missing_in_impulse_is_red() -> None:
    result = compare_segments([*HEADERS_GDL, "A", "X", "C"], [*HEADERS_LEG, "A", "C"])

    assert rows_as_tuples(result) == [
        *HEADER_MISMATCH_ROWS,
        ("A", "A", STATUS_MATCH, None),
        ("X", "", STATUS_MISSING_IN_IMPULSE, FILL_RED),
        ("C", "C", STATUS_MATCH, None),
    ]


def test_legacy_extra_line_is_missing_in_gdl_then_later_line_still_matches() -> None:
    result = compare_segments([*HEADERS_GDL, "A", "C"], [*HEADERS_LEG, "A", "B", "C"])

    assert rows_as_tuples(result) == [
        *HEADER_MISMATCH_ROWS,
        ("A", "A", STATUS_MATCH, None),
        ("", "B", STATUS_MISSING_IN_GDL, FILL_RED),
        ("C", "C", STATUS_MATCH, None),
    ]


def test_duplicate_lines_are_consumed_one_to_one() -> None:
    result = compare_segments(
        [*HEADERS_GDL, "QTY~33~0~EA", "LIN~~MG~AAA", "QTY~33~0~EA"],
        [*HEADERS_LEG, "QTY~33~0~EA", "LIN~~MG~BBB", "QTY~33~0~EA"],
    )

    assert rows_as_tuples(result) == [
        *HEADER_MISMATCH_ROWS,
        ("QTY~33~0~EA", "QTY~33~0~EA", STATUS_MATCH, None),
        ("", "LIN~~MG~BBB", STATUS_MISSING_IN_GDL, FILL_RED),
        ("LIN~~MG~AAA", "", STATUS_MISSING_IN_IMPULSE, FILL_RED),
        ("QTY~33~0~EA", "QTY~33~0~EA", STATUS_MATCH, None),
    ]


def test_empty_inputs_produce_empty_report() -> None:
    result = compare_segments([], [])
    assert result.rows == []
    assert result.gdl_line_count == 0
    assert result.legacy_line_count == 0


def test_read_segment_lines_keeps_inner_spaces_and_splits_crlf(tmp_path: Path) -> None:
    path = tmp_path / "sample.txt"
    path.write_bytes(b"N4~CP~0  ~~WH~00\r\nQTY~33~0~EA\r\n")

    assert read_segment_lines(path) == ["N4~CP~0  ~~WH~00", "QTY~33~0~EA"]


def test_read_segment_lines_skips_empty_lines(tmp_path: Path) -> None:
    path = tmp_path / "sample.txt"
    path.write_text("ST~846~0001\n\n\nIEA~1~1\n\n", encoding="utf-8")

    assert read_segment_lines(path) == ["ST~846~0001", "IEA~1~1"]


def test_excel_report_headers_status_and_fill_colors(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")

    result = compare_segments(
        [*HEADERS_GDL, "A", "C", "X"],
        [*HEADERS_LEG, "A", "B", "C"],
    )
    output = tmp_path / "report.xlsx"
    write_excel_report(result, output)

    workbook = openpyxl.load_workbook(output)
    sheet = workbook["Comparison"]
    assert [cell.value for cell in sheet[1]] == ["GDL Segment", "Legacy Segment", "status"]

    rows = [[sheet.cell(row=i, column=j).value for j in range(1, 4)] for i in range(2, 10)]
    assert rows == [
        ["ISA-GDL", "ISA-LEG", STATUS_MISMATCH],
        ["GS-GDL", "GS-LEG", STATUS_MISMATCH],
        ["ST-GDL", "ST-LEG", STATUS_MISMATCH],
        ["BIA-GDL", "BIA-LEG", STATUS_MISMATCH],
        ["A", "A", STATUS_MATCH],
        [None, "B", STATUS_MISSING_IN_GDL],
        ["C", "C", STATUS_MATCH],
        ["X", None, STATUS_MISSING_IN_IMPULSE],
    ]

    orange_mismatch = sheet.cell(row=2, column=3).fill.fgColor.rgb
    red_missing_legacy = sheet.cell(row=7, column=2).fill.fgColor.rgb
    red_missing_gdl = sheet.cell(row=9, column=3).fill.fgColor.rgb
    assert orange_mismatch.endswith("FFC000")
    assert red_missing_legacy.endswith("FF0000")
    assert red_missing_gdl.endswith("FF0000")

    summary = workbook["Summary"]
    metrics = {summary.cell(row=i, column=1).value: summary.cell(row=i, column=2).value for i in range(2, 10)}
    assert metrics["Sequential Match"] == 2
    assert metrics["Out-of-order Match (yellow)"] == 0
    assert metrics["Mismatch (first 4 lines)"] == 4
    assert metrics["Missing in Impulse"] == 1
    assert metrics["Missing in GDL"] == 1


def test_main_writes_a_timestamped_excel_file(tmp_path: Path) -> None:
    from compare_segments import main

    gdl = tmp_path / "gdl.txt"
    legacy = tmp_path / "legacy.txt"
    gdl.write_text("\n".join([*HEADERS_GDL, "LIN~A"]) + "\n", encoding="utf-8")
    legacy.write_text("\n".join([*HEADERS_LEG, "LIN~A"]) + "\n", encoding="utf-8")
    output = tmp_path / "report.xlsx"

    assert main(["--gdl", str(gdl), "--legacy", str(legacy), "-o", str(output)]) == 0
    reports = list(tmp_path.glob("report_*.xlsx"))
    assert len(reports) == 1
    assert reports[0].name.startswith("report_")
    assert not output.exists()
