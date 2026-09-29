from __future__ import annotations

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
    write_excel_report,
)


def rows_as_tuples(result) -> list[tuple[str, str, str, str | None]]:
    return [(row.gdl_segment, row.legacy_segment, row.status, row.fill) for row in result.rows]


def test_first_lines_are_always_written_as_mismatch() -> None:
    result = compare_segments(["ISA-GDL", "ST~846~0001", "LIN~A"], ["ISA-LEG", "ST~846~0001", "LIN~A"])

    assert rows_as_tuples(result) == [
        ("ISA-GDL", "ISA-LEG", STATUS_MISMATCH, FILL_ORANGE),
        ("ST~846~0001", "ST~846~0001", STATUS_MATCH, None),
        ("LIN~A", "LIN~A", STATUS_MATCH, None),
    ]
    assert result.mismatches == 1
    assert result.sequential_matches == 2


def test_first_lines_are_excluded_from_later_matching() -> None:
    result = compare_segments(["FOO", "BAR"], ["BAZ", "FOO"])

    assert rows_as_tuples(result) == [
        ("FOO", "BAZ", STATUS_MISMATCH, FILL_ORANGE),
        ("BAR", "", STATUS_MISSING_IN_IMPULSE, FILL_RED),
        ("", "FOO", STATUS_MISSING_IN_GDL, FILL_RED),
    ]


def test_sequential_matches_have_no_fill_after_first_line() -> None:
    result = compare_segments(["A", "B", "C"], ["A", "B", "C"])

    assert rows_as_tuples(result) == [
        ("A", "A", STATUS_MISMATCH, FILL_ORANGE),
        ("B", "B", STATUS_MATCH, None),
        ("C", "C", STATUS_MATCH, None),
    ]
    assert result.sequential_matches == 2
    assert result.out_of_order_matches == 0
    assert result.mismatches == 1
    assert result.missing_in_impulse == 0
    assert result.missing_in_gdl == 0


def test_out_of_order_match_is_yellow() -> None:
    result = compare_segments(["H", "A", "C", "B"], ["H", "A", "B", "C"])

    assert rows_as_tuples(result) == [
        ("H", "H", STATUS_MISMATCH, FILL_ORANGE),
        ("A", "A", STATUS_MATCH, None),
        ("C", "C", STATUS_MATCH, FILL_YELLOW),
        ("B", "B", STATUS_MATCH, None),
    ]


def test_gdl_line_missing_in_impulse_is_red() -> None:
    result = compare_segments(["H", "A", "X", "C"], ["H", "A", "C"])

    assert rows_as_tuples(result) == [
        ("H", "H", STATUS_MISMATCH, FILL_ORANGE),
        ("A", "A", STATUS_MATCH, None),
        ("X", "", STATUS_MISSING_IN_IMPULSE, FILL_RED),
        ("C", "C", STATUS_MATCH, None),
    ]


def test_legacy_extra_line_is_missing_in_gdl_then_later_line_still_matches() -> None:
    result = compare_segments(["H", "A", "C"], ["H", "A", "B", "C"])

    assert rows_as_tuples(result) == [
        ("H", "H", STATUS_MISMATCH, FILL_ORANGE),
        ("A", "A", STATUS_MATCH, None),
        ("", "B", STATUS_MISSING_IN_GDL, FILL_RED),
        ("C", "C", STATUS_MATCH, None),
    ]


def test_duplicate_lines_are_consumed_one_to_one() -> None:
    result = compare_segments(
        ["HDR", "QTY~33~0~EA", "LIN~~MG~AAA", "QTY~33~0~EA"],
        ["HDR", "QTY~33~0~EA", "LIN~~MG~BBB", "QTY~33~0~EA"],
    )

    assert rows_as_tuples(result) == [
        ("HDR", "HDR", STATUS_MISMATCH, FILL_ORANGE),
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

    result = compare_segments(["H", "A", "C", "X"], ["H", "A", "B", "C"])
    output = tmp_path / "report.xlsx"
    write_excel_report(result, output)

    workbook = openpyxl.load_workbook(output)
    sheet = workbook["Comparison"]
    assert [cell.value for cell in sheet[1]] == ["GDL Segment", "Legacy Segment", "status"]

    rows = [[sheet.cell(row=i, column=j).value for j in range(1, 4)] for i in range(2, 7)]
    assert rows == [
        ["H", "H", STATUS_MISMATCH],
        ["A", "A", STATUS_MATCH],
        [None, "B", STATUS_MISSING_IN_GDL],
        ["C", "C", STATUS_MATCH],
        ["X", None, STATUS_MISSING_IN_IMPULSE],
    ]

    orange_mismatch = sheet.cell(row=2, column=3).fill.fgColor.rgb
    red_missing_legacy = sheet.cell(row=4, column=2).fill.fgColor.rgb
    red_missing_gdl = sheet.cell(row=6, column=3).fill.fgColor.rgb
    assert orange_mismatch.endswith("FFC000")
    assert red_missing_legacy.endswith("FF0000")
    assert red_missing_gdl.endswith("FF0000")

    summary = workbook["Summary"]
    metrics = {summary.cell(row=i, column=1).value: summary.cell(row=i, column=2).value for i in range(2, 10)}
    assert metrics["Sequential Match"] == 2
    assert metrics["Out-of-order Match (yellow)"] == 0
    assert metrics["Mismatch (first lines)"] == 1
    assert metrics["Missing in Impulse"] == 1
    assert metrics["Missing in GDL"] == 1
