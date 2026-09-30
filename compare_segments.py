#!/usr/bin/env python3
"""Compare GDL and Legacy EDI 846 segment files and write an Excel report."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

STATUS_MATCH = "MATCH"
STATUS_MISMATCH = "MISMATCH"
STATUS_MISSING_IN_IMPULSE = "Missing in Impulse"
STATUS_MISSING_IN_GDL = "Missing in GDL"

FILL_ORANGE = "orange"
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
    gdl_block_count: int = 0
    legacy_block_count: int = 0
    paired_block_count: int = 0

    @property
    def matches(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MATCH)

    @property
    def mismatches(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISMATCH)

    @property
    def missing_in_impulse(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISSING_IN_IMPULSE)

    @property
    def missing_in_gdl(self) -> int:
        return sum(1 for row in self.rows if row.status == STATUS_MISSING_IN_GDL)


@dataclass
class LinQtyPair:
    lin: str
    qty: str | None = None


@dataclass
class StBlock:
    st: str | None = None
    bia: str | None = None
    dtm: list[str] = field(default_factory=list)
    n1: str | None = None
    n2: str | None = None
    n3: str | None = None
    n4: str | None = None
    per: str | None = None
    items: list[LinQtyPair] = field(default_factory=list)
    ctt: str | None = None
    se: str | None = None
    other: list[str] = field(default_factory=list)


@dataclass
class EdiDocument:
    isa: str | None = None
    gs: str | None = None
    blocks: list[StBlock] = field(default_factory=list)
    ge: str | None = None
    iea: str | None = None
    other: list[str] = field(default_factory=list)


def read_segment_lines(path: Path) -> list[str]:
    """Read a segment file as individual lines, preserving inner whitespace."""
    text = path.read_text(encoding="utf-8", errors="replace")
    return [line for line in text.splitlines() if line != ""]


def timestamped_output_path(path: Path, when: datetime | None = None) -> Path:
    """Insert a local timestamp before the file suffix so each run writes a new file."""
    stamp = (when or datetime.now()).strftime("%Y%m%d_%H%M%S")
    suffix = path.suffix or ".xlsx"
    candidate = path.with_name(f"{path.stem}_{stamp}{suffix}")
    if candidate.exists():
        stamp = (when or datetime.now()).strftime("%Y%m%d_%H%M%S_%f")
        candidate = path.with_name(f"{path.stem}_{stamp}{suffix}")
    return candidate


def segment_prefix(line: str) -> str:
    return line.split("~", 1)[0]


def parse_edi_document(lines: Iterable[str]) -> EdiDocument:
    """Split an 846 dump into ISA/GS, ST-SE blocks, and GE/IEA."""
    document = EdiDocument()
    block: StBlock | None = None
    pending_item: LinQtyPair | None = None

    def close_item() -> None:
        nonlocal pending_item
        if block is not None and pending_item is not None:
            block.items.append(pending_item)
        pending_item = None

    def start_block(st_line: str) -> None:
        nonlocal block, pending_item
        close_item()
        block = StBlock(st=st_line)
        document.blocks.append(block)
        pending_item = None

    for line in lines:
        prefix = segment_prefix(line)
        if prefix == "ISA":
            document.isa = line
            block = None
        elif prefix == "GS":
            document.gs = line
            block = None
        elif prefix == "GE":
            close_item()
            block = None
            document.ge = line
        elif prefix == "IEA":
            close_item()
            block = None
            document.iea = line
        elif prefix == "ST":
            start_block(line)
        elif block is None:
            document.other.append(line)
        elif prefix == "BIA":
            close_item()
            block.bia = line
        elif prefix == "DTM":
            close_item()
            block.dtm.append(line)
        elif prefix == "N1":
            close_item()
            block.n1 = line
        elif prefix == "N2":
            close_item()
            block.n2 = line
        elif prefix == "N3":
            close_item()
            block.n3 = line
        elif prefix == "N4":
            close_item()
            block.n4 = line
        elif prefix == "PER":
            close_item()
            block.per = line
        elif prefix == "LIN":
            close_item()
            pending_item = LinQtyPair(lin=line)
        elif prefix == "QTY":
            if pending_item is not None:
                pending_item.qty = line
                close_item()
            else:
                block.items.append(LinQtyPair(lin="", qty=line))
        elif prefix == "CTT":
            close_item()
            block.ctt = line
        elif prefix == "SE":
            close_item()
            block.se = line
        else:
            close_item()
            block.other.append(line)

    close_item()
    return document


def _row(gdl: str, legacy: str, status: str) -> ReportRow:
    fill = None
    if status == STATUS_MISMATCH:
        fill = FILL_ORANGE
    elif status in {STATUS_MISSING_IN_IMPULSE, STATUS_MISSING_IN_GDL}:
        fill = FILL_RED
    return ReportRow(gdl_segment=gdl, legacy_segment=legacy, status=status, fill=fill)


def _compare_optional(gdl: str | None, legacy: str | None) -> list[ReportRow]:
    if gdl is None and legacy is None:
        return []
    if gdl is not None and legacy is not None:
        status = STATUS_MATCH if gdl == legacy else STATUS_MISMATCH
        return [_row(gdl, legacy, status)]
    if gdl is not None:
        return [_row(gdl, "", STATUS_MISSING_IN_IMPULSE)]
    return [_row("", legacy or "", STATUS_MISSING_IN_GDL)]


def _compare_lists(gdl_values: list[str], legacy_values: list[str]) -> list[ReportRow]:
    rows: list[ReportRow] = []
    limit = max(len(gdl_values), len(legacy_values))
    for index in range(limit):
        gdl = gdl_values[index] if index < len(gdl_values) else None
        legacy = legacy_values[index] if index < len(legacy_values) else None
        rows.extend(_compare_optional(gdl, legacy))
    return rows


def _emit_block_missing(block: StBlock, side: str) -> list[ReportRow]:
    status = STATUS_MISSING_IN_IMPULSE if side == "gdl" else STATUS_MISSING_IN_GDL
    rows: list[ReportRow] = []

    def add(value: str | None) -> None:
        if not value:
            return
        if side == "gdl":
            rows.append(_row(value, "", status))
        else:
            rows.append(_row("", value, status))

    add(block.st)
    add(block.bia)
    for line in block.dtm:
        add(line)
    add(block.n1)
    add(block.n2)
    add(block.n3)
    add(block.n4)
    add(block.per)
    for extra in block.other:
        add(extra)
    for item in block.items:
        add(item.lin or None)
        add(item.qty)
    add(block.ctt)
    add(block.se)
    return rows


def _compare_items(gdl_items: list[LinQtyPair], legacy_items: list[LinQtyPair]) -> list[ReportRow]:
    """Match LIN/QTY pairs inside one ST-SE block. Order may differ."""
    rows: list[ReportRow] = []
    used = [False] * len(legacy_items)

    for gdl_item in gdl_items:
        found_index = None
        for index, legacy_item in enumerate(legacy_items):
            if used[index]:
                continue
            if gdl_item.lin and gdl_item.lin == legacy_item.lin:
                found_index = index
                break
        if found_index is None:
            if gdl_item.lin:
                rows.append(_row(gdl_item.lin, "", STATUS_MISSING_IN_IMPULSE))
            if gdl_item.qty:
                rows.append(_row(gdl_item.qty, "", STATUS_MISSING_IN_IMPULSE))
            continue

        used[found_index] = True
        legacy_item = legacy_items[found_index]
        rows.append(_row(gdl_item.lin, legacy_item.lin, STATUS_MATCH))
        rows.extend(_compare_optional(gdl_item.qty, legacy_item.qty))

    for index, legacy_item in enumerate(legacy_items):
        if used[index]:
            continue
        if legacy_item.lin:
            rows.append(_row("", legacy_item.lin, STATUS_MISSING_IN_GDL))
        if legacy_item.qty:
            rows.append(_row("", legacy_item.qty, STATUS_MISSING_IN_GDL))
    return rows


def _compare_blocks(gdl_block: StBlock, legacy_block: StBlock) -> list[ReportRow]:
    rows: list[ReportRow] = []
    rows.extend(_compare_optional(gdl_block.st, legacy_block.st))
    rows.extend(_compare_optional(gdl_block.bia, legacy_block.bia))
    rows.extend(_compare_lists(gdl_block.dtm, legacy_block.dtm))
    rows.extend(_compare_optional(gdl_block.n1, legacy_block.n1))
    rows.extend(_compare_optional(gdl_block.n2, legacy_block.n2))
    rows.extend(_compare_optional(gdl_block.n3, legacy_block.n3))
    rows.extend(_compare_optional(gdl_block.n4, legacy_block.n4))
    rows.extend(_compare_optional(gdl_block.per, legacy_block.per))
    rows.extend(_compare_lists(gdl_block.other, legacy_block.other))
    rows.extend(_compare_items(gdl_block.items, legacy_block.items))
    rows.extend(_compare_optional(gdl_block.ctt, legacy_block.ctt))
    rows.extend(_compare_optional(gdl_block.se, legacy_block.se))
    return rows


def compare_segments(gdl_lines: Iterable[str], legacy_lines: Iterable[str]) -> ComparisonResult:
    """Compare two 846 files: ISA/GS once, then each ST-SE block, then GE/IEA.

    ST-SE blocks with an N1 warehouse line are paired by that N1. LIN/QTY pairs
    are matched by LIN content inside the paired block, even when the order
    differs. Unmatched LIN/QTY rows are Missing in Impulse or Missing in GDL.
    """
    gdl_doc = parse_edi_document(list(gdl_lines))
    legacy_doc = parse_edi_document(list(legacy_lines))
    rows: list[ReportRow] = []

    rows.extend(_compare_optional(gdl_doc.isa, legacy_doc.isa))
    rows.extend(_compare_optional(gdl_doc.gs, legacy_doc.gs))
    rows.extend(_compare_lists(gdl_doc.other, legacy_doc.other))

    leftover_legacy: dict[str, list[int]] = {}
    for index, block in enumerate(legacy_doc.blocks):
        if block.n1:
            leftover_legacy.setdefault(block.n1, []).append(index)
    used_legacy: set[int] = set()
    paired_blocks = 0

    for gdl_block in gdl_doc.blocks:
        pair_index = None
        if gdl_block.n1:
            candidates = leftover_legacy.get(gdl_block.n1) or []
            if candidates:
                pair_index = candidates.pop(0)
        if pair_index is None:
            rows.extend(_emit_block_missing(gdl_block, "gdl"))
            continue
        used_legacy.add(pair_index)
        paired_blocks += 1
        rows.extend(_compare_blocks(gdl_block, legacy_doc.blocks[pair_index]))

    for index, legacy_block in enumerate(legacy_doc.blocks):
        if index not in used_legacy:
            rows.extend(_emit_block_missing(legacy_block, "legacy"))

    rows.extend(_compare_optional(gdl_doc.ge, legacy_doc.ge))
    rows.extend(_compare_optional(gdl_doc.iea, legacy_doc.iea))

    return ComparisonResult(
        rows=rows,
        gdl_line_count=_document_line_count(gdl_doc),
        legacy_line_count=_document_line_count(legacy_doc),
        gdl_block_count=len(gdl_doc.blocks),
        legacy_block_count=len(legacy_doc.blocks),
        paired_block_count=paired_blocks,
    )


def _document_line_count(document: EdiDocument) -> int:
    count = 0
    for value in (document.isa, document.gs, document.ge, document.iea):
        if value:
            count += 1
    count += len(document.other)
    for block in document.blocks:
        count += sum(
            1
            for value in (
                block.st,
                block.bia,
                block.n1,
                block.n2,
                block.n3,
                block.n4,
                block.per,
                block.ctt,
                block.se,
            )
            if value
        )
        count += len(block.dtm) + len(block.other)
        for item in block.items:
            count += 1 if item.lin else 0
            count += 1 if item.qty else 0
    return count


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
        if row.fill == FILL_ORANGE:
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
        ("GDL ST-SE blocks", result.gdl_block_count, value_fmt),
        ("Legacy ST-SE blocks", result.legacy_block_count, value_fmt),
        ("Paired warehouse blocks", result.paired_block_count, value_fmt),
        ("MATCH", result.matches, value_fmt),
        ("MISMATCH", result.mismatches, orange_value_fmt),
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
        description="Compare GDL and Legacy EDI 846 segment files and write an Excel report."
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
        help="Base path for the Excel report. A timestamp is added to the filename on each run.",
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
    output_path = timestamped_output_path(args.output)
    write_excel_report(result, output_path)

    print()
    print("Comparison complete")
    print(f"  GDL ST-SE blocks:          {result.gdl_block_count:,}")
    print(f"  Legacy ST-SE blocks:       {result.legacy_block_count:,}")
    print(f"  Paired warehouse blocks:   {result.paired_block_count:,}")
    print(f"  MATCH:                     {result.matches:,}")
    print(f"  MISMATCH:                  {result.mismatches:,}")
    print(f"  Missing in Impulse:        {result.missing_in_impulse:,}")
    print(f"  Missing in GDL:            {result.missing_in_gdl:,}")
    print(f"  Report rows:               {len(result.rows):,}")
    print(f"  Excel report:              {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
