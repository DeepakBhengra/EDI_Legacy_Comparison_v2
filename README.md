# GDL vs Legacy Segment Comparison

Python tool that ingests a GDL segment file and a Legacy (Impulse) segment file, compares them line by line, and writes an Excel report.

The bundled sample files are EDI 846 inventory segments. The same script works for any two line-oriented segment dumps.

## Matching rules

Each GDL line is compared against the current unused Legacy line, then against later unused Legacy lines if needed.

| Situation | GDL Segment | Legacy Segment | status | Row color |
| --- | --- | --- | --- | --- |
| Current GDL line equals the current Legacy line | GDL line | Legacy line | `Match` | none |
| Current GDL line does not equal the current Legacy line, but the same GDL line exists later in Legacy | GDL line | that later Legacy line | `Match` | yellow |
| Current GDL line is not present in any remaining Legacy line | GDL line | blank | `Missing in Impulse` | red |
| A Legacy line cannot match any remaining GDL line | blank | Legacy line | `Missing in GDL` | red |

Duplicate lines (for example many `QTY~33~0~EA` rows) are matched one-to-one. A line is never reused after it has been paired. Unique header mismatches (dates, control numbers) are reported as missing on both sides so later in-order product lines can still sequential-match.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

Compare the bundled sample files and write `reports/segment_comparison.xlsx`:

```bash
python compare_segments.py
```

Or pass explicit paths:

```bash
python compare_segments.py \
  --gdl data/GDL_Segment.txt \
  --legacy data/Legacy_Segment.txt \
  --output reports/segment_comparison.xlsx
```

The workbook has two sheets:

- **Comparison** — one row per GDL line, then leftover unmatched Legacy lines. Columns are `GDL Segment`, `Legacy Segment`, and `status`.
- **Summary** — counts of sequential matches, out-of-order (yellow) matches, and missing rows.

The report is `.xlsx` rather than legacy `.xls` so it can hold the full file size (these samples are ~55k lines each). Completely empty lines are ignored so trailing blank rows from file exports do not appear as mismatches.

## Tests

```bash
pytest -q
```
