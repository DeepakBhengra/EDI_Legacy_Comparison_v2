# GDL vs Legacy Segment Comparison

Python tool that compares a GDL EDI 846 dump with a Legacy (Impulse) 846 dump and writes an Excel report.

## File structure

Each file has this shape:

1. `ISA` — interchange header (once)
2. `GS` — functional group header (once)
3. One or more `ST`–`SE` transaction blocks:
   - `ST` → `BIA` → optional `DTM` → `N1` / `N2` / `N3` / `N4` → `PER`
   - repeating `LIN` + `QTY` pairs
   - `CTT` → `SE`
4. `GE` and `IEA` trailers (once each)

`LIN`/`QTY` order inside a block may differ between GDL and Legacy.

## Matching rules

| Situation | GDL Segment | Legacy Segment | status | Row color |
| --- | --- | --- | --- | --- |
| `ISA` or `GS` contents are equal | line | line | `MATCH` | none |
| `ISA` or `GS` contents differ | line | line | `MISMATCH` | orange |
| `ST`–`SE` header (`ST`, `BIA`, `N1`, `N2`, `N3`, `N4`, `PER`, `CTT`, `SE`) equal | line | line | `MATCH` | none |
| Those header lines differ | line | line | `MISMATCH` | orange |
| GDL `LIN` exists in the same warehouse block before `CTT` | GDL `LIN`/`QTY` | matching Legacy `LIN`/`QTY` | `MATCH` | none |
| GDL `LIN` is not in that Legacy block | GDL `LIN`/`QTY` | blank | `Missing in Impulse` | red |
| Legacy `LIN`/`QTY` has no GDL pair in that block | blank | Legacy `LIN`/`QTY` | `Missing in GDL` | red |

`ST`–`SE` blocks are paired by the `N1` warehouse line, not by `ST` control number. That keeps the same warehouse together when GDL inserts extra blocks without `N1`. Unpaired GDL blocks are `Missing in Impulse`. Unpaired Legacy blocks are `Missing in GDL`.

`LIN` matching stays inside one `ST`–`SE` block. A later `LIN` in the same block still counts as `MATCH`.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

Each run writes a new workbook under `reports/` with a timestamp, for example `reports/segment_comparison_20260930_052800.xlsx`:

```bash
python compare_segments.py
```

Or pass explicit paths. The timestamp is still added to the output name:

```bash
python compare_segments.py \
  --gdl data/GDL_Segment.txt \
  --legacy data/Legacy_Segment.txt \
  --output reports/segment_comparison.xlsx
```

The workbook has two sheets:

- **Comparison** — `GDL Segment`, `Legacy Segment`, `status`
- **Summary** — MATCH / MISMATCH / missing counts and ST–SE block counts

Install packages into the same Python that VS Code uses:

```bash
python -m pip install -r requirements.txt
python compare_segments.py
```

## Tests

```bash
pytest -q
```
