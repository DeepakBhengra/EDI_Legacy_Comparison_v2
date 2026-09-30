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

## How to run

### 1. Prerequisites

- Python 3.10 or later
- VS Code (optional, but recommended)
- This project folder opened as the workspace root (the folder that contains `compare_segments.py`)

Check Python:

```bash
python --version
```

On some Windows setups use `py -3 --version` instead.

### 2. Create and activate a virtual environment

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows PowerShell:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, run this once, then activate again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

In VS Code, press `Ctrl+Shift+P` → **Python: Select Interpreter** → choose `.venv`.

### 3. Install packages

Install into the same Python that will run the script:

```bash
python -m pip install -r requirements.txt
```

Confirm the Excel library is visible:

```bash
python -c "import xlsxwriter; print('ok', xlsxwriter.__version__)"
```

If that fails, VS Code is using a different Python. Select the `.venv` interpreter and run `python -m pip install -r requirements.txt` again.

### 4. Run the comparison

From the project root:

```bash
python compare_segments.py
```

That reads:

- `data/GDL_Segment.txt`
- `data/Legacy_Segment.txt`

and writes a **new** Excel file under `reports/` with a timestamp in the name, for example:

`reports/segment_comparison_20260930_091300.xlsx`

The terminal prints the exact output path when it finishes.

### 5. Use your own files

```bash
python compare_segments.py --gdl path/to/gdl.txt --legacy path/to/legacy.txt -o reports/segment_comparison.xlsx
```

A timestamp is still added, so `-o reports/segment_comparison.xlsx` becomes something like `reports/segment_comparison_20260930_091300.xlsx`.

### 6. Open the report

Open the newest `.xlsx` file in `reports/`. It has two sheets:

- **Comparison** — `GDL Segment`, `Legacy Segment`, `status`
- **Summary** — MATCH / MISMATCH / missing counts and ST–SE block counts

Do not reuse an older workbook from a previous run.

### Optional tests

```bash
python -m pytest -q
```
