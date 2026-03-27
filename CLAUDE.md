# CLAUDE.md — Owner Financial Extractor

## What This Project Does

Extracts owner financial data from property management Excel workbooks
for underwriting and proforma preparation. Supports multiple management
company formats. Outputs a structured Excel datapack and database-ready CSVs.

---

## Critical Rules

1. **Never break the Extra Space CLI workflow.** `run.bat` → `extract_owner_financials.py`
   → `extractor_core.py` is the primary production workflow. Any change that
   breaks this requires explicit instruction.

2. **`extractor_core.py` must never print to console or prompt for input.**
   It takes arguments and returns data. The CLI and webapp are the only I/O layers.

3. **New manager formats branch inside `process_workbook()` using `managed_by`.**
   Do not duplicate extraction logic across files. Add a new branch; keep EXR untouched.

4. **Prefer label-based extraction over hardcoded cell addresses.**
   Hardcoded cells break silently when the source file format changes.

5. **Do not invent a new COA taxonomy.** The target COA structure is defined
   in the proforma's `COA Translation` sheet. Map to that; do not create new labels.

6. **For major changes, propose a plan before editing code.** The user is
   learning Python. Explain what you're changing and why.

---

## Architecture

```
extractor_core.py            <- ALL extraction logic + output writers
extract_owner_financials.py  <- CLI wrapper (thin: prompts, file I/O only)
app.py                       <- Streamlit webapp (thin: upload, call core, download)
coa_mapper.py                <- COA mapping engine (rules-first, 5-step pipeline)
db_writer.py                 <- Writes accumulating CSVs to db_ready/
move_processed_files.py      <- Moves output/ -> completed/, input/ -> archive/
```

### Key design principle
`extractor_core.py` is the single source of truth for extraction logic.
Both the CLI and webapp import from it. Changes to extraction logic go there only.

### Folder layout
```
input/      <- drop source .xlsx files here (CLI workflow)
output/     <- CLI writes datapacks here
completed/  <- archive: output files after review
archive/    <- archive: processed input files
db_ready/   <- accumulating CSVs for database loading
reference/  <- sample/reference workbooks (not processed)
```

---

## Manager Format Profiles

### Managed By options
`Extra` | `Public Storage` | `CubeSmart` | `Other`

Each has its own approved_mappings and alias_mappings CSV files.
`Other` skips COA mapping entirely (manual review required).

### Extra Space (primary, fully implemented)
- Sheet prefix: `Rolling IS`, `Unit Rate`, `Ops Sum`, `Rent Roll`
- Dates: datetime objects or `"Feb 2025"` (space-separated)
- Property number: embedded in sheet name (e.g., `Rolling IS 7214`)
- COA files: `approved_mappings_exr.csv`, `alias_mappings_exr.csv`

### Public Storage (implemented, being validated)
- Sheet: `IS` (exact name, no prefix)
- Dates: `"Feb-2025"` (hyphen-separated) — normalised to `"Feb 2025"` by `format_date()`
- Skip section header rows: `Revenue`, `Contractually set fees`, `Other Expenses`, `Other items`
- Drop YTD column (col O) — excluded automatically (not a valid date value)
- Property number: parsed from cell B3 (e.g., `"77712 - Wentworth (Vacaville, CA)"`)
- Rent Roll: occupancy count only (column C, row 8 downward) — no rates/dates available
- Unit Rate / Ops Sum: not available in PS format
- COA files: `approved_mappings_ps.csv`, `alias_mappings_ps.csv`

### CubeSmart (stub — not yet implemented)
- COA files exist (`approved_mappings_cs.csv`) but are empty
- When implementing: follow the PS branching pattern in `process_workbook()`

---

## COA Mapping System

Five-step pipeline in `coa_mapper.py`:
1. Exact match against `approved_mappings_*.csv` (confidence 1.00)
2. Normalized match — strips GL codes like `(4000)`, lowercases (0.95)
3. Alias match via `alias_mappings_*.csv` (0.85)
4. Fuzzy match via difflib (0.50–0.84, always flagged for review)
5. No match (0.00)

`account_type` values: `Income` | `Expense` | `EXR_Rollup` | `PS_Rollup`
Rollup rows are subtotals calculated by the source system — tag them but
do not aggregate them in the model (double-counting risk).

`coa2` is only populated for four key rollup rows that represent structural
totals the proforma cares about:
`Net Rental Income`, `Total Operating Income`, `Total Operating Expense`, `Net Operating Income`

To improve mapping accuracy: edit `approved_mappings_*.csv` or `alias_mappings_*.csv`.
No code changes needed for routine mapping updates.

---

## Database Output (db_writer.py)

Five accumulating CSVs written to `db_ready/` per run:
- `fact_rolling_is.csv` — one row per account per month
- `fact_unit_rate_snapshot.csv` — occupancy metrics snapshot
- `fact_ops_sum.csv` — rental activity (EXR only)
- `fact_rent_roll_snapshot.csv` — tenant detail (EXR only)
- `etl_processing_log.csv` — extraction audit trail

**Schema stability rule:** Never rename, remove, or reorder existing columns.
Add new columns at the END only.

---

## Proforma Integration Intent

- `Rolling IS` → `Data Drop` sheet (columns: Actual/Budget, Entity, Account,
  Month, Year, Period, Amount, COA, COA 2, Type)
- `Unit Rate / occupancy` → `Inputs & Drivers` sheet
- COA mapping output respects the `COA Translation` sheet structure in the proforma
- `fact_rolling_is.csv` is a superset of what Data Drop needs

---

## Version History (brief)

- **v1/v2:** Monolithic script — all logic in one file
- **v3.0:** Refactored — `extractor_core.py` holds all logic; CLI and webapp
  are thin wrappers. COA mapper, db_writer, and Managed By routing added.
  Public Storage format support added.
