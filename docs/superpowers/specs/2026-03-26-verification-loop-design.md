# Verification-and-Correction Loop Design

## Problem

Generated Python scripts have three recurring quality issues:
1. **Unused code** — dead imports and unreachable functions
2. **Merged cells** — script crashes or produces NaN where Excel showed a value
3. **Incomplete pivot tables** — aggregation functions, filter values, and sort order not captured; silent extraction failures

## Solution

Approach C: upstream prompt/extraction fixes + a runtime verification-and-correction loop integrated into the `convert` pipeline. Disabled via `--no-verify`.

---

## Section 1 — Upstream Fixes

### System prompt additions (`prompts/templates.py:get_system_prompt`)
- Rule 7: For merged cells, forward-fill the anchor cell's value across all cells in the range when reading with pandas/openpyxl — never leave NaN where Excel showed a value.
- Rule 8: No unused imports, unused variables, or unreachable functions — every symbol in the output must be referenced.

### Pivot extraction fix (`parsers/xlsx_parser.py:_extract_pivot_tables`)
- Extend `PivotTableData` model with `aggregation_functions: dict[str, str]` (field name → "SUM"/"COUNT"/etc.)
- Read `subtotalFunctions` from each `dataField`
- Replace bare `except: pass` with `except Exception` that logs a warning

---

## Section 2 — Verification Module (`verifier.py`)

### Ground truth extraction
- Read Excel with `openpyxl` (`data_only=True`)
- Build `dict[str, SheetSnapshot]`: sheet name → `{cell_address: value}` for all non-empty cells
- For merged cells: record anchor cell value for all cells in the range
- Skip cells containing Excel errors (`#REF!`, `#VALUE!`, `#N/A`)
- Skip sheets with no data cells and no formulas

### Output comparison
- Read script output files back via pandas
- Match output files to sheets by column-name overlap + shape heuristic; multiple files matched by filename similarity to sheet name
- Compare cell-by-cell: numeric values within `rtol=1e-5`, strings exact, dates normalised to date only
- Produce `VerificationResult`: `passed: bool`, `errors: list[VerificationError]` (each: sheet, cell/row, expected, actual)

### Pass criteria
- Script exits with code 0
- All matched output values within tolerance
- No unmatched output sheets (script didn't skip a sheet entirely)

---

## Section 3 — Correction Loop (`converter.py`)

### Flow
```
generate → run script → verify → pass? → done
                   ↓ fail
          build correction prompt → regenerate → run script → verify → ...
```
Max attempts: 3 (configurable). Best-result tracker: keep the attempt with fewest errors.

### Running the script
- `subprocess.run` with configurable timeout (default 60s)
- Run in a temp directory; pass original Excel path as first argument
- Capture stdout, stderr, exit code

### Correction prompt content
- Original generated code
- Error type: crash / value mismatch / missing sheet
- For crashes: full traceback
- For mismatches: concise diff table capped at 20 rows (sheet, cell, expected, actual)
- Instruction: return only the corrected full Python file

### Failure handling
- After max attempts: write best result, emit warning listing remaining issues — never raise hard error

---

## Section 4 — CLI & Settings

### CLI additions (`cli.py`)
- `--no-verify` flag → `verify=False`
- `--max-verify-attempts N` (default 3)
- `--verbose` shows per-attempt summaries: `Attempt 1/3: 4 mismatches in Sheet1 (B2: expected 120.5, got 0)`

### Settings additions (`config.py`)
- `EXCEL2PY_VERIFY: bool = True`
- `EXCEL2PY_MAX_VERIFY_ATTEMPTS: int = 3`
- `EXCEL2PY_VERIFY_TIMEOUT: int = 60`

### `convert()` signature additions
- `verify: bool = True`
- `max_verify_attempts: int = 3`
- `verify_timeout: int = 60`

---

## Section 5 — Error Handling Edge Cases

| Situation | Handling |
|-----------|----------|
| Missing imports in script | Capture `ModuleNotFoundError` from stderr; include in correction prompt |
| Script hangs | Kill after timeout; treat as crash |
| Script writes no output files | Verification failure: "script produced no output" |
| Multiple output files | Match by column-name overlap + filename similarity to sheet name |
| Purely formatting sheets | Skip from comparison |
| Excel error cells (`#REF!` etc.) | Skip from ground truth |
| `--dry-run` | Verification never runs |

---

## Section 6 — Testing

### Unit tests
- `tests/test_verifier.py`: ground truth extraction, comparison logic (numeric tolerance, merged cell fill, error cell skipping)
- `tests/test_converter.py` additions: `verify=False` skips loop; best-result tracker; correction prompt construction

### Fixtures
- Synthetic `.xlsx` built with openpyxl in a fixture (not a binary file): formula column, merged cell range, one pivot table

### Integration tests
- `@pytest.mark.integration`: end-to-end convert with verification, assert passes in ≤ 3 attempts

---

## New Files

| File | Purpose |
|------|---------|
| `src/excel2py/verifier.py` | Ground truth extraction + output comparison |
| `tests/test_verifier.py` | Verifier unit tests |

## Modified Files

| File | Changes |
|------|---------|
| `src/excel2py/models.py` | Add `aggregation_functions` to `PivotTableData` |
| `src/excel2py/parsers/xlsx_parser.py` | Fix pivot extraction; capture aggregation functions |
| `src/excel2py/prompts/templates.py` | Add rules 7 and 8 to system prompt |
| `src/excel2py/converter.py` | Add verification loop; new params |
| `src/excel2py/config.py` | Add verify settings |
| `src/excel2py/cli.py` | Add `--no-verify`, `--max-verify-attempts` flags |
| `tests/test_converter.py` | New test cases for verification |
