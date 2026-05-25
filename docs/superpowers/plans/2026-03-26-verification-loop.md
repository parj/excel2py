# Verification-and-Correction Loop Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a runtime verification-and-correction loop to `excel2py convert` that runs the generated script, compares its output against Excel cached values, and asks the LLM to fix discrepancies — up to 3 attempts.

**Architecture:** Three layers of improvement: (1) upstream fixes to prompt and pivot extraction so the LLM gets better input, (2) a new `verifier.py` module that extracts ground truth from Excel and compares script output files, (3) a correction loop in `converter.py` that feeds failures back to the LLM. The loop is on by default and disabled via `--no-verify`.

**Tech Stack:** openpyxl, pandas, subprocess, pytest, click, pydantic-settings

---

## File Map

| File | Action | Responsibility |
|------|--------|---------------|
| `src/excel2py/models.py` | Modify | Add `aggregation_functions` to `PivotTableData` |
| `src/excel2py/parsers/xlsx_parser.py` | Modify | Read aggregation functions from pivot dataFields; replace silent `except: pass` |
| `src/excel2py/prompts/templates.py` | Modify | Add rules 7–9 to system prompt; serialize aggregation_functions |
| `src/excel2py/verifier.py` | Create | Ground truth extraction, script runner, output comparison, correction prompt |
| `src/excel2py/converter.py` | Modify | Extract `_strip_fences` helper; add verify loop with new params |
| `src/excel2py/config.py` | Modify | Add `verify`, `max_verify_attempts`, `verify_timeout` settings |
| `src/excel2py/cli.py` | Modify | Add `--no-verify` and `--max-verify-attempts` flags |
| `tests/conftest.py` | Modify | Add `tmp_xlsx_full` fixture (formula + merged cells) |
| `tests/test_verifier.py` | Create | Unit tests for all verifier functions |
| `tests/test_converter.py` | Modify | Tests for verify=False, best-result tracking, correction prompt |
| `tests/test_parsers/test_xlsx_parser.py` | Modify | Test aggregation_functions extraction |
| `tests/test_prompts/test_templates.py` | Modify | Test rules 7–9 present in system prompt |

---

## Task 1: Add `aggregation_functions` to `PivotTableData` and serialization

**Files:**
- Modify: `src/excel2py/models.py`
- Modify: `src/excel2py/prompts/templates.py`
- Test: `tests/test_prompts/test_templates.py`

- [ ] **Step 1: Write the failing test**

Add to `tests/test_prompts/test_templates.py`:

```python
from excel2py.models import PivotTableData

def test_pivot_serialization_includes_aggregation():
    from excel2py.prompts.templates import serialize_workbook
    wb = WorkbookData(
        filename="test.xlsx",
        format="xlsx",
        sheets=[],
        pivot_tables=[
            PivotTableData(
                sheet_name="Sheet1",
                source_range="Sheet1!A1:D10",
                row_fields=["Region"],
                col_fields=[],
                data_fields=["Sales"],
                filter_fields=[],
                aggregation_functions={"Sales": "SUM"},
            )
        ],
    )
    result = serialize_workbook(wb)
    assert "SUM" in result
    assert "Sales" in result
```

- [ ] **Step 2: Run test to verify it fails**

```bash
cd /Users/pmudu/sandbox/convert-excel-v5 && source .venv/bin/activate && pytest tests/test_prompts/test_templates.py::test_pivot_serialization_includes_aggregation -v
```

Expected: `TypeError` or `unexpected keyword argument 'aggregation_functions'`

- [ ] **Step 3: Add `aggregation_functions` to `PivotTableData` in `models.py`**

```python
@dataclass
class PivotTableData:
    sheet_name: str
    source_range: str
    row_fields: list[str] = field(default_factory=list)
    col_fields: list[str] = field(default_factory=list)
    data_fields: list[str] = field(default_factory=list)
    filter_fields: list[str] = field(default_factory=list)
    aggregation_functions: dict[str, str] = field(default_factory=dict)
```

- [ ] **Step 4: Update pivot serialization in `prompts/templates.py`**

Replace the pivot table serialization block (lines 224–233) with:

```python
    if workbook.pivot_tables:
        parts.append("=== Pivot Tables ===")
        for pt in workbook.pivot_tables:
            parts.append(f"Sheet: {pt.sheet_name}")
            parts.append(f"  Source range : {pt.source_range}")
            parts.append(f"  Row fields   : {', '.join(pt.row_fields) if pt.row_fields else '(none)'}")
            parts.append(f"  Column fields: {', '.join(pt.col_fields) if pt.col_fields else '(none)'}")
            data_field_strs = []
            for f in pt.data_fields:
                agg = pt.aggregation_functions.get(f, "SUM")
                data_field_strs.append(f"{f} ({agg})")
            parts.append(f"  Data fields  : {', '.join(data_field_strs) if data_field_strs else '(none)'}")
            parts.append(f"  Filter fields: {', '.join(pt.filter_fields) if pt.filter_fields else '(none)'}")
            parts.append("")
```

- [ ] **Step 5: Run test to verify it passes**

```bash
pytest tests/test_prompts/test_templates.py::test_pivot_serialization_includes_aggregation -v
```

Expected: PASS

- [ ] **Step 6: Run full test suite to check for regressions**

```bash
pytest tests/ -v --ignore=tests/test_llm -x
```

Expected: all pass

- [ ] **Step 7: Commit**

```bash
git add src/excel2py/models.py src/excel2py/prompts/templates.py tests/test_prompts/test_templates.py
git commit -m "feat: add aggregation_functions to PivotTableData and include in serialization"
```

---

## Task 2: Fix pivot extraction to capture aggregation functions

**Files:**
- Modify: `src/excel2py/parsers/xlsx_parser.py`
- Test: `tests/test_parsers/test_xlsx_parser.py`

- [ ] **Step 1: Write the failing test**

Add to `tests/test_parsers/test_xlsx_parser.py`:

```python
def test_pivot_aggregation_functions_default_sum(tmp_xlsx):
    """Even without a real pivot, PivotTableData.aggregation_functions has no silent failures."""
    from excel2py.models import PivotTableData
    # Directly test that _extract_pivot_tables returns empty list gracefully
    from excel2py.parsers.xlsx_parser import XlsxParser
    parser = XlsxParser()
    result = parser.parse(tmp_xlsx)
    # No pivot tables in tmp_xlsx — should return empty list without crashing
    assert result.pivot_tables == []
```

- [ ] **Step 2: Run test to verify it passes already (baseline)**

```bash
pytest tests/test_parsers/test_xlsx_parser.py::TestXlsxParser::test_pivot_aggregation_functions_default_sum -v
```

Expected: PASS (establishes baseline before the real fix)

- [ ] **Step 3: Update `_extract_pivot_tables` in `xlsx_parser.py`**

Replace the entire `_extract_pivot_tables` function with this version that captures aggregation functions and replaces silent exception swallowing with logged warnings:

```python
import logging as _logging
_logger = _logging.getLogger(__name__)

_SUBTOTAL_FUNCTION_MAP = {
    "sum": "SUM", "count": "COUNT", "average": "AVERAGE",
    "max": "MAX", "min": "MIN", "product": "PRODUCT",
    "countNums": "COUNT", "stdDev": "STDEV", "stdDevp": "STDEVP",
    "var": "VAR", "varp": "VARP",
}


def _extract_pivot_tables(ws, sheet_name: str) -> list[PivotTableData]:
    pivots = []
    raw_pivots = getattr(ws, "_pivots", None)
    if not raw_pivots:
        return pivots

    for pivot in raw_pivots:
        source_range = ""
        try:
            cache_def = pivot.cache
            ws_source = getattr(cache_def, "cacheSource", None)
            if ws_source is not None:
                ws_ref = getattr(ws_source, "worksheetSource", None)
                if ws_ref is not None:
                    ref = getattr(ws_ref, "ref", "") or ""
                    src_sheet = getattr(ws_ref, "sheet", "") or ""
                    source_range = f"'{src_sheet}'!{ref}" if src_sheet and ref else ref
        except Exception as e:
            _logger.warning("Could not extract pivot source range: %s", e)

        row_fields: list[str] = []
        col_fields: list[str] = []
        data_fields: list[str] = []
        filter_fields: list[str] = []
        aggregation_functions: dict[str, str] = {}

        try:
            cache_def = pivot.cache
            field_names: list[str] = []
            cache_fields = getattr(cache_def, "cacheFields", None)
            if cache_fields is not None:
                for cf in cache_fields:
                    field_names.append(getattr(cf, "name", "") or "")

            piv_fields = getattr(pivot, "rowFields", None)
            if piv_fields is not None:
                for rf in piv_fields:
                    idx = getattr(rf, "x", None)
                    if idx is not None and 0 <= idx < len(field_names):
                        row_fields.append(field_names[idx])

            piv_col_fields = getattr(pivot, "colFields", None)
            if piv_col_fields is not None:
                for cf in piv_col_fields:
                    idx = getattr(cf, "x", None)
                    if idx is not None and 0 <= idx < len(field_names):
                        col_fields.append(field_names[idx])

            piv_data_fields = getattr(pivot, "dataFields", None)
            if piv_data_fields is not None:
                for df in piv_data_fields:
                    name = getattr(df, "name", None)
                    if not name:
                        idx = getattr(df, "fld", None)
                        if idx is not None and 0 <= idx < len(field_names):
                            name = field_names[idx]
                    if name:
                        data_fields.append(name)
                        raw_fn = getattr(df, "subtotalFunction", "sum") or "sum"
                        aggregation_functions[name] = _SUBTOTAL_FUNCTION_MAP.get(raw_fn, raw_fn.upper())

            piv_page_fields = getattr(pivot, "pageFields", None)
            if piv_page_fields is not None:
                for pf in piv_page_fields:
                    idx = getattr(pf, "fld", None)
                    if idx is not None and 0 <= idx < len(field_names):
                        filter_fields.append(field_names[idx])
        except Exception as e:
            _logger.warning("Could not extract pivot field details: %s", e)

        pivots.append(
            PivotTableData(
                sheet_name=sheet_name,
                source_range=source_range,
                row_fields=row_fields,
                col_fields=col_fields,
                data_fields=data_fields,
                filter_fields=filter_fields,
                aggregation_functions=aggregation_functions,
            )
        )

    return pivots
```

- [ ] **Step 4: Run all parser tests**

```bash
pytest tests/test_parsers/ -v
```

Expected: all pass

- [ ] **Step 5: Commit**

```bash
git add src/excel2py/parsers/xlsx_parser.py tests/test_parsers/test_xlsx_parser.py
git commit -m "fix: capture pivot aggregation functions; replace silent except with logged warnings"
```

---

## Task 3: Improve system prompt (rules 7–9)

**Files:**
- Modify: `src/excel2py/prompts/templates.py`
- Test: `tests/test_prompts/test_templates.py`

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_prompts/test_templates.py`:

```python
from excel2py.prompts.templates import get_system_prompt

def test_system_prompt_merged_cell_rule():
    prompt = get_system_prompt()
    assert "merged" in prompt.lower()
    assert "forward-fill" in prompt.lower() or "forward fill" in prompt.lower()

def test_system_prompt_no_dead_code_rule():
    prompt = get_system_prompt()
    assert "unused" in prompt.lower()

def test_system_prompt_argv_rule():
    prompt = get_system_prompt()
    assert "sys.argv[1]" in prompt
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_prompts/test_templates.py::test_system_prompt_merged_cell_rule tests/test_prompts/test_templates.py::test_system_prompt_no_dead_code_rule tests/test_prompts/test_templates.py::test_system_prompt_argv_rule -v
```

Expected: FAIL

- [ ] **Step 3: Update `get_system_prompt` in `prompts/templates.py`**

Replace the `get_system_prompt` function body with:

```python
def get_system_prompt() -> str:
    return (
        "You are an expert Python developer specialising in converting Microsoft Excel "
        "spreadsheets to clean, idiomatic Python code.\n\n"
        "Rules you MUST follow:\n"
        "1. Use pandas for all tabular data manipulation and pivot-table reproduction.\n"
        "2. Reproduce every Excel formula as an equivalent Python function or inline "
        "expression; do not hard-code computed values.\n"
        "3. Convert any VBA macros to plain Python functions with equivalent logic.\n"
        "4. Add concise inline comments explaining non-obvious logic, especially for "
        "formula conversions and pivot-table construction.\n"
        "5. Output ONLY valid Python source code — no markdown fences, no prose, no "
        "explanations outside of Python comments.\n"
        "6. The output must be self-contained and runnable: import every library you use "
        "and define all helper functions before they are called.\n"
        "7. For merged cells, forward-fill the anchor cell's value across all cells in "
        "the merged range after reading — never leave NaN where Excel showed a value. "
        "Use: df.ffill() after reading merged regions, or explicitly fill with the anchor value.\n"
        "8. No unused imports, unused variables, or unreachable functions — every symbol "
        "in the output must be referenced at least once.\n"
        "9. Accept the input Excel file path as sys.argv[1]. Write all output files "
        "(one CSV per sheet, named <sheet_name>.csv) to the current working directory.\n"
    )
```

- [ ] **Step 4: Run the three new tests**

```bash
pytest tests/test_prompts/test_templates.py::test_system_prompt_merged_cell_rule tests/test_prompts/test_templates.py::test_system_prompt_no_dead_code_rule tests/test_prompts/test_templates.py::test_system_prompt_argv_rule -v
```

Expected: all PASS

- [ ] **Step 5: Run full templates test suite**

```bash
pytest tests/test_prompts/ -v
```

Expected: all pass

- [ ] **Step 6: Commit**

```bash
git add src/excel2py/prompts/templates.py tests/test_prompts/test_templates.py
git commit -m "feat: add merged-cell, no-dead-code, and argv rules to system prompt"
```

---

## Task 4: Create `verifier.py` — ground truth extraction

**Files:**
- Create: `src/excel2py/verifier.py`
- Create: `tests/test_verifier.py`
- Modify: `tests/conftest.py`

- [ ] **Step 1: Add `tmp_xlsx_full` fixture to `tests/conftest.py`**

Append to `tests/conftest.py`:

```python
@pytest.fixture
def tmp_xlsx_full(tmp_path):
    """Excel with data, formulas, and merged cells for verifier tests."""
    filepath = tmp_path / "full.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"

    # Row 1: merged header
    ws["A1"] = "Sales Report"
    ws.merge_cells("A1:C1")

    # Row 2: column headers
    ws["A2"] = "Product"
    ws["B2"] = "Price"
    ws["C2"] = "Total"

    # Row 3-4: data + formula
    ws["A3"] = "Widget"
    ws["B3"] = 10.0
    ws["C3"] = "=B3*2"

    ws["A4"] = "Gadget"
    ws["B4"] = 25.0
    ws["C4"] = "=B4*2"

    wb.save(filepath)
    wb.close()
    return filepath
```

- [ ] **Step 2: Write failing tests for `extract_ground_truth`**

Create `tests/test_verifier.py`:

```python
from __future__ import annotations

from pathlib import Path

import pytest
import openpyxl
import pandas as pd


class TestExtractGroundTruth:
    def test_returns_dict_keyed_by_sheet_name(self, tmp_xlsx_full):
        from excel2py.verifier import extract_ground_truth
        result = extract_ground_truth(tmp_xlsx_full)
        assert "Sales" in result

    def test_data_cells_captured(self, tmp_xlsx_full):
        from excel2py.verifier import extract_ground_truth
        result = extract_ground_truth(tmp_xlsx_full)
        df = result["Sales"]
        # DataFrame should have rows and columns
        assert df.shape[0] > 0
        assert df.shape[1] > 0

    def test_merged_cell_anchor_value_present(self, tmp_xlsx_full):
        from excel2py.verifier import extract_ground_truth
        result = extract_ground_truth(tmp_xlsx_full)
        df = result["Sales"]
        # "Sales Report" should appear somewhere in the first row
        first_row_values = [str(v) for v in df.iloc[0].tolist() if v is not None]
        assert any("Sales Report" in v for v in first_row_values)

    def test_empty_sheet_excluded(self, tmp_path):
        from excel2py.verifier import extract_ground_truth
        filepath = tmp_path / "empty.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Empty"
        wb.save(filepath)
        wb.close()
        result = extract_ground_truth(filepath)
        assert "Empty" not in result

    def test_excel_error_cells_skipped(self, tmp_path):
        from excel2py.verifier import extract_ground_truth
        filepath = tmp_path / "errors.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "#REF!"
        ws["A2"] = 42
        wb.save(filepath)
        wb.close()
        result = extract_ground_truth(filepath)
        if "Sheet1" in result:
            df = result["Sheet1"]
            flat = df.values.flatten().tolist()
            assert "#REF!" not in flat
```

- [ ] **Step 3: Run tests to verify they fail**

```bash
pytest tests/test_verifier.py::TestExtractGroundTruth -v
```

Expected: `ModuleNotFoundError: No module named 'excel2py.verifier'`

- [ ] **Step 4: Create `src/excel2py/verifier.py` with `extract_ground_truth`**

```python
from __future__ import annotations

import logging
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

_EXCEL_ERRORS = {"#REF!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!", "#DIV/0!"}


@dataclass
class VerificationError:
    sheet: str
    location: str
    expected: Any
    actual: Any
    error_type: str  # "mismatch" | "missing_sheet" | "crash" | "no_output" | "shape_mismatch"


@dataclass
class VerificationResult:
    passed: bool
    errors: list[VerificationError] = field(default_factory=list)

    def error_count(self) -> int:
        return len(self.errors)


def extract_ground_truth(excel_path: Path) -> dict[str, pd.DataFrame]:
    """Extract cached cell values from Excel as DataFrames, forward-filling merged cells."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    result: dict[str, pd.DataFrame] = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Build merged cell fill map: non-anchor address -> anchor value
        merged_fill: dict[str, Any] = {}
        for merged_range in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            anchor_val = ws.cell(row=min_row, column=min_col).value
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    if row == min_row and col == min_col:
                        continue
                    addr = f"{get_column_letter(col)}{row}"
                    merged_fill[addr] = anchor_val

        rows = []
        has_data = False
        for ws_row in ws.iter_rows():
            row_data = []
            for cell in ws_row:
                addr = f"{get_column_letter(cell.column)}{cell.row}"
                val = cell.value
                if isinstance(val, str) and val in _EXCEL_ERRORS:
                    val = None
                if val is None and addr in merged_fill:
                    val = merged_fill[addr]
                if val is not None:
                    has_data = True
                row_data.append(val)
            rows.append(row_data)

        if has_data:
            df = pd.DataFrame(rows).dropna(how="all").dropna(axis=1, how="all")
            result[sheet_name] = df

    wb.close()
    return result
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
pytest tests/test_verifier.py::TestExtractGroundTruth -v
```

Expected: all PASS

- [ ] **Step 6: Commit**

```bash
git add src/excel2py/verifier.py tests/test_verifier.py tests/conftest.py
git commit -m "feat: add verifier.py with extract_ground_truth; add tmp_xlsx_full fixture"
```

---

## Task 5: Add output comparison to `verifier.py`

**Files:**
- Modify: `src/excel2py/verifier.py`
- Modify: `tests/test_verifier.py`

- [ ] **Step 1: Write failing tests for `compare_outputs`**

Append to `tests/test_verifier.py`:

```python
class TestCompareOutputs:
    def test_no_output_files_fails(self, tmp_path):
        from excel2py.verifier import compare_outputs
        result = compare_outputs(tmp_path, {})
        # Empty ground truth + empty output = trivially passed (nothing to compare)
        assert result.passed

    def test_no_output_files_with_ground_truth_fails(self, tmp_path):
        from excel2py.verifier import compare_outputs, extract_ground_truth
        import pandas as pd
        ground_truth = {"Sales": pd.DataFrame([[1, 2], [3, 4]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert not result.passed
        assert any(e.error_type == "no_output" for e in result.errors)

    def test_matching_csv_passes(self, tmp_path):
        from excel2py.verifier import compare_outputs
        import pandas as pd
        # Write a CSV that matches the ground truth
        csv_file = tmp_path / "Sales.csv"
        csv_file.write_text("Widget,10.0\nGadget,25.0\n")
        ground_truth = {"Sales": pd.DataFrame([["Widget", 10.0], ["Gadget", 25.0]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert result.passed

    def test_value_mismatch_fails(self, tmp_path):
        from excel2py.verifier import compare_outputs
        import pandas as pd
        csv_file = tmp_path / "Sales.csv"
        csv_file.write_text("Widget,99.0\nGadget,25.0\n")
        ground_truth = {"Sales": pd.DataFrame([["Widget", 10.0], ["Gadget", 25.0]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert not result.passed
        assert any(e.error_type == "mismatch" for e in result.errors)

    def test_missing_sheet_fails(self, tmp_path):
        from excel2py.verifier import compare_outputs
        import pandas as pd
        # No CSV for "Sales"
        ground_truth = {"Sales": pd.DataFrame([[1, 2]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert not result.passed
        assert any(e.error_type in ("no_output", "missing_sheet") for e in result.errors)

    def test_numeric_tolerance(self, tmp_path):
        from excel2py.verifier import compare_outputs
        import pandas as pd
        csv_file = tmp_path / "Sheet1.csv"
        csv_file.write_text("1.000000001\n")
        ground_truth = {"Sheet1": pd.DataFrame([[1.0]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert result.passed
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_verifier.py::TestCompareOutputs -v
```

Expected: `AttributeError` or `ImportError` — `compare_outputs` does not exist yet

- [ ] **Step 3: Add `_find_matching_file`, `_compare_dataframes`, and `compare_outputs` to `verifier.py`**

Append to `src/excel2py/verifier.py`:

```python
def _find_matching_file(output_files: list[Path], sheet_name: str) -> Path | None:
    sheet_key = sheet_name.lower().replace(" ", "_")
    for f in output_files:
        if f.stem.lower().replace(" ", "_") == sheet_key:
            return f
    for f in output_files:
        stem = f.stem.lower()
        if sheet_key in stem or stem in sheet_key:
            return f
    if len(output_files) == 1:
        return output_files[0]
    return None


def _compare_dataframes(
    expected: pd.DataFrame, actual: pd.DataFrame, sheet_name: str
) -> list[VerificationError]:
    errors: list[VerificationError] = []
    expected = expected.reset_index(drop=True)
    actual = actual.reset_index(drop=True)

    if expected.shape != actual.shape:
        errors.append(VerificationError(
            sheet=sheet_name,
            location="shape",
            expected=str(expected.shape),
            actual=str(actual.shape),
            error_type="shape_mismatch",
        ))
        return errors

    rows = min(len(expected), 50)
    cols = min(len(expected.columns), 20)

    for r in range(rows):
        for c in range(cols):
            exp_val = expected.iat[r, c]
            act_val = actual.iat[r, c]

            exp_is_na = exp_val is None or (isinstance(exp_val, float) and pd.isna(exp_val))
            act_is_na = act_val is None or (isinstance(act_val, float) and pd.isna(act_val))
            if exp_is_na and act_is_na:
                continue

            try:
                if abs(float(exp_val) - float(act_val)) <= abs(float(exp_val)) * 1e-5 + 1e-9:
                    continue
            except (TypeError, ValueError):
                if str(exp_val) == str(act_val):
                    continue

            errors.append(VerificationError(
                sheet=sheet_name,
                location=f"row {r}, col {c}",
                expected=exp_val,
                actual=act_val,
                error_type="mismatch",
            ))
            if len(errors) >= 20:
                return errors

    return errors


def compare_outputs(
    output_dir: Path, ground_truth: dict[str, pd.DataFrame]
) -> VerificationResult:
    if not ground_truth:
        return VerificationResult(passed=True)

    output_files = list(output_dir.glob("*.csv")) + list(output_dir.glob("*.xlsx"))

    if not output_files:
        return VerificationResult(
            passed=False,
            errors=[VerificationError(
                sheet="", location="", expected="output files", actual="none written",
                error_type="no_output",
            )],
        )

    errors: list[VerificationError] = []

    for sheet_name, expected_df in ground_truth.items():
        output_file = _find_matching_file(output_files, sheet_name)
        if output_file is None:
            errors.append(VerificationError(
                sheet=sheet_name, location="", expected="output file", actual="not found",
                error_type="missing_sheet",
            ))
            continue

        try:
            if output_file.suffix == ".csv":
                actual_df = pd.read_csv(output_file, header=None)
            else:
                actual_df = pd.read_excel(output_file, header=None)
        except Exception as e:
            errors.append(VerificationError(
                sheet=sheet_name, location="", expected="readable file", actual=str(e),
                error_type="crash",
            ))
            continue

        errors.extend(_compare_dataframes(expected_df, actual_df, sheet_name))

    return VerificationResult(passed=len(errors) == 0, errors=errors)
```

- [ ] **Step 4: Run tests**

```bash
pytest tests/test_verifier.py::TestCompareOutputs -v
```

Expected: all PASS

- [ ] **Step 5: Run full verifier test suite**

```bash
pytest tests/test_verifier.py -v
```

Expected: all PASS

- [ ] **Step 6: Commit**

```bash
git add src/excel2py/verifier.py tests/test_verifier.py
git commit -m "feat: add compare_outputs to verifier with numeric tolerance and sheet matching"
```

---

## Task 6: Add `run_script` and `build_correction_prompt` to `verifier.py`

**Files:**
- Modify: `src/excel2py/verifier.py`
- Modify: `tests/test_verifier.py`

- [ ] **Step 1: Write failing tests**

Append to `tests/test_verifier.py`:

```python
class TestRunScript:
    def test_successful_script_returns_zero(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "ok.py"
        script.write_text("import sys\nprint('ok')\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")  # dummy file — script doesn't read it
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=10)
        assert exit_code == 0
        assert "ok" in stdout

    def test_crashing_script_returns_nonzero(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "crash.py"
        script.write_text("raise RuntimeError('boom')\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=10)
        assert exit_code != 0
        assert "boom" in stderr

    def test_timeout_returns_minus_one(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "hang.py"
        script.write_text("import time\ntime.sleep(999)\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=1)
        assert exit_code == -1
        assert "timed out" in stderr.lower()

    def test_output_files_written_to_output_dir(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "writer.py"
        script.write_text("open('out.csv', 'w').write('a,b\\n1,2\\n')\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=10)
        assert exit_code == 0
        assert (output_dir / "out.csv").exists()


class TestBuildCorrectionPrompt:
    def test_includes_original_code(self):
        from excel2py.verifier import VerificationResult, build_correction_prompt
        result = VerificationResult(passed=False, errors=[])
        prompt = build_correction_prompt("print('hello')", result, 0, "")
        assert "print('hello')" in prompt

    def test_includes_crash_stderr(self):
        from excel2py.verifier import VerificationResult, build_correction_prompt
        result = VerificationResult(passed=False, errors=[])
        prompt = build_correction_prompt("x", result, 1, "NameError: name 'x'")
        assert "NameError" in prompt

    def test_includes_mismatch_table(self):
        from excel2py.verifier import VerificationError, VerificationResult, build_correction_prompt
        errors = [VerificationError("Sales", "row 0, col 1", 10.0, 99.0, "mismatch")]
        result = VerificationResult(passed=False, errors=errors)
        prompt = build_correction_prompt("code", result, 0, "")
        assert "Sales" in prompt
        assert "10.0" in prompt
        assert "99.0" in prompt
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_verifier.py::TestRunScript tests/test_verifier.py::TestBuildCorrectionPrompt -v
```

Expected: `AttributeError` — `run_script` and `build_correction_prompt` not yet defined

- [ ] **Step 3: Add `run_script` and `build_correction_prompt` to `verifier.py`**

Append to `src/excel2py/verifier.py`:

```python
def run_script(
    script_path: Path, excel_path: Path, timeout: int = 60
) -> tuple[int, str, str, Path]:
    """Run the generated script in a temp dir. Returns (exit_code, stdout, stderr, output_dir)."""
    output_dir = Path(tempfile.mkdtemp())
    try:
        proc = subprocess.run(
            ["python", str(script_path), str(excel_path)],
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=str(output_dir),
        )
        return proc.returncode, proc.stdout, proc.stderr, output_dir
    except subprocess.TimeoutExpired:
        return -1, "", f"Script timed out after {timeout}s", output_dir


def build_correction_prompt(
    code: str, result: VerificationResult, exit_code: int, stderr: str
) -> str:
    lines = [
        "The Python script you generated has issues. Fix it and return ONLY the corrected Python file.",
        "",
        "## Issues",
    ]

    if exit_code != 0:
        lines.append(f"\n### Script crashed (exit code {exit_code})")
        lines.append("```")
        lines.append(stderr[:2000])
        lines.append("```")

    if result.errors:
        lines.append(
            f"\n### Verification errors ({len(result.errors)} total, showing first 20)"
        )
        lines.append("| Sheet | Location | Expected | Actual | Type |")
        lines.append("|-------|----------|----------|--------|------|")
        for err in result.errors[:20]:
            lines.append(
                f"| {err.sheet} | {err.location} | {err.expected} | {err.actual} | {err.error_type} |"
            )

    lines.append("\n## Original Code")
    lines.append("```python")
    lines.append(code)
    lines.append("```")
    lines.append(
        "\nReturn ONLY the corrected Python code, no markdown fences, no explanations."
    )

    return "\n".join(lines)
```

- [ ] **Step 4: Run all verifier tests**

```bash
pytest tests/test_verifier.py -v
```

Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add src/excel2py/verifier.py tests/test_verifier.py
git commit -m "feat: add run_script and build_correction_prompt to verifier"
```

---

## Task 7: Wire verification loop into `converter.py`

**Files:**
- Modify: `src/excel2py/converter.py`
- Modify: `tests/test_converter.py`

- [ ] **Step 1: Write failing tests**

Append to `tests/test_converter.py`:

```python
from unittest.mock import patch, MagicMock
from excel2py.verifier import VerificationResult, VerificationError


def _make_settings():
    s = MagicMock()
    s.default_provider = "openai"
    s.openai_api_key = "test-key"
    s.openai_model = "gpt-4o"
    s.max_tokens = 4096
    s.temperature = 0.2
    return s


@patch("excel2py.converter.create_provider")
def test_verify_false_skips_loop(mock_create, tmp_xlsx, tmp_path):
    mock_provider = MagicMock()
    mock_response = MagicMock()
    mock_response.content = "import pandas as pd\nprint('ok')"
    mock_provider.generate.return_value = mock_response
    mock_create.return_value = mock_provider

    with patch("excel2py.converter.extract_ground_truth") as mock_gt:
        result = convert(
            tmp_xlsx,
            provider="openai",
            api_key="test-key",
            settings=_make_settings(),
            verify=False,
        )
        mock_gt.assert_not_called()

    assert "import pandas" in result


@patch("excel2py.converter.create_provider")
def test_verify_loop_calls_llm_on_failure(mock_create, tmp_xlsx, tmp_path):
    """If first attempt fails verification, LLM is called again for correction."""
    mock_provider = MagicMock()
    good_code = "import pandas as pd\nimport sys\nprint('fixed')"
    mock_provider.generate.side_effect = [
        MagicMock(content="import pandas as pd\nprint('broken')"),  # initial
        MagicMock(content=good_code),  # correction
    ]
    mock_create.return_value = mock_provider

    failing_result = VerificationResult(
        passed=False,
        errors=[VerificationError("Sales", "row 0, col 0", 1, 2, "mismatch")],
    )
    passing_result = VerificationResult(passed=True)

    with patch("excel2py.converter.extract_ground_truth", return_value={"Sales": MagicMock()}):
        with patch("excel2py.converter.run_script", return_value=(0, "", "", tmp_path)):
            with patch(
                "excel2py.converter.compare_outputs",
                side_effect=[failing_result, passing_result],
            ):
                result = convert(
                    tmp_xlsx,
                    provider="openai",
                    api_key="test-key",
                    settings=_make_settings(),
                    verify=True,
                    max_verify_attempts=3,
                    verify_timeout=10,
                )

    assert mock_provider.generate.call_count == 2
    assert "fixed" in result


@patch("excel2py.converter.create_provider")
def test_verify_returns_best_result_after_max_attempts(mock_create, tmp_xlsx, tmp_path):
    """After max attempts, returns the attempt with fewest errors."""
    mock_provider = MagicMock()
    mock_provider.generate.side_effect = [
        MagicMock(content="import pandas as pd\n# attempt1"),
        MagicMock(content="import pandas as pd\n# attempt2"),
        MagicMock(content="import pandas as pd\n# attempt3"),
    ]
    mock_create.return_value = mock_provider

    results = [
        VerificationResult(passed=False, errors=[MagicMock()] * 5),   # 5 errors
        VerificationResult(passed=False, errors=[MagicMock()] * 2),   # 2 errors (best)
        VerificationResult(passed=False, errors=[MagicMock()] * 4),   # 4 errors
    ]

    with patch("excel2py.converter.extract_ground_truth", return_value={"S": MagicMock()}):
        with patch("excel2py.converter.run_script", return_value=(0, "", "", tmp_path)):
            with patch("excel2py.converter.compare_outputs", side_effect=results):
                result = convert(
                    tmp_xlsx,
                    provider="openai",
                    api_key="test-key",
                    settings=_make_settings(),
                    verify=True,
                    max_verify_attempts=3,
                    verify_timeout=10,
                )

    assert "attempt2" in result
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_converter.py::test_verify_false_skips_loop tests/test_converter.py::test_verify_loop_calls_llm_on_failure tests/test_converter.py::test_verify_returns_best_result_after_max_attempts -v
```

Expected: `TypeError` — `convert()` does not accept `verify` parameter yet

- [ ] **Step 3: Rewrite `converter.py`**

Replace the entire content of `src/excel2py/converter.py`:

```python
from __future__ import annotations

import ast
import logging
import tempfile
from pathlib import Path

from excel2py.config import Settings, get_settings
from excel2py.exceptions import CodeGenerationError, UnsupportedFormatError
from excel2py.llm.base import LLMRequest
from excel2py.llm.factory import create_provider
from excel2py.prompts.templates import get_system_prompt, serialize_workbook

logger = logging.getLogger(__name__)

PARSER_MAP = {
    ".xlsx": "excel2py.parsers.xlsx_parser:XlsxParser",
    ".xlsm": "excel2py.parsers.xlsx_parser:XlsxParser",
    ".xlsb": "excel2py.parsers.xlsb_parser:XlsbParser",
}


def _get_parser(ext: str):
    entry = PARSER_MAP.get(ext)
    if not entry:
        raise UnsupportedFormatError(f"Unsupported format: {ext}. Supported: {list(PARSER_MAP)}")
    module_path, class_name = entry.rsplit(":", 1)
    import importlib
    module = importlib.import_module(module_path)
    return getattr(module, class_name)()


def _get_api_key(settings: Settings, provider: str) -> str:
    key = getattr(settings, f"{provider}_api_key", None)
    if not key:
        raise ValueError(
            f"No API key configured for provider '{provider}'. "
            f"Set EXCEL2PY_{provider.upper()}_API_KEY"
        )
    return key


def _get_model(settings: Settings, provider: str) -> str:
    return getattr(settings, f"{provider}_model")


def _strip_fences(code: str) -> str:
    if code.startswith("```python"):
        code = code[len("```python"):].strip()
    if code.startswith("```"):
        code = code[3:].strip()
    if code.endswith("```"):
        code = code[:-3].strip()
    return code


def convert(
    input_file: Path,
    output_file: Path | None = None,
    provider: str | None = None,
    model: str | None = None,
    api_key: str | None = None,
    max_tokens: int | None = None,
    dry_run: bool = False,
    verify: bool = True,
    max_verify_attempts: int = 3,
    verify_timeout: int = 60,
    settings: Settings | None = None,
) -> str:
    """Convert an Excel file to a Python script.

    Returns the generated Python code.
    """
    settings = settings or get_settings()
    provider = provider or settings.default_provider

    # Parse Excel file
    ext = input_file.suffix.lower()
    parser = _get_parser(ext)
    logger.info("Parsing %s with %s", input_file, type(parser).__name__)
    workbook = parser.parse(input_file)

    # Build prompt
    system_prompt = get_system_prompt()
    user_prompt = serialize_workbook(workbook)

    if dry_run:
        logger.info("Dry run — prompt generated but not sent to LLM")
        return user_prompt

    # Call LLM (initial generation)
    api_key = api_key or _get_api_key(settings, provider)
    model = model or _get_model(settings, provider)
    llm = create_provider(provider, api_key, model)

    tokens = max_tokens or settings.max_tokens

    def _generate(prompt: str) -> str:
        request = LLMRequest(
            system_prompt=system_prompt,
            user_prompt=prompt,
            max_tokens=tokens,
            temperature=settings.temperature,
        )
        response = llm.generate(request)
        code = _strip_fences(response.content.strip())
        try:
            ast.parse(code)
        except SyntaxError as e:
            raise CodeGenerationError(f"Generated code has syntax errors: {e}") from e
        return code

    logger.info("Sending to %s (model: %s)", provider, model)
    code = _generate(user_prompt)

    # Verification-and-correction loop
    if verify:
        from excel2py.verifier import (
            build_correction_prompt,
            compare_outputs,
            extract_ground_truth,
            run_script,
        )

        ground_truth = extract_ground_truth(input_file)
        best_code = code
        best_error_count: float = float("inf")
        last_result = None

        for attempt in range(1, max_verify_attempts + 1):
            with tempfile.NamedTemporaryFile(suffix=".py", mode="w", delete=False) as f:
                f.write(code)
                script_path = Path(f.name)

            try:
                exit_code, stdout, stderr, output_dir = run_script(
                    script_path, input_file, verify_timeout
                )

                if exit_code != 0:
                    from excel2py.verifier import VerificationError, VerificationResult
                    last_result = VerificationResult(
                        passed=False,
                        errors=[VerificationError(
                            sheet="", location="", expected="exit 0",
                            actual=f"exit {exit_code}", error_type="crash",
                        )],
                    )
                else:
                    last_result = compare_outputs(output_dir, ground_truth)

                error_count = last_result.error_count()
                logger.info(
                    "Verification attempt %d/%d: %s (%d errors)",
                    attempt, max_verify_attempts,
                    "PASS" if last_result.passed else "FAIL",
                    error_count,
                )

                if error_count < best_error_count:
                    best_error_count = error_count
                    best_code = code

                if last_result.passed:
                    break

                if attempt < max_verify_attempts:
                    correction = build_correction_prompt(code, last_result, exit_code, stderr)
                    try:
                        code = _generate(correction)
                    except CodeGenerationError:
                        pass  # Keep previous code if correction has syntax error
            finally:
                script_path.unlink(missing_ok=True)

        if last_result is not None and not last_result.passed:
            logger.warning(
                "Verification incomplete after %d attempts. %d issues remain.",
                max_verify_attempts,
                best_error_count,
            )
        code = best_code

    # Write output
    if output_file:
        output_file.write_text(code)
        logger.info("Written to %s", output_file)

    return code
```

- [ ] **Step 4: Run the three new tests**

```bash
pytest tests/test_converter.py::test_verify_false_skips_loop tests/test_converter.py::test_verify_loop_calls_llm_on_failure tests/test_converter.py::test_verify_returns_best_result_after_max_attempts -v
```

Expected: all PASS

- [ ] **Step 5: Run full converter test suite**

```bash
pytest tests/test_converter.py -v
```

Expected: all PASS

- [ ] **Step 6: Commit**

```bash
git add src/excel2py/converter.py tests/test_converter.py
git commit -m "feat: add verification-and-correction loop to convert(); extract _strip_fences helper"
```

---

## Task 8: Add settings and CLI flags

**Files:**
- Modify: `src/excel2py/config.py`
- Modify: `src/excel2py/cli.py`
- Modify: `tests/test_cli.py`

- [ ] **Step 1: Write failing tests**

Read `tests/test_cli.py` first, then append:

```python
def test_no_verify_flag_accepted(tmp_xlsx, tmp_path):
    from click.testing import CliRunner
    from excel2py.cli import main
    from unittest.mock import patch, MagicMock

    runner = CliRunner()
    with patch("excel2py.converter.convert") as mock_convert:
        mock_convert.return_value = "print('ok')"
        result = runner.invoke(main, ["convert", str(tmp_xlsx), "--no-verify"])
    assert result.exit_code == 0
    call_kwargs = mock_convert.call_args[1]
    assert call_kwargs.get("verify") is False


def test_max_verify_attempts_flag(tmp_xlsx, tmp_path):
    from click.testing import CliRunner
    from excel2py.cli import main
    from unittest.mock import patch

    runner = CliRunner()
    with patch("excel2py.converter.convert") as mock_convert:
        mock_convert.return_value = "print('ok')"
        result = runner.invoke(main, ["convert", str(tmp_xlsx), "--max-verify-attempts", "5"])
    assert result.exit_code == 0
    call_kwargs = mock_convert.call_args[1]
    assert call_kwargs.get("max_verify_attempts") == 5
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_cli.py::test_no_verify_flag_accepted tests/test_cli.py::test_max_verify_attempts_flag -v
```

Expected: FAIL — `--no-verify` not recognised

- [ ] **Step 3: Add verify settings to `config.py`**

```python
class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_prefix="EXCEL2PY_")

    default_provider: str = "openai"
    openai_api_key: str | None = None
    openai_model: str = "gpt-4o"
    anthropic_api_key: str | None = None
    anthropic_model: str = "claude-sonnet-4-20250514"
    google_api_key: str | None = None
    google_model: str = "gemini-2.5-flash"
    openrouter_api_key: str | None = None
    openrouter_model: str = "openai/gpt-4o"
    max_tokens: int = 8192
    temperature: float = 0.2
    max_retries: int = 3
    log_level: str = "INFO"
    verify: bool = True
    max_verify_attempts: int = 3
    verify_timeout: int = 60
```

- [ ] **Step 4: Update `cli.py` to add `--no-verify` and `--max-verify-attempts`**

```python
from __future__ import annotations

import logging
import sys
from pathlib import Path

import click

from excel2py.converter import convert


@click.group()
@click.version_option(package_name="excel2py")
def main():
    """excel2py - Convert Excel spreadsheets to Python scripts using GenAI."""


@main.command()
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.option("-o", "--output", "output_file", type=click.Path(path_type=Path), default=None,
              help="Output Python file path. Defaults to <input>_converted.py")
@click.option("-p", "--provider", type=click.Choice(["openai", "anthropic", "google", "openrouter"]),
              default=None, help="LLM provider to use")
@click.option("-m", "--model", default=None, help="Model override")
@click.option("--api-key", default=None, help="API key (overrides env config)")
@click.option("--max-tokens", type=int, default=None, help="Max output tokens")
@click.option("--dry-run", is_flag=True, help="Parse Excel and print prompt without calling LLM")
@click.option("--no-verify", "no_verify", is_flag=True,
              help="Skip verification-and-correction loop")
@click.option("--max-verify-attempts", type=int, default=None,
              help="Max correction attempts (default 3)")
@click.option("-v", "--verbose", is_flag=True, help="Enable debug logging")
def convert_cmd(
    input_file, output_file, provider, model, api_key, max_tokens,
    dry_run, no_verify, max_verify_attempts, verbose,
):
    """Convert an Excel spreadsheet to a Python script."""
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if output_file is None:
        output_file = input_file.with_name(f"{input_file.stem}_converted.py")

    from excel2py.config import get_settings
    settings = get_settings()

    try:
        result = convert(
            input_file=input_file,
            output_file=None if dry_run else output_file,
            provider=provider,
            model=model,
            api_key=api_key,
            max_tokens=max_tokens,
            dry_run=dry_run,
            verify=not no_verify,
            max_verify_attempts=max_verify_attempts if max_verify_attempts is not None else settings.max_verify_attempts,
            verify_timeout=settings.verify_timeout,
            settings=settings,
        )
        if dry_run:
            click.echo(result)
        else:
            click.echo(f"Converted {input_file} -> {output_file}")
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)
```

- [ ] **Step 5: Run the two new CLI tests**

```bash
pytest tests/test_cli.py::test_no_verify_flag_accepted tests/test_cli.py::test_max_verify_attempts_flag -v
```

Expected: PASS

- [ ] **Step 6: Run full CLI test suite**

```bash
pytest tests/test_cli.py -v
```

Expected: all PASS

- [ ] **Step 7: Run full test suite**

```bash
pytest tests/ -v --ignore=tests/test_llm
```

Expected: all PASS

- [ ] **Step 8: Commit**

```bash
git add src/excel2py/config.py src/excel2py/cli.py tests/test_cli.py
git commit -m "feat: add --no-verify and --max-verify-attempts CLI flags; add verify settings"
```

---

## Task 9: Final smoke test

- [ ] **Step 1: Run the complete test suite including LLM unit tests**

```bash
pytest tests/ -v -x
```

Expected: all PASS (integration tests excluded — they require real API keys)

- [ ] **Step 2: Verify dry-run still works end-to-end**

```bash
excel2py convert tests/../tests/conftest.py --dry-run 2>&1 | head -5 || true
```

(This will fail on a non-Excel file as expected — just confirms the CLI loads without import errors.)

Actually run this instead to confirm CLI loads:

```bash
excel2py --help
excel2py convert --help
```

Expected: help text includes `--no-verify` and `--max-verify-attempts`

- [ ] **Step 3: Commit if any final fixes needed, otherwise tag**

```bash
git log --oneline -8
```

Expected: 8 commits covering Tasks 1–8
