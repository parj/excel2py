from __future__ import annotations

import re
from dataclasses import dataclass

from excel2py.models import CellData, WorkbookData

_MAX_DATA_ROWS = 50

# Matches cell references like A1, $A1, A$1, $A$1 — but not standalone numbers
_CELL_REF_RE = re.compile(r"(\$?[A-Za-z]+)(\$?)(\d+)")


def _parse_address(address: str) -> tuple[str, int]:
    """Return (column_letters, row_number) from a cell address like 'D12'."""
    m = re.match(r"([A-Z]+)(\d+)", address)
    if m:
        return m.group(1), int(m.group(2))
    return address, 0


def _normalize_formula(formula: str, row: int) -> str:
    """Replace relative row numbers with row-offset tokens so that structurally
    identical fill-down formulas share the same normalised string.

    Examples (formula at row 5):
      =B5*C5      ->  =B{r+0}*C{r+0}
      =B4*C5      ->  =B{r-1}*C{r+0}
      =SUM($A$1:$A$5)  ->  =SUM($A$1:$A$5)   (absolute refs unchanged)
    """

    def _replace(m: re.Match) -> str:
        col = m.group(1)  # e.g. 'B' or '$B'
        row_abs = m.group(2)  # '$' if row is absolute, else ''
        ref_row = int(m.group(3))
        if row_abs == "$":
            return m.group(0)  # absolute row — keep verbatim
        offset = ref_row - row
        sign = f"+{offset}" if offset >= 0 else str(offset)
        return f"{col}{{r{sign}}}"

    return _CELL_REF_RE.sub(_replace, formula)


@dataclass
class _FormulaGroup:
    """One or more consecutive same-column cells sharing a formula pattern."""

    col: str
    start_row: int
    end_row: int
    representative_formula: str  # formula of the first cell in the group
    representative_value: object  # cached value of the first cell


def _group_formula_cells(formula_cells: list[CellData]) -> list[_FormulaGroup]:
    """Group fill-down formula cells so that 100 identical-pattern rows collapse
    into a single entry.  Only consecutive cells in the same column with the same
    normalised formula are merged.
    """
    if not formula_cells:
        return []

    groups: list[_FormulaGroup] = []
    cells_by_col: dict[str, list[CellData]] = {}

    # Bucket by column, preserving row order
    for cell in formula_cells:
        col, _ = _parse_address(cell.address)
        cells_by_col.setdefault(col, []).append(cell)

    for col, col_cells in cells_by_col.items():
        # Sort by row number
        col_cells.sort(key=lambda c: _parse_address(c.address)[1])

        group_start = col_cells[0]
        group_start_row = _parse_address(group_start.address)[1]
        prev_row = group_start_row
        prev_norm = _normalize_formula(group_start.formula or "", group_start_row)

        for cell in col_cells[1:]:
            _, row = _parse_address(cell.address)
            norm = _normalize_formula(cell.formula or "", row)

            if norm == prev_norm and row == prev_row + 1:
                # Extends current group
                prev_row = row
            else:
                # Flush current group
                groups.append(
                    _FormulaGroup(
                        col=col,
                        start_row=group_start_row,
                        end_row=prev_row,
                        representative_formula=group_start.formula or "",
                        representative_value=group_start.value,
                    )
                )
                group_start = cell
                group_start_row = row
                prev_row = row
                prev_norm = norm

        groups.append(
            _FormulaGroup(
                col=col,
                start_row=group_start_row,
                end_row=prev_row,
                representative_formula=group_start.formula or "",
                representative_value=group_start.value,
            )
        )

    # Restore sheet order: sort by (start_row, col)
    groups.sort(key=lambda g: (g.start_row, g.col))
    return groups


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
        "7. Use xlsb_reader to read the Excel file. It handles both .xlsx/.xlsm and .xlsb "
        "files with a unified API. Use the mandatory helper below for ALL sheet reading.\n"
        "     from xlsb_reader import XlsxWorkbook, XlsbWorkbook, col_to_letter\n"
        "     for sheet_name, formulas in wb.iter_formulas():\n"
        "         # formulas is {(row, col): formula_string} with 0-based indices\n"
        "8. No unused imports, unused variables, or unreachable functions — every symbol "
        "in the output must be referenced at least once.\n"
        "9. Accept the input Excel file path as sys.argv[1]. Write all output files "
        "(one CSV per sheet, named <sheet_name>.csv) to the current working directory.\n\n"
        "## Output Contract — your script will be verified against this exact logic\n\n"
        "The verifier builds ground-truth DataFrames using this function:\n\n"
        "```python\n"
        "def _build_sheet_df(values: dict) -> pd.DataFrame:\n"
        "    \"\"\"MANDATORY: copy this verbatim into your script and use it for every sheet.\n"
        "    Merged cells: xlsb_reader returns only the anchor cell value (no forward-fill).\n"
        "    Non-anchor positions are absent from the dict and will become '' after fillna.\"\"\"\n"
        "    import pandas as pd\n"
        "    if not values:\n"
        "        return pd.DataFrame()\n"
        "    max_row = max(r for r, _ in values)\n"
        "    max_col = max(c for _, c in values)\n"
        "    data = [[None] * (max_col + 1) for _ in range(max_row + 1)]\n"
        "    for (r, c), v in values.items():\n"
        "        if v is not None:  # keep '' — do NOT use `if v:` (drops empty strings)\n"
        "            data[r][c] = v\n"
        "    df = pd.DataFrame(data)\n"
        "    return df.dropna(how='all').dropna(axis=1, how='all').fillna('')\n"
        "```\n\n"
        "The verifier reads your CSV output with:\n"
        "    actual_df = pd.read_csv(your_csv, header=None)\n"
        "Therefore: ALWAYS write `df.to_csv(path, index=False, header=False)`. "
        "A header line written by pandas becomes an extra data row when read back with "
        "header=None, causing a shape mismatch of exactly +1 row.\n"
    )


def serialize_workbook(workbook: WorkbookData) -> str:
    parts: list[str] = []

    # Header
    parts.append(f"Excel Workbook: {workbook.filename}")
    parts.append(f"Format: {workbook.format}")
    parts.append("")

    # Named ranges
    if workbook.named_ranges:
        parts.append("=== Named Ranges ===")
        for name, ref in workbook.named_ranges.items():
            parts.append(f"  {name} = {ref}")
        parts.append("")

    # Sheets
    for sheet in workbook.sheets:
        parts.append(f"=== Sheet: {sheet.name} ===")
        if sheet.dimensions:
            parts.append(f"Dimensions: {sheet.dimensions}")
        if sheet.merged_ranges:
            parts.append(f"Merged ranges: {', '.join(sheet.merged_ranges)}")

        if sheet.cells:
            parts.append("")
            parts.append("Address | Value | Formula | Type")
            parts.append("--------|-------|---------|-----")

            formula_cells = [c for c in sheet.cells if c.formula]
            formula_addresses = {c.address for c in formula_cells}
            non_formula_cells = [
                c for c in sheet.cells if c.address not in formula_addresses
            ]

            # Deduplicate fill-down formula groups so that 100 rows of =B{n}*C{n}
            # appear as a single "D2:D101" entry rather than 100 individual rows.
            formula_groups = _group_formula_cells(formula_cells)

            remaining_slots = max(0, _MAX_DATA_ROWS - len(formula_groups))
            included_data = non_formula_cells[:remaining_slots]
            included_data_addresses = {c.address for c in included_data}

            # Emit deduplicated formula groups first (in sheet order)
            shown_non_formula = [
                c for c in sheet.cells if c.address in included_data_addresses
            ]

            # Build unified output in approximate sheet order
            # Interleave: emit a data row when its row comes before the next formula group
            data_iter = iter(shown_non_formula)
            group_iter = iter(formula_groups)

            pending_data = next(data_iter, None)
            pending_group = next(group_iter, None)

            while pending_data is not None or pending_group is not None:
                data_row = (
                    _parse_address(pending_data.address)[1]
                    if pending_data
                    else float("inf")
                )
                group_row = pending_group.start_row if pending_group else float("inf")

                if data_row <= group_row:
                    value_str = (
                        "" if pending_data.value is None else str(pending_data.value)
                    )
                    parts.append(
                        f"{pending_data.address} | {value_str} | | {pending_data.data_type}"
                    )
                    pending_data = next(data_iter, None)
                else:
                    g = pending_group
                    if g.start_row == g.end_row:
                        addr_str = f"{g.col}{g.start_row}"
                    else:
                        addr_str = f"{g.col}{g.start_row}:{g.col}{g.end_row}"
                        count = g.end_row - g.start_row + 1
                        # Annotate the representative formula so the LLM knows it repeats
                        parts.append(
                            f"{addr_str} | {'' if g.representative_value is None else g.representative_value}"
                            f" | {g.representative_formula} [{count} rows, same pattern] | formula"
                        )
                        pending_group = next(group_iter, None)
                        continue
                    parts.append(
                        f"{addr_str} | {'' if g.representative_value is None else g.representative_value}"
                        f" | {g.representative_formula} | formula"
                    )
                    pending_group = next(group_iter, None)

            total = len(sheet.cells)
            shown_formula_count = sum(
                g.end_row - g.start_row + 1 for g in formula_groups
            )
            shown_data_count = len(included_data)
            shown = shown_formula_count + shown_data_count
            if total > shown:
                parts.append(f"... ({total - shown} more rows omitted)")

        parts.append("")

    # Pivot tables
    if workbook.pivot_tables:
        parts.append("=== Pivot Tables ===")
        for pt in workbook.pivot_tables:
            parts.append(f"Sheet: {pt.sheet_name}")
            parts.append(f"  Source range : {pt.source_range}")
            parts.append(
                f"  Row fields   : {', '.join(pt.row_fields) if pt.row_fields else '(none)'}"
            )
            parts.append(
                f"  Column fields: {', '.join(pt.col_fields) if pt.col_fields else '(none)'}"
            )
            data_field_strs = []
            for f in pt.data_fields:
                agg = pt.aggregation_functions.get(f, "SUM")
                data_field_strs.append(f"{f} ({agg})")
            parts.append(
                f"  Data fields  : {', '.join(data_field_strs) if data_field_strs else '(none)'}"
            )
            parts.append(
                f"  Filter fields: {', '.join(pt.filter_fields) if pt.filter_fields else '(none)'}"
            )
            parts.append("")

    # Macros
    if workbook.macros:
        parts.append("=== VBA Macros ===")
        for macro in workbook.macros:
            parts.append(
                f"--- Module: {macro.module_name} (type: {macro.macro_type}) ---"
            )
            parts.append(macro.code)
            parts.append("")

    return "\n".join(parts)
