from __future__ import annotations

from excel2py.models import CellData, PivotTableData, SheetData, WorkbookData
from excel2py.prompts.templates import _group_formula_cells, serialize_workbook


def _make_workbook(cells: list[CellData], sheet_name: str = "Sheet1") -> WorkbookData:
    return WorkbookData(
        filename="test.xlsx",
        format="xlsx",
        sheets=[SheetData(name=sheet_name, cells=cells)],
    )


class TestGroupFormulaCells:
    def test_single_cell_no_group(self):
        cells = [CellData("D2", 50, "=B2*C2", "formula")]
        groups = _group_formula_cells(cells)
        assert len(groups) == 1
        assert groups[0].start_row == 2
        assert groups[0].end_row == 2

    def test_100_identical_pattern_collapse_to_one(self):
        cells = [
            CellData(f"D{r}", r * 10, f"=B{r}*C{r}", "formula")
            for r in range(2, 102)  # rows 2-101
        ]
        groups = _group_formula_cells(cells)
        assert len(groups) == 1
        assert groups[0].start_row == 2
        assert groups[0].end_row == 101
        assert groups[0].representative_formula == "=B2*C2"

    def test_two_different_formulas_two_groups(self):
        # Rows 2-101: =B{n}*C{n}, row 102: =SUM(D2:D101)
        cells = [
            CellData(f"D{r}", r, f"=B{r}*C{r}", "formula") for r in range(2, 102)
        ] + [CellData("D102", 5050, "=SUM(D2:D101)", "formula")]
        groups = _group_formula_cells(cells)
        assert len(groups) == 2
        assert groups[0].end_row == 101
        assert groups[1].start_row == 102
        assert groups[1].representative_formula == "=SUM(D2:D101)"

    def test_non_consecutive_rows_not_merged(self):
        # Rows 2 and 4 have same formula pattern but row 3 is missing
        cells = [
            CellData("D2", 10, "=B2*C2", "formula"),
            CellData("D4", 30, "=B4*C4", "formula"),
        ]
        groups = _group_formula_cells(cells)
        assert len(groups) == 2

    def test_different_columns_separate_groups(self):
        cells = [
            CellData("D2", 10, "=B2*C2", "formula"),
            CellData("D3", 20, "=B3*C3", "formula"),
            CellData("E2", 5, "=A2+1", "formula"),
            CellData("E3", 6, "=A3+1", "formula"),
        ]
        groups = _group_formula_cells(cells)
        assert len(groups) == 2  # one D group, one E group

    def test_absolute_refs_do_not_normalise_away(self):
        # Both rows reference $A$1 (absolute) - still same pattern (trivially)
        cells = [
            CellData("D2", 100, "=$A$1*B2", "formula"),
            CellData("D3", 100, "=$A$1*B3", "formula"),
        ]
        groups = _group_formula_cells(cells)
        assert len(groups) == 1

    def test_mixed_absolute_breaks_group(self):
        # Row 2: =$A$1*B2  vs  Row 3: =$A$2*B3 — different absolute row => different patterns
        cells = [
            CellData("D2", 10, "=$A$1*B2", "formula"),
            CellData("D3", 20, "=$A$2*B3", "formula"),
        ]
        groups = _group_formula_cells(cells)
        assert len(groups) == 2


class TestSerializeWorkbook:
    def test_deduped_group_shows_range_address(self):
        cells = [CellData(f"D{r}", r, f"=B{r}*C{r}", "formula") for r in range(2, 102)]
        wb = _make_workbook(cells)
        output = serialize_workbook(wb)
        assert "D2:D101" in output
        # Should NOT have 100 separate D entries
        assert output.count("=B") == 1

    def test_single_formula_row_no_range_suffix(self):
        cells = [CellData("D2", 50, "=B2*C2", "formula")]
        wb = _make_workbook(cells)
        output = serialize_workbook(wb)
        assert "D2 |" in output
        assert "D2:D" not in output

    def test_two_groups_both_appear(self):
        cells = [
            CellData(f"D{r}", r, f"=B{r}*C{r}", "formula") for r in range(2, 102)
        ] + [CellData("D102", 5050, "=SUM(D2:D101)", "formula")]
        wb = _make_workbook(cells)
        output = serialize_workbook(wb)
        assert "D2:D101" in output
        assert "=SUM(D2:D101)" in output

    def test_non_formula_data_still_included(self):
        cells = [
            CellData("A1", "Header", None, "string"),
            CellData("D2", 50, "=B2*C2", "formula"),
        ]
        wb = _make_workbook(cells)
        output = serialize_workbook(wb)
        assert "Header" in output
        assert "=B2*C2" in output


def test_system_prompt_merged_cell_rule():
    from excel2py.prompts.templates import get_system_prompt

    prompt = get_system_prompt()
    assert "merged" in prompt.lower()
    assert "forward-fill" in prompt.lower() or "forward fill" in prompt.lower()


def test_system_prompt_no_dead_code_rule():
    from excel2py.prompts.templates import get_system_prompt

    prompt = get_system_prompt()
    assert "unused" in prompt.lower()


def test_system_prompt_argv_rule():
    from excel2py.prompts.templates import get_system_prompt

    prompt = get_system_prompt()
    assert "sys.argv[1]" in prompt


def test_pivot_serialization_includes_aggregation():
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
