from __future__ import annotations

from pathlib import Path

from xlsb_reader import XlsbWorkbook, col_to_letter

from excel2py.exceptions import ExcelParseError
from excel2py.models import CellData, MacroData, PivotTableData, SheetData, WorkbookData
from excel2py.parsers.base import BaseParser

_SUBTOTAL_MAP = {
    "sum": "SUM", "count": "COUNT", "average": "AVERAGE",
    "max": "MAX", "min": "MIN", "product": "PRODUCT",
    "countNums": "COUNT", "stdDev": "STDEV", "stdDevp": "STDEVP",
    "var": "VAR", "varp": "VARP",
}


def _to_address(row: int, col: int) -> str:
    return f"{col_to_letter(col)}{row + 1}"


def _determine_data_type(value, has_formula: bool) -> str:
    if has_formula:
        return "formula"
    if value is None:
        return "string"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    return "string"


def _parse_pivot(pt_dict: dict, sheet_name: str) -> PivotTableData:
    data_fields = []
    agg_functions: dict[str, str] = {}
    for df in pt_dict.get("data_fields", []):
        name = df.get("name", "")
        if name:
            data_fields.append(name)
            subtotal = str(df.get("subtotal", "sum"))
            agg_functions[name] = _SUBTOTAL_MAP.get(subtotal, subtotal.upper())

    location = pt_dict.get("location") or {}
    geom = location.get("rfx_geom") or {}
    tl = geom.get("top_left", "")
    br = geom.get("bottom_right", "")
    source_range = f"{tl}:{br}" if tl and br else ""

    return PivotTableData(
        sheet_name=sheet_name,
        source_range=source_range,
        row_fields=[],
        col_fields=[],
        data_fields=data_fields,
        filter_fields=[],
        aggregation_functions=agg_functions,
    )


class XlsbParser(BaseParser):
    """Parser for .xlsb Excel binary files using xlsb_reader."""

    def parse(self, filepath: Path) -> WorkbookData:
        try:
            return self._parse(filepath)
        except ExcelParseError:
            raise
        except Exception as e:
            raise ExcelParseError(f"Failed to parse {filepath}: {e}") from e

    def _parse(self, filepath: Path) -> WorkbookData:
        wb = XlsbWorkbook(filepath)

        values_by_sheet: dict[str, dict] = {
            name: vals for name, vals in wb.iter_values()
        }
        formulas_by_sheet: dict[str, dict] = {
            name: fmls for name, fmls in wb.iter_formulas()
        }

        sheets: list[SheetData] = []
        for sheet_name in wb.sheet_names:
            values = values_by_sheet.get(sheet_name, {})
            formulas = formulas_by_sheet.get(sheet_name, {})
            all_positions = set(values) | set(formulas)

            cells: list[CellData] = []
            for row, col in sorted(all_positions):
                formula = formulas.get((row, col))
                value = values.get((row, col))
                cells.append(CellData(
                    address=_to_address(row, col),
                    value=value,
                    formula=formula,
                    data_type=_determine_data_type(value, bool(formula)),
                ))

            if all_positions:
                rows = [r for r, _ in all_positions]
                cols = [c for _, c in all_positions]
                dimensions = f"{_to_address(min(rows), min(cols))}:{_to_address(max(rows), max(cols))}"
            else:
                dimensions = ""

            sheets.append(SheetData(
                name=sheet_name,
                cells=cells,
                merged_ranges=[],
                dimensions=dimensions,
            ))

        all_pivots = [
            _parse_pivot(pt, pt.get("sheet", ""))
            for pt in wb.iter_pivot_tables()
        ]

        macros: list[MacroData] = []  # xlsb_reader does not expose VBA modules

        return WorkbookData(
            filename=filepath.name,
            format="xlsb",
            sheets=sheets,
            pivot_tables=all_pivots,
            macros=macros,
            named_ranges={},
        )
