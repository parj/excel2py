from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class CellData:
    address: str
    value: Any
    formula: str | None = None
    data_type: str = "string"


@dataclass
class SheetData:
    name: str
    cells: list[CellData] = field(default_factory=list)
    merged_ranges: list[str] = field(default_factory=list)
    dimensions: str = ""


@dataclass
class PivotTableData:
    sheet_name: str
    source_range: str
    row_fields: list[str] = field(default_factory=list)
    col_fields: list[str] = field(default_factory=list)
    data_fields: list[str] = field(default_factory=list)
    filter_fields: list[str] = field(default_factory=list)
    aggregation_functions: dict[str, str] = field(default_factory=dict)


@dataclass
class MacroData:
    module_name: str
    code: str
    macro_type: str = "Module"


@dataclass
class WorkbookData:
    filename: str
    format: str
    sheets: list[SheetData] = field(default_factory=list)
    pivot_tables: list[PivotTableData] = field(default_factory=list)
    macros: list[MacroData] = field(default_factory=list)
    named_ranges: dict[str, str] = field(default_factory=dict)
