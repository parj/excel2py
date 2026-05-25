from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from excel2py.models import WorkbookData


class BaseParser(ABC):
    @abstractmethod
    def parse(self, filepath: Path) -> WorkbookData:
        """Parse an Excel file and return structured workbook data."""
