from __future__ import annotations

from pathlib import Path

from excel2py.exceptions import ExcelParseError
from excel2py.models import MacroData


def extract_macros(filepath: Path) -> list[MacroData]:
    """Extract VBA macros from an Excel file using oletools."""
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        raise ExcelParseError("oletools is required for macro extraction. Install with: pip install oletools")

    macros = []
    try:
        vba_parser = VBA_Parser(str(filepath))
        if vba_parser.detect_vba_macros():
            for filename, stream_path, vba_filename, vba_code in vba_parser.extract_all_macros():
                if vba_code.strip():
                    macro_type = "Module"
                    if "Class" in stream_path:
                        macro_type = "Class"
                    elif "Form" in stream_path:
                        macro_type = "Form"
                    macros.append(MacroData(
                        module_name=vba_filename or filename,
                        code=vba_code,
                        macro_type=macro_type,
                    ))
        vba_parser.close()
    except Exception as e:
        raise ExcelParseError(f"Failed to extract macros: {e}") from e

    return macros
