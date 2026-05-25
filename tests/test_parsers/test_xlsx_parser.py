from excel2py.parsers.xlsx_parser import XlsxParser


class TestXlsxParser:
    def test_parse_basic(self, tmp_xlsx):
        parser = XlsxParser()
        result = parser.parse(tmp_xlsx)
        assert result.filename == "test.xlsx"
        assert result.format == "xlsx"
        assert len(result.sheets) == 1
        sheet = result.sheets[0]
        assert sheet.name == "Sales"

    def test_formulas_extracted(self, tmp_xlsx):
        parser = XlsxParser()
        result = parser.parse(tmp_xlsx)
        sheet = result.sheets[0]
        formula_cells = [c for c in sheet.cells if c.formula]
        assert len(formula_cells) == 3
        formulas = {c.address: c.formula for c in formula_cells}
        assert formulas["D2"] == "=B2*C2"
        assert formulas["D3"] == "=B3*C3"
        assert formulas["D4"] == "=SUM(D2:D3)"

    def test_data_types(self, tmp_xlsx):
        parser = XlsxParser()
        result = parser.parse(tmp_xlsx)
        sheet = result.sheets[0]
        cell_map = {c.address: c for c in sheet.cells}
        assert cell_map["A1"].data_type == "string"
        assert cell_map["B2"].data_type == "number"
        assert cell_map["D2"].data_type == "formula"

    def test_named_ranges(self, tmp_xlsx_with_named_range):
        parser = XlsxParser()
        result = parser.parse(tmp_xlsx_with_named_range)
        assert "MyRange" in result.named_ranges

    def test_merged_cells(self, tmp_xlsx_merged):
        parser = XlsxParser()
        result = parser.parse(tmp_xlsx_merged)
        sheet = result.sheets[0]
        assert len(sheet.merged_ranges) > 0
        assert "A1:C1" in sheet.merged_ranges

    def test_unsupported_file(self, tmp_path):
        from excel2py.exceptions import ExcelParseError

        bad_file = tmp_path / "test.xlsx"
        bad_file.write_text("not an excel file")
        parser = XlsxParser()
        import pytest

        with pytest.raises(ExcelParseError):
            parser.parse(bad_file)
