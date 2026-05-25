import pytest
from excel2py.parsers.xlsb_parser import XlsbParser
from excel2py.exceptions import ExcelParseError


class TestXlsbParser:
    def test_instantiation(self):
        parser = XlsbParser()
        assert parser is not None

    def test_invalid_file_raises(self, tmp_path):
        bad_file = tmp_path / "bad.xlsb"
        bad_file.write_text("not a real xlsb file")
        parser = XlsbParser()
        with pytest.raises(ExcelParseError):
            parser.parse(bad_file)
