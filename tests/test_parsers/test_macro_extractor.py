from excel2py.parsers.macro_extractor import extract_macros


class TestMacroExtractor:
    def test_non_macro_file(self, tmp_xlsx):
        """Extracting macros from a plain xlsx should return an empty list."""
        result = extract_macros(tmp_xlsx)
        assert result == []
