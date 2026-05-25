from unittest.mock import MagicMock, patch

import pytest

from excel2py.converter import convert
from excel2py.exceptions import CodeGenerationError, UnsupportedFormatError
from excel2py.verifier import VerificationError, VerificationResult


class TestConverter:
    def test_unsupported_format(self, tmp_path):
        bad = tmp_path / "test.csv"
        bad.write_text("a,b,c")
        with pytest.raises(UnsupportedFormatError):
            convert(bad)

    def test_dry_run(self, tmp_xlsx):
        """Dry run should return the prompt without calling LLM."""
        result = convert(tmp_xlsx, dry_run=True)
        assert "Sales" in result
        assert "test.xlsx" in result

    @patch("excel2py.converter.create_provider")
    def test_full_conversion(self, mock_create, tmp_xlsx, tmp_path):
        mock_provider = MagicMock()
        mock_response = MagicMock()
        mock_response.content = "import pandas as pd\nprint('converted')"
        mock_provider.generate.return_value = mock_response
        mock_create.return_value = mock_provider

        output = tmp_path / "output.py"
        settings = MagicMock()
        settings.default_provider = "openai"
        settings.openai_api_key = "test-key"
        settings.openai_model = "gpt-4o"
        settings.max_tokens = 4096
        settings.temperature = 0.2
        settings.correction_backend = "langchain"

        result = convert(
            tmp_xlsx,
            output_file=output,
            provider="openai",
            api_key="test-key",
            settings=settings,
            verify=False,
        )
        assert "import pandas" in result
        assert output.read_text() == result

    @patch("excel2py.converter.create_provider")
    def test_invalid_code_raises(self, mock_create, tmp_xlsx):
        mock_provider = MagicMock()
        mock_response = MagicMock()
        mock_response.content = "def broken(:"  # syntax error
        mock_provider.generate.return_value = mock_response
        mock_create.return_value = mock_provider

        settings = MagicMock()
        settings.default_provider = "openai"
        settings.openai_api_key = "key"
        settings.openai_model = "gpt-4o"
        settings.max_tokens = 4096
        settings.temperature = 0.2

        with pytest.raises(CodeGenerationError):
            convert(tmp_xlsx, provider="openai", api_key="key", settings=settings)


def _make_settings(correction_backend="langchain"):
    s = MagicMock()
    s.default_provider = "openai"
    s.openai_api_key = "test-key"
    s.openai_model = "gpt-4o"
    s.max_tokens = 4096
    s.temperature = 0.2
    s.correction_backend = correction_backend
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
    mock_provider = MagicMock()
    good_code = "import pandas as pd\nimport sys\nprint('fixed')"
    mock_provider.generate.return_value = MagicMock(content="import pandas as pd\nprint('broken')")
    mock_create.return_value = mock_provider

    failing_result = VerificationResult(
        passed=False,
        errors=[VerificationError("Sales", "row 0, col 0", 1, 2, "mismatch")],
    )
    passing_result = VerificationResult(passed=True)

    with patch("excel2py.converter.extract_ground_truth", return_value={"Sales": MagicMock()}):
        with patch("excel2py.converter.run_script", return_value=(0, "", "", tmp_path)):
            with patch("excel2py.converter.compare_outputs", side_effect=[failing_result, passing_result]):
                with patch("excel2py.converter.run_langchain_correction", return_value=good_code):
                    result = convert(
                        tmp_xlsx,
                        provider="openai",
                        api_key="test-key",
                        settings=_make_settings(),
                        verify=True,
                        max_verify_attempts=3,
                        verify_timeout=10,
                    )

    # only 1 call: initial generation (correction goes through LangGraph team)
    assert mock_provider.generate.call_count == 1
    assert "fixed" in result


@patch("excel2py.converter.create_provider")
def test_verify_returns_best_result_after_max_attempts(mock_create, tmp_xlsx, tmp_path):
    mock_provider = MagicMock()
    mock_provider.generate.return_value = MagicMock(content="import pandas as pd\n# attempt1")
    mock_create.return_value = mock_provider

    results = [
        VerificationResult(passed=False, errors=[MagicMock()] * 5),
        VerificationResult(passed=False, errors=[MagicMock()] * 2),
        VerificationResult(passed=False, errors=[MagicMock()] * 4),
    ]
    # broadcast corrections return attempt2 then attempt3; attempt2 wins (fewest errors)
    broadcast_returns = [
        "import pandas as pd\n# attempt2",
        "import pandas as pd\n# attempt3",
    ]

    with patch("excel2py.converter.extract_ground_truth", return_value={"S": MagicMock()}):
        with patch("excel2py.converter.run_script", return_value=(0, "", "", tmp_path)):
            with patch("excel2py.converter.compare_outputs", side_effect=results):
                with patch(
                    "excel2py.converter.run_langchain_correction",
                    side_effect=broadcast_returns,
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

    assert "attempt2" in result
