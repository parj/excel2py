from unittest.mock import patch

from click.testing import CliRunner
from excel2py.cli import main


class TestCLI:
    def test_version(self):
        runner = CliRunner()
        result = runner.invoke(main, ["--version"])
        assert result.exit_code == 0
        assert "0.1.0" in result.output

    def test_convert_dry_run(self, tmp_xlsx):
        runner = CliRunner()
        result = runner.invoke(main, ["convert", str(tmp_xlsx), "--dry-run"])
        assert result.exit_code == 0
        assert "Sales" in result.output

    def test_convert_missing_file(self):
        runner = CliRunner()
        result = runner.invoke(main, ["convert", "nonexistent.xlsx"])
        assert result.exit_code != 0


def test_no_verify_flag_accepted(tmp_xlsx):
    runner = CliRunner()
    with patch("excel2py.cli.convert") as mock_convert:
        mock_convert.return_value = "print('ok')"
        result = runner.invoke(main, ["convert", str(tmp_xlsx), "--no-verify"])
    assert result.exit_code == 0
    call_kwargs = mock_convert.call_args[1]
    assert call_kwargs.get("verify") is False


def test_max_verify_attempts_flag(tmp_xlsx):
    runner = CliRunner()
    with patch("excel2py.cli.convert") as mock_convert:
        mock_convert.return_value = "print('ok')"
        result = runner.invoke(
            main, ["convert", str(tmp_xlsx), "--max-verify-attempts", "5"]
        )
    assert result.exit_code == 0
    call_kwargs = mock_convert.call_args[1]
    assert call_kwargs.get("max_verify_attempts") == 5
