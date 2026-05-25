from __future__ import annotations


import openpyxl
import pandas as pd


class TestExtractGroundTruth:
    def test_returns_dict_keyed_by_sheet_name(self, tmp_xlsx_full):
        from excel2py.verifier import extract_ground_truth
        result = extract_ground_truth(tmp_xlsx_full)
        assert "Sales" in result

    def test_data_cells_captured(self, tmp_xlsx_full):
        from excel2py.verifier import extract_ground_truth
        result = extract_ground_truth(tmp_xlsx_full)
        df = result["Sales"]
        assert df.shape[0] > 0
        assert df.shape[1] > 0

    def test_merged_cell_anchor_value_present(self, tmp_xlsx_full):
        from excel2py.verifier import extract_ground_truth
        result = extract_ground_truth(tmp_xlsx_full)
        df = result["Sales"]
        first_row_values = [str(v) for v in df.iloc[0].tolist() if v is not None]
        assert any("Sales Report" in v for v in first_row_values)

    def test_empty_sheet_excluded(self, tmp_path):
        from excel2py.verifier import extract_ground_truth
        filepath = tmp_path / "empty.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Empty"
        wb.save(filepath)
        wb.close()
        result = extract_ground_truth(filepath)
        assert "Empty" not in result

    def test_excel_error_cells_skipped(self, tmp_path):
        from excel2py.verifier import extract_ground_truth
        filepath = tmp_path / "errors.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "#REF!"
        ws["A2"] = 42
        wb.save(filepath)
        wb.close()
        result = extract_ground_truth(filepath)
        if "Sheet1" in result:
            df = result["Sheet1"]
            flat = df.values.flatten().tolist()
            assert "#REF!" not in flat


class TestCompareOutputs:
    def test_no_output_files_with_empty_ground_truth_passes(self, tmp_path):
        from excel2py.verifier import compare_outputs
        result = compare_outputs(tmp_path, {})
        assert result.passed

    def test_no_output_files_with_ground_truth_fails(self, tmp_path):
        from excel2py.verifier import compare_outputs
        ground_truth = {"Sales": pd.DataFrame([[1, 2], [3, 4]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert not result.passed
        assert any(e.error_type == "no_output" for e in result.errors)

    def test_matching_csv_passes(self, tmp_path):
        from excel2py.verifier import compare_outputs
        csv_file = tmp_path / "Sales.csv"
        csv_file.write_text("Widget,10.0\nGadget,25.0\n")
        ground_truth = {"Sales": pd.DataFrame([["Widget", 10.0], ["Gadget", 25.0]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert result.passed

    def test_value_mismatch_fails(self, tmp_path):
        from excel2py.verifier import compare_outputs
        csv_file = tmp_path / "Sales.csv"
        csv_file.write_text("Widget,99.0\nGadget,25.0\n")
        ground_truth = {"Sales": pd.DataFrame([["Widget", 10.0], ["Gadget", 25.0]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert not result.passed
        assert any(e.error_type == "mismatch" for e in result.errors)

    def test_missing_sheet_fails(self, tmp_path):
        from excel2py.verifier import compare_outputs
        ground_truth = {"Sales": pd.DataFrame([[1, 2]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert not result.passed
        assert any(e.error_type in ("no_output", "missing_sheet") for e in result.errors)

    def test_numeric_tolerance(self, tmp_path):
        from excel2py.verifier import compare_outputs
        csv_file = tmp_path / "Sheet1.csv"
        csv_file.write_text("1.000000001\n")
        ground_truth = {"Sheet1": pd.DataFrame([[1.0]])}
        result = compare_outputs(tmp_path, ground_truth)
        assert result.passed


class TestRunScript:
    def test_successful_script_returns_zero(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "ok.py"
        script.write_text("import sys\nprint('ok')\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=10)
        assert exit_code == 0
        assert "ok" in stdout

    def test_crashing_script_returns_nonzero(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "crash.py"
        script.write_text("raise RuntimeError('boom')\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=10)
        assert exit_code != 0
        assert "boom" in stderr

    def test_timeout_returns_minus_one(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "hang.py"
        script.write_text("import time\ntime.sleep(999)\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=1)
        assert exit_code == -1
        assert "timed out" in stderr.lower()

    def test_output_files_written_to_output_dir(self, tmp_path):
        from excel2py.verifier import run_script
        script = tmp_path / "writer.py"
        script.write_text("open('out.csv', 'w').write('a,b\\n1,2\\n')\n")
        xlsx = tmp_path / "dummy.xlsx"
        xlsx.write_bytes(b"")
        exit_code, stdout, stderr, output_dir = run_script(script, xlsx, timeout=10)
        assert exit_code == 0
        assert (output_dir / "out.csv").exists()


class TestBuildCorrectionPrompt:
    def test_includes_original_code(self):
        from excel2py.verifier import VerificationResult, build_correction_prompt
        result = VerificationResult(passed=False, errors=[])
        prompt = build_correction_prompt("print('hello')", result, 0, "")
        assert "print('hello')" in prompt

    def test_includes_crash_stderr(self):
        from excel2py.verifier import VerificationResult, build_correction_prompt
        result = VerificationResult(passed=False, errors=[])
        prompt = build_correction_prompt("x", result, 1, "NameError: name 'x'")
        assert "NameError" in prompt

    def test_includes_mismatch_table(self):
        from excel2py.verifier import VerificationError, VerificationResult, build_correction_prompt
        errors = [VerificationError("Sales", "row 0, col 1", 10.0, 99.0, "mismatch")]
        result = VerificationResult(passed=False, errors=errors)
        prompt = build_correction_prompt("code", result, 0, "")
        assert "Sales" in prompt
        assert "10.0" in prompt
        assert "99.0" in prompt
