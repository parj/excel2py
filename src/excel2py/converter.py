from __future__ import annotations

import ast
import hashlib
import logging
import tempfile
from pathlib import Path

from langchain_core.messages import HumanMessage, SystemMessage

from excel2py.config import Settings, get_settings
from excel2py.exceptions import CodeGenerationError, UnsupportedFormatError
from excel2py.llm.agno_team import run_broadcast_correction
from excel2py.llm.factory import create_chat_model
from excel2py.llm.langchain_team import (
    build_model_invoke,
    run_langchain_correction,
    run_output_judge,
    run_rubber_duck_diagnosis,
)
from excel2py.prompts.templates import get_system_prompt, serialize_workbook
from excel2py.verifier import (  # noqa: E402
    build_fresh_generation_prompt,
    build_last_resort_prompt,
    build_value_mismatch_last_resort_prompt,
    check_properties,
    compute_output_diagnostics,
)

logger = logging.getLogger(__name__)

PARSER_MAP = {
    ".xlsx": "excel2py.parsers.xlsx_parser:XlsxParser",
    ".xlsm": "excel2py.parsers.xlsx_parser:XlsxParser",
    ".xlsb": "excel2py.parsers.xlsb_parser:XlsbParser",
}


def _get_parser(ext: str):
    entry = PARSER_MAP.get(ext)
    if not entry:
        raise UnsupportedFormatError(
            f"Unsupported format: {ext}. Supported: {list(PARSER_MAP)}"
        )
    module_path, class_name = entry.rsplit(":", 1)
    import importlib

    module = importlib.import_module(module_path)
    return getattr(module, class_name)()


def _get_api_key(settings: Settings, provider: str) -> str:
    key = getattr(settings, f"{provider}_api_key", None)
    if not key:
        raise ValueError(
            f"No API key configured for provider '{provider}'. "
            f"Set EXCEL2PY_{provider.upper()}_API_KEY"
        )
    return key


def _get_model(settings: Settings, provider: str) -> str:
    return getattr(settings, f"{provider}_model")


# Error-type weights for best-code selection.
# Shape/structural errors are orders of magnitude harder to fix than value errors,
# so a script that gets the structure right (even with 60 wrong values) is a better
# base for the next correction than a script with 3 wrong shapes.
_ERROR_WEIGHTS: dict[str, float] = {
    "no_output": 100_000,  # wrote nothing — blind
    "missing_sheet": 50_000,  # sheet absent — structural
    "shape_mismatch": 10_000,  # shape wrong — structural
    "crash": 1_000_000,  # didn't run — worst
    "mismatch": 1,  # value wrong — easy fix
}


def _score_result(result, exit_code: int) -> float:
    """Weighted error score: lower = better. Structural errors massively outweigh value errors."""
    if exit_code != 0:
        return 1_000_000 + len(result.errors)
    return sum(_ERROR_WEIGHTS.get(e.error_type, 1) for e in result.errors)


def _extract_required_shapes(ground_truth: dict) -> dict[str, tuple[int, int]]:
    import pandas as pd

    return {
        name: (int(df.shape[0]), int(df.shape[1]))
        for name, df in ground_truth.items()
        if isinstance(df, pd.DataFrame)
    }


def _get_env_info() -> str:
    import sys
    import pandas as pd
    import openpyxl

    return (
        f"Python {sys.version.split()[0]}, "
        f"pandas {pd.__version__}, "
        f"openpyxl {openpyxl.__version__}"
    )


def _strip_fences(code: str) -> str:
    if code.startswith("```python"):
        code = code[len("```python") :].strip()
    if code.startswith("```"):
        code = code[3:].strip()
    if code.endswith("```"):
        code = code[:-3].strip()
    return code


# Re-export verifier symbols so tests can patch them at the converter module level
from excel2py.verifier import (  # noqa: E402
    build_correction_prompt,
    compare_outputs,
    extract_ground_truth,
    run_script,
)


def convert(
    input_file: Path,
    output_file: Path | None = None,
    provider: str | None = None,
    model: str | None = None,
    api_key: str | None = None,
    dry_run: bool = False,
    verify: bool = True,
    max_verify_attempts: int = 5,
    verify_timeout: int = 60,
    settings: Settings | None = None,
) -> str:
    """Convert an Excel file to a Python script.

    Returns the generated Python code.
    """
    settings = settings or get_settings()
    provider = provider or settings.default_provider

    # Parse Excel file
    ext = input_file.suffix.lower()
    parser = _get_parser(ext)
    logger.info("Parsing %s with %s", input_file, type(parser).__name__)
    workbook = parser.parse(input_file)

    # Build prompt
    system_prompt = get_system_prompt()
    user_prompt = serialize_workbook(workbook)

    if dry_run:
        logger.info("Dry run — prompt generated but not sent to LLM")
        return user_prompt

    # Call LLM (initial generation)
    api_key = api_key or _get_api_key(settings, provider)
    model = model or _get_model(settings, provider)
    lm = create_chat_model(provider, api_key, model, settings.temperature)

    def _generate(prompt: str) -> str:
        response = lm.invoke([SystemMessage(content=system_prompt), HumanMessage(content=prompt)])
        code = _strip_fences(response.content.strip())
        try:
            ast.parse(code)
        except SyntaxError as e:
            raise CodeGenerationError(f"Generated code has syntax errors: {e}") from e
        return code

    def _generate_at_temp(prompt: str, temperature: float) -> str:
        """Generate initial code at a specific temperature (used for diversity escapes)."""
        alt_lm = create_chat_model(provider, api_key, model, temperature)
        response = alt_lm.invoke([SystemMessage(content=system_prompt), HumanMessage(content=prompt)])
        alt_code = _strip_fences(response.content.strip())
        try:
            ast.parse(alt_code)
        except SyntaxError as e:
            raise CodeGenerationError(f"Alt-temperature generation has syntax errors: {e}") from e
        return alt_code

    def _generate_fix_with_qa(
        correction_prompt: str, temperature: float | None = None
    ) -> str:
        """Parallel correction: Fixer + QA run simultaneously, moderator produces final code.

        Backend is selected via EXCEL2PY_CORRECTION_BACKEND ("langchain" or "agno").
        Both agents receive the same prompt, the moderator reconciles in one shot.
        """
        temp = temperature if temperature is not None else settings.temperature
        backend = settings.correction_backend
        if backend == "agno":
            correction_lm = create_chat_model(provider, api_key, model, temp)
            raw = run_broadcast_correction(correction_prompt, correction_lm)
        else:
            invoke = build_model_invoke(provider, model, api_key, temp)
            raw = run_langchain_correction(correction_prompt, invoke)
        code = _strip_fences(raw.strip())
        try:
            ast.parse(code)
        except SyntaxError as e:
            raise CodeGenerationError(
                f"Correction team ({backend}) produced code with syntax errors: {e}"
            ) from e
        return code

    logger.info("Sending to %s (model: %s)", provider, model)
    code = _generate(user_prompt)

    # Verification-and-correction loop
    if verify:
        ground_truth = extract_ground_truth(input_file)
        best_code = code
        best_score: float = float("inf")  # weighted; lower = better
        best_error_count: float = float("inf")  # raw count, for display only
        best_result = None
        best_exit_code = 0
        best_stderr = ""
        best_output_dir: Path | None = None
        last_result = None
        attempt_history: list[dict] = []
        consecutive_stagnant = 0
        _last_improvement_attempt = 0  # for plateau detection
        _PLATEAU_WINDOW = 2  # exit if best_score unchanged for this many attempts
        # D — convergence detection
        seen_code_hashes: set[str] = {hashlib.md5(code.encode()).hexdigest()}
        seen_error_signatures: set[str] = set()
        # Semantic fingerprints: frozenset of (error_type, expected_value) — count-insensitive.
        # Same root cause (e.g. "Functions Used → nan") shows the same fingerprint even when
        # the error count varies between attempts.
        seen_error_fingerprints: set[frozenset] = set()

        for attempt in range(1, max_verify_attempts + 1):
            with tempfile.NamedTemporaryFile(suffix=".py", mode="w", delete=False) as f:
                f.write(code)
                script_path = Path(f.name)

            try:
                exit_code, stdout, stderr, output_dir = run_script(
                    script_path, input_file, verify_timeout
                )

                if exit_code != 0:
                    from excel2py.verifier import VerificationError, VerificationResult

                    last_result = VerificationResult(
                        passed=False,
                        errors=[
                            VerificationError(
                                sheet="",
                                location="",
                                expected="exit 0",
                                actual=f"exit {exit_code}",
                                error_type="crash",
                            )
                        ],
                    )
                else:
                    last_result = compare_outputs(output_dir, ground_truth)

                error_count = last_result.error_count()
                logger.info(
                    "Verification attempt %d/%d: %s (%d errors)",
                    attempt,
                    max_verify_attempts,
                    "PASS" if last_result.passed else "FAIL",
                    error_count,
                )

                if not last_result.passed:
                    if exit_code != 0:
                        logger.info("  Script exited with code %d", exit_code)
                        if stderr:
                            for line in stderr.strip().splitlines()[-10:]:
                                logger.info("  stderr: %s", line)
                    for err in last_result.errors[:10]:
                        logger.info(
                            "  [%s] sheet=%r location=%r expected=%r actual=%r",
                            err.error_type,
                            err.sheet,
                            err.location,
                            err.expected,
                            err.actual,
                        )
                    if error_count > 10:
                        logger.info("  ... and %d more errors", error_count - 10)
                    logger.debug("  stdout: %s", stdout[:500] if stdout else "(empty)")

                score = _score_result(last_result, exit_code)

                # Rubber duck diagnosis (Self-Debugging, arXiv 2304.05128):
                # Have the LLM explain WHY its code fails before it tries to fix it.
                # Stored in attempt_history as Reflexion-style verbal memory so subsequent
                # correction prompts know what root causes were already identified.
                diagnosis: str | None = None
                if not last_result.passed:
                    try:
                        diag_invoke = build_model_invoke(
                            provider, model, api_key, settings.temperature, lm=lm
                        )
                        diag_diagnostics = (
                            compute_output_diagnostics(output_dir, ground_truth)
                            if output_dir is not None and exit_code == 0
                            else None
                        )
                        import pandas as pd

                        diag_gt_samples = {
                            k: v
                            for k, v in ground_truth.items()
                            if isinstance(v, pd.DataFrame)
                            and any(e.sheet == k for e in last_result.errors)
                        }
                        diagnosis = run_rubber_duck_diagnosis(
                            code,
                            last_result.errors,
                            exit_code,
                            stderr,
                            diag_invoke,
                            output_diagnostics=diag_diagnostics,
                            ground_truth_samples=diag_gt_samples,
                        )
                        logger.debug("Rubber duck diagnosis:\n%s", diagnosis[:500])
                    except Exception as exc:
                        logger.debug("Rubber duck diagnosis failed: %s", exc)

                # Record this attempt in history before updating best
                attempt_history.append(
                    {
                        "n": attempt,
                        "error_count": error_count,
                        "score": score,
                        "exit_code": exit_code,
                        "errors": last_result.errors[:5],
                        "stderr_tail": stderr.strip().splitlines()[-5:]
                        if stderr
                        else [],
                        "diagnosis": diagnosis,
                    }
                )

                # Weighted best-code selection:
                # shape/structural errors score 10 000–1 000 000× more than value errors,
                # so a script with right shape + 60 wrong values beats one with 3 wrong shapes.
                if score < best_score:
                    best_score = score
                    best_error_count = error_count
                    best_code = code
                    best_result = last_result
                    best_exit_code = exit_code
                    best_stderr = stderr
                    best_output_dir = output_dir
                    consecutive_stagnant = 0
                    _last_improvement_attempt = attempt
                else:
                    consecutive_stagnant += 1

                if last_result.passed:
                    break

                # Plateau exit: if best score hasn't improved for _PLATEAU_WINDOW attempts,
                # spending more attempts on the same dead end is wasteful — return best code.
                if (
                    attempt >= 3
                    and (attempt - _last_improvement_attempt) >= _PLATEAU_WINDOW
                    and best_score < float("inf")
                ):
                    logger.info(
                        "Improvement plateau: best_score=%.0f unchanged for %d attempts — exiting with best code",
                        best_score,
                        attempt - _last_improvement_attempt,
                    )
                    break

                if attempt < max_verify_attempts:
                    # C — property-based failures (plain English, no enum categories)
                    prop_failures = (
                        check_properties(best_output_dir, ground_truth)
                        if best_output_dir is not None
                        else None
                    )

                    # B — LLM judge: natural-language description of what is wrong per sheet
                    judge_parts: list[str] = []
                    if best_output_dir is not None:
                        import pandas as pd
                        from excel2py.verifier import _find_matching_file

                        judge_invoke = build_model_invoke(
                            provider, model, api_key, settings.temperature, lm=lm
                        )
                        output_files = list(best_output_dir.glob("*.csv")) + list(
                            best_output_dir.glob("*.xlsx")
                        )
                        for sheet_name, expected_df in ground_truth.items():
                            if not isinstance(expected_df, pd.DataFrame):
                                continue
                            actual_file = _find_matching_file(output_files, sheet_name)
                            if actual_file is None:
                                continue
                            try:
                                actual_df = (
                                    pd.read_csv(actual_file, header=None)
                                    if actual_file.suffix == ".csv"
                                    else pd.read_excel(actual_file, header=None)
                                )
                            except Exception:
                                continue
                            if actual_df.shape != expected_df.shape:
                                verdict = run_output_judge(
                                    sheet_name, expected_df, actual_df, judge_invoke
                                )
                                judge_parts.append(f"**{sheet_name}**: {verdict}")

                    # Legacy diagnostics (concrete row/col diff)
                    diagnostics = (
                        compute_output_diagnostics(best_output_dir, ground_truth)
                        if best_output_dir is not None
                        else None
                    )

                    gt_samples = {
                        k: v
                        for k, v in ground_truth.items()
                        if isinstance(v, __import__("pandas").DataFrame)
                        and any(
                            e.sheet == k
                            for e in (best_result.errors if best_result else [])
                        )
                    }

                    # Collect dtype snapshot for mismatch errors to help diagnose
                    # NaN vs empty string and other type coercion bugs.
                    actual_state_snapshot = None
                    if (
                        best_output_dir is not None
                        and best_result is not None
                        and any(e.error_type == "mismatch" for e in best_result.errors)
                    ):
                        from excel2py.verifier import collect_actual_state_snapshot
                        actual_state_snapshot = collect_actual_state_snapshot(
                            best_output_dir, ground_truth, best_result.errors
                        )

                    # AST-based linter: deterministic anti-pattern detection.
                    # Findings are concrete line-level bugs (not prose rules) injected
                    # at the top of the correction prompt where the LLM sees them first.
                    from excel2py.verifier import lint_generated_code
                    lint_findings = lint_generated_code(best_code) or None

                    correction = build_correction_prompt(
                        best_code,
                        best_result,
                        best_exit_code,
                        best_stderr,
                        workbook_description=user_prompt,
                        attempt_history=attempt_history,
                        environment_info=_get_env_info(),
                        output_diagnostics=diagnostics,
                        ground_truth_samples=gt_samples,
                        stagnant_attempts=consecutive_stagnant,
                        property_failures=prop_failures,
                        judge_feedback="\n\n".join(judge_parts)
                        if judge_parts
                        else None,
                        actual_state_snapshot=actual_state_snapshot,
                        lint_findings=lint_findings,
                    )
                    logger.debug("Correction prompt:\n%s", correction[:1000])

                    # Temperature escalation: ramp 0.2 per stagnant attempt so the LLM
                    # explores different fixes rather than repeating the same wrong one.
                    # Research (LLMLOOP, Reflexion) shows temp=0 causes 48x more repetition
                    # loops than temp=1.0 — escalation is the primary escape mechanism.
                    correction_temp = min(
                        settings.temperature + consecutive_stagnant * 0.2,
                        0.9,
                    )
                    if consecutive_stagnant >= 1:
                        logger.info(
                            "Temperature escalated to %.1f (stagnant %d attempts)",
                            correction_temp,
                            consecutive_stagnant,
                        )
                    try:
                        new_code = _generate_fix_with_qa(
                            correction, temperature=correction_temp
                        )
                    except CodeGenerationError:
                        new_code = best_code

                    # D — convergence detection (two signals):
                    # 1. Identical code bytes — LLM reproduced exact same fix.
                    #    Don't stop immediately: try one high-temperature escape before giving up
                    #    (Hermes/OpenHands pattern: turn-scoped fallback before aborting).
                    new_hash = hashlib.md5(new_code.encode()).hexdigest()
                    if new_hash in seen_code_hashes:
                        escape_temp = min(settings.temperature + 0.6, 0.95)
                        logger.info(
                            "Identical code detected — escape attempt at temperature %.2f",
                            escape_temp,
                        )
                        try:
                            escape_code = _generate_fix_with_qa(
                                correction, temperature=escape_temp
                            )
                            escape_hash = hashlib.md5(escape_code.encode()).hexdigest()
                            if escape_hash not in seen_code_hashes:
                                seen_code_hashes.add(escape_hash)
                                new_code = escape_code
                            else:
                                logger.info(
                                    "Escape attempt also identical — stopping early"
                                )
                                break
                        except CodeGenerationError:
                            logger.info(
                                "Escape attempt invalid — falling back to best code"
                            )
                            new_code = best_code
                    else:
                        seen_code_hashes.add(new_hash)

                    # 2. Semantic loop detection (arXiv 2605.02236, OpenHands pattern):
                    #    Match on (error_type, expected_value) pairs — count-insensitive.
                    #    "Functions Used → nan" is the same root cause at attempt 3 (22 errors)
                    #    and attempt 5 (60 errors), even though the exact signature differs.
                    if exit_code == 0:
                        err_fingerprint = frozenset(
                            (e.error_type, str(e.expected)[:60])
                            for e in last_result.errors
                        )
                        # Also keep exact-signature detection as a secondary check
                        err_sig = "|".join(
                            f"{e.sheet}:{e.expected}:{e.actual}"
                            for e in last_result.errors
                        )
                        semantic_loop = (
                            bool(err_fingerprint)
                            and err_fingerprint in seen_error_fingerprints
                        ) or (bool(err_sig) and err_sig in seen_error_signatures)

                        if semantic_loop:
                            # Approach pivot: route to error-type-specific last resort.
                            # Use high temperature (0.7) per mixed-temperature strategy sampling
                            # (arXiv 2603.02045): strategy tokens at high temp, code at low temp.
                            lr_temp = min(settings.temperature + 0.5, 0.95)
                            primary_types = {
                                e.error_type
                                for e in (best_result.errors if best_result else [])
                            }
                            has_shape = bool(
                                primary_types
                                & {"shape_mismatch", "missing_sheet", "crash"}
                            )
                            if has_shape:
                                logger.info(
                                    "Semantic loop: same shape error pattern — shape-targeted pivot (temp=%.2f)",
                                    lr_temp,
                                )
                                lr_prompt = build_last_resort_prompt(
                                    best_code, ground_truth, _get_env_info()
                                )
                            else:
                                logger.info(
                                    "Semantic loop: same value mismatch pattern — merged-cell pivot (temp=%.2f)",
                                    lr_temp,
                                )
                                lr_prompt = build_value_mismatch_last_resort_prompt(
                                    best_code,
                                    ground_truth,
                                    best_result,
                                    _get_env_info(),
                                )
                            try:
                                lr_code = _generate_fix_with_qa(
                                    lr_prompt, temperature=lr_temp
                                )
                                lr_hash = hashlib.md5(lr_code.encode()).hexdigest()
                                if lr_hash not in seen_code_hashes:
                                    with tempfile.NamedTemporaryFile(
                                        suffix=".py", mode="w", delete=False
                                    ) as lrf:
                                        lrf.write(lr_code)
                                        lr_path = Path(lrf.name)
                                    try:
                                        lr_exit, _, lr_stderr, lr_odir = run_script(
                                            lr_path, input_file, verify_timeout
                                        )
                                        if lr_exit == 0:
                                            lr_result = compare_outputs(
                                                lr_odir, ground_truth
                                            )
                                            lr_score = _score_result(lr_result, lr_exit)
                                            logger.info(
                                                "Pivot attempt: %s (%d errors, score %.0f)",
                                                "PASS" if lr_result.passed else "FAIL",
                                                lr_result.error_count(),
                                                lr_score,
                                            )
                                            if lr_score < best_score:
                                                best_score = lr_score
                                                best_error_count = (
                                                    lr_result.error_count()
                                                )
                                                best_code = lr_code
                                                best_result = lr_result
                                                best_exit_code = lr_exit
                                                best_stderr = lr_stderr
                                                best_output_dir = lr_odir
                                    finally:
                                        lr_path.unlink(missing_ok=True)
                            except CodeGenerationError:
                                pass

                            # After the pivot, try fresh regeneration if attempts remain
                            # (Reflexion/LATS pattern: reset priors, inject failure context).
                            # If candidate A is already seen, try a high-temp candidate B for
                            # diversity (multipath decoding, arXiv:2509.07676).
                            if attempt < max_verify_attempts - 1:
                                remaining = max_verify_attempts - attempt - 1
                                logger.info(
                                    "Semantic loop: pivot done — fresh regeneration "
                                    "(%d attempt(s) left)",
                                    remaining,
                                )
                                try:
                                    fresh_prompt = build_fresh_generation_prompt(
                                        user_prompt,
                                        attempt_history,
                                        _extract_required_shapes(ground_truth),
                                    )
                                    candidate_a = _generate(fresh_prompt)
                                    hash_a = hashlib.md5(candidate_a.encode()).hexdigest()
                                    if hash_a in seen_code_hashes:
                                        # A is stale — try high-temperature B for novelty
                                        try:
                                            candidate_b = _generate_at_temp(
                                                fresh_prompt,
                                                min(settings.temperature + 0.4, 0.85),
                                            )
                                            hash_b = hashlib.md5(candidate_b.encode()).hexdigest()
                                            if hash_b not in seen_code_hashes:
                                                logger.info(
                                                    "Dual candidate: pivoted to novel candidate B (A was already seen)"
                                                )
                                                code = candidate_b
                                                seen_code_hashes.add(hash_b)
                                                continue
                                        except CodeGenerationError:
                                            pass
                                    code = candidate_a
                                    seen_code_hashes.add(hash_a)
                                    continue
                                except CodeGenerationError:
                                    logger.info(
                                        "Fresh regeneration produced invalid code — stopping"
                                    )
                            break

                        if err_fingerprint:
                            seen_error_fingerprints.add(err_fingerprint)
                        if err_sig:
                            seen_error_signatures.add(err_sig)

                    code = new_code
            finally:
                script_path.unlink(missing_ok=True)

        if last_result is not None and not last_result.passed:
            logger.warning(
                "Verification incomplete after %d attempts. %d issues remain.",
                max_verify_attempts,
                best_error_count,
            )
        code = best_code

    # Write output
    if output_file:
        output_file.write_text(code)
        logger.info("Written to %s", output_file)

    return code
