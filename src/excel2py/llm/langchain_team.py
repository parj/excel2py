from __future__ import annotations

import logging
from typing import Callable

logger = logging.getLogger(__name__)

# Single-agent system prompt grounded in concrete failure patterns.
_CORRECTION_SYSTEM = """\
You are an expert Python developer fixing a script that produces incorrect output files.

Workflow:
1. Read the "Root Cause Diagnosis" section first — it explains WHY the current code fails,
   not just what the symptoms are. Fix the identified root cause, not just the symptoms.
2. Read the "Concrete output diff" — it shows exactly what the script DID output versus
   what it MUST output. Use this to confirm the diagnosis and verify your fix will work.
3. Check the "Attempt History" — do NOT repeat any approach that already failed.
4. Use only the APIs listed in "Runtime Environment". Notable removals in modern pandas:
   - `DataFrame.applymap` → use `DataFrame.map`
   - `DataFrame.append` → use `pd.concat`
   - `reset_index(axis=1)` is invalid — use `rename_axis(None, axis=1)` to clear column names
5. Make MINIMUM targeted changes. Do not rewrite working parts of the script.
6. Read Excel with xlsb_reader: `from xlsb_reader import XlsxWorkbook, XlsbWorkbook, col_to_letter`.
   iter_values() returns {(row, col): value} sparse dict, 0-based. Convert to DataFrame:
   build a dense array then df.dropna(how='all').dropna(axis=1, how='all').
7. Common causes of extra rows: not calling dropna(how='all') on output DataFrames.
8. Common causes of extra columns: writing index (use index=False), not calling dropna(axis=1, how='all').

Return ONLY valid Python code — no markdown fences, no explanations."""

_JUDGE_SYSTEM = """\
You are a data analyst comparing actual DataFrame output against expected output.
In 2-4 concrete sentences:
1. Describe what is in the extra or missing rows/columns (e.g. "the first 4 rows appear to be column headers: Team, P, W, D, L…")
2. Name the specific code operation that is producing them (e.g. "the script is not skipping the header rows when reading with header=None")
3. State exactly what must change (e.g. "add skiprows=4 when reading sheet X, or read with header=0 and drop the first N rows")
Be specific and concrete. Do not write code."""

_RUBBER_DUCK_SYSTEM = """\
You are a senior code reviewer doing a rubber duck debugging session.

Your ONLY task: read the Python code and the verification errors, then explain step by step
WHY the code produces those errors — trace the actual runtime behaviour of the relevant
functions, do not just restate what the code looks like.

Structure your response as:
1. Trace: walk through the specific function(s) involved in the failing cases, explaining
   what each operation actually returns at runtime.
2. ROOT CAUSE: one sentence identifying the specific operation that must change.
3. FIX DIRECTION: one sentence describing what a correct implementation would do differently
   (do not write code — describe the approach).

Be concrete about runtime behaviour. If you are uncertain about an API's behaviour, say so."""

# Providers that have a direct SDK invoke path (no LangChain required).
_DIRECT_SDK_PROVIDERS = ("anthropic", "agno")
# All providers supported by this module.
_SUPPORTED_PROVIDERS = ("anthropic", "agno", "openai", "openrouter")

InvokeFn = Callable[[str, str], str]  # (system_prompt, user_prompt) -> response_text


def build_model_invoke(
    provider: str,
    model_id: str,
    api_key: str,
    temperature: float,
) -> InvokeFn:
    """Return a (system, user) -> text callable for the given provider.

    Uses the Anthropic SDK directly for anthropic/agno — no LangChain dependency.
    Falls back to LangChain for openai/openrouter.
    """
    if provider in _DIRECT_SDK_PROVIDERS:
        import anthropic as _anthropic

        client = _anthropic.Anthropic(api_key=api_key)

        def _anthropic_invoke(system: str, user: str) -> str:
            response = client.messages.create(
                model=model_id,
                system=system,
                messages=[{"role": "user", "content": user}],
                max_tokens=8192,
                temperature=temperature,
            )
            return response.content[0].text

        return _anthropic_invoke

    if provider == "openrouter":
        from langchain_core.messages import HumanMessage, SystemMessage
        from langchain_openai import ChatOpenAI

        lc = ChatOpenAI(
            model=model_id,
            api_key=api_key,
            temperature=temperature,
            base_url="https://openrouter.ai/api/v1",
        )

        def _openrouter_invoke(system: str, user: str) -> str:
            return lc.invoke(
                [SystemMessage(content=system), HumanMessage(content=user)]
            ).content

        return _openrouter_invoke

    if provider == "openai":
        from langchain_core.messages import HumanMessage, SystemMessage
        from langchain_openai import ChatOpenAI

        lc = ChatOpenAI(model=model_id, api_key=api_key, temperature=temperature)

        def _openai_invoke(system: str, user: str) -> str:
            return lc.invoke(
                [SystemMessage(content=system), HumanMessage(content=user)]
            ).content

        return _openai_invoke

    raise ValueError(
        f"Provider '{provider}' is not supported. Supported: {list(_SUPPORTED_PROVIDERS)}"
    )


def build_langchain_model(
    provider: str, model_id: str, api_key: str, temperature: float
):
    """Backward-compat shim — returns build_model_invoke() result.

    New code should call build_model_invoke() directly.
    """
    return build_model_invoke(provider, model_id, api_key, temperature)


def run_output_judge(
    sheet_name: str,
    expected_df,
    actual_df,
    invoke: InvokeFn,
) -> str:
    """LLM judge: natural-language description of why actual != expected for one sheet."""

    row_diff = actual_df.shape[0] - expected_df.shape[0]
    col_diff = actual_df.shape[1] - expected_df.shape[1]

    prompt_parts = [
        f"Sheet: {sheet_name!r}",
        f"Expected shape: {expected_df.shape} | Actual shape: {actual_df.shape}",
        "",
        "Expected — first 5 rows:",
        expected_df.head(5).to_string(index=False, header=False),
        "",
        "Actual — first 5 rows:",
        actual_df.head(5).to_string(index=False, header=False),
    ]
    if row_diff > 0:
        prompt_parts += [
            f"\nActual has {row_diff} extra row(s). Actual last 5 rows:",
            actual_df.tail(5).to_string(index=False, header=False),
            "Expected last 5 rows:",
            expected_df.tail(5).to_string(index=False, header=False),
        ]
    if col_diff > 0:
        prompt_parts.append(
            f"\nActual has {col_diff} extra column(s). "
            f"Extra column values in row 0: {list(actual_df.iloc[0, expected_df.shape[1] :])}"
        )

    logger.debug("Running output judge for sheet %r", sheet_name)
    return invoke(_JUDGE_SYSTEM, "\n".join(prompt_parts))


def run_rubber_duck_diagnosis(
    code: str,
    errors: list,
    exit_code: int,
    stderr: str,
    invoke: InvokeFn,
    output_diagnostics: str | None = None,
    ground_truth_samples: dict | None = None,
) -> str:
    """Self-Debugging (arXiv 2304.05128): LLM explains WHY its code fails before rewriting.

    Forcing a causal trace before the fix step breaks the symptom-patching loop: the LLM
    must reason through runtime behaviour rather than pattern-matching on error messages.
    The diagnosis is stored in attempt_history and fed back to subsequent correction prompts
    as Reflexion-style verbal memory (arXiv 2303.11366).
    """
    import pandas as pd

    error_lines = []
    if exit_code != 0 and stderr:
        error_lines.append(f"Script crashed (exit {exit_code}):")
        error_lines.extend(f"  {line}" for line in stderr.strip().splitlines()[-10:])
    else:
        for e in errors[:15]:
            error_lines.append(
                f"  [{e.error_type}] sheet={e.sheet!r} {e.location}: "
                f"expected={e.expected!r} actual={e.actual!r}"
            )
        if len(errors) > 15:
            error_lines.append(
                f"  ... and {len(errors) - 15} more errors of the same pattern"
            )

    parts = [
        f"## Verification errors ({len(errors)} total)",
        "\n".join(error_lines),
    ]

    if output_diagnostics:
        parts.append("\n## Concrete output diff (actual rows vs expected rows)")
        parts.append(output_diagnostics)

    if ground_truth_samples:
        parts.append("\n## Expected output (first 5 rows per failing sheet)")
        parts.append(
            "NOTE: ground truth uses dropna(how='all').dropna(axis=1, how='all')"
        )
        parts.append(
            "If your actual output has extra empty rows/columns the expected does not, that is the bug."
        )
        for sheet_name, df in ground_truth_samples.items():
            if not isinstance(df, pd.DataFrame):
                continue
            parts.append(f"\n### Sheet '{sheet_name}' expected (shape {df.shape}):")
            parts.append(df.head(5).to_string(index=False, header=False))

    parts.append(f"\n## Code\n```python\n{code[:4000]}\n```")
    parts.append(
        "\nTrace the runtime behaviour of the functions responsible for these errors. "
        "Pay attention to whether the actual output has EXTRA ROWS or COLUMNS compared "
        "to expected — this indicates missing dropna. End with ROOT CAUSE and FIX DIRECTION."
    )

    logger.info("Running rubber duck diagnosis")
    return invoke(_RUBBER_DUCK_SYSTEM, "\n".join(parts))


def run_langchain_correction(correction_prompt: str, invoke: InvokeFn) -> str:
    """Single-agent correction call."""
    logger.info("Running single-agent correction")
    return invoke(_CORRECTION_SYSTEM, correction_prompt)
