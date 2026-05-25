# Verification-and-Correction Loop Design

## Problem

Generated Python scripts have three recurring quality issues:
1. **Unused code** — dead imports and unreachable functions
2. **Merged cells** — script crashes or produces NaN where Excel showed a value
3. **Incomplete pivot tables** — aggregation functions, filter values, and sort order not captured; silent extraction failures

---

## Section 1 — Upstream Fixes

### System prompt (`prompts/templates.py:get_system_prompt`)
- Rule 7: Use xlsb_reader with a **mandatory `_build_sheet_df` helper** (see Output Contract below).
- Rule 8: No unused imports, unused variables, or unreachable functions.
- Rule 9: Write CSVs with `df.to_csv(path, index=False, header=False)`. The verifier reads with `header=None`; a pandas header line becomes an extra data row causing a shape mismatch of exactly +1 row.

### Output Contract (injected verbatim into system prompt)
The LLM is shown the verifier's own ground-truth construction code and CSV reading call rather than prose rules. This is self-documenting — if the verifier changes, the spec shown to the LLM updates automatically:

```python
def _build_sheet_df(values: dict) -> pd.DataFrame:
    """MANDATORY: copy verbatim and use for every sheet.
    Merged cells: xlsb_reader returns only the anchor cell value (no forward-fill).
    Non-anchor positions are absent from the dict and become '' after fillna."""
    if not values:
        return pd.DataFrame()
    max_row = max(r for r, _ in values)
    max_col = max(c for _, c in values)
    data = [[None] * (max_col + 1) for _ in range(max_row + 1)]
    for (r, c), v in values.items():
        if v is not None:  # keep '' — do NOT use `if v:` (drops empty strings)
            data[r][c] = v
    df = pd.DataFrame(data)
    return df.dropna(how='all').dropna(axis=1, how='all').fillna('')

# The verifier reads your CSV with:
actual_df = pd.read_csv(your_csv, header=None)
```

### Pivot extraction fix (`parsers/xlsx_parser.py:_extract_pivot_tables`)
- Extend `PivotTableData` model with `aggregation_functions: dict[str, str]`
- Read `subtotalFunctions` from each `dataField`
- Replace bare `except: pass` with `except Exception` that logs a warning

---

## Section 2 — Verification Module (`verifier.py`)

### Ground truth extraction
- Read Excel with `xlsb_reader` (`XlsbWorkbook` / `XlsxWorkbook`) via `extract_ground_truth()`
- Build `dict[str, pd.DataFrame]` using `_values_to_dataframe()`: same logic as `_build_sheet_df` above
- Skip cells containing Excel errors (`#REF!`, `#VALUE!`, `#N/A`, etc.)
- Skip sheets with no data cells

### Blank / NaN equivalence (`_is_blank`)
`None`, `''`, `float('nan')`, `np.float64(nan)`, `pd.NA`, and `pd.NaT` are treated as equivalent in cell comparison. An `expected=''` vs `actual=nan` is not a mismatch — both represent "blank". This avoids false positives from Excel cells that have an explicit empty-string cached value vs cells where the generated script produces NaN for a genuinely empty position.

### Output comparison (`compare_outputs`, `_compare_dataframes`)
- Read script output files with `pd.read_csv(path, header=None)` or `pd.read_excel(path, header=None)`
- Match output files to sheets by filename similarity (`_find_matching_file`)
- Compare cell-by-cell: numeric values within `rtol=1e-5`, strings exact, blanks via `_is_blank`
- Produce `VerificationResult`: `passed: bool`, `errors: list[VerificationError]`
  - Error types: `"mismatch"`, `"shape_mismatch"`, `"missing_sheet"`, `"no_output"`, `"crash"`

### Pass criteria
- Script exits with code 0
- All shapes match; all cell values within tolerance or both blank
- No unmatched output sheets

### AST-based linter (`lint_generated_code`)
Before every correction attempt, the generated code is parsed with Python's `ast` module. Known anti-patterns produce concrete, line-level findings injected at the top of the correction prompt. Adding a new check costs ~5 lines of AST code:

| Pattern detected | Finding |
|---|---|
| `to_csv()` without `header=False` | +1 row shape mismatch |
| `to_csv()` without `index=False` | extra index column |
| `if v:` inside `.items()` loop | drops `''` (falsy) silently |
| `dropna()` without `how='all'` | drops rows with any NaN, not all-NaN |
| `read_csv/read_excel` without `keep_default_na=False` | empty strings converted to NaN |

### Dtype state snapshot (`collect_actual_state_snapshot`)
When mismatch errors are present, the actual output CSVs are read back and their `.shape`, `.dtypes`, and `.head(3)` are injected into the correction prompt. A float64 column with NaN where object dtype is expected directly shows the LLM the type coercion bug.

### Value mismatch last-resort routing (`_classify_mismatch_errors`)
When the semantic loop triggers for value mismatches, errors are classified:
- **`nan_vs_empty`**: majority of mismatches are `expected=''` vs `actual=nan` → inject `fillna('')` + `if v is not None:` pattern
- **`merged_cell`**: other value mismatches → inject openpyxl merged-cell forward-fill pattern

---

## Section 3 — Correction Loop (`converter.py`)

### Flow
```
generate → run script → verify → pass? → done
                   ↓ fail
          AST lint → build correction prompt → regenerate → run script → verify → ...
```
Max attempts: 5 (configurable via `EXCEL2PY_MAX_VERIFY_ATTEMPTS`). Best-result tracker: keep the attempt with lowest weighted error score.

### Weighted scoring
```
crash=1,000,000 · no_output=100,000 · shape_mismatch=10,000 · mismatch=1
```
A script with correct shape + 60 wrong values beats one with 3 wrong shapes.

### Running the script
- `subprocess.run` with configurable timeout (default 60s)
- Run in a temp directory; pass original Excel path as first argument
- Capture stdout, stderr, exit code

### Correction prompt content
- **Static Analysis Findings** (from AST linter) — first section, deterministic
- Attempt history with rubber duck diagnoses (Reflexion-style verbal memory, arXiv:2303.11366)
- Root cause diagnosis from latest rubber duck session (Self-Debugging, arXiv:2304.05128)
- Property failures (plain English), LLM judge feedback, concrete row/col diff
- Dtype state snapshot for mismatch errors
- Ground truth samples (expected first/last rows with exact required shapes)

### Escape hierarchy (four levels)

| Level | Trigger | Action |
|---|---|---|
| 1 — Temperature escalation | Same weighted score for N consecutive attempts | +0.2 per stagnant step (max 0.9) |
| 2 — Improvement plateau | Best score unchanged for 2 attempts after attempt 3 | Exit early with best code (Debugging Decay Index, arXiv:2506.18403) |
| 3 — Identical-code hash | `md5(new_code)` already seen | Escape at base_temp + 0.6; if still identical, fall back to best code |
| 4 — Semantic loop | Same `frozenset(error_type, expected_value)` fingerprint recurs | Approach pivot (shape: `build_last_resort_prompt`; values: `build_value_mismatch_last_resort_prompt`) + fresh regeneration |

### Fresh regeneration with dual candidates
When semantic loop fires and fresh regeneration runs, if candidate A is already in `seen_code_hashes`, a second candidate B is generated at `base_temp + 0.4` for diversity before falling back. Based on multipath decoding approach (arXiv:2509.07676).

### Rubber duck diagnosis (`run_rubber_duck_diagnosis`)
Before every correction, the LLM traces WHY the code fails — not just what the symptoms are. The diagnosis is stored in `attempt_history` and fed into each subsequent correction prompt as verbal memory, preventing symptom-patching (Self-Debugging, arXiv:2304.05128).

---

## Section 4 — Correction Backends

### `langchain` (default, `llm/langchain_team.py`)
Single-agent correction via `run_langchain_correction`. The `_CORRECTION_SYSTEM` prompt is compact and generic — specific bug patterns are surfaced by the AST linter and output contract, not hardcoded prose rules.

### `agno` (opt-in via `EXCEL2PY_CORRECTION_BACKEND=agno`, `llm/agno_team.py`)
Sequential debate: **Fixer → QA critique of Fixer's proposed code → Moderator arbitrates**.

Unlike parallel broadcast (where QA sees only the error description), QA here receives the Fixer's actual proposed code and can identify specific remaining bugs (e.g., "your `to_csv()` call at line 42 is still missing `header=False`"). Based on SWE-Debate pattern (arXiv:2507.23348): Supporter → Skeptic → Judge.

Research note: DebateCoder (arXiv:2601.21469) and empirical studies on LLM self-correction for data science (arXiv:2408.15658) support 2–3 rounds as optimal; further rounds show diminishing returns or degradation.

---

## Section 5 — CLI & Settings

### CLI flags (`cli.py`)
- `--no-verify` → `verify=False`
- `--max-verify-attempts N` (default 5)

### Settings (`config.py`)
- `EXCEL2PY_VERIFY: bool = True`
- `EXCEL2PY_MAX_VERIFY_ATTEMPTS: int = 5`
- `EXCEL2PY_VERIFY_TIMEOUT: int = 60`
- `EXCEL2PY_CORRECTION_BACKEND: str = "langchain"` (or `"agno"`)

### `convert()` signature
- `verify: bool = True`
- `max_verify_attempts: int = 5`
- `verify_timeout: int = 60`

---

## Section 6 — Error Handling Edge Cases

| Situation | Handling |
|-----------|----------|
| Missing imports in script | Capture `ModuleNotFoundError` from stderr; include in correction prompt |
| Script hangs | Kill after timeout; treat as crash |
| Script writes no output files | Verification failure: "script produced no output" |
| Multiple output files | Match by filename similarity to sheet name |
| Purely formatting sheets | Skipped (no data cells in ground truth) |
| Excel error cells (`#REF!` etc.) | Skipped from ground truth |
| `--dry-run` | Verification never runs |
| Improvement plateau | Early exit after 2 stagnant attempts (post attempt 3) |
| Identical code loop | High-temp escape attempt then best-code fallback |

---

## Section 7 — Testing

### Unit tests
- `tests/test_verifier.py`: ground truth extraction, `_is_blank` equivalence, AST linter patterns, `_classify_mismatch_errors`, `build_value_mismatch_last_resort_prompt` routing, plateau exit, correction prompt construction
- `tests/test_converter.py`: `verify=False` skips loop; best-result tracker; correction loop calls LLM on failure
- `tests/test_prompts/test_templates.py`: system prompt rules, output contract presence

### Fixtures
- Synthetic `.xlsx` built with openpyxl in a fixture (not a binary file): formula column, merged cell range, one pivot table

### Integration tests
- `@pytest.mark.integration`: end-to-end convert with verification on real Excel files

---

## Files

| File | Purpose |
|------|---------|
| `src/excel2py/verifier.py` | Ground truth extraction, output comparison, AST linter, correction prompt builders |
| `src/excel2py/converter.py` | Full pipeline including verification-and-correction loop, escape hierarchy |
| `src/excel2py/prompts/templates.py` | System prompt with output contract + mandatory `_build_sheet_df` helper |
| `src/excel2py/llm/langchain_team.py` | Rubber duck diagnosis, single-agent correction, LLM judge |
| `src/excel2py/llm/agno_team.py` | Sequential debate: Fixer → QA → Moderator (SWE-Debate pattern) |
| `src/excel2py/llm/factory.py` | LangChain model factory (Anthropic, OpenAI, Google, OpenRouter) |
| `src/excel2py/models.py` | `WorkbookData`, `SheetData`, `CellData`, `PivotTableData`, `MacroData` |
| `src/excel2py/config.py` | Pydantic Settings — all config via `EXCEL2PY_*` env vars |
| `tests/test_verifier.py` | Verifier unit tests including linter and comparison tests |
| `tests/test_converter.py` | Converter unit tests |

---

## Research References

| Paper | arXiv | Used for |
|---|---|---|
| Self-Debugging: Teaching LLMs to Self-Debug | 2304.05128 | Rubber duck diagnosis before every correction — LLM traces WHY code fails rather than patching symptoms |
| Reflexion: Language Agents with Verbal Reinforcement Learning | 2303.11366 | Attempt history with diagnosed root causes as verbal memory; fresh regeneration injects anti-patterns as negative examples |
| Feedback-Triggered Self-Correction with Long-Term Multipath Decoding | 2509.07676 | Dual-candidate fresh regeneration — generate B at higher temperature when A is already seen |
| The Debugging Decay Index | 2506.18403 | Improvement plateau early exit — effectiveness drops exponentially per iteration; exit after 2 stagnant attempts rather than burning remaining budget |
| SWE-Debate: Competitive Multi-Agent Debate for Software Issue Resolution | 2507.23348 | Sequential debate backend: Fixer (Supporter) → QA (Skeptic) → Moderator (Judge); QA critiques Fixer's actual proposed code |
| DebateCoder: Adaptive Confidence Gating in Multi-Agent Collaboration | 2601.21469 | 2–3 rounds optimal; early exit on confidence convergence |
| An Empirical Study on Self-correcting LLMs for Data Science Code Generation | 2408.15658 | Informed decision to use structured feedback (dtype snapshots, AST linter) rather than generic self-correction |
| Code Repair with LLMs gives an Exploration-Exploitation Tradeoff | NeurIPS 2024 | Temperature escalation as exploration mechanism; stagnant loops need diversity not repetition |
