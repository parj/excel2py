# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`excel2py` converts Excel spreadsheets (`.xlsx`, `.xlsm`, `.xlsb`) to runnable Python scripts using an LLM. It extracts formulas, pivot tables, named ranges, and VBA macros, serializes them into a structured prompt, and sends it to a configurable LLM provider (OpenAI, Anthropic, or Google Gemini).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
cp .env.example .env  # then fill in API keys
```

## Commands

```bash
# Run all tests
pytest

# Run a single test file
pytest tests/test_converter.py -v

# Run a single test by name
pytest tests/test_converter.py::TestConverter::test_dry_run -v

# Integration tests (requires real API keys in .env)
pytest -m integration

# Lint
ruff check src/ tests/

# CLI usage
excel2py convert input.xlsx --dry-run          # preview prompt only
excel2py convert input.xlsx --provider anthropic
excel2py convert input.xlsx -o output.py --verbose
```

## Architecture

The conversion pipeline in `converter.py` runs sequentially:
1. **Parse** — `parsers/xlsx_parser.py` (openpyxl) or `parsers/xlsb_parser.py` (python-calamine) produces a `WorkbookData` object. `.xlsm` files also run `parsers/macro_extractor.py` (oletools) for VBA.
2. **Serialize** — `prompts/templates.py:serialize_workbook()` converts `WorkbookData` to a markdown table representation. It deduplicates fill-down formulas (e.g., 100 rows of `=B{n}*C{n}` collapse to one entry) and caps data rows at 50 per sheet.
3. **Generate** — The selected LLM provider (`llm/openai_provider.py`, `llm/anthropic_provider.py`, or `llm/google_provider.py`) is created via `llm/factory.py` and called with the serialized prompt.
4. **Validate** — The response is stripped of markdown fences and validated with `ast.parse()`.
5. **Write** — Result written to the output `.py` file.

### Verification-and-correction loop

After initial generation, the script is run against the original Excel file and its output is compared to ground-truth DataFrames extracted by `verifier.py`. Up to `EXCEL2PY_MAX_VERIFY_ATTEMPTS` (default 5) iterations are attempted. Each failed attempt runs through a four-level escape hierarchy:

| Level | Trigger | Action |
|-------|---------|--------|
| 1 — Temperature escalation | Same weighted error score for N consecutive attempts | Ramp correction temperature by +0.2 per stagnant step (max 0.9) |
| 2 — Identical-code escape | Correction hash matches a previously seen hash | Retry at temperature +0.6; if that produces invalid code, fall back to best known code and continue (does **not** stop early) |
| 3 — Semantic-loop pivot | Same `(error_type, expected_value)` fingerprint recurs | Route to error-type-specific last-resort prompt (`build_last_resort_prompt` for shape errors, `build_value_mismatch_last_resort_prompt` for value errors) |
| 4 — Fresh regeneration | Semantic-loop pivot exhausted with ≥2 attempts remaining | Regenerate from scratch using `build_fresh_generation_prompt`: injects diagnosed root causes as negative examples (Reflexion, arXiv 2303.11366) and hard shape constraints from ground truth |

Rubber duck diagnosis (Self-Debugging, arXiv 2304.05128) runs before every correction: the LLM traces WHY its code fails rather than symptom-patching. The diagnosis is stored in `attempt_history` and fed into each subsequent correction prompt as verbal memory.

### Key modules

| Module | Role |
|--------|------|
| `converter.py` | Orchestrates the full pipeline including the verification loop |
| `verifier.py` | Ground-truth extraction, script execution, output comparison, all correction/pivot prompt builders |
| `models.py` | Dataclasses: `WorkbookData`, `SheetData`, `CellData`, `PivotTableData`, `MacroData` |
| `config.py` | Pydantic Settings — all config via `EXCEL2PY_*` env vars |
| `prompts/templates.py` | Prompt construction + formula deduplication logic |
| `llm/langchain_team.py` | Rubber duck diagnosis, single-agent correction, LLM judge; supports `anthropic`, `openai`, `openrouter` |
| `llm/agno_team.py` | Broadcast correction team (Fixer + QA + Moderator); enabled via `EXCEL2PY_CORRECTION_BACKEND=agno` |
| `llm/base.py` | `BaseLLMProvider` abstract class + `LLMRequest`/`LLMResponse` dataclasses |
| `exceptions.py` | Exception hierarchy rooted at `Excel2PyError` |

### LLM providers

All providers retry with exponential backoff (up to 3 attempts) on rate limit / server errors. Default models: `gpt-4o` (OpenAI), `claude-sonnet-4-6` (Anthropic), `gemini-2.5-flash` (Google).

The `correction_backend` setting controls how corrections are generated:
- `langchain` (default) — single-agent correction via `llm/langchain_team.py`
- `agno` — broadcast team (Fixer + QA + Moderator) via `llm/agno_team.py`

### Known limitations

- **XLSB**: `python-calamine` does not expose pivot table metadata or named ranges.
- **Large sheets**: Data rows are capped at 50 per sheet; formula cells are always included regardless.
- **VBA**: Converted by LLM heuristically, not executed or validated.
