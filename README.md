# excel2py

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg) 

![Claude](https://img.shields.io/badge/Claude-D97757?style=for-the-badge&logo=claude&logoColor=white)

> [!WARNING]  
> This has been coded using Claude (sonnet 4.5)

# xlsb_reader

Convert Excel spreadsheets to Python scripts using GenAI.

Supports `.xlsx`, `.xlsm`, and `.xlsb` formats. Extracts formulas, pivot tables, and VBA macros, then sends the structured content to an LLM (OpenAI, Claude, or Gemini) to generate equivalent, runnable Python code.

## Features

- **Full format support**: `.xlsx` and `.xlsm` via openpyxl; `.xlsb` via python-calamine (Rust-based, with formula extraction)
- **Formula conversion**: Extracts Excel formula strings and converts them to Python functions
- **Pivot table conversion**: Reproduces pivot table logic using `pandas.pivot_table()`
- **VBA macro conversion**: Extracts VBA code via oletools and converts it to Python functions
- **Three LLM providers**: OpenAI, Anthropic (Claude), and Google (Gemini)
- **CLI interface**: Simple command-line tool with dry-run support

## Installation

Requires Python 3.11+.

```bash
# Clone the repo
git clone <repo-url>
cd convert-excel-v5

# Create a virtual environment
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install the package
pip install hatchling
pip install -e ".[dev]"
```

## Configuration

API keys can be set via environment variables or a `.env` file in the project root.

Copy the example:
```bash
cp .env.example .env
```

Then edit `.env`:
```env
EXCEL2PY_OPENAI_API_KEY=sk-...
EXCEL2PY_ANTHROPIC_API_KEY=sk-ant-...
EXCEL2PY_GOOGLE_API_KEY=AIza...

# Default provider (openai, anthropic, or google)
EXCEL2PY_DEFAULT_PROVIDER=openai
```

All settings use the `EXCEL2PY_` prefix:

| Variable | Default | Description |
|----------|---------|-------------|
| `EXCEL2PY_DEFAULT_PROVIDER` | `openai` | LLM provider to use |
| `EXCEL2PY_OPENAI_API_KEY` | — | OpenAI API key |
| `EXCEL2PY_OPENAI_MODEL` | `gpt-4o` | OpenAI model |
| `EXCEL2PY_ANTHROPIC_API_KEY` | — | Anthropic API key |
| `EXCEL2PY_ANTHROPIC_MODEL` | `claude-sonnet-4-20250514` | Anthropic model |
| `EXCEL2PY_GOOGLE_API_KEY` | — | Google API key |
| `EXCEL2PY_GOOGLE_MODEL` | `gemini-2.5-flash` | Gemini model |
| `EXCEL2PY_MAX_TOKENS` | `8192` | Max tokens for generated output |
| `EXCEL2PY_TEMPERATURE` | `0.2` | LLM temperature (low = more deterministic) |

## Usage

### Basic conversion

```bash
excel2py convert input.xlsx
```

Output is written to `input_converted.py` by default.

### Specify output file

```bash
excel2py convert input.xlsx -o my_script.py
```

### Choose a provider

```bash
excel2py convert input.xlsx --provider anthropic
excel2py convert input.xlsm --provider google
excel2py convert input.xlsb --provider openai
```

### Override the model

```bash
excel2py convert input.xlsx --provider openai --model gpt-4-turbo
excel2py convert input.xlsx --provider anthropic --model claude-opus-4-5
```

### Pass an API key directly

```bash
excel2py convert input.xlsx --provider openai --api-key sk-...
```

### Dry run (preview the prompt without calling the LLM)

Useful for inspecting what gets sent to the model, or for debugging large files.

```bash
excel2py convert input.xlsx --dry-run
```

### Verbose output

```bash
excel2py convert input.xlsx --verbose
```

## How It Works

1. **Parse**: The Excel file is parsed based on its format:
   - `.xlsx` / `.xlsm`: openpyxl reads the file twice — once for formula strings, once for cached values
   - `.xlsb`: python-calamine (Rust) reads cell values and formula strings
   - VBA macros are extracted from `.xlsm` / `.xlsb` files using oletools

2. **Serialize**: The parsed workbook is serialized into a structured prompt:
   - Sheet dimensions and cell data (capped at 50 data rows, but all formula cells are always included)
   - Named ranges
   - Pivot table definitions (source range, row/column/data/filter fields)
   - Full VBA macro source code

3. **Generate**: The prompt is sent to the chosen LLM with instructions to produce a self-contained, runnable Python script using pandas.

4. **Validate**: The output is checked for valid Python syntax via `ast.parse()` before being written to disk.

## Known Limitations

- **xlsb pivot tables and named ranges**: python-calamine does not expose pivot table metadata or named ranges for `.xlsb` files. Only cell values and formulas are available.
- **Large spreadsheets**: Data rows are capped at 50 per sheet to keep prompts within context window limits. All formula cells are always included regardless of position.
- **VBA execution**: Macros are converted to Python by the LLM — they are not executed or validated beyond syntax checking.

## Development

### Run tests

```bash
pytest
```

### Run linter

```bash
ruff check src/ tests/
```

### Integration tests (requires API keys)

```bash
pytest -m integration
```
