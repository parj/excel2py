# OpenRouter Provider Design

**Date:** 2026-03-26

## Overview

Add OpenRouter as an LLM provider. OpenRouter is an API aggregator with an OpenAI-compatible chat completions endpoint, so it can be integrated by reusing the existing `openai` SDK with a custom `base_url`.

## Architecture

No new dependencies. The `openai` SDK (already in `pyproject.toml`) handles all HTTP communication when pointed at `https://openrouter.ai/api/v1`.

## Components

### `src/excel2py/llm/openrouter_provider.py` (new file)

- Subclasses `BaseLLMProvider`
- Uses `OpenAI(api_key=self.api_key, base_url="https://openrouter.ai/api/v1")`
- `default_model`: `"openai/gpt-4o"` (OpenRouter model IDs use `provider/model` format)
- `provider_name`: `"openrouter"`
- Error handling: identical to `OpenAIProvider` (same `openai.*` exception types apply)
- Retry: 3 attempts, exponential backoff via `tenacity` (same as other providers)

### `src/excel2py/config.py`

Add two fields:
- `openrouter_api_key: str | None = None` — set via `EXCEL2PY_OPENROUTER_API_KEY`
- `openrouter_model: str = "openai/gpt-4o"` — set via `EXCEL2PY_OPENROUTER_MODEL`

### `src/excel2py/llm/factory.py`

Register `"openrouter": OpenRouterProvider` in `_get_providers()`.

### `src/excel2py/cli.py`

Add `"openrouter"` to the `click.Choice` list on the `--provider` option.

## Data Flow

No changes to the conversion pipeline. `converter.py` already resolves the API key via `_get_api_key(settings, provider)` using `getattr(settings, f"{provider}_api_key")`, and model via `_get_model`. Both work automatically once the config fields are added.

## Error Handling

OpenRouter returns standard OpenAI-compatible error responses, so the existing `openai.AuthenticationError`, `openai.RateLimitError`, and `openai.APIStatusError` handling in `OpenAIProvider` applies unchanged to `OpenRouterProvider`.

## Testing

Mirror the existing provider test patterns: mock the `OpenAI` client, assert correct `base_url` is set, verify `LLMResponse` fields are populated correctly.
