# OpenRouter Provider Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add OpenRouter as a fourth LLM provider using the existing `openai` SDK pointed at OpenRouter's OpenAI-compatible endpoint.

**Architecture:** A new `OpenRouterProvider` class mirrors `OpenAIProvider` exactly, differing only in `base_url`, `provider_name`, and `default_model`. Config, factory, and CLI each gain one line. No new dependencies.

**Tech Stack:** `openai` SDK (already installed), `tenacity` (already installed), `pydantic-settings`, `click`

---

### Task 1: Create OpenRouter provider with tests

**Files:**
- Create: `tests/test_llm/test_openrouter_provider.py`
- Create: `src/excel2py/llm/openrouter_provider.py`

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_llm/test_openrouter_provider.py
from unittest.mock import MagicMock, patch
from excel2py.llm.openrouter_provider import OpenRouterProvider
from excel2py.llm.base import LLMRequest


class TestOpenRouterProvider:
    def test_provider_properties(self):
        provider = OpenRouterProvider(api_key="test")
        assert provider.provider_name == "openrouter"
        assert provider.default_model == "openai/gpt-4o"

    @patch("excel2py.llm.openrouter_provider.OpenAI")
    def test_generate(self, mock_openai_class):
        mock_client = MagicMock()
        mock_openai_class.return_value = mock_client

        mock_usage = MagicMock()
        mock_usage.prompt_tokens = 10
        mock_usage.completion_tokens = 20
        mock_usage.total_tokens = 30

        mock_choice = MagicMock()
        mock_choice.message.content = "print('hello')"
        mock_choice.finish_reason = "stop"

        mock_response = MagicMock()
        mock_response.choices = [mock_choice]
        mock_response.model = "openai/gpt-4o"
        mock_response.usage = mock_usage

        mock_client.chat.completions.create.return_value = mock_response

        provider = OpenRouterProvider(api_key="test-key")
        request = LLMRequest(system_prompt="system", user_prompt="user")
        result = provider.generate(request)

        assert result.content == "print('hello')"
        assert result.model == "openai/gpt-4o"
        assert result.provider == "openrouter"
        assert result.usage["input_tokens"] == 10

    @patch("excel2py.llm.openrouter_provider.OpenAI")
    def test_uses_openrouter_base_url(self, mock_openai_class):
        mock_openai_class.return_value = MagicMock()
        mock_openai_class.return_value.chat.completions.create.return_value = MagicMock(
            choices=[MagicMock(message=MagicMock(content="x"), finish_reason="stop")],
            model="openai/gpt-4o",
            usage=None,
        )
        provider = OpenRouterProvider(api_key="test-key")
        provider.generate(LLMRequest(system_prompt="s", user_prompt="u"))

        mock_openai_class.assert_called_once_with(
            api_key="test-key",
            base_url="https://openrouter.ai/api/v1",
        )
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
pytest tests/test_llm/test_openrouter_provider.py -v
```

Expected: `ModuleNotFoundError` or `ImportError` — `openrouter_provider` does not exist yet.

- [ ] **Step 3: Implement the provider**

```python
# src/excel2py/llm/openrouter_provider.py
from __future__ import annotations

import openai
from openai import OpenAI
from tenacity import retry, retry_if_exception, stop_after_attempt, wait_exponential

from excel2py.exceptions import ProviderAuthError, ProviderError
from excel2py.llm.base import BaseLLMProvider, LLMRequest, LLMResponse

_OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"


def _is_retryable_openrouter_error(exc: BaseException) -> bool:
    if isinstance(exc, openai.RateLimitError):
        return True
    if isinstance(exc, openai.APIStatusError) and exc.status_code in (500, 502, 503):
        return True
    return False


class OpenRouterProvider(BaseLLMProvider):
    @property
    def default_model(self) -> str:
        return "openai/gpt-4o"

    @property
    def provider_name(self) -> str:
        return "openrouter"

    @retry(
        retry=retry_if_exception(_is_retryable_openrouter_error),
        wait=wait_exponential(min=1, max=30),
        stop=stop_after_attempt(3),
        reraise=True,
    )
    def generate(self, request: LLMRequest) -> LLMResponse:
        client = OpenAI(api_key=self.api_key, base_url=_OPENROUTER_BASE_URL)
        try:
            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": request.system_prompt},
                    {"role": "user", "content": request.user_prompt},
                ],
                max_tokens=request.max_tokens,
                temperature=request.temperature,
            )
        except openai.AuthenticationError as exc:
            raise ProviderAuthError(str(exc)) from exc
        except openai.RateLimitError:
            raise
        except openai.APIStatusError as exc:
            if exc.status_code in (500, 502, 503):
                raise
            raise ProviderError(str(exc)) from exc
        except openai.APIError as exc:
            raise ProviderError(str(exc)) from exc

        choice = response.choices[0]
        usage = {}
        if response.usage:
            usage = {
                "input_tokens": response.usage.prompt_tokens,
                "output_tokens": response.usage.completion_tokens,
                "total_tokens": response.usage.total_tokens,
            }

        return LLMResponse(
            content=choice.message.content or "",
            model=response.model,
            provider=self.provider_name,
            usage=usage,
            finish_reason=choice.finish_reason or "",
        )
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
pytest tests/test_llm/test_openrouter_provider.py -v
```

Expected: 3 tests PASSED.

- [ ] **Step 5: Commit**

```bash
git add src/excel2py/llm/openrouter_provider.py tests/test_llm/test_openrouter_provider.py
git commit -m "feat: add OpenRouterProvider"
```

---

### Task 2: Register in factory and update config + CLI

**Files:**
- Modify: `src/excel2py/llm/factory.py`
- Modify: `src/excel2py/config.py`
- Modify: `src/excel2py/cli.py`
- Modify: `tests/test_llm/test_factory.py`
- Modify: `.env.example`

- [ ] **Step 1: Add factory test**

Add to `tests/test_llm/test_factory.py` inside `class TestFactory`:

```python
    def test_create_openrouter(self):
        from excel2py.llm.openrouter_provider import OpenRouterProvider
        provider = create_provider("openrouter", api_key="test-key")
        assert isinstance(provider, OpenRouterProvider)
        assert provider.model == "openai/gpt-4o"
```

- [ ] **Step 2: Run to confirm it fails**

```bash
pytest tests/test_llm/test_factory.py::TestFactory::test_create_openrouter -v
```

Expected: FAIL — `ValueError: Unknown provider: openrouter`

- [ ] **Step 3: Update factory**

In `src/excel2py/llm/factory.py`, update `_get_providers()`:

```python
def _get_providers() -> dict[str, type[BaseLLMProvider]]:
    from excel2py.llm.anthropic_provider import AnthropicProvider
    from excel2py.llm.google_provider import GoogleProvider
    from excel2py.llm.openai_provider import OpenAIProvider
    from excel2py.llm.openrouter_provider import OpenRouterProvider

    return {
        "openai": OpenAIProvider,
        "anthropic": AnthropicProvider,
        "google": GoogleProvider,
        "openrouter": OpenRouterProvider,
    }
```

- [ ] **Step 4: Run factory tests to confirm they pass**

```bash
pytest tests/test_llm/test_factory.py -v
```

Expected: all tests PASSED.

- [ ] **Step 5: Update config**

In `src/excel2py/config.py`, add two fields after `google_model`:

```python
    openrouter_api_key: str | None = None
    openrouter_model: str = "openai/gpt-4o"
```

Full file after edit:

```python
"""Configuration management for excel2py."""

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_prefix="EXCEL2PY_")

    default_provider: str = "openai"
    openai_api_key: str | None = None
    openai_model: str = "gpt-4o"
    anthropic_api_key: str | None = None
    anthropic_model: str = "claude-sonnet-4-20250514"
    google_api_key: str | None = None
    google_model: str = "gemini-2.5-flash"
    openrouter_api_key: str | None = None
    openrouter_model: str = "openai/gpt-4o"
    max_tokens: int = 8192
    temperature: float = 0.2
    max_retries: int = 3
    log_level: str = "INFO"


def get_settings() -> Settings:
    return Settings()
```

- [ ] **Step 6: Update CLI choice list**

In `src/excel2py/cli.py`, change the `--provider` option:

```python
@click.option("-p", "--provider", type=click.Choice(["openai", "anthropic", "google", "openrouter"]), default=None,
              help="LLM provider to use")
```

- [ ] **Step 7: Update .env.example**

```ini
# LLM Provider API Keys
EXCEL2PY_OPENAI_API_KEY=your-openai-api-key
EXCEL2PY_ANTHROPIC_API_KEY=your-anthropic-api-key
EXCEL2PY_GOOGLE_API_KEY=your-google-api-key
EXCEL2PY_OPENROUTER_API_KEY=your-openrouter-api-key

# Default provider: openai, anthropic, google, or openrouter
EXCEL2PY_DEFAULT_PROVIDER=openai

# Model overrides (optional)
# EXCEL2PY_OPENAI_MODEL=gpt-4o
# EXCEL2PY_ANTHROPIC_MODEL=claude-sonnet-4-20250514
# EXCEL2PY_GOOGLE_MODEL=gemini-2.5-flash
# EXCEL2PY_OPENROUTER_MODEL=openai/gpt-4o
```

- [ ] **Step 8: Run full test suite**

```bash
pytest -v
```

Expected: all tests PASSED.

- [ ] **Step 9: Commit**

```bash
git add src/excel2py/llm/factory.py src/excel2py/config.py src/excel2py/cli.py tests/test_llm/test_factory.py .env.example
git commit -m "feat: register openrouter in factory, config, and CLI"
```
