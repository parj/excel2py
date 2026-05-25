from __future__ import annotations

import anthropic
from anthropic import Anthropic
from tenacity import retry, retry_if_exception, stop_after_attempt, wait_exponential

from excel2py.exceptions import ProviderAuthError, ProviderError
from excel2py.llm.base import BaseLLMProvider, LLMRequest, LLMResponse


def _is_retryable_anthropic_error(exc: BaseException) -> bool:
    if isinstance(exc, anthropic.RateLimitError):
        return True
    if isinstance(exc, anthropic.InternalServerError):
        return True
    return False


class AnthropicProvider(BaseLLMProvider):
    @property
    def default_model(self) -> str:
        return "claude-sonnet-4-20250514"

    @property
    def provider_name(self) -> str:
        return "anthropic"

    @retry(
        retry=retry_if_exception(_is_retryable_anthropic_error),
        wait=wait_exponential(min=1, max=30),
        stop=stop_after_attempt(3),
        reraise=True,
    )
    def generate(self, request: LLMRequest) -> LLMResponse:
        client = Anthropic(api_key=self.api_key)
        try:
            response = client.messages.create(
                model=self.model,
                system=request.system_prompt,
                messages=[{"role": "user", "content": request.user_prompt}],
                max_tokens=8192,
                temperature=request.temperature,
            )
        except anthropic.AuthenticationError as exc:
            raise ProviderAuthError(str(exc)) from exc
        except anthropic.RateLimitError:
            raise
        except anthropic.InternalServerError:
            raise
        except anthropic.APIError as exc:
            raise ProviderError(str(exc)) from exc

        usage = {}
        if response.usage:
            usage = {
                "input_tokens": response.usage.input_tokens,
                "output_tokens": response.usage.output_tokens,
            }

        return LLMResponse(
            content=response.content[0].text,
            model=response.model,
            provider=self.provider_name,
            usage=usage,
            finish_reason=response.stop_reason or "",
        )
