from __future__ import annotations

import openai
from openai import OpenAI
from tenacity import retry, retry_if_exception, stop_after_attempt, wait_exponential

from excel2py.exceptions import ProviderAuthError, ProviderError
from excel2py.llm.base import BaseLLMProvider, LLMRequest, LLMResponse


def _is_retryable_openai_error(exc: BaseException) -> bool:
    if isinstance(exc, openai.RateLimitError):
        return True
    if isinstance(exc, openai.APIStatusError) and exc.status_code in (500, 502, 503):
        return True
    return False


class OpenAIProvider(BaseLLMProvider):
    @property
    def default_model(self) -> str:
        return "gpt-4o"

    @property
    def provider_name(self) -> str:
        return "openai"

    @retry(
        retry=retry_if_exception(_is_retryable_openai_error),
        wait=wait_exponential(min=1, max=30),
        stop=stop_after_attempt(3),
        reraise=True,
    )
    def generate(self, request: LLMRequest) -> LLMResponse:
        client = OpenAI(api_key=self.api_key)
        try:
            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": request.system_prompt},
                    {"role": "user", "content": request.user_prompt},
                ],
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
