from __future__ import annotations

from google import genai
from google.genai import errors as genai_errors
from tenacity import retry, retry_if_exception, stop_after_attempt, wait_exponential

from excel2py.exceptions import ProviderAuthError, ProviderError
from excel2py.llm.base import BaseLLMProvider, LLMRequest, LLMResponse


def _is_retryable_google_error(exc: BaseException) -> bool:
    if isinstance(exc, genai_errors.APIError):
        # Rate limit errors typically have status 429; server errors are 5xx
        status = getattr(exc, "status_code", None) or getattr(exc, "code", None)
        if status in (429, 500, 502, 503):
            return True
    return False


class GoogleProvider(BaseLLMProvider):
    def __init__(self, api_key: str, model: str | None = None):
        super().__init__(api_key=api_key, model=model)
        self.client = genai.Client(api_key=api_key)

    @property
    def default_model(self) -> str:
        return "gemini-2.5-flash"

    @property
    def provider_name(self) -> str:
        return "google"

    @retry(
        retry=retry_if_exception(_is_retryable_google_error),
        wait=wait_exponential(min=1, max=30),
        stop=stop_after_attempt(3),
        reraise=True,
    )
    def generate(self, request: LLMRequest) -> LLMResponse:
        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=request.user_prompt,
                config=genai.types.GenerateContentConfig(
                    system_instruction=request.system_prompt,
                    temperature=request.temperature,
                ),
            )
        except genai_errors.APIError as exc:
            status = getattr(exc, "status_code", None) or getattr(exc, "code", None)
            if status == 401:
                raise ProviderAuthError(str(exc)) from exc
            if status in (429, 500, 502, 503):
                raise
            raise ProviderError(str(exc)) from exc
        except Exception as exc:
            # Catch-all for unexpected errors from the google-genai SDK
            raise ProviderError(str(exc)) from exc

        usage = {}
        if response.usage_metadata:
            meta = response.usage_metadata
            usage = {
                "input_tokens": getattr(meta, "prompt_token_count", None),
                "output_tokens": getattr(meta, "candidates_token_count", None),
                "total_tokens": getattr(meta, "total_token_count", None),
            }
            # Remove None values
            usage = {k: v for k, v in usage.items() if v is not None}

        finish_reason = ""
        if response.candidates:
            finish_reason = str(response.candidates[0].finish_reason or "")

        return LLMResponse(
            content=response.text,
            model=self.model,
            provider=self.provider_name,
            usage=usage,
            finish_reason=finish_reason,
        )
