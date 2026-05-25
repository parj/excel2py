from __future__ import annotations

from langchain_core.language_models.chat_models import BaseChatModel

_SUPPORTED_PROVIDERS = ("anthropic", "agno", "openai", "openrouter", "google")


def create_chat_model(
    provider: str,
    api_key: str,
    model: str,
    temperature: float = 0.2,
    max_tokens: int = 8192,
) -> BaseChatModel:
    """Return a LangChain BaseChatModel for the given provider, with exponential backoff.

    SDK-level retries are disabled (max_retries=0) to prevent rapid-fire 429 hammering.
    LangChain's with_retry() takes over with proper exponential backoff:
      wait_exponential_jitter=True  →  min=1s, max=60s, doubles each attempt + jitter
      stop_after_attempt=7          →  waits 1+2+4+8+16+32 = 63s before the final attempt,
                                       covering Anthropic's typical 60s rate-limit window.
    LangChain logs each retry at WARNING level via tenacity before_sleep_log.
    """
    if provider in ("anthropic", "agno"):
        from langchain_anthropic import ChatAnthropic
        lm: BaseChatModel = ChatAnthropic(
            model=model, api_key=api_key, temperature=temperature, max_tokens=max_tokens,
            max_retries=0,  # disable SDK retries; with_retry() owns the backoff
        )
    elif provider == "openai":
        from langchain_openai import ChatOpenAI
        lm = ChatOpenAI(
            model=model, api_key=api_key, temperature=temperature,
            max_retries=0,
        )
    elif provider == "openrouter":
        from langchain_openai import ChatOpenAI
        lm = ChatOpenAI(
            model=model, api_key=api_key, temperature=temperature,
            base_url="https://openrouter.ai/api/v1",
            max_retries=0,
        )
    elif provider == "google":
        from langchain_google_genai import ChatGoogleGenerativeAI
        lm = ChatGoogleGenerativeAI(
            model=model, google_api_key=api_key, temperature=temperature
        )
    else:
        raise ValueError(
            f"Unknown provider: {provider!r}. Choose from: {list(_SUPPORTED_PROVIDERS)}"
        )
    return lm.with_retry(stop_after_attempt=7, wait_exponential_jitter=True)
