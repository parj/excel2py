from __future__ import annotations

from excel2py.llm.base import BaseLLMProvider


def create_provider(provider_name: str, api_key: str, model: str | None = None) -> BaseLLMProvider:
    providers = _get_providers()
    if provider_name not in providers:
        raise ValueError(f"Unknown provider: {provider_name}. Choose from: {list(providers)}")
    return providers[provider_name](api_key=api_key, model=model)


def _get_providers() -> dict[str, type[BaseLLMProvider]]:
    from excel2py.llm.agno_provider import AgnoProvider
    from excel2py.llm.anthropic_provider import AnthropicProvider
    from excel2py.llm.google_provider import GoogleProvider
    from excel2py.llm.openai_provider import OpenAIProvider
    from excel2py.llm.openrouter_provider import OpenRouterProvider

    return {
        "openai": OpenAIProvider,
        "anthropic": AnthropicProvider,
        "google": GoogleProvider,
        "openrouter": OpenRouterProvider,
        "agno": AgnoProvider,
    }
