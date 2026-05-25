import pytest
from excel2py.llm.factory import create_provider
from excel2py.llm.openai_provider import OpenAIProvider
from excel2py.llm.anthropic_provider import AnthropicProvider
from excel2py.llm.google_provider import GoogleProvider
from excel2py.llm.openrouter_provider import OpenRouterProvider


class TestFactory:
    def test_create_openai(self):
        provider = create_provider("openai", api_key="test-key")
        assert isinstance(provider, OpenAIProvider)
        assert provider.api_key == "test-key"
        assert provider.model == "gpt-4o"

    def test_create_anthropic(self):
        provider = create_provider("anthropic", api_key="test-key")
        assert isinstance(provider, AnthropicProvider)

    def test_create_google(self):
        provider = create_provider("google", api_key="test-key")
        assert isinstance(provider, GoogleProvider)

    def test_create_with_model_override(self):
        provider = create_provider("openai", api_key="test-key", model="gpt-4")
        assert provider.model == "gpt-4"

    def test_create_openrouter(self):
        provider = create_provider("openrouter", api_key="test-key")
        assert isinstance(provider, OpenRouterProvider)
        assert provider.api_key == "test-key"
        assert provider.model == "openai/gpt-4o"

    def test_unknown_provider(self):
        with pytest.raises(ValueError, match="Unknown provider"):
            create_provider("unknown", api_key="test-key")
