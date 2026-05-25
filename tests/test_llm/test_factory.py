import pytest
from unittest.mock import patch, MagicMock
from langchain_core.language_models.chat_models import BaseChatModel

from excel2py.llm.factory import create_chat_model, _SUPPORTED_PROVIDERS


class TestFactory:
    def test_create_anthropic_returns_chat_model(self):
        with patch("excel2py.llm.factory.ChatAnthropic") as MockCls:
            mock_lm = MagicMock(spec=BaseChatModel)
            MockCls.return_value.with_retry.return_value = mock_lm
            result = create_chat_model("anthropic", "key", "claude-sonnet-4-6")
            MockCls.assert_called_once_with(
                model="claude-sonnet-4-6", api_key="key", temperature=0.2, max_tokens=8192
            )
            assert result is mock_lm

    def test_create_openai_returns_chat_model(self):
        with patch("excel2py.llm.factory.ChatOpenAI") as MockCls:
            mock_lm = MagicMock(spec=BaseChatModel)
            MockCls.return_value.with_retry.return_value = mock_lm
            result = create_chat_model("openai", "key", "gpt-4o")
            MockCls.assert_called_once_with(model="gpt-4o", api_key="key", temperature=0.2)
            assert result is mock_lm

    def test_create_openrouter_uses_base_url(self):
        with patch("excel2py.llm.factory.ChatOpenAI") as MockCls:
            mock_lm = MagicMock(spec=BaseChatModel)
            MockCls.return_value.with_retry.return_value = mock_lm
            create_chat_model("openrouter", "key", "openai/gpt-4o")
            _, kwargs = MockCls.call_args
            assert "openrouter.ai" in kwargs["base_url"]

    def test_create_google_returns_chat_model(self):
        with patch("excel2py.llm.factory.ChatGoogleGenerativeAI") as MockCls:
            mock_lm = MagicMock(spec=BaseChatModel)
            MockCls.return_value.with_retry.return_value = mock_lm
            result = create_chat_model("google", "key", "gemini-2.5-flash")
            MockCls.assert_called_once_with(
                model="gemini-2.5-flash", google_api_key="key", temperature=0.2
            )
            assert result is mock_lm

    def test_agno_maps_to_anthropic_backend(self):
        with patch("excel2py.llm.factory.ChatAnthropic") as MockCls:
            MockCls.return_value.with_retry.return_value = MagicMock()
            create_chat_model("agno", "key", "claude-sonnet-4-6")
            MockCls.assert_called_once()

    def test_unknown_provider_raises(self):
        with pytest.raises(ValueError, match="Unknown provider"):
            create_chat_model("unknown", "key", "model")

    def test_all_supported_providers_listed(self):
        assert "anthropic" in _SUPPORTED_PROVIDERS
        assert "openai" in _SUPPORTED_PROVIDERS
        assert "google" in _SUPPORTED_PROVIDERS
        assert "openrouter" in _SUPPORTED_PROVIDERS

    def test_retry_applied(self):
        with patch("excel2py.llm.factory.ChatAnthropic") as MockCls:
            inner = MagicMock()
            MockCls.return_value = inner
            create_chat_model("anthropic", "key", "claude-sonnet-4-6")
            inner.with_retry.assert_called_once_with(stop_after_attempt=3)
