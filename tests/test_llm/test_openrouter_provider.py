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
