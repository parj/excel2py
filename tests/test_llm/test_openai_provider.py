from unittest.mock import MagicMock, patch
from excel2py.llm.openai_provider import OpenAIProvider
from excel2py.llm.base import LLMRequest


class TestOpenAIProvider:
    def test_provider_properties(self):
        provider = OpenAIProvider(api_key="test")
        assert provider.provider_name == "openai"
        assert provider.default_model == "gpt-4o"

    @patch("excel2py.llm.openai_provider.OpenAI")
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
        mock_response.model = "gpt-4o"
        mock_response.usage = mock_usage

        mock_client.chat.completions.create.return_value = mock_response

        provider = OpenAIProvider(api_key="test-key")
        request = LLMRequest(system_prompt="system", user_prompt="user")
        result = provider.generate(request)

        assert result.content == "print('hello')"
        assert result.model == "gpt-4o"
        assert result.provider == "openai"
        assert result.usage["input_tokens"] == 10
