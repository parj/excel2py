from unittest.mock import MagicMock, patch
from excel2py.llm.anthropic_provider import AnthropicProvider
from excel2py.llm.base import LLMRequest


class TestAnthropicProvider:
    def test_provider_properties(self):
        provider = AnthropicProvider(api_key="test")
        assert provider.provider_name == "anthropic"
        assert provider.default_model == "claude-sonnet-4-20250514"

    @patch("excel2py.llm.anthropic_provider.Anthropic")
    def test_generate(self, mock_anthropic_class):
        mock_client = MagicMock()
        mock_anthropic_class.return_value = mock_client

        mock_usage = MagicMock()
        mock_usage.input_tokens = 15
        mock_usage.output_tokens = 25

        mock_content = MagicMock()
        mock_content.text = "import pandas as pd"

        mock_response = MagicMock()
        mock_response.content = [mock_content]
        mock_response.model = "claude-sonnet-4-20250514"
        mock_response.usage = mock_usage
        mock_response.stop_reason = "end_turn"

        mock_client.messages.create.return_value = mock_response

        provider = AnthropicProvider(api_key="test-key")
        request = LLMRequest(system_prompt="system", user_prompt="user")
        result = provider.generate(request)

        assert result.content == "import pandas as pd"
        assert result.provider == "anthropic"
        assert result.usage["input_tokens"] == 15
