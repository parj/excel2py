from __future__ import annotations

from excel2py.llm.base import BaseLLMProvider, LLMRequest, LLMResponse


class AgnoProvider(BaseLLMProvider):
    """Single-agent provider backed by Agno + Claude (Anthropic)."""

    @property
    def default_model(self) -> str:
        return "claude-sonnet-4-5-20250929"

    @property
    def provider_name(self) -> str:
        return "agno"

    def generate(self, request: LLMRequest) -> LLMResponse:
        from agno.agent import Agent
        from agno.models.anthropic import Claude

        agent = Agent(
            model=Claude(
                id=self.model,
                api_key=self.api_key,
                temperature=request.temperature,
            ),
            instructions=[request.system_prompt],
            markdown=False,
        )
        output = agent.run(request.user_prompt)
        content = (
            output.get_content_as_string()
            if callable(getattr(output, "get_content_as_string", None))
            else str(output.content)
        )
        return LLMResponse(
            content=content,
            model=self.model,
            provider=self.provider_name,
        )
