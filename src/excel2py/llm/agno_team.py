from __future__ import annotations

import logging

from agno.agent import Agent
from agno.team.mode import TeamMode
from agno.team.team import Team

logger = logging.getLogger(__name__)

_FIXER_INSTRUCTIONS = [
    "You are an expert Python developer fixing a script that produced incorrect outputs.",
    "Analyse the root cause of every reported verification error step by step.",
    "Then propose a complete corrected Python script.",
    "Structure your response EXACTLY as:",
    "## Analysis",
    "<step-by-step reasoning about what caused each error and how the fix addresses it>",
    "## Fixed Code",
    "```python",
    "<corrected Python code>",
    "```",
]

_QA_INSTRUCTIONS = [
    "You are a critical QA engineer reviewing a Python script that produced verification errors.",
    "For every reported error, specify EXACTLY:",
    "1. What the code currently produces (the actual value)",
    "2. What it must produce to pass (the expected value)",
    "3. The precise logic change required to achieve it",
    "Do NOT write or propose code. Focus only on precise, unambiguous requirements.",
    "Be specific about formulas, data transformations, output column order, and shapes.",
]

_MODERATOR_INSTRUCTIONS = [
    "You are a senior software engineer leading a code correction review.",
    "You receive two inputs from your team members:",
    "  - From the Fixer: a root-cause analysis and a proposed Python fix",
    "  - From QA: exact requirements every error correction must satisfy",
    "Your job: produce ONE final corrected Python script that:",
    "  1. Implements the Fixer's approach wherever it is correct",
    "  2. Satisfies EVERY requirement identified by QA — no exceptions",
    "  3. Runs without errors and produces exactly the right outputs",
    "Return ONLY valid Python code. No markdown fences, no explanations, no comments.",
]


_SUPPORTED_PROVIDERS = ("anthropic", "agno", "openai", "openrouter")


def build_agno_model(provider: str, model_id: str, api_key: str, temperature: float):
    """Return the correct Agno model adapter for the given provider."""
    if provider in ("anthropic", "agno"):
        from agno.models.anthropic import Claude

        return Claude(id=model_id, api_key=api_key, temperature=temperature)
    if provider == "openrouter":
        from agno.models.openrouter import OpenRouter

        return OpenRouter(id=model_id, api_key=api_key, temperature=temperature)
    if provider == "openai":
        from agno.models.openai import OpenAIChat

        return OpenAIChat(id=model_id, api_key=api_key, temperature=temperature)
    raise ValueError(
        f"Provider '{provider}' is not supported by the Agno broadcast team. "
        f"Supported: {list(_SUPPORTED_PROVIDERS)}"
    )


def run_broadcast_correction(correction_prompt: str, agno_model) -> str:
    """Fixer + QA run in broadcast mode; team leader synthesises final code in one shot.

    Both agents receive the same correction prompt simultaneously.  The Fixer
    proposes a fix with analysis; QA specifies exact requirements.  The team
    leader (moderator) combines both perspectives and returns the final code.
    """
    fixer = Agent(
        name="Fixer",
        role="Python code fixer",
        model=agno_model,
        instructions=_FIXER_INSTRUCTIONS,
        markdown=False,
    )
    qa = Agent(
        name="QA",
        role="QA requirements analyst",
        model=agno_model,
        instructions=_QA_INSTRUCTIONS,
        markdown=False,
    )

    team = Team(
        name="CodeCorrectionTeam",
        mode=TeamMode.broadcast,
        model=agno_model,
        members=[fixer, qa],
        instructions=_MODERATOR_INSTRUCTIONS,
        markdown=False,
        show_members_responses=False,
    )

    logger.info("Running Agno broadcast correction team (Fixer + QA → moderator)")
    output = team.run(correction_prompt)
    content = (
        output.get_content_as_string()
        if callable(getattr(output, "get_content_as_string", None))
        else str(output.content)
    )
    logger.debug("Broadcast team raw output:\n%s", content[:500])
    return content
