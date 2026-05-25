from __future__ import annotations

import logging

from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate

logger = logging.getLogger(__name__)

_FIXER_INSTRUCTIONS = """\
You are an expert Python developer fixing a script that produced incorrect outputs.
Analyse the root cause of every reported verification error step by step.
Then propose a complete corrected Python script.
Structure your response EXACTLY as:
## Analysis
<step-by-step reasoning about what caused each error and how the fix addresses it>
## Fixed Code
```python
<corrected Python code>
```"""

_QA_INSTRUCTIONS = """\
You are a critical QA skeptic reviewing a proposed code fix.
You receive: (a) the original verification errors and (b) the Fixer's proposed code.
Your task: identify any remaining bugs in the PROPOSED CODE that will still cause verification failures.
Focus on:
1. Shape errors — does the proposed code produce the exact required number of rows and columns?
   - Is to_csv() called with BOTH index=False AND header=False? A missing header=False adds an extra row.
   - Is dropna(how='all') applied before writing? Missing it leaves empty rows.
   - Is dropna(axis=1, how='all') applied? Missing it leaves empty columns.
2. Type errors — NaN vs empty string (''), float vs int, etc.
3. Off-by-one errors in row/column slicing or grid construction.
4. Any remaining logical errors the Fixer missed.

Be specific: quote the line number or code fragment in the proposed fix that is wrong and
explain exactly what it will produce vs. what is required.
If the fix is correct, say "FIX APPROVED" — do not invent problems.
Do NOT rewrite code. Focus only on critique."""

_MODERATOR_INSTRUCTIONS = """\
You are a senior software engineer making the final call after a code review debate.
You receive:
  - The original verification errors
  - The Fixer's proposed code (with analysis)
  - The QA critic's review of that proposed code

Your job: produce ONE final corrected Python script that:
  1. Starts from the Fixer's proposed code as the base
  2. Addresses EVERY valid concern raised by QA (ignore concerns marked FIX APPROVED)
  3. Runs without errors and produces exactly the right outputs

Return ONLY valid Python code. No markdown fences, no explanations, no comments."""


def run_broadcast_correction(correction_prompt: str, lm) -> str:
    """Sequential debate correction: Fixer proposes → QA critiques the proposal → Moderator arbitrates.

    Unlike parallel broadcast (where QA sees only the error description), the QA agent here
    sees the Fixer's actual proposed code and can point to specific remaining bugs.
    Based on SWE-Debate pattern (arXiv:2507.23348): Supporter → Skeptic → Judge.
    """
    parser = StrOutputParser()

    fixer_chain = (
        ChatPromptTemplate.from_messages([("system", _FIXER_INSTRUCTIONS), ("human", "{input}")])
        | lm
        | parser
    )
    qa_chain = (
        ChatPromptTemplate.from_messages([("system", _QA_INSTRUCTIONS), ("human", "{input}")])
        | lm
        | parser
    )
    moderator_chain = (
        ChatPromptTemplate.from_messages([
            ("system", _MODERATOR_INSTRUCTIONS),
            ("human", "{input}"),
        ])
        | lm
        | parser
    )

    logger.info("Running sequential debate correction (Fixer → QA critique → Moderator)")

    # Round 1: Fixer proposes a fix independently
    fixer_output = fixer_chain.invoke({"input": correction_prompt})

    # Round 2: QA critiques the Fixer's specific proposed code (not just the error description)
    qa_input = (
        f"## Original Verification Errors\n{correction_prompt}\n\n"
        f"## Fixer's Proposed Fix\n{fixer_output}\n\n"
        "Critique the proposed fix above. Is it correct? What bugs remain?"
    )
    qa_output = qa_chain.invoke({"input": qa_input})

    # Round 3: Moderator arbitrates and produces final code
    mod_input = (
        f"## Original Verification Errors\n{correction_prompt}\n\n"
        f"## Fixer's Proposed Fix\n{fixer_output}\n\n"
        f"## QA Critique\n{qa_output}"
    )
    return moderator_chain.invoke({"input": mod_input})
