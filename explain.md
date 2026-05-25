# Agentic Code-Repair Loop: B + C + D Pattern

This document explains the three-layer feedback strategy used to make an LLM-based code-repair loop converge reliably. It is self-contained enough to recreate the pattern from scratch in a new session.

---

## The Problem

A code-generation agent produces a Python script, runs it, and gets back error information. It then tries to fix the script — but it keeps failing. Two specific failure modes appear:

1. **Stuck in a loop** — attempts 4 and 5 produce identical errors because the agent generates the same wrong code each time.
2. **Abstract feedback** — the error messages say `shape_mismatch: expected (28, 50) actual (32, 52)` but the agent cannot reason backwards from a shape number to the specific line of code that produced the wrong shape.

The standard approaches (retry with same prompt, multi-agent chains) make this worse because:
- Each agent handoff loses the reasoning trace built by the previous agent
- Abstract numeric errors give the agent no causal signal
- Hard-coded error categories (`shape_mismatch`, `no_output`) are brittle and don't generalise

---

## The Solution: B + C + D

Three independent techniques that stack on top of each other.

---

### D — Convergence Detection (cheapest, runs first)

**What:** Hash every version of generated code (MD5 or SHA). Before running a new fix, check if that hash has been seen before. If yes, exit the loop immediately.

**Why it works:** When an agent is stuck, it doesn't generate slightly different wrong code — it generates *the same* wrong code. The hash check catches this in O(1) after generating the fix, before wasting another execution slot.

**Research basis:** OpenClaw (2024) formalises "boredom detection" — tracking repetitive tool-use patterns as a termination signal. ARCS (arXiv:2504.20434) proves that monotonic improvement requires correcting from the best-so-far code, not the most-recent code; convergence detection is the natural companion.

**Implementation sketch:**

```python
import hashlib

seen_hashes: set[str] = {hashlib.md5(initial_code.encode()).hexdigest()}

for attempt in range(1, max_attempts + 1):
    run_and_verify(code)
    if passed:
        break

    new_code = fix(code)
    new_hash = hashlib.md5(new_code.encode()).hexdigest()

    if new_hash in seen_hashes:
        logger.info("Convergence: identical code generated — stopping")
        break

    seen_hashes.add(new_hash)
    code = new_code
```

**Key rule:** Always correct from `best_code` (fewest errors so far), not the most-recently-generated code. If attempt 3 crashes but attempt 2 had 2 errors, fix attempt 2's code — not attempt 3's.

---

### C — Property-Based Failures (no hard-coded error categories)

**What:** Replace typed error enums (`shape_mismatch`, `no_output`, `mismatch`) with plain-English sentences derived at runtime from the ground truth.

**Why it works:** Hard-coded categories require maintenance and don't generalise. Plain-English sentences derived from the actual expected output adapt automatically to any new failure mode. The LLM also responds better to prose than to table rows with opaque type fields.

**Research basis:** arXiv:2506.18315 found LLMs respond better to "minimal, property-focused feedback" than verbose structured diffs. Property-based testing (arXiv:2510.09907) verifies invariants rather than exact equality — row count, column count, non-null checks are invariants that can be checked without knowing exact cell values.

**Implementation sketch:**

```python
def check_properties(output_dir, ground_truth) -> list[str]:
    failures = []
    for sheet_name, expected_df in ground_truth.items():
        actual_df = read_actual_output(output_dir, sheet_name)
        if actual_df is None:
            failures.append(f"Sheet '{sheet_name}': no output file was written.")
            continue

        row_diff = actual_df.shape[0] - expected_df.shape[0]
        col_diff = actual_df.shape[1] - expected_df.shape[1]

        if row_diff > 0:
            failures.append(
                f"Sheet '{sheet_name}': {row_diff} too many rows "
                f"(wrote {actual_df.shape[0]}, need {expected_df.shape[0]})."
            )
        if col_diff > 0:
            failures.append(
                f"Sheet '{sheet_name}': {col_diff} too many columns "
                f"(wrote {actual_df.shape[1]}, need {expected_df.shape[1]})."
            )
        if row_diff == 0 and col_diff == 0:
            mismatches = count_cell_mismatches(expected_df, actual_df)
            if mismatches:
                failures.append(
                    f"Sheet '{sheet_name}': shape correct but {mismatches} cell(s) differ."
                )
    return failures
```

These sentences go directly into the correction prompt under a `## What is wrong` heading — no enum, no table.

---

### B — LLM-as-Judge (most expensive, provides the "why")

**What:** For each failing sheet, run a second LLM call (the "judge") that receives the actual vs expected data and is asked to describe in 2–4 sentences: what is in the extra/missing rows, which specific operation is causing it, and what the code must change.

**Why it works:** Shape numbers alone (`(32,52) vs (28,50)`) give the fixer no causal signal. The judge converts the raw data diff into a diagnosis: *"The first 4 rows of your output are header rows ('Team', 'P', 'W'…). The script reads them as data because it uses `header=None` without skipping them. Add `skiprows=4` or use `header=0` and drop non-data rows."* This is the signal the fixer needs to change the right line of code.

**Research basis:** arXiv:2507.16587 (LLM-as-judge for code) found judges are unreliable for numerical correctness but reliable for structural differences (missing columns, extra rows). SWE-agent (arXiv:2405.15793) shows that execution traces + structured diagnosis outperform raw stderr. The key: pair the judge with actual execution output, not abstract error messages.

**Judge system prompt:**

```
You are a data analyst comparing actual DataFrame output against expected output.
In 2-4 concrete sentences:
1. Describe what is in the extra or missing rows/columns
   (e.g. "the first 4 rows appear to be column headers: Team, P, W, D, L…")
2. Name the specific code operation producing them
   (e.g. "the script reads with header=None without skipping header rows")
3. State exactly what must change
   (e.g. "add skiprows=4, or use header=0 and drop the first N rows")
Be specific and concrete. Do not write code.
```

**Implementation sketch:**

```python
def run_output_judge(sheet_name, expected_df, actual_df, llm) -> str:
    prompt = (
        f"Sheet: {sheet_name!r}\n"
        f"Expected shape: {expected_df.shape} | Actual shape: {actual_df.shape}\n\n"
        f"Expected first 5 rows:\n{expected_df.head(5).to_string(index=False, header=False)}\n\n"
        f"Actual first 5 rows:\n{actual_df.head(5).to_string(index=False, header=False)}\n"
    )
    if actual_df.shape[0] > expected_df.shape[0]:
        prompt += (
            f"\nActual last 5 rows (extra rows):\n"
            f"{actual_df.tail(5).to_string(index=False, header=False)}\n"
            f"Expected last 5 rows:\n"
            f"{expected_df.tail(5).to_string(index=False, header=False)}\n"
        )
    return llm.invoke([SystemMessage(JUDGE_SYSTEM), HumanMessage(prompt)]).content
```

Only run the judge for sheets where shape does not match — not for crashes or value mismatches (where it adds little signal).

---

## How They Combine in the Loop

```
for attempt in 1..max_attempts:

    run script → get exit_code, stderr, output_dir

    if exit_code != 0:
        record crash in correction prompt
    else:
        C: check_properties(output_dir, ground_truth) → plain-English failures

    update best_code if this attempt has fewer errors

    if passed: break
    if attempt == max_attempts: break

    B: for each failing sheet → run_output_judge → natural-language diagnosis
    build correction prompt with: env info, attempt history, C failures, B diagnosis, raw diff

    new_code = fixer_llm(correction_prompt)

    D: if hash(new_code) in seen_hashes → break (convergence)
    seen_hashes.add(hash(new_code))
    code = new_code
```

**Order matters:**
- D is checked *after* generating the fix, *before* running it — saves one execution slot
- C runs cheaply on every iteration
- B runs only when there are shape/structural failures — it's the most expensive call

---

## Correction Prompt Structure

The correction prompt that goes to the fixer agent contains these sections in order:

1. **Stagnation warning** (if no improvement for 2+ attempts) — forces the agent to try a different strategy class
2. **Runtime environment** — exact Python/pandas/openpyxl versions to prevent hallucinated deprecated APIs
3. **Attempt history** — compact summary of every attempt so far (error count, crash Y/N, first 3 errors)
4. **Original Excel structure** — the serialised workbook description from the initial parse
5. **What is wrong** (C) — plain-English property failures
6. **Why it is wrong** (B) — judge's natural-language diagnosis per sheet
7. **Concrete row/column diff** — actual vs expected first/last 5 rows
8. **Expected output samples** — first 5 and last 3 rows of ground truth for each failing sheet
9. **Best code so far** — the code with the fewest errors, not the most-recently-generated code

---

## Key Rules (easy to get wrong)

| Rule | Why |
|------|-----|
| Always fix `best_code`, not latest code | If attempt 3 crashes after attempt 2 had 2 errors, fixing attempt 3 makes things worse |
| Weight errors by type, not raw count | 60 value mismatches is *closer to correct* than 3 shape mismatches — raw count gets the ordering backwards |
| Reset `consecutive_stagnant` counter when **weighted** score improves | Stagnation warning must only trigger when genuinely stuck |
| D tracks both code hashes AND error signatures | LLM with temperature>0 generates different code that produces the same wrong output — hash alone won't catch it |
| D checks hash *after* generating, *before* running | Saves one expensive script execution |
| A crashed script must never beat a working script for `best_code` | A crash produces no output files; the judge and property checks are blind |
| B runs only for structural failures (shape) | For crashes, stderr is already the right signal; judge adds noise |
| Judge shows actual data, not just shapes | The fixer needs to see what is *in* the extra rows, not just that there are extra rows |
| Single fixer agent, not a chain | Multi-agent chains lose reasoning at each handoff; single agent with rich context wins |

## Error-Type Weights

Use weighted scoring for best-code selection, not raw error count:

```python
_ERROR_WEIGHTS = {
    "no_output":     100_000,   # wrote nothing — blind
    "missing_sheet":  50_000,   # sheet absent — structural
    "shape_mismatch": 10_000,   # shape wrong — structural
    "crash":       1_000_000,   # didn't run — worst
    "mismatch":            1,   # value wrong — easy fix
}

def score_result(result, exit_code):
    if exit_code != 0:
        return 1_000_000 + len(result.errors)
    return sum(_ERROR_WEIGHTS.get(e.error_type, 1) for e in result.errors)
```

**Why this matters:** Without weighting, a script with 3 shape mismatches (score=3) beats a script with 60 value mismatches (score=60), so `best_code` stays as the structurally-wrong code forever. With weighting: 3 shape mismatches = 30,000, 60 value mismatches = 60 — so the structurally-correct-but-imprecise script becomes the new best, and subsequent corrections fix the right thing (merged cell fill, value precision) rather than re-fixing the structure.

---

## What Not to Do

- **Do not use typed error categories** (`shape_mismatch`, `no_output`) in the feedback to the fixer — they are too abstract and require maintenance
- **Do not correct from the most-recently-generated code** — always use `best_code`
- **Do not run the judge for every sheet every time** — only sheets where the shape is wrong
- **Do not use multi-agent chains** (fixer → QA → moderator) for this task — research shows single-agent with rich context outperforms chains because no reasoning is lost at handoffs
- **Do not rely on max_attempts alone to terminate** — add D (hash convergence) to exit early when the agent is cycling
