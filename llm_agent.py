"""
llm_agent.py — the ONLY part of the system that talks to an LLM.

Its job is narrow on purpose: turn a natural-language chat instruction into
ONE structured edit op (see edit_engine.py for the schema), given a compact
graph summary — never the raw document text. That's what keeps cost flat:
DocGraph.summary_for_llm() is O(number of editable nodes), not O(document
length), so a 60-page contract with 40 tracked values costs roughly the
same per edit as a 2-page one with 40 tracked values.

Real usage (with an Anthropic API key set as ANTHROPIC_API_KEY):

    from llm_agent import propose_edit
    op = propose_edit(graph, "Bump the purchase price to $150,000")

This repo's demo.py runs with LIVE_LLM=0 by default (offline / no network
in this sandbox) and falls back to a tiny rule-based stand-in so the whole
pipeline is runnable end to end without a key. Swap in the real call for
production — the prompt/schema below is what actually ships.
"""
import json
import os
import re

SYSTEM_PROMPT = """You are a document-editing engine. You are given a compact \
list of the EDITABLE nodes in a Word document (bookmarks) and, separately, \
which fields are DERIVED (REF fields and formulas) and therefore off-limits \
for direct editing — those update automatically after you edit their source.

Given a user instruction, respond with ONE JSON object describing a single \
edit operation, and nothing else — no prose, no markdown fences.

Allowed ops:
  {"op": "set_bookmark_text", "bookmark": "<name>", "text": "<new text>"}
  {"op": "insert_paragraph_after", "bookmark": "<name>", "text": "<new text>"}

Rules:
- Never target a REF or FORMULA node — those are computed, not editable.
- If the instruction is ambiguous or targets something not in the node list,
  respond with {"op": "clarify", "question": "<what you need to know>"}.
- Preserve the existing formatting style of the value you're replacing
  (e.g. currency format) unless told otherwise.
"""


def build_user_prompt(graph_summary: str, instruction: str) -> str:
    return f"{graph_summary}\n\nUSER INSTRUCTION: {instruction}"


def propose_edit(graph, instruction: str) -> dict:
    """Returns a parsed edit-op dict. Calls the real API if ANTHROPIC_API_KEY
    is set and LIVE_LLM=1; otherwise uses a deterministic offline stand-in
    so this repo runs without network/credentials."""
    if os.environ.get("LIVE_LLM") == "1":
        return _call_claude(graph.summary_for_llm(), instruction)
    return _offline_stub(graph, instruction)


def _call_claude(graph_summary: str, instruction: str) -> dict:
    import anthropic  # pip install anthropic
    client = anthropic.Anthropic()
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=300,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": build_user_prompt(graph_summary, instruction)}],
    )
    text = "".join(b.text for b in msg.content if b.type == "text").strip()
    text = re.sub(r"^```json|```$", "", text, flags=re.MULTILINE).strip()
    return json.loads(text)


def _offline_stub(graph, instruction: str) -> dict:
    """Deterministic stand-in used only because this sandbox has no network.
    Mirrors what the real model call would return for the demo's specific
    instruction, so the rest of the pipeline (apply -> propagate -> repack)
    is exercised for real."""
    m = re.search(r"\$?([\d,]+(?:\.\d+)?)", instruction)
    if "purchase price" in instruction.lower() and m:
        amount = float(m.group(1).replace(",", ""))
        for name in graph.bookmarks:
            if "price" in name:
                return {"op": "set_bookmark_text", "bookmark": name,
                        "text": f"${amount:,.2f}"}
    return {"op": "clarify", "question": "Which value should I change?"}
