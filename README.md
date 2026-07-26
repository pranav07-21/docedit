# docedit — chat-driven Word editor prototype
## Overview

This project demonstrates a graph-based AI DOCX editor that performs chat-driven edits while preserving document consistency across references, bookmarks, and derived values.

Solves a specific failure mode: editing a value that's referenced in more
than one place in a long document (e.g. a price restated on "page 22" and
"page 56") without re-sending or re-generating the whole file, and without
the two copies drifting apart.

## Quickstart

```bash
git clone <your-repo-url>
cd docedit
pip install -r requirements.txt

python3 create_sample_contract.py   # generates the test fixture
python3 demo.py                     # runs the full pipeline (offline stub)
```

That's it — no API key needed to see the pipeline work end to end.

### Run it against the real Claude API

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
python3 demo.py --live
```

### CLI options

```bash
python3 demo.py --instruction "Change the purchase price to $200,000"
python3 demo.py --live
python3 demo.py --input mine.docx --output edited.docx
```

## The core idea

Don't treat the `.docx` as text. Treat it as a **dependency graph**:

- **Bookmarks** = source values (e.g. the Purchase Price defined in §4.2)
- **REF fields** = *derived* copies of a bookmark elsewhere in the doc
  (the same price quoted in the Recitals and the Signature page)
- **Formula fields** = *derived* numbers computed from table cells
  (line totals, `SUM(ABOVE)`)

The LLM's only job is turning a chat instruction into **one small
structured edit op** referencing a node id — never rewriting prose, never
touching fields it shouldn't. A deterministic (non-AI) propagation pass
then walks the graph from the changed node and fixes every REF and
formula that depends on it. Consistency is a correctness property, so
it's code, not a language-generation task.

```
chat instruction
      │
      ▼
appy.py (Tkinter gui/chat interface)
      │
      ▼
graph_builder.py   → parses word/document.xml into DocGraph
      │                (bookmarks, REF fields, formula fields)
      ▼
llm_agent.py       → LLM sees only a compact node summary
      │                (not the 60-page doc), returns ONE edit op
      ▼
edit_engine.py      → applies that op to exactly those runs
      │                (refuses if the op targets a derived field)
      ▼
propagate.py         → deterministic: syncs every REF, recomputes
      │                every formula, marks fields dirty as a fallback
      ▼
repack → validate → result.docx
```


## Why this hits the cost bar

A 60-page contract with 40 tracked values costs about the same per edit
as a 2-page one with 40 values, because the LLM call only ever sees
`DocGraph.summary_for_llm()` — one line per bookmark/field, never the
raw document text. Token cost scales with **number of tracked values**,
not document length.

## Why this hits the consistency bar

REF/formula fields are structurally forbidden as edit targets —
`edit_engine.py` raises `EditError` if you try. The model can't
independently edit page 56's copy and get it wrong; it can only edit the
one source bookmark, and propagation updates every derived copy by
walking real dependency edges — not by re-reading the document and
hoping the model remembers to touch both instances.

## What's demonstrated end to end (`demo.py`)

1. Parses `sample_contract.docx` — a price defined once in §4.2, quoted
   again via REF field in the Recitals and Signature Summary, plus a
   pricing table with formula fields.
2. Chat instruction: *"Update the purchase price to $150,000"*.
3. LLM proposes `{"op": "set_bookmark_text", "bookmark":
   "clause_4_2_price", "text": "$150,000.00"}`.
4. Engine applies it to §4.2 only.
5. Propagation updates both REF copies and recomputes every formula
   field, marking each `w:dirty="true"` as a fallback so Word itself
   recomputes correctly even in cases the small evaluator here doesn't
   cover.
6. Result is re-parsed from disk to prove the change is really there.

## Optional: validate against the OOXML schema

If you have LibreOffice + the Anthropic docx skill's validator available:

```bash
python3 validate.py result_contract.docx --original sample_contract.docx
```

Confirms no corruption and an unchanged paragraph count. (Not required to
run the core pipeline — this is an extra integrity check.)

## Honest limitations

This is a prototype meant to show the mechanism, not a finished product:

- **Formula evaluator is intentionally tiny**: numeric literals, one
  product form (`Bn*const`), and `SUM(ABOVE)`. Not a spreadsheet engine.
  The `w:dirty` fallback exists precisely because of this — a formula
  syntax it can't evaluate still gets corrected the next time the file
  is opened in Word, never silently left wrong.
- **Clause renumbering on insert is flagged, not solved.** True
  auto-renumbering needs Word's `<w:numPr>` list-numbering fields to be
  used for clause numbers (not manual "4.2" text) — at which point the
  same `w:dirty` mechanism handles it for free.
- **Bookmarks assumed single-paragraph.** Multi-paragraph bookmark spans
  need the run-collection logic to walk across paragraph boundaries.
- **No table row insert/delete handling** for the formula recompute —
  only value changes to existing rows.

## Files

| File | Role |
|---|---|
| `create_sample_contract.py` | Generates the test fixture (bookmarks, REF fields, formula fields) |
| `graph_builder.py` | Parses OOXML → `DocGraph` (nodes + edges) |
| `edit_engine.py` | Applies one structured op to one node; refuses derived-field edits |
| `propagate.py` | Deterministic consistency pass (REF sync + formula recompute + dirty-flag fallback) |
| `llm_agent.py` | Prompt/schema for instruction → edit-op; real + offline-stub call paths |
| `demo.py` | Runs the full pipeline, with CLI flags for custom instructions / live API calls |

## License

MIT — see [LICENSE](LICENSE).
