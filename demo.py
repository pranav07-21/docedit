"""
demo.py — runs the full pipeline once, end to end, and prints a before/after
diff so you can see the "page 22 vs page 56" problem actually get solved.

Pipeline:
  1. build_graph()      parse sample_contract.docx into a DocGraph
  2. propose_edit()     LLM (or offline stub) turns chat text into one EditOp
  3. apply_edit()        engine mutates ONLY that node's runs
  4. propagate()          deterministic pass fixes every dependent REF/formula
  5. repack()             write result_contract.docx

Run:
  python3 demo.py                                  # offline stub, default instruction
  python3 demo.py --instruction "..."               # custom instruction
  python3 demo.py --live                            # real Claude API call (needs ANTHROPIC_API_KEY)
  python3 demo.py --input mine.docx --output out.docx
"""
import argparse
import os
from graph_builder import build_graph, repack, W
from edit_engine import apply_edit, EditError
from llm_agent import propose_edit
from propagate import propagate


def dump_values(g):
    bm_vals = {name: ''.join(r.findtext(f'{{{W}}}t') or '' for r in bm.text_runs)
               for name, bm in g.bookmarks.items()}
    ref_vals = {r.id: ''.join(r.display_run.itertext()) if r.display_run is not None else None
                for r in g.refs}
    formula_vals = {f.id: ''.join(f.display_run.itertext()) if f.display_run is not None else None
                    for f in g.formulas}
    return bm_vals, ref_vals, formula_vals


def main():
    parser = argparse.ArgumentParser(description="Chat-driven docx edit demo")
    parser.add_argument('--instruction', default="Update the purchase price to $150,000",
                         help='Chat instruction to send through the pipeline')
    parser.add_argument('--live', action='store_true',
                         help='Call the real Claude API instead of the offline stub '
                              '(requires ANTHROPIC_API_KEY to be set)')
    parser.add_argument('--input', default='sample_contract.docx',
                         help='Path to the .docx to edit')
    parser.add_argument('--output', default='result_contract.docx',
                         help='Path to write the edited .docx to')
    args = parser.parse_args()

    if args.live:
        if not os.environ.get('ANTHROPIC_API_KEY'):
            print("ERROR: --live requires ANTHROPIC_API_KEY to be set in your environment.")
            return
        os.environ['LIVE_LLM'] = '1'

    instruction = args.instruction
    print(f"CHAT INSTRUCTION: {instruction!r}")
    print(f"MODE: {'LIVE (calling claude-sonnet-4-6)' if args.live else 'offline stub'}\n")

    g = build_graph(args.input, workdir='unpacked')
    before_bm, before_ref, before_formula = dump_values(g)

    print("BEFORE:")
    print(" bookmark (Section 4.2 source value):", before_bm)
    print(" REF fields (page 1 / page 56 copies):", before_ref)
    print(" formula fields (pricing table):        ", before_formula)
    print()

    op = propose_edit(g, instruction)
    print(f"LLM PROPOSED OP: {op}\n")

    if op.get('op') == 'clarify':
        print("Needs clarification:", op['question'])
        return

    try:
        applied_log = apply_edit(g, op)
    except EditError as e:
        print("REJECTED:", e)
        return
    print("APPLIED:", applied_log, "\n")

    changed = {op['bookmark']} if 'bookmark' in op else set()
    prop_log = propagate(g, changed)
    print("PROPAGATED (deterministic, no LLM involved):")
    for line in prop_log:
        print(" -", line)
    print()

    g.tree.write(g.doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    repack('unpacked', args.output)
    print(f"Wrote {args.output}")

    g2 = build_graph(args.output, workdir='unpacked_check')
    after_bm, after_ref, after_formula = dump_values(g2)
    print("\nAFTER (re-parsed from the saved file, proving it's really on disk):")
    print(" bookmark:", after_bm)
    print(" REF fields:", after_ref)
    print(" formula fields:", after_formula)

    all_price_refs_match = all(v == after_bm['clause_4_2_price'] for v in after_ref.values())
    print(f"\nCONSISTENCY CHECK — every REF matches the source bookmark: {all_price_refs_match}")


if __name__ == '__main__':
    main()
