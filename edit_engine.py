"""
edit_engine.py — applies ONE structured edit operation to the graph.

This is the piece that keeps token cost flat regardless of document size:
the LLM never sees or rewrites the whole document, it only ever emits one
of these small ops. Applying an op touches exactly the runs for that node
— nothing else in the 60-page document is re-serialized or re-sent.

Supported ops (deliberately small — this is a prototype, not the full
system; the brief explicitly says "get part way and show us how you
think"):

  {"op": "set_bookmark_text", "bookmark": "clause_4_2_price", "text": "$150,000.00"}
      Rewrites the text inside a named bookmark. This is a value DEFINITION
      changing (e.g. the actual contract price).

  {"op": "insert_paragraph_after", "bookmark": "clause_4_2_price", "text": "..."}
      Inserts a new paragraph after the paragraph containing a bookmark.
      (Included to show the "insert" case the brief calls out — numbering
      of subsequent clauses is out of scope for this prototype and is
      flagged, not silently guessed at.)

Fields (REF / formula) are deliberately NOT directly editable by op —
they are derived values. Attempting to target one is a validation error,
by design: the LLM should never be allowed to hand-edit a computed
number, only the source it's computed from. That constraint is what
prevents "invented values instead of computed ones."
"""
from lxml import etree
from graph_builder import W, DocGraph


class EditError(Exception):
    pass


def apply_edit(g: DocGraph, op: dict) -> str:
    """Applies op in place on g.tree. Returns a human-readable log line.
    Raises EditError on anything invalid — callers should surface this to
    the LLM as a retry signal, not silently ignore it."""
    kind = op.get('op')

    if kind == 'set_bookmark_text':
        return _set_bookmark_text(g, op['bookmark'], op['text'])

    if kind == 'insert_paragraph_after':
        return _insert_paragraph_after(g, op['bookmark'], op['text'])

    if op.get('bookmark') in (r.target_bookmark for r in g.refs) or \
       kind in ('set_ref_text', 'set_formula_text'):
        raise EditError(
            f"Refused: '{kind}' targets a derived field (REF or formula). "
            f"Derived values are computed by propagation, never edited "
            f"directly — edit the source bookmark instead."
        )

    raise EditError(f"Unknown or unsupported op: {kind}")


def _set_bookmark_text(g: DocGraph, bookmark_name: str, new_text: str) -> str:
    bm = g.bookmarks.get(bookmark_name)
    if bm is None:
        raise EditError(f"No such bookmark: {bookmark_name}")
    if not bm.text_runs:
        raise EditError(f"Bookmark {bookmark_name} has no runs to replace")

    old_text = ''.join(r.findtext(f'{{{W}}}t') or '' for r in bm.text_runs)

    # Preserve the formatting of the FIRST run; collapse all runs into one
    # so we don't leave orphan empty runs behind.
    first = bm.text_runs[0]
    t_el = first.find(f'{{{W}}}t')
    if t_el is None:
        t_el = etree.SubElement(first, f'{{{W}}}t')
    t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t_el.text = new_text

    for extra in bm.text_runs[1:]:
        extra.getparent().remove(extra)
    bm.text_runs = [first]

    return f"set_bookmark_text({bookmark_name}): '{old_text}' -> '{new_text}'"


def _insert_paragraph_after(g: DocGraph, bookmark_name: str, text: str) -> str:
    bm = g.bookmarks.get(bookmark_name)
    if bm is None:
        raise EditError(f"No such bookmark: {bookmark_name}")
    anchor = bm.text_runs[0]
    para = anchor
    while para is not None and etree.QName(para.tag).localname != 'p':
        para = para.getparent()
    if para is None:
        raise EditError("Could not locate containing paragraph")

    new_p = etree.SubElement(para.getparent(), f'{{{W}}}p')
    para.addnext(new_p)
    run = etree.SubElement(new_p, f'{{{W}}}r')
    t = etree.SubElement(run, f'{{{W}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

    return (f"insert_paragraph_after({bookmark_name}): inserted new paragraph. "
            f"NOTE: any manual clause numbering after this point is now "
            f"potentially stale — flagged for review, not silently renumbered "
            f"(numbering fields, if used instead of manual numbers, would "
            f"auto-correct on propagate).")
