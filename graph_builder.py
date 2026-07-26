"""
graph_builder.py — turns word/document.xml into a dependency graph instead
of a blob of text.

This is the "structural graph of the document" the brief hints at. Three
node types cover the failure modes named in the brief:

  BookmarkNode    - a named span of text that is the SOURCE of a value
                    (e.g. the Purchase Price defined in 4.2).
  RefFieldNode     - a REF/PAGEREF/etc field whose cached text must match
                    whatever its target bookmark currently says. This is
                    exactly "page 22 vs page 56."
  FormulaFieldNode - a table formula field (=B2*C2, =SUM(ABOVE)) whose
                    cached text must match a recomputation over its
                    operand cells. This is "invented values instead of
                    computed ones."

Edges point from a node to the node(s) it depends on. Propagation (see
propagate.py) walks edges backwards from whatever changed.

We work on the raw XML tree (lxml), not python-docx's object model, because
fields/bookmarks are exactly the kind of low-level OOXML structure
python-docx doesn't model richly — and because edits must be surgical
(edit one node's runs, touch nothing else) to keep formatting and
everything else in the 60-page document untouched.
"""
import re
import zipfile
import shutil
import os
from dataclasses import dataclass, field
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}
W = NS['w']


def unpack(docx_path, workdir):
    if os.path.exists(workdir):
        shutil.rmtree(workdir)
    os.makedirs(workdir)
    with zipfile.ZipFile(docx_path) as z:
        z.extractall(workdir)
    return workdir


def repack(workdir, out_path):
    if os.path.exists(out_path):
        os.remove(out_path)
    # zip -X (no extra attrs) keeps this deterministic; walk so paths are relative
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(workdir):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, workdir)
                z.write(full, rel)


@dataclass
class BookmarkNode:
    id: str            # bookmark name, e.g. "clause_4_2_price"
    text_runs: list     # list of <w:r> elements between start/end
    context: str        # short human-readable surrounding text, for LLM prompts


@dataclass
class RefFieldNode:
    id: str             # synthetic id, e.g. "ref_0"
    target_bookmark: str
    display_run: object  # the <w:r> holding the cached/displayed text
    context: str


@dataclass
class FormulaFieldNode:
    id: str             # synthetic id, e.g. "formula_0"
    formula: str        # raw instruction, e.g. "B2*80000" or "SUM(ABOVE)"
    display_run: object
    context: str
    table_index: int
    row_index: int
    col_index: int


@dataclass
class DocGraph:
    bookmarks: dict = field(default_factory=dict)
    refs: list = field(default_factory=list)
    formulas: list = field(default_factory=list)
    tree: object = None
    doc_path: str = None

    def summary_for_llm(self):
        """A compact text view of the graph — this is what gets sent to the
        model instead of the whole document. O(nodes near the edit), not
        O(document length)."""
        lines = ["DOCUMENT NODES (editable):"]
        for name, bm in self.bookmarks.items():
            text = ''.join(r.findtext(f'{{{W}}}t') or '' for r in bm.text_runs)
            lines.append(f'  [bookmark:{name}] "{text}"   context: ...{bm.context}...')
        for r in self.refs:
            lines.append(f'  [{r.id}] REF -> {r.target_bookmark} (auto-syncs, not directly editable)')
        for f in self.formulas:
            lines.append(f'  [{f.id}] FORMULA {f.formula} (auto-recomputes, not directly editable)')
        return '\n'.join(lines)


def _local(tag):
    return etree.QName(tag).localname


def build_graph(docx_path, workdir='unpacked') -> DocGraph:
    unpack(docx_path, workdir)
    doc_path = os.path.join(workdir, 'word', 'document.xml')
    tree = etree.parse(doc_path)
    root = tree.getroot()

    g = DocGraph(tree=tree, doc_path=doc_path)

    body = root.find(f'{{{W}}}body')
    # Walk every paragraph/table-cell for bookmarks and fields.
    ref_counter = 0
    formula_counter = 0

    # --- bookmarks: pair bookmarkStart/bookmarkEnd, collect runs between ---
    all_elems = list(root.iter())
    starts = {}
    for el in all_elems:
        if _local(el.tag) == 'bookmarkStart':
            bm_id = el.get(f'{{{W}}}id')
            name = el.get(f'{{{W}}}name')
            starts[bm_id] = (name, el)

    for el in all_elems:
        if _local(el.tag) == 'bookmarkEnd':
            bm_id = el.get(f'{{{W}}}id')
            if bm_id in starts:
                name, start_el = starts[bm_id]
                runs = _runs_between(start_el, el)
                context = _nearby_text(start_el)
                g.bookmarks[name] = BookmarkNode(id=name, text_runs=runs, context=context)

    # --- fields: walk paragraphs, find fldChar begin/separate/end triples ---
    for p in root.iter(f'{{{W}}}p'):
        _scan_fields_in_paragraph(p, g)

    # tag formula fields with table position (best-effort, for context only)
    for t_idx, tbl in enumerate(root.iter(f'{{{W}}}tbl')):
        for r_idx, tr in enumerate(tbl.findall(f'{{{W}}}tr')):
            for c_idx, tc in enumerate(tr.findall(f'{{{W}}}tc')):
                for fnode in g.formulas:
                    if fnode.table_index is None and _contains(tc, fnode.display_run):
                        fnode.table_index, fnode.row_index, fnode.col_index = t_idx, r_idx, c_idx

    return g


def _contains(ancestor, descendant):
    return descendant is not None and ancestor in list(descendant.iterancestors())


def _runs_between(start_el, end_el):
    """Collect <w:r> siblings that occur after start_el and before end_el
    within the same paragraph (bookmarks in this doc don't span paragraphs)."""
    runs = []
    el = start_el.getnext()
    while el is not None and el is not end_el:
        if _local(el.tag) == 'r':
            runs.append(el)
        el = el.getnext()
    return runs


def _nearby_text(el):
    p = el
    while p is not None and _local(p.tag) != 'p':
        p = p.getparent()
    if p is None:
        return ''
    text = ''.join(t.text or '' for t in p.iter(f'{{{W}}}t'))
    return text[:80]


def _scan_fields_in_paragraph(p, g):
    """
    A Word field looks like:
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> REF bookmark_name </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>cached display text</w:t></w:r>   <-- may be several runs
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    We find each begin/separate/end triple, read the instruction, and
    classify as a RefFieldNode or FormulaFieldNode.
    """
    runs = p.findall(f'{{{W}}}r')
    i = 0
    ref_counter = len(g.refs)
    formula_counter = len(g.formulas)
    while i < len(runs):
        fld = runs[i].find(f'{{{W}}}fldChar')
        if fld is not None and fld.get(f'{{{W}}}fldCharType') == 'begin':
            # gather instruction text until 'separate'
            j = i + 1
            instr = ''
            while j < len(runs):
                it = runs[j].find(f'{{{W}}}instrText')
                if it is not None:
                    instr += it.text or ''
                sep = runs[j].find(f'{{{W}}}fldChar')
                if sep is not None and sep.get(f'{{{W}}}fldCharType') == 'separate':
                    j += 1
                    break
                j += 1
            # display run(s) until 'end'
            display_start = j
            k = j
            while k < len(runs):
                endf = runs[k].find(f'{{{W}}}fldChar')
                if endf is not None and endf.get(f'{{{W}}}fldCharType') == 'end':
                    break
                k += 1
            display_run = runs[display_start] if display_start < k else None
            instr = instr.strip()
            context = ''.join(t.text or '' for t in p.iter(f'{{{W}}}t'))[:80]

            if instr.upper().startswith('REF '):
                target = instr.split()[1]
                g.refs.append(RefFieldNode(
                    id=f'ref_{ref_counter}', target_bookmark=target,
                    display_run=display_run, context=context))
                ref_counter += 1
            elif instr.startswith('='):
                g.formulas.append(FormulaFieldNode(
                    id=f'formula_{formula_counter}', formula=instr[1:].split('\\')[0].strip(),
                    display_run=display_run, context=context,
                    table_index=None, row_index=None, col_index=None))
                formula_counter += 1
            i = k
        i += 1


if __name__ == '__main__':
    g = build_graph('sample_contract.docx')
    print(g.summary_for_llm())
