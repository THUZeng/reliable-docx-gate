# -*- coding: utf-8 -*-
"""
gate_selftest — mutation-test the gate itself. Prove the checker can catch what it claims to,
BEFORE you trust it on real work.

Take a known-good file that already passes Gate A, inject one known fault at a time
(a stray XXX / deleted page breaks / all-Normal styles / a broken deviation cell /
a deleted TOC / an orphan style ...), run Gate A, and assert the matching check turns red.
If a fault survives, the gate is BLIND to it -> fix the gate first, THEN verify documents.

This is the cure for the "it accepts a defect it should catch" failure. It is exactly how
the deviation check got caught accepting "不满足" ("not satisfied") because that string
CONTAINS "满足" ("satisfied").

Reference: software mutation testing (arxiv 2102.11378) — coverage alone guarantees little;
catching seeded faults is what proves a checker works.

Usage:  python3 gate_selftest.py a-known-good.docx
Exit code 0 = every fault class was caught; non-zero = the gate has a blind spot (printed).
"""
import sys, os, tempfile
from docx import Document
from docx.oxml.ns import qn
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from format_gate import gate_a

# ---- each mutator takes a Document and injects one class of fault (in place) ----
# (injected strings stay Chinese because they mimic real Chinese-document defects)
def m_inject_xxx(d):
    if d.tables: d.tables[0].rows[-1].cells[0].paragraphs[0].add_run("XXX")
    else: d.add_paragraph("XXX")
def m_inject_seepage(d):
    tgt = d.tables[0].rows[-1].cells[-1] if d.tables else None
    (tgt.paragraphs[0].add_run("见投标文件第页") if tgt else d.add_paragraph("见投标文件第页"))
def m_inject_emptyparen(d):
    d.add_paragraph("（此处填写项目负责人）")   # a "here fill in..." placeholder
def m_strip_pagebreaks(d):
    # really delete them (not just set val=0), to mimic a genuinely un-paginated doc
    for pbb in d.element.body.findall('.//'+qn('w:pageBreakBefore')):
        pbb.getparent().remove(pbb)
    for br in d.element.body.findall('.//'+qn('w:br')):
        if br.get(qn('w:type')) == 'page': br.getparent().remove(br)
def m_flatten_styles(d):
    for p in d.paragraphs:
        try: p.style = d.styles['Normal']
        except Exception: pass
def m_break_deviation(d):
    for tb in d.tables:
        hdr = " ".join(c.text for c in tb.rows[0].cells)
        if "偏离" in hdr:                                   # deviation table
            col = next((j for j,c in enumerate(tb.rows[0].cells) if "偏离" in c.text), None)
            if col is not None and len(tb.rows) > 1:
                cell = tb.rows[1].cells[col]
                for r in cell.paragraphs[0].runs: r.text = ""
                cell.paragraphs[0].add_run("不满足")        # "not satisfied"
            break
def m_strip_toc(d):
    for p in list(d.paragraphs):
        if p.text.strip() in ("目录","目 录"): p._p.getparent().remove(p._p); break  # "Contents" heading
def m_orphan_style(d):
    st = d.styles.add_style('_orphan_test_', 1) if '_orphan_test_' not in [s.name for s in d.styles] else d.styles['_orphan_test_']
    if d.paragraphs: d.paragraphs[-1].style = st

# (name, mutator, substring of the Gate A check-label expected to turn red)
MUTATIONS = [
    ("inject XXX placeholder",           m_inject_xxx,        "leftover placeholders"),
    ("inject see-page-N placeholder",    m_inject_seepage,    "leftover placeholders"),
    ("inject fill-in placeholder",       m_inject_emptyparen, "leftover placeholders"),
    ("delete all page breaks",           m_strip_pagebreaks,  "page break"),
    ("flatten all to Normal style",      m_flatten_styles,    "heading styles"),
    ("set a deviation cell to negative", m_break_deviation,   "deviation table all satisfied"),
    ("delete the TOC",                   m_strip_toc,         "TOC"),
    ("inject an orphan style",           m_orphan_style,      "orphan"),
]

def find(R, sub):
    return next((it for it in R if sub in it[1]), None)

def run(good_path):
    print("=" * 70); print("mutation-test the gate:", os.path.basename(good_path)); print("=" * 70)
    # baseline: the good file must itself be clean
    base = gate_a(good_path)
    base_reds = [n for ok, n, _ in base if ok is False]
    print("baseline (good file) Gate A reds:", base_reds or "none (clean)")
    blind = []
    for name, mutate, expect in MUTATIONS:
        base_item = find(base, expect)
        # if this check doesn't apply to this file (e.g. no deviation table) -> N/A, not a blind spot
        if base_item is not None and base_item[0] is None:
            print(f"  [N/A ] {name:<32} -> '{expect}' not applicable to this file"); continue
        d = Document(good_path)
        try: mutate(d)
        except Exception as e:
            print(f"  [warn] mutation '{name}' failed to apply: {e!r} -- skipped"); continue
        tmp = os.path.join(tempfile.gettempdir(), "gate_mut.docx"); d.save(tmp)
        R = gate_a(tmp)
        item = find(R, expect)
        caught = item is not None and item[0] is False
        was_red_at_base = base_item is not None and base_item[0] is False
        ok = caught and not was_red_at_base
        print(f"  {'[caught]        ' if ok else '[MISS! blind gate]'} {name:<32} -> '{expect}' {'turned red' if caught else 'stayed green'}")
        if not ok: blind.append((name, expect))
    print("-" * 70)
    if blind:
        print("BLIND SPOTS (fix format_gate before verifying documents):")
        for n, e in blind: print(f"   x {n} -- '{e}' not caught")
        return 1
    print("OK — every known fault class is caught by Gate A. The gate is fit to verify documents.")
    return 0

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python3 gate_selftest.py a-known-good.docx"); sys.exit(2)
    sys.exit(run(sys.argv[1]))
