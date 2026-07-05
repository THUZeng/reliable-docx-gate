# -*- coding: utf-8 -*-
"""
format_gate — two-gate format verification for structured .docx.
Run before delivery; a document is "format-correct" only when BOTH gates are green.

- Gate A · structural lint (reads XML hard facts — the "declared" layer):
  TOC / cover / page breaks / heading styles / numbering continuity / orphan styles /
  mid-sentence splits / deviation-table-satisfied / leftover placeholders / fonts /
  structural fingerprint vs a reference document.

- Gate B · visual (render to PNG; a multimodal model or a human looks at EVERY page):
  missing-glyph boxes, a table past the page edge, a stranded heading, a figure over
  text, page-number continuity — the "rendered" layer XML cannot see.

XML is reliable only for the "declared" layer. The "rendered" layer (missing glyphs,
overflow, pagination) is not visible in XML — so Gate B is mandatory, not optional.

NOTE: built for Chinese government-procurement bid documents, so the *document-content
match patterns* are Chinese (kept as-is, with English comments) while the code and the
report labels are English. The method generalizes — swap the domain patterns for yours.

Usage:
    from format_gate import run_gate
    run_gate("draft.docx", reference="template.docx", out_png_dir="/tmp/gateB")
    # -> prints the Gate A report + renders Gate B PNGs (then look at every page)
"""
import os, re, subprocess, glob
from collections import Counter
from docx import Document
from docx.oxml.ns import qn

# sentence/clause-ending punctuation a paragraph should normally end with
TERMINAL = "。！？；：）」》】.!?;:)"
# style names considered "known-good" (incl. common Chinese Word style names)
GOOD_STYLES = {"Normal", "Normal Indent", "标题", "正文", "正文2", "CM14",
               "Heading 1", "Heading 2", "Heading 3", "toc 1", "toc 2", "TOC Heading",
               "List Paragraph", "List Paragraph1", "部分1", "Plain Text"}
# Chinese section ordinals: 一、二、… (matches headings in the document)
SEC_RE = re.compile(r"^\s*([一二三四五六七八九十]+)\s*、")

def _xml(doc):
    return doc.element.body.xml

def _has_toc_field(doc):
    x = _xml(doc)
    return ("instrText" in x and "TOC" in x) or ("fldSimple" in x and "TOC" in x)

def _has_manual_toc(doc):
    # a "目录" (Contents) heading followed by lines ending in page numbers (tab + number)
    # counts as a manual TOC.
    ps = doc.paragraphs
    for i, p in enumerate(ps):
        if p.text.strip() in ("目录", "目 录"):          # "Contents" heading text
            tail = " ".join(q.text for q in ps[i+1:i+15])
            if re.search(r"\d+\s*$", tail) or "\t" in tail:
                return True
    return False

def _effective_eastasia_font(p):
    for r in p.runs:
        rpr = r._element.rPr
        if rpr is not None:
            f = rpr.find(qn('w:rFonts'))
            if f is not None and f.get(qn('w:eastAsia')):
                return f.get(qn('w:eastAsia'))
    return None  # inherited from a style, not explicitly declared

def gate_a(docx_path, reference=None):
    d = Document(docx_path)
    ps = d.paragraphs
    x = _xml(d)
    R = []  # list of (ok, label, detail); ok = True / False / None(=needs Gate B)

    # 1. page breaks (their absence is why content runs together / page numbers break)
    #    both count as a real page break: explicit <w:br type=page> and <w:pageBreakBefore>
    pb = x.count('w:type="page"') + x.count('<w:pageBreakBefore')
    R.append((pb > 0, "has page break", f"{pb} found"))

    # 2. TOC (a real TOC field is preferred; a manual TOC also counts)
    toc = _has_toc_field(d) or _has_manual_toc(d)
    R.append((toc, "has TOC", "TOC field" if _has_toc_field(d) else ("manual TOC" if _has_manual_toc(d) else "none")))

    # 3. cover page (a large centered title near the top, before the body)
    cover = False
    for p in ps[:6]:
        sz = p.runs[0].font.size.pt if (p.runs and p.runs[0].font.size) else 0
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        if p.text.strip() and sz and sz >= 18 and p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            cover = True; break
    R.append((cover, "has cover page", "yes" if cover else "not detected"))

    # 4. heading styles (100% Normal = the structural layer was flattened)
    #    common pitfall: rebuilding from .text collapses every style to Normal
    sc = Counter(p.style.name for p in ps)
    head_styles = sum(v for k, v in sc.items() if ("eading" in k) or k.startswith("标题"))
    ratio_normal = sc.get("Normal", 0) / max(1, len(ps))
    R.append((head_styles > 0, "uses heading styles (not all-Normal)",
              f"heading paras={head_styles}, Normal ratio={ratio_normal:.0%}"))

    # 5. orphan / foreign styles (dragged in by paste, e.g. 纯文本_0_0 / AONormal)
    orphans = [k for k in sc if k not in GOOD_STYLES and sc[k] <= 3 and not k.startswith("Heading")]
    R.append((len(orphans) == 0, "no orphan styles", ("clean" if not orphans else "suspect: " + ", ".join(orphans))))

    # 6. section numbering continuous (no missing 三、/ no duplicate 5、)
    #    if heading styles exist, only check section HEADINGS (skip nested 一二三 sub-lists)
    seq = []
    zh = "一二三四五六七八九十"
    head_ps = [p for p in ps if ("eading" in p.style.name) or p.style.name.startswith("标题")]
    target = head_ps if head_ps else ps
    for p in target:
        m = SEC_RE.match(p.text)
        if m and len(p.text.strip()) < 45:
            seq.append(zh.index(m.group(1)[0]) + 1 if m.group(1) in zh else None)
    dup = [n for n in set(seq) if seq.count(n) > 1]
    gap = []
    s = [n for n in seq if n]
    if s:
        for n in range(min(s), max(s) + 1):
            if n not in s: gap.append(n)
    ok6 = (not dup and not gap)
    R.append((ok6, "section numbering continuous",
              f"seq={seq[:12]}{' dup='+str(dup) if dup else ''}{' gap='+str(gap) if gap else ''}"))

    # 7. mid-sentence line splits (short paragraph with no terminal punctuation =>
    #    OCR / text-rebuild split one visual line into its own paragraph)
    #    skip: cover region / headings / label lines with a colon / dates / captions / TOC placeholder
    DATE_RE = re.compile(r"\d+\s*年|\d+\s*月")               # matches Chinese dates
    splits = []
    for i, p in enumerate(ps):
        t = p.text.strip()
        if i < 18: continue
        if ("eading" in p.style.name) or p.style.name.startswith("标题"): continue
        if ("：" in t) or (":" in t) or DATE_RE.search(t): continue
        if t.startswith("图") or t.startswith("表") or ("更新域" in t) or ("目录" in t): continue  # 图/表=figure/table caption
        if 0 < len(t) <= 28 and t[-1] not in TERMINAL:
            nxt = ps[i+1].text.strip() if i+1 < len(ps) else ""
            if nxt and nxt[0] not in "一二三四五六七八九十（(0123456789注图表":
                splits.append(t[-14:])
    R.append((len(splits) == 0, "no mid-sentence line splits",
              ("ok" if not splits else f"{len(splits)} suspect, e.g. " + " / ".join(splits[:3]))))

    # 8. deviation table all "satisfied" (a negative ▲ deviation = disqualification)
    #    domain: 偏离表 = deviation table; 满足 = satisfied
    dev_ok, dev_detail = None, "no deviation table"
    for tb in d.tables:
        header = " ".join(c.text for c in tb.rows[0].cells)
        if "偏离" in header:                                  # "deviation" in header
            col = next((j for j, c in enumerate(tb.rows[0].cells) if "偏离" in c.text), None)
            if col is not None:
                vals = [tb.rows[i].cells[col].text.strip() for i in range(1, len(tb.rows))]
                POS = ("满足", "无偏离", "正偏离", "响应")       # satisfied / no-deviation / positive / responds
                # ⚠ "不满足" (NOT satisfied) CONTAINS the substring "满足" (satisfied) — you must
                #   check the negatives FIRST, or a negative deviation slips through. (Found by the
                #   mutation test in gate_selftest.py.)
                NEG = ("不满足", "未满足", "负偏离", "不响应", "部分满足", "有偏离", "不符", "偏高", "偏低")
                bad = [v for v in vals if v and (any(n in v for n in NEG) or not any(p in v for p in POS))]
                dev_ok = (len(bad) == 0)
                dev_detail = "all satisfied" if dev_ok else f"possible negative: {bad[:3]}"
            break
    R.append((dev_ok, "deviation table all satisfied", dev_detail))

    # 8b. no leftover placeholders (XX / XXX / "see page N" / empty brackets / underscores)
    #     common pitfall: filling only EMPTY cells leaves placeholder TEXT in place.
    #     whitelist: "（此处加盖电子公章）" (e-seal marker) and real doc numbers like
    #     「〔2017〕」/「【项目编号:...】」 are kept, not flagged.
    _ph_pats = [
        (re.compile(r"X{2,}"), "XX/XXX"),
        (re.compile(r"见投标文件第\s*页"), "see-page-N"),      # "see page __ of the bid" placeholder
        (re.compile(r"〔\s*〕"), "empty 〔〕"),
        (re.compile(r"【\s*】"), "empty 【】"),
        (re.compile(r"[_＿]{4,}"), "____underscore"),
    ]
    def _scan_placeholder(text):
        out = []
        for rgx, lbl in _ph_pats:
            if rgx.search(text):
                out.append(lbl)
        # （此处…） "here fill in…" is a placeholder EXCEPT the e-seal marker
        for m in re.finditer(r"（此处[^）]*）", text):
            if "加盖电子公章" not in m.group(0):              # keep the e-seal marker
                out.append("（此处…）fill-in")
        return out
    all_texts = [p.text for p in ps]
    for tb in d.tables:
        for row in tb.rows:
            for c in row.cells:
                all_texts.append(c.text)
    ph_hits = []
    for text in all_texts:
        for lbl in _scan_placeholder(text):
            ph_hits.append((lbl, text.strip().replace("\n", " ")[:24]))
    R.append((len(ph_hits) == 0, "no leftover placeholders",
              ("clean" if not ph_hits else f"{len(ph_hits)} found: " + " / ".join(f"[{l}]{s}" for l, s in ph_hits[:4]))))

    # 9. body CJK font (declared layer only; the rendered font is Gate B's job)
    fonts = Counter()
    for p in ps:
        f = _effective_eastasia_font(p)
        if f and len(p.text.strip()) > 10: fonts[f] += 1
    R.append((None, "body CJK font (declared)", dict(fonts.most_common(4)) or "mostly inherited; check Gate B"))

    # structural fingerprint vs a reference (template / known-good sample)
    if reference and os.path.exists(reference):
        rd = Document(reference); rx = _xml(rd)
        ref_pb = rx.count('w:type="page"'); ref_toc = _has_toc_field(rd) or _has_manual_toc(rd)
        rsc = Counter(p.style.name for p in rd.paragraphs)
        ref_head = sum(v for k, v in rsc.items() if "eading" in k or k.startswith("标题"))
        diffs = []
        if ref_pb > 0 and pb == 0: diffs.append("reference has page breaks, draft has none")
        if ref_toc and not toc: diffs.append("reference has a TOC, draft has none")
        if ref_head > 0 and head_styles == 0: diffs.append("reference has heading styles, draft is all-Normal")
        R.append((len(diffs) == 0, "structural fingerprint vs reference", ("aligned" if not diffs else "; ".join(diffs))))

    return R

def render_pages(docx_path, out_dir, dpi=110):
    """Gate B step 1: render to per-page PNGs; return the paths (then look at each one)."""
    os.makedirs(out_dir, exist_ok=True)
    for f in glob.glob(os.path.join(out_dir, "*")):
        try: os.remove(f)
        except: pass
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                   check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    pdf = glob.glob(os.path.join(out_dir, "*.pdf"))
    if not pdf: return []
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf[0], os.path.join(out_dir, "page")],
                   check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    return sorted(glob.glob(os.path.join(out_dir, "page*.png")))

def run_gate(docx_path, reference=None, out_png_dir=None):
    print("=" * 74)
    print("Gate A · structure:", os.path.basename(docx_path))
    print("=" * 74)
    R = gate_a(docx_path, reference)
    passed = True
    for ok, name, detail in R:
        mark = "PASS" if ok is True else ("FAIL" if ok is False else " -> ")
        if ok is False: passed = False
        print(f" [{mark}] {name:<34} {detail}")
    print(f"\nGate A: {'ALL GREEN' if passed else 'HAS RED -- must fix to all-green'}   ( -> = hand to Gate B )")
    if out_png_dir:
        pngs = render_pages(docx_path, out_png_dir)
        print(f"\nGate B · rendered {len(pngs)} pages -> LOOK at every page "
              f"(tofu / table off-edge / stranded heading / figure over text / page numbers):")
        for p in pngs: print("   ", p)
    return passed, R

if __name__ == "__main__":
    import sys
    ref = sys.argv[2] if len(sys.argv) > 2 else None
    out = sys.argv[3] if len(sys.argv) > 3 else "/tmp/gateB"
    run_gate(sys.argv[1], ref, out)
