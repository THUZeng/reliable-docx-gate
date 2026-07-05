# reliable-docx-gate

**Make an LLM produce Word documents that don't silently break — and *prove* it.**

A self-verifying approach for generating and editing structured `.docx` files (bids, contracts, reports) that keeps the table of contents, page breaks, heading styles, and numbering intact, and refuses to ship leftover `XXX` placeholders. Two independent gates catch defects; a **mutation-tester proves the gates themselves aren't blind.**

---

## The problem this really solves

An LLM hands you a document that *looks* finished but isn't — because it works from what it **produced in the source**, not from what a reader **actually gets once the file is rendered**. The same gap, over and over:

- the source is valid but the rendered document is broken — the table of contents, page breaks, and heading styles silently disappear;
- content overflows or hides itself — a table runs past the page edge, a figure sits on top of the text beneath it;
- fields that count as "filled" to the program are blank to a human — `XXX` placeholders survive;
- a value passes a shallow check while meaning the opposite — "not satisfied" slips through because the string *contains* "satisfied".

Every one is the same failure: **the generator trusts its own account of the document, instead of checking what the document actually is once it's rendered and read.**

| | Naive approach *(what most do — and what this did first)* | Refined approach *(this repo)* |
|---|---|---|
| generation | reconstruct from an abstraction (read the text, rebuild) | **work from the real thing** — clone the source, don't retype it |
| verification | trust the declared state — "I did X" = done | **check reality, not the claim** — render/execute, inspect the *actual* result, cover everything (not a sample) |
| the checker | assume the checker works | **mutation-test the checker** — seed known faults, confirm it catches them; a checker you haven't tried to fool is one you can't trust |

The throughline: **replace "I'll look carefully" with "the machine proves it can find it."**

---

## Concrete instantiation: Word / `.docx`

The failure that started this: ask an LLM to *"clean up this Word document."* It reads the text, rebuilds the file with a docx library, and quietly the table of contents vanishes, the page breaks disappear, every heading becomes body text, `XXX` placeholders survive — and it hands back something that looks finished.

**Root cause:** `paragraph.text` is a *lossy projection*. It keeps the characters and throws away the **structural layer** — named styles, auto-numbering, section/page breaks, the TOC field. Any pipeline shaped like *read text → rebuild* is **structurally guaranteed** to produce exactly those defects. The fix is to **clone the real structure and fill content into cloned format objects — never retype.**

---

## The architecture

```
                 CORRECT-FORMAT DOCX, PROVEN
        clone (don't rebuild) -> two gates -> test the gates -> ship

  [ format truth: real template ]     [ content truth: the data to fill ]
                     \                        /
                      v                      v
        GENERATION - clone, never rebuild
        (fill content into CLONED format objects; forced TOC + page
         breaks + real heading styles; NEVER read text & retype)
                              |
                              v
        +---------------- TWO GATES (ship only if BOTH green) ---------------+
        | GATE A - STRUCTURE (this repo)   | GATE B - VISUAL                 |
        | deterministic XML lint:          | render -> image -> LOOK:        |
        | page breaks / TOC / heading      | missing glyphs / table off the  |
        | styles / numbering / NO leftover | edge / stranded heading / figure|
        | placeholders / table content /   | over text / page numbers        |
        | fingerprint vs reference         | EVERY page, no sampling         |
        +----------------------------------+---------------------------------+
                              |
                              v
        MUTATION-TEST THE GATES  (gate_selftest.py)
        inject known faults -> confirm each check turns RED.
        a miss = a blind gate -> fix the gate BEFORE trusting it.
                              |
                              v
             ship only when: Gate A green AND Gate B every-page-clean
                             AND mutation-test passes
```

Gate A is **code** (deterministic, no model variance). Gate B is **looking** (a multimodal model or a human, but it must cover *every* page). They fail on opposite things — the "declared" layer vs the "rendered" layer — which is why removing either lets a whole class through.

---

## What's in here

| file | role |
|---|---|
| `gates/format_gate.py` | **Gate A** — deterministic structural lint. Checks page breaks, TOC, heading styles (not all-`Normal`), numbering continuity, orphan/pasted styles, **leftover placeholders**, table content, fonts, and a structural fingerprint vs. a reference document. |
| `gates/gate_selftest.py` | **the mutation-tester** — seeds known faults into a known-good file (stray `XXX`, deleted page breaks, all-`Normal` styles, a broken deviation cell, missing TOC …) and asserts each gate check turns red. If a fault survives → the gate is blind. |

> Built for Chinese government-procurement bid documents, so the check labels and some domain rules (e.g. the ▲-deviation table) are in Chinese and bid-specific. **The method generalizes to any structured `.docx`;** swap the domain checks for yours.

## Run it

```bash
pip install python-docx           # Gate A
# + LibreOffice (soffice) and poppler (pdftoppm/pdftotext) for rendering

# Gate A — structural lint of a draft against a reference template
python3 gates/format_gate.py draft.docx reference_template.docx

# Mutation-test the gate itself — prove it can catch what it claims to
python3 gates/gate_selftest.py a-known-good.docx
```

## The one idea worth stealing

**Mutation-test your checker.** Most people build a verifier and trust it. `gate_selftest` deliberately breaks a good file in each way the gate is *supposed* to catch, and confirms it does — which caught a real bug where the checker accepted `不满足` ("not satisfied") because that string *contains* `满足` ("satisfied"). **A verifier you haven't tried to fool is a verifier you can't trust.**

Grounded in the [generator-verifier gap](https://hazyresearch.stanford.edu/blog/2025-06-18-weaver), [chain-of-verification](https://arxiv.org/abs/2309.11495), and software [mutation testing](https://arxiv.org/pdf/2102.11378).

## License
[MIT](LICENSE)
