# Architecture

The whole design is **layered defense**: each layer assumes the one before it will sometimes fail, so a defect has to slip past *all* of them to reach the user.

## Four hurdles — miss one and defects leak

1. **Clone-generate, never rebuild.** Fill content into *cloned* format objects from the real source. Reconstructing from an abstraction (reading text, rebuilding) throws away the structural layer (styles, numbering, section/page breaks, fields). This single rule prevents most of the breakage at the source.

2. **Prior knowledge — a sourced checklist, not invented rules.** The checks come from authoritative domain sources (what actually causes rejection/failure), not from what the author *guessed* mattered. Invented checklists have blind spots exactly where you didn't think to look.

3. **Vision — inspect the rendered result, not the source.** Some defects only exist after rendering: missing-glyph boxes, tables past the page edge, a heading stranded at a page bottom, a figure covering text. The source can look perfect while the render is broken. This layer renders and *looks*.

4. **Enforced verification — coverage, not sampling; and test the checker.** The failure that ships is rarely "couldn't detect it" — it's "never ran the check on that part," or "the check was silently blind." So: cover *every* unit (every page, every cell), and **mutation-test the checker itself** (below).

## The two gates

| | Gate A · structure | Gate B · visual |
|---|---|---|
| reads | the **declared** layer (XML) | the **rendered** layer (pixels) |
| how | deterministic code | a multimodal model or human *looking* |
| catches | missing TOC, no page breaks, all-`Normal` styles, broken numbering, leftover placeholders, wrong table content | missing glyphs, table overflow, stranded headings, figures over text, page-number errors |
| discipline | scans every element | looks at **every page**, never a sample |

Neither alone is enough: XML is blind to rendering; rendering is slow to enumerate structure.

## Verification discipline (why the gates are trustworthy)

- **Separate verifier > self-check.** Verifying is easier than generating (the *generator-verifier gap*), and a separate review pass beats self-refinement. The gates are separate from generation, and one of them uses a tool (rendering) the generator didn't.
- **Chain-of-verification.** Answer "is every field filled?" by *enumerating fields*, not by recalling "I filled them." Verify against the artifact, not your memory of intent.
- **Coverage, not sampling.** "Done" is impossible below 100% coverage of the checklist.
- **Mutation-test the gate.** Before trusting a checker, inject each fault class it should catch and confirm it turns red. A surviving fault means a blind spot — fix the *checker* first. This is the one step most verification pipelines skip, and it's the one that catches a checker that has silently rotted.
- **Requirement traceability.** Each requirement maps to one machine assertion; "done" = every requirement traced to a passing check — not "looks finished."

## References

- Generator-verifier gap — https://hazyresearch.stanford.edu/blog/2025-06-18-weaver
- Chain-of-Verification (CoVe), ACL 2024 — https://arxiv.org/abs/2309.11495
- Mutation testing — https://arxiv.org/pdf/2102.11378
