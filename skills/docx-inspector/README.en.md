# docx-inspector

[![CI](https://github.com/ChrisBBmouse/aiwork/actions/workflows/ci.yml/badge.svg)](https://github.com/ChrisBBmouse/aiwork/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](../../LICENSE)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](pyproject.toml)

**Layout audit for any `.docx`.** It does not generate documents — it answers one question:
*will this Word file cause problems when I hand it over?*

[中文说明](README.md)

```bash
python -m docx_inspector report.docx
```

```
docx-inspector 0.1.0
file: report.docx  34.8 KB
content box: 14.66 x 24.62 cm   margins: L3.17 R3.17 T2.54 B2.54 cm

ERROR  B7   Table width 20.00cm exceeds content box 14.66cm (overflow 5.34cm).
            fixed layout — will overflow on render.   table #1 (4 cols)

9 checks · 1 error / 0 warning / 1 info
result: FAIL (exit 1)
```

## Why it exists

AI is now good at *generating* Word documents. But nobody owns the *acceptance* step —
checking whether the generated file is actually deliverable:

| | Generates | Edits | Audits **any** docx | Quantified verdict |
|---|---|---|---|---|
| Anthropic docx skill | ✅ | ✅ | ❌ | ❌ |
| gongwen-gbt9704 | ✅ | ❌ | ❌ own output only | partial |
| docx-cli | ✅ | ✅ | ❌ | ❌ |
| **docx-inspector** | ❌ | ❌ | ✅ | ✅ |

It does not compete with generators — it sits **downstream of all of them**.

## It is stricter than "it looks fine"

The same "table too wide" finding means two different things, depending on layout mode:

| Layout | Declared width | Actual render | Verdict |
|---|---|---|---|
| `auto` | 0.38cm over | Word **compresses** it back into the content box | ⚠️ warning (WPS / online preview may still overflow) |
| `fixed` | 5.34cm over | Word **draws it anyway** — the 4th column is cut off the page | 🔴 error |

![overflow annotation](docs/overflow.png)

`docx-inspector` distinguishes them because it reads `tblLayout` instead of shouting
at the first number that looks too big.

## Render-level evidence: `--annotate`

```bash
python -m docx_inspector report.docx --annotate out/
```

Renders the document to images and **draws a red box around the exact region that
crosses the right edge of the content box**, labelled with how far it goes.

Coordinates are measured from the rendered PDF — not read from the XML declarations.
The declaration says 14.66cm; the render is the truth.

## The 12 rules

Organised by **5 physical systems** on purpose. Classifying by document type
(paper / audit / contract / tender) never ends — but any docx only breaks in five places.

| System | ID | Rule | Default |
|---|---|---|---|
| Package | A1 | Invalid ZIP header (also detects DLP transparent encryption) | ✅ |
| | A3 | Relationship target missing (image / embedded object) | ✅ |
| | A4 | Media file corrupt or extension/format mismatch | ✅ |
| Geometry | B5 | Abnormal margins or content box | ✅ |
| | B7 | Table width exceeds content box (auto vs fixed aware) | ✅ |
| | B8 | Image width exceeds content box | ✅ |
| Fonts | C10 | Referenced font not installed (silent substitution) | ✅ |
| | C12 | `theme` font attributes override explicit font name | ⬜ |
| | C14 | Inconsistent fonts among same-class elements | ⬜ |
| Flow layout | D15 | Column widths disagree with `tblGrid` | ✅ |
| | D17 | Multi-row table missing repeated header row | ⬜ |
| Numbering | E20 | `numId` reused across list batches (numbering bleed) | ✅ |

Only 9 rules are on by default. Enabling everything is how a linter gets turned off.

```bash
python -m docx_inspector f.docx --preset thesis   # +C12 +C14 +D17
python -m docx_inspector f.docx --preset blockers # A1 A3 A4 B7 B8 only
python -m docx_inspector f.docx --only B7,E20     # explicit rule set
```

## Embedding in a pipeline

```bash
python -m docx_inspector out.docx --json > report.json
echo $?   # 0 pass / 1 has error / 2 unparsable
```

```json
{
  "tool": "docx-inspector",
  "counts": {"error": 1, "warning": 0, "info": 1},
  "findings": [
    {
      "rule": "B7",
      "severity": "error",
      "message": "Table width 20.00cm exceeds content box 14.66cm (overflow 5.34cm). fixed layout — will overflow on render.",
      "location": "table #1 (4 cols)",
      "fixable": true,
      "detail": {"table_index": 0, "overflow_cm": 5.34, "layout": "fixed"}
    }
  ]
}
```

## Positioning and pairing

This skill intentionally does **one** thing: inspection. The recommended workflow is
to pair it with a document-generation skill:

```
generate with any docx skill  →  docx-inspector  →  fix what it reports  →  deliver
```

- **Not a generator** — document creation belongs to dedicated skills
  (Anthropic `docx`, `docx-cli`, `gongwen-gbt9704`, …). Run one of those first.
- **Not a content reviewer** — wording, logic and completeness have no single right answer.
- **Not a business-rule engine** — audit cross-checks or journal formatting are *data*, not code.
- **Never more than 30 rules** — if a rule can't earn its place, something else must go.

## Install

```bash
pip install python-docx
# optional, only for --annotate
pip install pymupdf pillow
# render backend: Word COM on Windows, LibreOffice (soffice) elsewhere
```

## License

MIT
