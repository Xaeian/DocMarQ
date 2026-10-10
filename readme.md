# DocMarQ

DOCX generation with a fluent API.
Core is lean _(`python-docx` plus an SVG rasterizer)_.
Optional `[md]` extra adds markdown-to-DOCX rendering: documentation with banner headers, letters and contracts in one or two columns, native Word equations, mermaid and more.
Sibling library to [`pdfmarq`](https://github.com/Xaeian/PDFMarQ) with the same API shape and `.docx` output.

## Philosophy

DocMarQ wraps `python-docx` _(OOXML zip plumbing, content types, relationships)_ into a fluent paragraph/run API.
You describe document flow; Word handles layout, pagination, and reflow on open.

- **Fluent paragraph/run model**: `doc.para("First.")` opens a paragraph, `doc.text(" with bold", bold=True)` appends a styled run.
  Close with `enter()` or let the next block auto-open one.
- **One way per feature**: `doc.table()`, `doc.image()`, `doc.bullet()`, `doc.link()`, no overloaded call signatures
- **Markdown is optional**: core → `python-docx` + `svglib`/`rlPyCairo`/`Pillow` for SVG, `[md]` adds `markdown-it-py`, `mdit-py-plugins`, `PyYAML`
- **Engine and documents apart**: `docmarq.md` writes markdown content, a document package lays out the page around it
- **Banner as contract**: in `docmarq.paper` YAML frontmatter becomes a styled banner with logo, status badge, version, dates
- **Letters and contracts**: `docmarq.legal` sets a letter header from frontmatter and signing lines from the style, in one or two columns, or with language versions side by side
- **Templates and series**: `docmarq.merge` fills a markdown template with one record, or with each record of a table, a document per record
- **Cross-library parity**: API shape mirrors [`pdfmarq`](https://github.com/Xaeian/PDFMarQ), including `TableStyle`, `Styles`, `parse_color`, `rgb255`, page sizes, and `lang_style()` for i18n.
  The same markdown source and style can target both PDF and DOCX.
- **Word-native output**: opens cleanly in Word, LibreOffice, Google Docs.
  Templates `.dotx` / `.docx` are respected, so themes and styles carry over.

Trade-offs:
- No cursor or coordinate control.
  Word owns layout.
  Great for content-driven documents, not for pixel-perfect grids _(use `pdfmarq` if you need that)_.
- Math becomes a native Word equation, editable and selectable, for the common markdown-math subset.
  Beyond it the formula is drawn as a picture, and only what neither layer understands is printed as its own LaTeX source, with a warning.
- Syntax highlighting in code blocks is not rendered yet.
  The `language` argument is accepted but ignored.
- The dep tree is small but `python-docx` is the only path to OOXML.
  If it can't express something _(e.g. complex equation OMML)_, neither can DocMarQ.
  Drop to `doc.doc` for raw `python-docx` access.

## Install

```sh
pip install docmarq      # core: python-docx
pip install docmarq[md]  # + markdown rendering stack
```

## Examples

```py
from docmarq import DOCX
# Fluent core API
with DOCX("report.docx") as doc:
  doc.font("Calibri", 20, "Bold").para("Quarterly Report")
  doc.font(size=11, mode="Regular")
  doc.para("Revenue up 23% year-over-year.")
  doc.table(
    [["Q1", "120k"], ["Q2", "148k"], ["Q3", "172k"]],
    header=["Quarter", "Revenue"],
    aligns=["C", "R"],
  )
  doc.image("chart.png", width=180, height=80)
  doc.link("google.com", url="https://google.com")
```

```py
from docmarq.paper import md_to_docx, PaperStyle
# Markdown documentation to DOCX
style = PaperStyle(
  font_body="Calibri",
  font_mono="Consolas",
  page_number_label="Page", # "Page 1/5" in footer
)
md_to_docx(open("doc.md").read(), "doc.docx", style=style)
```

```py
from docmarq.legal import md_to_docx, md_parallel_to_docx, lang_style
# A contract in two balanced columns, signed by both parties
md_to_docx(open("contract.md").read(), "contract.docx",
  style=lang_style("pl", columns=2, sign="contract"))
# Language versions side by side, each from its own file
md_parallel_to_docx([pl_text, en_text], "contract-pl-en.docx",
  styles=[lang_style("pl"), lang_style("en")])
```

See [`example.py`](https://github.com/Xaeian/DocMarQ/blob/main/example.py) for an end-to-end CLI script.
It covers a language preset, Word-native fonts, `link_root` for cross-document references, and `base_dir` for relative images.

## Markdown features

- GitHub-flavored markdown _(tables, fenced code, lists, strikethrough)_
- YAML frontmatter rendered as a page-1 banner _(logo, status badge, version, dates)_ by `docmarq.paper`, with signing lines from the style
- Letters, statements and contracts by `docmarq.legal`: letter header from frontmatter, justified and hyphenated text, two balanced columns, language versions side by side _(see [`legal/guide`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/legal/guide.md))_
- Mail merge by `docmarq.merge`: placeholders, conditions, loops and included clauses, every record checked before the first document _(see [`merge/readme`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/merge/readme.md))_
- Columns in the core API: `doc.columns(2)`, and everything markdown writes flows through them
- Page geometry, fonts, banner toggles and locale come from caller's `style=`, never from the document _(see [`md/guide`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/md/guide.md))_
- Built-in language presets _(en|pl|de|fr|es|it|cs|sk)_ via `lang_style()`: covers banner labels, callouts, date format, page numbers
- Skip-duplicate-title: drops `# X` when it matches frontmatter `title`
- Auto-slugged headings with clickable `[text](#anchor)` internal links _(unicode-aware)_
- Local-path links configurable via `link_root` + `link_base`
- Math formulas inline `$x^2$` and block `$$...$$` as native Word equations
- Mermaid diagrams via `mermaid-cli` _(local)_ or `mermaid.ink` _(network fallback)_, with a shared cache with `pdfmarq`
- Footnotes, nested lists, blockquotes, GitHub callouts _(`> [!NOTE]`, `> [!WARNING]`, …)_
- Images alone in a paragraph become figures sharing text width, wrapping into a grid beyond it
- An image on the same line as text keeps only its alt text, so give a figure its own paragraph

Not supported _(use `pdfmarq` if you need them)_: syntax highlighting in code blocks, emoji shortcodes, task lists, inline HTML.

## Modules

- [`docmarq`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/readme.md): core DOCX API _(fluent paragraph/run model, columns)_
- [`docmarq.md`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/md/readme.md): markdown engine _(optional `[md]` extra)_
- [`docmarq.paper`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/paper/readme.md): documentation with a frontmatter banner
- [`docmarq.legal`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/legal/readme.md): letters, statements, contracts, parallel versions
- [`docmarq.merge`](https://github.com/Xaeian/DocMarQ/blob/main/docmarq/merge/readme.md): mail merge, one record or a document per record of a table

## See also

Need PDF instead of `.docx`?
Check [**PDFMarQ**](https://github.com/Xaeian/PDFMarQ), sibling library with the same API shape and PDF output.
It adds syntax highlighting and pre-measured page breaks, and covers a wider slice of LaTeX.
Everything else _(banner, callouts, mermaid, lang presets, letters and contracts)_ works the same in both.
