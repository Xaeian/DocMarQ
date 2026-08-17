# DocMarQ

DOCX generation with a fluent API. Core is lean _(just `python-docx`)_. Optional `[md]` extra adds markdown-to-DOCX rendering with banner headers, mermaid, GitHub callouts and more. Sibling library to [`pdfmarq`](https://github.com/Xaeian/PDFMarQ) with the same API shape and `.docx` output.

## Philosophy

DocMarQ wraps `python-docx` _(OOXML zip plumbing, content types, relationships)_ into a fluent paragraph/run API. You describe document flow; Word handles layout, pagination, and reflow on open.

- **Fluent paragraph/run model**: `doc.para("First.")` opens a paragraph, `doc.text(" with bold", bold=True)` appends a styled run. Close with `enter()` or let the next block auto-open one.
- **One way per feature**: `doc.table()`, `doc.image()`, `doc.bullet()`, `doc.link()`, no overloaded call signatures
- **Markdown is optional**: core → 1 dep _(`python-docx`)_, `[md]` adds `markdown-it-py`, `mdit-py-plugins`, `PyYAML`
- **Cross-library parity**: API shape mirrors [`pdfmarq`](https://github.com/Xaeian/PDFMarQ), including `TableStyle`, `Styles`, `parse_color`, `rgb255`, page sizes, and `lang_style()` for i18n. The same markdown source can target both PDF and DOCX.
- **Word-native output**: opens cleanly in Word, LibreOffice, Google Docs. Templates `.dotx` / `.docx` are respected, so themes and styles carry over.

Trade-offs:
- No cursor or coordinate control. Word owns layout. Great for content-driven documents, not for pixel-perfect grids _(use `pdfmarq` if you need that)_.
- Math becomes a native Word equation, editable and selectable, for the common markdown-math subset. Beyond it the formula is drawn as a picture, and only what neither layer understands is printed as its own LaTeX source, with a warning.
- Syntax highlighting in code blocks is not rendered yet. The `language` argument is accepted but ignored.
- The dep tree is small but `python-docx` is the only path to OOXML. If it can't express something _(e.g. complex equation OMML)_, neither can DocMarQ. Drop to `doc.doc` for raw `python-docx` access.

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
from docmarq.md import md_to_docx, MarkdownStyle
# Markdown to DOCX
style = MarkdownStyle(
  body_family="Calibri",
  mono_family="Consolas",
  line_height=1.4,
)
md_to_docx(open("doc.md").read(), "doc.docx", style=style)
```

See [`example.py`](example.py) for an end-to-end CLI script: language preset, Word-native fonts, `link_root` for cross-document references, and `base_dir` for relative images.

## Markdown features

- GitHub-flavored markdown _(tables, fenced code, lists, strikethrough)_
- YAML frontmatter rendered as a page-1 banner _(logo, status badge, version, sign block)_
- Page geometry, fonts, banner toggles and locale come from the caller's `style=`, never from the document _(see [`md-guide`](md-guide.md))_
- Built-in language presets _(en|pl|de|fr|es|it|cs|sk)_ via `lang_style()`: covers banner labels, callouts, date format
- Skip-duplicate-title: drops `# X` when it matches frontmatter `title`
- Auto-slugged headings with clickable `[text](#anchor)` internal links _(unicode-aware)_
- Local-path links configurable via `link_root` + `link_base` _(or per-doc YAML `base:`)_
- Mermaid diagrams via `mermaid-cli` _(local)_ or `mermaid.ink` _(network fallback)_, with a shared cache with `pdfmarq`
- Footnotes, emoji shortcodes `:rocket:`, nested lists, blockquotes, GitHub callouts _(`> [!NOTE]`, `> [!WARNING]`, …)_
- Images with size caps for block and inline use _(`![alt](logo.svg)` works inline at x-height)_
- Headerless single-row tables for label/value cards
- Setext-heading-with-image recovery: `![](img.svg)\n---` renders as block image + `<hr>` instead of a thumbnail-sized setext h2

Not supported _(use `pdfmarq` if you need them)_: syntax highlighting in code blocks, deferred `Page N/M` totals _(Word's page-numbering fields compute live)_.

## Modules

| Module       | Description                                         | Docs                                         |
| ------------ | --------------------------------------------------- | -------------------------------------------- |
| `docmarq`    | Core DOCX API _(fluent paragraph/run model)_        | [docmarq/readme.md](docmarq/readme.md)       |
| `docmarq.md` | Markdown-to-DOCX renderer _(optional `[md]` extra)_ | [docmarq/md/readme.md](docmarq/md/readme.md) |

## See also

Need PDF instead of `.docx`? Check [**PDFMarQ**](https://github.com/Xaeian/PDFMarQ), the sibling library with the same API shape and PDF output. It adds syntax highlighting and pre-measured page breaks, and covers a wider slice of LaTeX. Otherwise feature parity _(banner, callouts, mermaid, lang presets)_.
