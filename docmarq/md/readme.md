# `docmarq.md`

Markdown-to-DOCX renderer with YAML frontmatter. Requires `pip install docmarq[md]`.

## `md_to_docx`

```py
from docmarq.md import md_to_docx
md_to_docx(open("doc.md").read(), "doc.docx")
# Force landscape orientation (overrides `landscape: true` in YAML)
md_to_docx(md_text, "out.docx", landscape=True)
# Relative image paths resolved against base_dir
md_to_docx(md_text, "out.docx", base_dir="./assets")
```

## Banner (YAML frontmatter)

YAML block at the top becomes a styled banner on page 1.

```yaml
---
id: TXR-1991-007
title: Roundhouse kick deployment protocol
version: 1.3.2
author: Walker, Texas Ranger
status: approved
entity: Texas Ranger Division
address: 1 Lone Star Boulevard, Dallas TX 75201
created: 1993-04-21
updated: 2026-03-15
sign: true
logo: ./ranger-badge.svg
---

# Document body starts here
```

| Field       | Effect                                                                       |
| ----------- | ---------------------------------------------------------------------------- |
| `id`        | Document code in code-style box                                              |
| `title`     | Main title, centered, large                                                  |
| `version`   | Version in code-style box                                                    |
| `author`    | Author name, shown as `{banner_label_author}: ...`                           |
| `status`    | Badge: `draft` / `review` / `approved` / `deprecated` / `archived`           |
| `entity`    | Organization (left of banner, bold)                                          |
| `address`   | Address (right of banner, muted)                                             |
| `created`   | ISO date, formatted via `style.date_format`                                  |
| `updated`   | Same                                                                         |
| `sign`      | `true` adds a dashed signature line + label at the end                       |
| `logo`      | Path to `.svg`/`.png`/`.jpg`, aspect-aware                                   |
| `subject`   | Written to DOCX core properties `/Subject`, not rendered                     |
| `keywords`  | Written to DOCX core properties `/Keywords`, string or YAML list             |

Aliases: `code` → `id`, `company` → `entity` _(legacy)_.

DOCX core properties _(`/Title`, `/Author`, `/Subject`, `/Keywords`)_ auto-fill from matching YAML keys. Pass `metadata={...}` to `md_to_docx()` to override per-key.

If the first body block is `# X` and `X` matches `title` exactly, the h1 is dropped to avoid showing the title twice.

## Render block

Geometry, fonts, chrome, and locale live under a nested `render:` key in the same frontmatter. All optional; library defaults apply when absent. Mirrors `pdfmarq.md.render` field-for-field.

```yaml
---
title: Report
render:
  page: A4              # A4 / A3 / A5 / LETTER / LEGAL
  margin: 25            # mm; or list [top, right, bot, left]
  gutter: 0             # mm binding margin (Word native gutter)
  landscape: false      # flip page
  font_body: Calibri
  font_head: Sora       # defaults to `font_body`
  font_mono: Consolas
  font_size: 11         # body pt
  line_height: 1.0      # Word "Single" rule (~1.2x in render due to Calibri metrics)
  img_max_h: 120        # mm cap on every image (per-image DSL still overrides)
  banner: true          # page-1 banner from frontmatter
  banner_min: true      # mini-banner on continuation pages (alias: `header`)
  page_number: true     # footer numbering
  lang: pl              # banner/footer labels (en/pl/de/fr/es/it/cs/sk)
---
```

Precedence: `MarkdownStyle()` defaults < lang preset < frontmatter `render:` keys < caller's `style=` non-default fields.

Top-level `landscape:` in frontmatter is deprecated and warns at parse time; move it under `render.landscape:`.

## Internal links

Markdown anchor links work out of the box:

```md
See the [Notify characteristic](#bluetooth-low-energy) section.

## Bluetooth Low Energy
...
```

Each heading auto-registers a GitHub-style slug _(lowercase, spaces → hyphens, unicode preserved)_. Links to non-existent slugs render as plain text rather than crashing. Footnote refs `[^1]` jump to their definitions the same way.

## Local links

Paths without a schema _(`[x](file.md)`, `[x](folder/doc)`, `[x](/absolute/path)`)_ get the link style _(blue + underline)_ but no clickable action by default. Set `link_root` to make them real URLs:

```py
MarkdownStyle(
  link_root="https://docs.company.com",  # root to prepend
  link_base="projects/foo",              # subfolder this doc sits in
)
```

Resolution:
- `[x](file.md)` → `https://docs.company.com/projects/foo/file.md`
- `[x](/abs/path)` → `https://docs.company.com/abs/path` _(absolute ignores base)_

## Style

```py
from docmarq.md import MarkdownStyle
style = MarkdownStyle(
  body_family="Calibri",
  head_family="Sora",          # None = inherit from body_family
  mono_family="Consolas",
  line_height=1.0,             # 1.0 = Word "Single" rule
  para_gap_pt=6,               # pt between paragraphs
  page_number_label="Strona",  # "Strona 1/5" footer; None to disable
  page_number_total=True,      # False → "Strona 1" without total
  date_format="%d.%m.%Y",      # strftime pattern
  banner_render=True,          # page 1 full banner
  mini_banner_render=True,     # header line on pages 2+
  image_max_h=120,             # mm - cap tall images and diagrams
)
md_to_docx(md_text, "out.docx", style=style)
```

### Banner labels (i18n)

Labels in the banner, footer, and callouts are style fields. Defaults are English. Use `lang_style("pl"|"de"|...)` to apply a built-in preset, or override fields manually.

```py
from docmarq.md import lang_style, md_to_docx
style = lang_style("pl", body_family="Calibri")
md_to_docx(md_text, "out.docx", style=style)
```

Built-in presets ship in `docmarq/md/presets.py` and cover `en` _(defaults)_, `pl`, `de`, `fr`, `es`, `it`, `cs`, `sk`. Each preset configures `page_number_label`, `date_format`, banner labels _(author / created / updated / signature)_, and callout labels _(note / tip / important / warning / caution)_.

For ad-hoc overrides without a preset, set fields directly:

```py
MarkdownStyle(
  page_number_label="Page",
  banner_label_author="Author",
  callout_label_warning="Heads up",
)
```

### Status badge palette

```py
MarkdownStyle(banner_status_colors={
  "draft":      ((0.85, 0.87, 0.92), (0.30, 0.34, 0.45)),
  "review":     ((1.00, 0.95, 0.78), (0.62, 0.40, 0.05)),
  "approved":   ((0.86, 0.96, 0.87), (0.10, 0.45, 0.18)),
  "deprecated": ((1.00, 0.88, 0.85), (0.74, 0.20, 0.15)),
  "archived":   ((0.92, 0.87, 0.96), (0.45, 0.25, 0.55)),
})
```

Custom keys are allowed: `status: zatwierdzony` works if `"zatwierdzony"` is in the palette.

## Features

````md
# Headings h1 through h6

**bold** *italic* ***bold italic*** ~~strike~~ `inline code`

- Unordered lists
  - Nested
1. Ordered lists

| GFM | tables |
| --- | -----: |
| A   |  right |

```py
# Fenced code (no syntax highlighting in DOCX, language is accepted but ignored)
def hello(): pass
```

```mermaid
flowchart LR
  A --> B
```

> Blockquote with left border

> [!NOTE]
> GitHub-style callouts: NOTE / TIP / IMPORTANT / WARNING / CAUTION

Footnotes[^1] and emoji :rocket: :sparkles:
[^1]: Definition at end of doc.
````

Word handles pagination on open, so there's no analytical page-break pass. Paragraphs and tables reflow naturally.

## HTML support

A small whitelist of raw HTML tags is recognized inside markdown. Everything else _(`<table>`, `<div>`, `<span>`, `<style>`, attributes)_ is dropped silently.

| Tag                  | Effect                                            |
| -------------------- | ------------------------------------------------- |
| `<b>`, `<strong>`    | Bold                                              |
| `<i>`, `<em>`        | Italic                                            |
| `<code>`             | Inline code _(mono family, code colors)_          |
| `<br>`               | Hard line break                                   |
| `<hr>`               | Horizontal rule _(block-level)_                   |

### Headerless tables

Markdown tables require a header row per spec, but a single-row "card" layout is a common pattern. A table with header but no body is rendered as headerless, useful for label/value blocks and contact cards:

```md
| ![](logo.svg) | Pocket Diagnostics Poland sp. z o.o<br>80-890 Gdańsk<br>Jana Heweliusza 11/811 |
| --- | --- |
```

### Setext-heading-with-image recovery

```md
![diagram](schema.svg)
---
```

CommonMark parses this as a setext h2 with the image as heading text. `docmarq` detects the image-only setext case and renders it as a block image followed by an `<hr>`, matching the user's actual intent.

## Mixing with core API

```py
from docmarq import DOCX
from docmarq.md import MarkdownRenderer, MarkdownStyle
with DOCX("out.docx") as doc:
  doc.font("Calibri", 28, "Bold").para("Title", align="C")
  doc.page_break()
  # Markdown body (no banner - we already have a custom cover)
  MarkdownRenderer(doc, MarkdownStyle(banner_render=False)).render(open("body.md").read())
```

## Math formulas

Inline `$...$` and display `$$...$$` (also ` ```math ` fences) render as
**native Word equations (OMML)** - editable in Word's equation editor,
selectable, and scaling with font + zoom. The converter is self-contained
(built straight on OOXML, no extra runtime dependency) and covers the
common markdown-math subset:

````md
Inline $E = mc^2$ and the quadratic root $x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}$.

$$\sum_{i=1}^{n} i = \frac{n(n+1)}{2}$$

$$f(x) = \begin{cases} 1 & x > 0 \\ 0 & x \le 0 \end{cases}$$
````

Supported: superscripts/subscripts, `\frac` / `\dfrac` / `\tfrac` / `\binom`,
`\sqrt` and `\sqrt[n]`, big operators (`\sum`, `\prod`, `\int`, `\oint`, ...)
with limits, Greek + relation/operator/arrow/set symbols, functions
(`\sin`, `\log`, `\lim_{...}`, ...), delimiters (`\left ... \right`),
accents (`\hat`, `\vec`, `\bar`, `\dot`, ...), `\overline` / `\underline`,
matrices (`matrix` / `pmatrix` / `bmatrix` / `vmatrix` / `cases`), and font
commands (`\mathbb`, `\mathcal`, `\mathbf`, `\mathrm`, `\text`, ...).

LaTeX outside this subset falls back to a **matplotlib-rendered image** so
the formula still appears (never raw `$...$`). Inline fallback images are
baseline-aligned. Toggle / tune via `MarkdownStyle`:

```py
MarkdownStyle(
  math_enable=True,     # False leaves `$...$` as plain text
  math_fontset="stix",  # fallback image fontset (matplotlib mathtext)
)
```

The image fallback needs matplotlib (`pip install docmarq[math]`); without
it, unconvertible formulas degrade to italic LaTeX source text.

## Not supported (use `pdfmarq` if you need them)

- Syntax highlighting in fenced code _(language hint is accepted, ignored)_
- Deferred page-number rendering: Word uses field codes; configure via `page_number_label` only
- Rich inline content inside table cells _(math in a cell degrades to `$latex$` source text)_

## Optional deps for features

Installed by `pip install docmarq[md]`:
- `markdown-it-py`, `mdit-py-plugins`: parser + GFM plugins
- `PyYAML`: frontmatter
- `mermaid-cli` via npm for ` ```mermaid ` blocks: `npm install -g @mermaid-js/mermaid-cli` _(System tool, **not on PyPI**)_. Falls back to `mermaid.ink` HTTP service when `mmdc` is absent but network is available.

If a dep is missing, the feature silently degrades _(mermaid block renders as a regular code block)_.
