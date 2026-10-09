# `docmarq.paper`

Documentation DOCX: markdown content with a frontmatter banner around it.
Requires `pip install docmarq[md]`.
The content is written by [`docmarq.md`](../md/readme.md).

## `md_to_docx`

```py
from docmarq.paper import md_to_docx
md_to_docx(open("doc.md").read(), "doc.docx")
# Landscape, and relative images resolved against a directory
from docmarq.constants import A4
md_to_docx(md_text, "out.docx", page=A4.landscape(), base_dir="./assets")
```

`style` must be a `PaperStyle`, built directly or with `lang_style()`.

## `md_into`

```py
from docmarq import DOCX
from docmarq.paper import md_into, lang_style
doc = DOCX("report.docx")
doc.font("Calibri", 28, "Bold").para("Cover", align="C") # pages of its own
doc.font(size=11, mode="Regular").page_break()
md_into(doc, open("doc.md").read(), style=lang_style("pl", banner=False))
doc.save()
```

`md_into` writes a paper document after the last block of a `DOCX` that already holds pages, as a cover, and the caller saves it.
It sets the core properties, compact header, page number, content and signing lines as `md_to_docx` does.
Header and footer belong to the whole document, so the cover carries the page number too.
A banner keeps the first page clear of the compact header, so it belongs on page 1.

## Banner (YAML frontmatter)

YAML block at top becomes a styled banner on page 1, and its `id` or `title` a compact header on the pages after it.

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
logo: ./ranger-badge.svg
---

# Document body starts here
```

- `id`: document code, after the status and version
- `title`: main title, large
- `version`: version as written (no `v` prefix added)
- `author`: author name, shown as `{banner_label_author}: ...`
- `status`: badge: `draft` / `review` / `approved` / `deprecated` / `archived`
- `entity`: organization, bold, with `address` after it on one line
- `address`: address, muted
- `created`: ISO date, formatted via `style.date_format`
- `updated`: same as `created`
- `logo`: path to `.svg`/`.png`/`.jpg`, in a column of its own left of the banner text
- `subject`: written to the DOCX core properties, not rendered
- `keywords`: written to the DOCX core properties, string or YAML list

The DOCX core properties _(title, author, subject, keywords)_ are auto-filled from matching YAML keys.
Pass `metadata={...}` to `md_to_docx()` to override per-key.

If first body block is `# X` and `X` matches `title` exactly, h1 is dropped to avoid showing the title twice.
Only applies when banner actually printed that title, so `banner=False` never costs you the heading.
Disable with `skip_dup_title=False`.

## Presentation

Frontmatter carries content only; everything visual comes from the caller.
`style=` is used verbatim, so any value can be set, a `PaperStyle()` default included.

```python
from docmarq.paper import md_to_docx, lang_style
from docmarq.constants import A4, page_size

style = lang_style("pl", # banner/footer labels
  font_body="Calibri", font_head="Calibri", font_mono="Consolas",
  body_size=11, line_height=1.0, image_max_h=120,
  banner=True, banner_compact=True, sign="approval",
  mermaid_theme="default",
)
md_to_docx(md, "out.docx", style=style,
  page=A4, margin=25, gutter=0, base_dir=".", font_dir="./fonts")
```

`sign` takes `True` for one line, a `sign_labels` scenario or a list of custom labels.
Scenarios: `signature`, `approval` and `contract`, localized by the language preset.
Several lines sit side by side, sharing the content width.

`page` is a `PageSize` in mm: `A4`, `A4.landscape()`, `page_size("a3")` for a preset name (`A4`/`A3`/`A5`/`LETTER`/`LEGAL`, raises on anything else), or `PageSize(200, 250)` for a custom sheet.

Everything after `output_path` is keyword-only, so argument order cannot silently diverge from `pdfmarq.paper.md_to_pdf`.

## Style

Beyond fields shown above:

```py
PaperStyle(
  page_number_label="Strona",  # "Strona 1/5" footer; None to disable
  page_number_total=True,      # False → "Strona 1" without total
  date_format="%d.%m.%Y",      # strftime pattern
  skip_dup_title=True,         # drop `# X` if it matches frontmatter title
)
```

The engine's fields - fonts, sizes, `justify`, `keep_together` - are there too, since `PaperStyle` extends `MarkdownStyle`.

### Banner labels (i18n)

Labels in banner, footer, and callouts are style fields.
Defaults are English.
Use `lang_style("pl"|"de"|...)` to apply a built-in preset, or override fields manually.

Built-in presets cover `en` _(defaults)_, `pl`, `de`, `fr`, `es`, `it`, `cs`, `sk`.
Each configures `page_number_label`, `date_format`, banner labels _(author / created / updated)_, signing scenarios _(`sign_labels`)_, and the engine's callout labels.

```py
PaperStyle(
  page_number_label="Page",
  banner_label_author="Author",
  callout_label_warning="Heads up",
)
```

### Status badge palette

The default palette, `(background, text)` per status:

```py
PaperStyle(banner_status_colors={
  "draft": ((0.85, 0.87, 0.92), (0.30, 0.34, 0.45)),
  "review": ((1.00, 0.95, 0.78), (0.62, 0.40, 0.05)),
  "approved": ((0.86, 0.96, 0.87), (0.10, 0.45, 0.18)),
  "deprecated": ((1.00, 0.88, 0.85), (0.74, 0.20, 0.15)),
  "archived": ((0.92, 0.87, 0.96), (0.45, 0.25, 0.55)),
})
```

Custom keys are allowed: `status: zatwierdzony` works if `"zatwierdzony"` is in the palette.
