# `docmarq`

Fluent DOCX generation. Built on `python-docx` with a paragraph/run flow
model. Sibling library: [`pdfmarq`](../pdfmarq/readme.md) - same API
shape for PDF.

## `DOCX` context

```py
from docmarq import DOCX, A4
with DOCX("out.docx") as doc:
  doc.heading("Tytuł", level=1)
  doc.para("Pierwszy akapit.")
# Custom page size + margins (CSS-style: scalar / (v,h) / (t,h,b) / (t,r,b,l))
with DOCX("out.docx", width=A4.width, height=A4.height, margin=20) as doc:
  doc.margin(top=20, right=25, bot=20, left=25) # named per side
# Units (default mm)
DOCX("out.docx", unit="pt") # or "cm", "in", "px"
# Start from a Word template (.dotx / .docx) - keeps theme + styles
DOCX("out.docx", template="templates/corporate.docx")
```

## Style

```py
# Default font + size for subsequent runs
doc.font("Calibri", 11, "Bold") # `mode` accepts Regular/Bold/Italic/BoldItalic
doc.color("#1f2328")           # default text color
doc.style(line_height=1.4, space_before=4, space_after=4) # bulk update
```

## Paragraphs and runs

```py
# `para()` opens a new paragraph; `text()` appends a styled run
doc.para("First paragraph.")
doc.text(" with ").text("inline bold", bold=True).text(" then plain.")
doc.enter()                              # close paragraph
doc.text("Auto-opens", italic=True)      # opens a new para implicitly
# Soft line break inside current paragraph (Shift+Enter)
doc.text("Line one").line_break().text("Line two")
# Built-in Word style by name
doc.para("A pull quote.", style="Intense Quote", align="C")
# Run kwargs map to RichSegment fields - any combination works
doc.text("highlighted", highlight="yellow")
doc.text("monospace", code=True)
doc.text("super", superscript=True).text(" sub", subscript=True)
```

## Headings, lists

```py
doc.heading("Chapter 1", level=1)
doc.bullet("First item")
doc.bullet("Nested item", level=1)
doc.ordered("Step one")
doc.ordered("Step two")
```

## Block elements

```py
doc.blockquote("Quoted text.", indent=4, border_color="#d0d7de")
doc.hr(color="#d0d7de")
doc.code_block(
  "def foo():\n  return 42",
  language="python",        # accepted but not yet syntax-highlighted
  bg_color="#f6f8fa",
  border_color="#d0d7de",
  font_family="Consolas",
  font_size=9,
)
```

## Tables

```py
# Simple table
doc.table(
  [["a1", "b1"], ["a2", "b2"]],
  header=["A", "B"],
  aligns=["L", "R"],
  widths=[40, 60], # mm per column; omit to auto-fill
)
# Custom style (shape matches pdfmarq.TableStyle for cross-lib parity)
from docmarq import TableStyle
style = TableStyle(
  header_bg="#f6f8fa",
  header_color="#1f2328",
  header_bold=True,
  row_bg_odd=(0.985, 0.99, 0.995),
  border_color="#d0d7de",
  border_size=0.5,
  cell_pad_h=2.5,
  cell_pad_top=1.6,
  cell_pad_bot=0.6,
  font_size=10,
)
doc.table(rows, header=header, style=style)
# Or use a built-in Word table style name
doc.table(rows, word_style="Light Grid")
```

## Colors

```py
doc.color("#1f2328")           # hex
doc.color((0.03, 0.41, 0.85))  # tuple 0-1
# Conversion helpers (symmetric with pdfmarq)
from docmarq import parse_color, rgb255, color_hex
parse_color("#1f2328")    # → (0.121, 0.137, 0.156) floats 0-1
rgb255((0.03, 0.41, 0.85))    # → (8, 105, 217) ints for python-docx
color_hex((0.03, 0.41, 0.85)) # → "0869D9" uppercase, no '#'
```

## Images

```py
doc.image("photo.png", width=60, height=40)
doc.image("logo.svg", width=30, align="C")  # proportional, centered
```

## Page breaks, bookmarks, links

```py
doc.page_break()
doc.bookmark("section-1")
doc.link("see chapter 1", target="section-1")     # internal jump
doc.link("github", url="https://github.com/...")  # external
```

## Headers / footers

```py
# Static text + automatic page numbering via Word fields
doc.header("Corporate Report", align="C")
doc.footer(text="{page} / {pages}", align="C")
# Or just enable numbering without a template
doc.footer(align="R", page_number=True)
```

## Page sizes

```py
from docmarq import A4, A3, A5, LETTER, LEGAL
DOCX("out.docx", width=LETTER.width, height=LETTER.height)
DOCX("out.docx").landscape() # swap width/height for current section
```

## Metadata

```py
doc.metadata(
  title="Report",
  author="Xaeian",
  subject="Q4",
  keywords="report,q4",
  comments="Internal review",
  category="Finance",
)
```

## Style presets

```py
from docmarq import Styles
Styles.BOLD      # bold=True
Styles.ITALIC    # italic=True
Styles.HEADING1  # font_size=20, bold
Styles.HEADING2  # font_size=16, bold
Styles.HEADING3  # font_size=13, bold
Styles.HEADING4  # font_size=11, bold
Styles.SMALL     # font_size=9
Styles.CAPTION   # font_size=9, italic, muted grey
Styles.CODE      # font_family="Consolas", font_size=10
```

## Markdown

```py
from docmarq.md import md_to_docx, MarkdownStyle, lang_style
from docmarq import A4

md_to_docx(open("doc.md").read(), "doc.docx")
# With localization preset
style = lang_style("pl", footnote_label="Bibliografia")
md_to_docx(text, "doc.docx", style=style)
# Relative image paths resolved against base_dir
md_to_docx(text, "doc.docx", base_dir="./assets")
# Landscape, or any preset, via `page=`
md_to_docx(text, "doc.docx", page=A4.landscape())
```

Supported markdown: headings, paragraphs, lists, tables, code blocks,
blockquotes, GitHub callouts (`> [!NOTE]`), footnotes, mermaid diagrams
(via local `mmdc` or mermaid.ink fallback - shared cache with `pdfmarq`),
images, YAML frontmatter with banner rendering, and math formulas
(`$x^2$`, `$$...$$`) as native Word equations (OMML) with a matplotlib
image fallback for LaTeX outside the supported subset.

## Advanced

```py
# Direct python-docx Document access for anything not covered by the fluent API
doc.doc.add_section()
# Output path (symmetric with pdfmarq.PDF.output_path)
doc.output_path
```
