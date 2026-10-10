# `docmarq`

Fluent DOCX generation, built on `python-docx` with a paragraph/run flow model.
Sibling library: [`pdfmarq`](https://github.com/Xaeian/PDFMarQ), same API shape for PDF.

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
# Word's heading styles: size in pt, face, the rule under them; None keeps what is set
doc.heading_style(2, size=13, bold=True, italic=False, rule=False)
# Automatic hyphenation in a language, which becomes the language of the document
doc.hyphenation("pl")
```

Headings, the cells of `table()` and code blocks written after `hyphenation()` keep their words whole, so it comes before the blocks.

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
doc.text("Umowa", small_caps=True)       # lowercase letters as small capitals
doc.border("#1f2328", sides=("top", "bottom"), space=8) # rule the current paragraph
```

## Headings, lists

```py
doc.heading("Chapter 1", level=1)
doc.heading("Terms", level=2, space_before=12, space_after=4) # pt, instead of the level's gaps
doc.bullet("First item")
doc.bullet("Nested item", level=1)
doc.ordered("Step one", start=1, delim=")") # a new list counting from 1, printing `1)`
doc.ordered("Step two")                      # continues it
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

## Columns

```py
doc.columns(2, gap=8) # what follows in two columns, gap in mm
doc.para("Flows down the first column, then the second.")
doc.columns(1)        # back to one, below the columns
```

Each change starts a continuous section, so the text runs on down the same page.
Word ends the columns of a section level when it closes this way, on its last page too.
`doc.content_width` follows the column, so tables and images fit it.

## Writing into a table cell

```py
tbl = doc.new_table(1, 2)               # a bare table at the end of the write target
with doc.into(tbl.rows[0].cells[0], 80): # cell width in mm
  doc.para("Written in the cell.")
  doc.table([["a", "b"]])               # sized to the cell
doc.para("Back after the table.")
```

Every block call writes into the cell while the `with` lasts, and the cell's width stands in for `content_width`.

## Blocks of your own

The markdown engine and the document packages write their blocks through these calls too.

```py
from docmarq.tables import set_table_widths
from docmarq.images import preprocess_to_buffer, pin_offsets
doc.enter()                      # close the open paragraph
p = doc.new_para()               # bare: no gaps, and not the one text() writes to
doc.apply_para_spacing(p)        # the style's leading and gaps, the gap above collapsed
doc.keep_words(p)                # no hyphens in it under hyphenation()
p.add_run("A block of my own")
doc.track_block_spacing(6)       # its gap below in pt, for the next block to collapse into
tbl = doc.new_table(1, 2)        # a bare table, a hairline apart from a table before it
set_table_widths(tbl, [60, 100]) # fixed column widths in mm
run = tbl.rows[0].cells[0].paragraphs[0].add_run()
run.add_picture(preprocess_to_buffer("logo.png")) # a PNG Word takes, transparent margins cropped
pin_offsets(run)                 # the picture flush with the cell, in LibreOffice too
doc.chain(lambda: doc.para("Kept with the next one.")) # what it writes stays on one page
doc.hairline()                   # a 1pt paragraph: ends a section, keeps two tables apart
doc.paragraph, doc.text_style    # the paragraph text() writes to, and the style it writes in
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

`docmarq.fields` writes the fields themselves: `page_fields(p, "{page} / {pages}")` into any paragraph, and `page_fields_if_many(p, template)` as the field `IF NUMPAGES > 1`, which shows the numbers only in a document of more than one page.

## Signing lines and page number

`docmarq.parts` builds the pieces both document packages share, from plain values, the face of their text as a `RichSegment`:

```py
from docmarq import RichSegment
from docmarq.parts import signatures, page_number
label = RichSegment("", size=9, color=(0.4, 0.4, 0.4), italic=True)
signatures(doc, ["Client", "Contractor"], label, gap=20, width=70, thick=0.5, line="dotted")
page_number(doc.doc.sections[0].footer.paragraphs[0], "Page", show_total=True, many=True)
```

Their labels and the date format in each language are in `docmarq.parts.LANG_LABELS`, which the presets of both packages build on.

## Page sizes

```py
from docmarq import A4, A3, A5, LETTER, LEGAL
DOCX("out.docx", width=LETTER.width, height=LETTER.height)
DOCX("out.docx").landscape() # swap width/height for every section
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
from docmarq.paper import md_to_docx, lang_style
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

The markdown engine is [`docmarq.md`](md/readme.md).
Documentation with a banner comes from [`docmarq.paper`](paper/readme.md), letters and contracts from [`docmarq.legal`](legal/readme.md).
[`docmarq.merge`](merge/readme.md) fills a markdown template with one record, or one per record of a table.

## Advanced

```py
# Direct python-docx Document access for anything not covered by the fluent API
doc.doc.add_section()
# Output path (symmetric with pdfmarq.PDF.output_path)
doc.output_path
```
