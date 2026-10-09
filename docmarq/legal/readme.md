# `docmarq.legal`

Letters, statements and contracts: markdown content with a letter header and a signed end.
Requires `pip install docmarq[md]`.
The content is written by [`docmarq.md`](../md/readme.md).
How to write the source file is in [`guide.md`](guide.md).

## `md_to_docx`

```py
from docmarq.legal import md_to_docx, lang_style
style = lang_style("pl", sign=["Jan Kowalski"])
md_to_docx(open("letter.md").read(), "letter.docx", style=style, base_dir=".")
```

The arguments match `docmarq.paper.md_to_docx`: `style`, `page`, `margin`, `gutter`, `base_dir`, `font_dir`, `metadata`.
`style` must be a `LegalStyle`, built directly or with `lang_style()`.

The library lays out what the author wrote.
It numbers no paragraphs, detects no language, translates nothing and writes no legal wording of its own.

- **Markdown body**: the text - paragraphs, headings, lists, tables.
  Paragraph signs are typed by the author.
- **Frontmatter**: the facts of this document - `title`, `place`, `date`, `from`, `to`, `ref`, `logo`, `attachments`.
- **`LegalStyle`**: the look - columns, justification, signing lines, fonts, sizes, labels.

Word lays the pages out when the document opens: it breaks pages, balances columns and numbers pages.
The same style and source give the same document as `pdfmarq.legal`, set by Word instead of measured in advance.

## Header

Rows from frontmatter, each written only when its keys are set:

1. `logo`, aligned by `logo_align` (`L`, `C` or `R`)
2. `from` on the left, `place, date` on the right of its first line
3. `to` flush with the right margin, or the left one with `to_align="L"`, and under it `ref`, the case number, in bold after `ref_label` (`Sygn. akt` in the pl preset)
4. `title`, centred

A logo at the side shares the title's row when nothing stands between them, as on a contract.
It floats at the margin, and the preamble wraps around it.
The clauses start below both.
A title with no room beside the logo for its longest word gets the logo in a row of its own instead.
Values are plain text: a YAML block (`|`) or a list gives one line per entry, joined by commas in `place` and `ref`, which keep to one line.
A mapping is refused.
`date` takes the preset's `date_format` when it is a YAML date or ISO `YYYY-MM-DD` text, other text prints as written.

## Preamble

The blocks before the first `##` are the preamble, which names the parties of a contract.
Without any `##` there is none: the whole text is clauses.
The preamble is set in `header_font` at `preamble_size`, its paragraphs `preamble_para_gap` apart, and the clauses start `preamble_gap` below it.
Left at `None`, these fields follow the text, so a plain letter looks the same throughout.

## `LegalStyle`

`LegalStyle` extends `MarkdownStyle`, with defaults set for legal text:

- `justify=True`: justified text, see [Flowing text](../md/readme.md#flowing-text)
- `h1_size=13`, `h2_size`…`h6_size=None` (body size), no rules under `#` and `##`, tighter gaps

Its own fields:

```py
LegalStyle(
  columns=1, column_gap=8,          # 2 sets the text in two columns
  preamble_size=None,               # pt, None follows body_size
  preamble_para_gap=None,           # mm, None follows para_gap
  preamble_gap=4,                   # mm between the preamble and the clauses
  header_font=None,                 # face of header, title, preamble, signing lines, attachments, page number
  title_size=16, title_caps=False,  # pt; small capitals from the lowercase letters
  title_gap=8,                      # mm below the title, and above it under header rows
  header_size=10,                   # pt
  to_align="R",                     # recipient flush right, "L" flush left
  ref_label="Ref.",                 # before the bold `ref`; "Sygn. akt" in the pl preset
  logo_align="L", logo_max_h=20, logo_max_w=50,
  date_format="%Y-%m-%d",           # "%d.%m.%Y r." in the pl preset
  sign=False,                       # True, a sign_labels scenario, or a list of labels
  sign_gap=20, sign_w=70, sign_size=9,
  sign_bottom=False,                # signing lines and attachments at the bottom margin
  attachments_label="Attachments",
  page_number_label="Page", page_number_total=True,
)
```

Signing lines are dotted, as on a form, in the text colour, their labels in italics.
The text is set in `body_color`, and Word keeps two lines of a paragraph on each side of a page break.
`title_caps` sets the lowercase letters in Word's small capitals.
Write `title: Umowa zlecenie`, not in capitals.
A document of one page never carries a page number.
`lang_style(lang, **overrides)` layers a language preset (`en`, `pl`, `de`, `fr`, `es`, `it`, `cs`, `sk`) under the overrides.
A preset also sets `hyphenate` to its language, so Word splits words at line ends.
`LegalStyle()` alone keeps them whole.

## Columns

With `columns=2` the preamble spans the page, and from the first `##` on the text flows in columns.
Without any `##` the whole text flows in columns.
The signing lines and attachments span the page again below them.

Word ends the columns level on their last page, so a contract that fits one page with its end stays one page.
A longer one flows over the pages, its last page balanced.

## Parallel versions

```py
from docmarq.legal import md_parallel_to_docx, lang_style
md_parallel_to_docx([pl_text, en_text], "contract.docx",
  styles=[lang_style("pl"), lang_style("en")])
```

Versions of one document side by side, one column each.
Each is an ordinary legal document with its own frontmatter and its own style, so its own labels, date format and language.
Which language each one is in is the caller's knowledge.

- A borderless table holds them, a column per version and a row per block, so every row starts level.
- Headers fill the first row, the text starts below the lowest of them.
- A heading row keeps the row after it company, and a row taller than the rest of the page breaks across it.
- `<!-- pagebreak -->` starts a new page for all columns, and a `<!-- group -->` is one row.
- Each version's text is marked in its language, so Word hyphenates and spell-checks it in that one.
- Signing lines come from `styles[0]` across the page; bilingual labels go in its `sign` list, for example `["Zleceniodawca / Principal", "Zleceniobiorca / Contractor"]`.
- Each version's attachments stay in its column, below the signing lines.
- A logo keeps a row of its own above the title, and the preamble is set like the rest of the text.

These raise `ValueError` rather than write a wrong document:

- versions whose blocks differ in count or kind (`##` is not `###`), naming the first block where they part ways,
- a `<!-- pagebreak -->` inside a group, a quote or a list,
- fewer than two versions, or styles that do not match them in number.

## In Word

What each part becomes, for a document edited further in Word:

- **Columns**: a continuous section with two columns, between section breaks that start no page.
- **Page number**: the field `IF NUMPAGES > 1` holding `PAGE` and `NUMPAGES` in the footer.
- **Logo beside the title**: a picture floating at the margin, text wrapping square around it; a break that clears it ends the preamble.
- **`sign_bottom`**: a borderless table floating at the bottom margin of the page the text ends on.
- **Signing lines**: dotted bottom borders of paragraphs, kept with what follows them on one page.
- **Hyphenation**: Word's automatic hyphenation, with the text in the preset's language; headings, the header and table cells keep their words whole.

## Two-column contract

A one-page contract set like a typical template: the clauses in Barlow Semi Condensed at 10pt, the page around them in Barlow, the title in small capitals with the logo beside it.
Word sets text in fonts installed where the document opens, so Barlow has to be installed there.

```py
from docmarq import parse_color
from docmarq.legal import md_to_docx, lang_style
dark = parse_color("#231f20")
style = lang_style("pl",
  font_body="Barlow Semi Condensed", font_head="Barlow Semi Condensed", head_mode="Regular",
  header_font="Barlow", preamble_size=11, preamble_para_gap=1.41, preamble_gap=5.48,
  body_size=10, para_gap=0.7, head_gap_top=1.41, head_gap_bot=1.41,
  title_size=24, title_caps=True, title_gap=3.5, header_size=11,
  logo_align="R", logo_max_h=37.45, logo_max_w=23, columns=2, column_gap=9.3,
  body_color=dark, head_color=dark, hr_thick=0.5,
  sign=["podpis Zleceniodawcy", "podpis Zleceniobiorcy"], sign_w=90, sign_size=11,
  sign_gap=14, sign_bottom=True,
)
md_to_docx(text, "contract.docx", style=style, margin=(9.5, 10.35, 3.69, 10.35))
```

With `head_mode="Regular"`, a heading written `## **§1**. Subject` sets only the number in bold.
The small bottom margin lets the signing lines sit low, where the template has them.
