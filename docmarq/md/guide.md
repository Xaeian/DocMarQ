# Markdown document guide

How to prepare a markdown file that renders cleanly to PDF through [`pdfmarq`](https://github.com/Xaeian/PDFMarQ) and DOCX through [`docmarq`](https://github.com/Xaeian/DocMarQ).
Both read the same source, so one file serves both.

The document says what it is; how it looks is decided per conversion and never written into the file.
That is why there are no keys here for page size, margins or fonts.

## Quick start

```md
---
title: Quarterly report
author: Emilian Świtalski
---

# Summary

Revenue is up 23% year-over-year.
```

Everything else has sensible defaults.

## Frontmatter

YAML between `---` markers at top of the file.
Rendered as a paper document, it becomes the banner on page 1.

```yaml
---
id: RM-001           # unique document code
title: Markdown document guide
version: v2.3.1
author: Emilian Świtalski
status: approved     # draft/review/approved/deprecated/archived
entity: Gdynia Maritime University
address: Morska 81/87, 81-225 Gdynia
created: 2024-09-15  # first written, set once
updated: 2026-03-22  # last content change
logo: umg.svg        # .svg/.png/.jpg
subject: Short description for /Subject metadata
keywords: policy, workflow, reference
---
```

Every field is optional and unknown keys pass through untouched, so a document can carry whatever else your tooling needs.

`title` and `author` also fill `/Title` and `/Author`.
`subject` and `keywords` never show in the banner; they fill `/Subject` and `/Keywords`, searchable by DMS systems.
`keywords` takes a comma-separated string or a YAML list.
Write dates bare: a quoted date is carried through as text and the two formats then disagree on how to print it.

A first heading repeating `title` word for word is dropped so title is not printed twice, and its anchor goes with it, so do not link to it.

The language of everything the renderer writes itself, callout labels, the word for "page", the date format, is chosen at conversion time rather than in the file.

### Statuses

| Status       | Colour    | Meaning                                                       |
| ------------ | --------- | ------------------------------------------------------------- |
| `draft`      | blue-grey | Work in progress. Content is incomplete or unreviewed.        |
| `review`     | amber     | Content is complete and waiting for approval.                 |
| `approved`   | green     | Officially accepted. This is the current, binding version.    |
| `deprecated` | red       | Still accessible but no longer recommended. Being phased out. |
| `archived`   | violet    | Historical record only. Not valid for current use.            |

Status is an administrative decision, so changing it bumps neither `version` nor `updated`.
A value outside the table is not rejected: it prints as a grey badge in capitals, so a typo shows up as a colourless label rather than an error.

## Content

Standard GitHub-flavored markdown works in both.
For syntax itself, [Markdown](https://github.com/Xaeian/Markdown) is the reference; this guide covers only what these two renderers add, and where they part company.

### Where the two formats differ

PDF is the richer target.
If a document may end up as DOCX, do not lean on anything in the middle column.

| Written as                    | PDF                                 | DOCX                                   |
| ----------------------------- | ----------------------------------- | -------------------------------------- |
| fenced code with a language   | coloured by Pygments                | plain, hint is dropped                 |
| `$x^2$` and `$$…$$`           | drawn, blocks numbered `(1)`, `(2)` | a real Word equation, unnumbered       |
| a formula inside a table cell | drawn                               | drops back to `$…$` as text            |
| formatting in a table cell    | bold, links and images all work     | text survives, an image leaves its alt |
| footnote ref `[^1]`           | jumps to definition                 | superscript only, not clickable        |
| `- [x]` task list             | checkbox                            | literal `[x]`                          |
| `==mark==`, `^sup^`, `~sub~`  | rendered                            | literal characters                     |
| `:rocket:`                    | emoji                               | literal `:rocket:`                     |
| definition list               | rendered                            | one paragraph, literal `:`             |
| `<b> <i> <code> …`            | rendered                            | dropped, text survives                 |

A formula itself is safe in both.
What is not safe is a table cell mixing text with anything beyond plain text: in DOCX a cell is a string, so formatting flattens and a picture among the text is reduced to `[alt]`.
A picture alone in a cell is drawn in both formats.
PDF renders the HTML tags `<b>/<strong>`, `<i>/<em>`, `<code>`, `<br>` and `<hr>`.
A tag carrying attributes is stripped, so `<b class="x">` is not one of them.
All other HTML goes in both formats, and DOCX drops even those, so prefer `---` over `<hr>`.

### Lists

Under a numbered item a nested list needs three or four spaces; two makes it a sibling instead, in both formats.
In PDF that still looks about right, because the marker you typed is the marker you get, which is what makes the mistake easy to miss.
Under a bullet, two spaces are fine.

A nested list that starts at anything other than `1.` needs a blank line above it, or the parent item swallows it and text runs together on one line.
With that blank line any starting number is read as a list.
PDF then prints numbers as written, while Word counts on from the first one through its own list style.

### Images

Keep files next to the markdown and link them relatively.
Everything the document needs has to travel with it: an absolute path, a path leading out of the document's own folder, and an `https://` address all fail to load.

Give a figure its own paragraph and let it size itself.
It fills text width as far as its resolution allows, stops at 120mm tall and sits centred, so an exported chart spans the column while a small icon stays an icon.
PDF and DOCX land on the same box.

Images alone in a paragraph share that width: two side by side become two half-width figures, eight become two rows of four.
An image on the same line as text stays inline at line height, in both formats - so put caption in the next paragraph, never beside the picture.

The title slot sizes a figure by hand when automatic size is wrong:

```md
![diagram](schema.svg "max_h=80 align=C")
```

| Key     | Effect                                           |
| ------- | ------------------------------------------------ |
| `max_w` | Cap width in mm, aspect ratio preserved          |
| `max_h` | Cap height in mm, aspect ratio preserved         |
| `w`     | Set width in mm                                  |
| `h`     | Set height in mm                                 |
| `scale` | Multiplier on natural size, `1.0` means 100%     |
| `align` | `L`, `C` or `R`, block images default to `C`     |

`scale` runs first from natural size, then `w` and `h`, then `max_*` clamps the result, so a cap always wins.

A missing file never stops the conversion: PDF prints `[Image not found: …]`, DOCX leaves alt text, and the document comes out looking finished with a hole in it.

### Formulas

Inline `$x^2$` and block `$$...$$` both draw as real maths.
The LaTeX spellings `\(x^2\)` and `\[...\]` are read as the same thing, so a document pasted from a LaTeX source or an AI assistant needs no conversion.

PDF numbers display formulas `(1)`, `(2)` down the right margin.
A `\tag{X1.1}` names the label itself: it replaces the number in PDF and appears right-aligned next to the equation in DOCX.

### Mermaid diagrams

Fenced `mermaid` blocks compile to PNG and embed like any other figure.
The same image DSL applies, written as an info string after lang token.

````md
```mermaid max_h=80 align=C
flowchart TB
  A --> B
```
````

When the renderer is unavailable the source stays a plain fenced block, so document still builds.

### Callouts

```md
> [!NOTE]
> Plain information block.
```

Five types: `NOTE`, `TIP`, `IMPORTANT`, `WARNING`, `CAUTION`.
Each gets a coloured left border and a bold label in the document language, so Polish gives `Notatka`, `Wskazówka`, `Ważne`, `Ostrzeżenie` and `Uwaga`.

### Internal links

Each heading registers a GitHub-style slug: lowercase, spaces to hyphens, unicode kept, runs collapsed to one.
Two headings with the same text get numbered slugs, `setup` and then `setup-1`, as on GitHub.

A link to a slug that does not exist stays readable rather than dangling: plain text in DOCX, styled but inert in PDF.
The dropped duplicate title counts as one of those, so a link to it is text, not a jump.

Footnote refs `[^1]` jump to their definition in PDF.
DOCX prints superscript but registers no anchor, and definitions collect at the end of the document rather than becoming Word footnotes.

## Layout directives

HTML comments that control page flow, inert in ordinary markdown renderers.

```md
<!-- pagebreak -->

<!-- group -->
A block that must stay together.
<!-- /group -->
```

`pagebreak` forces a break at that point.
`group` keeps its content on one page, breaking pre-emptively when the block would fit on the next; a group larger than a page flows normally.
Both must stand alone in the comment, so `<!-- pagebreak now -->` is an ordinary comment and does nothing.

## Templates

A document with placeholders is a template: one record fills it into one document, a table of records into a series, as the same letter to many addresses.
`pdfmarq.merge` and `docmarq.merge` fill it with data read from CSV, JSON or anything else by the script that calls them.
A document that is never filled prints its braces as written.

```md
---
title: Service contract
date: "{{ signed }}"
to: ["{{ name }} {{ surname }}", "{{ city }}"]
vars:
  rate: "30.50"   # default, a record can change it
  id_card:        # left empty: every record must fill it
---

Concluded with **{{ name }} {{ surname }}**, identity card {{ id_card }}.

{% if student == "yes" %}
The contractor is a student.
{% endif %}

| Task | Hours |
|---|--:|
{% for task in tasks %}
| {{ task.name }} | {{ task.hours }} |
{% endfor %}
```

- A placeholder in frontmatter goes in quotes, as `date: "{{ signed }}"`.
  Unquoted, YAML reads the braces as a mapping, and the template stops with that hint.
- `vars` holds the variables of the file itself: defaults a record can change, and empty entries every record must fill.
  An empty cell takes the default.
  A placeholder inside a `vars` value prints as written, since `vars` hold values.
- A line holding a tag vanishes when it prints nothing and stays a line of its own otherwise.
  An optional table row or list item is one line, as `{% if note %}| {{ note }} |{% endif %}`, and keeps its indent.
- A value prints as written, `*`, `|` and `<` included, without the spaces around it.
  `{{ value | md }}` inserts markdown from the data, as `safe` does, and is the form inside code and math, where a backslash would show.
- Values join in the text, as `{{ name }} {{ surname }}`: beside a `{% set %}` block or an `md` value, `~`, `join` and `format` escape them the wrong way.
- A value of several lines keeps its lines, which inside a table cell breaks the row: join it there, as `{{ address | replace("\n", ", ") }}`.
- Any text is true in `{% if %}`, `no`, `0` and Excel's `FALSE` included.
  Compare instead, as `{% if student == "yes" %}`, or leave the cell empty.
- `{{ value | default("-", true) }}` covers an empty cell; `default("-")` alone covers only a field the record lacks.
- `{{ amount | number }}` and `{{ signed | date }}` print in the language and date format chosen at conversion, as `1 234,50` and `12.01.2026 r.` for Polish.
- Literal `{{` or `{%`, as in LaTeX or a mermaid hexagon, goes inside `{% raw %}` and `{% endraw %}`.
- Clauses kept in files of their own join with `{% include "clauses/" ~ clause ~ ".md" %}`, from the folder of the template.
  A blank line before `{% endfor %}` keeps one clause from running into the next.

When the script gives the whole table as one field, as `people`, the template makes one document with a section per record:

```md
{% for person in people %}
## {{ person.name }} {{ person.surname }}

Completed the course on {{ person.date }}.
{% if not loop.last %}

<!-- pagebreak -->
{% endif %}

{% endfor %}
```

The frontmatter, the page numbers and the signing lines then belong to the whole document, not to each record.

Data from Excel, for whatever reads it:
- save it as "CSV UTF-8";
- format columns with leading zeros, as document numbers, as text, or Excel drops the zeros;
- format dates as `YYYY-MM-DD`;
- name columns as variables, `first_name` rather than `First name`.

## Pre-flight checklist

1. `title:` and `author:` are set, and `status:` is honest.
2. `id:` is set if the document belongs to a controlled set, such as a DMS or an audit trail.
3. `created:` untouched, `updated:` current.
4. Every image and the logo sit beside the source and travel with it.
   Nothing points outside the document's own folder.
5. The output has been opened once.
   A file that would not load leaves a placeholder, not an error.
6. If the document may also become DOCX, nothing from the PDF column is load-bearing.
