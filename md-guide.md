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

YAML between `---` markers at the top of the file. It becomes the banner on page 1.

```yaml
---
id: RM-001           # unique document code
title: Markdown document guide
version: v2.2.1
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

A first heading repeating `title` word for word is dropped so the title is not printed twice, and its anchor goes with it, so do not link to it.

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
For the syntax itself, [Markdown](https://github.com/Xaeian/Markdown) is the reference; this guide covers only what these two renderers add, and where they part company.

### Where the two formats differ

PDF is the richer target.
If a document may end up as DOCX, do not lean on anything in the middle column.

| Written as | PDF | DOCX |
| --- | --- | --- |
| fenced code with a language | coloured by Pygments | plain, the hint is dropped |
| `$x^2$` and `$$…$$` | drawn, blocks numbered `(1)`, `(2)` | a real Word equation, unnumbered |
| a formula inside a table cell | drawn | drops back to `$…$` as text |
| formatting in a table cell | bold, links and images all work | text survives, an image leaves its alt |
| footnote ref `[^1]` | jumps to the definition | superscript only, not clickable |
| `- [x]` task list | checkbox | literal `[x]` |
| `==mark==`, `^sup^`, `~sub~` | rendered | literal characters |
| `:rocket:` | emoji | literal `:rocket:` |
| definition list | rendered | one paragraph, literal `:` |
| `<b>/<strong> <i>/<em> <code> <br> <hr>` | rendered | dropped, the text survives |

A formula itself is safe in both. What is not safe is a table cell carrying anything beyond plain text: in DOCX a cell is a string, so formatting flattens and a picture is reduced to its alt text.
A tag carrying attributes is stripped too, so `<b class="x">` is not the whitelisted `<b>`.
All other HTML goes in both formats, so prefer `---` over `<hr>`.

### Lists

Under a numbered item a nested list needs three or four spaces; two makes it a sibling instead, in both formats.
In PDF that still looks about right, because the marker you typed is the marker you get, which is what makes the mistake easy to miss.
Under a bullet, two spaces are fine.

A nested list that starts at anything other than `1.` needs a blank line above it, or the parent item swallows it and the text runs together on one line.
With that blank line any starting number is read as a list.
PDF then prints the numbers as written, while Word counts from one through its own list style.

### Images

Relative paths resolve against the markdown file's directory, and an `https://` address never loads.

```md
![diagram](schema.svg "max_w=120 max_h=80 scale=0.8 align=C")
```

| Key      | Effect                                                        |
| -------- | ------------------------------------------------------------- |
| `max_w`  | Cap width in mm, aspect ratio preserved                       |
| `max_h`  | Cap height in mm, aspect ratio preserved                      |
| `w`      | Set width in mm                                               |
| `h`      | Set height in mm                                              |
| `scale`  | Multiplier on the natural size, `1.0` means 100%              |
| `align`  | `L`, `C` or `R`, block images default to `C`                  |

Sizing resolves in one order: `scale` from the natural size, then `w` and `h`, then `max_*` clamps whatever came out.
So a cap is never overridden, an explicit `w` wider than `max_w` still lands on `max_w`, and `scale` replaces `w`/`h` rather than joining them.

With no DSL at all the two formats disagree, so set `w`, `h` or `scale` whenever the size matters.
A cap alone will not equalise them: PDF clamps the natural size, so a cap larger than the image changes nothing, while DOCX clamps the full text width, so the same cap scales the image up to it.
PDF keeps a raster at natural size and never enlarges it, while an SVG spreads across the full text width.
DOCX fits every block image to the text width, upwards as well, then clamps to the same 120mm height, so only wide images reach the full width and a small PNG still arrives blown up.

A missing image never stops the conversion: PDF prints `[Image not found: …]`, DOCX drops in the alt text, and the document comes out looking finished with a hole in it.
In DOCX an image also has to stand alone in its paragraph, or it turns into italic alt text.

### Mermaid diagrams

Fenced `mermaid` blocks compile to PNG and embed like any other figure.
The same image DSL applies, written as an info string after the lang token.

````md
```mermaid max_h=80 align=C
flowchart TB
  A --> B
```
````

When the renderer is unavailable the source stays a plain fenced block, so the document still builds.

### Callouts

```md
> [!NOTE]
> Plain information block.
```

Five types: `NOTE`, `TIP`, `IMPORTANT`, `WARNING`, `CAUTION`.
Each gets a coloured left border and a bold label in the document language, so Polish gives `Notatka`, `Wskazówka`, `Ważne`, `Ostrzeżenie` and `Uwaga`.

### Internal links

Each heading registers a GitHub-style slug: lowercase, spaces to hyphens, unicode kept, runs collapsed to one.
Two headings with the same text get the same slug and nothing tells them apart, so linking to either is a coin flip.

A link to a slug that does not exist stays readable rather than dangling: plain text in DOCX, styled but inert in PDF.
The dropped duplicate title counts as one of those, so a link to it is text, not a jump.

Footnote refs `[^1]` jump to their definition in PDF.
DOCX prints the superscript but registers no anchor, and the definitions collect at the end of the document rather than becoming Word footnotes.

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

## Pre-flight checklist

1. `title:` and `author:` are set, and `status:` is honest.
2. `id:` is set if the document belongs to a controlled set, such as a DMS or an audit trail.
3. `created:` untouched, `updated:` current.
4. Every image and the logo resolve to a file that travels with the source. Check the output, because a missing one is never reported.
5. If the document may also become DOCX, nothing from the PDF column is load-bearing.
