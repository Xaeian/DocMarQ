# Legal document guide

How to prepare a markdown file for a letter, statement or contract rendered through [`docmarq.legal`](readme.md).

The file is the document: its text, and in the frontmatter the facts that belong to it.
How it looks - columns, fonts, signing lines - is chosen when it is rendered and never written into the file.

## Quick start

```md
---
title: Statement
place: Gdynia
date: 2026-10-07
from: |
  Jan Kowalski
  81/87 Sea Street, 81-225 Gdynia
---

I declare that the information given in the application is true.
```

Rendered with `sign=["legible signature"]` this is a complete statement: sender and date on top, the title, the text, the signing line on the right.

## Frontmatter

YAML between `---` markers at the top.
Every key is optional; a key that is missing leaves its element out.

| Key | Where it lands | Notes |
| --- | --- | --- |
| `title` | centred under the header rows | the only title; do not repeat it as `#` in the text; write it in normal case, a style with small capitals sets it |
| `place`, `date` | right of the sender's first line | a bare date takes the language's format, other text prints as written |
| `from` | top left | a YAML block or a list, one line each |
| `to` | under the sender, flush with the right margin or the left one | same; the side comes from the style |
| `ref` | under the recipient, in bold after the language's label, `Sygn. akt` in Polish | the case number from the court's letter |
| `logo` | own row at the top, or beside the title with the preamble running beside it | position comes from the style, not the file |
| `attachments` | numbered list under the signing lines | a YAML list |

`title`, `author`, `subject` and `keywords` also fill the document properties Word shows.
Values are plain text: markdown inside them prints as written.

## Typical documents

- **Statement, power of attorney**: `title`, `place`, `date`, `from`, the text, one signing line.
- **Letter, application, complaint, notice**: add `to` and, where there are any, `attachments`; a letter to a court adds `ref`.
- **Contract**: `title`, a preamble naming the parties as ordinary text, then the clauses under `##` headings, two signing lines.
- **Protocol**: like a contract, often with a table in the text.

## Writing the text

Paragraph signs are part of the text, so the author types them: `## §1. Subject of the contract`.
No § is numbered for you: references like "under §5" are kept true by hand, and a statement or letter has no § at all.

- One clause per paragraph or list item.
- Numbered points use `1.` or `1)`: the delimiter you type is kept, and Word counts on from the first number.
- The preamble is everything before the first `##`.
  The style may set it apart in a face, size and spacing of its own.
- A `##` heading is set at body size in the legal style; write `**§1**.` to make only the number bold when the style sets headings in Regular.

## Two columns

With `columns=2` in the style the preamble spans the page and the clauses from the first `##` flow in two columns.
Word ends the columns level on their last page, so a contract that fits one page with its signing lines stays one page.

## Several language versions

Each language version is an ordinary document in its own file, with its own frontmatter.
Rendered side by side with `md_parallel_to_docx`, top-level blocks pair up in order and every pair starts level.

- Keep the structure identical: the same headings at the same levels, the same paragraphs, lists and tables in the same order.
- A version with a block more or less, or a heading where the other has a paragraph, is refused with the number of the first block where they part ways.
- A block taller than the rest of a page breaks across it, both versions together.
- A `<!-- pagebreak -->` stands between blocks, never inside a quote, list or group.
- Signing lines are shared: give them labels in both languages, `Zleceniodawca / Principal`.

## Pre-flight checklist

1. The title is in the frontmatter and not repeated as `#` in the text.
2. Every paragraph sign and every reference to one is typed and matches.
3. Dates are bare `2026-10-07`, unless the wording must stay exactly as written.
4. Language versions have the same block structure, heading for heading.
5. The output has been opened once in Word.
