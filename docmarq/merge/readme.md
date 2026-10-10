# `docmarq.merge`

Mail merge: a markdown template filled with one record gives a document, filled with a list of records a document per record.
Requires `pip install docmarq[md]`.
How to write a template is in the [markdown guide](../md/guide.md#templates).

The filled text is plain markdown, so it goes to any document package: `docmarq.paper`, `docmarq.legal`, a cover through `md_into`, or language versions side by side.

## `fill`: one record

```py
from docmarq.legal import md_to_docx, md_parallel_to_docx, lang_style
from docmarq.merge import fill

record = {"name": "Jan", "surname": "Kowalski", "signed": "2026-01-12"}
md_to_docx(fill(template, record), "contract.docx", style=lang_style("pl"))
md_to_docx(fill(notice), "notice.docx") # a template whose own `vars` fill it, no record
pl, en = fill([template_pl, template_en], record) # language versions, in the order given
md_parallel_to_docx([pl, en], "contract-pl-en.docx", styles=[lang_style("pl"), lang_style("en")])
md_to_docx(fill(certificates, {"people": rows}), "certificates.docx") # a table as one document
```

`fill` returns the markdown, or a list of it for a list of templates.
A table given as one field makes one document, the template looping over it with a page break between records.
Its rows then skip the checks of `merge`: empty rows, cells past the header and a byte-order mark.

## `merge`: a list of records

```py
import csv
from pathlib import Path
from docmarq.legal import md_to_docx, lang_style
from docmarq.merge import merge

HERE = Path(__file__).parent
STYLE = lang_style("pl")
template = (HERE / "contract.md").read_text(encoding="utf-8")
with open(HERE / "people.csv", encoding="utf-8-sig", newline="") as file:
  rows = list(csv.DictReader(file, delimiter=";"))
for doc in merge(template, rows, name="{{ surname }}-{{ name }}",
    base_dir=HERE, date_format=STYLE.date_format, lang="pl"):
  md_to_docx(doc.text, str(HERE / "out" / f"{doc.name}.docx"), style=STYLE, base_dir=str(HERE))
```

Each result is `Merged(name, text, data)`:
- `name`: the file name from the `name` pattern, which sees the fields and `n`, the record number, or the zero-padded number without a pattern;
- `text`: the filled markdown, a list of them when `template` is a list of language versions;
- `data`: the record, for anything the loop sets per record, as `replace(STYLE, sign=[...])`.

`base_dir` is the root of `{% include %}`, `date_format` the format of `date`, and `lang` the separators of `number`, in `fill` as in `merge`.

## Records

A record is a mapping, field name to value: one goes to `fill`, a list of them to `merge`.
Reading a file is the caller's: a JSON object or a YAML mapping is one record, and `csv.DictReader`, `xaeian.CSV.load` or a JSON array give a list of them as it is.
A CSV reader settles the separator, the encoding and quoting; the values are taken as they come, text or typed.
A list given to `fill`, or one mapping to `merge`, raises `TypeError` naming the other.

## Checks

`merge` checks every record before it returns any, and `fill` its one; a `MergeError` lists every problem:
- template syntax, with the line of the file, or the name and line of an included file;
- a frontmatter placeholder without quotes;
- records that are no mappings, are empty, or hold more cells than the header has names;
- field names that are no variable names, or that Jinja keeps for itself, as `true`, `none`, `loop`, `range` and `n`;
- a byte-order mark before the first field name, from a CSV file read without `utf-8-sig`;
- a `vars` entry left empty that a record does not fill;
- a field the template reads and the record lacks, with the line of the template or of the included file;
- a problem in one of several language versions, with the number of the version;
- file names that come out empty or repeat, letter case aside.

```text
2 of 120 records cannot be filled:
  record 7 (Kowalski-Jan): `id_card` is required and empty
  record 12 (Nowak-Anna): line 14: 'rate_gross' is undefined
```

A field no template reads gives a `RuntimeWarning`, since it is usually a typo.

File names lose the characters Windows refuses and a dot or space at the end, keep 100 characters at most, and `CON` or `CON.txt` gets a `_` before the first dot.

## Sandbox

Templates run in the immutable sandbox of Jinja2: a template cannot reach Python's internals, files or the network, and cannot change the data, so one record never leaks into the next.
- `include` reads only below `base_dir`, and fails without it.
- `**` takes exponents up to 100, and `*` builds at most 100000 characters or items.
- `random` and `lipsum` are gone, so one record always gives one document.

Other ways to spend time or memory remain, as nested loops or `center`, so a service filling templates from strangers runs them in a process it can stop.
