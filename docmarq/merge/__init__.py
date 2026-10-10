# docmarq/merge/__init__.py

"""
Mail merge: a markdown template and one record of data give a plain markdown document.

Placeholders `{{ field }}` and the tags `{% if %}`, `{% for %}`, `{% include %}` are Jinja2,
run in its sandbox, and values print as written.
`fill` takes one record, a mapping.
`merge` takes a list of them and checks all before it returns any.
The caller reads the data, from CSV, JSON or anything else.
The filled text goes to a document package, as `docmarq.legal`. Needs ``pip install docmarq[md]``.

Example:
  >>> import csv
  >>> from docmarq.legal import md_to_pdf, lang_style
  >>> from docmarq.merge import fill, merge
  >>> md_to_pdf(fill(template, {"name": "Anna", "surname": "Nowak"}), "one.pdf")
  >>> rows = list(csv.DictReader(open("people.csv", encoding="utf-8-sig"), delimiter=";"))
  >>> for doc in merge(template, rows, name="{{ surname }}"):
  ...   md_to_pdf(doc.text, f"out/{doc.name}.pdf", style=lang_style("pl"))
"""

__extras__ = ("md", ["Jinja2>=3.1.6"])

from .template import fill, merge, Merged, MergeError

__all__ = [
  "fill",
  "merge",
  "Merged",
  "MergeError",
]
