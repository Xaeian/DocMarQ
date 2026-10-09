# docmarq/legal/__init__.py

"""
Letters, statements and contracts: markdown content with a letter header and a signed end.

Frontmatter holds the facts of the document:
`title`, `place`, `date`, `from`, `to`, `ref`, `logo`, `attachments`.
The style holds the look: columns, justification, signing lines, fonts.
The content itself is written by `docmarq.md`. Needs ``pip install docmarq[md]``.

Example:
  >>> from docmarq.legal import md_to_docx, md_parallel_to_docx, lang_style
  >>> style = lang_style("pl", columns=2, sign="contract")
  >>> md_to_docx(open("contract.md").read(), "contract.docx", style=style)
  >>> md_parallel_to_docx([pl_text, en_text], "contract-pl-en.docx",
  ...   styles=[lang_style("pl"), lang_style("en")])
"""

from .style import LegalStyle
from .presets import lang_style, LANG_PRESETS
from .render import md_to_docx, md_parallel_to_docx

__all__ = [
  "LegalStyle",
  "md_to_docx",
  "md_parallel_to_docx",
  "lang_style",
  "LANG_PRESETS",
]
