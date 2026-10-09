# docmarq/paper/__init__.py

"""
Documentation-style DOCX: markdown content with a frontmatter banner around it.

Page 1 gets the banner (id, version, status, dates, logo), later pages a compact header,
every page a page number, and the end optional signing lines.
The content itself is written by `docmarq.md`. Needs ``pip install docmarq[md]``.

Example:
  >>> from docmarq.paper import md_to_docx, lang_style
  >>> md_to_docx(open("doc.md").read(), "doc.docx", style=lang_style("pl"))
"""

from .style import PaperStyle
from .presets import lang_style, LANG_PRESETS
from .render import md_to_docx, md_into

__all__ = [
  "PaperStyle",
  "md_to_docx",
  "md_into",
  "lang_style",
  "LANG_PRESETS",
]
