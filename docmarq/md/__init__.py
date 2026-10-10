# docmarq/md/__init__.py

"""
Markdown engine for `docmarq`: markdown content written into a `DOCX`, after its last block.

Pages around the content - banner, footer, signing lines - belong to the packages built on it:
`docmarq.paper` for documentation, `docmarq.legal` for letters and contracts.

Install the full dependency bundle with ``pip install docmarq[md]``.

Example:
  >>> from docmarq import DOCX
  >>> from docmarq.md import MarkdownRenderer, lang_style
  >>> doc = DOCX("notes.docx")
  >>> MarkdownRenderer(doc, lang_style("pl")).render(open("notes.md").read())
  >>> doc.save()
"""

#----------------------------------------------------------------------------- Extras for auto-toml

__extras__ = ("md", ["markdown-it-py", "mdit-py-plugins", "PyYAML"])

#--------------------------------------------------------------------------------------- Public API

from .style import MarkdownStyle
from .renderer import MarkdownRenderer, top_blocks
from .frontmatter import (split_frontmatter, peek_frontmatter, metadata_from_frontmatter,
  docx_metadata, read_yaml)
from .presets import lang_style, LANG_PRESETS
from .tokens import is_pagebreak_directive, is_group_open_directive

__all__ = [
  "MarkdownStyle",
  "MarkdownRenderer",
  "top_blocks",
  "split_frontmatter",
  "peek_frontmatter",
  "read_yaml",
  "metadata_from_frontmatter",
  "docx_metadata",
  "lang_style",
  "LANG_PRESETS",
  "is_pagebreak_directive",
  "is_group_open_directive",
]
