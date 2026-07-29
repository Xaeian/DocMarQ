# docmarq/md/__init__.py

"""Markdown-to-DOCX rendering for `docmarq`. Mirrors the `pdfmarq.md` API:
`md_to_docx` for one-shot conversion, `MarkdownRenderer` for embedding into
a larger document. Requires the `md` extra."""

#------------------------------------------------------------------------- Extras for auto-toml

__extras__ = ("md", ["markdown-it-py", "mdit-py-plugins", "PyYAML", "svglib", "rlPyCairo"])

#----------------------------------------------------------------------------------- Public API

from .style import MarkdownStyle
from .renderer import MarkdownRenderer, md_to_docx
from .presets import lang_style, LANG_PRESETS

__all__ = [
  "MarkdownStyle",
  "MarkdownRenderer",
  "md_to_docx",
  "lang_style",
  "LANG_PRESETS",
]
