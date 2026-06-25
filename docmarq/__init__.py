# docmarq/__init__.py

"""
DOCX generation with fluent API. Built on `python-docx`.

Mirror of `pdfmarq` philosophy: thin fluent layer over a heavy backend.
We use `python-docx` for OOXML zip plumbing, content types, and
relationships - everything user-facing is our own API.

Example:
  >>> from docmarq import DOCX, Align
  >>> with DOCX("out.docx") as doc:
  ...   doc.heading("Tytuł", level=1)
  ...   doc.para("Pierwszy akapit.")
  ...   doc.text("Drugi z ").text("bold", bold=True).text(" tekstem.")
"""

#----------------------------------------------------------------------- Metadata for auto-toml

__version__ = "0.2.0"
__repo__ = "Xaeian/docmarq"
__python__ = ">=3.10"
__description__ = "DOCX generation library with fluent API"
__author__ = "Xaeian"
__keywords__ = ["docx", "word", "ooxml", "document", "generation"]
__dependencies__ = ["python-docx"]
# pip-name → import-name mapping. `python-docx` is the pip package but
# the import is `import docx`; the hint surfaces in dependency diagnostics
# so users debugging `No module named 'docx'` see the right install command.
__import_names__ = {"python-docx": "docx"}

#----------------------------------------------------------------------------------- Public API

from .constants import (
  Unit, PageSize, Align, Colors, Defaults,
  A4, A3, A5, LETTER, LEGAL, EMU_PER_MM, EMU_PER_PT,
)
from .styles import Style, TableStyle, Styles
from .layout import PageGeometry
from .structure import Metadata, Bookmark
from .utils import to_mm, mm_to_emu, pt_to_emu, parse_color, parse_margin, color_hex, rgb255
from .inline import RichSegment
from .core import DOCX

__all__ = [
  "DOCX",
  "Unit", "PageSize", "Align", "Colors", "Defaults",
  "A4", "A3", "A5", "LETTER", "LEGAL", "EMU_PER_MM", "EMU_PER_PT",
  "Style", "TableStyle", "Styles",
  "PageGeometry",
  "Metadata", "Bookmark",
  "to_mm", "mm_to_emu", "pt_to_emu", "parse_color", "parse_margin", "color_hex", "rgb255",
  "RichSegment",
]
