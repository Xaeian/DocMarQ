# docmarq/constants.py

"""Constants for DOCX library - units, page sizes, alignment, defaults."""
from dataclasses import dataclass

# OOXML uses EMU (English Metric Units): 1 inch = 914400 EMU = 25.4 mm
EMU_PER_MM = 36000
EMU_PER_PT = 12700 # 1 pt = 1/72 inch

#---------------------------------------------------------------------------------------- Units

class Unit:
  """Unit conversion factors to millimeters."""
  MM = 1.0
  CM = 10.0
  INCH = 25.4
  PT = 25.4 / 72
  PX = 25.4 / 96

#------------------------------------------------------------------------------------- PageSize

@dataclass
class PageSize:
  """Common page sizes in mm."""
  width: float
  height: float
  def landscape(self) -> "PageSize":
    """Return a copy with width/height swapped (landscape orientation)."""
    return PageSize(self.height, self.width)

A4 = PageSize(210, 297)
A3 = PageSize(297, 420)
A5 = PageSize(148, 210)
LETTER = PageSize(215.9, 279.4)
LEGAL = PageSize(215.9, 355.6)

#---------------------------------------------------------------------------------------- Align

class Align:
  """Text/element alignment constants. Match pdfmarq values for cross-lib reuse."""
  LEFT = "L"
  RIGHT = "R"
  CENTER = "C"
  JUSTIFY = "J"

#--------------------------------------------------------------------------------------- Colors

class Colors:
  """Predefined colors as (r, g, b) tuples (0-1 range)."""
  BLACK = (0, 0, 0)
  WHITE = (1, 1, 1)
  RED = (1, 0, 0)
  GREEN = (0, 1, 0)
  BLUE = (0, 0, 1)
  GREY = (0.5, 0.5, 0.5)
  LIGHT_GREY = (0.8, 0.8, 0.8)
  DARK_GREY = (0.3, 0.3, 0.3)

#------------------------------------------------------------------------------------- Defaults

class Defaults:
  """Default values for DOCX generation."""
  PAGE_WIDTH = 210
  PAGE_HEIGHT = 297
  MARGIN = 20
  FONT_FAMILY = "Calibri"
  FONT_SIZE = 11
  FONT_MODE = "Regular"
  LINE_HEIGHT = 1.15
  UNIT = "mm"
  # Heading palette - GitHub-light, mirrors `pdfmarq.MarkdownStyle`
  HEAD_COLOR = (0.09, 0.11, 0.13) # near-black #1f2328
  RULE_COLOR = (0.82, 0.84, 0.87) # light grey #d0d7de (h1/h2 underline)
  # h1..h6 sizes in pt. Matches `pdfmarq.MarkdownStyle.h{1-6}_size` so
  # the same markdown source renders at the same scale in both libs.
  HEAD_SIZES = (20, 16, 13, 11, 11, 11)
  HEAD_UNDERLINE_LEVELS = (1, 2) # which heading levels get bottom border
  # Inline `code` runs; 0-1 tuples so `color_hex` emits `1F2328` / `F2F4F6`.
  CODE_FAMILY = "Consolas"
  CODE_SIZE_RATIO = 0.92 # of body size, when no explicit code size is set
  CODE_COLOR = (0.09, 0.11, 0.13)
  CODE_BG = (0.949, 0.957, 0.965)
