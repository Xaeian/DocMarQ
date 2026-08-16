# docmarq/styles.py

"""Style system - paragraph and run styling with inheritance support."""
from dataclasses import dataclass, fields, replace
from .constants import Defaults, Align

#---------------------------------------------------------------------------------------- Style

@dataclass
class Style:
  """Run/paragraph style. `None` values inherit from parent when merged.

  `font_mode` covers Bold/Italic/BoldItalic via separate flags below for
  cleaner toggling in inline runs - matches how Word stores them (each as
  its own `<w:b/>`, `<w:i/>` element on a run).
  """
  font_family: str|None = None
  font_size: float|None = None # pt
  color: tuple|str|None = None # `(r,g,b)` 0-1 or `#hex`
  bold: bool|None = None
  italic: bool|None = None
  underline: bool|None = None
  strike: bool|None = None
  align: str|None = None
  line_height: float|None = None # ratio of font size
  space_before: float|None = None # pt
  space_after: float|None = None # pt
  highlight: str|None = None # named highlight color: yellow/green/cyan/...
  # Inline `code` runs; `None` falls back to `Defaults.CODE_*`. The markdown
  # renderer feeds these from `MarkdownStyle`.
  code_family: str|None = None
  code_size: float|None = None # pt; None → `CODE_SIZE_RATIO` of body size
  code_color: tuple|str|None = None
  code_bg: tuple|str|None = None

  def merge(self, parent:"Style") -> "Style":
    """Create new style inheriting `None` values from parent."""
    out = Style()
    for f in fields(self):
      v = getattr(self, f.name)
      setattr(out, f.name, v if v is not None else getattr(parent, f.name))
    return out

  def with_defaults(self) -> "Style":
    """Fill `None` values with library defaults."""
    return Style(
      font_family=self.font_family or Defaults.FONT_FAMILY,
      font_size=self.font_size or Defaults.FONT_SIZE,
      color=self.color if self.color is not None else (0, 0, 0),
      bold=bool(self.bold),
      italic=bool(self.italic),
      underline=bool(self.underline),
      strike=bool(self.strike),
      align=self.align or Align.LEFT,
      line_height=self.line_height or Defaults.LINE_HEIGHT,
      space_before=self.space_before if self.space_before is not None else 4,
      space_after=self.space_after if self.space_after is not None else 4,
      highlight=self.highlight,
      code_family=self.code_family or Defaults.CODE_FAMILY,
      code_size=self.code_size,
      code_color=self.code_color if self.code_color is not None else Defaults.CODE_COLOR,
      code_bg=self.code_bg if self.code_bg is not None else Defaults.CODE_BG,
    )

  def copy(self, **overrides) -> "Style":
    """Return a copy with overridden fields."""
    valid = {k: v for k, v in overrides.items() if k in {f.name for f in fields(self)}}
    return replace(self, **valid)

#----------------------------------------------------------------------------------- TableStyle

@dataclass
class TableStyle:
  """Table-specific styling. Defaults match `pdfmarq` GitHub-light palette -
  light grey header, near-imperceptible zebra, light grey borders.

  Cell padding is split into `top`/`bot`/`h` (horizontal). Defaults are
  asymmetric vertical: more top than bottom - optical correction so text
  sits in the visual center, not the geometric one (font ascent > descent).
  """
  header_bg: tuple|str = (0.96, 0.97, 0.98) # #f6f8fa
  header_color: tuple|str|None = None
  header_bold: bool = True
  row_bg_even: tuple|str|None = None
  row_bg_odd: tuple|str|None = (0.985, 0.99, 0.995)
  border_color: tuple|str = (0.82, 0.84, 0.87) # #d0d7de
  border_size: float = 0.5 # pt
  cell_pad_top: float = 1.6 # mm - more, optical pushdown
  cell_pad_bot: float = 0.6 # mm - less
  cell_pad_h: float = 2.5 # mm horizontal
  vertical_align: str = "center"
  header_repeat: bool = True
  table_align: str|None = None
  fill_content_width: bool = True # auto-fill content area width
  # Cell font size in pt. `None` → auto-derive as `body_size - 1` rounded to
  # integer (for 11 pt body → 10 pt cells, common paper convention). Set
  # an explicit float to override.
  font_size: float|None = None

#-------------------------------------------------------------------------------------- Presets

class _Preset:
  """Descriptor returning a fresh `Style` per access - prevents shared mutation."""
  def __init__(self, **kwargs):
    self._kwargs = kwargs
  def __get__(self, obj, objtype=None) -> Style:
    return Style(**self._kwargs)

class Styles:
  """Predefined style presets. Each access returns a fresh `Style`."""
  DEFAULT = _Preset()
  BOLD = _Preset(bold=True)
  ITALIC = _Preset(italic=True)
  HEADING1 = _Preset(font_size=20, bold=True, space_before=12, space_after=6)
  HEADING2 = _Preset(font_size=16, bold=True, space_before=10, space_after=4)
  HEADING3 = _Preset(font_size=13, bold=True, space_before=8, space_after=3)
  HEADING4 = _Preset(font_size=11, bold=True, space_before=6, space_after=2)
  SMALL = _Preset(font_size=9)
  CAPTION = _Preset(font_size=9, italic=True, color=(0.4, 0.44, 0.5))
  CODE = _Preset(font_family="Consolas", font_size=10,
    highlight=None, color=(0.09, 0.11, 0.13))
