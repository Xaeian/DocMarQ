# docmarq/inline.py

"""
Rich inline runs - mixed-style spans within a single paragraph.

Emits Word runs (`<w:r>`) instead of PDF drawing ops.
`python-docx` owns line-breaking and justification.
This module only builds the run list.
"""
from dataclasses import dataclass
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from .constants import Defaults
from .utils import color_hex, rgb255

#-------------------------------------------------------------------------------------- RichSegment

@dataclass
class RichSegment:
  """
  Single styled run. Multiple segments compose into one paragraph.

  `link_url` makes the run a clickable external hyperlink.
  `link_target` references an internal bookmark by name.
  `code=True` flips the run to monospace family with subtle shading - matches markdown inline code.
  """
  text: str
  family: str|None = None # `None` → inherit current paragraph font
  size: float|None = None # pt, `None` → inherit
  color: tuple|str|None = None
  bold: bool = False
  italic: bool = False
  underline: bool = False
  strike: bool = False
  code: bool = False # render with mono family + light shading
  highlight: str|None = None # named highlight (yellow/green/cyan/...)
  link_url: str|None = None
  link_target: str|None = None # internal bookmark name
  break_line: bool = False # soft line break BEFORE this run
  bg_color: tuple|str|None = None # custom run background (status badges)
  superscript: bool = False # raise as superscript (footnote refs)
  subscript: bool = False # lower as subscript
  small_caps: bool = False # lowercase letters as Word's small capitals
  char_space: float = 0 # pt between letters beyond their own width, as Word's expanded spacing

#------------------------------------------------------------------------------------ Run rendering

def _apply_run_format(
  run, seg:RichSegment, default_family:str, default_size:float,
  code_family:str|None = None,
  code_size:float|None = None,
  code_color:tuple|str|None = None,
  code_bg:tuple|str|bool|None = None,
):
  """
  Apply `RichSegment` styling to a `python-docx` `Run`.
  Unset `code_*` arguments fall back to `Defaults.CODE_*`.
  `code_bg=False` leaves code unshaded, since `None` there is the grey of the defaults.
  """
  f = run.font
  if seg.code:
    f.name = code_family or Defaults.CODE_FAMILY
    f.size = Pt(code_size or default_size * Defaults.CODE_SIZE_RATIO)
    f.color.rgb = RGBColor(*rgb255(code_color or Defaults.CODE_COLOR))
  else:
    if seg.family or default_family:
      f.name = seg.family or default_family
    if seg.size or default_size:
      f.size = Pt(seg.size or default_size)
    if seg.color is not None:
      f.color.rgb = RGBColor(*rgb255(seg.color))
  if seg.bold: f.bold = True
  if seg.italic: f.italic = True
  if seg.underline: f.underline = True
  if seg.strike: f.strike = True
  if seg.highlight:
    f.highlight_color = _highlight_value(seg.highlight)
  # `bg_color` (status badges, callout titles) and code shading are mutually exclusive;
  # code always wins with its own background.
  if seg.code:
    if code_bg is not False: apply_run_shading(run, color_hex(code_bg or Defaults.CODE_BG))
  elif seg.bg_color is not None:
    apply_run_shading(run, color_hex(seg.bg_color))
  if seg.superscript:
    f.superscript = True
  if seg.subscript:
    f.subscript = True
  if seg.small_caps:
    f.small_caps = True
  if seg.char_space: set_char_space(run._element.get_or_add_rPr(), seg.char_space)

# CT_RPr child sequence (ISO/IEC 29500-1), where a run's or a style's properties place `spacing`
_RPR_ORDER = (
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
  "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color",
  "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
  "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
)

def set_char_space(rpr, pt:float):
  """
  Letters `pt` apart beyond their own width in the run or style properties `rpr`, as Word's
  expanded spacing, `<w:spacing>` in twentieths of a point at its place in the sequence.
  """
  old = rpr.find(qn("w:spacing"))
  if old is not None: rpr.remove(old)
  spacing = OxmlElement("w:spacing")
  spacing.set(qn("w:val"), str(round(pt * 20)))
  after = set(_RPR_ORDER[_RPR_ORDER.index("spacing") + 1:])
  later = next((child for child in rpr if child.tag.split("}")[-1] in after), None)
  if later is None: rpr.append(spacing)
  else: later.addprevious(spacing)

def _highlight_value(name:str):
  """Map color name to `WD_COLOR_INDEX`. Returns `None` if unknown."""
  table = {
    "yellow": WD_COLOR_INDEX.YELLOW, "green": WD_COLOR_INDEX.BRIGHT_GREEN,
    "cyan": WD_COLOR_INDEX.TURQUOISE, "magenta": WD_COLOR_INDEX.PINK,
    "blue": WD_COLOR_INDEX.BLUE, "red": WD_COLOR_INDEX.RED,
    "grey": WD_COLOR_INDEX.GRAY_25, "gray": WD_COLOR_INDEX.GRAY_25,
  }
  return table.get(name.lower())

def apply_run_shading(run, hex_color:str):
  """Apply background shading to a run via raw `<w:shd>` in `rPr`."""
  rpr = run._element.get_or_add_rPr()
  shd = OxmlElement("w:shd")
  shd.set(qn("w:val"), "clear")
  shd.set(qn("w:color"), "auto")
  shd.set(qn("w:fill"), hex_color)
  rpr.append(shd)
