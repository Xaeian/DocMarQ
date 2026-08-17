# docmarq/utils.py

"""Utility functions - unit conversion, color parsing, helpers."""
from .constants import Unit, EMU_PER_MM, EMU_PER_PT

#-------------------------------------------------------------------------------------------- Units

def to_mm(value:float, unit:str="mm") -> float:
  """Convert value from given unit to millimeters."""
  unit = unit.lower()
  factors = {"mm": Unit.MM, "cm": Unit.CM, "in": Unit.INCH, "inch": Unit.INCH,
    "pt": Unit.PT, "px": Unit.PX}
  if unit not in factors:
    raise ValueError(f"Unknown unit: {unit}. Use: mm, cm, in, pt, px")
  return value * factors[unit]

def mm_to_pt(*values:float) -> list[float]|float:
  """Convert mm values to points. Returns single value or list."""
  result = [v / Unit.PT for v in values]
  return result[0] if len(result) == 1 else result

def mm_to_emu(value_mm:float) -> int:
  """Convert millimeters to EMU (OOXML native unit)."""
  return int(round(value_mm * EMU_PER_MM))

def pt_to_emu(value_pt:float) -> int:
  """Convert points to EMU."""
  return int(round(value_pt * EMU_PER_PT))

def to_emu(value:float, unit:str="mm") -> int:
  """Convert value in given unit to EMU."""
  return mm_to_emu(to_mm(value, unit))

#------------------------------------------------------------------------------- Typographic ladder

# Word's font-size dropdown values - the de facto standard typographic
# ladder. Jumps are non-linear: 12→14→16 skips 13/15 because those don't
# read well in print. "One step smaller" should walk this list, not just
# subtract 1pt and hope for the best.
_SIZE_LADDER = (6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)

def smaller_size(body_pt:float, min_pt:float=7) -> float:
  """Next-smaller standard typographic size below `body_pt`. Used by tables
  and bibliography to derive a "one step smaller" reading size that hits
  expected Word ladder values (11→10, 12→11, 14→12, 16→14, 18→16) instead
  of arbitrary `body - 1` deltas that produce non-standard sizes like 13.
  Clamps at `min_pt` so small body sizes don't underflow.
  """
  for i, size in enumerate(_SIZE_LADDER):
    if size >= body_pt:
      return min_pt if i == 0 else max(min_pt, _SIZE_LADDER[i - 1])
  return max(min_pt, _SIZE_LADDER[-2])

def tight_line_height(size_pt:float) -> float:
  """Line-height multiplier appropriate for `size_pt`. Mirrors classical
  typographic rule: display text needs tighter leading than body text
  because the extra vertical air between lines reads as wasteful at large
  sizes. Linear interpolation between (11pt, 1.15) and (24pt, 1.00).

  Examples (rounded):
    11pt body  → 1.15
    13pt h3    → 1.12
    16pt h2    → 1.09
    20pt title → 1.05
    22pt h1    → 1.03
    24pt+      → 1.00

  Returned multiplier applies to Word's `lineRule="auto"` mode where the
  base unit is the font's recommended single-line height (`~1.2× font size`
  for Calibri). For really tight display, use `tight_line_height_pt`
  which returns an absolute pt value for `lineRule="exact"`.
  """
  if size_pt <= 11:
    return 1.15
  if size_pt >= 24:
    return 1.0
  return round(1.15 - (size_pt - 11) / (24 - 11) * 0.15, 3)

def tight_line_height_pt(size_pt:float) -> float:
  """Absolute line-height in points for `lineRule="exact"`. Same taper as
  `tight_line_height` but bypasses Calibri's internal leading (~1.2× font
  size in "single" mode), so the value is used verbatim by Word.

  Use over `tight_line_height` when auto-mode leading is too loose.

  Examples (Calibri, rounded):
    11pt body  → 12.7pt (`1.15×`)
    16pt h2    → 17.6pt (`1.10×`)
    20pt title → 21.6pt (`1.08×`)
    22pt h1    → 23.5pt (`1.07×`)
    24pt+      → 24.0pt (`1.00×` for very large)
  """
  if size_pt <= 11:
    return size_pt * 1.15
  if size_pt >= 24:
    return size_pt
  factor = 1.15 - (size_pt - 11) / (24 - 11) * 0.15
  return round(size_pt * factor, 2)

#------------------------------------------------------------------------------------------- Colors

_HEX_DIGITS = set("0123456789abcdefABCDEF")

def parse_color(color:tuple|str|None) -> tuple[float, float, float]:
  """Parse `(r,g,b)` tuple or hex string to 0-1 floats. Mirrors
  0-1 per channel - 0-255 conversion to OOXML happens at the
  boundary in `rgb255()`, keeping user-facing values consistently 0-1.
  """
  if color is None:
    return (0.0, 0.0, 0.0)
  if isinstance(color, tuple):
    if len(color) < 3:
      raise ValueError("Color tuple must have at least 3 values (r, g, b)")
    return (float(color[0]), float(color[1]), float(color[2]))
  if isinstance(color, str):
    raw = color
    h = color.lstrip("#")
    if len(h) == 3:
      h = "".join(c * 2 for c in h)
    if len(h) != 6:
      raise ValueError(
        f"Invalid hex color {raw!r}: expected #RGB or #RRGGBB, got {len(h)} digit(s)"
      )
    if not all(c in _HEX_DIGITS for c in h):
      raise ValueError(f"Invalid hex color {raw!r}: contains non-hex characters")
    return (int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255)
  raise ValueError(f"Invalid color type: {type(color)}")

def rgb255(color:tuple|str|None) -> tuple[int, int, int]:
  """Parse color and return `(r,g,b)` 0-255 ints - the form `python-docx`
  `RGBColor` wants. Internal boundary helper for OOXML emission."""
  r, g, b = parse_color(color)
  return (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))

def color_hex(color:tuple|str) -> str:
  """Return `RRGGBB` uppercase hex string (no `#` prefix) - OOXML format."""
  r, g, b = rgb255(color)
  return f"{r:02X}{g:02X}{b:02X}"

#------------------------------------------------------------------------------------------- Margin

def parse_margin(margin:float|tuple) -> tuple[float, float, float, float]:
  """Parse margin to `(top, right, bot, left)` tuple. CSS-style.

  Accepts:
    `n`                     → all sides `n`
    `(v,)`                  → all sides `v`
    `(v, h)`                → vertical `v`, horizontal `h`
    `(t, h, b)`             → top, horizontal, bottom
    `(t, r, b, l)`          → top, right, bottom, left
  """
  if isinstance(margin, (int, float)):
    return (margin, margin, margin, margin)
  if isinstance(margin, (tuple, list)):
    n = len(margin)
    if n == 1: return (margin[0], margin[0], margin[0], margin[0])
    if n == 2: return (margin[0], margin[1], margin[0], margin[1])
    if n == 3: return (margin[0], margin[1], margin[2], margin[1])
    if n == 4: return (margin[0], margin[1], margin[2], margin[3])
  raise ValueError(f"Invalid margin: {margin}")

#---------------------------------------------------------------------------------------- Alignment

def align_to_docx(align:str|None):
  """Map docmarq align constant to `python-docx` `WD_ALIGN_PARAGRAPH` value."""
  from docx.enum.text import WD_ALIGN_PARAGRAPH
  table = {
    "L": WD_ALIGN_PARAGRAPH.LEFT,
    "R": WD_ALIGN_PARAGRAPH.RIGHT,
    "C": WD_ALIGN_PARAGRAPH.CENTER,
    "J": WD_ALIGN_PARAGRAPH.JUSTIFY,
  }
  return table.get(align, WD_ALIGN_PARAGRAPH.LEFT) if align else None
