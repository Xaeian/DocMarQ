# docmarq/svg.py

"""
SVG rasterization for DOCX embedding.

python-docx has no SVG support, so vector sources are turned into PNG
before they reach `add_picture`. Lives at package level, next to `core`:
`svglib` and `rlPyCairo` are base dependencies and both the fluent API and
the markdown renderer embed images.

`register_fonts` is the other half: svglib resolves its own fonts and
defaults to Helvetica. The markdown renderer calls it whenever it was given
a `font_dir`; the fluent API cannot guess one, so call it yourself before
embedding an SVG that carries text.

Example:
  >>> from docmarq.svg import register_fonts, svg_to_png_buffer
  >>> register_fonts("./fonts", "IBMPlexSans")
  >>> buf = svg_to_png_buffer("logo.svg")
  >>> # buf is a BytesIO of PNG bytes, or None when rasterizing is impossible
"""
import io, threading
from .fonts import resolve_ttf

SVG_TARGET_PX = 2400 # longest raster side for SVG → PNG (~360 DPI at A4 content width)

_BACKEND_WARNED = False

#---------------------------------------------------------------------------------------- Rasterize

def _warn_backend_once(detail:str) -> None:
  """Report a missing rasterizer once per process.

  Embedding is best-effort: callers fall back to alt text, so a broken
  backend costs one figure rather than the whole document. The warning is
  the only trace that would otherwise be missing.
  """
  global _BACKEND_WARNED
  if _BACKEND_WARNED:
    return
  _BACKEND_WARNED = True
  import warnings
  warnings.warn(
    f"SVG rasterizing unavailable ({detail}); SVG images render as alt text. "
    "Install with: pip install svglib rlPyCairo",
    RuntimeWarning, stacklevel=3,
  )

def svg_to_png_buffer(path:str):
  """Rasterize an SVG into a `BytesIO` of PNG bytes.

  Returns `None` when the file cannot be rasterized - broken SVG, or no
  working backend. Scaled to `SVG_TARGET_PX` so a small viewBox stays sharp
  at page width.
  """
  try:
    import rlPyCairo # renderPM loads its backend lazily - check here, fail fast
    from svglib.svglib import svg2rlg
    from reportlab.graphics import renderPM
  except ImportError as e:
    _warn_backend_once(str(e))
    return None
  try:
    drawing = svg2rlg(path)
    if drawing is None:
      return None
    longest = max(drawing.width, drawing.height)
    if longest > 0:
      s = SVG_TARGET_PX / longest
      drawing.scale(s, s)
      drawing.width *= s
      drawing.height *= s
    png_bytes = renderPM.drawToString(drawing, fmt="PNG")
  except Exception:
    return None
  return io.BytesIO(png_bytes)

#------------------------------------------------------------------------------------- Font mapping

# `font-weight` values svglib can be asked for, and the mode fragment serving each.
# svglib keys its map on the literal attribute text, so `700` never reaches the entry
# registered under `bold` - every numeric weight needs its own alias. `900` is the
# only address `Black` has, because CSS has no name for it.
_WEIGHTS = {
  "normal": "", "100": "", "200": "", "300": "", "400": "", "500": "",
  "bold": "Bold", "600": "Bold", "700": "Bold", "800": "Bold", "900": "Black",
}

# `font-style`, composed onto the weight fragment above: bold + italic → `BoldItalic`.
_STYLES = {"normal": "", "italic": "Italic"}

# Every mode both tables compose. Six TTFs back twenty-two combinations, so a
# family resolves per mode and maps per combination.
_MODES = tuple(dict.fromkeys(
  (weight + style) or "Regular"
  for weight in _WEIGHTS.values() for style in _STYLES.values()
))

_MAPPED: set[tuple[str, str]] = set()
_LOCK = threading.Lock()

def register_fonts(font_dir:str, *families:str) -> None:
  """Make `families` resolvable for text drawn inside an SVG.

  Rasterizing hands `<text>` to svglib, which falls back to Helvetica and takes
  the typeface and every glyph outside Latin-1 with it. Registering the TTFs the
  document uses keeps a diagram on the page's own face.

  The registry is process-global, so a family stays mapped for process life,
  and a lock keeps a second thread off a half-mapped family.
  A family with no TTF is skipped in silence. Without svglib nothing is mapped at all,
  which is what `svg_to_png_buffer` warns about. Repeat calls are free.
  """
  try:
    from svglib.fonts import get_global_font_map
  except ImportError:
    return
  font_map = get_global_font_map()
  with _LOCK:
    for family in families:
      if (font_dir, family) in _MAPPED: continue
      paths = {mode: resolve_ttf(font_dir, family, mode) for mode in _MODES}
      for weight, weight_mode in _WEIGHTS.items():
        for style, style_mode in _STYLES.items():
          path = paths[(weight_mode + style_mode) or "Regular"]
          if path is None: continue
          # `<family>-<mode>` of the file that answered, not of the mode asked for,
          # so a weight served by a fallback reuses that fallback's registration.
          font_map.register_font(
            family, str(path), weight=weight, style=style, rlgFontName=path.stem,
          )
      _MAPPED.add((font_dir, family))

#---------------------------------------------------------------------------------------- Detection

def is_svg(path:str) -> bool:
  """True for a path naming an SVG file."""
  return bool(path) and path.lower().endswith(".svg")
