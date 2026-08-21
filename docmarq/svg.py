# docmarq/svg.py

"""
SVG rasterization for DOCX embedding.

python-docx has no SVG support, so vector sources are turned into PNG
before they reach `add_picture`. Lives at package level, next to `core`:
`svglib` and `rlPyCairo` are base dependencies and both the fluent API and
the markdown renderer embed images.

Example:
  >>> from docmarq.svg import svg_to_png_buffer
  >>> buf = svg_to_png_buffer("logo.svg")
  >>> # buf is a BytesIO of PNG bytes, or None when rasterizing is impossible
"""
import io

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

def is_svg(path:str) -> bool:
  """True for a path naming an SVG file."""
  return bool(path) and path.lower().endswith(".svg")
