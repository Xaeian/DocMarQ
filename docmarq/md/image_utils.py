"""
Sizing of markdown figures: the image title DSL, scaled dimensions, DPI and paths,
and the rules for an image in a line of text or in a table cell.
The `add_picture` call lives on the renderer,
where it needs access to the document and the paragraph layout state.
"""
import os, warnings
from dataclasses import dataclass
from svglib.svglib import svg2rlg
from PIL import Image as PILImage
from ..constants import MM_TO_PT

#------------------------------------------------------------------------------------------ Scaling

def read_dpi(path:str, default:float=96) -> float:
  """Metadata DPI _(pHYs / JFIF)_, `default` when absent."""
  try:
    with PILImage.open(path) as im:
      meta = im.info.get("dpi")
  except (OSError, ValueError): return default
  return meta[0] if isinstance(meta, tuple) and meta[0] else default

def _upscale_cap_mm(nat_w_px:int, dpi:float|None, min_dpi:float) -> float:
  """Widest the image may grow to before it turns soft."""
  if not dpi or min_dpi <= 0: return float("inf")
  return nat_w_px * 25.4 / min(dpi, min_dpi)

def compute_target_dims(
  nat_w:int, nat_h:int, content_w_mm:float,
  max_h_mm:float, dpi:float|None = 96, min_dpi:float = 150,
) -> tuple[float|None, float]:
  """
  Width-first scaling with height cap. Preserves aspect ratio.

  Fills `content_w_mm` while the raster stays above `min_dpi`.
  Vector sources _(`dpi=None`)_ and `min_dpi=0` always fill it.
  A result taller than `max_h_mm` scales by height instead.
  """
  if nat_w <= 0 or nat_h <= 0: return (None, max_h_mm)
  aspect = nat_h / nat_w
  width = min(content_w_mm, _upscale_cap_mm(nat_w, dpi, min_dpi))
  height = width * aspect
  if height > max_h_mm:
    height = max_h_mm
    width = height / aspect
  return (width, height)

#---------------------------------------------------------------------------------------- Title DSL

@dataclass
class ImageDSL:
  """Parsed `![alt](src "key=value ...")` title DSL."""
  exact_w_mm: float|None = None
  exact_h_mm: float|None = None
  max_w_mm: float|None = None
  max_h_mm: float|None = None
  scale: float|None = None
  align: str|None = None
  is_dsl: bool = False

_DSL_NUMERIC_KEYS = {"w", "h", "max_w", "max_h", "scale"}
_DSL_ALIGN_VALUES = {"L", "C", "R"}

def parse_image_dsl(title:str|None) -> ImageDSL:
  """
  Parse `key=value` space-separated DSL from an image title slot.

  Supported keys (case-insensitive):
    `w`, `h`         exact dimension in mm
    `max_w`, `max_h` soft cap in mm
    `scale`          float multiplier on natural size (absolute priority)
    `align`          `L` / `C` / `R` block-level horizontal alignment

  Title with no `=` token is treated as opaque (a plain caption).
  Unknown keys / malformed values / non-positive numerics each warn once and the key is dropped;
  parsing never raises.
  """
  out = ImageDSL()
  if not title or not title.strip(): return out
  tokens = title.split()
  if not any("=" in t for t in tokens): return out
  out.is_dsl = True
  for tok in tokens:
    if "=" not in tok:
      warnings.warn(
        f"image title token {tok!r} is not `key=value`, ignored",
        RuntimeWarning, stacklevel=2,
      )
      continue
    key, _, val = tok.partition("=")
    key = key.strip().lower()
    val = val.strip()
    if key == "align":
      v = val.upper()
      if v in _DSL_ALIGN_VALUES:
        out.align = v
      else:
        warnings.warn(
          f"image title align={val!r} must be L/C/R, ignored",
          RuntimeWarning, stacklevel=2,
        )
      continue
    if key not in _DSL_NUMERIC_KEYS:
      warnings.warn(
        f"unknown image title key {key!r}, ignored",
        RuntimeWarning, stacklevel=2,
      )
      continue
    try:
      fv = float(val)
    except ValueError:
      warnings.warn(
        f"image title {key}={val!r} not a number, ignored",
        RuntimeWarning, stacklevel=2,
      )
      continue
    if fv <= 0:
      warnings.warn(
        f"image title {key}={fv} must be > 0, ignored",
        RuntimeWarning, stacklevel=2,
      )
      continue
    if key == "w": out.exact_w_mm = fv
    elif key == "h": out.exact_h_mm = fv
    elif key == "max_w": out.max_w_mm = fv
    elif key == "max_h": out.max_h_mm = fv
    elif key == "scale": out.scale = fv
  return out

def apply_dsl_dims(
  nat_w_px:int, nat_h_px:int, content_w_mm:float,
  max_h_mm:float, dsl:ImageDSL, dpi:float|None = 96,
  min_dpi:float = 150,
) -> tuple[float|None, float]:
  """
  Combine DSL overrides with the default width-first / height-cap flow.

  Resolution order:
    1. `scale` wins absolutely - natural size × scale, clamped to caps.
    2. `w` + `h` both set - exact box (aspect can break), clamped.
    3. `w` or `h` alone - exact dim + aspect-locked other, clamped.
    4. None of the above - normal flow with optional `max_w`/`max_h` caps.

  Final size is always clamped to the effective caps
  so explicit overrides can't escape the page bounds.
  """
  eff_w = content_w_mm
  eff_h = max_h_mm
  if dsl.max_w_mm is not None: eff_w = min(eff_w, dsl.max_w_mm)
  if dsl.max_h_mm is not None: eff_h = min(eff_h, dsl.max_h_mm)
  if nat_w_px <= 0 or nat_h_px <= 0:
    return compute_target_dims(nat_w_px, nat_h_px, eff_w, eff_h, dpi, min_dpi)
  aspect = nat_h_px / nat_w_px
  if dsl.scale is not None:
    nat_w_mm = nat_w_px * 25.4 / (dpi or 96)
    nat_h_mm = nat_h_px * 25.4 / (dpi or 96)
    return _clamp_no_upscale(nat_w_mm * dsl.scale, nat_h_mm * dsl.scale, eff_w, eff_h)
  if dsl.exact_w_mm is not None and dsl.exact_h_mm is not None:
    return _clamp_no_upscale(dsl.exact_w_mm, dsl.exact_h_mm, eff_w, eff_h)
  if dsl.exact_w_mm is not None:
    return _clamp_no_upscale(dsl.exact_w_mm, dsl.exact_w_mm * aspect, eff_w, eff_h)
  if dsl.exact_h_mm is not None:
    return _clamp_no_upscale(dsl.exact_h_mm / aspect, dsl.exact_h_mm, eff_w, eff_h)
  return compute_target_dims(nat_w_px, nat_h_px, eff_w, eff_h, dpi, min_dpi)

#---------------------------------------------------------------------------------- Path resolution

def resolve_path(src:str, base_dir:str) -> str|None:
  """
  Resolve image `src` against `base_dir`. Absolute paths returned as-is.
  Remote URLs not supported (would require download) - returns `None`.
  """
  if not src: return None
  if src.startswith(("http://", "https://", "ftp://")): return None
  if os.path.isabs(src): return src
  return os.path.normpath(os.path.join(base_dir, src))

#--------------------------------------------------------------------------- Inline and cell images

# Vendored from `pdfmarq.md.md_images` on purpose: an image in a line or a cell sizes alike
# in both packages, and a cross-lib test keeps the copies equal.

@dataclass
class ImageInfo:
  """Resolved metadata + author hints for one image."""
  src: str
  is_svg: bool
  nat_w_mm: float  # natural width in mm at metadata-DPI (or 96 fallback)
  nat_h_mm: float  # natural height in mm
  dimensionless: bool = False # True only when SVG has no width/height/viewBox
  dpi: float|None = None # DPI behind `nat_*_mm`, None for vector
  generated: bool = False # rasterized here from vector source - upscaling is free
  explicit_w_mm: float|None = None # author-specified width (DSL `w=` or `scale=`)
  explicit_h_mm: float|None = None # author-specified height (DSL `h=` or `scale=`)
  dsl_max_w_mm: float|None = None  # DSL `max_w=` soft cap
  dsl_max_h_mm: float|None = None  # DSL `max_h=` soft cap
  align: str|None = None # DSL `align=L/C/R` block-level horizontal alignment
  alt: str = ""

def load_image_info(
  src:str,
  attrs:dict|None = None,
  alt:str = "",
  default_dpi:int = 96,
) -> ImageInfo|None:
  """
  Read image metadata + parse author overrides.
  Returns `None` if the file cannot be opened.

  Author overrides come from the markdown image title `![alt](src "DSL")`,
  parsed as space-separated `key=value` pairs (`w h max_w max_h scale align`).
  """
  if not src or not os.path.exists(src): return None
  is_svg = src.lower().endswith(".svg")
  if is_svg:
    dims = _load_svg_dims(src)
  else:
    dims = _load_raster_dims(src, default_dpi)
  if dims is None: return None
  nat_w_mm, nat_h_mm, dimensionless, dpi = dims
  info = ImageInfo(
    src=src, is_svg=is_svg,
    nat_w_mm=nat_w_mm, nat_h_mm=nat_h_mm,
    dimensionless=dimensionless, dpi=dpi, alt=alt,
  )
  title = None
  if attrs:
    title = (attrs.get("title") if isinstance(attrs, dict)
      else dict(attrs).get("title"))
  return apply_dsl(info, parse_image_dsl(title))

def _load_svg_dims(src:str) -> tuple[float, float, bool, None]|None:
  """
  Returns `(w_mm, h_mm, dimensionless, None)`.
  `dimensionless=True` when the SVG declares neither intrinsic dims nor viewBox -
  render context decides the size.
  """
  try:
    d = svg2rlg(src)
    if d is None: return None
    w, h = float(d.width or 0), float(d.height or 0)
    if w <= 0 or h <= 0: return 1.0, 1.0, True, None # dimensionless - aspect derived from context
    return w / MM_TO_PT, h / MM_TO_PT, False, None
  except Exception: return None

def _load_raster_dims(src:str, default_dpi:int) -> tuple[float, float, bool, float]|None:
  """
  Returns `(w_mm, h_mm, False, dpi)`.
  DPI is read from PIL `info["dpi"]` _(present in PNG/JPEG metadata as pHYs/JFIF resolution)_,
  fallback to `default_dpi`.
  """
  try:
    with PILImage.open(src) as im:
      px_w, px_h = im.size
      meta = im.info.get("dpi")
    real_dpi = meta[0] if isinstance(meta, tuple) and meta[0] else default_dpi
    return px_w * 25.4 / real_dpi, px_h * 25.4 / real_dpi, False, real_dpi
  except Exception: return None

def apply_dsl(info:ImageInfo, dsl:ImageDSL) -> ImageInfo:
  """
  Fold a parsed title DSL into `info` and return it.
  `scale` wins absolutely, `w`/`h` override the flow, `max_*` stay soft caps.
  """
  if not dsl.is_dsl: return info
  if dsl.scale is not None:
    info.explicit_w_mm = info.nat_w_mm * dsl.scale
    info.explicit_h_mm = info.nat_h_mm * dsl.scale
  else:
    if dsl.exact_w_mm is not None: info.explicit_w_mm = dsl.exact_w_mm
    if dsl.exact_h_mm is not None: info.explicit_h_mm = dsl.exact_h_mm
  info.dsl_max_w_mm = dsl.max_w_mm
  info.dsl_max_h_mm = dsl.max_h_mm
  info.align = dsl.align
  return info

def size_inline(
  info:ImageInfo,
  inline_cap_mm:float,
) -> tuple[float, float]:
  """
  Size an inline mid-paragraph image.
  Always height-capped at `inline_cap_mm` (LaTeX `height=2ex` idiom).
  Width scales proportionally.
  DSL `max_w`/`max_h` further narrow the cap, and `align` is ignored, the image riding the text.
  """
  cap_w = inline_cap_mm * 6
  cap_h = inline_cap_mm
  if info.dsl_max_w_mm is not None: cap_w = min(cap_w, info.dsl_max_w_mm)
  if info.dsl_max_h_mm is not None: cap_h = min(cap_h, info.dsl_max_h_mm)
  if info.explicit_w_mm or info.explicit_h_mm: return _explicit(info, cap_w, cap_h)
  effective_cap_h = inline_cap_mm
  if info.dsl_max_h_mm is not None:
    effective_cap_h = min(effective_cap_h, info.dsl_max_h_mm)
  nw, nh = info.nat_w_mm, info.nat_h_mm
  if nh <= 0: return effective_cap_h, effective_cap_h
  if nh > effective_cap_h:
    nw, nh = nw * (effective_cap_h / nh), effective_cap_h
  # Width is capped too: a very wide, very short image clears the height cap untouched
  # and would push the line box past the page edge.
  if cap_w > 0 and nw > cap_w:
    nh = nh * (cap_w / nw)
    nw = cap_w
  return nw, nh

def cell_intrinsic_w_mm(
  info:ImageInfo,
  inline_cap_mm:float,
  cell_image_max_w_mm:float,
  cell_image_scale:float = 0.5,
) -> tuple[float, float]:
  """
  Width range an image cell contributes to HTML auto-layout.

  Returns `(col_min, col_max)`:

  - **col_max**: `min(natural × cell_image_scale, cell_image_max_w_mm)`.
    Half-natural is a heuristic that works across both common table styles:
    icon-grids (small natural like 50mm symbols → render ~25mm)
    and procedure tables (large natural like 100mm photos → render ~50mm).
    The image-to-text visual ratio ends up roughly 1:2-1:3 in both cases,
    matching how GitHub renders tables with `td img { max-width: 100% }`
    after auto-layout pressures the column.
  - **col_min**: `max(inline_cap, 60% × col_max)`.
    Image cells participate in overflow shrink without collapsing to icon scale.

  Explicit author width wins, clamped to absolute cap only.
  Dimensionless SVG = inline cap (icon).
  """
  if info.explicit_w_mm:
    v = min(info.explicit_w_mm, cell_image_max_w_mm)
    return v, v
  if info.is_svg and info.dimensionless: return inline_cap_mm, inline_cap_mm
  col_max = min(info.nat_w_mm * cell_image_scale, cell_image_max_w_mm)
  # `inline_cap_mm` is a readability floor, clamped to the max:
  # an icon narrower than the cap would otherwise invert the range.
  col_min = min(max(inline_cap_mm, col_max * 0.6), col_max)
  return col_min, col_max

def size_cell(
  info:ImageInfo,
  cell_w_mm:float,
  max_h_mm:float,
  inline_cap_mm:float,
  cell_image_max_w_mm:float = 60,
  cell_image_scale:float = 0.5,
) -> tuple[float, float]:
  """
  Size an image inside a table cell.

  Render at `min(natural × cell_image_scale, cell_image_max_w_mm, cell_w_mm, max_h_mm)`.
  Never upscales (unlike block SVG) - a cell-filling SVG looks broken inside a multi-column table.
  The scale + abs cap mirror the intrinsic-width logic,
  so that when a mixed col gets stretched wider by long text, the image doesn't grow to fill -
  it stays at its preferred scaled size.
  Dimensionless SVGs render at inline cap (icon).
  """
  if info.explicit_w_mm or info.explicit_h_mm: return _explicit(info, cell_w_mm, max_h_mm)
  if info.is_svg and info.dimensionless: return inline_cap_mm, inline_cap_mm
  preferred_w = min(info.nat_w_mm * cell_image_scale, cell_image_max_w_mm)
  preferred_w = min(preferred_w, cell_w_mm)
  if preferred_w >= info.nat_w_mm:
    return _clamp_no_upscale(info.nat_w_mm, info.nat_h_mm, cell_w_mm, max_h_mm)
  s_factor = preferred_w / info.nat_w_mm
  w = preferred_w
  h = info.nat_h_mm * s_factor
  if max_h_mm > 0 and h > max_h_mm:
    s2 = max_h_mm / h
    w *= s2
    h = max_h_mm
  return w, h

def _clamp_no_upscale(
  nw:float, nh:float, max_w:float, max_h:float,
) -> tuple[float, float]:
  """
  Scale uniformly DOWN to fit `(max_w, max_h)`.
  Returns natural size when it already fits.
  """
  if nw <= 0 or nh <= 0: return nw, nh
  sw = max_w / nw if max_w > 0 and nw > max_w else 1.0
  sh = max_h / nh if max_h > 0 and nh > max_h else 1.0
  s = min(sw, sh)
  return nw * s, nh * s

def _explicit(
  info:ImageInfo, max_w:float, max_h:float,
) -> tuple[float, float]:
  """
  Resolve explicit width/height, a missing dimension from the natural aspect ratio.
  Final size still clamped to `(max_w, max_h)`.
  """
  ew, eh = info.explicit_w_mm, info.explicit_h_mm
  nw, nh = info.nat_w_mm, info.nat_h_mm
  if ew and eh:
    w, h = ew, eh
  elif ew:
    w = ew
    h = nh * (ew / nw) if nw > 0 else ew
  else:
    h = eh
    w = nw * (eh / nh) if nh > 0 else eh
  return _clamp_no_upscale(w, h, max_w, max_h)
