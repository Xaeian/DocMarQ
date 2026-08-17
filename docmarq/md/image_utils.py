"""
Pure image preprocessing for markdown figures and frontmatter logos.

Three pure operations: crop transparent borders, compute scaled dimensions,
parse the markdown image title DSL. The actual `add_picture` call lives on
the renderer where it needs access to the document and the paragraph
layout state.
"""
import io
import os
from dataclasses import dataclass

#--------------------------------------------------------------------------------------- Preprocess

def _svg_to_png_buffer(path:str):
  """Rasterize SVG to PNG bytes in a `BytesIO`. Returns `None` for a broken SVG.

  python-docx has no native SVG support, so we go SVG → reportlab `Drawing`
  → PNG via `renderPM` (rlPyCairo backend). DPI of 192 (~2x screen) keeps
  the raster crisp when Word scales it inside the page width.
  """
  try:
    import rlPyCairo  # renderPM loads its backend lazily - check here, fail fast
    from svglib.svglib import svg2rlg
    from reportlab.graphics import renderPM
  except ImportError:
    raise ImportError("Install with: pip install docmarq[md]")
  try:
    drawing = svg2rlg(path)
    if drawing is None:
      return None
    png_bytes = renderPM.drawToString(drawing, fmt="PNG", dpi=192)
  except Exception:
    return None
  return io.BytesIO(png_bytes)

def preprocess_to_buffer(path:str):
  """Open `path`, crop alpha-transparent borders, re-save as PNG into a
  `BytesIO`. Returns the buffer or `None` if Pillow / file unavailable.

  SVG inputs route through `_svg_to_png_buffer` first since neither Pillow
  nor python-docx can read them natively.

  Word honors the declared bounding box including transparent margins, so
  uncropped images appear smaller or misaligned. Cropping to the opaque
  content removes that bias.

  Also normalizes format: python-docx rejects JPEGs with non-standard APP
  segments (e.g. APP2 ICC profile); re-saving as PNG avoids that.
  """
  if not os.path.isfile(path):
    return None
  if path.lower().endswith(".svg"):
    return _svg_to_png_buffer(path)
  try:
    from PIL import Image
  except ImportError:
    return None
  try:
    im = Image.open(path)
    if im.mode == "P":
      # palette - convert so `getbbox` sees alpha
      im = im.convert("RGBA")
    if im.mode in ("RGBA", "LA"):
      bbox = im.getbbox()
      if bbox is not None:
        im = im.crop(bbox)
    elif im.mode != "RGB": im = im.convert("RGB")
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf
  except (OSError, ValueError):
    return None

#------------------------------------------------------------------------------------------ Scaling

def compute_target_dims(nat_w:int, nat_h:int, content_w_mm:float,
    max_h_mm:float) -> tuple[float|None, float]:
  """Width-first scaling with height cap. Preserves aspect ratio.

  Fits to `content_w_mm` unless the resulting height would exceed
  `max_h_mm`, in which case scales by height instead so the image never
  blows the page.
  """
  if nat_w <= 0 or nat_h <= 0:
    return (None, max_h_mm)
  aspect = nat_h / nat_w
  height_full = content_w_mm * aspect
  if height_full > max_h_mm:
    height = max_h_mm
    width = height / aspect
  else:
    width = content_w_mm
    height = height_full
  return (width, height)

def scale_dims_for_path(path:str, content_w_mm:float,
    max_h_mm:float) -> tuple[float|None, float]:
  """`compute_target_dims` driven by a file's natural dimensions."""
  try:
    from PIL import Image
    im = Image.open(path)
    nat_w, nat_h = im.size
  except (ImportError, OSError, ValueError):
    return (None, max_h_mm)
  return compute_target_dims(nat_w, nat_h, content_w_mm, max_h_mm)

def scale_dims_for_buffer(buf, content_w_mm:float,
    max_h_mm:float) -> tuple[float|None, float]:
  """`compute_target_dims` driven by a `BytesIO`. Resets buffer position
  when done so the caller can `add_picture` from it.
  """
  try:
    from PIL import Image
    buf.seek(0)
    im = Image.open(buf)
    nat_w, nat_h = im.size
    buf.seek(0)
  except (ImportError, OSError, ValueError):
    return (None, max_h_mm)
  return compute_target_dims(nat_w, nat_h, content_w_mm, max_h_mm)

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
  """Parse `key=value` space-separated DSL from an image title slot.

  Supported keys (case-insensitive):
    `w`, `h`         exact dimension in mm
    `max_w`, `max_h` soft cap in mm
    `scale`          float multiplier on natural size (absolute priority)
    `align`          `L` / `C` / `R` block-level horizontal alignment

  Title with no `=` token is treated as opaque (legacy "caption" titles).
  Unknown keys / malformed values / non-positive numerics each warn once
  and the key is dropped; parsing never raises.
  """
  out = ImageDSL()
  if not title or not title.strip():
    return out
  tokens = title.split()
  if not any("=" in t for t in tokens):
    return out
  out.is_dsl = True
  import warnings
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

def apply_dsl_dims(nat_w_px:int, nat_h_px:int, content_w_mm:float,
    max_h_mm:float, dsl:ImageDSL) -> tuple[float|None, float]:
  """Combine DSL overrides with the default width-first / height-cap flow.

  Resolution order:
    1. `scale` wins absolutely - natural size × scale, clamped to caps.
    2. `w` + `h` both set - exact box (aspect can break), clamped.
    3. `w` or `h` alone - exact dim + aspect-locked other, clamped.
    4. None of the above - normal flow with optional `max_w`/`max_h` caps.

  Final size is always clamped to the effective caps so explicit
  overrides can't escape the page bounds.
  """
  eff_w = content_w_mm
  eff_h = max_h_mm
  if dsl.max_w_mm is not None: eff_w = min(eff_w, dsl.max_w_mm)
  if dsl.max_h_mm is not None: eff_h = min(eff_h, dsl.max_h_mm)
  if nat_w_px <= 0 or nat_h_px <= 0:
    return compute_target_dims(nat_w_px, nat_h_px, eff_w, eff_h)
  aspect = nat_h_px / nat_w_px
  if dsl.scale is not None:
    # Natural size in mm at 96 DPI - matches Pillow assumption when no `dpi`
    # metadata is present; consistent with `compute_target_dims` semantics.
    nat_w_mm = nat_w_px * 25.4 / 96
    nat_h_mm = nat_h_px * 25.4 / 96
    return _clamp(nat_w_mm * dsl.scale, nat_h_mm * dsl.scale, eff_w, eff_h)
  if dsl.exact_w_mm is not None and dsl.exact_h_mm is not None:
    return _clamp(dsl.exact_w_mm, dsl.exact_h_mm, eff_w, eff_h)
  if dsl.exact_w_mm is not None:
    return _clamp(dsl.exact_w_mm, dsl.exact_w_mm * aspect, eff_w, eff_h)
  if dsl.exact_h_mm is not None:
    return _clamp(dsl.exact_h_mm / aspect, dsl.exact_h_mm, eff_w, eff_h)
  return compute_target_dims(nat_w_px, nat_h_px, eff_w, eff_h)

def _clamp(w:float, h:float, max_w:float, max_h:float) -> tuple[float, float]:
  """Scale down uniformly to fit `(max_w, max_h)` - no upscale."""
  if w <= 0 or h <= 0:
    return w, h
  sw = max_w / w if max_w > 0 and w > max_w else 1.0
  sh = max_h / h if max_h > 0 and h > max_h else 1.0
  s = min(sw, sh)
  return w * s, h * s

#---------------------------------------------------------------------------------- Path resolution

def resolve_path(src:str, base_dir:str) -> str|None:
  """Resolve image `src` against `base_dir`. Absolute paths returned as-is.
  Remote URLs not supported (would require download) - returns `None`.
  """
  if not src:
    return None
  if src.startswith(("http://", "https://", "ftp://")):
    return None
  if os.path.isabs(src):
    return src
  return os.path.normpath(os.path.join(base_dir, src))
