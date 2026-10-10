# docmarq/images.py

"""
Pictures for Word: an image as a PNG Word takes, and the offsets python-docx leaves out.
The markdown engine embeds its figures through these, the paper banner its logo.
"""
import io, os
from PIL import Image
from docx.oxml import OxmlElement
from .svg import svg_to_png_buffer, is_svg

#--------------------------------------------------------------------------------------- Preprocess

def preprocess_to_buffer(path:str):
  """Open `path`, crop alpha-transparent borders, re-save as PNG into a `BytesIO`.
  Returns the buffer or `None` if Pillow / file unavailable.

  SVG inputs route through `svg_to_png_buffer` first
  since neither Pillow nor python-docx can read them natively.

  Word honors the declared bounding box including transparent margins,
  so uncropped images appear smaller or misaligned.
  Cropping to the opaque content removes that bias.

  Also normalizes format:
  python-docx rejects JPEGs with non-standard APP segments (e.g. APP2 ICC profile);
  re-saving as PNG avoids that.
  """
  if not os.path.isfile(path):
    return None
  if is_svg(path):
    return svg_to_png_buffer(path)
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

#---------------------------------------------------------------------------------- Picture offsets

_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def pin_offsets(run):
  """
  Set `distT/B/L/R="0"` and inject a zero `<wp:effectExtent>` on every `<wp:inline>` in this run.
  python-docx omits these;
  LibreOffice and some Word versions render a small horizontal offset when they are absent.
  """
  def wp(tag):
    return f"{{{_WP_NS}}}{tag}"
  for inline in run._element.findall(f".//{wp('inline')}"):
    for attr in ("distT", "distB", "distL", "distR"):
      inline.set(attr, "0")
    eff = inline.find(wp("effectExtent"))
    if eff is None:
      eff = OxmlElement("wp:effectExtent")
      for attr in ("l", "t", "r", "b"):
        eff.set(attr, "0")
      # OOXML schema order: `extent` → `effectExtent` → `docPr` → ...
      extent = inline.find(wp("extent"))
      if extent is not None:
        idx = list(inline).index(extent) + 1
        inline.insert(idx, eff)
      else:
        inline.insert(0, eff)
