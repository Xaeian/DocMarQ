# docmarq/images.py

"""
Pictures for Word: an image as a PNG Word takes, an SVG drawing carried beside its PNG,
and the offsets python-docx leaves out.
The markdown engine embeds its figures through these, the paper banner its logo.
"""
import io, os
from pathlib import Path
from PIL import Image
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .svg import svg_to_png_buffer, is_svg, drawn_only

_SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" # the extension Word 2016 keeps an SVG in
_ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"

#--------------------------------------------------------------------------------------- Preprocess

def preprocess_to_buffer(path:str):
  """
  Open `path`, crop alpha-transparent borders, re-save as PNG into a `BytesIO`.
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
  if not os.path.isfile(path): return None
  if is_svg(path): return svg_to_png_buffer(path)
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
  except (OSError, ValueError): return None

#------------------------------------------------------------------------------------------ Picture

def add_picture(run, source, **size):
  """
  `source` as a picture of `run`, `size` as python-docx takes it.
  The PNG of an SVG drawing carries the drawing too, as Word 2016 keeps one: Word and LibreOffice
  draw its vectors, sharp at any size, an older reader the PNG.
  An SVG holding text stays a PNG alone, its text set here in faces a reader may lack.
  """
  shape = run.add_picture(source, **size)
  svg = getattr(source, "svg", "")
  if svg and drawn_only(svg): _carry_svg(run, svg)
  return shape

def _carry_svg(run, path:str):
  """The SVG at `path` beside the PNG of the last picture of `run`, one part for each drawing."""
  part, data = run.part, Path(path).read_bytes()
  svg = next((p for p in part.package.iter_parts()
    if p.content_type == "image/svg+xml" and p.blob == data), None)
  if svg is None:
    svg = Part(part.package.next_partname("/word/media/image%d.svg"), "image/svg+xml", data,
      part.package)
  ext_list, ext = OxmlElement("a:extLst"), OxmlElement("a:ext")
  ext.set("uri", _SVG_EXT)
  blip = ext.makeelement(f"{{{_ASVG}}}svgBlip", nsmap={"asvg": _ASVG})
  blip.set(qn("r:embed"), part.relate_to(svg, RT.IMAGE))
  ext.append(blip)
  ext_list.append(ext)
  run._r.xpath(".//a:blip")[-1].append(ext_list)

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
