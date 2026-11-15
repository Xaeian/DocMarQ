# docmarq/metrics.py

"""
Text widths for the layout calls Word leaves to the writer: table columns and letter header rows.

Word sets text in the fonts the document carries, else in those installed where it opens,
so a family is measured from the file a document carries, else from the system TTF Pillow finds,
or from a metric-compatible stand-in.
A family found nowhere measures as Helvetica.
"""
from functools import lru_cache
from PIL import ImageFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from .constants import Unit
from .fonts import resolve_ttf

#-------------------------------------------------------------------------------------------- Faces

# Regular and bold TTF per family, first found wins.
# Carlito, Caladea and Liberation share metrics with the Microsoft fonts a Linux host lacks.
_FILES = {
  "arial": (("arial.ttf", "LiberationSans-Regular.ttf"), ("arialbd.ttf", "LiberationSans-Bold.ttf")),
  "calibri": (("calibri.ttf", "Carlito-Regular.ttf"), ("calibrib.ttf", "Carlito-Bold.ttf")),
  "cambria": (("cambria.ttc", "Caladea-Regular.ttf"), ("cambriab.ttf", "Caladea-Bold.ttf")),
  "consolas": (("consola.ttf",), ("consolab.ttf",)),
  "courier new": (("cour.ttf", "LiberationMono-Regular.ttf"), ("courbd.ttf", "LiberationMono-Bold.ttf")),
  "georgia": (("georgia.ttf",), ("georgiab.ttf",)),
  "segoe ui": (("segoeui.ttf",), ("segoeuib.ttf",)),
  "tahoma": (("tahoma.ttf",), ("tahomabd.ttf",)),
  "times new roman": (("times.ttf", "LiberationSerif-Regular.ttf"), ("timesbd.ttf", "LiberationSerif-Bold.ttf")),
  "verdana": (("verdana.ttf",), ("verdanab.ttf",)),
}

# Size a face loads at: widths scale linearly, and a large size keeps hinting out of them.
_LOAD_PT = 1000

@lru_cache(maxsize=None)
def _face(family:str, bold:bool, font_dir:str|None=None) -> ImageFont.FreeTypeFont|None:
  """The TTF behind `family`, looked up in `font_dir` first, or `None` when none is found."""
  if font_dir:
    path = resolve_ttf(font_dir, family.replace(" ", ""), "Bold" if bold else "Regular",
      fallback=False)
    if path is not None: return ImageFont.truetype(str(path), _LOAD_PT)
  regular, heavy = _FILES.get(family.lower(), ((f"{family}.ttf",), ()))
  for name in heavy if bold else regular:
    try:
      return ImageFont.truetype(name, _LOAD_PT)
    except OSError: continue
  return None

#------------------------------------------------------------------------------------------ Measure

def text_width_mm(
  text:str, family:str, size_pt:float, bold:bool = False,
  font_dir:str|None = None,
) -> float:
  """
  Width of `text` set in `family` at `size_pt`, in mm,
  from its file in `font_dir` where that holds one, as a document carrying its faces sets it.
  """
  face = _face(family, bold, font_dir)
  if face is not None: return face.getlength(text) * size_pt / _LOAD_PT * Unit.PT
  # Helvetica is wider than Calibri, so a guess from it errs toward a column that holds its words
  return stringWidth(text, "Helvetica-Bold" if bold else "Helvetica", size_pt) * Unit.PT
