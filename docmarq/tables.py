# docmarq/tables.py

"""
Table rendering helpers.

Column widths come from content, HTML-style, and go out as a fixed layout:
Word's own autofit lays a table out differently in every renderer.
The rest is what Word doesn't hand us: zebra striping, header shading, borders, cell padding.
"""
from .utils import color_hex
from .metrics import text_width_mm

#------------------------------------------------------------------------------------ Column widths

# Word rounds widths to 1/20 pt and may set a glyph a hair wider than its TTF says.
# Without this room, a word given exactly its measured width could still break.
_FIT_SLACK_MM = 0.3

def content_widths(
  rows:list[list[str]], total:float, pad:float,
  family:str, size_pt:float,
  bold_rows:int = 0,
) -> list[float]:
  """
  Column widths in mm that fill `total`, sized by content the way HTML auto layout does.

  A column's min is its widest word, its max its longest line, both with `pad` a side.
  The first `bold_rows` rows measure bold, as a header does.
  """
  ncols = max(len(row) for row in rows)
  col_min, col_max = [0.0] * ncols, [0.0] * ncols
  for r, row in enumerate(rows):
    bold = r < bold_rows
    for c, cell in enumerate(row):
      for line in str(cell).split("\n"):
        col_max[c] = max(col_max[c], text_width_mm(line, family, size_pt, bold))
        for word in line.split():
          col_min[c] = max(col_min[c], text_width_mm(word, family, size_pt, bold))
  edge = 2 * pad + _FIT_SLACK_MM
  col_min = [w + edge for w in col_min]
  col_max = [w + edge for w in col_max]
  return _fit_columns(col_min, col_max, total, list(range(ncols)))

# `_fit_columns` and `_cap_widths` are vendored from `pdfmarq.md.md_table` on purpose:
# both packages lay the same markdown out alike, and a cross-lib test keeps the copies equal.

def _fit_columns(
  col_min:list[float], col_max:list[float],
  total:float, grow:list[int],
) -> list[float]:
  """
  HTML-style auto layout: widths that sum to `total`, each column between its min and max.

  Room to spare goes in equal shares to the `grow` columns, or to all when none may grow.
  A squeeze shrinks each column toward its min, in proportion to its max - min gap.
  Minimums wider than `total` go to `_cap_widths`.
  """
  if not col_max: return []
  if sum(col_max) <= total:
    grow = grow or list(range(len(col_max)))
    share = (total - sum(col_max)) / len(grow)
    return [w + share if i in grow else w for i, w in enumerate(col_max)]
  if sum(col_min) >= total: return _cap_widths(col_min, total)
  gaps = [hi - lo for lo, hi in zip(col_min, col_max)]
  scale = (total - sum(col_min)) / sum(gaps)
  return [lo + gap * scale for lo, gap in zip(col_min, gaps)]

def _cap_widths(widths:list[float], total:float) -> list[float]:
  """
  Cut the widest of `widths` down to one shared cap, so they sum to `total`.

  For column minimums that overflow the page.
  Scaling them all would break a short code like `PP-1` as hard as the URL behind the squeeze.
  A cap keeps every narrower column whole: only words too long to fit anyway get broken.
  """
  remaining = total
  left = len(widths)
  for w in sorted(widths):
    cap = remaining / left
    if w >= cap:
      return [min(x, cap) for x in widths]
    remaining -= w
    left -= 1
  return list(widths)

#----------------------------------------------------------------------------------- Border helpers

_BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")

def apply_table_borders(table, color:tuple|str, size_pt:float, sides:tuple=_BORDER_SIDES):
  """Apply uniform borders to a table.

  Args:
    table: `python-docx` `Table`.
    color: Border RGB (`(r,g,b)` 0-1 or `#hex`).
    size_pt: Border thickness in pt (Word stores as eighths-of-pt).
    sides: Which sides to draw - defaults to all + interior.
  """
  from docx.oxml.ns import qn
  from docx.oxml import OxmlElement
  tbl_pr = table._element.find(qn("w:tblPr"))
  if tbl_pr is None:
    tbl_pr = OxmlElement("w:tblPr")
    table._element.insert(0, tbl_pr)
  borders = tbl_pr.find(qn("w:tblBorders"))
  if borders is None:
    borders = OxmlElement("w:tblBorders")
    tbl_pr.append(borders)
  hex_color = color_hex(color)
  sz_eighths = max(1, int(round(size_pt * 8)))
  for side in sides:
    b = borders.find(qn(f"w:{side}"))
    if b is None:
      b = OxmlElement(f"w:{side}")
      borders.append(b)
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(sz_eighths))
    b.set(qn("w:space"), "0")
    b.set(qn("w:color"), hex_color)

def remove_table_borders(table):
  """Strip all borders from a table (set every side to `val='none'`).
  Useful for layout tables - 1xN borderless grids used to position content
  side-by-side (banner logo + text) without visible structure.
  """
  from docx.oxml.ns import qn
  from docx.oxml import OxmlElement
  tbl_pr = table._element.find(qn("w:tblPr"))
  if tbl_pr is None:
    tbl_pr = OxmlElement("w:tblPr")
    table._element.insert(0, tbl_pr)
  borders = tbl_pr.find(qn("w:tblBorders"))
  if borders is not None:
    tbl_pr.remove(borders)
  borders = OxmlElement("w:tblBorders")
  tbl_pr.append(borders)
  for side in _BORDER_SIDES:
    b = OxmlElement(f"w:{side}")
    b.set(qn("w:val"), "none")
    b.set(qn("w:sz"), "0")
    b.set(qn("w:space"), "0")
    b.set(qn("w:color"), "auto")
    borders.append(b)

def apply_cell_shading(cell, color:tuple|str):
  """Set background fill on a table cell via raw `<w:shd>` in `tcPr`."""
  from docx.oxml.ns import qn
  from docx.oxml import OxmlElement
  tc_pr = cell._tc.get_or_add_tcPr()
  shd = tc_pr.find(qn("w:shd"))
  if shd is None:
    shd = OxmlElement("w:shd")
    tc_pr.append(shd)
  shd.set(qn("w:val"), "clear")
  shd.set(qn("w:color"), "auto")
  shd.set(qn("w:fill"), color_hex(color))

def set_cell_align(cell, align:str|None):
  """Set horizontal alignment for all paragraphs in a cell."""
  from .utils import align_to_docx
  a = align_to_docx(align)
  if a is None: return
  for p in cell.paragraphs:
    p.alignment = a

def repeat_header_row(row):
  """Mark a table row as a repeating header (shown on each page)."""
  from docx.oxml.ns import qn
  from docx.oxml import OxmlElement
  tr_pr = row._tr.get_or_add_trPr()
  hdr = OxmlElement("w:tblHeader")
  hdr.set(qn("w:val"), "true")
  tr_pr.append(hdr)

#----------------------------------------------------------------------------- Cell margins / align

# OOXML cell margins use `dxa` = twentieths of a point. 1 mm = 1440/25.4 dxa.
_DXA_PER_MM = 1440 / 25.4

def set_cell_margins(cell, top:float=0, right:float=0, bot:float=0, left:float=0):
  """Set internal cell padding in mm via raw `<w:tcMar>`. Uses canonical
  `left` / `right` rather than bidi `start` / `end` for renderer
  compatibility (OnlyOffice in particular only honors the former).
  """
  from docx.oxml.ns import qn
  from docx.oxml import OxmlElement
  tc_pr = cell._tc.get_or_add_tcPr()
  mar = tc_pr.find(qn("w:tcMar"))
  if mar is None:
    mar = OxmlElement("w:tcMar")
    tc_pr.append(mar)
  for side, val in (("top", top), ("right", right), ("bottom", bot), ("left", left)):
    el = mar.find(qn(f"w:{side}"))
    if el is None:
      el = OxmlElement(f"w:{side}")
      mar.append(el)
    el.set(qn("w:w"), str(int(round(val * _DXA_PER_MM))))
    el.set(qn("w:type"), "dxa")

def set_cell_vertical_align(cell, align:str="center"):
  """Set vertical alignment of cell content. `align`: `top`/`center`/`bottom`."""
  from docx.enum.table import WD_ALIGN_VERTICAL
  va_map = {
    "top": WD_ALIGN_VERTICAL.TOP,
    "center": WD_ALIGN_VERTICAL.CENTER,
    "bottom": WD_ALIGN_VERTICAL.BOTTOM,
  }
  cell.vertical_alignment = va_map.get(align, WD_ALIGN_VERTICAL.CENTER)
