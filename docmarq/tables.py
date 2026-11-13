# docmarq/tables.py

"""
Table rendering helpers.

Column widths come from content, HTML-style, and go out as a fixed layout:
Word's own autofit lays a table out differently in every renderer.
The rest is what Word doesn't hand us: zebra striping, header shading, borders, cell padding.
"""
from dataclasses import dataclass
from typing import NamedTuple
from docx.shared import Mm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_ALIGN_VERTICAL
from .constants import DXA_PER_MM, MM_TO_PT
from .utils import color_hex, align_to_docx, eighths
from .metrics import text_width_mm

#------------------------------------------------------------------------------------ Column widths

# Word rounds widths to 1/20 pt and may set a glyph a hair wider than its TTF says.
# Without this room, a word given exactly its measured width could still break.
_FIT_SLACK_MM = 0.3

CELL_LINE_SPACING = 1.15 # line spacing of the text in a cell, in lines

# How much smaller than the text beside it a column of pictures stays, as pdfmarq's default.
_IMAGE_BALANCE_BIAS = 0.7

@dataclass
class TablePicture:
  """
  A picture a table cell holds instead of text, for `DOCX.table`.
  `source` is a path or a buffer of PNG bytes, and `width` by `height` mm its size,
  scaled down to fit a narrower column.
  `low` and `high` are the narrowest and widest it lets its column be, padding aside.
  """
  source: object
  width: float
  height: float
  low: float
  high: float

def content_widths(
  rows:list[list], total:float, pad:float,
  family:str, size_pt:float,
  bold_rows:int = 0,
) -> list[float]:
  """
  Column widths in mm that fill `total`, sized by content the way HTML auto layout does.

  A column's min is its widest word, its max its longest line, both with `pad` a side.
  The first `bold_rows` rows measure bold, as a header does.
  A `TablePicture` asks for the range it carries,
  and a column of pictures alone trades width with the text beside it, as in pdfmarq.
  """
  ncols = max(len(row) for row in rows)
  asks = [[_ask(cell, family, size_pt, r < bold_rows) for cell in row]
    for r, row in enumerate(rows)]
  line_h = size_pt * CELL_LINE_SPACING / MM_TO_PT
  # the slack splits between the two sides of a column
  widths, _ = _layout_columns(asks, ncols, total, pad + _FIT_SLACK_MM / 2, line_h,
    _IMAGE_BALANCE_BIAS)
  return widths

def words_width(
  rows:list[list], pad:float, family:str, size_pt:float, bold_rows:int = 0,
) -> float:
  """
  mm `rows` need to hold their longest words: each column's widest word, `pad` a side,
  as `content_widths` lays them out. Narrower, words break inside cells.
  """
  ncols = max(len(row) for row in rows)
  asks = [[_ask(cell, family, size_pt, r < bold_rows) for cell in row]
    for r, row in enumerate(rows)]
  return sum(max((row[i].low for row in asks if i < len(row) and not row[i].image), default=0)
    + 2 * pad + _FIT_SLACK_MM for i in range(ncols))

def _ask(cell, family:str, size_pt:float, bold:bool) -> "_Cell":
  """What `cell` asks of its column: a picture its range, text its widest word and longest line."""
  if isinstance(cell, TablePicture):
    return _Cell(cell.low, cell.high, True, cell.height / cell.width if cell.width > 0 else None)
  low = high = 0.0
  for line in str(cell).split("\n"):
    high = max(high, text_width_mm(line, family, size_pt, bold))
    for word in line.split():
      low = max(low, text_width_mm(word, family, size_pt, bold))
  return _Cell(low, high)

# `_fit_columns`, `_cap_widths` and the column layout under them
# are vendored from `pdfmarq.md.md_table` on purpose:
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
    if w >= cap: return [min(x, cap) for x in widths]
    remaining -= w
    left -= 1
  return list(widths)

class _Cell(NamedTuple):
  """
  What one cell asks of its column, in mm of content:
  `low` the narrowest it fits in, `high` the widest it would use.
  An image cell sets `image`, and `aspect`, its height over width, where it is known.
  """
  low: float
  high: float
  image: bool = False
  aspect: float|None = None

def _layout_columns(
  rows:list[list[_Cell]], ncols:int, total:float, pad:float,
  line_h:float, bias:float,
) -> tuple[list[float], list[float]]:
  """
  HTML-style auto layout of the cells of `rows`, `pad` a side: `(col_widths, text_widths)` in mm.

  Text and image contributions are tracked separately per column,
  so an image-only row cannot pin the column max below what the text rows need.
  Spare room goes to text columns only: growing an image column inflates its images.
  A squeeze shrinks image columns too, and their `low` floors them at icon scale.
  `_balance_image_cols` then trades width between a pure image column and a text one,
  `line_h` being a text line's height and `bias` how much smaller than the text an image stays.
  """
  col_min_text = [0.0] * ncols
  col_max_text = [0.0] * ncols
  col_min_img = [0.0] * ncols
  col_max_img = [0.0] * ncols
  has_text = [False] * ncols
  has_image = [False] * ncols
  for cells in rows:
    for i, cell in enumerate(cells[:ncols]):
      if cell.image:
        if cell.low > col_min_img[i]: col_min_img[i] = cell.low
        if cell.high > col_max_img[i]: col_max_img[i] = cell.high
        has_image[i] = True
      else:
        if cell.low > col_min_text[i]: col_min_text[i] = cell.low
        if cell.high > col_max_text[i]: col_max_text[i] = cell.high
        has_text[i] = True
  col_min: list[float] = []
  col_max: list[float] = []
  for i in range(ncols):
    cmin = max(col_min_img[i], col_min_text[i])
    if has_text[i]:
      cmax = max(col_max_text[i], col_min_img[i]) # text width, floored at image min
    else:
      cmax = col_max_img[i]
    col_min.append(cmin)
    col_max.append(cmax)
  col_min = [m + 2 * pad for m in col_min]
  col_max = [m + 2 * pad for m in col_max]
  text_cols = [i for i in range(ncols) if not has_image[i]]
  col_widths = _fit_columns(col_min, col_max, total, text_cols)
  text_widths = [max(1.0, cw - 2 * pad) for cw in col_widths]
  return _balance_image_cols(col_widths, text_widths, rows, has_image, has_text,
    ncols, pad, col_min, line_h, bias)

def _balance_image_cols(
  col_widths:list[float],
  text_width_mm:list[float],
  rows:list[list[_Cell]],
  is_image_col:list[bool],
  has_text:list[bool],
  ncols:int,
  h_pad:float,
  col_min:list[float],
  line_h_mm:float,
  bias:float,
) -> tuple[list[float], list[float]]:
  """
  Reflow widths so a pure-image col's height matches its widest text-col partner.
  Mixed cols (both image and text rows) are skipped -
  they need full text width and balancing would truncate text.

  Per image col `i` paired with text col `j`:
    image_h(w_i) = w_text_i × aspect_mean
    text_h(w_j)  ≈ K / w_text_j  (K = total_text_w × line_h)
  Budget constraint: w_i + w_j = const.
  Equating heights yields a quadratic in w_i.
  """
  if not any(is_image_col) or ncols < 2: return col_widths, text_width_mm
  for ic in range(ncols):
    if not is_image_col[ic]: continue
    if has_text[ic]: continue # mixed col: text rows need full width
    aspects = [row[ic].aspect for row in rows
      if ic < len(row) and row[ic].image and row[ic].aspect is not None]
    if not aspects: continue
    aspect = sum(aspects) / len(aspects)
    # Partner: widest non-image col (donor for width redistribution)
    partners = [j for j in range(ncols) if j != ic and not is_image_col[j]]
    if not partners: continue
    tc = max(partners, key=lambda j: col_widths[j])
    # K = text char-area of the widest row (total_text_w × line_h)
    K = 0.0
    for row in rows:
      if tc >= len(row) or row[tc].image: continue
      k = row[tc].high * line_h_mm
      if k > K: K = k
    if K <= 0: continue
    # bias < 1 targets image_h = text_h × bias (smaller image for short text);
    # equivalent to inflating K by 1/bias².
    K = K / (bias ** 2)
    budget = col_widths[ic] + col_widths[tc]
    pad2 = 2 * h_pad
    # Work in text widths directly: budget_text = (w_i + w_j) − 4·h_pad
    budget_text = budget - 2 * pad2
    if budget_text <= 0: continue
    # quadratic: aspect·w_i² − aspect·budget_text·w_i + K = 0
    a = aspect
    b = -aspect * budget_text
    c = K
    disc = b*b - 4*a*c
    if disc < 0:
      # Text dominates, so the image takes its minimum width.
      w_i_text = max(col_min[ic] - pad2, 1.0)
    else:
      w_i_text = (-b - disc**0.5) / (2 * a) # smaller root keeps image narrower
    img_min_text = col_min[ic] - pad2
    img_max_text = col_widths[ic] - pad2 # don't exceed linear-solver result
    img_max_text = max(img_max_text, budget_text * 0.45) # allow growth for heavy text
    w_i_text = max(img_min_text, min(w_i_text, img_max_text))
    w_j_text = budget_text - w_i_text
    partner_min_text = col_min[tc] - pad2 # partner must not shrink below its min
    if w_j_text < partner_min_text:
      w_j_text = partner_min_text
      w_i_text = budget_text - w_j_text
    col_widths[ic] = w_i_text + pad2
    col_widths[tc] = w_j_text + pad2
    text_width_mm[ic] = max(1.0, w_i_text)
    text_width_mm[tc] = max(1.0, w_j_text)
  return col_widths, text_width_mm

#----------------------------------------------------------------------------------- Border helpers

_BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")
# `w:tblPr` children that follow `w:tblBorders` (ISO/IEC 29500-1);
# strict renderers drop a child out of sequence.
_AFTER_BORDERS = ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
  "w:tblDescription", "w:tblPrChange")
# `w:tcPr` children that follow `w:tcBorders`
_AFTER_TC_BORDERS = ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
  "w:vAlign", "w:hideMark")
# sides of `w:tcBorders` in schema order, `left` and `right` the transitional names of the ends
_CELL_SIDES = ("top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br",
  "tr2bl")

def apply_table_borders(table, color:tuple|str, size_pt:float, sides:tuple=_BORDER_SIDES):
  """
  Apply uniform borders to a table.

  Args:
    table: `python-docx` `Table`.
    color: Border RGB (`(r,g,b)` 0-1 or `#hex`).
    size_pt: Border thickness in pt (Word stores as eighths-of-pt).
    sides: Which sides to draw - defaults to all + interior.
  """
  tbl_pr = table._element.find(qn("w:tblPr"))
  if tbl_pr is None:
    tbl_pr = OxmlElement("w:tblPr")
    table._element.insert(0, tbl_pr)
  borders = tbl_pr.find(qn("w:tblBorders"))
  if borders is None:
    borders = OxmlElement("w:tblBorders")
    tbl_pr.insert_element_before(borders, *_AFTER_BORDERS)
  hex_color = color_hex(color)
  sz_eighths = eighths(size_pt)
  for side in sides:
    b = borders.find(qn(f"w:{side}"))
    if b is None:
      b = OxmlElement(f"w:{side}")
      borders.append(b)
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(sz_eighths))
    b.set(qn("w:space"), "0")
    b.set(qn("w:color"), hex_color)

_RULED_SIDES = {
  "grid": _BORDER_SIDES,
  "rows": ("top", "bottom", "insideH"),
  "head": ("top", "bottom"),
}

def check_table_rules(rules:str):
  """Refuse a `rules` value no table is ruled by, before any table is written."""
  if rules not in _RULED_SIDES:
    raise ValueError(f"table rules must be one of {', '.join(_RULED_SIDES)}, got {rules!r}")

def apply_table_rules(table, color:tuple|str, size_pt:float, rules:str, double:bool):
  """
  Rule `table` as `rules` names it: `grid`, `rows` or `head`, see `TableStyle`.
  Every other side is set to none, so a table style of the template adds no rule of its own.
  Under `head` the caller rules the header row with `rule_cell`.
  `double` draws the rules above and below the table double.
  """
  check_table_rules(rules)
  remove_table_borders(table)
  apply_table_borders(table, color, size_pt, _RULED_SIDES[rules])
  if not double: return
  borders = table._element.tblPr.find(qn("w:tblBorders"))
  for side in ("top", "bottom"):
    borders.find(qn(f"w:{side}")).set(qn("w:val"), "double")

def rule_cell(cell, color:tuple|str, size_pt:float, side:str="bottom"):
  """
  Single rule on one `side` of `cell`, as under a table's header row.
  A rule already on that side gives way, and the sides keep the order of the schema.
  """
  tc_pr = cell._tc.get_or_add_tcPr()
  borders = tc_pr.find(qn("w:tcBorders"))
  if borders is None:
    borders = OxmlElement("w:tcBorders")
    tc_pr.insert_element_before(borders, *_AFTER_TC_BORDERS)
  old = borders.find(qn(f"w:{side}"))
  if old is not None: borders.remove(old)
  size = str(eighths(size_pt))
  rule = OxmlElement(f"w:{side}", {qn("w:val"): "single", qn("w:sz"): size,
    qn("w:space"): "0", qn("w:color"): color_hex(color)})
  later = {qn(f"w:{name}") for name in _CELL_SIDES[_CELL_SIDES.index(side) + 1:]}
  following = next((child for child in borders if child.tag in later), None)
  if following is None: borders.append(rule)
  else: following.addprevious(rule)

def remove_table_borders(table):
  """
  Strip all borders from a table (set every side to `val='none'`).
  Useful for layout tables - 1xN borderless grids
  used to position content side-by-side (banner logo + text) without visible structure.
  """
  tbl_pr = table._element.find(qn("w:tblPr"))
  if tbl_pr is None:
    tbl_pr = OxmlElement("w:tblPr")
    table._element.insert(0, tbl_pr)
  borders = tbl_pr.find(qn("w:tblBorders"))
  if borders is not None:
    tbl_pr.remove(borders)
  borders = OxmlElement("w:tblBorders")
  tbl_pr.insert_element_before(borders, *_AFTER_BORDERS)
  for side in _BORDER_SIDES:
    b = OxmlElement(f"w:{side}")
    b.set(qn("w:val"), "none")
    b.set(qn("w:sz"), "0")
    b.set(qn("w:space"), "0")
    b.set(qn("w:color"), "auto")
    borders.append(b)

def apply_cell_shading(cell, color:tuple|str):
  """Set background fill on a table cell via raw `<w:shd>` in `tcPr`."""
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
  a = align_to_docx(align)
  if a is None: return
  for p in cell.paragraphs:
    p.alignment = a

def repeat_header_row(row):
  """Mark a table row as a repeating header (shown on each page)."""
  tr_pr = row._tr.get_or_add_trPr()
  hdr = OxmlElement("w:tblHeader")
  hdr.set(qn("w:val"), "true")
  tr_pr.append(hdr)

# `w:trPr` children that follow `w:cantSplit`
_AFTER_CANT_SPLIT = ("w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden")

def keep_rows(table, ends:int=0):
  """
  Each row of `table` on one page, and a break between pages or columns kept from its ends:
  the first `ends + 1` rows go together, as a header and `ends` rows of the body,
  and so do the last `ends`.
  """
  rows = table.rows
  last = len(rows) - 1
  for k, row in enumerate(rows):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
      tr_pr.insert_element_before(OxmlElement("w:cantSplit"), *_AFTER_CANT_SPLIT)
    if k < ends or last - ends < k < last:
      for cell in row.cells:
        for p in cell.paragraphs: p.paragraph_format.keep_with_next = True

#----------------------------------------------------------------------------- Cell margins / align

def set_cell_margins(cell, top:float=0, right:float=0, bot:float=0, left:float=0):
  """
  Set internal cell padding in mm via raw `<w:tcMar>`.
  Uses canonical `left` / `right` rather than bidi `start` / `end` for renderer compatibility
  (OnlyOffice in particular only honors the former).
  """
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
    el.set(qn("w:w"), str(int(round(val * DXA_PER_MM))))
    el.set(qn("w:type"), "dxa")

def set_cell_vertical_align(cell, align:str="center"):
  """Set vertical alignment of cell content. `align`: `top`/`center`/`bottom`."""
  va_map = {
    "top": WD_ALIGN_VERTICAL.TOP,
    "center": WD_ALIGN_VERTICAL.CENTER,
    "bottom": WD_ALIGN_VERTICAL.BOTTOM,
  }
  cell.vertical_alignment = va_map.get(align, WD_ALIGN_VERTICAL.CENTER)

#-------------------------------------------------------------------------------------- Table width

def set_table_widths(tbl, col_widths_mm:list[float], left_pad_mm:float=0):
  """
  Anchor a table's total/per-column widths and left edge via raw XML.

  `<w:tblPr>` children must be in canonical schema order
  (`tblW → jc → tblInd → tblBorders → tblLayout → tblCellMar → tblLook`);
  strict renderers (OnlyOffice) fall back to defaults on out-of-order children.

  `jc=left` is explicit because `tblInd` is measured from the leading edge, which `jc` defines.
  `tblInd` equals the per-cell left padding (`tcMar/left`):
  renderers anchor the table by cell content,
  pulling the border left of the margin by that padding,
  so an equal `tblInd` shifts it back onto the margin.
  `tblCellMar` stays 0 so only the per-cell margin participates.
  """
  total_dxa = int(round(sum(col_widths_mm) * DXA_PER_MM))
  left_dxa = int(round(left_pad_mm * DXA_PER_MM))
  tbl_el = tbl._element
  tbl_pr = tbl_el.find(qn("w:tblPr"))
  # Drop the children we manage;
  # keep refs to borders/look first so they re-insert at the right schema slot.
  preserved_borders = tbl_pr.find(qn("w:tblBorders"))
  preserved_look = tbl_pr.find(qn("w:tblLook"))
  for tag in ("tblW", "jc", "tblInd", "tblBorders", "tblLayout",
      "tblCellMar", "tblLook"):
    for el in tbl_pr.findall(qn(f"w:{tag}")):
      tbl_pr.remove(el)
  # Build in canonical schema order:
  # `tblW`, `jc`, `tblInd`, [borders], `tblLayout`, `tblCellMar`, [look] -
  # borders/look carried over from the preserved refs.
  tbl_w = OxmlElement("w:tblW")
  tbl_w.set(qn("w:w"), str(total_dxa))
  tbl_w.set(qn("w:type"), "dxa")
  jc = OxmlElement("w:jc") # left - defines the `tblInd` leading edge
  jc.set(qn("w:val"), "left")
  tbl_ind = OxmlElement("w:tblInd") # match per-cell left padding (see docstring)
  tbl_ind.set(qn("w:w"), str(left_dxa))
  tbl_ind.set(qn("w:type"), "dxa")
  tbl_layout = OxmlElement("w:tblLayout")
  tbl_layout.set(qn("w:type"), "fixed")
  tbl_cell_mar = OxmlElement("w:tblCellMar") # 0 on all sides, per-cell `tcMar` owns the padding
  for side in ("top", "left", "bottom", "right"):
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:w"), "0")
    el.set(qn("w:type"), "dxa")
    tbl_cell_mar.append(el)
  ordered = [tbl_w, jc, tbl_ind]
  if preserved_borders is not None:
    ordered.append(preserved_borders)
  ordered.append(tbl_layout)
  ordered.append(tbl_cell_mar)
  if preserved_look is not None:
    ordered.append(preserved_look)
  for el in ordered:
    tbl_pr.append(el)
  # Replace grid columns
  old_grid = tbl_el.find(qn("w:tblGrid"))
  if old_grid is not None:
    tbl_el.remove(old_grid)
  new_grid = OxmlElement("w:tblGrid")
  for w_mm in col_widths_mm:
    gc = OxmlElement("w:gridCol")
    gc.set(qn("w:w"), str(int(round(w_mm * DXA_PER_MM))))
    new_grid.append(gc)
  # Insert `tblGrid` right after `tblPr` (OOXML schema requires this order).
  tbl_el.insert(list(tbl_el).index(tbl_pr) + 1, new_grid)
  # Also set each cell width so cell-level overrides are consistent.
  for row in tbl.rows:
    for i, w_mm in enumerate(col_widths_mm):
      if i < len(row.cells):
        row.cells[i].width = Mm(w_mm)
