# docmarq/paper/banner.py

"""
Banner of a documentation DOCX, from frontmatter: the masthead on page 1 and the compact header.

The masthead is document content: issuer, title, status badge, version, id, author and dates,
beside the logo when frontmatter names one.
The compact header carries the id, or the title, on the pages after it.
"""
import warnings
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.image.exceptions import UnrecognizedImageError
from ..core import DOCX, _set_table_widths
from ..inline import _apply_run_shading
from ..tables import remove_table_borders, set_cell_margins, set_cell_vertical_align
from ..utils import color_hex, rgb255, tight_line_height_pt, format_date, eighths
from ..md import image_utils
from .style import PaperStyle, UNKNOWN_STATUS

_SEP = "  •  " # U+2022 BULLET between the parts of a line, more visible than a middle dot
_LOGO_COL_MM = 30 # leaves room for square-ish logos
_COMPACT_SIZE_PT = 9
_COMPACT_COLOR = (0x66, 0x6E, 0x77)

#------------------------------------------------------------------------------------------- Banner

class Banner:
  """Masthead of one document, from its frontmatter `data` and resolved `logo` path."""
  def __init__(self, doc:DOCX, style:PaperStyle, data:dict, logo:str|None):
    self.doc = doc
    self.style = style
    self.data = data
    self.logo = logo

  def draw(self):
    """Masthead after the last block, then a rule under it."""
    if self.logo: self._table()
    else: self._flow()
    self.doc.hr(color=self.style.hr_color, size_eighths=eighths(self.style.hr_thick))

  def _flow(self):
    """Masthead with no logo, one paragraph per line through the fluent `DOCX` API."""
    doc = self.doc
    def new_para(*, space_before:float=0, space_after:float=0, exact_pt:float|None=None):
      doc.para()
      p = doc._current_para
      _space(p, space_before, space_after, exact_pt)
      doc._keep_words(p)
    def add_run(text, *, bold=False, color=None, bg=None, size=None):
      doc.text(text, bold=bold, color=color, bg_color=bg, size=size)
    self._content(new_para, add_run)

  def _table(self):
    """Masthead beside the logo: a borderless 1×2 table, logo left, text right."""
    doc = self.doc
    doc._flush_para()
    tbl = doc._new_table(1, 2)
    _set_table_widths(tbl, [_LOGO_COL_MM, doc.content_width - _LOGO_COL_MM])
    remove_table_borders(tbl)
    left, right = tbl.rows[0].cells
    set_cell_margins(left, top=0, right=4, bot=0, left=0)
    set_cell_vertical_align(left, "center")
    self._logo_cell(left, _LOGO_COL_MM - 4)
    set_cell_margins(right, top=0, right=0, bot=0, left=2)
    set_cell_vertical_align(right, "center")
    self._text_cell(right)

  def _logo_cell(self, cell, max_w:float):
    """
    The logo in `cell`, `max_w` mm wide.
    Pillow crops its transparent margins first, which would shift what shows off the cell edge.
    """
    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    run = p.add_run()
    buf = image_utils.preprocess_to_buffer(self.logo)
    try:
      run.add_picture(buf if buf is not None else self.logo, width=Mm(max_w))
      image_utils.pin_offsets(run)
    except (OSError, ValueError, UnrecognizedImageError) as e:
      warnings.warn(
        f"frontmatter logo {self.logo!r} could not be embedded, banner renders without it: {e}",
        RuntimeWarning, stacklevel=2,
      )

  def _text_cell(self, cell):
    """Masthead text in `cell`, written into its paragraphs directly, at Word's default leading."""
    p_now = []
    def new_para(*, space_before:float=0, space_after:float=0, exact_pt:float|None=None):
      if p_now:
        p = cell.add_paragraph()
      else:
        p = cell.paragraphs[0]
        p.text = ""
      _space(p, space_before, space_after, exact_pt)
      self.doc._keep_words(p)
      p_now[:] = [p]
    def add_run(text, *, bold=False, color=None, bg=None, size=None):
      run = p_now[0].add_run(text)
      if bold: run.bold = True
      if size is not None: run.font.size = Pt(size)
      if color is not None: run.font.color.rgb = RGBColor(*rgb255(color))
      if bg is not None: _apply_run_shading(run, color_hex(bg))
    self._content(new_para, add_run)

  def _content(self, new_para, add_run):
    """
    The masthead lines, the same in both layouts.
    `new_para(space_before=, space_after=, exact_pt=)` opens a line,
    `add_run(text, bold=, color=, bg=, size=)` writes into it.
    """
    s, fm = self.style, self.data
    # Issuer first, the name bold and the address muted, on one line.
    entity = fm.get("entity")
    address = fm.get("address")
    if entity or address:
      new_para(space_after=1)
      if entity:
        add_run(str(entity), bold=True, size=s.banner_meta_size)
        if address:
          add_run(_SEP, color=s.muted_color, size=s.banner_meta_size)
      if address:
        add_run(str(address), size=s.banner_meta_size, color=s.muted_color)
    title = fm.get("title")
    if title:
      # `EXACTLY` bypasses Calibri's ~1.2× internal leading,
      # so a title that wraps keeps its lines tight.
      new_para(exact_pt=tight_line_height_pt(s.banner_title_size))
      add_run(str(title), bold=True, size=s.banner_title_size)
    status = str(fm.get("status") or "").strip().lower()
    version = fm.get("version")
    id_ = fm.get("id")
    if status or version or id_:
      new_para(space_before=2)
      if status:
        bg, fg = s.banner_status_colors.get(status, UNKNOWN_STATUS)
        add_run(f" {status.upper()} ", bold=True, color=fg, bg=bg)
        if version or id_:
          add_run(_SEP, color=s.muted_color)
      if version:
        add_run(str(version), color=s.muted_color)
        if id_:
          add_run(_SEP, color=s.muted_color)
      if id_:
        add_run(str(id_), color=s.muted_color)
    author = fm.get("author")
    if author:
      new_para()
      add_run(f"{s.banner_label_author}: {author}", size=s.banner_meta_size, color=s.muted_color)
    created = fm.get("created")
    updated = fm.get("updated")
    if created or updated:
      new_para()
      parts = []
      if created:
        parts.append(f"{s.banner_label_created}: {format_date(created, s.date_format)}")
      if updated:
        parts.append(f"{s.banner_label_updated}: {format_date(updated, s.date_format)}")
      add_run(_SEP.join(parts), size=s.banner_meta_size, color=s.muted_color)

#------------------------------------------------------------------------------------------ Helpers

def _space(p, before:float, after:float, exact_pt:float|None):
  """Spacing of a masthead line in pt: `before`, `after`, and an exact leading when given."""
  pf = p.paragraph_format
  pf.space_before = Pt(before)
  pf.space_after = Pt(after)
  if exact_pt is not None:
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(exact_pt)

def compact(p, text:str):
  """Header paragraph `p` set to `text`: centred, muted grey, small."""
  p.text = ""
  p.alignment = WD_ALIGN_PARAGRAPH.CENTER
  run = p.add_run(text)
  run.font.size = Pt(_COMPACT_SIZE_PT)
  run.font.color.rgb = RGBColor(*_COMPACT_COLOR)
