# docmarq/core.py

"""Core DOCX class - main facade for document generation."""
import copy, warnings
from contextlib import contextmanager
from dataclasses import dataclass
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_BREAK, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .constants import Defaults
from .utils import (to_mm, parse_margin, align_to_docx, color_hex,
  rgb255, smaller_size, tight_line_height)
from .styles import Style, TableStyle
from .layout import PageGeometry
from .structure import Metadata
from .inline import RichSegment, _apply_run_format
from .svg import svg_to_png_buffer, is_svg
from .fields import page_fields
from .tables import (
  apply_table_borders, apply_cell_shading, set_cell_align, repeat_header_row,
  set_cell_margins, set_cell_vertical_align, content_widths,
)

# Language tags Word names a language by, for the codes of the presets.
_LANG_TAGS = {"en": "en-US", "pl": "pl-PL", "de": "de-DE", "fr": "fr-FR", "es": "es-ES",
  "it": "it-IT", "cs": "cs-CZ", "sk": "sk-SK"}

#------------------------------------------------------------------------------------------- Target

@dataclass
class _Target:
  """Where blocks go: the document body, or a table cell."""
  box: object # `Document` or `_Cell`, both with `add_paragraph` and `add_table`
  width: float|None = None # mm, a cell's; `None` follows the page and its columns
  spare: object = None # empty paragraph the next block fills, as a fresh cell's first

#--------------------------------------------------------------------------------------------- DOCX

class DOCX:
  """Main DOCX generator with fluent API.

  Example:
    >>> with DOCX("out.docx") as doc:
    ...   doc.heading("Tytuł", level=1)
    ...   doc.para("Pierwszy akapit.")
    ...   doc.text("Drugi z ").text("bold", bold=True).text(" fragmentem.")
  """
  def __init__(
    self,
    path:str,
    width:float = Defaults.PAGE_WIDTH,
    height:float = Defaults.PAGE_HEIGHT,
    margin:float|tuple = Defaults.MARGIN,
    unit:str = Defaults.UNIT,
    template:str|None = None,
    neutral_style:bool = True,
    gutter:float = 0,
    font_body:str = Defaults.FONT_FAMILY,
    font_head:str|None = None,
    body_size:float = Defaults.FONT_SIZE,
  ):
    self.path = path
    self.unit = unit
    top, right, bot, left = parse_margin(margin)
    self._page = PageGeometry(
      width=to_mm(width, unit), height=to_mm(height, unit),
      margin_top=to_mm(top, unit), margin_right=to_mm(right, unit),
      margin_bot=to_mm(bot, unit), margin_left=to_mm(left, unit),
      gutter=to_mm(gutter, unit),
    )
    self._doc = Document(template) if template else Document()
    # a justified line a manual break ends stays flush left, as in pdfmarq
    _compat_option(self._doc.settings.element, "doNotExpandShiftReturn")
    self._apply_section_geometry()
    self._style = Style().with_defaults()
    self._style.font_family = font_body
    self._style.font_size = body_size
    self._metadata = Metadata()
    self._current_para = None # open paragraph; runs append here
    self._num_ids: dict[int, int|None] = {} # list level → numbering instance
    self._para_default_color = None # per-paragraph color override; reset on flush
    self._last_space_after = 0 # pt - previous block's `space_after` (margin collapse)
    self._extra_block_after = 0 # pt - one-shot `space_before` bonus (see `_track_block_spacing`)
    self._bookmark_id = 0
    self._at = _Target(self._doc)
    self._cols = (1, 0.0) # columns of the current section: count, gap (mm)
    self._hyphenate = False # set by `hyphenation()`: headings, cells and code keep words whole
    # `False` keeps the template's own styles: the markdown engine leaves its headings alone
    self.neutral_style = neutral_style
    if neutral_style:
      self._override_heading_styles(font_body, font_head or font_body)

  #-------------------------------------------------------------------------------- Context manager

  def __enter__(self) -> "DOCX":
    return self

  def __exit__(self, exc_type, exc_val, exc_tb):
    if exc_type is None:
      self.save()

  #------------------------------------------------------------------------------------- Properties

  @property
  def page_width(self) -> float:
    """Current section's page width in mm."""
    return self._page.width

  @property
  def page_height(self) -> float:
    """Current section's page height in mm."""
    return self._page.height

  @property
  def content_width(self) -> float:
    """Width in mm blocks are laid out in: a column of the text area, or the cell `into` fills."""
    if self._at.width is not None: return self._at.width
    count, gap = self._cols
    return (self._page.content_width - gap * (count - 1)) / count

  @property
  def doc(self):
    """
    Direct `python-docx` `Document` access for advanced operations not exposed by the fluent API.
    Bypasses fluent state tracking.
    """
    return self._doc

  @property
  def output_path(self) -> str:
    """Destination file path. Symmetric with `PDF.output_path`."""
    return self.path

  #--------------------------------------------------------------------------- Section / page setup

  def _override_heading_styles(self, font_body:str, font_head:str):
    """Recolor Word's built-in `Heading 1..6` to the GitHub-light palette:
    near-black text, thin bottom rule on h1/h2.
    `font.color.rgb` pins the color so it holds regardless of the theme.

    Families arrive as parameters
    so `MarkdownStyle.font_body` and `font_head` reach Word's `Normal` and `Heading n` styles.
    """
    near_black = RGBColor(*rgb255(Defaults.HEAD_COLOR))
    # Style-level spacing defaults.
    # `heading()` recomputes `space_before` per-paragraph (margin collapse);
    # these cover headings made outside it.
    spacing_pt = self._HEADING_SPACING_PT
    for i, size_pt in enumerate(Defaults.HEAD_SIZES, start=1):
      try:
        st = self._doc.styles[f"Heading {i}"]
      except KeyError:
        continue
      st.font.color.rgb = near_black
      st.font.name = font_head
      # `st.font.name` sets only `ascii`/`hAnsi` and leaves the theme refs,
      # which resolve to Calibri Light;
      # `_force_font` strips them so the body font wins.
      _force_font(st, font_head)
      before_pt, after_pt = spacing_pt[i-1]
      st.paragraph_format.space_before = Pt(before_pt)
      st.paragraph_format.space_after = Pt(after_pt)
      self.heading_style(i, size=size_pt, bold=True, rule=i in Defaults.HEAD_UNDERLINE_LEVELS)
    # Pin `Normal` body font too, across all script ranges.
    try:
      _force_font(self._doc.styles["Normal"], font_body)
    except KeyError:
      pass

  def heading_style(
    self, level:int,
    size:float|None = None,
    bold:bool|None = None,
    italic:bool|None = None,
    color:tuple|str|None = None,
    rule:bool|None = None,
  ) -> "DOCX":
    """
    Restyle Word's `Heading {level}`, so every heading of that level follows.

    `size` in pt, `bold` and `italic` the face, `color` its text, `rule` the thin line under it.
    `None` leaves that property as the document has it.
    """
    try:
      st = self._doc.styles[f"Heading {level}"]
    except KeyError:
      return self
    if size is not None:
      st.font.size = Pt(size)
      # Tighter leading for display sizes: 1.15 at body, → 1.0 at 24pt+.
      st.paragraph_format.line_spacing = tight_line_height(size)
    if bold is not None: st.font.bold = bold
    if italic is not None: st.font.italic = italic
    if color is not None: st.font.color.rgb = RGBColor(*rgb255(color))
    if rule: _set_pbdr(st.element, color_hex(Defaults.RULE_COLOR), sides=("bottom",))
    elif rule is False: _clear_border(st.element, "bottom")
    return self

  def _apply_section_geometry(self):
    """Apply current `PageGeometry` to every section: columns change sections, not the page."""
    for sec in self._doc.sections:
      sec.page_width = Mm(self._page.width)
      sec.page_height = Mm(self._page.height)
      sec.top_margin = Mm(self._page.margin_top)
      sec.right_margin = Mm(self._page.margin_right)
      sec.bottom_margin = Mm(self._page.margin_bot)
      sec.left_margin = Mm(self._page.margin_left)
      sec.header_distance = Mm(self._page.effective_header_dist())
      sec.footer_distance = Mm(self._page.effective_footer_dist())
      if self._page.gutter:
        sec.gutter = Mm(self._page.gutter)

  def page(self, width:float|None=None, height:float|None=None) -> "DOCX":
    """Update the page size, of every section alike."""
    if width: self._page.width = to_mm(width, self.unit)
    if height: self._page.height = to_mm(height, self.unit)
    self._apply_section_geometry()
    return self

  def margin(
    self, top:float|None = None, right:float|None = None,
    bot:float|None = None, left:float|None = None,
  ) -> "DOCX":
    """Update the margins, of every section alike."""
    if top is not None: self._page.margin_top = to_mm(top, self.unit)
    if right is not None: self._page.margin_right = to_mm(right, self.unit)
    if bot is not None: self._page.margin_bot = to_mm(bot, self.unit)
    if left is not None: self._page.margin_left = to_mm(left, self.unit)
    self._apply_section_geometry()
    return self

  def landscape(self) -> "DOCX":
    """Swap page width and height, of every section alike."""
    self._page.width, self._page.height = self._page.height, self._page.width
    self._apply_section_geometry()
    return self

  def columns(self, count:int, gap:float=8) -> "DOCX":
    """
    Set the blocks that follow in `count` columns `gap` apart; `columns(1)` returns to one.

    Each change starts a continuous section, so the text runs on down the same page.
    Word balances the columns of a section that such a break ends, so they finish level.
    """
    if not isinstance(count, int) or count < 1:
      raise ValueError(f"columns must be a whole number from 1 up, got {count!r}")
    if self._at.width is not None:
      raise RuntimeError("columns() cannot change inside a table cell")
    gap_mm = to_mm(gap, self.unit) if count > 1 else 0.0
    if (count, gap_mm) == self._cols: return self
    self._flush_para()
    sect = self._new_section()
    cols = sect.find(qn("w:cols"))
    if cols is None:
      cols = OxmlElement("w:cols")
      sect.insert_element_before(cols, *_SECT_AFTER_COLS)
    cols.set(qn("w:num"), str(count))
    if count > 1: cols.set(qn("w:space"), str(round(gap_mm * _DXA_PER_MM)))
    self._cols = (count, gap_mm)
    return self

  def _new_section(self):
    """
    End the section so far at the last paragraph and open a continuous one after it.
    Returns the `w:sectPr` of the new section, the last in the body.

    The new section takes its headers and footers from the one before,
    and never a first-page variant: it does not start a page.
    """
    last = self._doc.element.body.get_or_add_sectPr()
    p = last.getprevious()
    if p is None or p.tag != qn("w:p") or p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
      p = self._hairline()._p
    p.set_sectPr(last.clone())
    for ref in last.xpath("w:headerReference|w:footerReference"):
      last.remove(ref)
    last.titlePg_val = False
    last.start_type = WD_SECTION.CONTINUOUS
    return last

  def _hairline(self):
    """
    Empty paragraph 1pt high after the last block:
    it ends a section where the last block is no paragraph, keeps two tables apart,
    or carries a break or a floating table's anchor.
    """
    self._flush_para()
    p = self._new_para()
    _hairline_format(p)
    self._track_block_spacing(0)
    return p

  def _chain(self, write):
    """
    Call `write()`, then chain the paragraphs it wrote, those in its tables too,
    with keep-with-next, all but the last, so Word keeps them on one page.
    """
    self._flush_para()
    body = self._at.box is self._doc
    box = self._doc.element.body if body else self._at.box._tc
    # blocks go before the body's closing `w:sectPr`, or after a cell's last block
    tail = 1 if body else 0
    children = list(box)
    spare = self._at.spare
    start = children.index(spare._p) if spare is not None else len(children) - tail
    write()
    self._flush_para()
    children = list(box)
    paras = [p for el in children[start:len(children) - tail] for p in el.iter(qn("w:p"))]
    for p in paras[:-1]:
      p.get_or_add_pPr().keepNext_val = True

  #------------------------------------------------------------------------------------------ Style

  def font(
    self, family:str|None = None, size:float|None = None,
    mode:str|None = None,
  ) -> "DOCX":
    """Set default font for subsequent runs.

    `mode` takes the modes of `pdfmarq`: `Regular`, `Bold`, `Italic` or `BoldItalic`.
    It sets the `bold` and `italic` flags.
    """
    if family: self._style.font_family = family
    if size: self._style.font_size = size
    if mode:
      m = mode.lower()
      self._style.bold = "bold" in m
      self._style.italic = "italic" in m or "oblique" in m
    return self

  def color(self, color:tuple|str|None) -> "DOCX":
    """Set default text color for subsequent runs."""
    self._style.color = color
    return self

  def style(self, **overrides) -> "DOCX":
    """Update multiple style fields at once."""
    self._style = self._style.copy(**overrides)
    return self

  def hyphenation(self, lang:str) -> "DOCX":
    """
    Split words at line ends by the hyphenation rules of `lang`, as `pl`.

    Word hyphenates text by its language, so `lang` becomes the language of the whole document.
    Headings, cells of `table()` and code written after the call keep their words whole,
    so it comes before the blocks.
    """
    settings = self._doc.settings.element
    if settings.find(qn("w:autoHyphenation")) is None:
      _insert_ordered(settings, OxmlElement("w:autoHyphenation"), _SETTINGS_ORDER)
    defaults = self._doc.styles.element.find(qn("w:docDefaults"))
    if defaults is None:
      defaults = OxmlElement("w:docDefaults")
      self._doc.styles.element.insert(0, defaults)
    rpr = _child(_child(defaults, "w:rPrDefault", first=True), "w:rPr")
    _child(rpr, "w:lang", after=_RPR_AFTER_LANG).set(qn("w:val"), _lang_tag(lang))
    for i in range(1, 7):
      try:
        _no_hyphens(self._doc.styles[f"Heading {i}"].element)
      except KeyError:
        continue
    self._hyphenate = True
    return self

  def _keep_words(self, p):
    """Keep the words of paragraph `p` whole under `hyphenation()`; without it there is no need."""
    if self._hyphenate: _no_hyphens(p._p)

  #----------------------------------------------------------------------------------- Write target

  @contextmanager
  def into(self, cell, width:float):
    """
    Write the blocks of the `with` body into table `cell`, `width` wide, then go back.

    The cell's width stands in for `content_width`, so tables and images fit inside it.
    The first block fills the cell's empty paragraph, and spacing starts afresh.
    """
    self._flush_para()
    outer = (self._at, self._last_space_after, self._extra_block_after)
    paras = cell.paragraphs
    spare = paras[-1] if paras and _is_empty(paras[-1]) else None
    self._at = _Target(cell, to_mm(width, self.unit), spare)
    self._last_space_after = self._extra_block_after = 0
    try:
      yield self
    finally:
      self._flush_para()
      # the empty paragraph a cell keeps, as Word wants after a table, adds no height
      if self._at.spare is not None: _hairline_format(self._at.spare)
      self._at, self._last_space_after, self._extra_block_after = outer

  def _new_para(self, text:str="", style:str|None=None):
    """New paragraph at the end of the write target, in its spare empty one when it has one."""
    at = self._at
    if at.spare is None: return at.box.add_paragraph(text, style=style)
    p, at.spare = at.spare, None
    if p._p.pPr is not None: p._p.remove(p._p.pPr) # a spare formatted as a hairline
    if text: p.add_run(text)
    if style is not None: p.style = style
    return p

  def _new_table(self, rows:int, cols:int):
    """
    New table at the end of the write target, a hairline apart from a table just before it.

    In a cell, a table ahead of every other block takes the place of the empty paragraph,
    and the paragraph Word wants after a table becomes the spare.
    """
    at = self._at
    if at.box is self._doc:
      # Word joins two tables with nothing between them into one
      if _last_block(self._doc).tag == qn("w:tbl"): self._hairline()
      return self._doc.add_table(rows=rows, cols=cols)
    if at.spare is not None and _first_block(at.spare._p):
      at.spare._p.getparent().remove(at.spare._p)
    table = at.box.add_table(rows, cols)
    at.spare = at.box.paragraphs[-1]
    return table

  #------------------------------------------------------------------------------------- Paragraphs

  def para(
    self, text:str|None = None, style:str|None = None,
    align:str|None = None,
  ) -> "DOCX":
    """Start a new paragraph. Closes any previous active paragraph.

    Args:
      text: Optional initial text run.
      style: Built-in Word style name (e.g. `Quote`, `Intense Quote`).
      align: `L` / `R` / `C` / `J`.
    """
    self._flush_para()
    p = self._new_para(style=style)
    if align:
      p.alignment = align_to_docx(align)
    self._apply_para_spacing(p)
    self._current_para = p
    if text:
      self._add_run(p, RichSegment(text=text))
    return self

  def text(self, content:str, **seg_kwargs) -> "DOCX":
    """Add a styled run to the current paragraph.
    Auto-creates a paragraph if none active.
    Kwargs match `RichSegment` fields (`bold`, `italic`, `code`, `color`, `link_url`, etc.).
    """
    if self._current_para is None:
      self.para()
    seg = RichSegment(text=content, **seg_kwargs)
    self._add_run(self._current_para, seg)
    return self

  def link(self, content:str, url:str|None=None, target:str|None=None) -> "DOCX":
    """Add a hyperlink run.
    Either `url` for external link or `target` for internal bookmark name."""
    if self._current_para is None:
      self.para()
    self._add_hyperlink(self._current_para, content, url=url, target=target)
    return self

  def enter(self) -> "DOCX":
    """End the current paragraph - next `text()` starts a new one."""
    self._flush_para()
    return self

  def line_break(self) -> "DOCX":
    """Insert a soft line break inside the current paragraph (shift+enter)."""
    if self._current_para is None:
      self.para()
    run = self._current_para.add_run()
    run.add_break(WD_BREAK.LINE)
    return self

  def border(
    self, color:tuple|str,
    sides:tuple = ("top", "left", "bottom", "right"),
    size:float = 0.5,
    space:float = 1,
  ) -> "DOCX":
    """
    Rule the current paragraph on `sides`, `size` pt thick, `space` pt clear of its text.
    Word keeps the width in eighths of a point and the space in whole points.
    """
    if self._current_para is None:
      self.para()
    _set_pbdr(self._current_para._p, color_hex(color), sides=sides,
      size_eighths=max(1, round(size * 8)), space=round(space))
    return self

  def _flush_para(self):
    """Close the active paragraph - block elements must call this first."""
    self._current_para = None
    self._para_default_color = None

  def _apply_para_spacing(self, p):
    """Apply line-height and before/after spacing with CSS-style margin collapse.
    Word SUMS adjacent paragraph margins (CSS takes the max),
    so `space_before` is shrunk by the previous block's `space_after`
    to land a gap of `max(prev_after, declared_before)`.
    See `_collapsed_before`.
    """
    pf = p.paragraph_format
    if self._style.line_height is not None:
      pf.line_spacing = self._style.line_height
    declared_before = self._style.space_before or 0
    declared_after = self._style.space_after or 0
    pf.space_before = Pt(self._collapsed_before(declared_before))
    pf.space_after = Pt(declared_after)
    self._last_space_after = declared_after

  def _collapsed_before(self, declared_before:float) -> float:
    """Compute effective `space_before`:
    collapse against previous block's `space_after`, then add any pending `extra_block_after`.
    Consuming this helper resets `_extra_block_after` so the bonus applies once only.
    """
    val = max(0, declared_before - self._last_space_after) + self._extra_block_after
    self._extra_block_after = 0
    return val

  def _track_block_spacing(self, after_pt:float, extra_for_next:float=0):
    """Record this block's `space_after` (for collapse against the next paragraph)
    plus an optional one-shot `space_before` bonus for it.

    Tables and lists set `extra_for_next`: their inner paragraphs use `space_after=0`,
    so nothing else would give the next block breathing room.
    """
    self._last_space_after = after_pt
    if extra_for_next:
      self._extra_block_after = extra_for_next

  @staticmethod
  def _style_owns_font(p) -> bool:
    """True when the paragraph style carries its own font.

    Runs appended to such a paragraph inherit family, size and color from the style;
    writing body values onto them would show a heading in body type.
    """
    try:
      name = p.style.name or ""
    except Exception:
      return False
    return name.startswith("Heading") or name in ("Title", "Subtitle")

  def _add_run(self, p, seg:RichSegment):
    """Add a run with `seg` styling, falling back to current `_style`."""
    if seg.break_line:
      p.add_run().add_break(WD_BREAK.LINE)
    if seg.link_url or seg.link_target:
      # `w:hyperlink` wraps its own run, so the link path replaces the plain run entirely.
      self._add_hyperlink(p, seg.text, url=seg.link_url, target=seg.link_target)
      return
    run = p.add_run(seg.text)
    # Effective family/size: segment wins, else current style.
    # A styled paragraph (heading) contributes neither, so its own font survives.
    inherited = self._style_owns_font(p)
    family = seg.family or (None if inherited else self._style.font_family)
    size = seg.size or (None if inherited else self._style.font_size)
    # Color: the segment's own, else the paragraph's default set by `blockquote()` and the like,
    # else the current `_style.color`.
    if seg.color is not None:
      eff_color = seg.color
    elif inherited:
      eff_color = None
    elif self._para_default_color is not None:
      eff_color = self._para_default_color
    else:
      eff_color = self._style.color
    # Style-level bold/italic OR-merged with segment-level
    merged = RichSegment(
      text=seg.text,
      family=family, size=size,
      color=eff_color,
      bold=seg.bold or bool(self._style.bold),
      italic=seg.italic or bool(self._style.italic),
      underline=seg.underline or bool(self._style.underline),
      strike=seg.strike or bool(self._style.strike),
      code=seg.code, highlight=seg.highlight or self._style.highlight,
      bg_color=seg.bg_color,
      superscript=seg.superscript,
      subscript=seg.subscript,
      small_caps=seg.small_caps,
    )
    _apply_run_format(run, merged, family, size,
      code_family=self._style.code_family, code_size=self._style.code_size,
      code_color=self._style.code_color, code_bg=self._style.code_bg)

  #--------------------------------------------------------------------------------------- Headings

  # Per-level (`space_before`, `space_after`) in pt.
  # Shared by the style defaults in `_override_heading_styles`
  # and the margin-collapse math in `heading()`.
  _HEADING_SPACING_PT = [(10, 4), (8, 3), (6, 3), (4, 2), (3, 2), (3, 2)]

  def heading(
    self, text:str|None = None, level:int = 1,
    space_before:float|None = None,
    space_after:float|None = None,
  ) -> "DOCX":
    """Add a heading using Word's built-in `Heading 1..9` styles (TOC-aware).
    Leaves `_current_para` set so `text()` can append rich inline runs.
    `space_before` and `space_after` in pt replace the gaps of the level.
    `space_before` is recomputed for margin collapse against the previous block.
    """
    self._flush_para()
    level = max(1, min(9, level))
    p = self._new_para(text or "", style=f"Heading {level}")
    spec_before, spec_after = self._HEADING_SPACING_PT[min(level - 1, 5)]
    if space_before is not None: spec_before = space_before
    if space_after is not None: spec_after = space_after
    eff_before = self._collapsed_before(spec_before)
    p.paragraph_format.space_before = Pt(eff_before)
    p.paragraph_format.space_after = Pt(spec_after)
    self._track_block_spacing(spec_after)
    self._current_para = p
    return self

  #------------------------------------------------------------------------------------------ Lists

  def bullet(self, text:str|None=None, level:int=0) -> "DOCX":
    """Bullet list item. Uses built-in `List Bullet` / `List Bullet 2..3` styles.

    For multi-level numbering with custom formats,
    define a numbering style via `doc.doc.styles` directly - this method covers the 99% case.
    """
    self._flush_para()
    style_name = "List Bullet" if level == 0 else f"List Bullet {level + 1}"
    p = self._new_para(text or "", style=self._known_style(style_name, "List Bullet"))
    self._tighten_list_spacing(p)
    # Items pack tight (0/0); the last one needs air before the next block.
    # Which item is last isn't known yet, so every item parks the bonus -
    # only a following non-list paragraph consumes it (list items bypass it).
    self._track_block_spacing(0, extra_for_next=4)
    self._current_para = p
    return self

  def ordered(
    self, text:str|None = None, level:int = 0,
    start:int|None = None,
    delim:str = ".",
  ) -> "DOCX":
    """Ordered list item. Uses built-in `List Number` / `List Number 2..3`.

    `start` opens a new list counted from that number;
    items that follow without it continue the same list.
    The numbering of the built-in styles is shared,
    so every ordered list in a document runs as one sequence
    until one of them asks for its own instance.
    `delim` follows the number of a list `start` opens: `.` or `)`.
    """
    self._flush_para()
    style_name = self._known_style(
      "List Number" if level == 0 else f"List Number {level + 1}",
      "List Number")
    p = self._new_para(text or "", style=style_name)
    if start is not None:
      self._num_ids[level] = self._fresh_num_id(style_name, start, delim)
    num_id = self._num_ids.get(level)
    if num_id is not None:
      self._set_para_num(p, num_id)
    self._tighten_list_spacing(p)
    self._track_block_spacing(0, extra_for_next=4)
    self._current_para = p
    return self

  def _style_abstract_num(self, style_name:str):
    """`(numbering_element, abstractNumId)` behind a list style, or `None`.

    `None` covers every way the lookup can come up empty:
    a document with no numbering part, a style without `numPr`,
    or a `numId` naming an instance that is not there.
    """
    try:
      numbering = self._doc.part.numbering_part.element
      pPr = self._doc.styles[style_name].element.find(qn("w:pPr"))
      src_id = pPr.find(qn("w:numPr")).find(qn("w:numId")).get(qn("w:val"))
    except Exception:
      return None
    for num in numbering.findall(qn("w:num")):
      if num.get(qn("w:numId")) != src_id:
        continue
      abstract = num.find(qn("w:abstractNumId"))
      if abstract is not None:
        return numbering, abstract.get(qn("w:val"))
    return None

  def _fresh_num_id(self, style_name:str, start:int, delim:str=".") -> int|None:
    """New numbering instance cloned from `style_name`, counting from `start`.

    `w:numId` is what Word counts by, so lists sharing one instance run as a single sequence.
    A new instance pointing at the same abstract definition keeps the look
    and gets its own counter.
    `None` when the document has no numbering to extend,
    which leaves the style's own numbering in place.
    A `delim` other than `.` overrides the level with a copy that prints it.
    """
    found = self._style_abstract_num(style_name)
    if found is None:
      return None
    numbering, abstract_id = found
    taken = [int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))]
    new_id = max(taken, default=0) + 1
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_id))
    abstract = OxmlElement("w:abstractNumId")
    abstract.set(qn("w:val"), abstract_id)
    num.append(abstract)
    # `startOverride` is written even for 1:
    # it states the restart rather than leaning on the abstract definition's own start value.
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start_el = OxmlElement("w:startOverride")
    start_el.set(qn("w:val"), str(max(1, int(start))))
    override.append(start_el)
    if delim != ".":
      override.append(_level_with_delim(numbering, abstract_id, start, delim))
    num.append(override)
    numbering.append(num) # `w:num` entries follow every `w:abstractNum`
    return new_id

  @staticmethod
  def _set_para_num(p, num_id:int):
    """Point one paragraph at a numbering instance."""
    pPr = p._element.get_or_add_pPr()
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
      numPr = OxmlElement("w:numPr")
      _ppr_insert(pPr, numPr)
    num_el = numPr.find(qn("w:numId"))
    if num_el is None:
      num_el = OxmlElement("w:numId")
      numPr.append(num_el)
    num_el.set(qn("w:val"), str(num_id))

  def _known_style(self, wanted:str, fallback:str) -> str:
    """`wanted` when the document defines that style, else `fallback`.

    Word ships `List Bullet 2..3` but not deeper levels, and the built-in set varies by template,
    so the level a caller asks for is checked rather than assumed.
    """
    try:
      self._doc.styles[wanted]
      return wanted
    except KeyError:
      return fallback

  def _tighten_list_spacing(self, p):
    """Override Word's loose default list spacing - keep items packed."""
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.1

  #--------------------------------------------------------------------------------- Block elements

  def code_block(
    self, content:str, language:str|None = None,
    bg_color:tuple|str = (0.96, 0.97, 0.98),
    border_color:tuple|str = (0.82, 0.84, 0.87),
    font_family:str = "Consolas", font_size:float = 9,
  ) -> "DOCX":
    """Insert a fenced code block - monospace font, light grey background, thin border.
    `language` is accepted but not used (no syntax highlight).
    """
    self._flush_para()
    p = self._new_para()
    pf = p.paragraph_format
    pf.space_before = Pt(self._collapsed_before(2))
    pf.space_after = Pt(6)
    pf.line_spacing = 1.2
    _apply_paragraph_shading(p, color_hex(bg_color))
    _set_pbdr(p._element, color_hex(border_color), size_eighths=4)
    self._keep_words(p)
    lines = content.rstrip("\n").split("\n")
    for i, line in enumerate(lines):
      if i > 0:
        p.add_run().add_break(WD_BREAK.LINE)
      run = p.add_run(line)
      run.font.name = font_family
      run.font.size = Pt(font_size)
    self._track_block_spacing(6)
    return self

  def blockquote(
    self, text:str|None = None,
    border_color:tuple|str = (0.82, 0.84, 0.87),
    text_color:tuple|str = (0.40, 0.44, 0.50),
    indent:float = 4,
    space_before:float = 3, space_after:float = 3,
  ) -> "DOCX":
    """Insert a blockquote paragraph - thick left border, muted text, indent.

    For multi-paragraph blockquotes / callouts pass `space_before=0` on middle/last paragraphs
    and `space_after=0` on first/middle ones
    so the left bar reads as a single continuous block, not stacked stripes.

    Returns with `_current_para` set so `text()` can append rich runs.
    """
    self._flush_para()
    p = self._new_para()
    pf = p.paragraph_format
    pf.left_indent = Mm(to_mm(indent, self.unit))
    pf.space_before = Pt(self._collapsed_before(space_before))
    pf.space_after = Pt(space_after)
    _set_pbdr(p._element, color_hex(border_color), sides=("left",), size_eighths=24, space=4)
    self._current_para = p
    self._para_default_color = text_color
    self._track_block_spacing(space_after)
    if text:
      self._add_run(p, RichSegment(text=text))
    return self

  def hr(self, color:tuple|str=(0.82, 0.84, 0.87), size_eighths:int=4) -> "DOCX":
    """Insert a horizontal rule - empty paragraph with bottom border."""
    self._flush_para()
    p = self._new_para()
    pf = p.paragraph_format
    pf.space_before = Pt(self._collapsed_before(4))
    pf.space_after = Pt(4)
    _set_pbdr(p._element, color_hex(color), sides=("bottom",),
      size_eighths=size_eighths)
    self._track_block_spacing(4)
    return self

  #----------------------------------------------------------------------------------------- Tables

  def table(
    self,
    body:list[list[str]],
    header:list[str]|None = None,
    aligns:list[str]|None = None,
    widths:list[float]|None = None, # mm per column
    style:TableStyle|None = None,
    word_style:str|None = None, # built-in Word table style name
  ) -> "DOCX":
    """Add a table.

    Args:
      body: List of rows (each row a list of cell strings).
      header: Optional header row.
      aligns: Per-column alignment - `L`/`R`/`C`/`J`.
      widths: Per-column widths in mm.
        `None` sizes columns by content to fill the content area.
        With `style.fill_content_width` off, Word sizes them itself.
      style: `TableStyle` for borders/shading. Ignored when `word_style` set.
      word_style: Built-in Word style name (e.g. `Light Grid`).
    """
    self._flush_para()
    rows = ([header] if header else []) + (body or [])
    if not rows:
      return self
    ncols = max(len(r) for r in rows)
    tbl = self._new_table(len(rows), ncols)
    if word_style:
      try:
        tbl.style = word_style
      except KeyError:
        warnings.warn(
          f"table style {word_style!r} is not defined in this document; "
          "the table renders unstyled",
          RuntimeWarning, stacklevel=2,
        )
    style = style or TableStyle()
    # Cell size: explicit `style.font_size`,
    # else one ladder step below body (11→10, 14→12) via `smaller_size`.
    # Tables read better slightly smaller.
    if style.font_size is None:
      body_pt = self._style.font_size or Defaults.FONT_SIZE
      eff_size_pt = smaller_size(body_pt)
    else:
      eff_size_pt = style.font_size
    # Column widths: explicit `widths=`, else sized by content to fill the content area.
    # Left to `python-docx`, a narrow table drifts left.
    # Explicit widths and no autofit anchor it to the full width.
    if widths:
      col_w_mm = [to_mm(w, self.unit) for w in widths[:ncols]]
      while len(col_w_mm) < ncols:
        col_w_mm.append(self.content_width / ncols)
    elif style.fill_content_width:
      # cell runs name no font, so they set in the `Normal` face, not the current `font()`
      family = self._doc.styles["Normal"].font.name or Defaults.FONT_FAMILY
      bold_rows = 1 if header and style.header_bold else 0
      col_w_mm = content_widths(rows, self.content_width, style.cell_pad_h,
        family, eff_size_pt, bold_rows)
    else:
      col_w_mm = None
    # Borders first: `_set_table_widths` reorders `tblPr` into canonical schema order
    # and slots these borders correctly (strict renderers need the order).
    apply_table_borders(tbl, style.border_color, style.border_size)
    if col_w_mm:
      tbl.autofit = False
      tbl.allow_autofit = False
      _set_table_widths(tbl, col_w_mm, left_pad_mm=style.cell_pad_h)
    if style.table_align:
      tbl.alignment = {"L": WD_TABLE_ALIGNMENT.LEFT, "C": WD_TABLE_ALIGNMENT.CENTER,
        "R": WD_TABLE_ALIGNMENT.RIGHT}.get(style.table_align, WD_TABLE_ALIGNMENT.LEFT)
    if header:
      head_row = tbl.rows[0]
      for j, txt in enumerate(header):
        cell = head_row.cells[j]
        cell.text = "" # clear placeholder
        p = cell.paragraphs[0]
        run = p.add_run(txt)
        run.font.size = Pt(eff_size_pt)
        if style.header_bold:
          run.bold = True
        if style.header_color is not None:
          run.font.color.rgb = RGBColor(*rgb255(style.header_color))
        apply_cell_shading(cell, style.header_bg)
        set_cell_margins(cell, top=style.cell_pad_top, right=style.cell_pad_h,
          bot=style.cell_pad_bot, left=style.cell_pad_h)
        set_cell_vertical_align(cell, style.vertical_align)
        set_cell_align(cell, aligns[j] if aligns and j < len(aligns) else None)
        self._tighten_cell_paras(cell)
      if style.header_repeat:
        repeat_header_row(head_row)
    # Body rows
    body_start = 1 if header else 0
    for i, row in enumerate(body or []):
      tr = tbl.rows[body_start + i]
      bg = style.row_bg_odd if i % 2 else style.row_bg_even
      for j, txt in enumerate(row):
        if j >= len(tr.cells):
          break
        cell = tr.cells[j]
        cell.text = "" # clear placeholder
        body_run = cell.paragraphs[0].add_run(str(txt))
        body_run.font.size = Pt(eff_size_pt)
        if bg is not None:
          apply_cell_shading(cell, bg)
        set_cell_margins(cell, top=style.cell_pad_top, right=style.cell_pad_h,
          bot=style.cell_pad_bot, left=style.cell_pad_h)
        set_cell_vertical_align(cell, style.vertical_align)
        set_cell_align(cell, aligns[j] if aligns and j < len(aligns) else None)
        self._tighten_cell_paras(cell)
    # Cell spacing stays inside the table, so park a `space_before` bonus for the next paragraph -
    # else it sticks to the table's bottom border.
    self._track_block_spacing(0, extra_for_next=6)
    return self

  def _tighten_cell_paras(self, cell):
    """Strip the `Normal` style's `space_after` from cell paragraphs -
    else every cell carries phantom padding and rows look bottom-heavy.
    """
    for p in cell.paragraphs:
      pf = p.paragraph_format
      pf.space_before = Pt(0)
      pf.space_after = Pt(0)
      pf.line_spacing = 1.15
      self._keep_words(p)

  #----------------------------------------------------------------------------------------- Images

  def image(
    self, path:str, width:float|None = None, height:float|None = None,
    align:str|None = None,
  ) -> "DOCX":
    """Insert an image. Width/height in current unit (mm by default).
    Provide only `width` for proportional scaling.

    SVG sources are rasterized first - `python-docx` reads raster formats only.
    Raises `UnrecognizedImageError` for a format Word cannot embed.
    """
    self._flush_para()
    p = self._new_para()
    if align:
      p.alignment = align_to_docx(align)
    run = p.add_run()
    kwargs = {}
    if width: kwargs["width"] = Mm(to_mm(width, self.unit))
    if height: kwargs["height"] = Mm(to_mm(height, self.unit))
    source = path
    if is_svg(path):
      buf = svg_to_png_buffer(path)
      if buf is not None:
        source = buf
    run.add_picture(source, **kwargs)
    # Treat image paragraph like a regular body block for collapse purposes.
    self._track_block_spacing(0)
    return self

  #------------------------------------------------------------------------------------ Page breaks

  def page_break(self) -> "DOCX":
    """Insert a hard page break."""
    self._flush_para()
    p = self._new_para()
    p.add_run().add_break(WD_BREAK.PAGE)
    return self

  #-------------------------------------------------------------------------------------- Bookmarks

  def bookmark(self, name:str) -> "DOCX":
    """Mark a bookmark anchor at the current point (zero-width range).
    Use `link(..., target=name)` to jump here from another spot.
    """
    if self._current_para is None:
      self.para()
    self._bookmark_id += 1
    bid = str(self._bookmark_id)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), bid)
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bid)
    self._current_para._element.append(start)
    self._current_para._element.append(end)
    return self

  def _add_hyperlink(self, p, text:str, url:str|None=None, target:str|None=None):
    """Append a hyperlink run to paragraph. Either `url` or `target`."""
    hl = OxmlElement("w:hyperlink")
    if url:
      r_id = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
      hl.set(qn("r:id"), r_id)
    elif target:
      hl.set(qn("w:anchor"), target)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    # Hyperlink style: blue + underline
    rstyle = OxmlElement("w:rStyle")
    rstyle.set(qn("w:val"), "Hyperlink")
    rpr.append(rstyle)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    run.append(t)
    hl.append(run)
    p._element.append(hl)

  #-------------------------------------------------------------------------------- Headers/footers

  def header(self, text:str|None=None, align:str|None=None) -> "DOCX":
    """Set the header text of the last section, which the sections `columns()` opens share.
    For richer headers use `doc.doc.sections[0].header` directly.
    """
    sec = self._doc.sections[-1]
    p = sec.header.paragraphs[0]
    p.text = text or ""
    if align: p.alignment = align_to_docx(align)
    return self

  def footer(
    self, text:str|None = None, align:str|None = None,
    page_number:bool = False,
  ) -> "DOCX":
    """Set the footer text of the last section, which the sections `columns()` opens share.

    Args:
      text: Static footer content. Use `{page}` / `{pages}` for field codes.
      align: `L` / `R` / `C` / `J`.
      page_number: Append `{page} / {pages}` as Word fields, worked out on every page.
    """
    sec = self._doc.sections[-1]
    p = sec.footer.paragraphs[0]
    p.text = ""
    if align: p.alignment = align_to_docx(align)
    if text:
      if "{page}" in text or "{pages}" in text:
        page_fields(p, text)
      else:
        p.add_run(text)
    if page_number:
      if text: p.add_run("  ")
      page_fields(p, "{page} / {pages}")
    return self

  #--------------------------------------------------------------------------------------- Metadata

  def metadata(
    self, title:str|None = None, author:str|None = None,
    subject:str|None = None, keywords:str|None = None,
    comments:str|None = None, category:str|None = None,
  ) -> "DOCX":
    """Set core document properties."""
    if title: self._metadata.title = title
    if author: self._metadata.author = author
    if subject: self._metadata.subject = subject
    if keywords: self._metadata.keywords = keywords
    if comments: self._metadata.comments = comments
    if category: self._metadata.category = category
    return self

  #------------------------------------------------------------------------------------------- Save

  def save(self) -> "DOCX":
    """Write the document to disk."""
    self._flush_para()
    # a document ends with a paragraph: Word adds a full-size one after a final table itself
    if _last_block(self._doc).tag == qn("w:tbl"): self._hairline()
    self._metadata.apply(self._doc)
    self._doc.save(self.path)
    return self

#------------------------------------------------------------------------- Paragraph shading/border

# CT_PPrBase child sequence (ISO/IEC 29500-1).
# Word tolerates any order, strict renderers may drop what arrives out of sequence,
# so elements built by hand are placed by this table.
_PPR_ORDER = (
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
  "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
  "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
  "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
  "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
  "suppressOverlap", "jc", "textDirection", "textAlignment",
  "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
  "pPrChange",
)

# CT_Settings child sequence, up to the last element a document from the default template holds.
_SETTINGS_ORDER = (
  "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
  "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
  "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts",
  "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges",
  "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
  "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate",
  "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge",
  "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting",
  "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet",
  "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone",
  "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle",
  "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
  "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
  "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
  "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin",
  "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning",
  "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
  "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml",
  "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml",
  "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags",
  "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr",
  "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
  "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions",
  "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults",
  "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
)

# CT_Compat child sequence, up to `doNotExpandShiftReturn`; the options after it rank last.
_COMPAT_ORDER = (
  "useSingleBorderforContiguousCells", "wpJustification", "noTabHangInd", "noLeading",
  "spaceForUL", "noColumnBalance", "balanceSingleByteDoubleByteWidth", "noExtraLineSpacing",
  "doNotLeaveBackslashAlone", "ulTrailSpace", "doNotExpandShiftReturn",
)

# Children that follow `w:cols` in a `w:sectPr`, and `w:lang` in a `w:rPr`.
_SECT_AFTER_COLS = ("w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
  "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange")
_RPR_AFTER_LANG = ("w:eastAsianLayout", "w:specVanish", "w:oMath")

def _insert_ordered(parent, el, order:tuple):
  """
  Insert `el` into `parent` at its place in the schema sequence `order` of local names.

  Elements outside `order` go last -
  the tail of the sequence is the only place a name of unknown rank cannot break a known one.
  """
  def rank(tag:str) -> int:
    name = tag.split("}")[-1]
    return order.index(name) if name in order else len(order)
  target = rank(el.tag)
  for child in parent:
    if rank(child.tag) > target:
      child.addprevious(el)
      return
  parent.append(el)

def _ppr_insert(pPr, el):
  """Insert `el` into `pPr` at its schema position."""
  _insert_ordered(pPr, el, _PPR_ORDER)

def _compat_option(settings, name:str):
  """Turn Word's compatibility option `name` on, in a `w:compat` made when missing."""
  compat = settings.find(qn("w:compat"))
  if compat is None:
    compat = OxmlElement("w:compat")
    _insert_ordered(settings, compat, _SETTINGS_ORDER)
  if compat.find(qn(f"w:{name}")) is None:
    _insert_ordered(compat, OxmlElement(f"w:{name}"), _COMPAT_ORDER)

def _child(parent, tag:str, first:bool=False, after:tuple=()):
  """Child `tag` of `parent`, made when missing: first, ahead of the tags `after`, or last."""
  el = parent.find(qn(tag))
  if el is not None: return el
  el = OxmlElement(tag)
  if first: parent.insert(0, el)
  elif after: parent.insert_element_before(el, *after)
  else: parent.append(el)
  return el

def _lang_tag(lang:str) -> str:
  """Language tag Word names `lang` by: `pl` → `pl-PL`; a tag with a region passes through."""
  return _LANG_TAGS.get(lang.lower(), lang.replace("_", "-"))

def _mark_language(element, lang:str):
  """Set every run under `element` in `lang`, which Word spell-checks and hyphenates it by."""
  tag = _lang_tag(lang)
  for r in element.iter(qn("w:r")):
    _child(r.get_or_add_rPr(), "w:lang", after=_RPR_AFTER_LANG).set(qn("w:val"), tag)

def _no_hyphens(element):
  """Keep the words of a paragraph or paragraph style whole under automatic hyphenation."""
  pPr = element.get_or_add_pPr()
  if pPr.find(qn("w:suppressAutoHyphens")) is None:
    _ppr_insert(pPr, OxmlElement("w:suppressAutoHyphens"))

def _hairline_format(p):
  """Paragraph `p` set 1pt high with no space around it."""
  pf = p.paragraph_format
  pf.space_before = pf.space_after = Pt(0)
  pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
  pf.line_spacing = Pt(1)

def _last_block(document):
  """Last paragraph or table of the body, the `w:sectPr` after it aside, or the body when empty."""
  body = document.element.body
  last = body.get_or_add_sectPr().getprevious()
  return body if last is None else last

def _is_empty(p) -> bool:
  """True when paragraph `p` holds nothing but its properties."""
  return all(child.tag == qn("w:pPr") for child in p._p)

def _first_block(el) -> bool:
  """True when no paragraph or table comes before `el` in its table cell."""
  prev = el.getprevious()
  return prev is None or prev.tag == qn("w:tcPr")

def _apply_paragraph_shading(p, hex_color:str):
  """Set background fill on a paragraph via raw `<w:shd>` in `pPr`."""
  pPr = p._element.get_or_add_pPr()
  shd = pPr.find(qn("w:shd"))
  if shd is None:
    shd = OxmlElement("w:shd")
    _ppr_insert(pPr, shd)
  shd.set(qn("w:val"), "clear")
  shd.set(qn("w:color"), "auto")
  shd.set(qn("w:fill"), hex_color)

_BORDER_SIDES_ALL = ("top", "left", "bottom", "right")
# `w:pBdr` children in schema order, whatever order a caller names the sides in
_PBDR_ORDER = ("top", "left", "bottom", "right", "between", "bar")

def _set_pbdr(
  element, hex_color:str, sides:tuple = _BORDER_SIDES_ALL,
  size_eighths:int = 4, space:int = 1, val:str = "single",
):
  """Set paragraph borders on a paragraph or style element via `<w:pBdr>`.
  Works on both `<w:p>` and `<w:style>` since both expose `get_or_add_pPr()`.

  Args:
    element: `paragraph._element` or `style.element`.
    hex_color: Border color as `RRGGBB` hex string.
    sides: Subset of `("top", "left", "bottom", "right")`. Default: all four.
    size_eighths: Width in eighths-of-a-point (OOXML native unit).
    space: Gap from text in pt.
    val: Word border style, as `single` or `dotted`.
  """
  pPr = element.get_or_add_pPr()
  pBdr = pPr.find(qn("w:pBdr"))
  if pBdr is None:
    pBdr = OxmlElement("w:pBdr")
    _ppr_insert(pPr, pBdr)
  for side in sides:
    el = pBdr.find(qn(f"w:{side}"))
    if el is None:
      el = OxmlElement(f"w:{side}")
      _insert_ordered(pBdr, el, _PBDR_ORDER)
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), str(size_eighths))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), hex_color)

def _clear_border(element, side:str):
  """Remove the `side` border of a paragraph or style element, and an emptied `w:pBdr` with it."""
  pPr = element.find(qn("w:pPr"))
  pBdr = None if pPr is None else pPr.find(qn("w:pBdr"))
  if pBdr is None: return
  el = pBdr.find(qn(f"w:{side}"))
  if el is not None: pBdr.remove(el)
  if len(pBdr) == 0: pPr.remove(pBdr)

def _force_font(style, font_name:str):
  """Strip theme font references on a style and pin `font_name` for every script range.
  Office's `majorHAnsi` theme font resolves to Calibri Light,
  and some renderers prefer the theme ref over the direct `ascii` font.
  """
  el = style.element
  rpr = el.find(qn("w:rPr"))
  if rpr is None:
    rpr = OxmlElement("w:rPr")
    el.append(rpr)
  rfonts = rpr.find(qn("w:rFonts"))
  if rfonts is None:
    rfonts = OxmlElement("w:rFonts")
    rpr.append(rfonts)
  # Drop the theme refs (the source of the Calibri-vs-Calibri-Light mismatch).
  for theme_attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
    full = qn(f"w:{theme_attr}")
    if full in rfonts.attrib:
      del rfonts.attrib[full]
  # Pin direct fonts for every script range.
  for direct_attr in ("ascii", "hAnsi", "eastAsia", "cs"):
    rfonts.set(qn(f"w:{direct_attr}"), font_name)

#---------------------------------------------------------------------------------------- Numbering

def _level_with_delim(numbering, abstract_id:str, start:int, delim:str):
  """
  Copy of level 0 of abstract numbering `abstract_id`,
  counting from `start` and printing `delim` after the number, for a `w:lvlOverride`:
  Word takes a whole level there, not just its text.
  """
  abstract = next(a for a in numbering.findall(qn("w:abstractNum"))
    if a.get(qn("w:abstractNumId")) == abstract_id)
  levels = abstract.findall(qn("w:lvl"))
  lvl = copy.deepcopy(next(el for el in levels if el.get(qn("w:ilvl")) == "0"))
  _child(lvl, "w:start", first=True).set(qn("w:val"), str(max(1, int(start))))
  lvl.find(qn("w:lvlText")).set(qn("w:val"), f"%1{delim}")
  return lvl

#-------------------------------------------------------------------------------------- Table width

# `dxa` (twentieths of a point) is the OOXML native length unit for tables.
_DXA_PER_MM = 1440 / 25.4

def _set_table_widths(tbl, col_widths_mm:list[float], left_pad_mm:float=0):
  """Anchor a table's total/per-column widths and left edge via raw XML.

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
  total_dxa = int(round(sum(col_widths_mm) * _DXA_PER_MM))
  left_dxa = int(round(left_pad_mm * _DXA_PER_MM))
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
  tbl_cell_mar = OxmlElement("w:tblCellMar") # 0 all sides; per-cell `tcMar` owns padding
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
    gc.set(qn("w:w"), str(int(round(w_mm * _DXA_PER_MM))))
    new_grid.append(gc)
  # Insert `tblGrid` right after `tblPr` (OOXML schema requires this order).
  tbl_el.insert(list(tbl_el).index(tbl_pr) + 1, new_grid)
  # Also set each cell width so cell-level overrides are consistent.
  for row in tbl.rows:
    for i, w_mm in enumerate(col_widths_mm):
      if i < len(row.cells):
        row.cells[i].width = Mm(w_mm)
