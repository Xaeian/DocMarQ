# docmarq/core.py

"""Core DOCX class - main facade for document generation."""
import copy, warnings
from contextlib import contextmanager
from dataclasses import dataclass
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.run import Run
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .constants import Defaults, DXA_PER_MM
from .utils import (to_mm, parse_margin, align_to_docx, color_hex,
  rgb255, smaller_size, tight_line_height, eighths)
from .styles import Style, TableStyle
from .layout import PageGeometry
from .structure import Metadata, new_properties
from .inline import RichSegment, _apply_run_format, set_char_space
from .svg import svg_to_png_buffer, is_svg
from .fields import complex_field, mark_field, page_fields
from .images import add_picture, pin_offsets
from .emoji import font_emoji, EMOJI_FONT
from .fonts import embed_fonts
from .metrics import text_width_mm
from . import dropcap, notes
from .tables import (
  apply_table_rules, check_table_rules, rule_cell, apply_cell_shading, set_cell_align,
  repeat_header_row, apply_table_borders, remove_table_borders, keep_rows,
  set_cell_margins, set_cell_vertical_align, content_widths, set_table_widths,
  TablePicture, CELL_LINE_SPACING, words_width,
)

# Language tags Word names a language by, for the codes of the presets.
_LANG_TAGS = {"en": "en-US", "pl": "pl-PL", "de": "de-DE", "fr": "fr-FR", "es": "es-ES",
  "it": "it-IT", "cs": "cs-CZ", "sk": "sk-SK", "pt": "pt-PT"}
_TABS = {"L": WD_TAB_ALIGNMENT.LEFT, "C": WD_TAB_ALIGNMENT.CENTER, "R": WD_TAB_ALIGNMENT.RIGHT}

#------------------------------------------------------------------------------------------- Target

@dataclass
class _Target:
  """Where blocks go: the document body, a table cell or a footnote."""
  box: object # `Document`, `_Cell` or `notes.Note`, each with `add_paragraph` and `add_table`
  width: float|None = None # mm, a cell's, or `None` to follow the page and its columns
  spare: object = None # empty paragraph the next block fills, as a fresh cell's first

#--------------------------------------------------------------------------------------------- DOCX

class DOCX:
  """
  Main DOCX generator with fluent API.

  A block of one's own, as the markdown engine and the document packages write theirs,
  starts with `enter()` and takes `new_para()` or `new_table()` at the end of the write target.
  It reports its gap below to `track_block_spacing()`, so the next block's gap collapses into it.

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
    two_sided:bool = False,
    font_dir:str|None = None,
  ):
    """
    `gutter` is Word's binding margin, added to the left one.
    `two_sided` prints on both sides of the sheet: Word mirrors the margins of even pages,
    so the left one and the gutter stay at the bound edge, as in pdfmarq.
    `font_dir` holds faces as pdfmarq takes them, `Family/Family-Mode.ttf`: those the document
    names go into its file as it is saved, so Word shows them where they are not installed.
    """
    self.path = path
    self._font_dir = font_dir
    self.unit = unit
    top, right, bot, left = parse_margin(margin)
    self._page = PageGeometry(
      width=to_mm(width, unit), height=to_mm(height, unit),
      margin_top=to_mm(top, unit), margin_right=to_mm(right, unit),
      margin_bot=to_mm(bot, unit), margin_left=to_mm(left, unit),
      gutter=to_mm(gutter, unit),
    )
    self._doc = Document(template) if template else Document()
    new_properties(self._doc, blank=not template)
    settings = self._doc.settings.element
    # a justified line a manual break ends stays flush left, as in pdfmarq
    _compat_option(settings, "doNotExpandShiftReturn")
    if two_sided and settings.find(qn("w:mirrorMargins")) is None:
      _insert_ordered(settings, OxmlElement("w:mirrorMargins"), _SETTINGS_ORDER)
    self._apply_section_geometry()
    self._style = Style().with_defaults()
    self._style.font_family = font_body
    self._style.font_size = body_size
    self._metadata = Metadata()
    self._current_para = None # open paragraph, where runs append
    self._num_ids: dict[int, int|None] = {} # list level → numbering instance
    self._para_default_color = None # per-paragraph color override, reset on flush
    self._last_space_after = 0 # pt - previous block's `space_after` (margin collapse)
    self._extra_block_after = 0 # pt - one-shot `space_before` bonus (see `track_block_spacing`)
    self._bookmark_id = 0
    self._table = None # the table written last, which `keep_rows()` holds
    self._at = _Target(self._doc)
    self._cols = (1, 0.0) # columns of the current section: count, gap (mm)
    self._hyphenate = False # set by `hyphenation()`: headings, cells and code keep words whole
    self._emoji_font = None # set by `emoji()`: the face emoji take when the document is saved
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
  def in_columns(self) -> bool:
    """Whether the current section sets its text in more than one column."""
    return self._cols[0] > 1

  @property
  def content_width(self) -> float:
    """Width in mm blocks are laid out in: a column of the text area, or the cell `into` fills."""
    if self._at.width is not None: return self._at.width
    count, gap = self._cols
    return (self._page.content_width - gap * (count - 1)) / count

  def text_width(self, text:str, family:str, size:float, bold:bool=False) -> float:
    """mm `text` takes set in `family` at `size` pt, from the faces this file carries first."""
    return text_width_mm(text, family, size, bold, self._font_dir)

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

  @property
  def paragraph(self):
    """The open paragraph `text()` writes to, as python-docx has it, or `None`."""
    return self._current_para

  @property
  def hyphenates(self) -> bool:
    """Whether `hyphenation()` is on."""
    return self._hyphenate

  @property
  def text_style(self) -> Style:
    """Style of what is written next: font, colour, leading and gaps, as `style()` sets them."""
    return self._style

  @text_style.setter
  def text_style(self, style:Style):
    self._style = style

  #--------------------------------------------------------------------------- Section / page setup

  def _override_heading_styles(self, font_body:str, font_head:str):
    """
    Recolor Word's built-in `Heading 1..6` to the GitHub-light palette:
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
      except KeyError: continue
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
    except KeyError: pass

  def heading_style(
    self, level:int,
    size:float|None = None,
    bold:bool|None = None,
    italic:bool|None = None,
    color:tuple|str|None = None,
    rule:bool|None = None,
    family:str|None = None,
    caps:bool|None = None,
    upper:bool|None = None,
    align:str|None = None,
    indent:float|None = None,
    page:bool|None = None,
    lines:float|None = None,
    at_least:float|None = None,
    char_space:float|None = None,
  ) -> "DOCX":
    """
    Restyle Word's `Heading {level}`, so every heading of that level follows.

    `size` in pt, `bold` and `italic` the face, `color` its text, `rule` the thin line under it,
    `family` its font, `caps` small capitals, `upper` capitals alone,
    `align` and `indent` its place, `page` a new page before it,
    its lines `lines` apart as a multiple or `at_least` pt, its letters `char_space` pt apart.
    `None` leaves that property as the document has it.
    """
    try:
      st = self._doc.styles[f"Heading {level}"]
    except KeyError: return self
    pf = st.paragraph_format
    if size is not None:
      st.font.size = Pt(size)
      # Tighter leading for display sizes: 1.15 at body, → 1.0 at 24pt+.
      pf.line_spacing = tight_line_height(size)
    if bold is not None: st.font.bold = bold
    if italic is not None: st.font.italic = italic
    if color is not None: st.font.color.rgb = RGBColor(*rgb255(color))
    if rule: _set_pbdr(st.element, color_hex(Defaults.RULE_COLOR), sides=("bottom",))
    elif rule is False: _clear_border(st.element, "bottom")
    if family is not None: _force_font(st, family)
    if caps is not None: st.font.small_caps = caps
    if upper is not None: st.font.all_caps = upper
    if align is not None: pf.alignment = align_to_docx(align)
    if indent is not None: pf.left_indent = Mm(to_mm(indent, self.unit))
    if page is not None: pf.page_break_before = page
    if char_space is not None: set_char_space(st.element.get_or_add_rPr(), char_space)
    _leading(pf, lines, at_least)
    return self

  def paragraph_style(
    self, name:str,
    size:float|None = None,
    bold:bool|None = None,
    italic:bool|None = None,
    color:tuple|str|None = None,
    family:str|None = None,
    upper:bool|None = None,
    char_space:float|None = None,
    before:float|None = None,
    after:float|None = None,
    at_least:float|None = None,
  ) -> "DOCX":
    """
    Restyle Word's paragraph style `name`, as `toc 1` the contents fill their chapters with:
    `size` in pt, `bold` and `italic` the face, `color` its text, `family` its font,
    `upper` capitals alone, its letters `char_space` pt apart, `before` and `after` pt around
    its paragraphs, their lines `at_least` pt apart; `None` leaves a property as it is.
    A style the document lacks is left alone, as `heading_style` leaves one.
    """
    if name not in self._doc.styles: return self
    st = self._doc.styles[name]
    pf = st.paragraph_format
    if before is not None: pf.space_before = Pt(before)
    if after is not None: pf.space_after = Pt(after)
    if at_least is not None:
      pf.line_spacing, pf.line_spacing_rule = Pt(at_least), WD_LINE_SPACING.AT_LEAST
    if size is not None: st.font.size = Pt(size)
    if bold is not None: st.font.bold = bold
    if italic is not None: st.font.italic = italic
    if color is not None: st.font.color.rgb = RGBColor(*rgb255(color))
    if family is not None: _force_font(st, family)
    if upper is not None: st.font.all_caps = upper
    if char_space is not None: set_char_space(st.element.get_or_add_rPr(), char_space)
    return self

  def body_style(self, lines:float|None=None, at_least:float|None=None) -> "DOCX":
    """
    Lines of Word's `Normal` style, which every paragraph of no leading of its own takes:
    `lines` apart as a multiple, or `at_least` pt.
    """
    _leading(self._doc.styles["Normal"].paragraph_format, lines, at_least)
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
    Set the blocks that follow in `count` columns `gap` apart.
    `columns(1)` returns to one.

    Each change starts a continuous section, so the text runs on down the same page,
    unless nothing is set in the section yet: that one takes the columns itself.
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
    if count > 1: cols.set(qn("w:space"), str(round(gap_mm * DXA_PER_MM)))
    self._cols = (count, gap_mm)
    return self

  def _new_section(self):
    """
    End the section so far at the last paragraph and open a continuous one after it.
    Returns the `w:sectPr` of the new section, the last in the body.
    A section nothing is set in yet is returned as it is, so no empty paragraph ends it.

    The new section takes its headers and footers from the one before,
    and never a first-page variant: it does not start a page.
    """
    last = self._doc.element.body.get_or_add_sectPr()
    p = last.getprevious()
    if p is None or p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None: return last
    if p.tag != qn("w:p"): p = self.hairline()._p
    p.set_sectPr(last.clone())
    for ref in last.xpath("w:headerReference|w:footerReference"):
      last.remove(ref)
    last.titlePg_val = False
    last.start_type = WD_SECTION.CONTINUOUS
    return last

  def hairline(self):
    """
    Empty paragraph 1pt high after the last block:
    it ends a section where the last block is no paragraph, keeps two tables apart,
    or carries a break or a floating table's anchor.
    """
    self._flush_para()
    p = self.new_para()
    _hairline_format(p)
    self.track_block_spacing(0)
    return p

  def chain(self, write):
    """
    Call `write()`, then chain the paragraphs it wrote, those in its tables too,
    with keep-with-next, all but the last, so Word keeps them on one page.
    """
    self._flush_para()
    target = self._at.box
    body = target is self._doc
    if body: box = self._doc.element.body
    elif isinstance(target, notes.Note): box = target.element
    else: box = target._tc
    # blocks go before the body's closing `w:sectPr`, or after the last block of a cell or a note
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
    """
    Set default font for subsequent runs.

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
        no_hyphens(self._doc.styles[f"Heading {i}"].element)
      except KeyError: continue
    self._hyphenate = True
    return self

  def emoji(self, font:str|None=EMOJI_FONT) -> "DOCX":
    """
    Set emoji in `font` when the document is saved, each stretch of them a run of its own,
    since a text face holds none and Word would print a box.
    `None` leaves them as written.
    """
    self._emoji_font = font
    return self

  def keep_words(self, p):
    """Keep the words of paragraph `p` whole while `hyphenation()` is on."""
    if self._hyphenate: no_hyphens(p._p)

  def drop_cap(
    self,
    size:float,
    line:float,
    lines:int = 2,
    sink:float = 0,
    caps:bool = False,
  ) -> "DOCX":
    """
    Drop the first letter of the paragraph just written `lines` lines, as Word's own drop cap:
    `size` pt tall, sunk `sink` pt, in a frame of lines `line` pt apart.
    `caps` sets the rest of its first word in capitals, as IEEE journals do.
    A paragraph opening with no letter stays as it is.
    """
    if self._current_para is not None:
      dropcap.drop_cap(self._current_para, size, line, lines, sink, caps)
    return self

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

  @contextmanager
  def footnote(self, mark:str|None=None):
    """
    Write the paragraphs of the `with` body as a footnote of the open paragraph,
    then go on writing that paragraph.

    `mark=None` numbers the note as Word does, a string marks it with itself,
    and `""` leaves it unmarked, as a note of the authors' affiliations.
    The note's paragraphs take Word's `footnote text` style and no gaps,
    their runs the face `font()` sets, inside the `with` body as anywhere.
    A note holds paragraphs only, so one inside another refuses, as a table there does.
    A note whose body fails leaves neither the note nor its mark behind.
    """
    if isinstance(self._at.box, notes.Note): raise ValueError("a footnote holds no footnote")
    if self._current_para is None: self.para()
    # the paragraph goes on after the note, in its colour too, as a quote's
    host, color = self._current_para, self._para_default_color
    self._footnote_settings()
    note = notes.add_note(self._doc, host, mark)
    outer = (self._at, self._style, self._last_space_after, self._extra_block_after)
    self._at = _Target(note, self.content_width)
    self._style = self._style.copy(space_before=0, space_after=0)
    self._flush_para()
    self._last_space_after = self._extra_block_after = 0
    try:
      yield self
      self._flush_para()
      notes.mark_note(self._doc, note, mark)
    except Exception:
      notes.drop_note(host, note)
      raise
    finally:
      self._at, self._style, self._last_space_after, self._extra_block_after = outer
      self._current_para, self._para_default_color = host, color

  def footnote_rule(self, show:bool) -> "DOCX":
    """Draw the short rule Word sets above the footnotes, or leave it out."""
    self._footnote_settings()
    notes.rule(self._doc, show)
    return self

  def _footnote_settings(self):
    """Word's footnote settings, naming the separators of the footnotes part, added once."""
    settings = self._doc.settings.element
    if settings.find(qn("w:footnotePr")) is None:
      _insert_ordered(settings, notes.settings_entry(), _SETTINGS_ORDER)

  def new_para(self, text:str="", style:str|None=None):
    """
    New paragraph at the end of the write target, in its spare empty one when it has one.
    It is bare: no gaps, and `text()` still writes to the open paragraph.
    """
    at = self._at
    if at.spare is None: return at.box.add_paragraph(text, style=style)
    p, at.spare = at.spare, None
    if p._p.pPr is not None: p._p.remove(p._p.pPr) # a spare formatted as a hairline
    if text: p.add_run(text)
    if style is not None: p.style = style
    return p

  def new_table(self, rows:int, cols:int):
    """
    New table at the end of the write target, a hairline apart from a table just before it.

    In a cell, a table ahead of every other block takes the place of the empty paragraph,
    and the paragraph Word wants after a table becomes the spare.
    """
    at = self._at
    if at.box is self._doc:
      # Word joins two tables with nothing between them into one
      if _last_block(self._doc).tag == qn("w:tbl"): self.hairline()
      self._table = self._doc.add_table(rows=rows, cols=cols)
      return self._table
    if at.spare is not None and _first_block(at.spare._p):
      at.spare._p.getparent().remove(at.spare._p)
    self._table = at.box.add_table(rows, cols)
    at.spare = at.box.paragraphs[-1]
    return self._table

  def grid(
    self, widths:list[float], rows:int = 1,
    rule:tuple|None = None,
    height:float|None = None,
    middle:bool = False,
  ) -> list[list]:
    """
    A table to lay out by: `rows` rows of cells `widths` wide, ruled `(color, pt)` by `rule`
    or bare, each row `height` tall at least, what its cells hold in their `middle` with `middle`.
    Returns its cells by row, which `into()` writes in and `span()` joins.
    """
    table = self.new_table(rows, len(widths))
    if rule: apply_table_borders(table, *rule)
    else: remove_table_borders(table)
    set_table_widths(table, [to_mm(w, self.unit) for w in widths])
    for row in table.rows:
      if height is not None:
        row.height, row.height_rule = Mm(to_mm(height, self.unit)), WD_ROW_HEIGHT_RULE.AT_LEAST
      if middle:
        for cell in row.cells: set_cell_vertical_align(cell, "center")
    return [list(row.cells) for row in table.rows]

  def span(self, first, last):
    """The cells of a row from `first` to `last` of `grid()` joined into one, which it returns."""
    return first.merge(last)

  def keep_rows(self, ends:int=0) -> "DOCX":
    """
    Each row of the table written last on one page, and a break between pages or columns
    kept from its ends: its first `ends + 1` rows go together, and so do its last `ends`.
    """
    if self._table is not None: keep_rows(self._table, ends)
    return self

  #------------------------------------------------------------------------------------- Paragraphs

  def para(
    self, text:str|None = None, style:str|None = None,
    align:str|None = None,
  ) -> "DOCX":
    """
    Start a new paragraph. Closes any previous active paragraph.

    Args:
      text: Optional initial text run.
      style: Built-in Word style name (e.g. `Quote`, `Intense Quote`).
      align: `L` / `R` / `C` / `J`.
    """
    self._flush_para()
    p = self.new_para(style=style)
    if align:
      p.alignment = align_to_docx(align)
    self.apply_para_spacing(p)
    self._current_para = p
    if text:
      self._add_run(p, RichSegment(text=text))
    return self

  def text(self, content:str, **seg_kwargs) -> "DOCX":
    """
    Add a styled run to the current paragraph.
    Auto-creates a paragraph if none active.
    Kwargs match `RichSegment` fields (`bold`, `italic`, `code`, `color`, `link_url`, etc.).
    """
    if self._current_para is None:
      self.para()
    seg = RichSegment(text=content, **seg_kwargs)
    self._add_run(self._current_para, seg)
    return self

  def link(self, content:str, url:str|None=None, target:str|None=None, **seg_kwargs) -> "DOCX":
    """
    Add a hyperlink run in the look of the text around it, `seg_kwargs` as `text()` takes them:
    in `color` where given, as a PDF draws a link, else in Word's underlined blue.
    Either `url` for external link or `target` for internal bookmark name.
    """
    return self.text(content, link_url=url, link_target=target, **seg_kwargs)

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
    dotted:bool = False,
  ) -> "DOCX":
    """
    Rule the current paragraph on `sides`, `size` pt thick, `space` pt clear of its text,
    in dots by `dotted`, as a line to sign on.
    Word keeps the width in eighths of a point and the space in whole points.
    """
    if self._current_para is None:
      self.para()
    _set_pbdr(self._current_para._p, color_hex(color), sides=sides,
      size_eighths=eighths(size), space=round(space), val="dotted" if dotted else "single")
    return self

  def cell_rule(
    self, cell, color:tuple|str,
    side:str = "bottom",
    size:float = 0.5,
    dotted:bool = False,
  ) -> "DOCX":
    """
    Rule one `side` of a `cell` of `grid()`, `size` pt thick, in dots by `dotted`,
    as the light lines between the rows of a form.
    """
    rule_cell(cell, color, size, side, "dotted" if dotted else "single")
    return self

  def shape(
    self,
    indent:float|None = None,
    first:float|None = None,
    before:float|None = None,
    after:float|None = None,
    lines:float|None = None,
    at_least:float|None = None,
    keep:bool|None = None,
    tabs:tuple = (),
    align:str|None = None,
    style:str|None = None,
    page:bool|None = None,
    right:float|None = None,
    exact:float|None = None,
  ) -> "DOCX":
    """
    Lay out the open paragraph: `indent` in from the left and `right` from the right,
    its first line `first` further, a negative one hanging, `before` and `after` pt around it,
    its lines `lines` apart as a multiple, `at_least` pt or `exact` pt, kept with the next one
    by `keep`,
    a tab stop at each `(position, "L"|"C"|"R")` of `tabs`, `(position, "R", ".")` dots led to it,
    in Word's `style` by name,
    on a new page by `page`, with no empty paragraph before it as `page_break()` leaves.
    Its gaps are its own: the next paragraph's gap does not collapse into them.
    `None` leaves a property as it is.
    """
    if self._current_para is None: self.para()
    p = self._current_para
    pf = p.paragraph_format
    if style is not None: p.style = self._doc.styles[style]
    if indent is not None: pf.left_indent = Mm(to_mm(indent, self.unit))
    if right is not None: pf.right_indent = Mm(to_mm(right, self.unit))
    if first is not None: pf.first_line_indent = Mm(to_mm(first, self.unit))
    if before is not None: pf.space_before = Pt(before)
    if after is not None: pf.space_after = Pt(after)
    _leading(pf, lines, at_least)
    if exact is not None:
      pf.line_spacing, pf.line_spacing_rule = Pt(exact), WD_LINE_SPACING.EXACTLY
    if keep is not None: pf.keep_with_next = keep
    if page is not None: pf.page_break_before = page
    for at, kind, *lead in tabs:
      leader = WD_TAB_LEADER.DOTS if lead == ["."] else WD_TAB_LEADER.SPACES
      pf.tab_stops.add_tab_stop(Mm(to_mm(at, self.unit)), _TABS[kind], leader)
    if align: p.alignment = align_to_docx(align)
    return self

  def restyle(
    self,
    bold:bool|None = None,
    italic:bool|None = None,
    caps:bool|None = None,
    size:float|None = None,
    family:str|None = None,
    color:tuple|str|None = None,
    lang:str|None = None,
  ) -> "DOCX":
    """
    Set the runs of the open paragraph written so far in this look, over the style they take:
    `False` turns off a face their style turns on, `caps` sets small capitals,
    `lang` the language Word checks and hyphenates them by.
    `None` leaves a property as it is.
    """
    p = self._current_para
    if p is None: return self
    for r in p._p.iter(qn("w:r")): # a link's words too, which `p.runs` leaves out
      font = Run(r, p).font
      if bold is not None: font.bold = bold
      if italic is not None: font.italic = italic
      if caps is not None: font.small_caps = caps
      if size is not None: font.size = Pt(size)
      if family is not None: font.name = family
      if color is not None: font.color.rgb = RGBColor(*rgb255(color))
    if lang is not None: mark_language(p._p, lang)
    return self

  def _flush_para(self):
    """Close the active paragraph - block elements must call this first."""
    self._current_para = None
    self._para_default_color = None

  def apply_para_spacing(self, p):
    """
    Apply line-height and before/after spacing with CSS-style margin collapse.
    Word SUMS adjacent paragraph margins (CSS takes the max),
    so `space_before` is shrunk by the previous block's `space_after`
    to land a gap of `max(prev_after, declared_before)`.
    See `collapsed_before`.
    """
    pf = p.paragraph_format
    if self._style.line_height is not None:
      pf.line_spacing = self._style.line_height
    declared_before = self._style.space_before or 0
    declared_after = self._style.space_after or 0
    pf.space_before = Pt(self.collapsed_before(declared_before))
    pf.space_after = Pt(declared_after)
    self._last_space_after = declared_after

  def collapsed_before(self, declared_before:float) -> float:
    """
    Effective `space_before` of a block declaring `declared_before` pt:
    collapsed against the previous block's `space_after`,
    plus the one-shot bonus `track_block_spacing()` left, which this uses up.
    """
    val = max(0, declared_before - self._last_space_after) + self._extra_block_after
    self._extra_block_after = 0
    return val

  def track_block_spacing(self, after_pt:float, extra_for_next:float=0):
    """
    Record this block's `space_after` (for collapse against the next paragraph)
    plus an optional one-shot `space_before` bonus for it.

    Tables and lists set `extra_for_next`: their inner paragraphs use `space_after=0`,
    so nothing else would give the next block breathing room.
    """
    self._last_space_after = after_pt
    if extra_for_next:
      self._extra_block_after = extra_for_next

  @staticmethod
  def _style_owns_font(p) -> bool:
    """
    True when the paragraph style carries its own font.

    Runs appended to such a paragraph inherit family, size and color from the style;
    writing body values onto them would show a heading in body type.
    """
    try:
      name = p.style.name or ""
    except Exception: return False
    return name.startswith("Heading") or name in ("Title", "Subtitle")

  def _add_run(self, p, seg:RichSegment):
    """Add a run with `seg` styling, falling back to current `_style`."""
    if seg.break_line:
      p.add_run().add_break(WD_BREAK.LINE)
    linked = bool(seg.link_url or seg.link_target)
    run = p.add_run(seg.text)
    # Effective family/size: segment wins, else current style.
    # A styled paragraph (heading) contributes neither, so its own font survives.
    inherited = self._style_owns_font(p)
    family = seg.family or (None if inherited else self._style.font_family)
    size = seg.size or (None if inherited else self._style.font_size)
    # Color: the segment's own, else the paragraph's default set by `blockquote()` and the like,
    # else the current `_style.color`; a link naming none takes Word's blue below.
    if seg.color is not None:
      eff_color = seg.color
    elif inherited or linked:
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
      small_caps=seg.small_caps or bool(self._style.small_caps),
      char_space=seg.char_space,
    )
    _apply_run_format(run, merged, family, size,
      code_family=self._style.code_family, code_size=self._style.code_size,
      code_color=self._style.code_color, code_bg=self._style.code_bg)
    if linked: _link_run(p, run, seg)

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
    """
    Add a heading using Word's built-in `Heading 1..9` styles (TOC-aware).
    Leaves `_current_para` set so `text()` can append rich inline runs.
    `space_before` and `space_after` in pt replace the gaps of the level.
    `space_before` is recomputed for margin collapse against the previous block.
    """
    self._flush_para()
    level = max(1, min(9, level))
    p = self.new_para(text or "", style=f"Heading {level}")
    spec_before, spec_after = self._HEADING_SPACING_PT[min(level - 1, 5)]
    if space_before is not None: spec_before = space_before
    if space_after is not None: spec_after = space_after
    eff_before = self.collapsed_before(spec_before)
    p.paragraph_format.space_before = Pt(eff_before)
    p.paragraph_format.space_after = Pt(spec_after)
    self.track_block_spacing(spec_after)
    self._current_para = p
    return self

  #------------------------------------------------------------------------------------------ Lists

  def bullet(self, text:str|None=None, level:int=0) -> "DOCX":
    """
    Bullet list item. Uses built-in `List Bullet` / `List Bullet 2..3` styles.

    For multi-level numbering with custom formats,
    define a numbering style via `doc.doc.styles` directly - this method covers the 99% case.
    """
    self._flush_para()
    style_name = "List Bullet" if level == 0 else f"List Bullet {level + 1}"
    p = self.new_para(text or "", style=self._known_style(style_name, "List Bullet"))
    self._tighten_list_spacing(p)
    # Items pack tight (0/0).
    # The last one needs air before the next block.
    # Which item is last isn't known yet, so every item parks the bonus -
    # only a following non-list paragraph consumes it (list items bypass it).
    self.track_block_spacing(0, extra_for_next=4)
    self._current_para = p
    return self

  def ordered(
    self, text:str|None = None, level:int = 0,
    start:int|None = None,
    delim:str = ".",
  ) -> "DOCX":
    """
    Ordered list item. Uses built-in `List Number` / `List Number 2..3`.

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
    p = self.new_para(text or "", style=style_name)
    if start is not None:
      self._num_ids[level] = self._fresh_num_id(style_name, start, delim)
    num_id = self._num_ids.get(level)
    if num_id is not None:
      self._set_para_num(p, num_id)
    self._tighten_list_spacing(p)
    self.track_block_spacing(0, extra_for_next=4)
    self._current_para = p
    return self

  def _style_abstract_num(self, style_name:str):
    """
    `(numbering_element, abstractNumId)` behind a list style, or `None`.

    `None` covers every way the lookup can come up empty:
    a document with no numbering part, a style without `numPr`,
    or a `numId` naming an instance that is not there.
    """
    try:
      numbering = self._doc.part.numbering_part.element
      pPr = self._doc.styles[style_name].element.find(qn("w:pPr"))
      src_id = pPr.find(qn("w:numPr")).find(qn("w:numId")).get(qn("w:val"))
    except Exception: return None
    for num in numbering.findall(qn("w:num")):
      if num.get(qn("w:numId")) != src_id: continue
      abstract = num.find(qn("w:abstractNumId"))
      if abstract is not None: return numbering, abstract.get(qn("w:val"))
    return None

  def _fresh_num_id(self, style_name:str, start:int, delim:str=".") -> int|None:
    """
    New numbering instance cloned from `style_name`, counting from `start`.

    `w:numId` is what Word counts by, so lists sharing one instance run as a single sequence.
    A new instance pointing at the same abstract definition keeps the look
    and gets its own counter.
    `None` when the document has no numbering to extend,
    which leaves the style's own numbering in place.
    A `delim` other than `.` overrides the level with a copy that prints it.
    """
    found = self._style_abstract_num(style_name)
    if found is None: return None
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
    """
    `wanted` when the document defines that style, else `fallback`.

    Word ships `List Bullet 2..3` but not deeper levels, and the built-in set varies by template,
    so the level a caller asks for is checked rather than assumed.
    """
    try:
      self._doc.styles[wanted]
      return wanted
    except KeyError: return fallback

  def _tighten_list_spacing(self, p):
    """
    Override Word's loose default list spacing - keep items packed.
    With `line_height=None` the leading stays Word's paragraph style's, as for the text.
    """
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if self._style.line_height is not None: pf.line_spacing = 1.1

  #--------------------------------------------------------------------------------- Block elements

  def code_block(
    self, content:str, language:str|None = None,
    bg_color:tuple|str = (0.96, 0.97, 0.98),
    border_color:tuple|str = (0.82, 0.84, 0.87),
    font_family:str = "Consolas", font_size:float = 9,
  ) -> "DOCX":
    """
    Insert a fenced code block - monospace font, light grey background, thin border.
    `language` is accepted but not used (no syntax highlight).
    """
    self._flush_para()
    p = self.new_para()
    pf = p.paragraph_format
    pf.space_before = Pt(self.collapsed_before(2))
    pf.space_after = Pt(6)
    pf.line_spacing = 1.2
    _apply_paragraph_shading(p, color_hex(bg_color))
    _set_pbdr(p._element, color_hex(border_color), size_eighths=4)
    self.keep_words(p)
    lines = content.rstrip("\n").split("\n")
    for i, line in enumerate(lines):
      if i > 0:
        p.add_run().add_break(WD_BREAK.LINE)
      run = p.add_run(line)
      run.font.name = font_family
      run.font.size = Pt(font_size)
    self.track_block_spacing(6)
    return self

  def blockquote(
    self, text:str|None = None,
    border_color:tuple|str = (0.82, 0.84, 0.87),
    text_color:tuple|str = (0.40, 0.44, 0.50),
    indent:float = 4,
    space_before:float = 3, space_after:float = 3,
  ) -> "DOCX":
    """
    Insert a blockquote paragraph - thick left border, muted text, indent.

    For multi-paragraph blockquotes / callouts pass `space_before=0` on middle/last paragraphs
    and `space_after=0` on first/middle ones
    so the left bar reads as a single continuous block, not stacked stripes.

    Returns with `_current_para` set so `text()` can append rich runs.
    """
    self._flush_para()
    p = self.new_para()
    pf = p.paragraph_format
    pf.left_indent = Mm(to_mm(indent, self.unit))
    pf.space_before = Pt(self.collapsed_before(space_before))
    pf.space_after = Pt(space_after)
    _set_pbdr(p._element, color_hex(border_color), sides=("left",), size_eighths=24, space=4)
    self._current_para = p
    self._para_default_color = text_color
    self.track_block_spacing(space_after)
    if text:
      self._add_run(p, RichSegment(text=text))
    return self

  def hr(
    self, color:tuple|str = (0.82, 0.84, 0.87), size_eighths:int = 4,
    length:float|None = None,
    align:str = "L",
    before:float = 4,
    after:float = 4,
    dotted:bool = False,
  ) -> "DOCX":
    """
    A horizontal rule, an empty paragraph with a bottom border, `before` and `after` pt around it,
    across the text, or `length` long at `align` as an accent under a title: a short one stands
    a point tall, so it keeps close to the line above. `dotted` sets it in dots, a line to sign on.
    """
    self._flush_para()
    p = self.new_para()
    pf = p.paragraph_format
    pf.space_before = Pt(self.collapsed_before(before))
    pf.space_after = Pt(after)
    if length is not None:
      rest = max(0.0, self.content_width - to_mm(length, self.unit)) # mm beside the rule
      left = {"L": 0.0, "C": rest / 2, "R": rest}[align]
      pf.left_indent, pf.right_indent = Mm(left), Mm(rest - left)
      pf.line_spacing, pf.line_spacing_rule = Pt(1), WD_LINE_SPACING.EXACTLY
      mark = OxmlElement("w:rPr")
      mark.append(OxmlElement("w:sz", {qn("w:val"): "2"}))
      _ppr_insert(p._p.get_or_add_pPr(), mark)
    _set_pbdr(p._element, color_hex(color), sides=("bottom",),
      size_eighths=size_eighths, val="dotted" if dotted else "single")
    self.track_block_spacing(after)
    return self

  #----------------------------------------------------------------------------------------- Tables

  def table_words(self, header:list|None, body:list[list], style:TableStyle|None=None) -> float:
    """
    mm `table()` needs for these rows to hold their longest words,
    each column's widest word with its padding: a narrower table breaks words inside cells.
    """
    style = style or TableStyle()
    rows = ([header] if header else []) + (body or [])
    if not rows: return 0.0
    bold_rows = 1 if header and style.header_bold else 0
    return words_width(rows, style.cell_pad_h, self._cell_family(style), self._cell_size(style),
      bold_rows)

  def _cell_size(self, style:TableStyle) -> float:
    """
    Size (pt) of table text: `style.font_size`, else one step below the body (11→10, 14→12),
    as tables read better a little smaller.
    """
    if style.font_size is not None: return style.font_size
    return smaller_size(self._style.font_size or Defaults.FONT_SIZE)

  def _cell_family(self, style:TableStyle) -> str:
    """
    Face of table text: `style.font_family`, else `Normal`'s,
    since cell runs name no font of their own then, not the one `font()` set.
    """
    return style.font_family or self._doc.styles["Normal"].font.name or Defaults.FONT_FAMILY

  def table(
    self,
    body:list[list[str|TablePicture]],
    header:list[str|TablePicture]|None = None,
    aligns:list[str]|None = None,
    widths:list[float]|None = None, # mm per column
    style:TableStyle|None = None,
    word_style:str|None = None, # built-in Word table style name
  ) -> "DOCX":
    """
    Add a table.

    Args:
      body: List of rows, each a list of cell strings, or of a `TablePicture` for a picture.
      header: Optional header row.
      aligns: Per-column alignment - `L`/`R`/`C`/`J`.
      widths: Per-column widths in mm.
        `None` sizes columns by content to fill the content area.
        With `style.fill_content_width` off, Word sizes them itself.
      style: `TableStyle` for borders/shading. Ignored when `word_style` set.
      word_style: Built-in Word style name (e.g. `Light Grid`).
    """
    self._flush_para()
    if style is not None: check_table_rules(style.rules) # before a table is left half made
    rows = ([header] if header else []) + (body or [])
    if not rows: return self
    ncols = max(len(r) for r in rows)
    tbl = self.new_table(len(rows), ncols)
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
    eff_size_pt = self._cell_size(style)
    # Column widths: explicit `widths=`, else sized by content to fill the content area.
    # Left to `python-docx`, a narrow table drifts left.
    # Explicit widths and no autofit anchor it to the full width.
    if widths:
      col_w_mm = [to_mm(w, self.unit) for w in widths[:ncols]]
      while len(col_w_mm) < ncols:
        col_w_mm.append(self.content_width / ncols)
    elif style.fill_content_width:
      bold_rows = 1 if header and style.header_bold else 0
      col_w_mm = content_widths(rows, self.content_width, style.cell_pad_h,
        self._cell_family(style), eff_size_pt, bold_rows)
    else:
      col_w_mm = None
    # Borders first: `set_table_widths` reorders `tblPr` into canonical schema order
    # and slots these borders correctly (strict renderers need the order).
    apply_table_rules(tbl, style.border_color, style.border_size, style.rules, style.double)
    if col_w_mm:
      tbl.autofit = False
      tbl.allow_autofit = False
      set_table_widths(tbl, col_w_mm, left_pad_mm=style.cell_pad_h)
    if style.table_align:
      tbl.alignment = {"L": WD_TABLE_ALIGNMENT.LEFT, "C": WD_TABLE_ALIGNMENT.CENTER,
        "R": WD_TABLE_ALIGNMENT.RIGHT}.get(style.table_align, WD_TABLE_ALIGNMENT.LEFT)
    def room(j:int) -> float:
      """mm of content in column `j`, which a picture scales down to fit."""
      return (col_w_mm[j] if col_w_mm else self.content_width / ncols) - 2 * style.cell_pad_h
    if header:
      head_row = tbl.rows[0]
      for j, txt in enumerate(header):
        cell = head_row.cells[j]
        run = self._cell_run(cell, txt, eff_size_pt, room(j), style.font_family)
        if style.header_bold:
          run.bold = True
        if style.header_color is not None:
          run.font.color.rgb = RGBColor(*rgb255(style.header_color))
        if style.header_bg is not None: apply_cell_shading(cell, style.header_bg)
        if style.rules == "head" and body: rule_cell(cell, style.border_color, style.border_size)
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
        if j >= len(tr.cells): break
        cell = tr.cells[j]
        self._cell_run(cell, txt, eff_size_pt, room(j), style.font_family)
        if bg is not None:
          apply_cell_shading(cell, bg)
        set_cell_margins(cell, top=style.cell_pad_top, right=style.cell_pad_h,
          bot=style.cell_pad_bot, left=style.cell_pad_h)
        set_cell_vertical_align(cell, style.vertical_align)
        set_cell_align(cell, aligns[j] if aligns and j < len(aligns) else None)
        self._tighten_cell_paras(cell)
    # Cell spacing stays inside the table, so park a `space_before` bonus for the next paragraph -
    # else it sticks to the table's bottom border.
    self.track_block_spacing(0, extra_for_next=6)
    return self

  def _cell_run(self, cell, value, size_pt:float, room:float, family:str|None=None):
    """
    The run holding `value` in `cell`, emptied of its placeholder first:
    text at `size_pt`, in `family` where given, or a `TablePicture` scaled down to fit `room` mm.
    """
    cell.text = ""
    run = cell.paragraphs[0].add_run()
    if isinstance(value, TablePicture):
      scale = min(1.0, room / value.width) if value.width > 0 else 1.0
      add_picture(run, value.source, width=Mm(value.width * scale),
        height=Mm(value.height * scale))
      pin_offsets(run)
    else:
      run.text = str(value)
      run.font.size = Pt(size_pt)
      if family: run.font.name = family
    return run

  def _tighten_cell_paras(self, cell):
    """
    Strip the `Normal` style's `space_after` from cell paragraphs -
    else every cell carries phantom padding and rows look bottom-heavy.
    """
    for p in cell.paragraphs:
      pf = p.paragraph_format
      pf.space_before = Pt(0)
      pf.space_after = Pt(0)
      pf.line_spacing = CELL_LINE_SPACING
      self.keep_words(p)

  #----------------------------------------------------------------------------------------- Images

  def image(
    self, path:str, width:float|None = None, height:float|None = None,
    align:str|None = None,
    before:float = 0,
    after:float = 0,
  ) -> "DOCX":
    """
    Insert an image, `before` and `after` pt around its paragraph.
    Width/height in current unit (mm by default).
    Provide only `width` for proportional scaling.

    SVG sources are rasterized first - `python-docx` reads raster formats only.
    Raises `UnrecognizedImageError` for a format Word cannot embed.
    """
    self._flush_para()
    p = self.new_para()
    if align:
      p.alignment = align_to_docx(align)
    if before: p.paragraph_format.space_before = Pt(self.collapsed_before(before))
    if after: p.paragraph_format.space_after = Pt(after)
    run = p.add_run()
    kwargs = {}
    if width: kwargs["width"] = Mm(to_mm(width, self.unit))
    if height: kwargs["height"] = Mm(to_mm(height, self.unit))
    source = path
    if is_svg(path):
      buf = svg_to_png_buffer(path)
      if buf is not None:
        source = buf
    add_picture(run, source, **kwargs)
    # Treat image paragraph like a regular body block for collapse purposes.
    self.track_block_spacing(after)
    return self

  #------------------------------------------------------------------------------------ Page breaks

  def page_break(self) -> "DOCX":
    """Insert a hard page break."""
    self._flush_para()
    p = self.new_para()
    p.add_run().add_break(WD_BREAK.PAGE)
    return self

  #-------------------------------------------------------------------------------------- Bookmarks

  def bookmark(self, name:str) -> "DOCX":
    """
    Mark a bookmark anchor at the current point (zero-width range).
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

  #--------------------------------------------------------------------------------------- Contents

  def toc(
    self, levels:int=3, shown:str="", indent:float=6, after:float=6, kind:str="",
  ) -> "DOCX":
    """
    A table of contents of the headings of levels 1 to `levels`, as Word builds one:
    a field it fills as it opens the file, asking the reader first, `shown` until then.
    Its entries link to their headings and take Word's own `TOC 1` to `TOC 9` styles:
    each level `indent` in from the one above, `after` pt apart, single lines,
    the page number at the right edge after dots.
    Given `kind`, a letter, it lists instead what `toc_entry()` marks with it,
    as a list of figures lists their captions.
    """
    if not 1 <= levels <= 9: raise ValueError(f"levels run from 1 to 9, got {levels!r}")
    if kind: _list_letter(kind)
    styles = self._doc.styles
    for level in range(1, levels + 1):
      name = f"toc {level}"
      style = styles[name] if name in styles else styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
      # Word's own `TOC 1` it fills the contents with, not a style of that name beside it
      style.element.set(qn("w:styleId"), f"TOC{level}")
      style.element.attrib.pop(qn("w:customStyle"), None)
      style.base_style = styles["Normal"]
      pf = style.paragraph_format
      pf.line_spacing, pf.space_before, pf.space_after = 1.0, Pt(0), Pt(after)
      pf.left_indent = Mm(to_mm(indent, self.unit) * (level - 1))
      pf.tab_stops.clear_all() # one, though the contents are set twice
      pf.tab_stops.add_tab_stop(Mm(self.content_width), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    self._flush_para()
    source = f"\\f {kind} \\h \\z" if kind else f'\\o "1-{levels}" \\h \\z \\u'
    complex_field(self.new_para(), f"TOC {source}", shown)
    settings = self._doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
      update = OxmlElement("w:updateFields")
      update.set(qn("w:val"), "true")
      _insert_ordered(settings, update, _SETTINGS_ORDER)
    self.track_block_spacing(0)
    return self

  def toc_entry(self, words:str, kind:str, level:int=1) -> "DOCX":
    """
    Mark the open paragraph as an entry of the list `toc(kind=kind)` builds, `words` there
    at `level` with this paragraph's page, as a caption enters a list of figures.
    The mark is a field Word reads, and nothing of it shows.
    """
    _list_letter(kind)
    if not 1 <= level <= 9: raise ValueError(f"level runs from 1 to 9, got {level!r}")
    if self._current_para is None: self.para()
    words = words.replace("\\", "\\\\").replace('"', '\\"')
    mark_field(self._current_para, f'TC "{words}" \\f {kind} \\l {level}')
    return self

  #-------------------------------------------------------------------------------- Headers/footers

  def header(self, text:str|None=None, align:str|None=None) -> "DOCX":
    """
    Set the header text of the last section, which the sections `columns()` opens share.
    `in_header()` writes paragraphs of any look there, on the first, odd or even pages.
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
    """
    Set the footer text of the last section, which the sections `columns()` opens share.

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

  @contextmanager
  def in_header(self, pages:str="all"):
    """
    Write the `with` body into the header of the last section on `pages`: `all` its pages,
    its `first` page apart, or `even` pages apart from odd ones, which then holds for every
    section. The header stops following the section before. Its paragraph is open from the start,
    in Word's `Header` style, and `para()` adds more.
    """
    with self._band("header", pages): yield self

  @contextmanager
  def in_footer(self, pages:str="all"):
    """Write the `with` body into the footer of the last section on `pages`, as `in_header`."""
    with self._band("footer", pages): yield self

  @contextmanager
  def _band(self, kind:str, pages:str):
    """The header or footer `kind` of the last section on `pages` as the write target."""
    if pages not in ("all", "first", "even"):
      raise ValueError(f"pages are all, first or even, got {pages!r}")
    sec = self._doc.sections[-1]
    if pages == "first": sec.different_first_page_header_footer = True
    if pages == "even": self._doc.settings.odd_and_even_pages_header_footer = True
    band = getattr(sec, kind if pages == "all" else f"{pages}_page_{kind}")
    band.is_linked_to_previous = False
    self._flush_para()
    outer = (self._at, self._last_space_after, self._extra_block_after)
    self._at = _Target(band, self._page.content_width)
    self._current_para = band.paragraphs[0]
    try:
      yield self
    finally:
      self._flush_para()
      self._at, self._last_space_after, self._extra_block_after = outer

  def fields(self, template:str) -> "DOCX":
    """`template` into the open paragraph in the font set, `{page}` and `{pages}` as fields."""
    if self._current_para is None: self.para()
    page_fields(self._current_para, template, self._look)
    return self

  def _look(self, run):
    """`run` in the font, size, face and colour of what `text()` writes now."""
    _apply_run_format(run, RichSegment(text=run.text, color=self._style.color,
      bold=bool(self._style.bold), italic=bool(self._style.italic)),
      self._style.font_family, self._style.font_size)

  def section(self) -> "DOCX":
    """
    Start the pages after this one in a section of their own: `in_header()` and `in_footer()`
    set its headers and footers apart, and its first page is like the others.
    """
    self.hairline()
    self._doc.add_section(WD_SECTION.NEW_PAGE).different_first_page_header_footer = False
    return self

  def from_edge(self, header:float|None=None, footer:float|None=None) -> "DOCX":
    """How far the header and the footer of the last section stand from the page's edge."""
    sec = self._doc.sections[-1]
    if header is not None: sec.header_distance = Mm(to_mm(header, self.unit))
    if footer is not None: sec.footer_distance = Mm(to_mm(footer, self.unit))
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
    if _last_block(self._doc).tag == qn("w:tbl"): self.hairline()
    self._metadata.apply(self._doc)
    if self._emoji_font:
      for part in self._doc.part.package.iter_parts(): # body, headers, footers and footnotes
        element = getattr(part, "element", None)
        if element is not None: font_emoji(element, self._emoji_font)
    embedded = embed_fonts(self._doc, self._font_dir) if self._font_dir else []
    settings = self._doc.settings.element
    if embedded and settings.find(qn("w:embedTrueTypeFonts")) is None:
      _insert_ordered(settings, OxmlElement("w:embedTrueTypeFonts"), _SETTINGS_ORDER)
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

# CT_Compat child sequence, up to `doNotExpandShiftReturn`.
# The options after it rank last.
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
  """
  Language tag Word names `lang` by: `pl` → `pl-PL`.
  A tag with a region passes through.
  """
  return _LANG_TAGS.get(lang.lower(), lang.replace("_", "-"))

def mark_language(element, lang:str):
  """Set every run under `element` in `lang`, which Word spell-checks and hyphenates it by."""
  tag = _lang_tag(lang)
  for r in element.iter(qn("w:r")):
    _child(r.get_or_add_rPr(), "w:lang", after=_RPR_AFTER_LANG).set(qn("w:val"), tag)

def no_hyphens(element):
  """Keep the words of a paragraph or paragraph style whole under automatic hyphenation."""
  pPr = element.get_or_add_pPr()
  if pPr.find(qn("w:suppressAutoHyphens")) is None:
    _ppr_insert(pPr, OxmlElement("w:suppressAutoHyphens"))

def _list_letter(kind:str):
  """`kind` of a list of entries refused unless one letter, the name Word knows such a list by."""
  if not (len(kind) == 1 and kind.isalpha()): raise ValueError(f"kind is one letter, got {kind!r}")

def _leading(pf, lines:float|None, at_least:float|None):
  """Lines of paragraph format `pf` `lines` apart as a multiple, or `at_least` pt apart."""
  if lines is not None: pf.line_spacing = lines
  if at_least is not None:
    pf.line_spacing, pf.line_spacing_rule = Pt(at_least), WD_LINE_SPACING.AT_LEAST

def _link_run(p, run, seg:RichSegment):
  """
  `run` of paragraph `p` wrapped in a hyperlink to the URL or bookmark of `seg`,
  in Word's underlined blue where `seg` names no colour.
  """
  link = OxmlElement("w:hyperlink")
  if seg.link_url:
    link.set(qn("r:id"), p.part.relate_to(seg.link_url, RT.HYPERLINK, is_external=True))
  else: link.set(qn("w:anchor"), seg.link_target)
  run._r.addprevious(link)
  link.append(run._r)
  if seg.color is not None: return
  run._r.get_or_add_rPr().insert(0, OxmlElement("w:rStyle", {qn("w:val"): "Hyperlink"}))
  run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
  run.font.underline = True

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
  """
  Set paragraph borders on a paragraph or style element via `<w:pBdr>`.
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
  """
  Strip theme font references on a style and pin `font_name` for every script range.
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
