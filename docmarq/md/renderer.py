# docmarq/md/renderer.py

"""
Markdown → DOCX content: what a markdown document says, written after the last block of a `DOCX`.

Walks `markdown-it-py` tokens and emits `DOCX` API calls.
Word's native layout - page breaks, keep-with-next on heading styles, paragraph spacing -
does the heavy lifting, so this renderer is a token dispatcher with an inline state machine.

Pure helpers live in sibling modules: `tokens` (attributes, closing tokens, directives),
`slug` (heading slugs and anchor resolution), `image_utils` (Pillow preprocessing and sizing),
`mermaid` (the `mmdc` CLI).
The class owns everything that touches `doc` state.

Example:
  >>> from docmarq import DOCX
  >>> from docmarq.md import MarkdownRenderer, lang_style
  >>> doc = DOCX("notes.docx")
  >>> MarkdownRenderer(doc, lang_style("pl")).render(open("notes.md").read())
  >>> doc.save()
"""
import os, re, warnings
from markdown_it import MarkdownIt
from markdown_it.token import Token
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.dollarmath import dollarmath_plugin
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.shared import Mm, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.image.exceptions import UnrecognizedImageError
from PIL import Image
from ..core import DOCX
from ..constants import Align, Defaults
from ..utils import mm_to_pt, eighths, color_hex, align_to_docx, smaller_size
from ..svg import register_fonts
from .style import MarkdownStyle
from .frontmatter import split_frontmatter
from .._warn import caller_level
from .tokens import (get_attr, find_close, CALLOUT_RE,
  is_pagebreak_directive, is_group_open_directive, is_group_close_directive)
from . import slug, image_utils, mermaid
from .math import latex_to_omath, MathConversionError, build_omath_para, pop_tag, render_math_png

#------------------------------------------------------------------------------------ Image helpers

_ROW_MIN_SLOT_MM = 20 # figures narrower than this wrap to another row

def _image_row(inline_token:Token) -> list[Token]|None:
  """Image children of a paragraph holding nothing but local images.
  Anything else gives `None`, a remote reference included: it has no file to size."""
  imgs = []
  for c in inline_token.children or []:
    if c.type == "image":
      src = get_attr(c, "src")
      if not src or src.startswith(("http://", "https://", "data:")): return None
      imgs.append(c)
    elif c.type in ("softbreak", "hardbreak"): continue
    elif c.type == "text" and not c.content.strip(): continue
    else: return None
  return imgs or None

def _row_columns(avail_mm:float, gap_mm:float) -> int:
  """How many figures fit across before each falls below `_ROW_MIN_SLOT_MM`."""
  return max(1, int((avail_mm + gap_mm) // (_ROW_MIN_SLOT_MM + gap_mm)))

def _split_rows(n:int, cols:int) -> list[int]:
  """Row sizes for `n` figures: balanced, none wider than `cols`.
  Eight figures over seven columns is `4+4`, not `7+1`."""
  rows = -(-n // cols)
  base, extra = divmod(n, rows)
  return [base + (i < extra) for i in range(rows)]

#---------------------------------------------------------------------------------- Math delimiters

_MATH_INLINE_RE = re.compile(r"\\\((.+?)\\\)")
_MATH_BLOCK_LINE_RE = re.compile(r"^(\s*)\\\[\s*(.+?)\s*\\\]\s*$")

def normalize_math_delimiters(md_text:str) -> str:
  r"""LaTeX math delimiters `\[...\]` / `\(...\)` become `$$` / `$...$`.

  The bracket forms are what LaTeX sources and AI assistants emit;
  only the dollar forms reach the parser.
  Fenced code passes through untouched and inline code spans keep their backslashes.
  """
  out: list[str] = []
  in_fence = False
  for line in md_text.split("\n"):
    stripped = line.strip()
    if in_fence:
      out.append(line)
      if stripped.startswith(("```", "~~~")): in_fence = False
      continue
    if stripped.startswith(("```", "~~~")):
      in_fence = True
      out.append(line)
      continue
    indent = line[:len(line) - len(line.lstrip())]
    if stripped in ("\\[", "\\]"):
      out.append(indent + "$$")
      continue
    m = _MATH_BLOCK_LINE_RE.match(line)
    if m:
      out.extend([m.group(1) + "$$", m.group(1) + m.group(2), m.group(1) + "$$"])
      continue
    parts = line.split("`")
    for i in range(0, len(parts), 2):
      parts[i] = _MATH_INLINE_RE.sub(lambda x: "$" + x.group(1).strip() + "$", parts[i])
    out.append("`".join(parts))
  return "\n".join(out)

#--------------------------------------------------------------------------------- MarkdownRenderer

class MarkdownRenderer:
  """Render a `markdown-it` token stream into a `DOCX`, after its last block."""
  def __init__(
    self, doc:DOCX, style:MarkdownStyle|None = None,
    base_dir:str|None = None, font_dir:str|None = None,
  ):
    """Create a renderer bound to `doc`.

    Args:
      doc: Target `DOCX` instance.
      style: Optional `MarkdownStyle`. `None` uses defaults.
      base_dir: Root for resolving relative image paths (e.g. `![alt](./img/x.png)`).
        Defaults to current working dir.
      font_dir: Optional TTF root.
        When set, mermaid diagrams render with `style.font_body`
        instead of the system default sans-serif,
        and the document families become resolvable for any text drawn inside an SVG
        (both match the rest of the document).
        Layout: `<font_dir>/<family>/<family>-<mode>.ttf`.
    """
    if doc.unit != "mm":
      raise ValueError(f"the markdown engine writes in mm: make the DOCX with unit=\"mm\", "
        f"not {doc.unit!r}")
    self.doc = doc
    self.style = style or MarkdownStyle()
    self.base_dir = base_dir or os.getcwd()
    self.font_dir = font_dir
    if font_dir:
      register_fonts(
        font_dir, self.style.font_body, self.style.font_head, self.style.font_mono,
      )
    md = MarkdownIt("commonmark", {"html": True, "breaks": False})
    md.enable(["table", "strikethrough"])
    md.use(footnote_plugin)
    # `$...$` / `$$...$$` → `math_inline` / `math_block` tokens (OMML)
    if self.style.math_enable: md.use(dollarmath_plugin)
    self._md = md
    self._list_depth = 0
    # Pre-scanned slugs for `[text](#slug)` internal links - built before the first token
    # so broken anchors render as plain text, never as dangling Word hyperlinks.
    self._known_slugs: set[str] = set() # populated by `index()`
    self._slug_uses: dict = {} # slug → times used, for numbering repeats
    self._anchors = "" # prefix of this document's bookmark names, set by `index()`
    if self.style.hyphenate:
      doc.hyphenation(self.style.hyphenate)
    if doc.neutral_style:
      self._head_styles()
    self._typeset()

  def _head_styles(self):
    """
    Word's `Heading 1..6` in the sizes, face and rules of `style`, for the whole document.
    A face without `Italic` leaves the slant Word gives a level, as to `Heading 4`.
    """
    s = self.style
    mode = s.head_mode.lower()
    sizes = (s.h1_size, s.h2_size, s.h3_size, s.h4_size, s.h5_size, s.h6_size)
    for level, size in enumerate(sizes, 1):
      self.doc.heading_style(level, size=s.body_size if size is None else size,
        bold="bold" in mode, italic=True if "italic" in mode else None, color=s.head_color,
        rule=(level == 1 and s.h1_underline) or (level == 2 and s.h2_underline))

  def _typeset(self):
    """Body typography of `style` on the document: leading, paragraph gap, inline code."""
    s = self.style
    self.doc.style(
      line_height=s.line_height,
      space_after=mm_to_pt(s.para_gap),
      code_family=s.font_mono,
      code_color=s.code_inline_color,
      code_bg=s.code_inline_bg,
    )

  #------------------------------------------------------------------------------------------ Entry

  def render(self, md_text:str):
    """Parse, index and write a whole markdown document after the last block."""
    tokens = self.parse(md_text)
    self.index(tokens)
    self._render_tokens(tokens)

  def parse(self, md_text:str) -> list[Token]:
    """
    Tokens of `md_text`, ready for `index` and `render_tokens`.

    Frontmatter is dropped, so a whole document can come in as it is.
    Reading it is the document package's business, through `split_frontmatter`.
    """
    _, md_text = split_frontmatter(md_text, warn=False)
    md_text = normalize_math_delimiters(md_text)
    return self._md.parse(md_text)

  def index(self, tokens:list[Token], prefix:str=""):
    """
    Register the heading slugs of the whole document.

    Once per document, before any part of it is written:
    a link is live only when its target is a heading here.
    A heading dropped before this call, like a repeated title, leaves links to it as plain text.
    """
    self._known_slugs = slug.collect_heading_slugs(tokens)
    self._slug_uses = {}
    self._anchors = prefix

  def render_tokens(self, tokens:list[Token], style:MarkdownStyle|None=None):
    """
    Write `tokens`, the whole parsed document or any slice of it, after the last block.

    `style` sets this part in a look of its own: its font, size, leading and paragraph gap.
    The document's state is kept: heading slugs run on across parts.
    """
    if style is None: return self._render_tokens(tokens)
    own, self.style = self.style, style
    body = self.doc._style.copy()
    try:
      self.doc.font(family=style.font_body, size=style.body_size)
      self._typeset()
      self._render_tokens(tokens)
    finally:
      self.style = own
      self.doc._style = body

  def frontmatter_file(self, fm:dict, key:str, warn:bool=True) -> str|None:
    """
    The file frontmatter `key` names, resolved against `base_dir`.
    `None` when the key is unset, or when the file is missing, which warns unless `warn` is off.
    The warning points at the first line outside `docmarq`: the one that asked for the document.
    """
    raw = fm.get(key)
    if not raw: return None
    path = image_utils.resolve_path(str(raw), self.base_dir)
    if path and os.path.isfile(path): return path
    if warn:
      warnings.warn(f"frontmatter `{key}: {raw}` not found (base_dir={self.base_dir}); "
        "written without it", RuntimeWarning, stacklevel=caller_level())
    return None

  #------------------------------------------------------------------------------------- Directives

  def _find_group_close(self, tokens:list[Token], start:int) -> int:
    """Return index of the matching `<!-- /group -->` for the open at `start`.
    Returns `len(tokens)` for unclosed groups (warns, renders to EOF).
    """
    close = _group_close(tokens, start)
    if close is not None: return close
    warnings.warn(
      "unclosed `<!-- group -->` directive - rendering to end of document",
      RuntimeWarning, stacklevel=2,
    )
    return len(tokens)

  def _render_group(self, tokens:list[Token]):
    """Render group content with Word `keep_with_next` chaining.

    Word handles pagination itself - we tell it to keep our paragraphs together
    via `<w:keepNext/>` on each paragraph except the last in the group.
    Tables embedded in the group don't get special handling
    (no clean public API for whole-table keep-together);
    their leading paragraph's `keep_with_next` keeps them tethered to it,
    which covers the common case well enough.
    """
    if not tokens: return
    self.doc._chain(lambda: self._render_tokens(tokens))

  #--------------------------------------------------------------------------------- Block dispatch

  def _render_tokens(self, tokens:list[Token]):
    i = 0
    n = len(tokens)
    while i < n:
      t = tokens[i]
      tt = t.type
      if tt == "heading_open":
        # `![alt](src)` over a `---` or `===` line parses as a setext heading holding the picture:
        # it stays a figure, and `---` a rule
        imgs = _image_row(tokens[i+1]) if t.markup and t.markup[0] in "-=" else None
        if imgs:
          self._render_images(imgs)
          if t.markup[0] == "-": self._hr()
        else:
          self._render_heading(tokens[i+1], int(t.tag[1]))
        i += 3
      elif tt == "paragraph_open":
        # An image-only paragraph becomes a figure, or a row of them;
        # mixed inline content stays a regular paragraph.
        inline = tokens[i+1]
        imgs = _image_row(inline)
        if imgs:
          self._render_images(imgs)
        else:
          self._render_paragraph(inline)
        i += 3
      elif tt == "fence" or tt == "code_block":
        info_parts = (t.info or "").strip().split(maxsplit=1) # [lang, rest_dsl]
        lang = info_parts[0] if info_parts else None
        info_rest = info_parts[1] if len(info_parts) > 1 else ""
        if lang == "mermaid": self._render_mermaid(t.content, info_rest)
        elif lang == "math": self._render_math_block(t.content)
        else:
          self.doc.code_block(
            t.content,
            language=lang,
            bg_color=self.style.code_block_bg,
            border_color=self.style.code_block_border,
            font_family=self.style.font_mono,
          )
        i += 1
      elif tt == "math_block":
        self._render_math_block(t.content)
        i += 1
      elif tt == "bullet_list_open": i = self._render_list(tokens, i, ordered=False)
      elif tt == "ordered_list_open": i = self._render_list(tokens, i, ordered=True)
      elif tt == "blockquote_open": i = self._render_blockquote(tokens, i)
      elif tt == "hr":
        self._hr()
        i += 1
      elif tt == "table_open": i = self._render_table(tokens, i)
      elif tt == "footnote_block_open": i = self._render_footnote_block(tokens, i)
      elif tt == "html_block":
        # Only directive comments are handled; arbitrary HTML is dropped.
        content = t.content or ""
        if is_pagebreak_directive(content):
          self.doc.page_break()
        elif is_group_open_directive(content):
          end_i = self._find_group_close(tokens, i)
          self._render_group(tokens[i+1:end_i])
          i = end_i + 1
          continue
        elif is_group_close_directive(content):
          warnings.warn(
            "stray `<!-- /group -->` directive (no matching open)",
            RuntimeWarning, stacklevel=2,
          )
        i += 1
      else:
        i += 1

  #--------------------------------------------------------------------------------------- Headings

  def _render_heading(self, inline_token:Token, level:int):
    """Emit a heading.
    Bold/italic/strike pass through;
    font family/size/color inherit from the Heading paragraph style
    (so inline `code` keeps heading size, just switches family).
    Registers the slug as a Word bookmark so `[text](#slug)` internal links resolve here.
    """
    self.doc.heading(level=level, **self._head_gaps())
    p = self.doc._current_para
    anchor = slug.slugify_inline(inline_token)
    if anchor:
      self.doc.bookmark(self._anchors + slug.dedupe(anchor, self._slug_uses))
    self._add_runs_to_heading(p, inline_token)

  def _hr(self):
    """Horizontal rule in the style's color and thickness."""
    self.doc.hr(color=self.style.hr_color, size_eighths=eighths(self.style.hr_thick))

  def _head_gaps(self) -> dict:
    """
    `heading()` gaps in pt from `head_gap_top` and `head_gap_bot`, none when unset.
    The gap above counts on top of the paragraph gap, which collapses into it.
    """
    s = self.style
    gaps = {}
    if s.head_gap_top is not None: gaps["space_before"] = mm_to_pt(s.para_gap + s.head_gap_top)
    if s.head_gap_bot is not None: gaps["space_after"] = mm_to_pt(s.head_gap_bot)
    return gaps

  def _add_runs_to_heading(self, p, inline_token:Token):
    """Walk inline children and add runs to a Heading paragraph.
    Inherits size/color from the Heading style; overrides font when `font_head` set."""
    state = {"bold": False, "italic": False, "strike": False}
    font_head = self.style.font_head
    for c in (inline_token.children or []):
      ct = c.type
      if ct == "text":
        run = p.add_run(c.content)
        if font_head: run.font.name = font_head
        if state["bold"]: run.bold = True
        if state["italic"]: run.italic = True
        if state["strike"]: run.font.strike = True
      elif ct == "code_inline":
        run = p.add_run(c.content)
        run.font.name = self.style.font_mono
        if state["bold"]: run.bold = True
        if state["italic"]: run.italic = True
        if state["strike"]: run.font.strike = True
      elif ct == "strong_open": state["bold"] = True
      elif ct == "strong_close": state["bold"] = False
      elif ct == "em_open": state["italic"] = True
      elif ct == "em_close": state["italic"] = False
      elif ct == "s_open": state["strike"] = True
      elif ct == "s_close": state["strike"] = False
      elif ct == "math_inline":
        # Heading math inherits the heading run size (`size_halfpt=None`)
        # so it scales with the heading rather than rendering at body size.
        _, color_hex = self._math_run_props()
        self._append_inline_math(p, c.content, None, color_hex)
      elif ct == "softbreak": p.add_run(" ")
      elif ct == "hardbreak": p.add_run("\n")

  #-------------------------------------------------------------------------------------- Paragraph

  def _render_paragraph(self, inline_token:Token):
    """Emit a regular body paragraph and stream styled runs into it."""
    self.doc.para()
    self._body_look()
    self._add_runs_to_current_para(inline_token)

  def _body_look(self):
    """Justify the body paragraph just opened, and keep it on one page, as `style` says."""
    p = self.doc._current_para
    if self.style.justify: p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if self.style.keep_together: p.paragraph_format.keep_together = True

  def _add_runs_to_current_para(self, inline_token:Token):
    """State-machine walk of inline children.
    Translates each child into one or more `doc.text()` / `doc.link()` / `doc.line_break()` calls.
    """
    state = {"bold": False, "italic": False, "strike": False,
      "link_url": None, "link_target": None}

    def kwargs():
      k = {}
      if state["bold"]: k["bold"] = True
      if state["italic"]: k["italic"] = True
      if state["strike"]: k["strike"] = True
      return k

    for c in (inline_token.children or []):
      ct = c.type
      if ct == "text":
        if state["link_target"]:
          self.doc.link(c.content, target=state["link_target"])
        elif state["link_url"]:
          self.doc.link(c.content, url=state["link_url"])
        else:
          self.doc.text(c.content, **kwargs())
      elif ct == "code_inline":
        # Inherits wrapping bold/italic/strike - e.g. `*foo `code` bar*`.
        self.doc.text(c.content, code=True, **kwargs())
      elif ct == "strong_open": state["bold"] = True
      elif ct == "strong_close": state["bold"] = False
      elif ct == "em_open": state["italic"] = True
      elif ct == "em_close": state["italic"] = False
      elif ct == "s_open": state["strike"] = True
      elif ct == "s_close": state["strike"] = False
      elif ct == "link_open":
        href = get_attr(c, "href") or ""
        # `#anchor` validated against pre-scanned slugs; unknown → plain text.
        anchor = slug.resolve_anchor(href, self._known_slugs)
        if anchor is not None:
          state["link_target"] = self._anchors + anchor
        else:
          state["link_url"] = self._resolve_link(href)
      elif ct == "link_close":
        state["link_url"] = None
        state["link_target"] = None
      elif ct == "math_inline":
        self._render_inline_math(c.content)
      elif ct == "softbreak": self.doc.text(" ")
      elif ct == "hardbreak": self.doc.line_break()
      elif ct == "image":
        # an image in a line of text leaves its alt text in italics; a figure takes a paragraph
        alt = c.content or "[image]"
        self.doc.text(alt, italic=True)
      elif ct == "footnote_ref":
        # `[^N]` → superscript `[1]` link style.
        label = (c.meta or {}).get("label") or ""
        self.doc.text(f"[{label}]", superscript=True, color=self.style.link_color)

  def _resolve_link(self, href:str) -> str|None:
    """Resolve local hrefs against `link_root`/`link_base`.
    URLs with a schema or hash-anchor are returned as-is.
    `None` means "render text but don't linkify" (for unresolvable local refs).
    """
    if not href:
      return None
    if href.startswith(("http://", "https://", "mailto:", "ftp://")):
      return href
    if href.startswith("#"):
      return None # internal anchors handled via bookmark targets
    if self.style.link_root:
      base = self.style.link_base.strip("/")
      root = self.style.link_root.rstrip("/")
      if href.startswith("/"):
        return f"{root}{href}"
      return f"{root}/{base}/{href}" if base else f"{root}/{href}"
    return None

  #------------------------------------------------------------------------------------------ Lists

  def _render_list(self, tokens:list[Token], start:int, ordered:bool) -> int:
    open_type = "ordered_list_open" if ordered else "bullet_list_open"
    close_type = "ordered_list_close" if ordered else "bullet_list_close"
    end = find_close(tokens, start, open_type, close_type)
    if self._list_depth == 0 and self.style.keep_together:
      self.doc._chain(lambda: self._render_items(tokens, start, end, ordered))
    else:
      self._render_items(tokens, start, end, ordered)
    return end + 1

  def _render_items(self, tokens:list[Token], start:int, end:int, ordered:bool):
    """Items of the list opened at `start` and closed at `end`."""
    depth = self._list_depth
    self._list_depth += 1
    # Only the first item opens a numbering instance; the rest continue it.
    num_start = _list_start(tokens[start]) if ordered else None
    delim = (tokens[start].markup or ".") if ordered else None # the author's `1.` or `1)`
    j = start + 1
    try:
      while j < end:
        if tokens[j].type == "list_item_open":
          item_end = find_close(tokens, j, "list_item_open", "list_item_close")
          self._render_list_item(tokens, j + 1, item_end, depth, delim, num_start)
          num_start = None
          j = item_end + 1
        else:
          j += 1
    finally:
      self._list_depth -= 1

  # Block tokens a list item can hold besides paragraphs and nested lists.
  # Each maps to its closing token, or to `None` when the block is a single self-contained token.
  _ITEM_BLOCK_CLOSE = {
    "fence": None,
    "code_block": None,
    "math_block": None,
    "hr": None,
    "html_block": None,
    "blockquote_open": "blockquote_close",
    "table_open": "table_close",
    "dl_open": "dl_close",
  }

  def _render_list_item(
    self, tokens:list[Token], start:int, end:int,
    depth:int, delim:str|None,
    num_start:int|None = None,
  ):
    """Render one list item.
    First inline content goes into a list-styled paragraph; nested lists recurse with deeper depth.
    Code, quotes, tables and math under the bullet go through the top-level block dispatch,
    so a list item carries the same content a document body can."""
    first_block = True
    k = start
    while k < end:
      t = tokens[k]
      tt = t.type
      if tt == "paragraph_open":
        inline = tokens[k+1]
        if first_block:
          if delim:
            self.doc.ordered(level=depth, start=num_start, delim=delim)
          else:
            self.doc.bullet(level=depth)
          self._body_look()
          self._add_runs_to_current_para(inline)
          first_block = False
        else:
          # Continuation paragraph inside a list item - plain para
          self._render_paragraph(inline)
        k += 3
      elif tt == "bullet_list_open": k = self._render_list(tokens, k, ordered=False)
      elif tt == "ordered_list_open": k = self._render_list(tokens, k, ordered=True)
      elif tt in self._ITEM_BLOCK_CLOSE:
        close = self._ITEM_BLOCK_CLOSE[tt]
        stop = k if close is None else find_close(tokens, k, tt, close)
        self._render_tokens(tokens[k:stop+1])
        k = stop + 1
      else:
        k += 1

  #------------------------------------------------------------------------------------- Blockquote

  @staticmethod
  def _spacing_for(idx:int, total:int, before:float, after:float) -> tuple[float, float]:
    """Pick `(space_before, space_after)` for paragraph `idx` in a group of `total`.
    First gets `before`, last gets `after`, inner gets 0/0 -
    so the left bar reads as one continuous block rather than stacked stripes."""
    if total <= 1:
      return (before, after)
    if idx == 0:
      return (before, 0)
    if idx == total - 1:
      return (0, after)
    return (0, 0)

  # Quote content is flattened to `inline` tokens;
  # a fence carries its text on `token.content` instead, so it would vanish.
  # Reported, not dropped silently.
  _QUOTE_UNSUPPORTED = ("fence", "code_block")

  def _collect_blockquote_inlines(self, tokens:list[Token], start:int, end:int) -> list[Token]:
    """Flat list of every `inline` token directly inside the blockquote.
    Nested blockquotes / lists / code blocks not supported -
    they'd require recursion and per-block grouping."""
    out = []
    dropped = []
    j = start + 1
    while j < end:
      t = tokens[j]
      if t.type == "inline": out.append(t)
      elif t.type in self._QUOTE_UNSUPPORTED: dropped.append(t)
      j += 1
    if dropped:
      warnings.warn(
        f"{len(dropped)} code block(s) inside a blockquote are not rendered; "
        "move them outside the `>` quote to keep the content",
        RuntimeWarning, stacklevel=2,
      )
    return out

  def _render_blockquote(self, tokens:list[Token], start:int) -> int:
    end = find_close(tokens, start, "blockquote_open", "blockquote_close")
    callout_type = self._detect_callout(tokens, start, end) # `> [!NOTE]` form
    if callout_type:
      return self._render_callout(tokens, start, end, callout_type)
    inlines = self._collect_blockquote_inlines(tokens, start, end)
    total = len(inlines)
    for idx, inline in enumerate(inlines):
      sb, sa = self._spacing_for(idx, total, 3, 3)
      self.doc.blockquote(text=None,
        border_color=self.style.quote_border,
        text_color=self.style.muted_color,
        space_before=sb, space_after=sa)
      self._body_look()
      self._add_runs_to_current_para(inline)
    return end + 1

  def _detect_callout(self, tokens:list[Token], start:int, end:int) -> str|None:
    """Callout type ('note', 'tip', ...) if the blockquote opens with `[!TYPE]`, else `None`.
    Looks at the first inline token only."""
    j = start + 1
    while j < end:
      if tokens[j].type == "inline":
        content = tokens[j].content or ""
        first_line = content.split("\n", 1)[0]
        m = CALLOUT_RE.match(first_line)
        if m:
          return m.group(1).lower()
        return None
      j += 1
    return None

  def _render_callout(
    self, tokens:list[Token], start:int, end:int,
    kind:str,
  ) -> int:
    """Render a GitHub callout: colored title row + body in muted blockquote.
    All paragraphs in the group share tight inner spacing
    so the left bar reads as a single continuous block regardless of body length.
    """
    label = getattr(self.style, f"callout_label_{kind}", kind.title())
    border_rgb, text_rgb = self.style.callout_colors.get(kind,
      (self.style.quote_border, self.style.muted_color))
    inlines = self._collect_blockquote_inlines(tokens, start, end)
    # Strip the `[!TYPE]` marker from the FIRST inline's children.
    # markdown-it puts marker + body into a single inline (soft-broken),
    # so we can't just skip the whole inline - we'd lose the body that follows the marker line.
    body_inlines = []
    if inlines:
      stripped = self._strip_callout_marker(inlines[0])
      if stripped is not None:
        body_inlines.append(stripped)
      body_inlines.extend(inlines[1:])
    total = 1 + len(body_inlines)
    sb, sa = self._spacing_for(0, total, 3, 3)
    self.doc.blockquote(text=None, border_color=border_rgb, text_color=text_rgb,
      space_before=sb, space_after=sa)
    self.doc.text(label, bold=True, color=text_rgb)
    for j, inline in enumerate(body_inlines):
      sb, sa = self._spacing_for(j + 1, total, 3, 3)
      self.doc.blockquote(text=None, border_color=border_rgb,
        text_color=self.style.muted_color,
        space_before=sb, space_after=sa)
      self._body_look()
      self._add_runs_to_current_para(inline)
    return end + 1

  @staticmethod
  def _strip_callout_marker(inline_token:Token) -> Token|None:
    """Return a shallow copy of `inline_token`
    with the leading `[!TYPE]` marker text + the softbreak right after it removed from `children`.
    Returns `None` if no real content remains after stripping.
    """
    children = inline_token.children or []
    if not children:
      return None
    new_children = list(children)
    # Drop the `[!TYPE]` text token if present; leave unchanged if it doesn't match.
    if new_children[0].type == "text" and CALLOUT_RE.match(new_children[0].content):
      new_children.pop(0)
      # Drop the following softbreak so the body doesn't start with a blank line.
      if new_children and new_children[0].type in ("softbreak", "hardbreak"):
        new_children.pop(0)
    if not new_children:
      return None
    # `Token(type, tag, nesting)` - copy remaining fields manually.
    clone = type(inline_token)(inline_token.type, inline_token.tag,
      inline_token.nesting)
    clone.children = new_children
    clone.content = inline_token.content
    clone.attrs = inline_token.attrs
    return clone

  #------------------------------------------------------------------------------------------- Math

  def _math_run_props(self) -> tuple[int, str]:
    """`(size_halfpt, color_hex)` for OMML runs matching body font size/color."""
    body_pt = self.doc._style.font_size or Defaults.FONT_SIZE
    return round(body_pt * 2), color_hex(self.style.body_color)

  def _render_inline_math(self, latex:str):
    """Inline `$...$`: append a native OMML equation to the current paragraph,
    falling back to an inline image when the LaTeX is outside the OMML subset.
    """
    if self.doc._current_para is None:
      self.doc.para()
    size_halfpt, color_hex = self._math_run_props()
    self._append_inline_math(self.doc._current_para, latex, size_halfpt, color_hex)

  def _append_inline_math(self, p, latex:str, size_halfpt:int|None, color_hex:str):
    """Append inline math to paragraph `p` (works for body paragraphs and headings).
    Native OMML first; image fallback on conversion failure.

    `size_halfpt=None` lets the math runs inherit the paragraph's run size -
    used for headings so `# Result $x^2$` sizes the formula to the heading, not the body.
    Body callers pass the explicit body size
    so a non-default `font_size` doesn't fall back to Word's 11pt Normal default.
    """
    try:
      omath = latex_to_omath(latex, size_halfpt=size_halfpt,
        color_hex=color_hex, display=False)
      p._p.append(omath)
      return
    except (MathConversionError, RecursionError):
      # `RecursionError` can escape the depth guard through nested sub-parsers.
      pass
    self._append_inline_math_image(p, latex)

  def _render_math_block(self, latex:str):
    r"""Display `$$...$$` (or ```math): a centered native OMML equation on its own paragraph,
    with an image fallback for unsupported LaTeX.
    A `\tag{...}` becomes the right-aligned equation label, Word style."""
    self.doc._flush_para()
    latex, tag = pop_tag(latex)
    size_halfpt, color_hex = self._math_run_props()
    if tag:
      try:
        omath = latex_to_omath(latex, size_halfpt=size_halfpt, color_hex=color_hex)
      except (MathConversionError, RecursionError):
        self._render_math_block_image(latex, tag)
        return
      p = self._tagged_equation_para(tag)
      # order in the XML is the order on the line: tab, equation, tab, label
      p._p.insert(list(p._p).index(p.runs[-1]._r), omath)
      self.doc._track_block_spacing(self.doc._style.space_after or 0)
      return
    try:
      omath_para = build_omath_para(latex, size_halfpt=size_halfpt,
        color_hex=color_hex, align="center")
    except (MathConversionError, RecursionError):
      self._render_math_block_image(latex)
      return
    p = self.doc._new_para()
    self.doc._apply_para_spacing(p)
    p.alignment = align_to_docx(Align.CENTER)
    p._p.append(omath_para)
    self.doc._track_block_spacing(self.doc._style.space_after or 0)

  def _tagged_equation_para(self, tag:str):
    """Paragraph carrying the classic Word equation layout:
    a center tab in the middle of the text width, a right tab at its edge,
    the label after the second tab.
    The equation itself goes between the tabs."""
    content_w = self.doc.content_width
    p = self.doc._new_para()
    self.doc._apply_para_spacing(p)
    pf = p.paragraph_format
    pf.tab_stops.add_tab_stop(Mm(content_w / 2), WD_TAB_ALIGNMENT.CENTER)
    pf.tab_stops.add_tab_stop(Mm(content_w), WD_TAB_ALIGNMENT.RIGHT)
    p.add_run("\t")
    p.add_run("\t" + tag)
    return p

  #---------------------------------------------------------------------------- Math image fallback

  def _append_inline_math_image(self, p, latex:str):
    """Render `latex` to a PNG and embed it inline in paragraph `p`,
    dropped onto the text baseline.
    Last resort (no matplotlib, or mathtext also can't parse it):
    render the LaTeX body as inline `code` - never the raw `$...$` delimiters,
    so the no-`$`-leak contract holds."""
    body_pt = self.doc._style.font_size or Defaults.FONT_SIZE
    result = render_math_png(latex, fontsize_pt=body_pt,
      color=self.style.body_color, fontset=self.style.math_fontset)
    if result is None:
      self._warn_math_fallback()
      self.doc.text(latex, code=True)
      return
    buf, _w_mm, h_mm, baseline_mm = result
    run = p.add_run()
    try:
      run.add_picture(buf, height=Mm(h_mm))
    except Exception:
      self._warn_math_fallback()
      self.doc.text(latex, code=True)
      return
    image_utils.pin_offsets(run)
    # Shift down by the descent so formula baseline aligns with text baseline.
    # `<w:position>` is in half-points; positive = raised.
    descent_pt = baseline_mm / 25.4 * 72.0
    if descent_pt > 0.1:
      self._set_run_vertical_position(run, -round(descent_pt * 2))

  def _render_math_block_image(self, latex:str, tag:str|None=None):
    """Centered block image for a display formula outside the OMML subset.
    With a `tag`, the image sits between the label tabs instead."""
    body_pt = self.doc._style.font_size or Defaults.FONT_SIZE
    # Display math is a touch larger than inline, matching LaTeX conventions.
    result = render_math_png(latex, fontsize_pt=body_pt * 1.25,
      color=self.style.body_color, fontset=self.style.math_fontset)
    if result is None:
      # Render the LaTeX body as centered code - never echo the `$$` delimiters.
      self._warn_math_fallback()
      self.doc.para(align=Align.CENTER)
      self.doc.text(latex + (f"  {tag}" if tag else ""), code=True)
      return
    buf, w_mm, h_mm, _baseline = result
    content_w = self.doc.content_width
    if w_mm > content_w: # cap to page width
      h_mm = h_mm * content_w / w_mm
      w_mm = content_w
    if tag is None:
      self._insert_picture(buf, w_mm, h_mm)
      return
    p = self._tagged_equation_para(tag)
    run = p.runs[0]
    run.add_picture(buf, height=Mm(h_mm))
    image_utils.pin_offsets(run)
    self.doc._track_block_spacing(self.doc._style.space_after or 0)

  def _set_run_vertical_position(self, run, half_points:int):
    """Set `<w:position>` on a run (raises/lowers it; half-point units)."""
    rpr = run._element.get_or_add_rPr()
    pos = rpr.find(qn("w:position"))
    if pos is None:
      pos = OxmlElement("w:position")
      rpr.append(pos)
    pos.set(qn("w:val"), str(half_points))

  def _warn_math_fallback(self):
    """Warn on total render failure.
    Message varies: matplotlib missing vs. present but unable to parse this LaTeX."""
    try:
      import matplotlib # noqa: F401
      msg = ("math formula is outside both the OMML subset and matplotlib "
        "mathtext; rendering LaTeX source as code text")
    except ImportError:
      msg = ("math formula could not be rendered (install matplotlib for the "
        "image fallback: pip install docmarq[math]); rendering LaTeX as text")
    warnings.warn(msg, RuntimeWarning, stacklevel=2)

  #----------------------------------------------------------------------------------------- Images

  def _render_images(self, imgs:list[Token]):
    """Draw an image-only paragraph: one figure, or a balanced grid of rows.
    Never hands it back: a row that cannot be prepared becomes stacked figures."""
    if len(imgs) == 1:
      self._render_block_image(get_attr(imgs[0], "src"), imgs[0])
      return
    content_w = self.doc.content_width
    gap = self.style.para_gap
    sizes = _split_rows(len(imgs), _row_columns(content_w, gap))
    slot = (content_w - gap * (sizes[0] - 1)) / sizes[0]
    start = 0
    for count in sizes:
      row = imgs[start:start + count]
      start += count
      if not self._render_image_row(row, slot):
        for img in row:
          self._render_block_image(get_attr(img, "src"), img)

  def _render_image_row(self, imgs:list[Token], slot:float) -> bool:
    """Lay one row of figures side by side, each in a `slot`-wide cell.
    All-or-nothing: a file that will not load cancels the row."""
    items = []
    try:
      for img in imgs:
        path = image_utils.resolve_path(get_attr(img, "src") or "", self.base_dir)
        if not path or not os.path.isfile(path):
          return False
        dsl = image_utils.parse_image_dsl(get_attr(img, "title"))
        items.append(self._prepare_picture(path, slot, dsl))
      self._insert_pictures(items)
    except (OSError, ValueError, UnrecognizedImageError):
      return False
    return True

  def _render_block_image(self, src:str, img:Token):
    """Embed an image as a centered block paragraph.
    Falls back to italic alt text when the file is missing.
    Title DSL `![alt](src "key=val")` controls sizing and alignment
    (see `image_utils.parse_image_dsl`).
    """
    alt = get_attr(img, "alt") or ""
    if not alt and img.children:
      alt = "".join(c.content for c in img.children if c.type == "text")
    dsl = image_utils.parse_image_dsl(get_attr(img, "title"))
    path = image_utils.resolve_path(src, self.base_dir)
    if not path or not os.path.isfile(path):
      self.doc.para()
      self.doc.text(alt or f"[image: {src}]", italic=True)
      return
    if not self._try_insert_image(path, dsl=dsl):
      self.doc.para()
      self.doc.text(alt or f"[image: {src}]", italic=True)

  def _try_insert_image(
    self, path:str, dsl:image_utils.ImageDSL|None = None,
    generated:bool = False,
  ) -> bool:
    """Insert image scaled per DSL overrides with `style.image_max_h` cap.

    Every image goes through Pillow to:
      1. Crop transparent padding - Word renders the full transparent canvas,
         causing visible right-shift / shrunk content.
      2. Normalize format - JPEGs with APP2 ICC profile are rejected by python-docx;
         re-saving as PNG fixes it.
      3. Rasterize SVG (python-docx has no SVG support).

    `dsl.align` is applied to the paragraph after the insert.
    """
    dsl = dsl or image_utils.ImageDSL()
    try:
      item = self._prepare_picture(path, self.doc.content_width, dsl, generated)
      self._insert_pictures([item], dsl.align or "C")
    except (OSError, ValueError, UnrecognizedImageError):
      return False
    return True

  def _prepare_picture(
    self, path:str, content_w:float, dsl,
    generated:bool = False,
  ) -> tuple:
    """`(src, width_mm, height_mm)` ready for `_insert_pictures`.
    `src` is the preprocessed buffer when Pillow could open the file, else the path.
    `generated` marks a raster made here from vector source,
    which carries no resolution the width has to respect."""
    # The buffer is a re-save - DPI comes from the source. Vector has none.
    dpi = None if generated or image_utils.is_svg(path) else image_utils.read_dpi(path)
    buf = image_utils.preprocess_to_buffer(path)
    src = path if buf is None else buf
    w, h = self._compute_dims(src, content_w, self.style.image_max_h, dsl, dpi)
    if buf is not None: buf.seek(0)
    return src, w, h

  def _compute_dims(
    self, src, content_w:float, max_h:float,
    dsl:image_utils.ImageDSL, dpi:float|None,
  ) -> tuple[float|None, float]:
    """Read natural dims from a path or buffer and apply DSL overrides."""
    seek = getattr(src, "seek", None)
    try:
      if seek: seek(0)
      with Image.open(src) as im:
        nat_w, nat_h = im.size
      if seek: seek(0)
    except (OSError, ValueError):
      return (None, max_h)
    return image_utils.apply_dsl_dims(nat_w, nat_h, content_w, max_h, dsl,
      dpi, self.style.image_min_dpi)

  #---------------------------------------------------------------------------------------- Mermaid

  def _render_mermaid(self, source:str, info_rest:str=""):
    """Render mermaid via `mmdc` (or `mermaid.ink`) to PNG and embed.
    `info_rest` is parsed as image DSL. Falls back to a code block on failure."""
    s = self.style
    if not s.mermaid_enable:
      self._fallback_mermaid_code(source)
      return
    png_path = mermaid.compile_to_png(source,
      cli=s.mermaid_cli, theme=s.mermaid_theme,
      background=s.mermaid_background, scale=s.mermaid_scale,
      font_family=s.font_body, font_dir=self.font_dir,
      remote=s.mermaid_remote,
    )
    dsl = image_utils.parse_image_dsl(info_rest) if info_rest else None
    if png_path is None or not self._try_insert_image(png_path, dsl=dsl, generated=True):
      self._fallback_mermaid_code(source)

  def _fallback_mermaid_code(self, source:str):
    """Render mermaid source as a regular fenced code block -
    used when `mmdc` is unavailable or the diagram fails to compile."""
    self.doc.code_block(
      source,
      language="mermaid",
      bg_color=self.style.code_block_bg,
      border_color=self.style.code_block_border,
      font_family=self.style.font_mono,
    )

  def _insert_picture(self, src, width_mm:float|None, height_mm:float|None):
    """Centered paragraph + picture run with explicit dimensions.
    `src` may be a path or `BytesIO`.
    Bypasses `doc.image()` to pass both width and height simultaneously
    after scaled dims are computed.
    """
    self._insert_pictures([(src, width_mm, height_mm)])

  def _insert_pictures(self, items:list, align:str="C"):
    """One paragraph holding every `(src, width_mm, height_mm)` given, separated by a space run,
    aligned `L`, `C` or `R`."""
    self.doc._flush_para()
    p = self.doc._new_para()
    p.alignment = align_to_docx(align)
    pf = p.paragraph_format
    pf.left_indent = Mm(0)
    pf.right_indent = Mm(0)
    pf.first_line_indent = Mm(0)
    pf.space_before = Pt(self.doc._collapsed_before(2))
    pf.space_after = Pt(2)
    for n, (src, width_mm, height_mm) in enumerate(items):
      if n: p.add_run(" ")
      run = p.add_run()
      kwargs = {}
      if width_mm is not None:
        kwargs["width"] = Mm(width_mm)
      if height_mm is not None:
        kwargs["height"] = Mm(height_mm)
      run.add_picture(src, **kwargs)
      image_utils.pin_offsets(run)
    self.doc._track_block_spacing(2)

  #--------------------------------------------------------------------------------- Footnote block

  def _render_footnote_block(self, tokens:list[Token], start:int) -> int:
    """Render the footnote bibliography section appended by markdown-it.
    Emits an H2 heading (when `style.footnote_label` is set) or a thin HR,
    then one paragraph per footnote prefixed with `[N]`.
    Runs are one pt smaller than body (`smaller_size`)
    so the bibliography reads as compact reference material.
    Heading keeps its style-driven size.
    """
    end = find_close(tokens, start, "footnote_block_open", "footnote_block_close")
    if self.style.footnote_label:
      self.doc.heading(self.style.footnote_label, level=2, **self._head_gaps())
    else:
      self._hr()
    body_pt = self.doc._style.font_size or Defaults.FONT_SIZE
    biblio_pt = smaller_size(body_pt)
    self.doc.font(size=biblio_pt)
    try:
      j = start + 1
      while j < end:
        if tokens[j].type == "footnote_open":
          item_end = find_close(tokens, j, "footnote_open", "footnote_close")
          label = (tokens[j].meta or {}).get("label") or "?"
          self._render_footnote_item(tokens, j + 1, item_end, label)
          j = item_end + 1
        else:
          j += 1
    finally:
      self.doc.font(size=body_pt)
    return end + 1

  def _render_footnote_item(
    self, tokens:list[Token], start:int, end:int,
    label:str,
  ):
    """Render a single footnote item:
    a paragraph beginning with `[label] ` followed by the footnote's inline content.
    Multi-paragraph footnotes emit additional plain paragraphs (no further prefix).
    """
    first = True
    k = start
    while k < end:
      t = tokens[k]
      if t.type == "paragraph_open":
        inline = tokens[k+1]
        if first:
          self.doc.para()
          self._body_look()
          self.doc.text(f"[{label}] ", bold=True)
          self._add_runs_to_current_para(inline)
          first = False
        else:
          self._render_paragraph(inline)
        k += 3
      elif t.type == "footnote_anchor":
        k += 1 # back-reference decoration; no DOCX equivalent
      else:
        k += 1

  #------------------------------------------------------------------------------------------ Table

  def _render_table(self, tokens:list[Token], start:int) -> int:
    end = find_close(tokens, start, "table_open", "table_close")
    header_cells: list[str] = []
    body_rows: list[list[str]] = []
    aligns: list[str] = []
    in_header = False
    current_row: list[str] = []
    j = start + 1
    while j < end:
      t = tokens[j]
      tt = t.type
      if tt == "thead_open": in_header = True
      elif tt == "thead_close": in_header = False
      elif tt == "tr_open": current_row = []
      elif tt == "tr_close":
        if in_header: header_cells = current_row
        else: body_rows.append(current_row)
      elif tt in ("th_open", "td_open"):
        if tt == "th_open":
          style_attr = get_attr(t, "style") or ""
          if "center" in style_attr: aligns.append(Align.CENTER)
          elif "right" in style_attr: aligns.append(Align.RIGHT)
          else: aligns.append(Align.LEFT)
        close_type = "th_close" if tt == "th_open" else "td_close"
        k = j + 1
        cell_text = ""
        while k < end and tokens[k].type != close_type:
          if tokens[k].type == "inline": cell_text = self._inline_to_plain(tokens[k])
          k += 1
        current_row.append(cell_text)
        j = k
      j += 1
    self.doc.table(
      header=header_cells if header_cells else None,
      body=body_rows,
      aligns=aligns if aligns else None,
    )
    return end + 1

  #---------------------------------------------------------------------------------------- Helpers

  @staticmethod
  def _inline_to_plain(inline_token:Token) -> str:
    """Flatten an inline token into a plain string for cells/short snippets.

    A cell holds text, so formatting that needs its own XML does not survive:
    a link keeps its words and loses the href, an image leaves its alt text,
    and math stays in `$...$` notation.
    Footnote references keep their `[n]` marker
    so the note at the end of the document still has something pointing at it.
    """
    parts = []
    for c in (inline_token.children or []):
      if c.type in ("text", "code_inline"):
        parts.append(c.content)
      elif c.type == "footnote_ref":
        parts.append(f"[{(c.meta or {}).get('label') or ''}]")
      elif c.type == "math_inline":
        # Tables take plain strings; preserve LaTeX in `$...$` notation.
        parts.append(f"${c.content}$")
      elif c.type == "image":
        # A cell holds a string, so the picture cannot travel - its alt can
        alt = (c.content or "").strip()
        parts.append(alt or "[image]")
      elif c.type == "softbreak": parts.append(" ")
      elif c.type == "hardbreak": parts.append("\n")
    return "".join(parts)

#------------------------------------------------------------------------------------------ Helpers

def _list_start(token:Token) -> int:
  """Opening number of an ordered list from its `start` attribute.

  CommonMark lets a list open at any number ("5.");
  markdown-it puts it on the token only when it differs from 1.
  """
  raw = get_attr(token, "start")
  if raw is None:
    return 1
  try:
    return max(1, int(raw))
  except (TypeError, ValueError):
    return 1

def _group_close(tokens:list[Token], start:int) -> int|None:
  """
  Index of the `<!-- /group -->` closing the group opened at `start`, `None` when unclosed.
  Depth-tracked, so nested group directives don't confuse the outer close.
  """
  depth = 1
  for j in range(start + 1, len(tokens)):
    t = tokens[j]
    if t.type != "html_block": continue
    content = t.content or ""
    if is_group_open_directive(content): depth += 1
    elif is_group_close_directive(content):
      depth -= 1
      if depth == 0: return j
  return None

def top_blocks(tokens:list[Token]) -> list[tuple[int, int]]:
  """
  `(start, end)` of every top-level block, so `tokens[start:end]` renders it alone.

  A block runs from an opening token to its close, or is a single leaf token.
  A `<!-- group -->` through its `<!-- /group -->` counts as one block.
  """
  blocks = []
  i = 0
  while i < len(tokens):
    t = tokens[i]
    if t.type == "html_block" and is_group_open_directive(t.content or ""):
      close = _group_close(tokens, i)
      end = len(tokens) if close is None else close + 1
    elif t.nesting == 1:
      # matched by type and depth: the footnote plugin opens each note at level 0 too
      end = find_close(tokens, i, t.type, t.type.replace("_open", "_close")) + 1
    else:
      end = i + 1
    blocks.append((i, end))
    i = end
  return blocks
