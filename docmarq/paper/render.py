# docmarq/paper/render.py

"""
`md_to_docx` and `md_into` for documentation: the page around the markdown content.

`md_to_docx` makes a new file. `md_into` writes into a `DOCX` that already holds pages of its own,
as a cover, and leaves saving to the caller.
"""
from markdown_it.token import Token
from ..core import DOCX
from ..constants import PageSize, A4
from ..inline import RichSegment
from ..parts import signatures, page_number
from ..md import MarkdownRenderer, split_frontmatter, docx_metadata
from .style import PaperStyle
from .banner import Banner, compact

#--------------------------------------------------------------------------------------------- Page

def md_to_docx(
  md_text:str,
  output_path:str,
  *,
  style:PaperStyle|None = None,
  page:PageSize = A4,
  margin:float|tuple = 20,
  gutter:float = 0,
  base_dir:str|None = None,
  font_dir:str|None = None,
  metadata:dict|None = None,
) -> DOCX:
  """
  Convert markdown text to a `.docx` file.

  Presentation comes from the caller, content from the document.
  `style` is used verbatim, so a caller can set any value, a `PaperStyle()` default included.

  Frontmatter is read for content only:
  `title`, `author`, `subject` and `keywords` seed the core properties,
  which `metadata=` overrides per key, and the rest feeds the banner.

  Keyword-only after `output_path`, so the parameter order cannot drift from `md_to_pdf`.

  Args:
    page: `PageSize` in mm. `page_size("a3")` resolves a preset name,
      `A4.landscape()` flips it, `PageSize(200, 250)` is a custom sheet.
    margin: mm, scalar or CSS-order 1-4 sequence.
    gutter: binding margin in mm - Word's native gutter, on the left of every page.
    base_dir: root for relative image paths. `None` uses cwd.
    font_dir: TTF root for mermaid diagram font sync.
      When set, diagrams render in `style.font_body` instead of the system default,
      and the document families also back any text drawn inside an SVG.
  """
  style = _checked(style or PaperStyle())
  doc = DOCX(output_path, width=page.width, height=page.height,
    margin=margin, gutter=gutter,
    font_body=style.font_body, font_head=style.font_head, body_size=style.body_size)
  md_into(doc, md_text, style=style, base_dir=base_dir, font_dir=font_dir, metadata=metadata)
  doc.save()
  return doc

def md_into(
  doc:DOCX,
  md_text:str,
  *,
  style:PaperStyle|None = None,
  base_dir:str|None = None,
  font_dir:str|None = None,
  metadata:dict|None = None,
):
  """
  Write markdown after the last block of `doc` as `md_to_docx` writes a new file:
  core properties, banner, compact header, page number, content and signing lines.

  For a document that opens with pages of its own, as a cover; the caller saves it.
  Header and footer belong to the whole document, so a cover carries the page number too,
  and a banner, which keeps the first page clear of the compact header, belongs on page 1.
  """
  style = _checked(style or PaperStyle())
  fm, _ = split_frontmatter(md_text)
  meta = docx_metadata(fm, metadata)
  if meta: doc.metadata(**meta)
  renderer = MarkdownRenderer(doc, style, base_dir=base_dir, font_dir=font_dir)
  banner = Banner(doc, style, fm, renderer.frontmatter_file(fm, "logo")) if fm and style.banner \
    else None
  _chrome(doc, style, fm)
  _draw(doc, renderer, style, md_text, banner)

def _checked(style) -> PaperStyle:
  """`style` when it is a `PaperStyle`. A wrong one fails here, before anything is written."""
  if not isinstance(style, PaperStyle):
    raise TypeError("style must be a PaperStyle, build one with docmarq.paper.lang_style()")
  return style

def _chrome(doc:DOCX, style:PaperStyle, fm:dict):
  """
  Footer with the page number on every page, and the compact header on pages 2+.

  Word's `titlePg` flag is section-wide:
  it splits header and footer alike into first-page and default variants.
  Page 1 goes without a header only under the banner, and then both footers get the number.
  """
  sec = doc.doc.sections[-1]
  head = (fm.get("id") or fm.get("title")) if style.banner_compact and fm else None
  face = RichSegment("", family=style.font_body, size=style.banner_meta_size,
    color=style.muted_color)
  def number(p):
    page_number(p, style.page_number_label, show_total=style.page_number_total, face=face)
  if fm and style.banner and head:
    sec.different_first_page_header_footer = True
    sec.first_page_header.paragraphs[0].text = ""
    compact(sec.header.paragraphs[0], str(head))
    if style.page_number_label:
      number(sec.first_page_footer.paragraphs[0])
      number(sec.footer.paragraphs[0])
  else:
    if head:
      compact(sec.header.paragraphs[0], str(head))
    if style.page_number_label:
      number(sec.footer.paragraphs[0])

def _draw(
  doc:DOCX, renderer:MarkdownRenderer, style:PaperStyle,
  md_text:str, banner:Banner|None,
):
  """Banner, content and signing lines, in writing order."""
  if banner:
    banner.draw()
  tokens = renderer.parse(md_text)
  title = banner.data.get("title") if banner else None
  if style.skip_dup_title and title:
    tokens = _skip_matching_h1(tokens, str(title))
  # indexed after the drop, so a link to the removed title is plain text
  renderer.index(tokens)
  renderer.render_tokens(tokens)
  if style.sign:
    face = RichSegment("", size=style.sign_size, color=style.muted_color, italic=True)
    signatures(doc, style.sign_lines(), face,
      gap=style.sign_gap, width=style.sign_w, thick=style.hr_thick)

#------------------------------------------------------------------------------------------ Helpers

def _skip_matching_h1(tokens:list[Token], title:str) -> list[Token]:
  r"""
  Drop first 3 tokens if they are `# <title>` matching the frontmatter title.
  Written or read alike: `# Sp\. z o\.o\.` and `# **Report**` match `Sp. z o.o.` and `Report`.
  A heading holding more than text, code and emphasis, as a formula, must match as written.
  """
  if len(tokens) < 3: return tokens
  t0, t1, t2 = tokens[0], tokens[1], tokens[2]
  if not (t0.type == "heading_open" and t0.tag == "h1"
      and t1.type == "inline" and t2.type == "heading_close"):
    return tokens
  children = t1.children or []
  kinds = {"text", "code_inline", "softbreak", "em_open", "em_close", "strong_open",
    "strong_close", "s_open", "s_close"}
  text = "".join(c.content for c in children if c.type in ("text", "code_inline"))
  read_alike = all(c.type in kinds for c in children) and text.strip() == title.strip()
  return tokens[3:] if read_alike or (t1.content or "").strip() == title.strip() else tokens
