# tests/test_flow.py

"""Engine options: justified, hyphenated and kept text, heading looks, list delimiters, parts."""

import pytest
from lxml import etree
from docmarq import DOCX
from docmarq.md import MarkdownRenderer, MarkdownStyle, top_blocks
from conftest import NS, docx_xml

TEXT = "# Title\n\n## Section\n\nA paragraph.\n\n- one\n- two\n\n> quoted\n\n1) first\n2) second\n"

@pytest.fixture
def render(tmp_path):
  """Write `md` with `MarkdownStyle(**style)` into a fresh document, return its body."""
  def run(md:str, **style):
    doc = DOCX(str(tmp_path / "flow.docx"))
    MarkdownRenderer(doc, MarkdownStyle(**style)).render(md)
    doc.save()
    return docx_xml(doc.path)
  return run

@pytest.fixture
def para_of():
  """`w:pPr` of the paragraph holding `text`."""
  def of(root, text:str):
    return root.xpath(f"//w:p[.//w:t='{text}']/w:pPr", namespaces=NS)[0]
  return of

#---------------------------------------------------------------------------------------- Text flow

def defaults_leave_text_ragged_and_free_to_break(render):
  root = render(TEXT)
  assert not root.xpath("//w:jc[@w:val='both']|//w:keepLines", namespaces=NS)

def justify_sets_body_text_flush(render, para_of):
  root = render(TEXT, justify=True)
  for text in ("A paragraph.", "one", "quoted", "first"):
    assert para_of(root, text).xpath("string(w:jc/@w:val)", namespaces=NS) == "both", text
  assert not root.xpath("//w:p[w:pPr/w:pStyle[starts-with(@w:val, 'Heading')]]/w:pPr/w:jc",
    namespaces=NS)

def keep_together_keeps_paragraphs_and_lists_whole(render, para_of):
  root = render(TEXT, keep_together=True)
  assert para_of(root, "A paragraph.").xpath("w:keepLines", namespaces=NS)
  assert para_of(root, "one").xpath("w:keepNext", namespaces=NS)
  assert not para_of(root, "two").xpath("w:keepNext", namespaces=NS) # the list ends there

def hyphenate_turns_word_hyphenation_on(tmp_path):
  doc = DOCX(str(tmp_path / "h.docx"))
  MarkdownRenderer(doc, MarkdownStyle(hyphenate="de")).render("Text.")
  doc.save()
  assert docx_xml(doc.path, "word/settings.xml").xpath("w:autoHyphenation", namespaces=NS)
  styles = docx_xml(doc.path, "word/styles.xml")
  assert styles.xpath("string(w:docDefaults//w:lang/@w:val)", namespaces=NS) == "de-DE"

#----------------------------------------------------------------------------------------- Headings

def heading_options_restyle_the_levels(tmp_path):
  doc = DOCX(str(tmp_path / "h.docx"))
  MarkdownRenderer(doc, MarkdownStyle(h1_size=13, h2_size=None, h1_underline=False,
    h2_underline=False, head_mode="Regular"))
  doc.save()
  styles = docx_xml(doc.path, "word/styles.xml")
  def style(level:int):
    return styles.xpath(f"w:style[@w:styleId='Heading{level}']", namespaces=NS)[0]
  assert style(1).xpath("string(w:rPr/w:sz/@w:val)", namespaces=NS) == "26"
  assert style(2).xpath("string(w:rPr/w:sz/@w:val)", namespaces=NS) == "22" # body size
  assert not style(1).xpath("w:pPr/w:pBdr", namespaces=NS)
  assert style(2).xpath("string(w:rPr/w:b/@w:val)", namespaces=NS) == "0"

def head_gaps_add_to_the_paragraph_gap(render):
  root = render("Text.\n\n## Section\n\nMore.", para_gap=2, head_gap_top=2, head_gap_bot=1)
  spacing = root.xpath("//w:p[w:pPr/w:pStyle/@w:val='Heading2']/w:pPr/w:spacing", namespaces=NS)[0]
  before, after = (int(spacing.get(f"{{{NS['w']}}}{k}")) for k in ("before", "after"))
  # 4mm declared above, less the 2mm the paragraph leaves under it; 1mm below
  assert before == pytest.approx(2 / 25.4 * 1440, abs=2)
  assert after == pytest.approx(1 / 25.4 * 1440, abs=2)

@pytest.mark.parametrize("line, rules", [("---", 1), ("===", 0)])
def image_over_a_line_stays_a_figure(tmp_path, line, rules):
  # CommonMark reads it as a setext heading holding the picture; only `---` is a rule
  from PIL import Image
  Image.new("RGB", (40, 20)).save(tmp_path / "pic.png")
  doc = DOCX(str(tmp_path / "s.docx"))
  renderer = MarkdownRenderer(doc, MarkdownStyle(), base_dir=str(tmp_path))
  renderer.render(f"![diagram](pic.png)\n{line}\n")
  doc.save()
  root = docx_xml(doc.path)
  assert root.xpath("//wp:inline", namespaces=NS)
  assert len(root.xpath("//w:pBdr/w:bottom", namespaces=NS)) == rules
  assert not root.xpath("//w:pStyle[starts-with(@w:val, 'Heading')]", namespaces=NS)

#-------------------------------------------------------------------------------------------- Lists

def ordered_list_keeps_the_authors_delimiter(tmp_path):
  doc = DOCX(str(tmp_path / "l.docx"))
  MarkdownRenderer(doc, MarkdownStyle()).render("1) one\n2) two\n\n5. five\n")
  doc.save()
  numbering = docx_xml(doc.path, "word/numbering.xml")
  ids = docx_xml(doc.path).xpath("//w:numPr/w:numId/@w:val", namespaces=NS)
  def delim(num_id:str) -> str:
    return numbering.xpath(f"string(w:num[@w:numId='{num_id}']//w:lvlText/@w:val)",
      namespaces=NS)
  assert (delim(ids[0]), delim(ids[2])) == ("%1)", "") # `.` stays the style's own level

#-------------------------------------------------------------------------------------------- Parts

def parts_render_in_a_look_of_their_own(tmp_path):
  doc = DOCX(str(tmp_path / "p.docx"))
  base = MarkdownStyle()
  renderer = MarkdownRenderer(doc, base)
  tokens = renderer.parse("First.\n\nSecond.\n")
  renderer.index(tokens)
  (a, b), (c, d) = top_blocks(tokens)
  renderer.render_tokens(tokens[a:b], MarkdownStyle(font_body="Cambria", body_size=14))
  renderer.render_tokens(tokens[c:d])
  doc.save()
  root = docx_xml(doc.path)
  def font(text:str) -> tuple:
    rpr = root.xpath(f"//w:r[w:t='{text}']/w:rPr", namespaces=NS)[0]
    return (rpr.xpath("string(w:rFonts/@w:ascii)", namespaces=NS),
      rpr.xpath("string(w:sz/@w:val)", namespaces=NS))
  assert font("First.") == ("Cambria", "28")
  assert font("Second.") == ("Calibri", "22")
  assert renderer.style is base

def top_blocks_tile_the_document(tmp_path):
  renderer = MarkdownRenderer(DOCX(str(tmp_path / "t.docx")))
  tokens = renderer.parse(
    "# Head\n\nText.\n\n- a\n- b\n\n<!-- group -->\n\nOne.\n\nTwo.\n\n<!-- /group -->\n\n---\n")
  blocks = top_blocks(tokens)
  assert blocks[0][0] == 0 and blocks[-1][1] == len(tokens)
  assert all(end == start for (_, end), (start, _) in zip(blocks, blocks[1:]))
  assert [tokens[s].type for s, _ in blocks] \
    == ["heading_open", "paragraph_open", "bullet_list_open", "html_block", "hr"]

def parse_drops_frontmatter(tmp_path):
  renderer = MarkdownRenderer(DOCX(str(tmp_path / "f.docx")))
  tokens = renderer.parse("---\ntitle: Kept out\n---\n\n# Body\n")
  assert tokens[1].content == "Body"

def head_color_and_rule_thickness_reach_word(tmp_path):
  doc = DOCX(str(tmp_path / "c.docx"))
  MarkdownRenderer(doc, MarkdownStyle(head_color=(1, 0, 0), hr_thick=1)).render("# A\n\n---\n")
  doc.save()
  styles = docx_xml(doc.path, "word/styles.xml")
  assert styles.xpath("string(w:style[@w:styleId='Heading1']/w:rPr/w:color/@w:val)",
    namespaces=NS) == "FF0000"
  sizes = docx_xml(doc.path).xpath("//w:p[not(w:pPr/w:pStyle)]/w:pPr/w:pBdr/w:bottom/@w:sz",
    namespaces=NS)
  assert sizes == ["8"] # eighths of a point

def template_headings_stay_without_the_neutral_style(tmp_path):
  def heading1(neutral:bool, md:str|None) -> bytes:
    doc = DOCX(str(tmp_path / f"t{neutral}{md is None}.docx"), neutral_style=neutral)
    if md is not None: MarkdownRenderer(doc, MarkdownStyle(h1_size=30)).render(md)
    doc.save()
    styles = docx_xml(doc.path, "word/styles.xml")
    return etree.tostring(styles.xpath("w:style[@w:styleId='Heading1']", namespaces=NS)[0])
  assert heading1(False, "# Title\n") == heading1(False, None)
  assert heading1(True, "# Title\n") != heading1(True, None)

def repeated_headings_get_numbered_bookmarks(render):
  root = render("## Setup\n\nOne.\n\n## Setup\n\nSee [the second](#setup-1).\n")
  assert root.xpath("//w:bookmarkStart/@w:name", namespaces=NS) == ["setup", "setup-1"]
  assert root.xpath("//w:hyperlink/@w:anchor", namespaces=NS) == ["setup-1"]

#----------------------------------------------------------------------------------------- Up front

def engine_refuses_a_docx_in_another_unit(tmp_path):
  with pytest.raises(ValueError, match="unit"):
    MarkdownRenderer(DOCX(str(tmp_path / "u.docx"), unit="cm"))
