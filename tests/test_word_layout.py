# tests/test_word_layout.py

"""
Word's layout through the API alone: the open paragraph and its runs, the styles,
headers and footers, sections, layout tables and the styles of the contents.
"""

import pytest
from docmarq import DOCX
from conftest import NS, attrs, docx_xml, twips

W = f"{{{NS['w']}}}"

@pytest.fixture
def doc(tmp_path):
  return DOCX(str(tmp_path / "layout.docx"))

#---------------------------------------------------------------------------------------- Paragraph

def shape_sets_the_open_paragraph(doc):
  doc.para("Hanging").shape(indent=10, first=-10, before=3, after=6, at_least=14, keep=True,
    tabs=[(10, "L"), (150, "R")], align="J")
  doc.para("Lines").shape(lines=1.5)
  doc.save()
  body = docx_xml(doc.path)
  ppr = body.xpath("//w:p[w:r/w:t='Hanging']/w:pPr", namespaces=NS)[0]
  assert attrs(ppr.find(W + "ind"), "left", "hanging") == [twips(10), twips(10)]
  assert attrs(ppr.find(W + "spacing"), "before", "after", "line", "lineRule") == [
    "60", "120", "280", "atLeast"]
  assert ppr.find(W + "keepNext") is not None
  assert [attrs(t, "val", "pos") for t in ppr.find(W + "tabs")] == [
    ["left", twips(10)], ["right", twips(150)]]
  assert attrs(ppr.find(W + "jc"), "val") == ["both"]
  lines = body.xpath("//w:p[w:r/w:t='Lines']/w:pPr/w:spacing", namespaces=NS)[0]
  assert attrs(lines, "line", "lineRule") == ["360", "auto"]

def shape_page_opens_a_new_page_before_the_paragraph(doc):
  doc.para("Fresh").shape(page=True)
  doc.para("Plain").shape(keep=True)
  doc.save()
  body = docx_xml(doc.path)
  assert body.xpath("//w:p[w:r/w:t='Fresh']/w:pPr/w:pageBreakBefore", namespaces=NS)
  assert not body.xpath("//w:p[w:r/w:t='Plain']/w:pPr/w:pageBreakBefore", namespaces=NS)
  assert not body.xpath("//w:br[@w:type='page']", namespaces=NS) # no paragraph of a break

def restyle_sets_the_runs_written_so_far(doc):
  doc.heading("Plain", level=1)
  doc.restyle(bold=False, italic=True, caps=True, size=10, color=(1, 0, 0), lang="en")
  doc.text(" later")
  doc.save()
  first, later = docx_xml(doc.path).xpath("//w:p[w:r/w:t='Plain']/w:r", namespaces=NS)
  rpr = first.find(W + "rPr")
  assert attrs(rpr.find(W + "b"), "val") == ["0"]
  assert rpr.find(W + "i") is not None and rpr.find(W + "smallCaps") is not None
  assert attrs(rpr.find(W + "sz"), "val") == ["20"]
  assert attrs(rpr.find(W + "color"), "val") == ["FF0000"]
  assert attrs(rpr.find(W + "lang"), "val") == ["en-US"]
  assert later.find(f"{W}rPr/{W}smallCaps") is None and later.find(f"{W}rPr/{W}lang") is None

def small_caps_of_the_style_set_the_runs_after_it(doc):
  doc.style(small_caps=True).para("Small")
  doc.style(small_caps=False).text(" then plain")
  doc.save()
  runs = docx_xml(doc.path).xpath("//w:p[w:r/w:t='Small']/w:r", namespaces=NS)
  assert [r.find(f"{W}rPr/{W}smallCaps") is not None for r in runs] == [True, False]

def restyle_and_small_caps_reach_the_words_of_a_link(doc):
  doc.para("See ").link("the site", url="https://example.com")
  doc.restyle(caps=True)
  doc.style(small_caps=True).para("Then ").link("a part", target="part")
  doc.style(small_caps=False)
  doc.save()
  links = docx_xml(doc.path).xpath("//w:hyperlink/w:r", namespaces=NS)
  assert [r.find(f"{W}rPr/{W}smallCaps") is not None for r in links] == [True, True]

#------------------------------------------------------------------------------------------- Styles

def heading_style_takes_a_look_of_its_own(doc):
  doc.heading_style(1, size=12, bold=True, family="Arial", caps=True, upper=True, align="C",
    indent=5, page=True, lines=1.5)
  doc.save()
  style = docx_xml(doc.path, "word/styles.xml").xpath(
    "//w:style[@w:styleId='Heading1']", namespaces=NS)[0]
  rpr, ppr = style.find(W + "rPr"), style.find(W + "pPr")
  assert attrs(rpr.find(W + "sz"), "val") == ["24"]
  assert rpr.find(W + "smallCaps") is not None and rpr.find(W + "caps") is not None
  assert attrs(rpr.find(W + "rFonts"), "ascii", "hAnsi") == ["Arial", "Arial"]
  assert attrs(ppr.find(W + "jc"), "val") == ["center"]
  assert attrs(ppr.find(W + "ind"), "left") == [twips(5)]
  assert ppr.find(W + "pageBreakBefore") is not None
  assert attrs(ppr.find(W + "spacing"), "line", "lineRule") == ["360", "auto"]

def body_style_sets_the_leading_of_every_paragraph(doc):
  doc.body_style(at_least=13.2)
  doc.save()
  spacing = docx_xml(doc.path, "word/styles.xml").xpath(
    "//w:style[@w:styleId='Normal']/w:pPr/w:spacing", namespaces=NS)[0]
  assert attrs(spacing, "line", "lineRule") == ["264", "atLeast"]

def toc_entries_take_words_own_styles(doc):
  doc.toc(2)
  doc.save()
  styles = docx_xml(doc.path, "word/styles.xml")
  for level, indent in ((1, 0), (2, 6)):
    style = styles.xpath(f"//w:style[@w:styleId='TOC{level}']", namespaces=NS)
    assert len(style) == 1 and style[0].get(W + "customStyle") is None
    assert attrs(style[0].find(W + "basedOn"), "val") == ["Normal"]
    ppr = style[0].find(W + "pPr")
    assert attrs(ppr.find(f"{W}tabs/{W}tab"), "val", "leader", "pos") == [
      "right", "dot", twips(doc.content_width)]
    assert attrs(ppr.find(W + "ind"), "left") == [twips(indent)]

def toc_called_again_keeps_one_tab_to_the_page_number(doc):
  doc.toc(2)
  doc.toc(2)
  doc.save()
  styles = docx_xml(doc.path, "word/styles.xml")
  assert len(styles.xpath("//w:style[@w:styleId='TOC1']/w:pPr/w:tabs/w:tab", namespaces=NS)) == 1

#----------------------------------------------------------------------------------- Heads and feet

def headers_and_footers_take_paragraphs_as_the_body_does(doc):
  with doc.in_header("first"):
    doc.shape(align="C").text("Title page head")
  with doc.in_header():
    doc.shape(style="Normal", tabs=[(doc.content_width, "R")]).fields("Odd\t{page}")
  with doc.in_header("even"):
    doc.fields("{page}\tEven")
  with doc.in_footer("first"):
    doc.font(size=12).text("Gdynia 2027")
  doc.para("Body")
  doc.from_edge(header=12, footer=8)
  doc.save()
  body = docx_xml(doc.path)
  sect = body.xpath("//w:body/w:sectPr", namespaces=NS)[0]
  assert sect.find(W + "titlePg") is not None
  assert attrs(sect.find(W + "pgMar"), "header", "footer") == [twips(12), twips(8)]
  kinds = {ref.get(W + "type") for ref in sect.findall(W + "headerReference")}
  assert kinds == {"first", "default", "even"}
  settings = docx_xml(doc.path, "word/settings.xml")
  assert settings.find(W + "evenAndOddHeaders") is not None
  assert "Body" not in "".join(sect.xpath("//w:hdr//w:t/text()", namespaces=NS))
  import zipfile
  with zipfile.ZipFile(doc.path) as z:
    parts = {name: z.read(name).decode() for name in z.namelist() if "/header" in name
      or "/footer" in name}
  first = next(xml for xml in parts.values() if "Title page head" in xml)
  assert 'w:val="center"' in first and '<w:pStyle w:val="Header"/>' in first # the band's style
  odd = next(xml for xml in parts.values() if "Odd" in xml)
  assert "PAGE" in odd and 'w:val="right"' in odd and "w:pStyle" not in odd
  foot = next(xml for xml in parts.values() if "Gdynia 2027" in xml)
  assert 'w:sz w:val="24"' in foot

def a_section_from_a_new_page_takes_feet_of_its_own(doc):
  doc.para("Title")
  with doc.in_footer("first"):
    doc.text("Place")
  doc.section()
  with doc.in_footer():
    doc.shape(align="R").fields("{page}")
  with doc.in_footer("even"):
    doc.shape(align="L").fields("{page}")
  doc.para("Numbered")
  doc.save()
  body = docx_xml(doc.path)
  title, numbered = body.xpath("//w:sectPr", namespaces=NS)
  assert title.find(W + "titlePg") is not None
  assert [r.get(W + "type") for r in title.findall(W + "footerReference")] == ["first"]
  assert numbered.find(W + "titlePg") is None
  start = numbered.find(W + "type") # none: a new page, as Word starts a section
  assert start is None or attrs(start, "val") == ["nextPage"]
  assert {r.get(W + "type") for r in numbered.findall(W + "footerReference")} == {
    "default", "even"}

#------------------------------------------------------------------------------------------- Tables

def grid_lays_out_cells_ruled_and_spanned(doc):
  rows = doc.grid([30, 60, 40], rows=2, rule=((0, 0, 0), 0.5), height=8, middle=True)
  across = doc.span(rows[1][1], rows[1][2])
  with doc.into(rows[0][0], 30):
    doc.para("Label")
  with doc.into(across, 100):
    doc.para("Across")
  bare = doc.grid([80, 80])
  with doc.into(bare[0][1], 80):
    doc.para("Right")
  doc.save()
  ruled, plain = docx_xml(doc.path).xpath("//w:tbl", namespaces=NS)
  assert [attrs(c, "w") for c in ruled.find(W + "tblGrid")] == [[twips(30)], [twips(60)],
    [twips(40)]]
  assert attrs(ruled.find(f"{W}tblPr/{W}tblBorders/{W}top"), "val", "sz") == ["single", "4"]
  assert {tuple(attrs(h, "val", "hRule")) for h in ruled.iter(W + "trHeight")} == {
    (twips(8), "atLeast")}
  assert {v.get(W + "val") for v in ruled.iter(W + "vAlign")} == {"center"}
  assert [attrs(s, "val") for s in ruled.iter(W + "gridSpan")] == [["2"]]
  assert attrs(plain.find(f"{W}tblPr/{W}tblBorders/{W}top"), "val") == ["none"]

def keep_rows_holds_the_table_written_last(doc):
  doc.table([["a"], ["b"], ["c"]], header=["head"])
  doc.keep_rows(1)
  doc.save()
  rows = docx_xml(doc.path).xpath("//w:tbl/w:tr", namespaces=NS)
  assert all(r.find(f"{W}trPr/{W}cantSplit") is not None for r in rows)
