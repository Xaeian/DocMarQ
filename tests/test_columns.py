# tests/test_columns.py

"""Core: continuous sections with columns, the write target `into()`, fields, hyphenation."""

import pytest
from docmarq import DOCX
from docmarq.fields import page_fields_if_many
from conftest import NS, docx_xml

@pytest.fixture
def doc(tmp_path):
  return DOCX(str(tmp_path / "core.docx"))

#------------------------------------------------------------------------------------------ Columns

def columns_open_a_continuous_section_and_close_it(doc):
  doc.para("Above.")
  doc.columns(2, gap=8)
  assert doc.content_width == pytest.approx((170 - 8) / 2)
  doc.para("Inside.")
  doc.columns(1)
  doc.para("Below.")
  doc.save()
  sections = docx_xml(doc.path).xpath("//w:sectPr", namespaces=NS)
  assert [s.xpath("string(w:cols/@w:num)", namespaces=NS) for s in sections] == ["", "2", "1"]
  assert [s.xpath("string(w:type/@w:val)", namespaces=NS) for s in sections][1:] \
    == ["continuous", "continuous"]
  assert sections[1].xpath("string(w:cols/@w:space)", namespaces=NS) == str(round(8 * 1440 / 25.4))

def section_ends_at_the_last_paragraph(doc):
  doc.para("Above.")
  doc.columns(2)
  doc.save()
  root = docx_xml(doc.path)
  ending = root.xpath("//w:p[w:pPr/w:sectPr]", namespaces=NS)
  assert len(ending) == 1 and ending[0].xpath("string(.)", namespaces=NS) == "Above."

def section_after_a_table_gets_a_hairline_paragraph(doc):
  doc.table([["a", "b"]])
  doc.columns(2)
  doc.save()
  ending = docx_xml(doc.path).xpath("//w:p[w:pPr/w:sectPr]", namespaces=NS)[0]
  assert ending.xpath("string(w:pPr/w:spacing/@w:lineRule)", namespaces=NS) == "exact"

def new_section_inherits_headers_and_has_no_first_page(doc):
  sec = doc.doc.sections[0]
  sec.different_first_page_header_footer = True
  sec.footer.paragraphs[0].text = "foot"
  doc.para("Above.")
  doc.columns(2)
  doc.save()
  last = docx_xml(doc.path).xpath("/w:document/w:body/w:sectPr", namespaces=NS)[0]
  assert not last.xpath("w:footerReference|w:headerReference|w:titlePg", namespaces=NS)

def same_columns_add_no_section(doc):
  doc.columns(1).columns(1)
  doc.save()
  assert len(docx_xml(doc.path).xpath("//w:sectPr", namespaces=NS)) == 1

def columns_refuse_a_bad_count_and_a_cell(doc):
  with pytest.raises(ValueError):
    doc.columns(0)
  tbl = doc.new_table(1, 1)
  with doc.into(tbl.rows[0].cells[0], 50), pytest.raises(RuntimeError):
    doc.columns(2)

def geometry_reaches_every_section(doc):
  doc.para("Above.")
  doc.columns(2)
  doc.margin(left=30)
  doc.save()
  lefts = docx_xml(doc.path).xpath("//w:sectPr/w:pgMar/@w:left", namespaces=NS)
  assert set(lefts) == {str(round(30 * 1440 / 25.4))}

#------------------------------------------------------------------------------------- Write target

def into_writes_in_the_cell_and_returns(doc):
  tbl = doc.new_table(1, 2)
  with doc.into(tbl.rows[0].cells[0], 60):
    assert doc.content_width == 60
    doc.para("first")
    doc.para("second")
  doc.para("after")
  assert doc.content_width == 170
  doc.save()
  root = docx_xml(doc.path)
  cell = root.xpath("//w:tc", namespaces=NS)[0]
  assert [p.xpath("string(.)", namespaces=NS) for p in cell.xpath("w:p", namespaces=NS)] \
    == ["first", "second"]
  assert root.xpath("string(/w:document/w:body/w:p[last()])", namespaces=NS) == "after"

def table_first_in_a_cell_takes_the_empty_paragraph(doc):
  tbl = doc.new_table(1, 1)
  with doc.into(tbl.rows[0].cells[0], 80):
    doc.table([["x", "y"]])
    doc.para("under")
  doc.save()
  cell = docx_xml(doc.path).xpath("/w:document/w:body/w:tbl/w:tr/w:tc", namespaces=NS)[0]
  kids = [el.tag.split("}")[1] for el in cell if not el.tag.endswith("tcPr")]
  assert kids == ["tbl", "p"]
  assert cell.xpath("string(w:p)", namespaces=NS) == "under"

def table_in_a_cell_is_sized_to_it(doc):
  tbl = doc.new_table(1, 1)
  with doc.into(tbl.rows[0].cells[0], 80):
    doc.table([["x", "y"]])
  doc.save()
  inner = docx_xml(doc.path).xpath("//w:tc//w:tbl/w:tblPr/w:tblW/@w:w", namespaces=NS)
  assert int(inner[0]) == pytest.approx(80 * 1440 / 25.4, abs=2)

#----------------------------------------------------------------------------------- Words and text

def hyphenation_sets_the_language_and_spares_headings(doc):
  doc.hyphenation("pl")
  doc.heading("Tytuł", level=2)
  doc.para("tekst")
  doc.table([["komórka"]])
  doc.code_block("x = 1")
  doc.save()
  settings = docx_xml(doc.path, "word/settings.xml")
  assert settings.xpath("w:autoHyphenation", namespaces=NS)
  order = [el.tag.split("}")[1] for el in settings]
  assert order.index("autoHyphenation") == order.index("defaultTabStop") + 1
  styles = docx_xml(doc.path, "word/styles.xml")
  assert styles.xpath("string(w:docDefaults//w:lang/@w:val)", namespaces=NS) == "pl-PL"
  assert styles.xpath("w:style[@w:styleId='Heading2']/w:pPr/w:suppressAutoHyphens",
    namespaces=NS)
  root = docx_xml(doc.path)
  assert root.xpath("//w:tc//w:suppressAutoHyphens", namespaces=NS)
  assert len(root.xpath("//w:p[w:pPr/w:suppressAutoHyphens]", namespaces=NS)) == 2 # cell, code
  assert not root.xpath("//w:p[.//w:t='tekst']/w:pPr/w:suppressAutoHyphens", namespaces=NS)

def justified_line_ending_in_a_break_stays_flush_left(doc):
  # Word stretches it unless its compatibility option says otherwise; pdfmarq leaves it ragged
  doc.para("one").save()
  compat = docx_xml(doc.path, "word/settings.xml").xpath("w:compat", namespaces=NS)[0]
  assert [el.tag.split("}")[1] for el in compat][:2] == ["doNotExpandShiftReturn", "useFELayout"]

def small_caps_reach_the_run(doc):
  doc.text("Umowa", small_caps=True)
  doc.save()
  assert docx_xml(doc.path).xpath("//w:r[w:t='Umowa']/w:rPr/w:smallCaps", namespaces=NS)

def heading_gaps_override_the_level(doc):
  doc.heading("A", level=2, space_before=20, space_after=5)
  doc.save()
  spacing = docx_xml(doc.path).xpath("//w:p/w:pPr/w:spacing", namespaces=NS)[0]
  assert (spacing.get(f"{{{NS['w']}}}before"), spacing.get(f"{{{NS['w']}}}after")) \
    == ("400", "100")

def heading_style_sets_size_face_and_rule(doc):
  doc.heading_style(3, size=15, bold=False, italic=True, rule=True)
  doc.heading_style(1, rule=False)
  doc.save()
  styles = docx_xml(doc.path, "word/styles.xml")
  h3 = styles.xpath("w:style[@w:styleId='Heading3']", namespaces=NS)[0]
  assert h3.xpath("string(w:rPr/w:sz/@w:val)", namespaces=NS) == "30"
  assert h3.xpath("string(w:rPr/w:b/@w:val)", namespaces=NS) == "0"
  assert h3.xpath("w:rPr/w:i", namespaces=NS) and h3.xpath("w:pPr/w:pBdr/w:bottom", namespaces=NS)
  assert not styles.xpath("w:style[@w:styleId='Heading1']/w:pPr/w:pBdr", namespaces=NS)

def ordered_list_prints_its_delimiter(doc):
  doc.ordered("one", start=1, delim=")")
  doc.ordered("two")
  doc.save()
  numbering = docx_xml(doc.path, "word/numbering.xml")
  num_id = docx_xml(doc.path).xpath("string(//w:numId/@w:val)", namespaces=NS)
  text = numbering.xpath(f"string(w:num[@w:numId='{num_id}']//w:lvl/w:lvlText/@w:val)",
    namespaces=NS)
  assert text == "%1)"

#------------------------------------------------------------------------------------------- Fields

def page_number_field_shows_only_on_many_pages(doc):
  p = doc.doc.sections[0].footer.paragraphs[0]
  page_fields_if_many(p, 'Page {page} of "{pages}"')
  doc.save()
  footer = docx_xml(doc.path, "word/footer1.xml")
  instr = "".join(footer.xpath("//w:instrText/text()", namespaces=NS))
  assert instr == ' IF  NUMPAGES  > 1 "Page  PAGE  of \\" NUMPAGES \\"" "" '
  marks = footer.xpath("//w:fldChar/@w:fldCharType", namespaces=NS)
  assert marks.count("begin") == marks.count("end") == 4
  # the stored result is empty, so a reader that never updates fields shows nothing
  last_separate = max(i for i, m in enumerate(marks) if m == "separate")
  assert marks[last_separate + 1:] == ["end"]

#--------------------------------------------------------------------------------------------- Body

def tables_in_a_row_stay_apart_and_the_body_ends_on_a_paragraph(doc):
  doc.table([["a"]])
  doc.table([["b"]])
  doc.save()
  body = docx_xml(doc.path).xpath("/w:document/w:body/*", namespaces=NS)
  assert [el.tag.split("}")[1] for el in body] == ["tbl", "p", "tbl", "p", "sectPr"]
  for p in (body[1], body[3]):
    assert p.xpath("string(w:pPr/w:spacing/@w:lineRule)", namespaces=NS) == "exact"

def unused_cell_paragraph_is_a_hairline(doc):
  tbl = doc.new_table(1, 1)
  with doc.into(tbl.rows[0].cells[0], 50):
    pass
  doc.save()
  cell = docx_xml(doc.path).xpath("/w:document/w:body/w:tbl/w:tr/w:tc", namespaces=NS)[0]
  assert cell.xpath("string(w:p/w:pPr/w:spacing/@w:lineRule)", namespaces=NS) == "exact"

def border_writes_its_sides_in_schema_order(doc):
  doc.para("ruled")
  doc.border("#1f2328", sides=("right", "bottom", "top", "left"), size=1, space=4)
  doc.save()
  pbdr = docx_xml(doc.path).xpath("//w:pBdr", namespaces=NS)[0]
  assert [el.tag.split("}")[1] for el in pbdr] == ["top", "left", "bottom", "right"]
  assert {(el.get(f"{{{NS['w']}}}sz"), el.get(f"{{{NS['w']}}}space")) for el in pbdr} \
    == {("8", "4")}

def gutter_narrows_the_content(tmp_path):
  assert DOCX(str(tmp_path / "g.docx"), gutter=10).content_width == pytest.approx(160)
