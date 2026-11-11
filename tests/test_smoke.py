# tests/test_smoke.py

"""
End-to-end smoke: the public API writes a valid DOCX to disk.
No crash, and the ZIP magic header is there.
"""

import pytest
from conftest import assert_valid_docx
from docmarq import DOCX, Align, Styles, TableStyle

#--------------------------------------------------------------------------------------- DOCX basic

def docx_empty_produces_valid_file(tmp_path):
  path = tmp_path / "empty.docx"
  with DOCX(str(path)) as doc:
    doc.para("hi")
  assert_valid_docx(path)

def docx_paragraphs_and_runs(tmp_path):
  path = tmp_path / "runs.docx"
  with DOCX(str(path)) as doc:
    doc.text("plain ").text("bold", bold=True).text(" italic", italic=True).enter()
    doc.text("code", code=True)
  assert_valid_docx(path)

def docx_headings_all_levels(tmp_path):
  path = tmp_path / "head.docx"
  with DOCX(str(path)) as doc:
    for level in range(1, 7):
      doc.heading(f"Heading level {level}", level=level)
      doc.para(f"Body under h{level}.")
  assert_valid_docx(path)

def docx_table_simple_with_header(tmp_path):
  path = tmp_path / "tab.docx"
  with DOCX(str(path)) as doc:
    doc.table(
      [["a1", "b1"], ["a2", "b2"]],
      header=["A", "B"],
      aligns=[Align.LEFT, Align.RIGHT],
    )
  assert_valid_docx(path)

def docx_lists_bullet_and_ordered(tmp_path):
  path = tmp_path / "lists.docx"
  with DOCX(str(path)) as doc:
    doc.bullet("one").bullet("two").bullet("three")
    doc.ordered("first").ordered("second")
  assert_valid_docx(path)

def docx_blockquote_hr_code(tmp_path):
  path = tmp_path / "blocks.docx"
  with DOCX(str(path)) as doc:
    doc.blockquote("Quote text.")
    doc.hr()
    doc.code_block("def foo():\n  return 42", language="python")
  assert_valid_docx(path)

def docx_colors_hex_and_tuple(tmp_path):
  # `parse_color` returns floats 0-1 in both libs.
  # `rgb255` converts at the OOXML boundary.
  # Mixing hex strings and float tuples in the same document must produce no crash.
  path = tmp_path / "colors.docx"
  with DOCX(str(path)) as doc:
    doc.text("hex", color="#1f2328").enter()
    doc.text("tuple", color=(0.03, 0.41, 0.85))
  assert_valid_docx(path)

def docx_metadata(tmp_path):
  path = tmp_path / "meta.docx"
  with DOCX(str(path)) as doc:
    doc.metadata(title="T", author="X", subject="S", keywords="k")
    doc.para("body")
  assert_valid_docx(path)

def docx_without_metadata_names_no_tool_and_is_dated_now(tmp_path):
  import docx
  from datetime import datetime, timezone
  path = tmp_path / "bare.docx"
  with DOCX(str(path)) as doc:
    doc.para("body")
  props = docx.Document(str(path)).core_properties
  assert (props.author, props.comments, props.title) == ("", "", "")
  assert props.created.year == datetime.now(timezone.utc).year # not the template's 2013

def a_template_of_the_users_keeps_its_own_properties(tmp_path):
  import docx
  template = docx.Document()
  template.core_properties.author = "Template Author"
  template.save(str(tmp_path / "template.docx"))
  path = tmp_path / "made.docx"
  with DOCX(str(path), template=str(tmp_path / "template.docx")) as doc:
    doc.para("body")
  assert docx.Document(str(path)).core_properties.author == "Template Author"

def drop_cap_sets_the_first_letter_in_a_frame_of_its_own(tmp_path):
  from conftest import NS, docx_xml
  path = str(tmp_path / "drop.docx")
  with DOCX(path) as doc:
    doc.para("Hello world").drop_cap(30, 12, sink=2.5, caps=True)
    doc.para("[1] opens with no letter").drop_cap(30, 12)
    doc.para().link("Linked", url="https://example.com").text(" words after").drop_cap(30, 12)
  root = docx_xml(path)
  frames = root.xpath("//w:p[w:pPr/w:framePr/@w:dropCap='drop']", namespaces=NS)
  assert [frame.xpath("string(.)") for frame in frames] == ["H"]
  assert frames[0].xpath("string(w:pPr/w:spacing/@w:line)", namespaces=NS) == "480"
  assert frames[0].xpath("string(w:r/w:rPr/w:position/@w:val)", namespaces=NS) == "-5"
  assert frames[0].getnext().xpath("string(.)") == "ELLO world"

def drop_cap_wears_the_face_of_its_word_and_skips_a_picture(tmp_path):
  from PIL import Image
  from conftest import NS, docx_xml
  Image.new("RGB", (8, 8)).save(tmp_path / "dot.png")
  path = str(tmp_path / "face.docx")
  with DOCX(path) as doc:
    doc.para().text("Bold opening", bold=True, color="#AA0000").drop_cap(30, 12)
    doc.para()
    doc.paragraph.add_run().add_picture(str(tmp_path / "dot.png"))
    doc.text(" words after a picture").drop_cap(30, 12)
  frames = docx_xml(path).xpath("//w:p[w:pPr/w:framePr]", namespaces=NS)
  assert [frame.xpath("string(.)") for frame in frames] == ["B"] # none after the picture
  face = frames[0].xpath("w:r/w:rPr", namespaces=NS)[0]
  assert face.xpath("w:b", namespaces=NS)
  assert face.xpath("string(w:color/@w:val)", namespaces=NS) == "AA0000"
  assert [c.tag.split("}")[1] for c in face if c.tag.split("}")[1] in ("position", "sz")] == [
    "position", "sz"]

def keep_rows_keeps_each_row_whole_and_the_ends_together(tmp_path):
  from docmarq.tables import keep_rows
  from conftest import NS, docx_xml
  path = str(tmp_path / "rows.docx")
  with DOCX(path) as doc:
    doc.table([[str(k)] for k in range(6)], header=["h"])
    keep_rows(doc.doc.tables[-1], 2)
  rows = docx_xml(path).xpath("//w:tbl/w:tr", namespaces=NS)
  assert all(row.xpath("w:trPr/w:cantSplit", namespaces=NS) for row in rows)
  kept = [bool(row.xpath(".//w:keepNext", namespaces=NS)) for row in rows]
  assert kept == [True, True, False, False, False, True, False]

def a_table_with_unknown_rules_refuses_before_it_is_made(tmp_path):
  doc = DOCX(str(tmp_path / "rules.docx"))
  with pytest.raises(ValueError, match="table rules"):
    doc.table([["a"]], style=TableStyle(rules="box"))
  assert not doc.doc.tables

#------------------------------------------------------------------------------------------- Styles

def style_presets_are_independent_copies():
  # each access must yield a fresh `Style` so users can't mutate shared state
  a = Styles.BOLD
  b = Styles.BOLD
  assert a is not b
  a.font_size = 99
  assert b.font_size != 99
