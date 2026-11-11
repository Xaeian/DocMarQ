# tests/test_notes.py

"""Footnotes of the core API: numbered, marked or unmarked, in a template or not."""

import pytest
from docmarq import DOCX
from conftest import NS, docx_xml

def numbered_note_lives_in_its_part_and_text_goes_on(tmp_path):
  doc = DOCX(str(tmp_path / "note.docx"))
  doc.para("Body")
  with doc.footnote():
    doc.para("A note.")
  doc.text(" goes on.")
  doc.save()
  body, notes = docx_xml(doc.path), docx_xml(doc.path, "word/footnotes.xml")
  w = f"{{{NS['w']}}}"
  ref = body.xpath("//w:footnoteReference", namespaces=NS)[0]
  assert ref.get(f"{w}customMarkFollows") is None
  note = notes.xpath(f"//w:footnote[@w:id='{ref.get(f'{w}id')}']", namespaces=NS)[0]
  assert note.xpath("count(.//w:footnoteRef)", namespaces=NS) == 1
  assert "A note." in "".join(note.itertext())
  assert body.xpath("string(//w:p[.//w:footnoteReference])", namespaces=NS) == "Body goes on."
  assert docx_xml(doc.path, "word/settings.xml").xpath("count(//w:footnotePr)", namespaces=NS) == 1

def a_quote_goes_on_in_its_colour_after_a_note(tmp_path):
  doc = DOCX(str(tmp_path / "quote.docx"))
  doc.blockquote("Quoted", text_color="#666666")
  with doc.footnote():
    doc.para("A note.")
  doc.text(" goes on.")
  doc.save()
  colors = docx_xml(doc.path).xpath("//w:r[w:t=' goes on.']/w:rPr/w:color/@w:val", namespaces=NS)
  assert colors == ["666666"]

def a_list_or_a_heading_in_a_note_refuses_as_a_table_does(tmp_path):
  doc = DOCX(str(tmp_path / "l.docx"))
  doc.para("Body")
  with pytest.raises(ValueError, match="plain paragraphs"):
    with doc.footnote():
      doc.bullet("item")
  with pytest.raises(ValueError, match="plain paragraphs"):
    with doc.footnote():
      doc.heading("Title", level=2)
  doc.save()
  body, notes = docx_xml(doc.path), docx_xml(doc.path, "word/footnotes.xml")
  assert not body.xpath("//w:footnoteReference", namespaces=NS) # a refused note leaves no mark
  assert notes.xpath("count(//w:footnote[not(@w:type)])", namespaces=NS) == 0

def a_note_chains_its_paragraphs(tmp_path):
  doc = DOCX(str(tmp_path / "c.docx"))
  doc.para("Body")
  with doc.footnote():
    doc.chain(lambda: (doc.para("One."), doc.para("Two.")))
  notes = docx_xml(doc.save().path, "word/footnotes.xml")
  kept = notes.xpath("//w:footnote[not(@w:type)]/w:p[w:pPr/w:keepNext]", namespaces=NS)
  assert ["".join(p.itertext()).strip() for p in kept] == ["One."]

def unmarked_note_carries_no_mark_and_rule_can_go(tmp_path):
  doc = DOCX(str(tmp_path / "note.docx"))
  doc.para("Authors")
  with doc.footnote(mark=""):
    doc.para("With a university.")
  doc.footnote_rule(False)
  doc.save()
  body, notes = docx_xml(doc.path), docx_xml(doc.path, "word/footnotes.xml")
  assert body.xpath("count(//w:footnoteReference[@w:customMarkFollows='1'])", namespaces=NS) == 1
  assert body.xpath("count(//w:r[w:footnoteReference]/w:sym)", namespaces=NS) == 1
  assert notes.xpath("count(//w:footnoteRef)", namespaces=NS) == 0
  assert notes.xpath("count(//w:separator)", namespaces=NS) == 0

def notes_join_a_template_that_has_some(tmp_path):
  first = DOCX(str(tmp_path / "a.docx"))
  first.para("Body")
  with first.footnote():
    first.para("One.")
  first.save()
  doc = DOCX(str(tmp_path / "b.docx"), template=first.path)
  doc.para("Again")
  with doc.footnote():
    doc.para("Two.")
  notes = docx_xml(doc.save().path, "word/footnotes.xml")
  assert notes.xpath("count(//w:footnote[not(@w:type)])", namespaces=NS) == 2

def an_empty_note_gets_a_paragraph_and_a_note_in_a_note_refuses(tmp_path):
  doc = DOCX(str(tmp_path / "n.docx"))
  doc.para("Body")
  with doc.footnote():
    pass
  with pytest.raises(ValueError, match="no footnote"):
    with doc.footnote():
      with doc.footnote():
        pass
  notes = docx_xml(doc.save().path, "word/footnotes.xml")
  assert notes.xpath("count(//w:footnote[not(@w:type)][not(w:p)])", namespaces=NS) == 0

def the_rule_alone_names_the_separators(tmp_path):
  doc = DOCX(str(tmp_path / "r.docx"))
  doc.para("Body")
  doc.footnote_rule(False)
  settings = docx_xml(doc.save().path, "word/settings.xml")
  assert settings.xpath("count(//w:footnotePr)", namespaces=NS) == 1

def notes_join_a_template_without_their_styles(tmp_path):
  # Word saves a footnotes part, but adds its two styles only when a note first uses them
  import docx
  first = DOCX(str(tmp_path / "a.docx"))
  first.para("Body")
  with first.footnote():
    first.para("One.")
  word = docx.Document(first.save().path)
  for name in ("footnote text", "footnote reference"):
    style = word.styles[name].element
    style.getparent().remove(style)
  word.save(str(tmp_path / "word.docx"))
  doc = DOCX(str(tmp_path / "b.docx"), template=str(tmp_path / "word.docx"))
  doc.para("Again")
  with doc.footnote():
    doc.para("Two.")
  notes = docx_xml(doc.save().path, "word/footnotes.xml")
  assert notes.xpath("count(//w:footnote[not(@w:type)])", namespaces=NS) == 2

def a_note_takes_a_picture(tmp_path):
  from PIL import Image
  Image.new("RGB", (40, 30), (200, 0, 0)).save(tmp_path / "p.png")
  doc = DOCX(str(tmp_path / "i.docx"))
  doc.para("Body")
  with doc.footnote():
    doc.image(str(tmp_path / "p.png"), width=10)
  rels = docx_xml(doc.save().path, "word/_rels/footnotes.xml.rels")
  assert rels.xpath("count(//*[contains(@Type, '/image')])") == 1
