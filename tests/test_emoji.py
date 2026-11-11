# tests/test_emoji.py

"""Emoji set in a face that has them, so Word draws no box in their place."""

from docmarq import DOCX
from docmarq.emoji import EMOJI_FONT
from conftest import NS, docx_xml

def emoji_get_a_run_in_a_face_of_their_own_when_asked(tmp_path):
  def fonts(doc:DOCX) -> dict:
    runs = docx_xml(doc.path).xpath("//w:r[w:t]", namespaces=NS)
    face = "string(w:rPr/w:rFonts/@w:ascii)"
    return {r.xpath("string(w:t)", namespaces=NS): r.xpath(face, namespaces=NS) for r in runs}
  plain = DOCX(str(tmp_path / "plain.docx"))
  plain.para("low 🟢 high")
  assert list(fonts(plain.save())) == ["low 🟢 high"] # off unless asked
  doc = DOCX(str(tmp_path / "emoji.docx")).emoji()
  doc.para("low 🟢 high ✓")
  doc.table([["🔴 red", "plain"]])
  found = fonts(doc.save())
  assert found["🟢"] == found["🔴"] == EMOJI_FONT
  assert found["low "] != EMOJI_FONT and found[" high ✓"] != EMOJI_FONT # a tick stays text

def emoji_keep_the_spaces_and_the_tab_around_them(tmp_path):
  doc = DOCX(str(tmp_path / "e.docx")).emoji()
  doc.para("\tlow 🟢 high")
  body = docx_xml(doc.save().path)
  assert body.xpath("count(//w:tab)", namespaces=NS) == 1
  assert body.xpath("//w:t[@xml:space='preserve']/text()", namespaces=NS) == ["low ", "🟢", " high"]

def a_selector_or_a_joiner_alone_stays_in_the_text(tmp_path):
  doc = DOCX(str(tmp_path / "s.docx")).emoji()
  doc.para("©\ufe0f and क्\u200dष")
  body = docx_xml(doc.save().path)
  face = "//w:r[w:rPr/w:rFonts/@w:ascii='Segoe UI Emoji']/w:t/text()"
  assert body.xpath(face, namespaces=NS) == ["©\ufe0f"]
