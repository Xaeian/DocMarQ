# tests/test_fonts.py

"""
Faces carried in the file: the families a document names, found in `font_dir`,
as Word embeds them, and the widths of words measured by them.
"""

import shutil, zipfile
from pathlib import Path
import pytest
import reportlab
from PIL import ImageFont
from docmarq import DOCX, text_width_mm
from conftest import NS, docx_xml

VERA = Path(reportlab.__file__).parent / "fonts" # the Bitstream Vera reportlab carries
W = f"{{{NS['w']}}}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

@pytest.fixture
def faces(tmp_path) -> Path:
  """A font folder in pdfmarq's layout holding Vera Sans in Regular and Bold, no italics."""
  folder = tmp_path / "faces" / "BitstreamVeraSans"
  folder.mkdir(parents=True)
  shutil.copyfile(VERA / "Vera.ttf", folder / "BitstreamVeraSans-Regular.ttf")
  shutil.copyfile(VERA / "VeraBd.ttf", folder / "BitstreamVeraSans-Bold.ttf")
  return folder.parent

def a_named_family_found_in_font_dir_goes_into_the_file(tmp_path, faces):
  path = tmp_path / "embedded.docx"
  doc = DOCX(str(path), font_body="Bitstream Vera Sans", font_dir=str(faces))
  doc.para("Carried").text(" and bold", bold=True)
  doc.save()
  table = docx_xml(str(path), "word/fontTable.xml")
  font = table.xpath("//w:font[@w:name='Bitstream Vera Sans']", namespaces=NS)[0]
  assert [e.tag[len(W):] for e in font if e.tag.startswith(W + "embed")] == [
    "embedRegular", "embedBold"] # no italic in the folder, so Word slants the Regular
  with zipfile.ZipFile(path) as z:
    rels = {r.get("Id"): r.get("Target") for r in docx_xml(str(path),
      "word/_rels/fontTable.xml.rels").iter() if r.get("Id")}
    regular = font.find(W + "embedRegular")
    stored = z.read("word/" + rels[regular.get(R + "id")])
    key = bytes.fromhex(regular.get(W + "fontKey").strip("{}").replace("-", ""))[::-1]
    shown = bytes(b ^ key[k % 16] for k, b in enumerate(stored[:32])) + stored[32:]
    assert shown == (VERA / "Vera.ttf").read_bytes() # hidden as ECMA-376 hides it, nothing lost
    assert "obfuscatedFont" in z.read("[Content_Types].xml").decode()
  assert docx_xml(str(path), "word/settings.xml").find(W + "embedTrueTypeFonts") is not None

def nothing_is_embedded_without_font_dir_or_a_face_found_there(tmp_path, faces):
  plain, other = tmp_path / "plain.docx", tmp_path / "other.docx"
  DOCX(str(plain), font_body="Bitstream Vera Sans").para("Plain").save()
  DOCX(str(other), font_body="Times New Roman", font_dir=str(faces)).para("Other").save()
  for path in (plain, other):
    table = docx_xml(str(path), "word/fontTable.xml")
    assert not table.xpath("//w:embedRegular", namespaces=NS)
    assert docx_xml(str(path), "word/settings.xml").find(W + "embedTrueTypeFonts") is None

def words_measure_by_the_faces_the_document_carries(tmp_path, faces):
  words = "Carried words"
  carried = DOCX(str(tmp_path / "m.docx"), font_dir=str(faces)).text_width(words,
    "Bitstream Vera Sans", 10, bold=True)
  vera = ImageFont.truetype(str(VERA / "VeraBd.ttf"), 1000).getlength(words) / 100 * 25.4 / 72
  assert carried == pytest.approx(vera)
  # Elsewhere the family is found nowhere, so it measures as Helvetica
  assert text_width_mm(words, "Bitstream Vera Sans", 10, bold=True) != pytest.approx(vera)
