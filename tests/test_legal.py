# tests/test_legal.py

"""`docmarq.legal`: letter header, preamble, columns, signing lines, attachments, page number."""

import inspect, zipfile
from pathlib import Path
import pytest
from PIL import Image, features
from docmarq.paper import md_to_docx as paper_md_to_docx
from docmarq.legal import md_to_docx, LegalStyle, lang_style, LANG_PRESETS
from conftest import NS, docx_xml

LETTER = """---
title: Termination notice
place: Gdynia
date: 2026-10-07
from: |
  Sender Name
  1 Sender Street
to: [Recipient Ltd, 2 Recipient Avenue]
ref: I C 123/26
attachments: [copy of the contract, delivery receipt]
---

I hereby terminate the contract with one month's notice, effective at the end of the month.
"""

SECTION = "Sectionmark text that runs long enough to fill a few lines of a narrow column. "
CONTRACT = "---\ntitle: Service contract\nlogo: logo.png\n---\n\n" \
  "Preamble naming the parties.\n\n" \
  + "\n\n".join(f"## §{i}. Clause {i}\n\n" + SECTION * 3 for i in range(1, 5))

@pytest.fixture
def render(tmp_path):
  """Render `md` with `style`, a 60×20 logo beside it; return the document root and its path."""
  Image.new("RGB", (60, 20), (30, 60, 120)).save(tmp_path / "logo.png")
  def run(md:str, style:LegalStyle):
    path = str(tmp_path / "legal.docx")
    md_to_docx(md, path, style=style, base_dir=str(tmp_path))
    return docx_xml(path), path
  return run

@pytest.fixture
def para_of():
  """The paragraph holding `text`."""
  def of(root, text:str):
    return root.xpath(f"//w:p[.//w:t[contains(., '{text}')]]", namespaces=NS)[0]
  return of

#-------------------------------------------------------------------------------- md2paper contract

def md_to_docx_matches_the_paper_signature():
  def params(f) -> list[tuple]:
    return [(name, p.default) for name, p in inspect.signature(f).parameters.items()]
  assert params(md_to_docx) == params(paper_md_to_docx)

def presets_hyphenate_in_their_language():
  assert all(preset["hyphenate"] == lang for lang, preset in LANG_PRESETS.items())
  assert LegalStyle().hyphenate is None
  assert lang_style("pl").date_format == "%d.%m.%Y r."

@pytest.mark.parametrize("field, value, error", [
  ("logo_align", "X", ValueError), ("to_align", "C", ValueError), ("columns", 0, ValueError),
])
def wrong_style_fails_before_writing(tmp_path, field, value, error):
  with pytest.raises(error, match=field):
    md_to_docx("Text.", str(tmp_path / "x.docx"), style=LegalStyle(**{field: value}))

def paper_style_is_refused(tmp_path):
  from docmarq.paper import PaperStyle
  with pytest.raises(TypeError, match="LegalStyle"):
    md_to_docx("Text.", str(tmp_path / "x.docx"), style=PaperStyle())

#------------------------------------------------------------------------------------------- Header

def sender_and_stamp_share_a_row(render):
  root, _ = render(LETTER, lang_style("pl"))
  cells = root.xpath("/w:document/w:body/w:tbl[1]/w:tr/w:tc", namespaces=NS)
  assert [c.xpath("string(.)", namespaces=NS) for c in cells] \
    == ["Sender Name1 Sender Street", "Gdynia, 07.10.2026 r."]
  assert cells[1].xpath("string(.//w:jc/@w:val)", namespaces=NS) == "right"

def stamp_goes_above_a_sender_with_no_room(render, para_of):
  md = LETTER.replace("Sender Name", "Sendername" * 8)
  root, _ = render(md, lang_style("pl"))
  assert not root.xpath("/w:document/w:body/w:tbl", namespaces=NS)
  body = root.xpath("/w:document/w:body", namespaces=NS)[0]
  assert body.index(para_of(root, "Gdynia")) < body.index(para_of(root, "Sendername"))

def recipient_sits_in_the_right_half_with_the_case_number(render, para_of):
  root, _ = render(LETTER, lang_style("pl"))
  to = para_of(root, "Recipient Ltd")
  assert int(to.xpath("string(w:pPr/w:ind/@w:left)", namespaces=NS)) \
    == round(85 * 1440 / 25.4)
  assert to.xpath("string(w:pPr/w:jc/@w:val)", namespaces=NS) == "right"
  ref = to.xpath("w:r[w:t='I C 123/26']/w:rPr/w:b", namespaces=NS)
  assert ref and "Sygn. akt: " in to.xpath("string(.)", namespaces=NS)

def recipient_flush_left(render, para_of):
  root, _ = render(LETTER, lang_style("pl", to_align="L"))
  to = para_of(root, "Recipient Ltd")
  assert to.xpath("string(w:pPr/w:ind/@w:right)", namespaces=NS)
  assert to.xpath("string(w:pPr/w:jc/@w:val)", namespaces=NS) == "left"

def title_is_centred_in_small_capitals(render, para_of):
  root, _ = render(LETTER, lang_style("pl", title_caps=True, title_size=18))
  title = para_of(root, "Termination notice")
  assert title.xpath("string(w:pPr/w:jc/@w:val)", namespaces=NS) == "center"
  rpr = title.xpath("w:r/w:rPr", namespaces=NS)[0]
  assert rpr.xpath("w:smallCaps", namespaces=NS)
  assert rpr.xpath("string(w:sz/@w:val)", namespaces=NS) == "36"

def logo_floats_beside_a_lone_title(render, para_of):
  root, _ = render(CONTRACT, lang_style("pl", logo_align="R"))
  title = para_of(root, "Service contract")
  anchor = title.xpath(".//wp:anchor", namespaces=NS)[0]
  assert anchor.xpath("string(wp:positionH/wp:align)", namespaces=NS) == "right"
  assert anchor.xpath("wp:wrapSquare", namespaces=NS)
  indents = title.xpath("w:pPr/w:ind", namespaces=NS)[0]
  assert indents.get(f"{{{NS['w']}}}left") == indents.get(f"{{{NS['w']}}}right")
  # the preamble ends below the logo, then the clauses begin
  assert root.xpath("//w:br[@w:clear='all']", namespaces=NS)

def logo_takes_a_row_of_its_own_above_header_rows(render):
  root, _ = render(LETTER.replace("---\n\nI hereby", "logo: logo.png\n---\n\nI hereby"),
    lang_style("pl"))
  first = root.xpath("/w:document/w:body/*[1]", namespaces=NS)[0]
  assert first.xpath(".//wp:inline", namespaces=NS)
  assert not root.xpath("//wp:anchor", namespaces=NS)

def preamble_takes_the_header_face(render, para_of):
  root, _ = render(CONTRACT, lang_style("pl", header_font="Cambria", preamble_size=12))
  rpr = para_of(root, "Preamble naming").xpath("w:r/w:rPr", namespaces=NS)[0]
  assert rpr.xpath("string(w:rFonts/@w:ascii)", namespaces=NS) == "Cambria"
  assert rpr.xpath("string(w:sz/@w:val)", namespaces=NS) == "24"
  clause = para_of(root, "Sectionmark").xpath("w:r/w:rPr", namespaces=NS)[0]
  assert clause.xpath("string(w:rFonts/@w:ascii)", namespaces=NS) == "Calibri"

#------------------------------------------------------------------------------------------ Columns

def clauses_run_in_columns_after_the_preamble(render):
  root, _ = render(CONTRACT, lang_style("pl", columns=2))
  sections = root.xpath("//w:sectPr", namespaces=NS)
  assert [s.xpath("string(w:cols/@w:num)", namespaces=NS) for s in sections] == ["", "2", "1"]
  # the preamble's section ends before the first clause
  first_end = root.xpath("//w:p[w:pPr/w:sectPr]", namespaces=NS)[0]
  assert "Clause 1" not in first_end.xpath("string(.)", namespaces=NS)

def one_column_adds_no_section(render):
  root, _ = render(CONTRACT, lang_style("pl"))
  assert len(root.xpath("//w:sectPr", namespaces=NS)) == 1

#---------------------------------------------------------------------------------------------- End

def signing_lines_are_dotted_and_attachments_numbered(render, para_of):
  root, _ = render(LETTER, lang_style("pl", sign=["Sender Name"]))
  rule = root.xpath("//w:p[w:pPr/w:pBdr/w:bottom]", namespaces=NS)[-1]
  assert rule.xpath("string(w:pPr/w:pBdr/w:bottom/@w:val)", namespaces=NS) == "dotted"
  listing = para_of(root, "Załączniki:")
  assert listing.xpath("string(.)", namespaces=NS) \
    == "Załączniki:1. copy of the contract2. delivery receipt"

def end_is_kept_together(render):
  root, _ = render(LETTER, lang_style("pl", sign=["Sender Name"]))
  body = root.xpath("/w:document/w:body/w:p", namespaces=NS)
  tail = body[-3:] # rule, label, attachments
  assert [bool(p.xpath("w:pPr/w:keepNext", namespaces=NS)) for p in tail] == [True, True, False]

def sign_bottom_floats_the_end_at_the_bottom_margin(render):
  root, _ = render(CONTRACT, lang_style("pl", sign="contract", sign_bottom=True))
  place = root.xpath("//w:tblpPr", namespaces=NS)[0]
  assert (place.get(f"{{{NS['w']}}}vertAnchor"), place.get(f"{{{NS['w']}}}tblpYSpec")) \
    == ("margin", "bottom")
  # the signing lines sit inside it, in a table of their own
  assert place.xpath("../../w:tr//w:tbl//w:t/text()", namespaces=NS) \
    == ["Zleceniodawca", "Zleceniobiorca"]

def page_number_hides_on_a_single_page(render):
  _, path = render(LETTER, lang_style("pl"))
  footer = docx_xml(path, "word/footer1.xml")
  instr = "".join(footer.xpath("//w:instrText/text()", namespaces=NS))
  assert instr.startswith(" IF  NUMPAGES  > 1 \"Strona ")

def legal_text_takes_the_body_colour_and_widow_control(render, para_of):
  root, path = render(LETTER, lang_style("pl", body_color=(0.2, 0.0, 0.0)))
  rpr = para_of(root, "I hereby").xpath("w:r/w:rPr", namespaces=NS)[0]
  assert rpr.xpath("string(w:color/@w:val)", namespaces=NS) == "330000"
  styles = docx_xml(path, "word/styles.xml")
  assert styles.xpath("w:style[@w:styleId='Normal']/w:pPr/w:widowControl", namespaces=NS)

#--------------------------------------------------------------------------------------------- Logo

@pytest.mark.parametrize("name, mode", [("logo.webp", "RGB"), ("logo.jpg", "CMYK")])
def logo_word_cannot_take_goes_in_as_png(tmp_path, name, mode):
  if name.endswith("webp") and not features.check("webp"): pytest.skip("Pillow without WebP")
  Image.new(mode, (60, 20)).save(tmp_path / name)
  path = str(tmp_path / "l.docx")
  md_to_docx(f"---\ntitle: T\nlogo: {name}\n---\n\nBody.\n", path, style=lang_style("pl"),
    base_dir=str(tmp_path))
  media = [n for n in zipfile.ZipFile(path).namelist() if n.startswith("word/media/")]
  assert len(media) == 1 and media[0].endswith(".png")

def unreadable_logo_warns_and_is_left_out(tmp_path):
  (tmp_path / "logo.png").write_bytes(b"not a picture")
  path = str(tmp_path / "l.docx")
  with pytest.warns(RuntimeWarning, match="could not be read") as caught:
    md_to_docx("---\ntitle: T\nlogo: logo.png\n---\n\nBody.\n", path,
      style=lang_style("pl"), base_dir=str(tmp_path))
  assert {Path(w.filename).name for w in caught} == {Path(__file__).name}
  assert not [n for n in zipfile.ZipFile(path).namelist() if n.startswith("word/media/")]
