# tests/test_parallel.py

"""`md_parallel_to_docx`: versions side by side in a table, a row per block, checked to line up."""

import pytest
from docmarq.legal import md_parallel_to_docx, lang_style, LegalStyle
from conftest import NS, docx_xml

PL = "---\ntitle: Umowa\nattachments: [załącznik]\n---\n\nStrony umowy.\n\n## §1\n\nTreść.\n\n" \
  "<!-- pagebreak -->\n\n## §2\n\nDalej.\n"
EN = "---\ntitle: Contract\nattachments: [annex]\n---\n\nThe parties.\n\n## §1\n\nText.\n\n" \
  "<!-- pagebreak -->\n\n## §2\n\nMore.\n"

@pytest.fixture
def render(tmp_path):
  """Render versions side by side, return the document root."""
  def run(texts:list[str], styles:list|None=None, **sign):
    path = str(tmp_path / "parallel.docx")
    styles = styles or [lang_style("pl", **sign), lang_style("en")]
    md_parallel_to_docx(texts, path, styles=styles)
    return docx_xml(path)
  return run

@pytest.fixture
def rows_of():
  """Text of each cell, row by row, of every top-level table."""
  def rows(root) -> list[list[str]]:
    trs = root.xpath("/w:document/w:body/w:tbl/w:tr", namespaces=NS)
    return [[tc.xpath("string(.)", namespaces=NS) for tc in tr.xpath("w:tc", namespaces=NS)]
      for tr in trs]
  return rows

#-------------------------------------------------------------------------------------------- Pairs

def blocks_pair_up_in_rows(render, rows_of):
  rows = rows_of(render([PL, EN]))
  assert rows[0] == ["Umowa", "Contract"]
  assert rows[1:4] == [["Strony umowy.", "The parties."], ["§1", "§1"], ["Treść.", "Text."]]

def page_break_starts_a_new_table_on_a_new_page(render):
  root = render([PL, EN])
  body = root.xpath("/w:document/w:body/*", namespaces=NS)
  tables = [i for i, el in enumerate(body) if el.tag.endswith("}tbl")]
  breaks = [i for i, el in enumerate(body) if el.xpath("w:pPr/w:pageBreakBefore", namespaces=NS)]
  assert len(breaks) == 1 and tables[0] < breaks[0] < tables[1]

def heading_row_keeps_with_the_next_but_not_across_a_break(render):
  pl = "## A\n\nTekst.\n\n## B\n\n<!-- pagebreak -->\n\nDalej.\n"
  en = "## A\n\nText.\n\n## B\n\n<!-- pagebreak -->\n\nMore.\n"
  root = render([pl, en])
  def kept(text:str) -> list[bool]:
    trs = root.xpath(f"//w:tr[.//w:t='{text}']", namespaces=NS)
    return [bool(tr.xpath(".//w:keepNext", namespaces=NS)) for tr in trs]
  assert (kept("A"), kept("B"), kept("Tekst.")) == ([True], [False], [False])

def each_version_is_marked_in_its_language(render):
  root = render([PL, EN])
  def langs(text:str) -> set[str]:
    return set(root.xpath(f"//w:r[w:t='{text}']/w:rPr/w:lang/@w:val", namespaces=NS))
  assert (langs("Treść."), langs("Text.")) == ({"pl-PL"}, {"en-US"})

def signing_lines_span_the_page_and_attachments_keep_their_column(render):
  root = render([PL, EN], sign=["Zleceniodawca / Principal", "Zleceniobiorca / Contractor"])
  labels = root.xpath("//w:tbl[w:tr/w:tc//w:pBdr]//w:t/text()", namespaces=NS)
  assert labels == ["Zleceniodawca / Principal", "Zleceniobiorca / Contractor"]
  last = root.xpath("/w:document/w:body/w:tbl[last()]/w:tr/w:tc", namespaces=NS)
  assert [tc.xpath("string(.)", namespaces=NS) for tc in last] \
    == ["Załączniki:1. załącznik", "Attachments:1. annex"]

#------------------------------------------------------------------------------------------- Checks

def versions_that_part_ways_name_the_block(render):
  with pytest.raises(ValueError, match="block 2: .*heading_open h2.*heading_open h3"):
    render([PL, EN.replace("## §1", "### §1")])

def missing_block_is_named(render):
  with pytest.raises(ValueError, match="nothing"):
    render([PL, EN.replace("\n\nMore.\n", "\n")])

def page_break_inside_a_block_is_refused(render):
  quoted = "> quote\n>\n> <!-- pagebreak -->\n>\n> more\n"
  with pytest.raises(ValueError, match="pagebreak"):
    render([quoted, quoted])

def one_version_or_unmatched_styles_are_refused(tmp_path):
  with pytest.raises(ValueError, match="two versions"):
    md_parallel_to_docx([PL], str(tmp_path / "x.docx"), styles=[lang_style("pl")])
  with pytest.raises(ValueError, match="as many styles"):
    md_parallel_to_docx([PL, EN], str(tmp_path / "x.docx"), styles=[lang_style("pl")])

#-------------------------------------------------------------------------------------------- Apart

def versions_keep_their_bookmarks_apart(render):
  text = "## Scope\n\nSee [scope](#scope).\n"
  root = render([text, text])
  assert root.xpath("//w:bookmarkStart/@w:name", namespaces=NS) == ["scope", "v2-scope"]
  assert root.xpath("//w:hyperlink/@w:anchor", namespaces=NS) == ["scope", "v2-scope"]

def no_header_row_without_a_header(render, rows_of):
  assert rows_of(render(["Tekst.\n", "Text.\n"])) == [["Tekst.", "Text."]]

def version_naming_no_language_keeps_its_words_whole(render):
  root = render([PL, EN], styles=[lang_style("pl"), LegalStyle()])
  def cell(text:str):
    return root.xpath(f"//w:tc[.//w:t='{text}']", namespaces=NS)[0]
  first, second = cell("Treść."), cell("Text.")
  assert first.xpath(".//w:lang/@w:val", namespaces=NS) \
    and not first.xpath(".//w:suppressAutoHyphens", namespaces=NS)
  assert second.xpath(".//w:suppressAutoHyphens", namespaces=NS) \
    and not second.xpath(".//w:lang", namespaces=NS)
