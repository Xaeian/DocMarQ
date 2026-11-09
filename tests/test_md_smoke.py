# tests/test_md_smoke.py

"""Markdown rendering smoke: small samples through `md_to_docx` produce valid DOCX (no crash)."""

import pytest
from conftest import NS, assert_valid_docx, docx_text, docx_xml
from docmarq.paper import md_to_docx, lang_style
from docmarq.constants import A4

#----------------------------------------------------------------------------------------- Markdown

def md_minimal_renders(tmp_path):
  path = tmp_path / "min.docx"
  md_to_docx("# Title\n\nHello world.", str(path))
  assert_valid_docx(path)

def md_all_headings_render(tmp_path):
  path = tmp_path / "head.docx"
  md_to_docx(
    "# H1\n## H2\n### H3\n#### H4\n##### H5\n###### H6\n\nbody",
    str(path),
  )
  assert_valid_docx(path)

def md_emphasis_and_code(tmp_path):
  path = tmp_path / "em.docx"
  src = "Plain **bold**, *italic*, `code`, ~~strike~~, [link](https://x.com)."
  md_to_docx(src, str(path))
  assert_valid_docx(path)

def md_lists_bullet_and_ordered(tmp_path):
  path = tmp_path / "lists.docx"
  md_to_docx("- a\n- b\n\n1. one\n2. two", str(path))
  assert_valid_docx(path)

def md_table_renders(tmp_path):
  path = tmp_path / "tab.docx"
  src = (
    "| A | B | C |\n"
    "|---|--:|:-:|\n"
    "| a | 1 | x |\n"
    "| b | 2 | y |\n"
  )
  md_to_docx(src, str(path))
  assert_valid_docx(path)

def md_blockquote_and_callout(tmp_path):
  path = tmp_path / "bq.docx"
  src = "> normal quote\n\n> [!NOTE]\n> note callout"
  md_to_docx(src, str(path))
  assert_valid_docx(path)

def md_frontmatter_metadata(tmp_path):
  path = tmp_path / "fm.docx"
  src = "---\ntitle: Doc\nauthor: Xaeian\n---\n\n# Doc\n\nBody."
  md_to_docx(src, str(path))
  assert_valid_docx(path)

def md_page_landscape_flips_page(tmp_path):
  # orientation comes from `page=`, never from the document
  path = tmp_path / "ls.docx"
  doc = md_to_docx("# Wide content", str(path), page=A4.landscape())
  assert doc.page_width > doc.page_height
  assert_valid_docx(path)

def md_frontmatter_never_sets_geometry(tmp_path):
  # a leftover `render:` block is inert content, not configuration
  src = "---\nrender:\n  landscape: true\n---\n\n# Content"
  doc = md_to_docx(src, str(tmp_path / "inert.docx"))
  assert doc.page_width < doc.page_height

@pytest.mark.parametrize("label, title", [
  (None, None), ("Bibliografia", "Bibliografia"), (True, "Literatura"),
], ids=["rule", "own word", "language word"])
def md_footnote_label_titles_the_bibliography(tmp_path, label, title):
  # a rule by default, or a title: the style's own word, or its language's with `True`,
  # which starts no page of its own: a page break is the author's, put where it is wanted
  path = str(tmp_path / "fn.docx")
  style = lang_style("pl", footnote_label=label)
  md_to_docx("Body[^1].\n\n[^1]: footnote text.", path, style=style)
  titles = {"Literatura", "Bibliografia", "References"} & set(docx_text(path).split())
  assert titles == ({title} if title else set())
  assert not docx_xml(path, "word/document.xml").xpath(
    "//w:br[@w:type='page'] | //w:pageBreakBefore", namespaces=NS)
  assert not docx_xml(path, "word/styles.xml").xpath(
    "//w:style[@w:styleId='Heading2']//w:pageBreakBefore", namespaces=NS)

@pytest.mark.parametrize("src, title", [
  ("Body[^1].\n\n[^1]: footnote text.", "Literatura"),
  ("Body[^1].\n\n### Bibliografia\n\n[^1]: footnote text.", "Bibliografia"),
], ids=["language word", "own heading"])
def md_paper_titles_the_bibliography_unasked(tmp_path, src, title):
  # a paper titles its footnotes in its language, unless it ends with a heading of its own
  path = str(tmp_path / "fn.docx")
  md_to_docx(src, path, style=lang_style("pl"))
  words = docx_text(path).split()
  assert [w for w in words if w in {"Literatura", "Bibliografia"}] == [title]
  assert not docx_xml(path, "word/document.xml").xpath("//w:pBdr", namespaces=NS) # no rule

def md_paper_bibliography_title_sits_where_an_authors_would(tmp_path):
  # the title a paper adds is written as an author's `## Literatura` is, to the last attribute
  from lxml import etree
  bodies = []
  for name, heading in (("auto", ""), ("authored", "## Literatura\n\n")):
    path = str(tmp_path / f"{name}.docx")
    md_to_docx(f"Body[^1].\n\n{heading}[^1]: footnote text.", path, style=lang_style("pl"))
    bodies.append(etree.tostring(docx_xml(path)))
  assert bodies[0] == bodies[1]

def md_only_a_paper_titles_the_bibliography_unasked():
  # the bare engine and an act keep the thin rule over their footnotes
  from docmarq.md import MarkdownStyle
  from docmarq.legal import LegalStyle
  from docmarq.paper import PaperStyle
  assert MarkdownStyle().footnote_label is None
  assert LegalStyle().footnote_label is None
  assert PaperStyle().footnote_label is True

def md_bibliography_title_is_a_heading_like_any_other(tmp_path):
  # in the headings' font, with an anchor a link can reach
  path = str(tmp_path / "fn.docx")
  md_to_docx("## Section\n\nBody[^1].\n\n[^1]: footnote text.", path,
    style=lang_style("pl", font_head="Georgia"))
  title = docx_xml(path, "word/document.xml").xpath(
    "//w:p[w:pPr/w:pStyle/@w:val='Heading2'][.//w:t='Literatura']", namespaces=NS)[0]
  assert title.xpath(".//w:rFonts/@w:ascii", namespaces=NS) == ["Georgia"]
  assert title.xpath(".//w:bookmarkStart/@w:name", namespaces=NS)
