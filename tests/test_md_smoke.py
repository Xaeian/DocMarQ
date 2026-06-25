# tests/test_md_smoke.py

"""Markdown rendering smoke: small samples through `md_to_docx` produce valid DOCX (no crash)."""

import pytest
from conftest import assert_valid_docx
from docmarq.md import md_to_docx, MarkdownStyle

#---------------------------------------------------------------------------------- Markdown

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

def md_landscape_from_frontmatter(tmp_path):
  # `render.landscape: true` auto-flips when `landscape=` is None (default).
  # Top-level `landscape:` is no longer honored (warns).
  path = tmp_path / "ls.docx"
  src = "---\nrender:\n  landscape: true\n---\n\n# Wide content"
  doc = md_to_docx(src, str(path))
  assert doc.page_width > doc.page_height
  assert_valid_docx(path)

def md_landscape_explicit_overrides_frontmatter(tmp_path):
  # explicit `landscape=False` wins over `render.landscape: true`
  path = tmp_path / "ls_false.docx"
  src = "---\nrender:\n  landscape: true\n---\n\n# Content"
  doc = md_to_docx(src, str(path), landscape=False)
  assert doc.page_width < doc.page_height

def md_footnote_no_label_default(tmp_path):
  # default `footnote_label=None` emits HR + smaller-font footnotes
  path = tmp_path / "fn_hr.docx"
  src = "Body[^1].\n\n[^1]: footnote text."
  md_to_docx(src, str(path))
  assert_valid_docx(path)

def md_footnote_with_label(tmp_path):
  # setting `footnote_label="References"` emits an H2 heading above the section
  path = tmp_path / "fn_label.docx"
  src = "Body[^1].\n\n[^1]: footnote text."
  md_to_docx(src, str(path), style=MarkdownStyle(footnote_label="References"))
  assert_valid_docx(path)
