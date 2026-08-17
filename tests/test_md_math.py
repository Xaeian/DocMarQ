# tests/test_md_math.py

"""Math (OMML): LaTeX->OMML converter structures + `md_to_docx` integration.

Converter tests assert the right OMML structures appear and that unsupported
input raises `MathConversionError` (so the renderer can fall back to an image).
Integration tests confirm `$...$` / `$$...$$` produce native `<m:oMath>` in
`word/document.xml`, never leak raw LaTeX, and keep the file valid OOXML.
"""

import re
import zipfile
import pytest
from docx.oxml.ns import qn
from conftest import assert_valid_docx
from docmarq.md import md_to_docx, MarkdownStyle
from docmarq.md.math import latex_to_omath, build_omath_para, MathConversionError

#------------------------------------------------------------------------------------------ Helpers

@pytest.fixture
def document_xml():
  """Return `word/document.xml` text from a .docx (a ZIP container)."""
  def _read(path) -> str:
    with zipfile.ZipFile(path) as z:
      return z.read("word/document.xml").decode("utf-8")
  return _read

@pytest.fixture
def child_tags():
  """Local tag names of an element's direct children (namespace stripped)."""
  def _tags(el) -> list[str]:
    return [c.tag.split("}")[-1] for c in el]
  return _tags

#---------------------------------------------------------------------------- Converter: structures

def omml_superscript():
  el = latex_to_omath(r"E = mc^2")
  assert el.tag == qn("m:oMath")
  assert el.find(qn("m:sSup")) is not None

def omml_subscript_and_subsup():
  assert latex_to_omath(r"x_i").find(qn("m:sSub")) is not None
  assert latex_to_omath(r"x_i^2").find(qn("m:sSubSup")) is not None

def omml_fraction_child_order(child_tags):
  f = latex_to_omath(r"\frac{a}{b}").find(qn("m:f"))
  assert f is not None
  # CT_F order: (fPr?) num, den
  assert child_tags(f) == ["num", "den"]

def omml_sqrt_and_nthroot(child_tags):
  rad = latex_to_omath(r"\sqrt{2}").find(qn("m:rad"))
  assert rad is not None
  # Plain sqrt hides the degree: radPr/degHide present, empty deg, then e.
  assert child_tags(rad) == ["radPr", "deg", "e"]
  rad3 = latex_to_omath(r"\sqrt[3]{x}").find(qn("m:rad"))
  assert child_tags(rad3) == ["deg", "e"]  # explicit degree, no degHide

def omml_nary_sum_limits(child_tags):
  nary = latex_to_omath(r"\sum_{i=1}^{n} i", display=True).find(qn("m:nary"))
  assert nary is not None
  # CT_Nary order: naryPr, sub, sup, e
  assert child_tags(nary) == ["naryPr", "sub", "sup", "e"]
  pr = nary.find(qn("m:naryPr"))
  # CT_NaryPr canonical order: chr, limLoc, grow, subHide, supHide
  assert child_tags(pr) == ["chr", "limLoc", "grow"]  # both limits present -> no hide
  assert pr.find(qn("m:limLoc")).get(qn("m:val")) == "undOvr"  # display style

def omml_nary_hides_absent_limits():
  pr = latex_to_omath(r"\sum_{i} x").find(qn("m:nary")).find(qn("m:naryPr"))
  assert pr.find(qn("m:supHide")) is not None
  assert pr.find(qn("m:subHide")) is None

def omml_integral_inline_uses_side_limits():
  pr = latex_to_omath(r"\int_0^1 x", display=True).find(qn("m:nary")).find(qn("m:naryPr"))
  # Integrals keep side limits even in display mode.
  assert pr.find(qn("m:limLoc")).get(qn("m:val")) == "subSup"

def omml_delimiter_default_parens_omit_props():
  d = latex_to_omath(r"\left( x \right)").find(qn("m:d"))
  assert d is not None
  assert d.find(qn("m:dPr")) is None  # parens are the default, no dPr needed

def omml_delimiter_custom_brackets():
  d = latex_to_omath(r"\left[ x \right]").find(qn("m:d"))
  dpr = d.find(qn("m:dPr"))
  assert dpr.find(qn("m:begChr")).get(qn("m:val")) == "["
  assert dpr.find(qn("m:endChr")).get(qn("m:val")) == "]"

def omml_matrix_dimensions():
  mat = latex_to_omath(r"\begin{matrix} a & b \\ c & d \end{matrix}").find(qn("m:m"))
  assert mat is not None
  rows = mat.findall(qn("m:mr"))
  assert len(rows) == 2
  assert len(rows[0].findall(qn("m:e"))) == 2
  count = mat.find(qn("m:mPr")).find(qn("m:mcs")).find(qn("m:mc")).find(
    qn("m:mcPr")).find(qn("m:count")).get(qn("m:val"))
  assert count == "2"

def omml_matrix_braces_not_split_as_separators():
  # `&` / `\\` inside a braced cell must NOT split the matrix.
  mat = latex_to_omath(
    r"\begin{matrix} {a & b} \\ c \end{matrix}").find(qn("m:m"))
  rows = mat.findall(qn("m:mr"))
  assert len(rows) == 2  # two rows, first row has a single (braced) cell

def omml_pmatrix_wrapped_in_parens():
  d = latex_to_omath(r"\begin{pmatrix} a \\ b \end{pmatrix}").find(qn("m:d"))
  assert d is not None and d.find(qn("m:e")).find(qn("m:m")) is not None

def omml_cases_left_aligned_brace():
  d = latex_to_omath(
    r"\begin{cases} 1 & x>0 \\ 0 & x<0 \end{cases}").find(qn("m:d"))
  dpr = d.find(qn("m:dPr"))
  assert dpr.find(qn("m:begChr")).get(qn("m:val")) == "{"
  assert dpr.find(qn("m:endChr")).get(qn("m:val")) == ""

def omml_accent_and_bar():
  assert latex_to_omath(r"\hat{x}").find(qn("m:acc")) is not None
  bar = latex_to_omath(r"\overline{AB}").find(qn("m:bar"))
  assert bar.find(qn("m:barPr")).find(qn("m:pos")).get(qn("m:val")) == "top"

def omml_lim_uses_limlow(child_tags):
  ll = latex_to_omath(r"\lim_{x \to 0} x").find(qn("m:limLow"))
  assert ll is not None
  assert child_tags(ll) == ["e", "lim"]

def omml_font_scripts_mapping():
  # mathbb -> double-struck script on the run.
  el = latex_to_omath(r"\mathbb{R}")
  r = el.iter(qn("m:r")).__next__()
  scr = r.find(qn("m:rPr")).find(qn("m:scr"))
  assert scr.get(qn("m:val")) == "double-struck"

def omml_run_carries_color_and_size():
  el = latex_to_omath(r"x", size_halfpt=22, color_hex="1F2328")
  r = el.find(qn("m:r"))
  wrpr = r.find(qn("w:rPr"))
  assert wrpr.find(qn("w:color")).get(qn("w:val")) == "1F2328"
  assert wrpr.find(qn("w:sz")).get(qn("w:val")) == "22"

def omath_para_is_centered():
  para = build_omath_para(r"x^2", align="center")
  assert para.tag == qn("m:oMathPara")
  jc = para.find(qn("m:oMathParaPr")).find(qn("m:jc"))
  assert jc.get(qn("m:val")) == "center"
  assert para.find(qn("m:oMath")) is not None

#------------------------------------------------------------------------ Converter: error contract

@pytest.mark.parametrize("bad", [
  "", "   ",
  r"\unknowncommand{x}",
  r"\frac{a}",                 # missing 2nd arg -> unexpected end
  r"x }",                       # unbalanced close
  r"{x",                        # unbalanced open
  r"\left( x",                 # \left without \right
  r"\begin{xyz} a \end{xyz}",  # unsupported environment
  r"\overset{a}{b}",           # genuinely unsupported construct
])
def unsupported_input_raises_conversion_error(bad):
  with pytest.raises(MathConversionError):
    latex_to_omath(bad)

def deep_nesting_terminates_without_other_error():
  # Pathological nesting must not hang or hit the recursion ceiling silently;
  # it either converts or raises MathConversionError - never anything else.
  deep = r"\frac{1}{" * 50 + "x" + "}" * 50
  try:
    latex_to_omath(deep)
  except MathConversionError:
    pass

#-------------------------------------------------------------------------- Integration: md_to_docx

def md_inline_math_produces_omml(tmp_path, document_xml):
  path = tmp_path / "inline.docx"
  md_to_docx("Energy is $E = mc^2$ exactly.", str(path))
  assert_valid_docx(path)
  xml = document_xml(path)
  assert "<m:oMath>" in xml
  assert "$" not in xml  # no raw LaTeX leakage

def md_block_math_produces_omath_para(tmp_path, document_xml):
  path = tmp_path / "block.docx"
  md_to_docx(r"$$\int_0^1 x^2\,dx = \frac{1}{3}$$", str(path))
  assert_valid_docx(path)
  xml = document_xml(path)
  assert "<m:oMathPara>" in xml
  assert "<m:nary>" in xml and "<m:f>" in xml

def md_math_fenced_block(tmp_path, document_xml):
  path = tmp_path / "fence.docx"
  md_to_docx("```math\n\\frac{a}{b}\n```", str(path))
  xml = document_xml(path)
  assert "<m:oMathPara>" in xml

def md_math_in_heading_and_list(tmp_path, document_xml):
  path = tmp_path / "mixed.docx"
  src = "# Result $x^2$\n\n- item $a_i$\n- plain item\n\n> quote $\\pi$"
  md_to_docx(src, str(path))
  xml = document_xml(path)
  assert xml.count("<m:oMath>") >= 3

def md_math_disabled_renders_literal(tmp_path, document_xml):
  path = tmp_path / "disabled.docx"
  md_to_docx("Inline $x^2$ here.", str(path),
    style=MarkdownStyle(math_enable=False))
  xml = document_xml(path)
  assert "<m:oMath>" not in xml
  assert "x^2" in xml  # left as plain text

def md_unsupported_math_falls_back_to_image(tmp_path):
  # `\overset` is outside the OMML subset -> image fallback, never raw $.
  path = tmp_path / "fallback.docx"
  md_to_docx(r"Inline $\overset{a}{b}$ formula.", str(path))
  assert_valid_docx(path)
  with zipfile.ZipFile(path) as z:
    media = [n for n in z.namelist() if n.startswith("word/media/")]
    xml = z.read("word/document.xml").decode("utf-8")
  assert media, "expected an embedded image fallback"
  assert "<m:oMath>" not in xml
  assert "$" not in xml

def md_table_cell_math_preserved_as_text(tmp_path, document_xml):
  path = tmp_path / "tablemath.docx"
  src = "| sym | val |\n|---|---|\n| $x^2$ | 4 |\n"
  md_to_docx(src, str(path))
  xml = document_xml(path)
  # Cell math degrades to LaTeX source rather than being dropped.
  assert "x^2" in xml

#----------------------------------------------------------------- Regressions (adversarial review)

def nested_font_scr_before_sty():
  # CT_RPr requires m:scr before m:sty (mathbb wrapping mathbf).
  el = latex_to_omath(r"\mathbb{\mathbf{R}}")
  rpr = el.iter(qn("m:r")).__next__().find(qn("m:rPr"))
  tags = [c.tag.split("}")[-1] for c in rpr if c.tag.split("}")[-1] in ("scr", "sty")]
  assert tags == ["scr", "sty"]

def deep_nesting_raises_conversion_error_not_recursion():
  with pytest.raises(MathConversionError):
    latex_to_omath("{" * 500 + "x" + "}" * 500)
  with pytest.raises(MathConversionError):
    latex_to_omath(r"\frac{1}{" * 300 + "x" + "}" * 300)

def nary_double_script_raises():
  with pytest.raises(MathConversionError):
    latex_to_omath(r"\sum_i^n^m a")
  with pytest.raises(MathConversionError):
    latex_to_omath(r"\int_a^b_c f")

def lim_with_sub_and_sup_has_no_empty_base():
  el = latex_to_omath(r"\lim_x^2 f", display=True)
  for r in el.iter(qn("m:r")):
    t = r.find(qn("m:t"))
    assert t is not None and (t.text or "") != "", "empty-base run leaked"

def array_column_spec_consumed():
  el = latex_to_omath(r"\begin{array}{cc} a & b \\ c & d \end{array}")
  assert [t.text for t in el.iter(qn("m:t"))] == ["a", "b", "c", "d"]
  el2 = latex_to_omath(r"\begin{array}{c|c} a & b \\ c & d \end{array}")
  assert [t.text for t in el2.iter(qn("m:t"))] == ["a", "b", "c", "d"]

def text_mode_preserves_spaces():
  el = latex_to_omath(r"\text{hello world}")
  assert "".join(t.text for t in el.iter(qn("m:t"))) == "hello world"

def textbf_is_text_mode_not_math():
  el = latex_to_omath(r"\textbf{x^2}")
  assert el.findall(".//" + qn("m:sSup")) == []  # ^2 must NOT become a superscript
  assert "".join(t.text for t in el.iter(qn("m:t"))) == "x^2"
  sty = el.iter(qn("m:r")).__next__().find(qn("m:rPr")).find(qn("m:sty"))
  assert sty.get(qn("m:val")) == "b"

def align_environment_native():
  el = latex_to_omath(r"\begin{align} a &= b \\ c &= d \end{align}", display=True)
  assert el.find(qn("m:m")) is not None  # rendered as an aligned grid, not a fallback

def nested_matrix_native():
  el = latex_to_omath(
    r"\begin{pmatrix} \begin{matrix} a \\ b \end{matrix} & c \\ d & e \end{pmatrix}")
  assert el.find(qn("m:d")) is not None

def left_right_rejects_arbitrary_char():
  with pytest.raises(MathConversionError):
    latex_to_omath(r"\left a x \right b")

def nul_byte_sanitized_not_valueerror():
  # Must not raise the low-level lxml ValueError; NUL is stripped.
  el = latex_to_omath("x\x00y")
  assert "\x00" not in "".join(t.text or "" for t in el.iter(qn("m:t")))

def runs_are_coalesced():
  assert len(latex_to_omath("abcdef").findall(qn("m:r"))) == 1

def oversized_formula_raises():
  with pytest.raises(MathConversionError):
    latex_to_omath("x" * 20000)

def trailing_space_is_not_trailing_tokens():
  assert latex_to_omath("x ").find(qn("m:r")) is not None  # no "trailing tokens" error

def no_dollar_leak_when_matplotlib_present(tmp_path, document_xml):
  # \cancel + \begin{align} + \big are unsupported by OMML; with matplotlib
  # present they must become images / native, never raw `$...$`.
  path = tmp_path / "noleak.docx"
  src = (r"Inline $\cancel{x}$ here." + "\n\n"
    r"$$\begin{align} a &= b \\ c &= d \end{align}$$" + "\n\n"
    r"And $\big(\frac{a}{b}\big)$.")
  md_to_docx(src, str(path))
  xml = document_xml(path)
  leaks = [t for t in re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml) if "$" in t]
  assert leaks == [], f"raw $ leaked: {leaks}"
