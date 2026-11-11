# docmarq/dropcap.py

"""
Drop caps: the first letter of a paragraph dropped beside its first lines, as Word's own drop cap.

Word keeps the letter in a frame paragraph of its own before the text, `w:framePr w:dropCap`,
for which python-docx has no API, so the frame is built here.
"""
import copy, re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm
from docx.text.run import Run
from .utils import text_element

_BEFORE_TEXT = ("w:pPr", "w:bookmarkStart", "w:bookmarkEnd") # before a paragraph's words
_BLANK = ("w:rPr", "w:t", "w:tab", "w:br") # what a run holds that is no picture, symbol or field
# `w:rPr` children that follow `w:sz`, in schema order
_AFTER_SIZE = ("w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
  "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath")

def drop_cap(p, size:float, line:float, lines:int=2, sink:float=0, caps:bool=False):
  """
  The first letter of paragraph `p` dropped `lines` lines, `size` pt tall and sunk `sink` pt,
  in a frame of `lines` lines `line` pt apart, in the face of its word, as bold or a colour.
  `caps` sets the rest of its first word in capitals, as IEEE journals do.
  A paragraph opening with no letter stays as it is, as one opening with a link, a formula
  or a picture.
  """
  run = _opening_run(p)
  words = run.text.lstrip() if run is not None else ""
  if not words[:1].isalpha(): return
  cut = re.match(r"\w*", words[1:]).end() if caps else 0
  run.text = words[1:cut + 1].upper() + words[cut + 1:]
  frame = OxmlElement("w:p")
  ppr = frame.get_or_add_pPr()
  ppr.append(OxmlElement("w:keepNext"))
  ppr.append(OxmlElement("w:framePr", {qn("w:dropCap"): "drop", qn("w:lines"): str(lines),
    qn("w:wrap"): "around", qn("w:vAnchor"): "text", qn("w:hAnchor"): "text"}))
  ppr.append(OxmlElement("w:spacing", {qn("w:after"): "0",
    qn("w:line"): str(round(lines * line * 20)), qn("w:lineRule"): "exact"}))
  ppr.append(OxmlElement("w:ind", {qn("w:firstLine"): "0"}))
  ppr.append(OxmlElement("w:textAlignment", {qn("w:val"): "baseline"}))
  letter = OxmlElement("w:r")
  letter.append(_letter_face(run, size, sink))
  letter.append(text_element(words[0].upper()))
  frame.append(letter)
  p._p.addprevious(frame)
  p.paragraph_format.first_line_indent = Mm(0)

def _opening_run(p) -> Run|None:
  """
  The run whose words open paragraph `p`, `None` when a link, a formula, a picture or a field
  opens it, since `p.runs` passes over them and would hand on a letter from mid-sentence.
  """
  skipped = {qn(tag) for tag in _BEFORE_TEXT}
  blank = {qn(tag) for tag in _BLANK}
  for child in p._p:
    if child.tag in skipped: continue
    if child.tag != qn("w:r"): return None
    if "".join(t.text or "" for t in child.iter(qn("w:t"))).strip(): return Run(child, p)
    if any(part.tag not in blank for part in child): return None
  return None

def _letter_face(run:Run, size:float, sink:float):
  """`w:rPr` of the letter: the face of `run`, `size` pt tall and sunk `sink` pt."""
  rpr = copy.deepcopy(run._r.rPr) if run._r.rPr is not None else OxmlElement("w:rPr")
  for tag in ("w:position", "w:sz", "w:szCs"):
    for old in rpr.findall(qn(tag)): rpr.remove(old)
  size_el = OxmlElement("w:sz", {qn("w:val"): str(round(size * 2))}) # half-points
  following = next((c for c in rpr if c.tag in {qn(tag) for tag in _AFTER_SIZE}), None)
  if following is None: rpr.append(size_el)
  else: following.addprevious(size_el)
  size_el.addprevious(OxmlElement("w:position", {qn("w:val"): str(round(-sink * 2))}))
  return rpr
