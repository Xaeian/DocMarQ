# docmarq/notes.py

"""
Footnotes: the part Word keeps them in, their styles, and the marks that tie a note to its text.

`DOCX.footnote()` writes a note's text through `Note`, this module the OOXML around it.
python-docx reads and writes no footnotes, so the part is built here.
"""
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.parts.story import StoryPart
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from .utils import text_element

_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
# the notes of a template saved by Word load as a story a note and its pictures can join
PartFactory.part_type_for.setdefault(_CONTENT_TYPE, StoryPart)
_TEXT_STYLE = "footnote text" # Word's own names, so its footnote tools take them up
_MARK_STYLE = "footnote reference"
# The separator and its continuation open every footnotes part, the ids Word gives them.
_SEPARATORS = (
  f'<w:footnotes {nsdecls("w", "r")}>' # `r` names the links and pictures of a note
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
  '<w:r><w:separator/></w:r></w:p></w:footnote>'
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0"/>'
  '</w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
  '</w:footnotes>'
)

class Note:
  """A footnote as a write target, as `DOCX.into` makes one of a table cell."""
  def __init__(self, element, part:StoryPart, style_id:str):
    self.element = element
    self.part = part # hyperlinks and pictures in the note relate to it
    self._style_id = style_id

  def add_paragraph(self, text:str="", style:str|None=None) -> Paragraph:
    """
    New paragraph at the end of the note, in the `footnote text` style.
    One of another style, as a list item or a heading, refuses, as a table does.
    """
    if style is not None: raise ValueError(f"a footnote holds plain paragraphs, not {style}")
    p = OxmlElement("w:p")
    p.get_or_add_pPr().get_or_add_pStyle().val = self._style_id
    self.element.append(p)
    paragraph = Paragraph(p, self)
    if text: paragraph.add_run(text)
    return paragraph

  def add_table(self, rows:int, cols:int):
    """A note holds paragraphs only, so a table refuses."""
    raise ValueError("a footnote holds paragraphs, not tables")

def notes_part(document) -> StoryPart:
  """
  The footnotes part of `document`, made on first use, and the two styles of its notes.
  A part Word wrote may come without them, since Word adds a style the first time it uses it.
  """
  _note_styles(document)
  main = document.part
  for rel in main.rels.values():
    if rel.reltype == RT.FOOTNOTES: return rel.target_part
  part = StoryPart(PackURI("/word/footnotes.xml"), _CONTENT_TYPE, parse_xml(_SEPARATORS),
    main.package)
  main.relate_to(part, RT.FOOTNOTES)
  return part

def _note_styles(document):
  """Word's `footnote text` and `footnote reference` styles, added where `document` lacks them."""
  names = {s.name for s in document.styles}
  if _TEXT_STYLE not in names:
    text = document.styles.add_style(_TEXT_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    text.base_style = document.styles["Normal"]
    text.font.size = Pt(9)
    text.paragraph_format.space_after = Pt(0)
  if _MARK_STYLE not in names:
    document.styles.add_style(_MARK_STYLE, WD_STYLE_TYPE.CHARACTER).font.superscript = True

def settings_entry():
  """`w:footnotePr` naming the separators, for the document's settings."""
  entry = OxmlElement("w:footnotePr")
  for note_id in ("-1", "0"):
    entry.append(OxmlElement("w:footnote", {qn("w:id"): note_id}))
  return entry

def add_note(document, host:Paragraph, mark:str|None) -> Note:
  """
  New note of paragraph `host`, referred to at its end: by Word's next number with `mark=None`,
  by `mark` itself, or invisibly with `""`.
  """
  part = notes_part(document)
  root = part.element
  note_id = 1 + max(int(n.get(qn("w:id"))) for n in root.findall(qn("w:footnote")))
  element = OxmlElement("w:footnote", {qn("w:id"): str(note_id)})
  root.append(element)
  run = _mark_run(document)
  reference = OxmlElement("w:footnoteReference", {qn("w:id"): str(note_id)})
  run.append(reference)
  if mark is not None:
    reference.set(qn("w:customMarkFollows"), "1")
    # an empty mark still needs a character after it, a space of the Symbol font as Word writes
    run.append(text_element(mark) if mark else
      OxmlElement("w:sym", {qn("w:font"): "Symbol", qn("w:char"): "F020"}))
  host._p.append(run)
  return Note(element, part, document.styles[_TEXT_STYLE].style_id)

def mark_note(document, note:Note, mark:str|None):
  """
  Set the note's own mark before its first words: the number, `mark`, or nothing for `""`.
  A note left empty gets a paragraph, since Word holds a note without one broken.
  """
  first = note.element.find(qn("w:p"))
  if first is None: first = note.add_paragraph()._p
  if mark == "": return
  run = _mark_run(document)
  run.append(OxmlElement("w:footnoteRef") if mark is None else text_element(mark))
  space = OxmlElement("w:r")
  space.append(text_element(" "))
  ppr = first.find(qn("w:pPr"))
  at = 0 if ppr is None else list(first).index(ppr) + 1
  first.insert(at, run)
  first.insert(at + 1, space)

def drop_note(host:Paragraph, note:Note):
  """`note` taken back out, with the mark that refers to it in paragraph `host`."""
  note_id = note.element.get(qn("w:id"))
  for reference in host._p.iter(qn("w:footnoteReference")):
    if reference.get(qn("w:id")) == note_id:
      run = reference.getparent()
      run.getparent().remove(run)
      break
  note.element.getparent().remove(note.element)

def rule(document, show:bool):
  """The short rule above the footnotes drawn, or left out."""
  root = notes_part(document).element
  separator = next(n for n in root.findall(qn("w:footnote")) if n.get(qn("w:type")) == "separator")
  p = separator.find(qn("w:p"))
  for run in p.findall(qn("w:r")):
    p.remove(run)
  if show:
    run = OxmlElement("w:r")
    run.append(OxmlElement("w:separator"))
    p.append(run)

def _mark_run(document):
  """Empty run in the `footnote reference` style."""
  run = OxmlElement("w:r")
  rpr = OxmlElement("w:rPr")
  rpr.append(OxmlElement("w:rStyle", {qn("w:val"): document.styles[_MARK_STYLE].style_id}))
  run.append(rpr)
  return run
