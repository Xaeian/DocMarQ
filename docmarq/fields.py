# docmarq/fields.py

"""
Word fields: values Word works out as it lays the pages out, as the page number and page count.

A simple field is one element.
A complex one is a run of marks - begin, instruction, separate, result, end -
and its instruction may hold other fields, as `IF NUMPAGES > 1` holds `PAGE`.
"""
import re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run

_PLACEHOLDER = re.compile(r"(\{page\}|\{pages\})")
_CODES = {"{page}": "PAGE", "{pages}": "NUMPAGES"}

#------------------------------------------------------------------------------------------- Simple

def simple_field(p, instr:str, look=None):
  """
  Append a field running `instr` to paragraph `p`, showing `0` until Word works it out.
  `look(run)` formats the run it shows in.
  """
  fld = OxmlElement("w:fldSimple")
  fld.set(qn("w:instr"), instr)
  run = Run(OxmlElement("w:r"), p)
  if look: look(run)
  run.text = "0"
  fld.append(run._r)
  p._p.append(fld)

def page_fields(p, template:str, look=None):
  """`template` into paragraph `p`, `{page}` and `{pages}` as fields, `look(run)` on each run."""
  for part in _PLACEHOLDER.split(template):
    if part in _CODES: simple_field(p, _CODES[part], look)
    elif part:
      run = p.add_run(part)
      if look: look(run)

#------------------------------------------------------------------------------------------ Complex

def complex_field(p, instr:str, shown:str="", look=None):
  """
  A field running `instr` at the end of paragraph `p`, as a table of contents is one:
  `shown` stands in its place until Word works it out, `look(run)` formats every run.
  """
  run = _runs(p, look)
  _mark(run(), "begin")
  _instr(run(), f" {instr} ")
  _mark(run(), "separate")
  if shown: run().text = shown
  _mark(run(), "end")

def page_fields_if_many(p, template:str, look=None):
  """
  `template` into paragraph `p`, shown only in a document of more than one page.

  It goes in as the field `IF NUMPAGES > 1 "template" ""`,
  each `{page}` and `{pages}` a field within it.
  `look(run)` formats every run, so the text Word puts in its place takes that look too.
  The result stored for readers that never work fields out is empty.
  """
  run = _runs(p, look)
  def nested(code:str):
    _mark(run(), "begin")
    _instr(run(), f" {code} ")
    _mark(run(), "separate")
    run().text = "1"
    _mark(run(), "end")
  _mark(run(), "begin")
  _instr(run(), " IF ")
  nested("NUMPAGES")
  _instr(run(), ' > 1 "')
  for part in _PLACEHOLDER.split(template):
    if part in _CODES: nested(_CODES[part])
    elif part: _instr(run(), part.replace("\\", "\\\\").replace('"', '\\"'))
  _instr(run(), '" "" ')
  _mark(run(), "separate")
  _mark(run(), "end")

def _runs(p, look):
  """Maker of the runs of a field in paragraph `p`, each formatted by `look(run)`."""
  def run():
    r = p.add_run()
    if look: look(r)
    return r
  return run

def _mark(run, kind:str):
  """`run` as the field mark `kind`: begin, separate or end."""
  el = OxmlElement("w:fldChar")
  el.set(qn("w:fldCharType"), kind)
  run._r.append(el)

def _instr(run, text:str):
  """`run` holding `text` of a field's instruction."""
  el = OxmlElement("w:instrText")
  el.set(qn("xml:space"), "preserve")
  el.text = text
  run._r.append(el)
