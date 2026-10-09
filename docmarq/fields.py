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

def page_fields_if_many(p, template:str, look=None):
  """
  `template` into paragraph `p`, shown only in a document of more than one page.

  It goes in as the field `IF NUMPAGES > 1 "template" ""`,
  each `{page}` and `{pages}` a field within it.
  `look(run)` formats every run, so the text Word puts in its place takes that look too.
  The result stored for readers that never work fields out is empty.
  """
  def run():
    r = p.add_run()
    if look: look(r)
    return r
  def mark(kind:str):
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), kind)
    run()._r.append(el)
  def instr(text:str):
    el = OxmlElement("w:instrText")
    el.set(qn("xml:space"), "preserve")
    el.text = text
    run()._r.append(el)
  def nested(code:str):
    mark("begin")
    instr(f" {code} ")
    mark("separate")
    run().text = "1"
    mark("end")
  mark("begin")
  instr(" IF ")
  nested("NUMPAGES")
  instr(' > 1 "')
  for part in _PLACEHOLDER.split(template):
    if part in _CODES: nested(_CODES[part])
    elif part: instr(part.replace("\\", "\\\\").replace('"', '\\"'))
  instr('" "" ')
  mark("separate")
  mark("end")
