# docmarq/emoji.py

"""
Emoji set in a face that has them, for `DOCX.emoji()`.

Word draws a character its run's font lacks as an empty box, and text faces hold no emoji.
`font_emoji` gives each stretch of emoji a run of its own in an emoji face, before saving.
A symbol of the text block, as a tick or a star, stays in the text face,
which Word fills in from a symbol face, unless `U+FE0F` after it asks for an emoji.
"""
import copy, re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .utils import text_element

EMOJI_FONT = "Segoe UI Emoji" # Word on macOS puts Apple Color Emoji in its place
# one emoji: a pictograph, or a symbol `U+FE0F` asks to show as one, as a heart, `©` or a keycap
_ONE = r"(?:[\U0001F000-\U0001FAFF]|[#*0-9\u00A9\u00AE\u203C-\u2BFF]\uFE0F\u20E3?)"
# emoji side by side, a selector after one and a joiner between two kept inside, never alone
_EMOJI = re.compile(rf"{_ONE}(?:\uFE0F|\u200D?{_ONE})*")
_FONT_SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")

def font_emoji(root, font:str=EMOJI_FONT) -> None:
  """
  Every run under `root` whose text holds emoji cut where they start and end, in order,
  its emoji in runs in `font`, and every other child, as a tab or a break, in a run of its own.
  """
  for run in list(root.iter(qn("w:r"))):
    if not any(t.text and _EMOJI.search(t.text) for t in run.findall(qn("w:t"))): continue
    rpr = run.find(qn("w:rPr"))
    for child in [c for c in run if c is not rpr]:
      pieces = _pieces(child.text) if child.tag == qn("w:t") and child.text else [None]
      for piece in pieces:
        part = run.makeelement(run.tag, run.attrib)
        if rpr is not None: part.append(copy.deepcopy(rpr))
        part.append(child if piece is None else text_element(piece))
        if piece is not None and _EMOJI.fullmatch(piece): _font(part, font)
        run.addprevious(part)
    run.getparent().remove(run)

def _pieces(text:str) -> list[str]:
  """`text` cut where emoji start and end."""
  out, start = [], 0
  for m in _EMOJI.finditer(text):
    if m.start() > start: out.append(text[start:m.start()])
    out.append(m.group())
    start = m.end()
  if start < len(text): out.append(text[start:])
  return out

def _font(run, font:str):
  """`run` set in `font` for every script, its theme fonts dropped."""
  rpr = run.find(qn("w:rPr"))
  if rpr is None:
    rpr = OxmlElement("w:rPr")
    run.insert(0, rpr)
  fonts = rpr.find(qn("w:rFonts"))
  if fonts is None:
    fonts = OxmlElement("w:rFonts")
    style = rpr.find(qn("w:rStyle"))
    rpr.insert(0 if style is None else 1, fonts) # schema order: `w:rStyle`, then `w:rFonts`
  for name in list(fonts.attrib):
    if name.lower().endswith("theme"): del fonts.attrib[name]
  for slot in _FONT_SLOTS:
    fonts.set(qn(f"w:{slot}"), font)
