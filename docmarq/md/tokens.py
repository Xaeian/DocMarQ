"""
Helpers for walking markdown-it token lists.
`*_open` / `*_close` pair structure requires attr lookup and depth-tracking close-finding;
this module centralises both so callers stay free of boilerplate.
"""
import re
from markdown_it.token import Token

#-------------------------------------------------------------------------------------------- Attrs

def get_attr(token:Token, name:str) -> str|None:
  """Return the value of HTML attribute `name` from a token, or `None`."""
  if not token.attrs: return None
  if isinstance(token.attrs, dict): return token.attrs.get(name)
  for k, v in token.attrs:
    if k == name: return v
  return None

#----------------------------------------------------------------------------------- Range scanning

def find_close(tokens:list[Token], start:int, open_type:str, close_type:str) -> int:
  """
  Return index of the matching close token for the open token at `start`.
  Falls back to last index on imbalance -
  markdown-it always produces balanced output for valid input, so this is a safety net only.
  """
  depth = 0
  for j in range(start, len(tokens)):
    tt = tokens[j].type
    if tt == open_type:
      depth += 1
    elif tt == close_type:
      depth -= 1
      if depth == 0: return j
  return len(tokens) - 1

#----------------------------------------------------------------------------------------- Callouts

# `> [!NOTE]` style marker matching GitHub-flavored markdown callouts.
CALLOUT_RE = re.compile(r"^\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$",
  re.IGNORECASE)

#--------------------------------------------------------------------------------------- Directives

# HTML-comment directives for layout control. Whitespace-tolerant, case-insensitive.
# Any extra content disqualifies the match (e.g. `<!-- pagebreak xxx -->` is not a directive).
_PAGEBREAK_RE = re.compile(r"\s*<!--\s*pagebreak\s*-->\s*", re.IGNORECASE)
# `wide`, the one option a group takes, spans it across columns where a package sets them
_GROUP_OPEN_RE = re.compile(r"\s*<!--\s*group(\s+wide)?\s*-->\s*", re.IGNORECASE)
_GROUP_CLOSE_RE = re.compile(r"\s*<!--\s*/\s*group\s*-->\s*", re.IGNORECASE)

def is_pagebreak_directive(content:str) -> bool:
  """True for `<!-- pagebreak -->` `html_block` content."""
  return bool(_PAGEBREAK_RE.fullmatch(content))

def is_group_open_directive(content:str) -> bool:
  """True for `<!-- group -->` opening directive, or `<!-- group wide -->`."""
  return bool(_GROUP_OPEN_RE.fullmatch(content))

def group_options(content:str) -> list[str]:
  """
  Options of a group opening: `["wide"]` of `<!-- group wide -->`, `[]` of a plain group.
  The engine sets every group alike, and the options are for a package that lays groups out.
  """
  match = _GROUP_OPEN_RE.fullmatch(content)
  return (match[1] or "").lower().split() if match else []

def is_group_close_directive(content:str) -> bool:
  """True for `<!-- /group -->` closing directive."""
  return bool(_GROUP_CLOSE_RE.fullmatch(content))

#------------------------------------------------------------------------------------------ Binding

NBSP = "\u00a0"
# a number as text writes it: `5`, `3,5`, `1.5`
_NUMBER = re.compile(r"\d+(?:[.,]\d+)*")
# a word after which a sentence or a clause starts: a full stop or a colon, maybe in quotes
_SENTENCE_END = re.compile(r"""[.!?…:][]"”’')»]*$""")
# the units a number keeps to: a letter after a number is one of them, never a word
_UNITS = frozenset((
  "s", "ms", "µs", "us", "ns", "min", "h",
  "V", "mV", "kV", "A", "mA", "µA", "W", "mW", "kW", "Wh", "kWh",
  "Ω", "kΩ", "MΩ", "F", "µF", "nF", "pF", "H", "mH", "µH",
  "Hz", "kHz", "MHz", "GHz", "dB", "B", "kB", "MB", "GB", "TB", "bit",
  "m", "mm", "cm", "km", "µm", "nm", "g", "kg", "mg", "t", "l", "ml",
  "N", "J", "Pa", "kPa", "bar", "K", "°C", "%", "pt", "px", "rpm",
))
# brackets and quotes a word may open with, as `(w` or `„w`
_OPENERS = "([„«'" + '"'
# a formula, a picture or code within a word, which is neither a letter nor a number
_OBJECT = chr(0xFFFC)

def binds(before:str|None, word:str, after:str, letters:str) -> bool:
  """
  Whether no line may end between `word` and the word `after` it.
  `before` is the word before `word`, `None` where a paragraph or a line starts.
  A one-letter word of `letters` keeps to the next one, as `aiouwz` in Polish,
  and so does one opening a bracket or a quote, as `(w`.
  A unit after a number never does, as `V` in `5 V` or Czech `s` in `5 s`.
  A capital binds where a sentence or a clause opens, after a stop, a colon or a bracket,
  or among capitals, as in a title.
  Where `letters` binds anything, a number keeps to its unit as well.
  """
  if not letters: return False
  if _NUMBER.fullmatch(word): return after.rstrip(".,;:!?)") in _UNITS
  letter = word.lstrip(_OPENERS)
  if len(letter) != 1 or letter.lower() not in letters: return False
  if before is not None and _NUMBER.fullmatch(before) and word in _UNITS: return False
  if letter.islower() or letter != word or before is None: return True
  return bool(_SENTENCE_END.search(before)) or (before + after).isupper()

def _bound_spaces(text:str, letters:str) -> list[int]:
  """Positions of the spaces in `text` no line may end at, as `binds` decides from its words."""
  words = list(re.finditer(r"\S+", text))
  found = []
  for k, word in enumerate(words[:-1]):
    after = words[k + 1]
    gap = text[word.end():after.start()]
    if not gap.startswith(" ") or "\n" in gap: continue
    opens = not k or "\n" in text[words[k - 1].end():word.start()]
    before = None if opens else words[k - 1].group()
    if binds(before, word.group(), after.group(), letters): found.append(word.end())
  return found

def bind_letters(text:str, letters:str) -> str:
  """`text` with a no-break space wherever `binds` keeps two words together, which Word keeps."""
  chars = list(text)
  for at in _bound_spaces(text, letters): chars[at] = NBSP
  return "".join(chars)

def bind_children(children:list[Token], letters:str) -> list[Token]:
  """
  Inline `children` with a no-break space wherever `binds` keeps two words together.
  A word runs across emphasis and links, so a bold one-letter word binds as well,
  and a soft break that binds gets a no-break space for its content.
  Code, a formula, a picture or a note mark in a word makes it one no rule binds.
  The children come back as copies where they change, so the parsed tokens stay as they were.
  """
  pieces = [_piece(child) for child in children]
  bound = _bound_spaces("".join(pieces), letters)
  out, at = [], 0
  for child, piece in zip(children, pieces):
    hits = [k - at for k in bound if at <= k < at + len(piece)]
    at += len(piece)
    if not hits:
      out.append(child)
      continue
    if child.type == "softbreak":
      out.append(child.copy(content=NBSP))
      continue
    chars = list(child.content)
    for k in hits: chars[k] = NBSP
    out.append(child.copy(content="".join(chars)))
  return out

def _piece(child:Token) -> str:
  """What `child` reads as among the words around it."""
  if child.type == "text": return child.content
  if child.type == "softbreak": return " "
  if child.type == "hardbreak": return "\n"
  if child.type in ("code_inline", "math_inline", "image", "footnote_ref"): return _OBJECT
  return ""
