"""Helpers for walking markdown-it token lists.
`*_open` / `*_close` pair structure requires attr lookup and depth-tracking
close-finding; this module centralises both so callers stay free of boilerplate."""
import re
from markdown_it.token import Token

#-------------------------------------------------------------------------------------- Attrs

def get_attr(token:Token, name:str) -> str|None:
  """Return the value of HTML attribute `name` from a token, or None."""
  if not token.attrs:
    return None
  if isinstance(token.attrs, dict):
    return token.attrs.get(name)
  for k, v in token.attrs:
    if k == name:
      return v
  return None

#----------------------------------------------------------------------------- Range scanning

def find_close(tokens:list[Token], start:int, open_type:str, close_type:str) -> int:
  """Return index of the matching close token for the open token at `start`.
  Falls back to last index on imbalance - markdown-it always produces
  balanced output for valid input, so this is a safety net only."""
  depth = 0
  for j in range(start, len(tokens)):
    tt = tokens[j].type
    if tt == open_type:
      depth += 1
    elif tt == close_type:
      depth -= 1
      if depth == 0:
        return j
  return len(tokens) - 1

#---------------------------------------------------------------------------------- Callouts

# `> [!NOTE]` style marker matching GitHub-flavored markdown callouts.
CALLOUT_RE = re.compile(r"^\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$",
  re.IGNORECASE)

#---------------------------------------------------------------------------- Directives

# HTML-comment directives for layout control. Whitespace-tolerant, case-insensitive.
# Any extra content disqualifies the match (e.g. `<!-- pagebreak xxx -->` is not a
# directive). Symmetric with `pdfmarq.md.md_html` detectors.
_PAGEBREAK_RE = re.compile(r"\s*<!--\s*pagebreak\s*-->\s*", re.IGNORECASE)
_GROUP_OPEN_RE = re.compile(r"\s*<!--\s*group\s*-->\s*", re.IGNORECASE)
_GROUP_CLOSE_RE = re.compile(r"\s*<!--\s*/\s*group\s*-->\s*", re.IGNORECASE)

def is_pagebreak_directive(content:str) -> bool:
  """True for `<!-- pagebreak -->` html_block content."""
  return bool(_PAGEBREAK_RE.fullmatch(content))

def is_group_open_directive(content:str) -> bool:
  """True for `<!-- group -->` opening directive."""
  return bool(_GROUP_OPEN_RE.fullmatch(content))

def is_group_close_directive(content:str) -> bool:
  """True for `<!-- /group -->` closing directive."""
  return bool(_GROUP_CLOSE_RE.fullmatch(content))
