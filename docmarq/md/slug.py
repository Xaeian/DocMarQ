"""
Heading slug generation and `#anchor` link resolution.

GitHub-style slugs: lower-case, spaces to hyphens, drop non-word chars.
Unicode-aware so Polish / German / Czech headings get usable slugs
(`# Błędy` → `błędy`). Pre-scanning all heading slugs before rendering
lets us catch broken `[text](#typo)` links and render them as plain text
instead of producing dangling jumps.
"""
import urllib.parse
from markdown_it.token import Token

#------------------------------------------------------------------------------------------ Slugify

def slugify_inline(inline_token:Token) -> str:
  """GitHub-style slug from an inline token's children. Lower-case, spaces
  collapsed to single hyphens, non-word chars dropped. Unicode letters and
  digits pass through `str.isalnum`.
  """
  text = "".join(c.content for c in (inline_token.children or [])
    if c.type in ("text", "code_inline"))
  text = text.strip().lower()
  out = []
  prev_dash = False
  for ch in text:
    if ch.isspace() or ch == "-":
      if not prev_dash and out:
        out.append("-")
        prev_dash = True
    elif ch.isalnum() or ch == "_":
      out.append(ch)
      prev_dash = False
  return "".join(out).strip("-")

def collect_heading_slugs(tokens:list[Token]) -> set[str]:
  """Walk tokens and collect every heading's slug. Feeds `resolve_anchor`
  so unknown `#anchor` hrefs render as plain text instead of producing
  broken jumps.
  """
  slugs:set[str] = set()
  for i, t in enumerate(tokens):
    if t.type != "heading_open": continue
    if i + 1 < len(tokens) and tokens[i + 1].type == "inline":
      slug = slugify_inline(tokens[i + 1])
      if slug:
        slugs.add(slug)
  return slugs

#------------------------------------------------------------------------------------ Anchor lookup

def resolve_anchor(href:str, known:set[str]) -> str|None:
  """Map a `#slug` href to a known heading bookmark name. Returns the slug
  on match, `None` otherwise. Handles markdown-it's URL-encoded non-ASCII
  slugs (PL/DE/CZ/SK characters get percent-encoded by the parser).
  """
  if not href.startswith("#"):
    return None
  slug = urllib.parse.unquote(href[1:])
  return slug if slug in known else None
