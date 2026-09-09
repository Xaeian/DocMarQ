# docmarq/fonts.py

"""Font configuration helpers.

Word fonts are referenced by family name only; rasterization is host-side.
No embedding: if the target machine lacks the font, Word substitutes silently.
To embed, add `<w:embedTrueTypeFonts/>` to `settings.xml`.

TTFs still matter for what this package draws itself rather than handing to Word:
mermaid diagrams and text inside an SVG. `resolve_ttf` is the one lookup both use.
"""
from pathlib import Path

#------------------------------------------------------------------------------------------ Helpers

# Present on virtually every Word install, so a document naming one of these
# renders as intended without embedding. Exposed as data, not only through the
# predicate, because a caller building a font picker needs the list itself.
SAFE_FAMILIES: tuple[str, ...] = (
  "Arial", "Calibri", "Cambria", "Consolas", "Courier New",
  "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Verdana",
)

def is_safe_default(family:str) -> bool:
  """Returns `True` for fonts present on virtually all Word installs."""
  return family.lower() in {f.lower() for f in SAFE_FAMILIES}

#--------------------------------------------------------------------------------------- TTF lookup

# Mode fallback when a variant TTF is missing (drops styling, keeps content).
_MODE_FALLBACK = {
  "Bold": ["Regular"],
  "Italic": ["Regular"],
  "BoldItalic": ["Bold", "Italic", "Regular"],
  "Black": ["Bold", "Regular"],
  "BlackItalic": ["BoldItalic", "Black", "Bold", "Italic", "Regular"],
}

def resolve_ttf(font_dir:str, family:str, mode:str="Regular") -> Path|None:
  """Find the TTF backing one face, walking `_MODE_FALLBACK` when it is absent.

  Layout is `<font_dir>/<family>/<family>-<mode>.ttf`, with the subfolder named
  either as-is or lowercase, and a flat `<font_dir>` accepted as well.
  """
  base = Path(font_dir)
  for try_mode in [mode] + _MODE_FALLBACK.get(mode, []):
    for sub in (family.lower(), family):
      path = base / sub / f"{family}-{try_mode}.ttf"
      if path.is_file(): return path
    path = base / f"{family}-{try_mode}.ttf"
    if path.is_file(): return path
  return None
