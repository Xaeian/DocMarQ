# docmarq/fonts.py

"""Font configuration helpers.

Word fonts are referenced by family name only; rasterization is host-side.
No embedding: if the target machine lacks the font, Word substitutes silently.
To embed, add `<w:embedTrueTypeFonts/>` to `settings.xml`.
"""

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
