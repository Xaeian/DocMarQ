# docmarq/fonts.py

"""Font configuration helpers.

Word fonts are referenced by family name only; rasterization is host-side.
No embedding: if the target machine lacks the font, Word substitutes silently.
To embed, add `<w:embedTrueTypeFonts/>` to `settings.xml`.
"""

#-------------------------------------------------------------------------------------- Helpers

def is_safe_default(family:str) -> bool:
  """Returns `True` for fonts present on virtually all Word installs."""
  return family.lower() in {
    "calibri", "cambria", "arial", "times new roman", "verdana",
    "tahoma", "georgia", "courier new", "consolas", "segoe ui",
  }
