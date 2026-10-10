# docmarq/paper/presets.py

"""
Language presets for `PaperStyle`: the engine's labels, those of the parts and the banner's own.

Example:
  >>> from docmarq.paper import lang_style
  >>> style = lang_style("pl", font_body="Cambria")
"""
import copy
from ..md import LANG_PRESETS as _ENGINE
from ..parts import LANG_LABELS
from .style import PaperStyle

#------------------------------------------------------------------------------------------ Presets

# Banner labels per language; English is the `PaperStyle` default.
_LABELS: dict[str, dict] = {
  "pl": {
    "banner_label_author": "Autor",
    "banner_label_created": "Utworzono",
    "banner_label_updated": "Zaktualizowano",
  },
  "de": {
    "banner_label_author": "Autor",
    "banner_label_created": "Erstellt",
    "banner_label_updated": "Aktualisiert",
  },
  "fr": {
    "banner_label_author": "Auteur",
    "banner_label_created": "Créé",
    "banner_label_updated": "Mis à jour",
  },
  "es": {
    "banner_label_author": "Autor",
    "banner_label_created": "Creado",
    "banner_label_updated": "Actualizado",
  },
  "it": {
    "banner_label_author": "Autore",
    "banner_label_created": "Creato",
    "banner_label_updated": "Aggiornato",
  },
  "cs": {
    "banner_label_author": "Autor",
    "banner_label_created": "Vytvořeno",
    "banner_label_updated": "Aktualizováno",
  },
  "sk": {
    "banner_label_author": "Autor",
    "banner_label_created": "Vytvorené",
    "banner_label_updated": "Aktualizované",
  },
}

# Full preset per language: the engine's callout labels, the parts' labels, then the banner's.
LANG_PRESETS: dict[str, dict] = {
  lang: {**_ENGINE[lang], **LANG_LABELS.get(lang, {}), **_LABELS.get(lang, {})} for lang in _ENGINE
}

#------------------------------------------------------------------------------------------ Builder

def lang_style(lang:str, **overrides) -> PaperStyle:
  """
  `PaperStyle` from a language preset and caller overrides, which win.

  An unknown `lang` falls back to `en`, the `PaperStyle` defaults,
  so callers don't have to validate the language code themselves.
  Each style gets a copy of the preset: changing its labels leaves other styles alone.
  """
  return PaperStyle(**{**copy.deepcopy(LANG_PRESETS.get(lang, {})), **overrides})
