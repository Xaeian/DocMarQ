# docmarq/legal/presets.py

"""
Language presets for `LegalStyle`: the engine's labels, those of the parts and the letter's own.

Example:
  >>> from docmarq.legal import lang_style
  >>> style = lang_style("pl", columns=2)
"""
import copy
from ..md.presets import LANG_PRESETS as _ENGINE
from ..parts import LANG_LABELS
from .style import LegalStyle

#------------------------------------------------------------------------------------------ Presets

# Labels of the end of a letter per language; English is the `LegalStyle` default.
_LABELS: dict[str, dict] = {
  "pl": {"attachments_label": "Załączniki", "ref_label": "Sygn. akt",
    "date_format": "%d.%m.%Y r."},
  "de": {"attachments_label": "Anlagen", "ref_label": "Az."},
  "fr": {"attachments_label": "Pièces jointes", "ref_label": "Réf."},
  "es": {"attachments_label": "Anexos", "ref_label": "Ref."},
  "it": {"attachments_label": "Allegati", "ref_label": "Rif."},
  "cs": {"attachments_label": "Přílohy", "ref_label": "Sp. zn."},
  "sk": {"attachments_label": "Prílohy", "ref_label": "Sp. zn."},
}

# Full preset per language: the engine's callout labels, the parts' labels, then the letter's,
# with the text hyphenated in that language.
LANG_PRESETS: dict[str, dict] = {
  lang: {**_ENGINE[lang], **LANG_LABELS.get(lang, {}), **_LABELS.get(lang, {}), "hyphenate": lang}
  for lang in _ENGINE
}

#------------------------------------------------------------------------------------------ Builder

def lang_style(lang:str, **overrides) -> LegalStyle:
  """
  `LegalStyle` from a language preset and caller overrides, which win.

  An unknown `lang` falls back to the `LegalStyle` defaults, English labels and no hyphenation,
  so callers don't have to validate the language code themselves.
  Each style gets a copy of the preset: changing its labels leaves other styles alone.
  """
  return LegalStyle(**{**copy.deepcopy(LANG_PRESETS.get(lang, {})), **overrides})
