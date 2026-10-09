# docmarq/md/presets.py

"""Built-in language presets for `MarkdownStyle`: the labels the engine writes itself.

Each preset is a dict of fields applied to a fresh `MarkdownStyle`
via `lang_style(lang, **overrides)`.
Ships `en`, `pl`, `de`, `fr`, `es`, `it`, `cs`, `sk`.
Extend by adding entries to `LANG_PRESETS`.
The document packages layer their own labels on top.

  >>> from docmarq.md import lang_style
  >>> style = lang_style("pl", font_body="Cambria")
"""

from .style import MarkdownStyle

#------------------------------------------------------------------------------------------ Presets

# `en` keeps the `MarkdownStyle` defaults - empty preset is intentional.
LANG_PRESETS:dict[str, dict] = {
  "en": {},
  "pl": {
    "callout_label_note": "Notatka",
    "callout_label_tip": "Wskazówka",
    "callout_label_important": "Ważne",
    "callout_label_warning": "Ostrzeżenie",
    "callout_label_caution": "Uwaga",
  },
  "de": {
    "callout_label_note": "Hinweis",
    "callout_label_tip": "Tipp",
    "callout_label_important": "Wichtig",
    "callout_label_warning": "Warnung",
    "callout_label_caution": "Achtung",
  },
  "fr": {
    "callout_label_note": "Note",
    "callout_label_tip": "Astuce",
    "callout_label_important": "Important",
    "callout_label_warning": "Avertissement",
    "callout_label_caution": "Attention",
  },
  "es": {
    "callout_label_note": "Nota",
    "callout_label_tip": "Consejo",
    "callout_label_important": "Importante",
    "callout_label_warning": "Advertencia",
    "callout_label_caution": "Precaución",
  },
  "it": {
    "callout_label_note": "Nota",
    "callout_label_tip": "Suggerimento",
    "callout_label_important": "Importante",
    "callout_label_warning": "Avviso",
    "callout_label_caution": "Attenzione",
  },
  "cs": {
    "callout_label_note": "Poznámka",
    "callout_label_tip": "Tip",
    "callout_label_important": "Důležité",
    "callout_label_warning": "Varování",
    "callout_label_caution": "Pozor",
  },
  "sk": {
    "callout_label_note": "Poznámka",
    "callout_label_tip": "Tip",
    "callout_label_important": "Dôležité",
    "callout_label_warning": "Upozornenie",
    "callout_label_caution": "Pozor",
  },
}

#------------------------------------------------------------------------------------------ Builder

def lang_style(lang:str, **overrides) -> MarkdownStyle:
  """Build a `MarkdownStyle` from a language preset + caller overrides.
  Unknown `lang` silently falls back to `en` so callers need not pre-validate.
  `overrides` win over preset values."""
  base = LANG_PRESETS.get(lang, {})
  return MarkdownStyle(**{**base, **overrides})
