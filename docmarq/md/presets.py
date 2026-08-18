# docmarq/md/presets.py

"""Built-in language presets for `MarkdownStyle`.

Each preset is a dict of fields applied to a fresh `MarkdownStyle` via
`lang_style(lang, **overrides)`. Ships `en`, `pl`, `de`, `fr`, `es`,
`it`, `cs`, `sk` - add more by extending `LANG_PRESETS`.

  >>> from docmarq.md import lang_style
  >>> style = lang_style("pl", font_body="IBMPlexSans")
"""

from .style import MarkdownStyle

#------------------------------------------------------------------------------------------ Presets

# `en` keeps the `MarkdownStyle` defaults - empty preset is intentional.
LANG_PRESETS:dict[str, dict] = {
  "en": {},
  "pl": {
    "page_number_label": "Strona",
    "date_format": "%d.%m.%Y",
    "banner_label_author": "Autor",
    "banner_label_created": "Utworzono",
    "banner_label_updated": "Zaktualizowano",
    "sign_labels": {"signature": ["Podpis"], "approval": ["Opracował", "Zatwierdził"],
      "contract": ["Zleceniodawca", "Zleceniobiorca"]},
    "callout_label_note": "Notatka",
    "callout_label_tip": "Wskazówka",
    "callout_label_important": "Ważne",
    "callout_label_warning": "Ostrzeżenie",
    "callout_label_caution": "Uwaga",
  },
  "de": {
    "page_number_label": "Seite",
    "date_format": "%d.%m.%Y",
    "banner_label_author": "Autor",
    "banner_label_created": "Erstellt",
    "banner_label_updated": "Aktualisiert",
    "sign_labels": {"signature": ["Unterschrift"], "approval": ["Erstellt", "Genehmigt"],
      "contract": ["Auftraggeber", "Auftragnehmer"]},
    "callout_label_note": "Hinweis",
    "callout_label_tip": "Tipp",
    "callout_label_important": "Wichtig",
    "callout_label_warning": "Warnung",
    "callout_label_caution": "Achtung",
  },
  "fr": {
    "page_number_label": "Page",
    "date_format": "%d/%m/%Y",
    "banner_label_author": "Auteur",
    "banner_label_created": "Créé",
    "banner_label_updated": "Mis à jour",
    "sign_labels": {"signature": ["Signature"], "approval": ["Rédigé par", "Approuvé par"],
      "contract": ["Donneur d'ordre", "Prestataire"]},
    "callout_label_note": "Note",
    "callout_label_tip": "Astuce",
    "callout_label_important": "Important",
    "callout_label_warning": "Avertissement",
    "callout_label_caution": "Attention",
  },
  "es": {
    "page_number_label": "Página",
    "date_format": "%d/%m/%Y",
    "banner_label_author": "Autor",
    "banner_label_created": "Creado",
    "banner_label_updated": "Actualizado",
    "sign_labels": {"signature": ["Firma"], "approval": ["Elaborado por", "Aprobado por"],
      "contract": ["Cliente", "Contratista"]},
    "callout_label_note": "Nota",
    "callout_label_tip": "Consejo",
    "callout_label_important": "Importante",
    "callout_label_warning": "Advertencia",
    "callout_label_caution": "Precaución",
  },
  "it": {
    "page_number_label": "Pagina",
    "date_format": "%d/%m/%Y",
    "banner_label_author": "Autore",
    "banner_label_created": "Creato",
    "banner_label_updated": "Aggiornato",
    "sign_labels": {"signature": ["Firma"], "approval": ["Redatto da", "Approvato da"],
      "contract": ["Committente", "Appaltatore"]},
    "callout_label_note": "Nota",
    "callout_label_tip": "Suggerimento",
    "callout_label_important": "Importante",
    "callout_label_warning": "Avviso",
    "callout_label_caution": "Attenzione",
  },
  "cs": {
    "page_number_label": "Strana",
    "date_format": "%d.%m.%Y",
    "banner_label_author": "Autor",
    "banner_label_created": "Vytvořeno",
    "banner_label_updated": "Aktualizováno",
    "sign_labels": {"signature": ["Podpis"], "approval": ["Vypracoval", "Schválil"],
      "contract": ["Objednatel", "Zhotovitel"]},
    "callout_label_note": "Poznámka",
    "callout_label_tip": "Tip",
    "callout_label_important": "Důležité",
    "callout_label_warning": "Varování",
    "callout_label_caution": "Pozor",
  },
  "sk": {
    "page_number_label": "Strana",
    "date_format": "%d.%m.%Y",
    "banner_label_author": "Autor",
    "banner_label_created": "Vytvorené",
    "banner_label_updated": "Aktualizované",
    "sign_labels": {"signature": ["Podpis"], "approval": ["Vypracoval", "Schválil"],
      "contract": ["Objednávateľ", "Zhotoviteľ"]},
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
