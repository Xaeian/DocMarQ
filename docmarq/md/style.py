# docmarq/md/style.py

"""
Markdown-rendering style: fonts, sizes, colors, spacing and the switches of the engine.

Word lays the pages out itself, so there are no measuring or page-break fields.
"""
from dataclasses import dataclass, field

#-------------------------------------------------------------------------- Callout palette default

def _default_callout_colors() -> dict:
  """GitHub-style callout colors per type (border, text/icon).
  Each entry: lowercase type → `(border_rgb, text_rgb)` in 0..1 range.
  """
  return {
    "note": ((0.035, 0.41, 0.855), (0.035, 0.41, 0.855)), # blue
    "tip": ((0.12, 0.53, 0.24), (0.12, 0.53, 0.24)), # green
    "important": ((0.51, 0.31, 0.87), (0.51, 0.31, 0.87)), # purple
    "warning": ((0.60, 0.40, 0.0), (0.60, 0.40, 0.0)), # amber
    "caution": ((0.81, 0.13, 0.18), (0.81, 0.13, 0.18)), # red
  }

#------------------------------------------------------------------------------------ MarkdownStyle

@dataclass
class MarkdownStyle:
  """Visual style for markdown rendering. GitHub-light defaults."""
  # Fonts - both ship with Windows out of the box (Vista+),
  # so Word users don't need to install anything.
  # Calibri is the canonical Office body font; Consolas is its companion monospace.
  font_body: str = "Calibri"
  font_head: str = "Calibri"
  font_mono: str = "Consolas"

  # Body size in pt.
  body_size: float = 11

  # Body line-height multiplier.
  # 1.0 = Word's "Single" rule, which renders ~1.2x font size due to Calibri's internal leading,
  # tight without being cramped.
  line_height: float = 1.0

  # Paragraph gap in mm. Converted to Word's pt-based `space_after` at apply time.
  para_gap: float = 3

  #--------------------------------------------------------------------------------------- Headings

  # Sizes of Word's `Heading 1..6` styles in pt. `None` follows `body_size`.
  h1_size: float|None = 20
  h2_size: float|None = 16
  h3_size: float|None = 13
  h4_size: float|None = 11
  h5_size: float|None = 11
  h6_size: float|None = 11
  # Thin rule under `#` and `##`.
  h1_underline: bool = True
  h2_underline: bool = True
  # Face of the headings: `Regular`, `Bold`, `Italic` or `BoldItalic`.
  head_mode: str = "Bold"
  # mm above a heading, on top of the paragraph gap, and below it.
  # `None` keeps the gaps of each level: 10/4, 8/3, 6/3, 4/2, 3/2 and 3/2 pt.
  head_gap_top: float|None = None
  head_gap_bot: float|None = None

  #-------------------------------------------------------------------------------------- Text flow

  # Paragraphs, list items, quotes and notes set flush to both edges. Their last line stays left.
  justify: bool = False
  # Language whose hyphenation rules split words at line ends, as `pl`. `None` keeps words whole.
  # Word hyphenates by the language of the text, so this sets the language of the document.
  hyphenate: str|None = None
  # `True` keeps each paragraph on one page, and a list on one page with its items.
  # `False` leaves page breaks to Word.
  keep_together: bool = False

  #------------------------------------------------------------------------------------------- Math

  # Math formulas (`$...$`, `$$...$$`) render as native Word equations (OMML):
  # editable, selectable, scalable.
  # Constructs outside the supported LaTeX subset fall back to a matplotlib-rendered image,
  # so raw LaTeX never leaks into the output.
  # Set `math_enable=False` to leave `$...$` as plain text.
  math_enable: bool = True
  # Fallback image fontset (matplotlib mathtext):
  # `stix` / `stixsans` / `cm` / `dejavusans` / `dejavuserif`.
  # `stix` is closest to Word's Cambria Math,
  # so a fallback formula blends with the native ones around it.
  math_fontset: str = "stix"

  #---------------------------------------------------------------------------------------- Mermaid

  # Mermaid fenced blocks (```mermaid) render through the `mmdc` CLI to a PNG
  # and get embedded as a regular figure.
  # When `mmdc` is missing or the diagram fails to compile,
  # the mermaid.ink HTTP service is tried next (see `mermaid_remote`),
  # and only then does the source degrade to a regular code block.
  # `mermaid_enable=False` skips straight to the code block without touching either backend.
  mermaid_enable: bool = True
  # Allow the mermaid.ink fallback.
  # It uploads the diagram source to a public third-party service and warns once per process;
  # `False` keeps rendering strictly local.
  mermaid_remote: bool = True
  mermaid_theme: str = "default" # `mmdc` theme: default / dark / forest / neutral
  mermaid_background: str = "transparent" # bg color; transparent matches page
  mermaid_scale: float = 3 # PNG oversampling factor
  mermaid_cli: str = "mmdc" # CLI binary name / path

  # Colors (rgb 0-1) - GitHub-light palette
  body_color: tuple = (0.09, 0.11, 0.13) # #1f2328
  head_color: tuple = (0.09, 0.11, 0.13) # Word's `Heading 1..6`, and a letter's title
  muted_color: tuple = (0.40, 0.44, 0.50) # for blockquote text
  link_color: tuple = (0.03, 0.41, 0.85) # #0969da
  code_inline_color: tuple = (0.09, 0.11, 0.13)
  code_inline_bg: tuple = (0.96, 0.97, 0.98) # subtle grey
  code_block_bg: tuple = (0.96, 0.97, 0.98)
  code_block_border: tuple = (0.82, 0.84, 0.87) # #d0d7de
  quote_border: tuple = (0.82, 0.84, 0.87)
  hr_color: tuple = (0.82, 0.84, 0.87)
  # Inert: `==mark==` stays literal text; kept so a pdfmarq style carries over.
  mark_bg: str = "yellow"
  hr_thick: float = 0.5 # pt - rules and signing lines; Word keeps eighths of a point

  # Image sizing (mm)
  image_max_h: float = 120 # cap block image height
  image_min_dpi: int = 150 # block image fills text width down to this; 0 = always
  # Inert: inline images render as italic alt text, which no height cap applies to.
  # Block figures answer to `image_max_h`.
  inline_image_max_h: float = 5.5 # cap inline image height (~2ex @ 11pt)

  # Local-link handling.
  # `[x](file.md)` with no schema gets resolved to `{link_root}/{link_base}/file.md`
  # when `link_root` is set.
  # Without it, the text is rendered but not linked.
  link_root: str|None = None
  link_base: str = ""

  # GitHub callout titles - override for localization
  callout_label_note: str = "Note"
  callout_label_tip: str = "Tip"
  callout_label_important: str = "Important"
  callout_label_warning: str = "Warning"
  callout_label_caution: str = "Caution"

  # Callout palette (border + text colors per type)
  callout_colors: dict = field(default_factory=_default_callout_colors)

  #-------------------------------------------------------------------------------------- Footnotes

  # Footnote section heading.
  # `None` (default) emits just a thin HR above the bibliography:
  # the smaller font signals reference matter.
  # Set to a string (e.g. `"References"`, `"Bibliografia"`)
  # to add an H2 heading above the footnotes.
  footnote_label: str|None = None
