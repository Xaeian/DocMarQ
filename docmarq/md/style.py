# docmarq/md/style.py

"""Markdown-rendering style. Mirrors `pdfmarq.MarkdownStyle` semantics but
trimmed - Word handles layout/sizing itself so we don't need height/size
fields, just color palette, fonts, and a few format toggles."""
from dataclasses import dataclass, field

#---------------------------------------------------------------------- Callout palette default

def _default_callout_colors() -> dict:
  """GitHub-style callout colors per type (border, text/icon).
  Each entry: lowercase type → `(border_rgb, text_rgb)` in 0..1 range.
  Mirrors `pdfmarq.md.markdown_style._default_callout_colors`."""
  return {
    "note": ((0.035, 0.41, 0.855), (0.035, 0.41, 0.855)),  # blue
    "tip": ((0.12, 0.53, 0.24), (0.12, 0.53, 0.24)),  # green
    "important": ((0.51, 0.31, 0.87), (0.51, 0.31, 0.87)),  # purple
    "warning": ((0.60, 0.40, 0.0), (0.60, 0.40, 0.0)),  # amber
    "caution": ((0.81, 0.13, 0.18), (0.81, 0.13, 0.18)),  # red
  }

#------------------------------------------------------------------ Status badge palette default

def _default_status_colors() -> dict:
  """Status badge colors `(bg, text)` keyed by lowercase status name."""
  return {
    "draft": ((0.85, 0.87, 0.92), (0.30, 0.34, 0.45)),
    "review": ((1.00, 0.95, 0.78), (0.62, 0.40, 0.05)),
    "approved": ((0.86, 0.96, 0.87), (0.10, 0.45, 0.18)),
    "deprecated": ((1.00, 0.88, 0.85), (0.74, 0.20, 0.15)),
    "archived": ((0.92, 0.87, 0.96), (0.45, 0.25, 0.55)),
  }

#-------------------------------------------------------------------------------- MarkdownStyle

@dataclass
class MarkdownStyle:
  """Visual style for markdown rendering. GitHub-light defaults."""

  # Fonts - both ship with Windows out of the box (Vista+) so Word users
  # don't need to install anything. Calibri is the canonical Office body
  # font; Consolas is its companion monospace.
  body_family: str = "Calibri"
  head_family: str|None = None # `None` = heading inherits `body_family`
  mono_family: str = "Consolas"

  # Body line-height multiplier. 1.0 = Word's "Single" rule, which renders
  # ~1.2x font size due to Calibri's internal leading - tight without being
  # cramped. Headings have their own line spacing via Word's built-in
  # Heading styles (see `core.py` `tight_line_height`).
  line_height: float = 1.0

  # Paragraph gap in pt (Word `space_after`). 8pt ≈ pdfmarq's 3mm body gap,
  # so PDF and DOCX share the same vertical rhythm.
  para_gap_pt: float = 8

  #---------------------------------------------------------------------------------- Math

  # Math formulas (`$...$`, `$$...$$`) render as native Word equations (OMML)
  # - editable, selectable, scalable. Constructs outside the supported LaTeX
  # subset fall back to a matplotlib-rendered image so raw LaTeX never leaks
  # into the output. Set `math_enable=False` to leave `$...$` as plain text.
  math_enable: bool = True
  # Fallback image fontset (matplotlib mathtext): stix / stixsans / cm /
  # dejavusans / dejavuserif. `stix` is closest to Word's Cambria Math, so a
  # fallback formula blends with the native ones around it.
  math_fontset: str = "stix"

  #------------------------------------------------------------------------------- Mermaid

  # Mermaid fenced blocks (```mermaid) render through the `mmdc` CLI to a
  # PNG and get embedded as a regular figure. When `mmdc` is missing or
  # the diagram fails to compile, falls back to rendering the source as
  # a regular code block so the document still produces output.
  mermaid_enable: bool = True
  mermaid_theme: str = "default" # mmdc theme: default / dark / forest / neutral
  mermaid_background: str = "transparent" # bg color; transparent matches page
  mermaid_scale: float = 3 # oversampling factor; matches `pdfmarq` for shared cache
  mermaid_cli: str = "mmdc" # CLI binary name / path

  # Colors (rgb 0-1) - mirror `pdfmarq` GitHub-light palette
  body_color: tuple = (0.09, 0.11, 0.13) # #1f2328
  muted_color: tuple = (0.40, 0.44, 0.50) # for blockquote text
  link_color: tuple = (0.03, 0.41, 0.85) # #0969da
  code_inline_color: tuple = (0.09, 0.11, 0.13)
  code_inline_bg: tuple = (0.96, 0.97, 0.98) # subtle grey
  code_block_bg: tuple = (0.96, 0.97, 0.98)
  code_block_border: tuple = (0.82, 0.84, 0.87) # #d0d7de
  quote_border: tuple = (0.82, 0.84, 0.87)
  hr_color: tuple = (0.82, 0.84, 0.87)
  mark_bg: str = "yellow" # named Word highlight

  # Image sizing (mm)
  image_max_h: float = 120 # cap block image height
  inline_image_max_h: float = 5.5 # cap inline image height (~2ex @ 11pt)

  # Local-link handling (see `pdfmarq` for the rationale).
  # `[x](file.md)` with no schema gets resolved to `{link_root}/{link_base}/file.md`
  # when `link_root` is set. Without it, the text is rendered but not linked.
  link_root: str|None = None
  link_base: str = ""

  # GitHub callout titles - override for localization
  callout_label_note:      str = "Note"
  callout_label_tip:       str = "Tip"
  callout_label_important: str = "Important"
  callout_label_warning:   str = "Warning"
  callout_label_caution:   str = "Caution"

  # Callout palette (border + text colors per type)
  callout_colors: dict = field(default_factory=_default_callout_colors)

  #---------------------------------------------------------------------------- Page chrome

  # Footer page numbering. `None` disables the footer entirely.
  page_number_label: str|None = "Page"
  page_number_total: bool = True # `Page X / Y` vs just `Page X`

  # Mini-banner on continuation pages (page 2+). Word header gets the
  # document id (if frontmatter has one) or short title fragment.
  mini_banner_render: bool = True

  #------------------------------------------------------------------- First-page banner

  # Rendered as document content on page 1 when frontmatter is present.
  banner_render: bool = True
  banner_title_size: float = 20 # pt
  banner_meta_size: float = 9 # pt
  banner_status_colors: dict = field(default_factory=_default_status_colors)
  date_format: str = "%Y-%m-%d"
  banner_label_author: str = "Author"
  banner_label_created: str = "Created"
  banner_label_updated: str = "Updated"

  #------------------------------------------------------------------------- Footnotes

  # Footnote section heading. `None` (default) emits just a thin HR above
  # the bibliography - the smaller font signals reference matter. Set to a
  # string (e.g. `"References"`, `"Bibliografia"`) to add an H2 heading
  # above the footnotes. Mirrors `pdfmarq.MarkdownStyle.footnote_label`.
  footnote_label: str|None = None
