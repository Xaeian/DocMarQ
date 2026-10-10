# docmarq/paper/style.py

"""Style of a documentation DOCX: the markdown style plus banner, footer and signing lines."""
from dataclasses import dataclass, field
from ..md import MarkdownStyle
from ..parts import sign_labels, default_sign_labels

#----------------------------------------------------------------------------------- Status palette

def _default_status_colors() -> dict:
  """Status badge colors `(bg, text)` keyed by lowercase status name."""
  return {
    "draft": ((0.85, 0.87, 0.92), (0.30, 0.34, 0.45)),
    "review": ((1.00, 0.95, 0.78), (0.62, 0.40, 0.05)),
    "approved": ((0.86, 0.96, 0.87), (0.10, 0.45, 0.18)),
    "deprecated": ((1.00, 0.88, 0.85), (0.74, 0.20, 0.15)),
    "archived": ((0.92, 0.87, 0.96), (0.45, 0.25, 0.55)),
  }

# Badge colors `(bg, text)` of a status the palette does not name.
UNKNOWN_STATUS = ((0.93, 0.93, 0.95), (0.40, 0.44, 0.50))

#--------------------------------------------------------------------------------------- PaperStyle

@dataclass
class PaperStyle(MarkdownStyle):
  """`MarkdownStyle` plus the frontmatter banner, compact banner, page number and signatures."""
  # Footer page numbering. `None` disables the footer entirely.
  page_number_label: str|None = "Page"
  page_number_total: bool = True # `Page X/Y` vs just `Page X`

  # Mini-banner on continuation pages (page 2+).
  # Word header gets the document id (if frontmatter has one) or short title fragment.
  banner_compact: bool = True

  # Rendered as document content on page 1 when frontmatter is present.
  banner: bool = True
  # Drop a leading `# title` duplicating the frontmatter title.
  # Only applies when the banner rendered that title, otherwise the document loses its head.
  skip_dup_title: bool = True
  banner_title_size: float = 20 # pt
  banner_meta_size: float = 9 # pt
  banner_status_colors: dict = field(default_factory=_default_status_colors)
  date_format: str = "%Y-%m-%d"
  banner_label_author: str = "Author"
  banner_label_created: str = "Created"
  banner_label_updated: str = "Updated"

  # Signing lines at the very end, independent of `banner`.
  # `True` or a `sign_labels` scenario draws its labels; a list draws custom ones verbatim.
  sign: bool|str|list = False
  sign_labels: dict = field(default_factory=default_sign_labels)
  sign_gap: float = 25 # mm - blank space above the line, for the handwriting
  sign_w: float = 70 # mm - width of one line
  sign_size: float = 9 # pt - label

  def sign_lines(self) -> list[str]:
    """Labels to draw: a scenario from `sign_labels`, or a list verbatim."""
    return sign_labels(self.sign, self.sign_labels)
