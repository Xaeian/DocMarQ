# docmarq/legal/style.py

"""Style of a legal document: the markdown style set for letters and contracts, plus their page."""
from dataclasses import dataclass, field
from ..md.style import MarkdownStyle
from ..parts import sign_labels, default_sign_labels

#--------------------------------------------------------------------------------------- LegalStyle

@dataclass
class LegalStyle(MarkdownStyle):
  """`MarkdownStyle` set for letters, statements and contracts, plus header, columns and end."""
  # Engine defaults a legal page sets its own way:
  # justified text, `##` and lower at body size (`None`), no rules under headings, tighter spacing.
  justify: bool = True
  h1_size: float|None = 13
  h2_size: float|None = None
  h3_size: float|None = None
  h4_size: float|None = None
  h5_size: float|None = None
  h6_size: float|None = None
  h1_underline: bool = False
  h2_underline: bool = False
  head_gap_top: float|None = 2 # mm
  head_gap_bot: float|None = 1 # mm
  para_gap: float = 2 # mm

  # `2` sets the text in two columns from the first `##` on, the preamble above stays full width.
  columns: int = 1
  column_gap: float = 8 # mm
  # The preamble is the text before the first `##`, naming the parties of a contract.
  preamble_size: float|None = None # pt, `None` follows `body_size`
  preamble_para_gap: float|None = None # mm between its paragraphs, `None` follows `para_gap`
  preamble_gap: float = 4 # mm between the preamble and the clauses

  # Letter header, from frontmatter `title`, `place`, `date`, `from`, `to`, `ref` and `logo`.
  # The page around the text: header, title, preamble, signing lines, attachments, page number.
  # `header_font` sets it in a face of its own;
  # `None` keeps `font_body`, and `font_head` for the title.
  header_font: str|None = None
  title_size: float = 16 # pt
  title_caps: bool = False # lowercase letters as Word's small capitals
  title_gap: float = 8 # mm below the title, and above it under header rows
  header_size: float = 10 # pt - place and date, sender, recipient, attachments
  to_align: str = "R" # L | R: the recipient flush with the left or the right margin
  ref_label: str = "Ref." # before `ref`: a case or reference number, bold, under the recipient
  logo_align: str = "L" # L | C | R; at a side, the title and preamble may run beside it
  logo_max_h: float = 20 # mm
  logo_max_w: float = 50 # mm
  # strftime syntax, applied to a YAML date or an ISO `YYYY-MM-DD` text.
  date_format: str = "%Y-%m-%d"

  # End of the document: signing lines, then the attachments listed in frontmatter.
  # `sign` takes `True`, a `sign_labels` scenario, or a list of labels drawn verbatim.
  sign: bool|str|list = False
  sign_labels: dict = field(default_factory=default_sign_labels)
  sign_gap: float = 20 # mm - blank space above the line, for the handwriting
  sign_w: float = 70 # mm - width of one line
  sign_size: float = 9 # pt - label
  # Signing lines and attachments together at the bottom margin of the page the text ends on.
  sign_bottom: bool = False
  attachments_label: str = "Attachments"

  # Page number prefix; `None` disables the footer. A single page never carries one.
  page_number_label: str|None = "Page"
  # `Page X / Y` vs just `Page X`.
  page_number_total: bool = True

  @property
  def page_font(self) -> str:
    """Face of the page around the text: `header_font`, or `font_body` when unset."""
    return self.header_font or self.font_body

  def sign_lines(self) -> list[str]:
    """Labels to draw: a scenario from `sign_labels`, or a list verbatim."""
    return sign_labels(self.sign, self.sign_labels)
