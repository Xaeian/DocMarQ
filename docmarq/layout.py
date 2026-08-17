# docmarq/layout.py

"""Page geometry - section configuration mirroring Word's section model."""
from dataclasses import dataclass

#------------------------------------------------------------------------------------- PageGeometry

@dataclass
class PageGeometry:
  """Page dimensions and margins in mm. Maps 1:1 to a Word section.

  Word stores margins per-section (top/right/bot/left), plus header/footer
  distances from page edge. `header_dist` / `footer_dist` default to half
  the corresponding margin so chrome sits midway in the margin strip.
  """
  width: float
  height: float
  margin_top: float = 20
  margin_right: float = 20
  margin_bot: float = 20
  margin_left: float = 20
  header_dist: float|None = None # mm from top edge to header content
  footer_dist: float|None = None # mm from bottom edge to footer content
  gutter: float = 0 # binding margin

  @property
  def content_width(self) -> float:
    return self.width - self.margin_left - self.margin_right

  @property
  def content_height(self) -> float:
    return self.height - self.margin_top - self.margin_bot

  def effective_header_dist(self) -> float:
    """Distance from page top to header content in mm. Defaults to half
    the top margin when `header_dist` is unset (chrome midline convention)."""
    return self.header_dist if self.header_dist is not None else self.margin_top / 2

  def effective_footer_dist(self) -> float:
    """Distance from page bottom to footer content in mm. Defaults to half
    the bottom margin when `footer_dist` is unset."""
    return self.footer_dist if self.footer_dist is not None else self.margin_bot / 2
