# docmarq/structure.py

"""Document structure - metadata, bookmarks, hyperlinks."""
from dataclasses import dataclass

#----------------------------------------------------------------------------------------- Metadata

@dataclass
class Metadata:
  """Core docx properties (`core.xml`)."""
  title: str|None = None
  author: str|None = None
  subject: str|None = None
  keywords: str|None = None
  comments: str|None = None
  category: str|None = None

  def apply(self, document):
    """Write to `python-docx` `core_properties`."""
    cp = document.core_properties
    if self.title: cp.title = self.title
    if self.author: cp.author = self.author
    if self.subject: cp.subject = self.subject
    if self.keywords: cp.keywords = self.keywords
    if self.comments: cp.comments = self.comments
    if self.category: cp.category = self.category

#----------------------------------------------------------------------------------------- Bookmark

@dataclass
class Bookmark:
  """Internal bookmark anchor. In OOXML this is a `<w:bookmarkStart/>` +
  `<w:bookmarkEnd/>` pair around a range of content, referenced by id.
  """
  name: str
  id: int
