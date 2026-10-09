# docmarq/md/frontmatter.py

"""
YAML frontmatter: split it off a markdown document and read it.

Frontmatter is content only, and nothing here draws it.
The document packages decide what to make of each key.
"""
import re, warnings
from .._warn import caller_level

#---------------------------------------------------------------------------------------- Constants

_FM_RE = re.compile(r"\A---\s*\n(.*?\n)---\s*(?:\n|$)", re.DOTALL)

#--------------------------------------------------------------------------------------------- Read

def peek_frontmatter(md_text:str) -> dict|None:
  """Lightweight YAML frontmatter parse without markdown rendering.
  Returns the parsed dict or `None` if no frontmatter / parse failure.
  Public so a caller can read a document's content keys and decide how to render it.
  The library itself never derives form from them.
  """
  m = _FM_RE.match(md_text)
  if not m: return None
  try:
    import yaml
    data = yaml.safe_load(m.group(1))
    return data if isinstance(data, dict) else None
  except Exception:
    return None

def split_frontmatter(md_text:str, warn:bool=True) -> tuple[dict, str]:
  """
  `(data, body)`: the frontmatter mapping, `{}` when there is none, and the markdown after it.

  Only a YAML mapping counts as frontmatter.
  A block that parses to a scalar or a list is a document opening with a horizontal rule,
  so the text is handed back whole - everything down to the next `---` is ordinary content.
  A block that looks like frontmatter but does not parse is dropped with a warning,
  since it was meant as metadata; `warn=False` drops it quietly, for text read a second time.
  """
  m = _FM_RE.match(md_text)
  if not m: return {}, md_text
  try:
    import yaml
    data = yaml.safe_load(m.group(1))
  except ModuleNotFoundError:
    warnings.warn("YAML frontmatter needs PyYAML (pip install docmarq[md]); its keys are skipped",
      RuntimeWarning, stacklevel=2)
    return {}, md_text[m.end():]
  except Exception as e:
    if warn:
      warnings.warn(f"invalid YAML frontmatter, block ignored: {e}", RuntimeWarning,
        stacklevel=caller_level())
    return {}, md_text[m.end():]
  if not isinstance(data, dict): return {}, md_text
  return data, md_text[m.end():]

def docx_metadata(fm:dict, overrides:dict|None=None) -> dict:
  """`DOCX.metadata()` arguments from frontmatter, `overrides` winning key by key."""
  return {**metadata_from_frontmatter(fm), **(overrides or {})}

def metadata_from_frontmatter(fm:dict) -> dict:
  """Map YAML keys to `DOCX.metadata()` kwargs. Keys absent in `fm` are skipped."""
  out: dict = {}
  for yaml_key, meta_key in (("title", "title"), ("author", "author"), ("subject", "subject")):
    val = fm.get(yaml_key)
    if val is not None:
      out[meta_key] = _plain(val)
  kw = fm.get("keywords")
  if kw is not None:
    if isinstance(kw, (list, tuple)):
      out["keywords"] = ", ".join(str(k) for k in kw)
    else:
      out["keywords"] = str(kw)
  return out

def _plain(value) -> str:
  """A frontmatter value on one line: the entries of a list, or the lines of a block, by spaces."""
  items = value if isinstance(value, (list, tuple)) else [value]
  return " ".join(" ".join(str(item).split()) for item in items)
