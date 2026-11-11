# docmarq/md/frontmatter.py

"""
YAML frontmatter: split it off a markdown document and read it.

Frontmatter is content only, and nothing here draws it.
The document packages decide what to make of each key.
Values stay text as written: YAML 1.1 types would print `id: 0012` as `10` and `1.10` as `1.1`.
"""
import re, warnings
from functools import cache
from ..warn import caller_level

#---------------------------------------------------------------------------------------- Constants

_FM_RE = re.compile(r"\A\ufeff?---\s*\n(.*?\n)---\s*(?:\n|$)", re.DOTALL) # a BOM may lead
_BOOL_RE = re.compile(r"^(?:true|True|TRUE|false|False|FALSE)$")

#--------------------------------------------------------------------------------------------- Read

def read_yaml(text:str):
  """
  YAML with values kept as written text, but `None` (empty, `null`, `~`), booleans and dates.

  PyYAML follows YAML 1.1, which changes data without a word:
  `02210112345` turns into an octal number, `12:30` into 750, `NO` into `False`, `1.10` into `1.1`.
  Aliases are refused, since a few lines of them can stand for gigabytes of text.
  Raises `ModuleNotFoundError` without PyYAML and `yaml.YAMLError` on text that does not parse.
  """
  import yaml
  return yaml.load(text, Loader=_text_loader())

@cache
def _text_loader():
  """`yaml.SafeLoader` typing nothing but null, dates and `true`/`false`, and refusing aliases."""
  import yaml
  class TextLoader(yaml.SafeLoader):
    """Scalars stay strings unless null, a date or a YAML 1.2 boolean, and aliases fail."""
    def compose_node(self, parent, index):
      if self.check_event(yaml.AliasEvent):
        raise yaml.composer.ComposerError(None, None, "aliases are not allowed",
          self.peek_event().start_mark)
      return super().compose_node(parent, index)
  keep = ("tag:yaml.org,2002:null", "tag:yaml.org,2002:timestamp")
  TextLoader.yaml_implicit_resolvers = {first: [(tag, rx) for tag, rx in rules if tag in keep]
    for first, rules in yaml.SafeLoader.yaml_implicit_resolvers.items()}
  TextLoader.add_implicit_resolver("tag:yaml.org,2002:bool", _BOOL_RE, list("tTfF"))
  return TextLoader

def peek_frontmatter(md_text:str) -> dict|None:
  """
  Lightweight YAML frontmatter parse without markdown rendering.
  Returns the parsed dict or `None` if no frontmatter / parse failure.
  Public so a caller can read a document's content keys and decide how to render it.
  The library itself never derives form from them.
  """
  m = _FM_RE.match(md_text)
  if not m: return None
  try:
    data = read_yaml(m.group(1))
    return data if isinstance(data, dict) else None
  except Exception: return None

def split_frontmatter(md_text:str, warn:bool=True, strict:bool=False) -> tuple[dict, str]:
  """
  `(data, body)`: the frontmatter mapping, `{}` when there is none, and the markdown after it.

  Only a YAML mapping counts as frontmatter.
  A block that parses to a scalar or a list is a document opening with a horizontal rule,
  so the text is handed back whole - everything down to the next `---` is ordinary content.
  A block that looks like frontmatter but does not parse is dropped with a warning,
  since it was meant as metadata.
  `warn=False` drops it quietly, for text read a second time.
  `strict=True` raises instead, for a template that must not lose its header:
  `ValueError` for a block that does not parse, `ModuleNotFoundError` without PyYAML.
  """
  m = _FM_RE.match(md_text)
  if not m: return {}, md_text
  try:
    data = read_yaml(m.group(1))
  except ModuleNotFoundError:
    if strict: raise
    warnings.warn("YAML frontmatter needs PyYAML (pip install docmarq[md]); its keys are skipped",
      RuntimeWarning, stacklevel=2)
    return {}, md_text[m.end():]
  except Exception as e:
    if strict: raise ValueError(f"invalid YAML frontmatter: {e}") from e
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
  if fm.get("title") is not None: out["title"] = _plain(fm["title"])
  if fm.get("author") is not None: out["author"] = _plain(byline(fm["author"]))
  if fm.get("subject") is not None: out["subject"] = _plain(fm["subject"])
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

#------------------------------------------------------------------------------------------- People

# An ORCID and an e-mail, each known by its shape, its value the first group.
# A phone number has no shape of its own: dates and ranges of years would read as one.
_DETAILS = (
  ("orcid", re.compile(r"(?:(?:https?://)?orcid\.org/)?(\d{4}-\d{4}-\d{4}-\d{3}[\dX])")),
  ("email", re.compile(r"([\w.%+'-]+@[\w-]+(?:\.[\w-]+)+)")),
)
# a word that only names the detail beside it, as `ORCID:` or `e-mail`
_LABEL = re.compile(r"\b(?:orcid(?:\s+id)?|(?:e-?)?mail)\b[.:]*", re.IGNORECASE)

def people(value) -> list[dict]:
  """
  Persons an `author`-like frontmatter value names, each a mapping with its `name` first.

  The value is one mapping, a list of names and mappings, or names in a comma-separated string.
  In a string an ORCID or an e-mail, known by its shape, joins the name it follows:
  `Anna Nowak, 0000-0002-1825-0097, a.nowak@example.com`.
  A mapping keeps its other keys as written, as `orcid` or `aff` for a package that sets them,
  its empty ones dropped. A person with no name is left out with a warning.
  """
  if not value: return []
  out = []
  for item in value if isinstance(value, (list, tuple)) else [value]:
    if not isinstance(item, dict):
      if item is not None: out += _named(str(item))
      continue
    person = dict(item)
    name = str(person.pop("name", None) or "").strip()
    if name:
      out.append({"name": name, **{k: v for k, v in person.items() if v not in (None, "")}})
    else:
      warnings.warn(f"an author with no name, left out: {item}", RuntimeWarning,
        stacklevel=caller_level())
  return out

def _named(text:str) -> list[dict]:
  """
  People of a comma-separated string.
  A detail joins the name beside it, or the name before it when it stands alone,
  so a comma left out between a name and its detail loses no one.
  A label beside a detail, as `ORCID:`, or a stray letter is no name.
  A person keeps the first detail of a kind, and a second one warns.
  """
  out = []
  for part in (piece.strip() for piece in text.split(",")):
    details, rest = _details(part)
    if not details:
      if part: out.append({"name": part})
      continue
    name = " ".join(_LABEL.sub(" ", rest).split()).strip("()[]<>:; ")
    if re.search(r"[^\W\d_]{2}", name): out.append({"name": name})
    elif not out:
      warnings.warn(f"`{part}` follows no name, left out", RuntimeWarning,
        stacklevel=caller_level())
      continue
    for key, value in details:
      if key in out[-1]:
        warnings.warn(f"a second {key} of {out[-1]['name']}, `{value}`, left out", RuntimeWarning,
          stacklevel=caller_level())
        continue
      out[-1][key] = value
  return out

def _details(part:str) -> tuple[list[tuple[str, str]], str]:
  """`(key, value)` of each ORCID and e-mail in `part`, and the rest of it."""
  found = []
  for key, shape in _DETAILS:
    found += [(key, match[1]) for match in shape.finditer(part)]
    part = shape.sub(" ", part)
  return found, part

def byline(value, bind:bool=False) -> str:
  """
  The names of `people(value)` on one line, comma-separated, as banners and metadata show them.
  A string of names alone stays as written, so a name holding a comma keeps it.
  `bind` ties the words of each name by no-break spaces where there are several,
  so a line breaks only between names, while a single name breaks as any text does.
  """
  if isinstance(value, str) and not _details(value)[0]: line = value
  else: line = ", ".join(person["name"] for person in people(value))
  if not bind or ", " not in line: return line
  return re.sub(r"(?<!,) ", "\u00a0", " ".join(line.split()))
