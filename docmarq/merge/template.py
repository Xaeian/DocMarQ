# docmarq/merge/template.py

"""
Fill a markdown template with records: Jinja2 in its sandbox, set up for markdown.

Frontmatter values and the body are filled apart, so a value cannot break the YAML around it.
In the body a value prints literally: markdown specials get a backslash, a newline a hard break.
A line holding a tag vanishes when it prints nothing and stays a line of its own otherwise,
so tables, lists and their indents keep their shape whatever the conditions decide.
"""
import os, re, traceback, warnings
from collections.abc import Mapping
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from pathlib import Path
try:
  import yaml
  from jinja2 import (BaseLoader, FileSystemLoader, StrictUndefined, Template,
    TemplateNotFound, TemplateSyntaxError, meta)
  from jinja2.sandbox import ImmutableSandboxedEnvironment, SecurityError
  from markupsafe import Markup
except ModuleNotFoundError as e:
  raise ImportError("Install with: pip install docmarq[md]") from e
from .._warn import caller_level
from ..utils import format_date
from ..md import split_frontmatter

#---------------------------------------------------------------------------------------- Constants

# Specials of markdown and of the preprocessing both libraries run before parsing.
# `(` and `]` stay bare: `\(…\)` and `\[…\]` are LaTeX math, turned into `$` before markdown-it.
_SPECIAL_RE = re.compile(r"([\\`*_\[<>#+\-.!|~^=$&:)])")
_TAG_RE = re.compile(r"\{%.*?%\}|\{#.*?#\}")
_RAW_RE = re.compile(r"\{%[-+]?\s*(raw|endraw)\s*[-+]?%\}")
# Ends every line of the template holding a tag.
# No text carries this private-use character.
_MARK = "\ue000"
_BAD_NAME_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_WINDOWS_NAMES = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
  *(f"LPT{i}" for i in range(1, 10))}
# Field names Jinja reads as something else: literals, operators, globals, `loop` and `self`.
# `n` is the record number in the name pattern.
_RESERVED = {"true", "false", "none", "True", "False", "None", "and", "or", "not", "in", "is",
  "if", "else", "loop", "self", "range", "dict", "cycler", "joiner", "namespace", "n"}
# Decimal point and thousands separator of `number`.
# The no-break space keeps a number on one line.
_SEPARATORS = {
  "pl": (",", "\u00a0"), "cs": (",", "\u00a0"), "sk": (",", "\u00a0"), "fr": (",", "\u00a0"),
  "de": (",", "."), "es": (",", "."), "it": (",", "."),
  "en": (".", ","),
}
_MAX_POWER = 100 # largest exponent of `**`
_MAX_REPEAT = 100_000 # longest text or list `*` may build
_MAX_EXPONENT = 100 # a number past 10^100 is a typo, not an amount
_MAX_PLACES = 100 # places `number` rounds to, either side of the point
_NAME_MAX = 100 # characters of a file name, so a path stays under the 260 of Windows
_SHOWN = 20 # problems a `MergeError` lists
_UNSET = object()

#------------------------------------------------------------------------------------------- Result

class MergeError(ValueError):
  """A template or records that cannot be filled, every problem listed at once."""
  def __init__(self, problems:list[str], head:str="cannot fill the template:"):
    self.problems = list(problems)
    shown = [f"  {p}" for p in self.problems[:_SHOWN]]
    more = len(self.problems) - len(shown)
    super().__init__("\n".join([head, *shown] + ([f"  and {more} more"] if more else [])))

@dataclass
class Merged:
  """One filled document: its file name, its markdown (a list for versions) and its record."""
  name: str
  text: str|list[str]
  data: dict

class _Problem(Exception):
  """A record that cannot be filled, with the place in the template."""

#------------------------------------------------------------------------------------------ Filling

def fill(
  template:str|list[str],
  data:Mapping|None = None,
  *,
  base_dir:str|Path|None = None,
  date_format:str|None = None,
  lang:str|None = None,
) -> str|list[str]:
  """
  Markdown of `template` filled with one record, `data`, over the defaults in its `vars`.

  `data` is a mapping, as a JSON object or a YAML mapping.
  Without it, only `vars` fill the template.
  A list of templates, one per language version, gives a list of their markdown in that order.
  A list of records goes to `merge`.

  Args:
    base_dir: Root of `{% include %}`, the folder images are resolved against too.
      Without it, an include fails.
    date_format: strftime pattern of `date` without an argument, usually `style.date_format`.
    lang: Separators of `number`, as `pl` for `1 234,50`, and `1234.50` without it.

  Raises:
    MergeError: The template or the record cannot be filled, with the line it happened on.
    TypeError: `data` is no mapping, as a list of records.
  """
  if data is not None and not isinstance(data, Mapping):
    raise TypeError("fill takes one record, a mapping of fields; a list of records goes to merge")
  versions = [template] if isinstance(template, str) else list(template)
  body_env, plain_env = _environments(base_dir, date_format, lang)
  parsed = [_parse(text, body_env, plain_env, base_dir) for text in versions]
  record = dict(data or {})
  problems = _field_problems([record])
  if problems: raise MergeError(problems)
  _warn_unused([record], parsed, plain_env, None)
  texts, problems = _fill_versions(parsed, record)
  if problems: raise MergeError(problems)
  return texts[0] if isinstance(template, str) else texts

def merge(
  template:str|list[str],
  rows:list[Mapping],
  *,
  name:str|None = None,
  base_dir:str|Path|None = None,
  date_format:str|None = None,
  lang:str|None = None,
) -> list[Merged]:
  """
  Every record of `rows` filled into `template`, checked all before any comes back.

  `rows` is a list of mappings, one record each, as `csv.DictReader` or `json.load` give them:
  reading a file is the caller's, so CSV is read by whatever reader suits it.
  One record goes to `fill`.
  A list of templates, one per language version, fills each from the same record,
  and `text` of every result is then a list in the same order.
  `name` is a Jinja pattern of the file name, with the fields and `n`, the record number from 1;
  without it, the name is the zero-padded number.
  The other arguments are those of `fill`.

  Raises:
    MergeError: Every record that cannot be filled, with its number and name, or no records.
    TypeError: `rows` is one record, a mapping, which goes to `fill`.
  """
  if isinstance(rows, Mapping):
    raise TypeError("merge takes a list of records; one record, a mapping, goes to fill")
  versions = [template] if isinstance(template, str) else list(template)
  rows = list(rows)
  body_env, plain_env = _environments(base_dir, date_format, lang)
  parsed = [_parse(text, body_env, plain_env, base_dir) for text in versions]
  try:
    pattern = plain_env.from_string(_mark_lines(name)) if name else None
  except TemplateSyntaxError as e:
    raise MergeError([f"name: {e.message}"]) from None
  if not rows: raise MergeError(["no records to fill"])
  problems = _field_problems(rows)
  if problems: raise MergeError(problems)
  _warn_unused(rows, parsed, plain_env, name)
  width = len(str(len(rows)))
  out, seen, failed = [], {}, set()
  for i, row in enumerate(rows, 1):
    number = str(i).zfill(width)
    if all(_empty(v) for v in row.values()): # a spreadsheet row of separators only
      problems.append(f"record {i} is empty")
      failed.add(i)
      continue
    try:
      file_name = _file_name(_unmark(pattern.render({**row, "n": number})) if pattern else number)
    except Exception as e:
      problems.append(f"record {i}: name: {_message(e)}")
      failed.add(i)
      continue
    label = f"record {i} ({file_name})"
    first = seen.setdefault(file_name.casefold(), i)
    if first != i:
      problems.append(f"{label}: the name repeats record {first}")
      failed.add(i)
    texts, failures = _fill_versions(parsed, row)
    if failures:
      problems += [f"{label}: {p}" for p in failures]
      failed.add(i)
    if i not in failed:
      out.append(Merged(file_name, texts[0] if isinstance(template, str) else texts, row))
  if problems:
    raise MergeError(problems, f"{len(failed)} of {len(rows)} records cannot be filled:")
  return out

#----------------------------------------------------------------------------------------- Template

@dataclass
class _Template:
  """A template parsed once: frontmatter to fill, `vars`, the body, and what it reads."""
  front: dict
  defaults: dict
  required: list[str]
  body: Template
  offset: int # lines above the body, for line numbers of the file
  fields: set[str]|None # names it reads, `None` when an include is named from data
  base_dir: str|Path|None

def _parse(text:str, body_env, plain_env, base_dir) -> _Template:
  """The template split and compiled, or `MergeError` when it does not compile."""
  text = text.removeprefix("\ufeff")
  try:
    fm, body = split_frontmatter(text, strict=True)
  except ValueError as e:
    # PyYAML's own message runs over several lines of context.
    # Its line and problem are enough.
    mark = getattr(e.__cause__, "problem_mark", None)
    where = f"line {mark.line + 2}: " if mark else "" # the block starts below the first `---`
    problem = getattr(e.__cause__, "problem", None) or str(e.__cause__ or e)
    hint = " - a placeholder in frontmatter goes in quotes, as `to: \"{{ name }}\"`"
    raise MergeError([f"{where}invalid YAML frontmatter, {problem}"
      + (hint if "{{" in text else "")]) from None
  fm = dict(fm)
  values = fm.pop("vars", None) or {}
  if not isinstance(values, dict): raise MergeError(["`vars` maps names to values"])
  offset = text.count("\n") - body.count("\n")
  problems, sources = [], []
  front = _compile_front(fm, plain_env, problems, sources, "")
  source = _mark_lines(body)
  compiled = fields = None
  try:
    compiled = body_env.from_string(source)
    fields = _fields(body_env, source, plain_env, sources)
  except TemplateSyntaxError as e:
    problems.append(f"{e.name}:{e.lineno}: {e.message}" if e.name
      else f"line {e.lineno + offset}: {e.message}")
  except Exception as e:
    problems.append(_message(e))
  if problems: raise MergeError(problems)
  return _Template(front,
    defaults={k: v for k, v in values.items() if not _empty(v)},
    required=[k for k, v in values.items() if _empty(v)],
    body=compiled, offset=offset, fields=fields, base_dir=base_dir)

def _compile_front(value, env, problems:list, sources:list, path:str):
  """Frontmatter with every text holding a tag compiled, the rest as it is."""
  if isinstance(value, str):
    if "{{" not in value and "{%" not in value: return value
    source = _mark_lines(value)
    sources.append(source)
    try:
      return env.from_string(source)
    except TemplateSyntaxError as e:
      problems.append(f"frontmatter {path}: {e.message}")
      return value
  if isinstance(value, list):
    return [_compile_front(v, env, problems, sources, f"{path}[{i}]") for i, v in enumerate(value)]
  if isinstance(value, dict):
    return {k: _compile_front(v, env, problems, sources, f"{path}.{k}" if path else str(k))
      for k, v in value.items()}
  return value

def _fields(body_env, body:str, plain_env, fronts:list[str]) -> set[str]|None:
  """Names the template reads, with files included under a fixed name, `None` for one from data."""
  fields, seen = set(), set()
  todo = [(body_env, None, body)] + [(plain_env, None, s) for s in fronts]
  while todo:
    env, name, source = todo.pop()
    ast = env.parse(source, name=name)
    fields |= meta.find_undeclared_variables(ast)
    for included in meta.find_referenced_templates(ast):
      if included is None: return None
      if included in seen: continue
      seen.add(included)
      try:
        todo.append((body_env, included, body_env.loader.get_source(body_env, included)[0]))
      except TemplateNotFound:
        return None # rendering names the missing file
  return fields

#-------------------------------------------------------------------------------------------- Lines

def _mark_lines(source:str) -> str:
  """
  `_MARK` at the end of every line holding a tag, outside `{% raw %}` text.
  Jinja then trims nothing, and `_unmark` drops a marked line that printed nothing.
  Whitespace in front of a tag stays, so a list item behind `{% if %}` keeps its indent.
  """
  out, raw = [], False
  for line in source.split("\n"):
    toggles = list(_RAW_RE.finditer(line))
    inside = raw and not toggles
    for m in toggles: raw = m.group(1) == "raw"
    out.append(line if inside or not _TAG_RE.search(line) else line + _MARK)
  return "\n".join(out)

def _unmark(text:str) -> str:
  """The lines a tag left empty dropped, the marks of the others removed."""
  return "\n".join(line.replace(_MARK, "") for line in text.split("\n")
    if _MARK not in line or line.replace(_MARK, "").strip())

#------------------------------------------------------------------------------------------ Records

def _fill_versions(parsed:list[_Template], record:dict) -> tuple[list[str], list[str]]:
  """Markdown of every version filled with `record`, and its problems, by version when many."""
  texts, problems = [], []
  for k, t in enumerate(parsed, 1):
    try:
      texts.append(_fill(t, record))
    except _Problem as e:
      problems.append(f"version {k}: {e}" if len(parsed) > 1 else str(e))
  return texts, problems

def _fill(t:_Template, record:dict) -> str:
  """One template filled with one record, or `_Problem` raised with the place it failed."""
  missing = [k for k in t.required if _empty(record.get(k))]
  if missing: raise _Problem("; ".join(f"`{k}` is required and empty" for k in missing))
  # an empty cell takes the default, as a blank in a spreadsheet means "as usual"
  given = {k: v for k, v in record.items() if not (k in t.defaults and _empty(v))}
  context = {**t.defaults, **given}
  front = _render_front(t.front, context, "")
  try:
    body = _unmark(t.body.render(context))
  except Exception as e:
    raise _Problem(_located(e, t)) from None
  # a body opening with `---` needs a frontmatter block before it, or it is read as one
  if not front and not body.startswith("---"): return body
  return "---\n" + yaml.safe_dump(front, allow_unicode=True, sort_keys=False,
    width=1_000_000) + "---\n\n" + body.lstrip("\n")

def _render_front(value, context:dict, path:str):
  """Frontmatter with its compiled texts filled: YAML is written from it, never parsed."""
  if isinstance(value, Template):
    try:
      return _unmark(value.render(context))
    except Exception as e:
      raise _Problem(f"frontmatter {path}: {_message(e)}") from None
  if isinstance(value, list):
    return [_render_front(v, context, f"{path}[{i}]") for i, v in enumerate(value)]
  if isinstance(value, dict):
    return {k: _render_front(v, context, f"{path}.{k}" if path else str(k))
      for k, v in value.items()}
  return value

def _field_problems(rows:list[dict]) -> list[str]:
  """Records a template cannot read: no mapping, cells with no name, or names Jinja cannot use."""
  names, problems = {}, []
  for i, row in enumerate(rows, 1):
    if not isinstance(row, Mapping): return [f"record {i} is no mapping of fields"]
    # `csv.DictReader` files the cells past the header under `None`
    if None in row: problems.append(f"record {i} has more cells than the header has names")
    for key in row:
      if key is not None: names.setdefault(key, i)
  for key in names:
    if isinstance(key, str) and key.startswith("\ufeff"): # Excel's "CSV UTF-8" read as utf-8
      problems.append(f"field `{key[1:]}` carries a byte-order mark, read the file as utf-8-sig")
    elif not isinstance(key, str) or not key.isidentifier():
      wanted = re.sub(r"\W+", "_", str(key).strip()).strip("_").lower() or "field"
      problems.append(f"field `{key}` is no variable name, rename it, as `{wanted}`")
    elif key in _RESERVED:
      problems.append(f"field `{key}` is a name Jinja keeps for itself, rename it")
  return problems

def _warn_unused(rows:list[dict], parsed:list[_Template], plain_env, name:str|None):
  """Warn about fields no template reads, which are usually typos."""
  used = set()
  for t in parsed:
    if t.fields is None: return
    used |= t.fields
  if name: used |= meta.find_undeclared_variables(plain_env.parse(name))
  unused = sorted({key for row in rows for key in row} - used)
  if unused:
    warnings.warn(f"fields no template reads: {', '.join(unused)}", RuntimeWarning,
      stacklevel=caller_level())

def _file_name(raw:str) -> str:
  """`raw` as a file name Windows takes: no forbidden characters, no dot or space at the end."""
  name = " ".join(_BAD_NAME_RE.sub("", raw).split())[:_NAME_MAX].rstrip(". ")
  if not name: raise ValueError("the name comes out empty")
  # Windows reserves `CON` with any extension, so the mark goes before the first dot
  base, dot, rest = name.partition(".")
  if base.rstrip().upper() in _WINDOWS_NAMES: name = f"{base}_{dot}{rest}"
  return name

def _empty(value) -> bool:
  return value is None or (isinstance(value, str) and not value.strip())

#----------------------------------------------------------------------------------------- Messages

def _located(error:Exception, t:_Template) -> str:
  """`error` after the line it came from: `line 14` of the template, `name:3` of an include."""
  if isinstance(error, RecursionError): return _message(error)
  if isinstance(error, TemplateSyntaxError) and error.name:
    return f"{error.name}:{error.lineno}: {error.message}"
  for frame in reversed(traceback.extract_tb(error.__traceback__)):
    if frame.filename == "<template>":
      return f"line {frame.lineno + t.offset}: {_message(error)}"
    if t.base_dir and not frame.filename.endswith(".py"):
      where = Path(os.path.relpath(frame.filename, t.base_dir)).as_posix()
      return f"{where}:{frame.lineno}: {_message(error)}"
  return _message(error)

def _message(error:Exception) -> str:
  if isinstance(error, RecursionError): return "the template nests too deep or includes itself"
  if isinstance(error, TemplateNotFound): return f"cannot include {error.message}"
  return str(error) or type(error).__name__

#-------------------------------------------------------------------------------------- Environment

class _Sandbox(ImmutableSandboxedEnvironment):
  """
  Jinja's sandbox where data cannot change:
  an `append` to a list in `vars` would reach every record after it.
  `**` and `*` are bounded besides.
  A key of a mapping wins over a method of the same name, so `p.items` reads a column `items`.
  Other ways to spend time or memory remain, as nested loops or `center`,
  so a service filling templates from strangers runs them in a process it can stop.
  """
  intercepted_binops = frozenset(["**", "*"])

  def call_binop(self, context, operator, left, right):
    if operator == "**" and isinstance(right, (int, float)) and abs(right) > _MAX_POWER:
      raise SecurityError(f"`**` with an exponent above {_MAX_POWER}")
    if operator == "*" and _repeat_size(left, right) > _MAX_REPEAT:
      raise SecurityError(f"`*` building more than {_MAX_REPEAT} characters or items")
    return super().call_binop(context, operator, left, right)

  def getattr(self, obj, attribute):
    if isinstance(obj, Mapping) and attribute in obj: return obj[attribute]
    return super().getattr(obj, attribute)

class _Loader(FileSystemLoader):
  """Files under `base_dir`, read as the template is, a byte-order mark dropped."""
  def __init__(self, base_dir:str):
    super().__init__(base_dir, encoding="utf-8-sig")

  def get_source(self, environment, template):
    source, filename, uptodate = super().get_source(environment, template)
    return _mark_lines(source), filename, uptodate

class _NoFiles(BaseLoader):
  """The loader without `base_dir`: an include fails saying what it needs."""
  def get_source(self, environment, template):
    raise TemplateNotFound(template, f"`{template}`: an include needs base_dir")

def _environments(base_dir, date_format:str|None, lang:str|None) -> tuple:
  """`(body, plain)`: the body escapes values for markdown, `plain` fills frontmatter and names."""
  if lang is not None and lang not in _SEPARATORS:
    raise ValueError(f"lang {lang!r} has no number format, use one of: {', '.join(_SEPARATORS)}")
  def make(markdown:bool) -> _Sandbox:
    # autoescape in the body marks what `{% set %}` blocks, macros and `call` catch as markdown,
    # so they are not escaped a second time.
    # Values themselves are escaped by `_literal`.
    env = _Sandbox(undefined=StrictUndefined, keep_trailing_newline=True, autoescape=markdown,
      finalize=_literal if markdown else _plain,
      loader=_Loader(str(base_dir)) if base_dir else _NoFiles())
    env.filters.update({
      "date": lambda value, fmt=None: _date(value, fmt or date_format),
      "number": lambda value, places=2: _number(value, places, lang),
      "md": lambda value: Markup("" if value is None else str(value)),
      "int": _int,
      "float": _float,
    })
    if markdown:
      env.filters.update({"e": _literal, "escape": _literal,
        "forceescape": lambda value: _literal(str(value))})
    # one record gives one document, and markdown has no use for HTML or JSON
    for name in ("random", "tojson", "urlize", "xmlattr"): del env.filters[name]
    del env.globals["lipsum"]
    return env
  return make(True), make(False)

def _literal(value) -> Markup:
  """
  A value as literal markdown: specials escaped, outer whitespace and `None` gone,
  a newline a hard break.
  Markdown already, as from `md` or a `{% set %}` block, passes as it is.
  """
  if hasattr(value, "__html__"): return Markup(value)
  if value is None: return Markup("")
  text = str(value).replace("\r\n", "\n").replace(_MARK, "").strip()
  return Markup(_SPECIAL_RE.sub(r"\\\1", text).replace("\n", "\\\n"))

def _plain(value):
  """A value as plain text, `None` empty: frontmatter is drawn as text, not as markdown."""
  if value is None: return ""
  return value.replace(_MARK, "") if isinstance(value, str) else value

def _repeat_size(left, right) -> int:
  """Length of a text or list repeated by `*`, 0 for a product of numbers."""
  for seq, times in ((left, right), (right, left)):
    if isinstance(seq, (str, list, tuple)) and isinstance(times, int):
      return len(seq) * times
  return 0

#------------------------------------------------------------------------------------------ Filters

def _date(value, fmt:str|None) -> str:
  """`value` in the strftime pattern `fmt`, as `format_date` does, or as written without one."""
  if not fmt: return "" if value is None else str(value)
  return format_date(value, fmt)

def _number(value, places:int, lang:str|None) -> str:
  """`value` rounded half up to `places` decimals, in the separators of `lang`, empty for none."""
  if abs(places) > _MAX_PLACES: raise ValueError(f"`number` takes {_MAX_PLACES} places at most")
  number = _decimal(value)
  if number is None: return ""
  with localcontext() as context:
    context.prec = max(context.prec, number.adjusted() + places + 2) # all digits of a large amount
    number = number.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
  # `copy_abs` ignores the context, where `abs` would round back to its 28 digits
  whole, _, fraction = f"{number.copy_abs():f}".partition(".")
  point, group = _SEPARATORS.get(lang, (".", ""))
  if group: whole = f"{int(whole):,}".replace(",", group)
  return ("-" if number < 0 else "") + whole + (point + fraction if fraction else "")

def _int(value, default=_UNSET, base:int=10):
  """Jinja's `int`, refusing text that is no number where Jinja gives 0, unless `default`."""
  try:
    if base != 10: return int(str(value).strip(), base)
    return int(_given_number(value))
  except ValueError:
    if default is _UNSET: raise
    return default

def _float(value, default=_UNSET):
  """Jinja's `float`, refusing text that is no number where Jinja gives 0, unless `default`."""
  try:
    return float(_given_number(value))
  except ValueError:
    if default is _UNSET: raise
    return default

def _given_number(value) -> Decimal:
  number = _decimal(value)
  if number is None: raise ValueError("an empty value is no number")
  return number

def _decimal(value) -> Decimal|None:
  """
  A number from a value, or from text with `,` or `.` in it.
  `None` for no value or empty text.

  With both marks, the later one is the decimal point and the other groups thousands.
  One mark once is the decimal point, so `1,234` is one and a bit.
  Several times, it groups.
  """
  if value is None: return None
  if isinstance(value, bool): raise ValueError(f"{value} is no number")
  if isinstance(value, (int, float, Decimal)):
    text = str(value)
  else:
    text = re.sub(r"\s", "", str(value))
    if not text: return None
    comma, dot = text.rfind(","), text.rfind(".")
    if comma >= 0 and dot >= 0:
      point = "," if comma > dot else "."
      text = text.replace("." if point == "," else ",", "").replace(point, ".")
    elif comma >= 0 or dot >= 0:
      mark = "," if comma >= 0 else "."
      text = text.replace(mark, "." if text.count(mark) == 1 else "")
  try:
    number = Decimal(text)
  except InvalidOperation:
    raise ValueError(f"{value!r} is no number") from None
  if not number.is_finite() or abs(number.adjusted()) > _MAX_EXPONENT:
    raise ValueError(f"{value!r} is no number")
  return number
