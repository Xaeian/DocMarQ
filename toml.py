# toml.py

"""
Auto-generate pyproject.toml from package source analysis.
Scans Python files, detects modules, reads __extras__ declarations,
auto-discovers third-party dependencies from imports, and detects
non-Python files for package-data.

Each module/subpackage declares its own extras via __extras__:
  Tuple form: __extras__ = ("group", ["pkg1", "pkg2"])
  Dict form:  __extras__ = {"group": ["pkg1"], "group-async": ["pkg2"]}

Example:
  >>> from toml import generate
  >>> generate("xaeian")

CLI:
  py toml.py xaeian
  py toml.py xaeian -o pyproject.toml
"""

import ast, sys
from xaeian import FILE, DIR, PATH, Print, Color as c

p = Print()

#--------------------------------------------------------------------------- Import-to-PyPI map

IMPORT_MAP = {
  "PIL": "Pillow",
  "cv2": "opencv-python",
  "yaml": "PyYAML",
  "serial": "pyserial",
  "usb": "pyusb",
  "gi": "PyGObject",
  "bs4": "beautifulsoup4",
  "attr": "attrs",
  "sklearn": "scikit-learn",
  "skimage": "scikit-image",
  "wx": "wxPython",
  "Crypto": "pycryptodome",
  "nacl": "PyNaCl",
  "google.protobuf": "protobuf",
  "jose": "python-jose",
  "dotenv": "python-dotenv",
  "magic": "python-magic",
  "dateutil": "python-dateutil",
  "webview": "pywebview",
  "docx": "python-docx",
  "pptx": "python-pptx",
  "lxml": "lxml",
  "fitz": "PyMuPDF",
}

#------------------------------------------------------------------------------------ Internals

def _parse_extras(node) -> dict[str, list[str]]:
  """Extract extras from AST node value (tuple or dict)."""
  if isinstance(node, ast.Tuple) and len(node.elts) == 2:
    name_node, list_node = node.elts
    if isinstance(name_node, ast.Constant) and isinstance(list_node, ast.List):
      pkgs = [e.value for e in list_node.elts if isinstance(e, ast.Constant)]
      return {str(name_node.value): pkgs}
  elif isinstance(node, ast.Dict):
    result = {}
    for k, v in zip(node.keys, node.values):
      if isinstance(k, ast.Constant) and isinstance(v, ast.List):
        pkgs = [e.value for e in v.elts if isinstance(e, ast.Constant)]
        result[str(k.value)] = pkgs
    return result
  return {}

def _scan_extras_from_file(path:str) -> dict[str, list[str]]:
  """Parse `__extras__` from a Python file."""
  if not PATH.is_file(path): return {}
  try:
    tree = ast.parse(FILE.load(path))
  except Exception:
    return {}
  for node in ast.walk(tree):
    if isinstance(node, ast.Assign):
      for t in node.targets:
        if isinstance(t, ast.Name) and t.id == "__extras__":
          return _parse_extras(node.value)
  return {}

def _top_level(name:str) -> str:
  """Extract top-level module from dotted import name."""
  return name.split(".")[0]

#------------------------------------------------------------------------------------- Analysis

def scan_package(pkg_dir:str) -> tuple[set[str], set[str]]:
  """Return (modules, subpackages) present in package."""
  modules = set()
  subpackages = set()
  for name in DIR.file_list(pkg_dir, exts=[".py"], local=True):
    if "/" in name or name.startswith("__"): continue
    modules.add(name.removesuffix(".py"))
  for name in DIR.folder_list(pkg_dir, basename=True):
    if PATH.is_file(PATH.join(pkg_dir, name, "__init__.py")):
      subpackages.add(name)
  return modules, subpackages

def scan_imports(pkg_dir:str, pkg_name:str) -> set[str]:
  """Scan all `.py` files in package for third-party imports.

  Args:
    pkg_dir: Package directory path.
    pkg_name: Package name (to exclude self-imports).

  Returns:
    Set of PyPI package names detected from imports.
  """
  stdlib = sys.stdlib_module_names
  internal = {pkg_name}
  py_files = DIR.file_list(pkg_dir, exts=[".py"])
  raw_imports: set[str] = set()
  for fpath in py_files:
    try:
      tree = ast.parse(FILE.load(fpath))
    except Exception:
      continue
    for node in ast.walk(tree):
      if isinstance(node, ast.Import):
        for alias in node.names:
          raw_imports.add(alias.name)
      elif isinstance(node, ast.ImportFrom):
        if node.module and node.level == 0:
          raw_imports.add(node.module)
  third_party: set[str] = set()
  for name in raw_imports:
    top = _top_level(name)
    if top in stdlib: continue
    if top in internal: continue
    if top.startswith("_"): continue
    pypi = IMPORT_MAP.get(name) or IMPORT_MAP.get(top) or top
    third_party.add(pypi)
  return third_party

def scan_package_data(pkg_dir:str) -> list[str]:
  """Detect non-Python files that need `package-data` declaration.

  Returns:
    List of glob patterns like `"files/**"`, `"*.cfg"`.
  """
  skip_exts = {".py", ".pyc", ".pyo", ".md"}
  all_files = DIR.file_list(pkg_dir, local=True, blacklist=["__pycache__"])
  top_dirs: set[str] = set()
  root_exts: set[str] = set()
  for f in all_files:
    if f.startswith("__"): continue
    ext = PATH.ext(f)
    if ext in skip_exts: continue
    parts = f.split("/")
    if len(parts) > 1:
      top_dirs.add(parts[0])
    else:
      if ext:
        root_exts.add(f"*{ext}")
      else:
        root_exts.add(f)
  patterns = sorted(root_exts)
  for d in sorted(top_dirs):
    patterns.append(f"{d}/**")
  return patterns

def build_extras(
  pkg_dir:str, modules:set[str], subpackages:set[str],
) -> dict[str, list[str]]:
  """Build extras dict by scanning `__extras__` in modules and subpackages."""
  extras: dict[str, set[str]] = {}
  for mod in modules:
    found = _scan_extras_from_file(PATH.join(pkg_dir, f"{mod}.py"))
    for name, pkgs in found.items():
      extras.setdefault(name, set()).update(pkgs)
  for subpkg in subpackages:
    for path in DIR.file_list(PATH.join(pkg_dir, subpkg), exts=[".py"]):
      found = _scan_extras_from_file(path)
      for name, pkgs in found.items():
        extras.setdefault(name, set()).update(pkgs)
  if extras:
    all_deps = set()
    for deps in extras.values():
      all_deps.update(deps)
    extras["all"] = all_deps
  return {k: sorted(v) for k, v in extras.items()}

def get_meta(pkg_dir:str) -> dict:
  """Extract metadata from `__init__.py`."""
  meta = {
    "version": "0.0.0", "repo": "", "python": ">=3.10",
    "description": "", "author": "", "keywords": [],
    "dependencies": [], "scripts": {},
  }
  init = PATH.join(pkg_dir, "__init__.py")
  if not PATH.is_file(init): return meta
  STR_FIELDS = {
    "__version__": "version", "__repo__": "repo",
    "__python__": "python", "__description__": "description",
    "__author__": "author",
  }
  LIST_FIELDS = {"__keywords__": "keywords", "__dependencies__": "dependencies"}
  try:
    tree = ast.parse(FILE.load(init))
  except Exception:
    return meta
  for node in ast.walk(tree):
    if not isinstance(node, ast.Assign): continue
    for t in node.targets:
      if not isinstance(t, ast.Name): continue
      if isinstance(node.value, ast.Constant) and t.id in STR_FIELDS:
        meta[STR_FIELDS[t.id]] = str(node.value.value)
      elif isinstance(node.value, ast.List) and t.id in LIST_FIELDS:
        meta[LIST_FIELDS[t.id]] = [
          e.value for e in node.value.elts if isinstance(e, ast.Constant)
        ]
      elif isinstance(node.value, ast.Dict) and t.id == "__scripts__":
        meta["scripts"] = {
          k.value: v.value
          for k, v in zip(node.value.keys, node.value.values)
          if isinstance(k, ast.Constant) and isinstance(v, ast.Constant)
        }
  return meta

#------------------------------------------------------------------------------------- Generate

def generate_toml(
  pkg_name:str, meta:dict,
  extras:dict[str, list[str]],
  package_data:list[str]|None=None,
) -> str:
  """Generate pyproject.toml content."""
  lines = [
    '[build-system]',
    'requires = ["setuptools>=61.0", "wheel"]',
    'build-backend = "setuptools.build_meta"',
    '',
    '[project]',
    f'name = "{pkg_name}"',
    f'version = "{meta["version"]}"',
    f'description = "{meta["description"]}"',
    'readme = "readme.md"',
    'license = {text = "MIT"}',
    f'requires-python = "{meta["python"]}"',
  ]
  if meta["author"]:
    lines.append(f'authors = [{{name = "{meta["author"]}"}}]')
  if meta["keywords"]:
    kw_str = ", ".join(f'"{k}"' for k in meta["keywords"])
    lines.append(f'keywords = [{kw_str}]')
  deps_str = ", ".join(f'"{d}"' for d in meta["dependencies"])
  lines.append(f'dependencies = [{deps_str}]')
  lines.append('')
  if extras:
    lines.append('[project.optional-dependencies]')
    for name in sorted(extras.keys(), key=lambda x: (x == "all", x)):
      dep_str = ", ".join(f'"{d}"' for d in extras[name])
      lines.append(f'{name} = [{dep_str}]')
    lines.append('')
  if meta.get("scripts"):
    lines.append('[project.scripts]')
    for cmd, entry in meta["scripts"].items():
      lines.append(f'{cmd} = "{entry}"')
    lines.append('')
  lines.append('[project.urls]')
  if meta["repo"]:
    lines.append(f'Repository = "https://github.com/{meta["repo"]}"')
  else:
    lines.append(f'Repository = "https://github.com/.../{pkg_name}"')
  lines.append('')
  lines.append('[tool.setuptools.packages.find]')
  lines.append(f'include = ["{pkg_name}*"]')
  lines.append('')
  if package_data:
    lines.append('[tool.setuptools.package-data]')
    pat_str = ", ".join(f'"{pat}"' for pat in package_data)
    lines.append(f'{pkg_name} = [{pat_str}]')
    lines.append('')
  lines.append('[tool.pytest.ini_options]')
  lines.append('testpaths = ["tests"]')
  lines.append('python_functions = ["*"]')
  lines.append('')
  return "\n".join(lines)

#-------------------------------------------------------------------------------------- Logging

def _log_summary(
  pkg_name:str, meta:dict,
  modules:set[str], subpackages:set[str],
  extras:dict[str, list[str]],
  package_data:list[str],
):
  p.inf(f"Package: {c.TURQUS}{pkg_name}{c.END} {meta['version']}")
  if meta["repo"]:
    p.gap(f"https://github.com/{c.SKY}{meta['repo']}{c.END}")
  p.inf(f"Modules: {c.GREY}{', '.join(sorted(modules))}{c.END}")
  if subpackages:
    p.inf(f"Subpackages: {c.GREY}{', '.join(sorted(subpackages))}{c.END}")
  if meta["dependencies"]:
    p.inf(f"Dependencies: {c.GREY}{', '.join(meta['dependencies'])}{c.END}")
  if extras:
    for name, deps in sorted(extras.items(), key=lambda x: (x[0] == "all", x[0])):
      p.item(f"[{c.CREAM}{name}{c.END}]: {c.GREY}{', '.join(deps)}{c.END}")
  if package_data:
    p.inf(f"Package data: {c.GREY}{', '.join(package_data)}{c.END}")
  if meta.get("scripts"):
    for cmd, entry in meta["scripts"].items():
      p.item(f"Script: {c.TURQUS}{cmd}{c.END} → {c.GREY}{entry}{c.END}")

#--------------------------------------------------------------------------------------- Public

def generate(package:str, output:str|None=None, auto_deps:bool=False):
  """Generate pyproject.toml for given package directory.

  Args:
    package: Package directory path.
    output: Output file path (default: parent/pyproject.toml).
    auto_deps: Scan imports for third-party dependencies.
  """
  pkg_dir = PATH.resolve(package)
  if not PATH.is_dir(pkg_dir):
    p.err(f"{c.ORANGE}{pkg_dir}{c.END} is not a directory")
    sys.exit(1)
  pkg_name = PATH.basename(pkg_dir)
  meta = get_meta(pkg_dir)
  modules, subpackages = scan_package(pkg_dir)
  extras = build_extras(pkg_dir, modules, subpackages)
  package_data = scan_package_data(pkg_dir)
  if auto_deps:
    scanned = scan_imports(pkg_dir, pkg_name)
    extras_all = set()
    for deps in extras.values():
      extras_all.update(deps)
    scanned -= extras_all
    declared = set(meta["dependencies"])
    new_deps = scanned - declared
    if new_deps:
      p.wrn(f"Auto-detected: {c.TURQUS}{', '.join(sorted(new_deps))}{c.END}")
    meta["dependencies"] = sorted(declared | scanned)
  _log_summary(pkg_name, meta, modules, subpackages, extras, package_data)
  toml = generate_toml(pkg_name, meta, extras, package_data)
  out = output or PATH.join(PATH.dirname(pkg_dir), "pyproject.toml")
  FILE.save(out, toml)
  p.ok(f"Generated {c.GREY}{PATH.dirname(out)}/{c.END}{c.ORANGE}{PATH.basename(out)}{c.END}")

#------------------------------------------------------------------------------------------ CLI

EXAMPLES = """
examples:
  py toml.py xaeian              Generate from package dir
  py toml.py xaeian -o out.toml  Custom output path
  py toml.py xaeian --auto-deps  Auto-detect dependencies from imports
"""

if __name__ == "__main__":
  from xaeian.cli._args import _make_parser, _add_help
  parser = _make_parser(
    f"Generate {c.ORANGE}pyproject.toml{c.END} from package source", EXAMPLES,
  )
  parser.add_argument("package", metavar="PACKAGE", help="Package directory to scan")
  parser.add_argument("-o", "--output", default=None, metavar="PATH",
    help="Output file (default: parent/pyproject.toml)")
  parser.add_argument("-a", "--auto-deps", action="store_true",
    help="Auto-detect third-party dependencies from imports")
  _add_help(parser)
  args = parser.parse_args()
  generate(args.package, args.output, args.auto_deps)