# tests/test_layers.py

"""
Layer boundaries: core ← md ← paper, legal and merge, the three strangers to each other.
A layer reaches the ones below through their public names only.
"""

import ast, importlib.util
from pathlib import Path
import pytest
import docmarq

PACKAGE = Path(docmarq.__file__).resolve().parent # under test: the wheel in CI, the source here
SUBPACKAGES = ("md", "paper", "legal", "merge")

# layer → layers it must never import
FORBIDDEN = {
  "core": {"md", "paper", "legal", "merge"},
  "md": {"paper", "legal", "merge"},
  "paper": {"legal", "merge"},
  "legal": {"paper", "merge"},
  "merge": {"paper", "legal"},
}

#------------------------------------------------------------------------------------------ Helpers

@pytest.fixture
def layer_of():
  """Layer of a dotted module name inside `docmarq`."""
  def of(module:str) -> str:
    parts = module.split(".")
    return parts[1] if len(parts) > 1 and parts[1] in SUBPACKAGES else "core"
  return of

@pytest.fixture
def module_of():
  """Dotted module name of a file under `PACKAGE`, `__init__` kept as a module of its own."""
  def of(path:Path) -> str:
    return ".".join(path.relative_to(PACKAGE.parent).with_suffix("").parts)
  return of

@pytest.fixture
def imported(layer_of, module_of):
  """
  Layers of `docmarq` imported by the file at `path`, lazy imports included.
  `package` places a file outside the package, for relative imports.
  """
  def scan(path:Path, package:str|None=None) -> set[str]:
    package = package or module_of(path).rsplit(".", 1)[0]
    names = []
    for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
      if isinstance(node, ast.Import):
        names += [alias.name for alias in node.names]
      elif isinstance(node, ast.ImportFrom):
        base = node.module or ""
        if node.level:
          base = importlib.util.resolve_name("." * node.level + base, package)
        # one name per alias: `from docmarq import legal` reaches into another layer
        names += [f"{base}.{a.name}" for a in node.names]
    return {layer_of(n) for n in names if n == "docmarq" or n.startswith("docmarq.")}
  return scan

@pytest.fixture
def trees(module_of):
  """Every module of `PACKAGE`, parsed, by dotted name."""
  return {module_of(p): ast.parse(p.read_text(encoding="utf-8")) for p in PACKAGE.rglob("*.py")}

@pytest.fixture
def private_crossings(layer_of):
  """
  Underscored names a module of `trees` takes from another layer:
  imported by name, or read as an attribute only other layers define.
  A name a class defines is read as a member, so the classes judge it, and the modules any other.
  """
  def defined(body:list) -> list[str]:
    names = []
    for node in body:
      if isinstance(node, (ast.FunctionDef, ast.ClassDef)): names.append(node.name)
      elif isinstance(node, (ast.Assign, ast.AnnAssign)):
        targets = node.targets if isinstance(node, ast.Assign) else [node.target]
        names += [t.id for t in targets if isinstance(t, ast.Name)]
    return names
  def scan(trees:dict) -> list[str]:
    members, globals_ = {}, {} # underscored name → layers defining it in a class, in a module
    for module, tree in trees.items():
      found = [(globals_, name) for name in defined(tree.body)]
      for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
          found += [(members, name) for name in defined(node.body)]
        elif isinstance(node, ast.Attribute) and isinstance(node.ctx, ast.Store):
          found.append((members, node.attr))
      for homes, name in found:
        if name.startswith("_") and not name.startswith("__"):
          homes.setdefault(name, set()).add(layer_of(module))
    bad = []
    for module, tree in trees.items():
      layer, package = layer_of(module), module.rsplit(".", 1)[0]
      for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom):
          base = node.module or ""
          if node.level: base = importlib.util.resolve_name("." * node.level + base, package)
          if base.split(".")[0] != "docmarq" or layer_of(base) == layer: continue
          bad += [f"{module} imports {base}.{a.name}" for a in node.names
            if a.name.startswith("_") and not a.name.startswith("__")]
        elif isinstance(node, ast.Attribute) and isinstance(node.ctx, ast.Load):
          if layer not in (members.get(node.attr) or globals_.get(node.attr) or {layer}):
            bad.append(f"{module}:{node.lineno} reads .{node.attr}")
    return bad
  return scan

#--------------------------------------------------------------------------------------- Boundaries

def every_module_stays_in_its_layer(imported, layer_of, module_of):
  bad = []
  for path in sorted(PACKAGE.rglob("*.py")):
    module = module_of(path)
    crossed = imported(path) & FORBIDDEN[layer_of(module)]
    bad += [f"{module} imports docmarq.{layer}" for layer in sorted(crossed)]
  assert not bad, "\n".join(bad)

def scanner_sees_a_subpackage_imported_by_name(imported, tmp_path):
  src = tmp_path / "probe.py"
  src.write_text("from docmarq import legal\nfrom .. import paper\n", encoding="utf-8")
  assert {"legal", "paper"} <= imported(src, package="docmarq.md")

def scanner_resolves_relative_imports(imported):
  # guards `every_module_stays_in_its_layer` against passing on an empty scan
  assert "md" in imported(PACKAGE / "paper" / "render.py")
  assert "core" in imported(PACKAGE / "md" / "renderer.py")

def no_underscored_name_crosses_a_layer(trees, private_crossings):
  # a layer reaches the ones below through their public names, never an underscored one
  bad = private_crossings(trees)
  assert not bad, "\n".join(bad)

def scanner_sees_an_underscored_name_crossing(trees, private_crossings):
  trees["docmarq.paper.probe"] = ast.parse("from ..core import _secret\ntop = pdf._page.height\n")
  assert private_crossings(trees) == ["docmarq.paper.probe imports docmarq.core._secret",
    "docmarq.paper.probe:2 reads ._page"]
