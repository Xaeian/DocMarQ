# tests/test_layers.py

"""Layer boundaries: core ← md ← paper and core ← md ← legal, with paper and legal strangers."""

import ast, importlib.util
from pathlib import Path
import pytest
import docmarq

PACKAGE = Path(docmarq.__file__).resolve().parent # under test: the wheel in CI, the source here
SUBPACKAGES = ("md", "paper", "legal")

# layer → layers it must never import
FORBIDDEN = {
  "core": {"md", "paper", "legal"},
  "md": {"paper", "legal"},
  "paper": {"legal"},
  "legal": {"paper"},
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
