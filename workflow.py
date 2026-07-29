# workflow.py

"""
Auto-generate GitHub Actions workflow for PyPI publishing.

Reads __repo__ and __python__ from package __init__.py,
generates a publish.yml with trusted publishing via OIDC.

Example:
  >>> from workflow import generate
  >>> generate("xaeian")

CLI:
  py workflow.py xaeian
  py workflow.py xaeian -o .github/workflows/publish.yml
"""

import sys
from xaeian import FILE, PATH, Print, Color as c
from toml import get_meta

p = Print()

#------------------------------------------------------------------------------------- Analysis

def has_svglib(root:str) -> bool:
  """Check if `svglib` is in `pyproject.toml` dependencies."""
  toml = PATH.join(root, "pyproject.toml")
  if not PATH.is_file(toml): return False
  try: return "svglib" in FILE.load(toml)
  except Exception: return False

#------------------------------------------------------------------------------------- Generate

def generate_workflow(meta:dict, cairo:bool=False) -> str:
  """Generate publish.yml content."""
  python_ver = meta["python"].replace(">=", "").replace(">", "")
  cairo_step = (
    "      - run: sudo apt-get update && "
    "sudo apt-get install -y libcairo2-dev pkg-config python3-dev\n"
  ) if cairo else ""
  return f'''name: Publish PyPI

on:
  release:
    types: [published]

jobs:
  publish:
    runs-on: ubuntu-latest
    environment: pypi
    permissions:
      id-token: write
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "{python_ver}"
{cairo_step}      - run: pip install build
      - run: python -m build
      - run: pip install -e .
      - uses: pypa/gh-action-pypi-publish@release/v1
'''

#--------------------------------------------------------------------------------------- Public

def generate(package:str, output:str|None=None):
  """Generate GitHub Actions workflow for given package.

  Args:
    package: Package directory path.
    output: Output file path (default: .github/workflows/publish.yml).
  """
  pkg_dir = PATH.resolve(package)
  if not PATH.is_dir(pkg_dir):
    p.err(f"{c.ORANGE}{pkg_dir}{c.END} is not a directory")
    sys.exit(1)
  meta = get_meta(pkg_dir)
  root = PATH.dirname(pkg_dir)
  cairo = has_svglib(root)
  python_ver = meta["python"].replace(">=", "").replace(">", "")
  p.inf(f"Python: {c.TURQUS}{python_ver}{c.END}")
  if meta["repo"]:
    p.gap(f"https://github.com/{c.SKY}{meta['repo']}{c.END}")
  if cairo:
    p.inf(f"Cairo: {c.VIOLET}libcairo2-dev{c.GREY} (svglib detected){c.END}")
  workflow = generate_workflow(meta, cairo)
  out = output or PATH.join(root, ".github", "workflows", "publish.yml")
  FILE.save(out, workflow)
  p.ok(f"Generated {c.GREY}{PATH.dirname(out)}/{c.END}{c.ORANGE}{PATH.basename(out)}{c.END}")

#------------------------------------------------------------------------------------------ CLI

EXAMPLES = """
examples:
  py workflow.py xaeian                 Generate with defaults
  py workflow.py xaeian -o publish.yml  Custom output path
"""

if __name__ == "__main__":
  from xaeian.cli._args import _make_parser, _add_help
  parser = _make_parser("Generate GitHub Actions workflow for PyPI publishing", EXAMPLES)
  parser.add_argument("package", metavar="PACKAGE", help="Package directory to scan")
  parser.add_argument("-o", "--output", default=None, metavar="PATH",
    help="Output file (default: .github/workflows/publish.yml)")
  _add_help(parser)
  args = parser.parse_args()
  generate(args.package, args.output)