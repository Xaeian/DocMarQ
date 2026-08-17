"""
Render a markdown file to DOCX via `docmarq`.

Usage:
  python example.py <input.md>

Showcases:
  - language preset
  - Word-native fonts (no font registration needed)
  - `link_root` for cross-document references in the rendered DOCX
  - `base_dir` so relative image paths in the markdown resolve correctly
  - YAML frontmatter → page-1 banner (content only; chrome comes from `style`)
"""
import sys
from docmarq.md import md_to_docx, lang_style
from xaeian import PATH, FILE, Print, Color as c

p = Print()

LANG = "en" # en | pl | de | fr | es | it | cs | sk
LINK_ROOT = "https://github.com/{owner}/docs/blob/main"
FONTS = dict(body_family="Calibri", mono_family="Consolas")

#----------------------------------------------------------------------------------------- Renderer

def render(in_path:str) -> str:
  """Convert markdown at `in_path` to a sibling `.docx`. Returns output path."""
  in_path = PATH.resolve(in_path, read=False)
  if not PATH.is_file(in_path):
    p.err(f"input not found | {c.RED}{in_path}{c.END}")
    sys.exit(1)
  out_path = PATH.with_suffix(in_path, ".docx")
  md_to_docx(
    FILE.load(in_path), out_path,
    style=lang_style(LANG, link_root=LINK_ROOT, **FONTS),
    base_dir=PATH.dirname(in_path),
  )
  return out_path

#-------------------------------------------------------------------------------------------- Entry

def main():
  if len(sys.argv) != 2:
    p.err(f"usage: {c.SKY}python example.py <input.md>{c.END}")
    sys.exit(1)
  out = render(sys.argv[1])
  s = FILE.size(out)
  p.ok(f"rendered | {c.SKY}{out}{c.END} | {c.GREEN}{s/1024:.1f} kB{c.END}")

if __name__ == "__main__":
  main()
