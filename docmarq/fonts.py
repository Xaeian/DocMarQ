# docmarq/fonts.py

"""
Font configuration helpers.

Word fonts are referenced by family name, and a machine lacking one substitutes it silently.
`embed_fonts` carries faces in the file instead, as Word does when told to embed them,
so a document in faces of its own reads the same everywhere.

TTFs also matter for what this package draws itself rather than handing to Word:
mermaid diagrams and text inside an SVG. `resolve_ttf` is the one lookup all three use.
"""
import hashlib, uuid
from pathlib import Path
from lxml import etree
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.oxml.ns import qn

#------------------------------------------------------------------------------------------ Helpers

# Present on virtually every Word install,
# so a document naming one of these renders as intended without embedding.
# Exposed as data, not only through the predicate,
# because a caller building a font picker needs the list itself.
SAFE_FAMILIES: tuple[str, ...] = (
  "Arial", "Calibri", "Cambria", "Consolas", "Courier New",
  "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Verdana",
)

def is_safe_default(family:str) -> bool:
  """Returns `True` for fonts present on virtually all Word installs."""
  return family.lower() in {f.lower() for f in SAFE_FAMILIES}

#--------------------------------------------------------------------------------------- TTF lookup

# Mode fallback when a variant TTF is missing (drops styling, keeps content).
_MODE_FALLBACK = {
  "Bold": ["Regular"],
  "Italic": ["Regular"],
  "BoldItalic": ["Bold", "Italic", "Regular"],
  "Black": ["Bold", "Regular"],
  "BlackItalic": ["BoldItalic", "Black", "Bold", "Italic", "Regular"],
}

def resolve_ttf(font_dir:str, family:str, mode:str="Regular", fallback:bool=True) -> Path|None:
  """
  Find the TTF backing one face, walking `_MODE_FALLBACK` when it is absent unless `fallback`
  is off, as embedding wants: Word slants or thickens a missing face itself.

  Layout is `<font_dir>/<family>/<family>-<mode>.ttf`,
  with the subfolder named either as-is or lowercase, and a flat `<font_dir>` accepted as well.
  """
  base = Path(font_dir)
  for try_mode in [mode] + (_MODE_FALLBACK.get(mode, []) if fallback else []):
    for sub in (family.lower(), family):
      path = base / sub / f"{family}-{try_mode}.ttf"
      if path.is_file(): return path
    path = base / f"{family}-{try_mode}.ttf"
    if path.is_file(): return path
  return None

#---------------------------------------------------------------------------------------- Embedding

_OBFUSCATED = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
# Each mode of a family and the element of Word's font table that carries its file
_EMBEDS = (("Regular", "embedRegular"), ("Bold", "embedBold"), ("Italic", "embedItalic"),
  ("BoldItalic", "embedBoldItalic"))
_NAMES = ("ascii", "hAnsi", "cs", "eastAsia") # where a run or style names its face

def embed_fonts(document, font_dir:str) -> list[str]:
  """
  The faces of `font_dir` that `document` names, carried in its file, so Word shows them where
  they are not installed: of each family whatever of Regular, Bold, Italic and BoldItalic is there,
  hidden as ECMA-376 hides an embedded font. A family's files go by its name without spaces,
  `Source Serif 4` as `SourceSerif4/SourceSerif4-Bold.ttf`, as pdfmarq lays them out.
  Returns the families embedded, none when `font_dir` holds none of them.
  """
  package = document.part.package
  named = {el.get(qn(f"w:{attr}")) for part in package.iter_parts()
    if getattr(part, "element", None) is not None
    for el in part.element.iter(qn("w:rFonts")) for attr in _NAMES} - {None}
  table = document.part.part_related_by(RT.FONT_TABLE)
  root, embedded = etree.fromstring(table.blob), []
  for name in sorted(named):
    files = [(tag, resolve_ttf(font_dir, name.replace(" ", ""), mode, fallback=False))
      for mode, tag in _EMBEDS]
    files = [(tag, path) for tag, path in files if path]
    if not files: continue
    font = _font_entry(root, name)
    for tag, path in files:
      data = path.read_bytes()
      key = "{" + str(uuid.uuid5(uuid.NAMESPACE_OID, hashlib.sha1(data).hexdigest())).upper() + "}"
      partname = package.next_partname("/word/fonts/font%d.odttf")
      part = Part(partname, _OBFUSCATED, _obfuscated(data, key), package)
      embed = etree.SubElement(font, qn(f"w:{tag}"))
      embed.set(qn("r:id"), table.relate_to(part, RT.FONT))
      embed.set(qn("w:fontKey"), key)
    embedded.append(name)
  if embedded:
    table._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
  return embedded

def _font_entry(root, name:str):
  """The entry of family `name` in Word's font table, made bare when the table has none."""
  for font in root.iter(qn("w:font")):
    if font.get(qn("w:name")) == name: return font
  font = etree.SubElement(root, qn("w:font"))
  font.set(qn("w:name"), name)
  for tag, value in (("charset", "00"), ("family", "auto"), ("pitch", "variable")):
    etree.SubElement(font, qn(f"w:{tag}")).set(qn("w:val"), value)
  return font

def _obfuscated(data:bytes, key:str) -> bytes:
  """
  A font file as Word stores it embedded: its first 32 bytes XOR-ed with the bytes of `key`,
  a GUID, taken in reverse order, as ECMA-376 Part 1 §17.8.1 asks; the same step undoes it.
  """
  mask = bytes.fromhex(key.strip("{}").replace("-", ""))[::-1]
  return bytes(b ^ mask[k % 16] for k, b in enumerate(data[:32])) + data[32:]
