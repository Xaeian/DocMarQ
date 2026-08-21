# docmarq/md/math.py

"""
Math formula rendering for DOCX - native Office Math (OMML) with an image
fallback.

Two layers:

  1. `latex_to_omath` / `build_omath_para` - a self-contained LaTeX→OMML
     converter for the common markdown-math subset (fractions, scripts,
     roots, Greek, big operators, delimiters, matrices, accents, ...).
     The result is a true Word equation: editable in the equation editor,
     selectable, scales with font + zoom, tiny file size. No third-party
     dependency - OMML is built straight from `python-docx`'s OOXML layer.

  2. `render_math_png` - a matplotlib (mathtext) raster fallback used when
     the converter hits a construct outside the supported subset.

The converter raises `MathConversionError` for anything it can't represent;
the renderer catches it and routes that single formula through the image
fallback. Everything it *does* understand renders as native OMML.

Example:
  >>> from docmarq.md.math import latex_to_omath
  >>> el = latex_to_omath(r"E = mc^2", size_halfpt=22, color_hex="1F2328")
  >>> # el is an <m:oMath> lxml element - append it to a <w:p>
"""

__extras__ = ("math", ["matplotlib"])

import re as _re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

#------------------------------------------------------------------------------------------- Errors

class MathConversionError(Exception):
  """Raised when a LaTeX construct falls outside the supported OMML subset.
  The renderer catches this and falls back to an image render of the whole
  formula."""

#------------------------------------------------------------------------------------ Symbol tables

# LaTeX command → unicode glyph. Covers Greek, relations, operators, arrows,
# set theory and the punctuation/dots that show up in real markdown math.
SYMBOLS: dict[str, str] = {
  # lowercase Greek
  "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ϵ",
  "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
  "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
  "omicron": "ο", "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ",
  "sigma": "σ", "varsigma": "ς", "tau": "τ", "upsilon": "υ", "phi": "ϕ",
  "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
  # uppercase Greek
  "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
  "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ",
  "Omega": "Ω",
  # relations
  "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "neq": "≠", "ne": "≠",
  "equiv": "≡", "approx": "≈", "cong": "≅", "simeq": "≃", "sim": "∼",
  "propto": "∝", "ll": "≪", "gg": "≫", "doteq": "≐", "asymp": "≍",
  "prec": "≺", "succ": "≻", "preceq": "⪯", "succeq": "⪰",
  # binary operators
  "pm": "±", "mp": "∓", "times": "×", "div": "÷", "cdot": "⋅", "ast": "∗",
  "star": "⋆", "circ": "∘", "bullet": "∙", "oplus": "⊕", "ominus": "⊖",
  "otimes": "⊗", "oslash": "⊘", "odot": "⊙", "setminus": "∖",
  "wedge": "∧", "vee": "∨", "land": "∧", "lor": "∨", "neg": "¬", "lnot": "¬",
  # set theory & logic
  "in": "∈", "notin": "∉", "ni": "∋", "subset": "⊂", "supset": "⊃",
  "subseteq": "⊆", "supseteq": "⊇", "subsetneq": "⊊", "cup": "∪", "cap": "∩",
  "emptyset": "∅", "varnothing": "∅", "forall": "∀", "exists": "∃",
  "nexists": "∄", "complement": "∁", "therefore": "∴", "because": "∵",
  # arrows
  "to": "→", "rightarrow": "→", "leftarrow": "←", "leftrightarrow": "↔",
  "Rightarrow": "⇒", "Leftarrow": "⇐", "Leftrightarrow": "⇔",
  "implies": "⇒", "iff": "⇔", "mapsto": "↦", "longrightarrow": "⟶",
  "longleftarrow": "⟵", "uparrow": "↑", "downarrow": "↓", "hookrightarrow": "↪",
  # misc symbols
  "infty": "∞", "partial": "∂", "nabla": "∇", "hbar": "ℏ", "ell": "ℓ",
  "Re": "ℜ", "Im": "ℑ", "aleph": "ℵ", "wp": "℘", "prime": "′",
  "angle": "∠", "measuredangle": "∡", "triangle": "△", "square": "□",
  "diamond": "⋄", "perp": "⊥", "parallel": "∥", "mid": "∣", "nmid": "∤",
  "top": "⊤", "bot": "⊥", "vdash": "⊢", "models": "⊨",
  "degree": "°", "checkmark": "✓", "dagger": "†", "ddagger": "‡",
  "flat": "♭", "sharp": "♯", "natural": "♮", "surd": "√",
  "backslash": "\\", "%": "%", "&": "&", "#": "#", "$": "$", "_": "_",
  "{": "{", "}": "}", " ": " ",
  # dots
  "ldots": "…", "dots": "…", "cdots": "⋯", "vdots": "⋮", "ddots": "⋱",
  # named numbers / constants often used
}

# Dots/spacing that just emit literal glyphs handled above; spacing commands:
SPACES: dict[str, str] = {
  ",": " ",  # thin space
  ":": " ",  # medium space
  ";": " ",  # thick space
  "!": "",   # negative thin space - drop
  "quad": " ",
  "qquad": "  ",
  " ": " ",
}

# Function-like names rendered upright. `lim`-family additionally turns a
# trailing `_{...}` into an under-script (limLow) in display style.
FUNCTIONS: set[str] = {
  "sin", "cos", "tan", "cot", "sec", "csc",
  "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "coth",
  "log", "ln", "lg", "exp", "deg", "dim", "hom", "ker", "arg",
  "det", "gcd", "Pr", "min", "max", "sup", "inf", "lim", "limsup",
  "liminf",
}
LIM_FUNCTIONS: set[str] = {"lim", "limsup", "liminf", "max", "min", "sup", "inf",
  "argmax", "argmin"}

# Big operators → (glyph, default-omit-when-integral). Integrals use OMML's
# default n-ary char so we omit `chr`; everything else sets it explicitly.
BIG_OPERATORS: dict[str, str] = {
  "sum": "∑", "prod": "∏", "coprod": "∐",
  "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮",
  "bigcup": "⋃", "bigcap": "⋂", "bigvee": "⋁", "bigwedge": "⋀",
  "bigoplus": "⨁", "bigotimes": "⨂", "bigodot": "⨀", "biguplus": "⨄",
  "bigsqcup": "⨆",
}
# Operators whose limits sit to the side (subSup) even in display mode.
SIDE_LIMIT_OPS = {"int", "iint", "iiint", "oint"}

# Accents: command → combining char placed above the base via <m:acc>.
ACCENTS: dict[str, str] = {
  "hat": "̂", "widehat": "̂", "tilde": "̃", "widetilde": "̃",
  "bar": "̅", "vec": "⃗", "overrightarrow": "⃗",
  "dot": "̇", "ddot": "̈", "dddot": "⃛",
  "check": "̌", "acute": "́", "grave": "̀", "breve": "̆",
  "mathring": "̊",
}

# Stretchy delimiters for \left ... \right and matrix wrappers.
DELIMS: dict[str, str] = {
  "(": "(", ")": ")", "[": "[", "]": "]", "|": "|",
  "\\{": "{", "\\}": "}", "\\|": "‖", "/": "/", "\\backslash": "\\",
  "\\langle": "⟨", "\\rangle": "⟩",
  "\\lfloor": "⌊", "\\rfloor": "⌋", "\\lceil": "⌈", "\\rceil": "⌉",
  "\\vert": "|", "\\Vert": "‖", "\\lvert": "|", "\\rvert": "|",
  "\\lVert": "‖", "\\rVert": "‖",
  ".": "",  # \left. / \right. - invisible delimiter
}

# Font commands → (math script value or None, style value or None).
# scr: roman/script/fraktur/double-struck/sans-serif/monospace
# sty: p (plain/upright), b (bold), i (italic), bi (bold-italic)
FONT_CMDS: dict[str, tuple[str|None, str|None]] = {
  "mathrm": (None, "p"), "text": (None, "p"), "textrm": (None, "p"),
  "operatorname": (None, "p"), "mathnormal": (None, "i"), "mathit": (None, "i"),
  "textit": (None, "i"),
  "mathbf": (None, "b"), "textbf": (None, "b"), "boldsymbol": (None, "bi"),
  "bm": (None, "bi"), "mathsf": ("sans-serif", "p"), "textsf": ("sans-serif", "p"),
  "mathtt": ("monospace", "p"), "texttt": ("monospace", "p"),
  "mathbb": ("double-struck", "p"), "mathcal": ("script", "p"),
  "mathscr": ("script", "p"), "mathfrak": ("fraktur", "p"),
}

# Matrix environments → (begin delimiter, end delimiter). `None` = no wrapper.
MATRIX_ENVS: dict[str, tuple[str|None, str|None]] = {
  "matrix": (None, None), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
  "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"),
  "cases": ("{", ""), "array": (None, None), "aligned": (None, None),
  "smallmatrix": (None, None),
  # Multi-line equation environments: rendered as a borderless aligned grid
  # (the most common markdown display blocks). `_read_env_name` strips `*`.
  "align": (None, None), "gather": (None, None), "eqnarray": (None, None),
  "alignat": (None, None), "split": (None, None), "gathered": (None, None),
}

#------------------------------------------------------------------------------------- OMML helpers

def _el(tag:str, **attrs) -> "OxmlElement":
  """Create an OMML/WML element; attrs are `m:val`-style (namespace `m`)."""
  e = OxmlElement(tag)
  for k, v in attrs.items():
    e.set(qn(f"m:{k}"), str(v))
  return e

def _prop(parent, tag:str, **attrs):
  """Append a property child like `<m:chr m:val='...'/>` to `parent`."""
  child = _el(tag, **attrs)
  parent.append(child)
  return child

def _xml_safe(text:str) -> str:
  """Drop XML-1.0-illegal control chars (NUL etc.). lxml raises ValueError on
  them, which would escape the MathConversionError contract; stripping keeps a
  stray control byte from crashing the converter."""
  return "".join(
    ch for ch in text
    if ord(ch) in (9, 10, 13)
    or 0x20 <= ord(ch) <= 0xD7FF
    or 0xE000 <= ord(ch) <= 0xFFFD
    or 0x10000 <= ord(ch) <= 0x10FFFF
  )

class _Build:
  """OMML builder bound to a run color + size so every leaf run is styled
  consistently. Kept tiny - just `run()` plus the structural constructors
  live on the parser."""

  def __init__(self, size_halfpt:int|None, color_hex:str|None):
    self.size_halfpt = size_halfpt
    self.color_hex = color_hex

  def run(self, text:str, sty:str|None=None, scr:str|None=None) -> "OxmlElement":
    """Build `<m:r>` with optional math style/script and shared color/size.

    `sty`: p/b/i/bi (plain/bold/italic/bold-italic). Letters default to
    italic in Word's math font, so we only emit a style when overriding.
    """
    r = _el("m:r")
    if sty or scr:
      rpr = _el("m:rPr")
      if scr: _prop(rpr, "m:scr", val=scr)
      if sty: _prop(rpr, "m:sty", val=sty)
      r.append(rpr)
    if self.color_hex or self.size_halfpt:
      wrpr = OxmlElement("w:rPr")
      if self.color_hex:
        c = OxmlElement("w:color"); c.set(qn("w:val"), self.color_hex); wrpr.append(c)
      if self.size_halfpt:
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(self.size_halfpt)); wrpr.append(sz)
        szc = OxmlElement("w:szCs"); szc.set(qn("w:val"), str(self.size_halfpt)); wrpr.append(szc)
      r.append(wrpr)
    t = _el("m:t")
    if text != text.strip() or text == "":
      t.set(qn("xml:space"), "preserve")
    t.text = _xml_safe(text)
    r.append(t)
    return r

  def wrap(self, tag:str, nodes:list) -> "OxmlElement":
    """Wrap a node list into a container element (`m:e`, `m:num`, ...)."""
    box = _el(tag)
    for n in nodes:
      box.append(n)
    return box

#---------------------------------------------------------------------------------------- Tokenizer

def _tokenize(src:str) -> list[str]:
  """Split LaTeX math source into tokens: commands (`\\frac`, `\\,`), the
  structural chars `{} ^ _ &`, the row break `\\\\`, single characters, and a
  single `" "` token for any whitespace run.

  Whitespace is insignificant in math mode (the cursor skips `" "` tokens),
  but `\\text{a b}` must keep its spaces - so we emit one collapsed space
  token rather than dropping whitespace, letting text mode recover it."""
  tokens: list[str] = []
  i, n = 0, len(src)
  while i < n:
    c = src[i]
    if c.isspace():
      while i < n and src[i].isspace():
        i += 1
      tokens.append(" ")
      continue
    if c == "\\":
      if i + 1 < n and src[i+1] == "\\":
        tokens.append("\\\\")  # row break
        i += 2
        continue
      j = i + 1
      if j < n and src[j].isalpha():
        k = j
        while k < n and src[k].isalpha():
          k += 1
        tokens.append(src[i:k])
        i = k
      elif j < n:
        tokens.append(src[i:j+1])  # \{  \,  \|  \%  etc.
        i = j + 1
      else:
        tokens.append("\\")
        i += 1
    elif c in "{}^_&":
      tokens.append(c)
      i += 1
    else:
      tokens.append(c)
      i += 1
  return tokens

#------------------------------------------------------------------------------------------- Parser

class _Parser:
  """Recursive-descent LaTeX-math → OMML converter.

  `parse_sequence` produces a flat node list (runs + structures) for a
  container; scripts (`^`/`_`) bind to the preceding atom; big operators
  and `\\lim` consume their limits + one operand. Anything unrecognized
  raises `MathConversionError` so the caller can fall back to an image.
  """

  # Recursion guard: raise MathConversionError (→ image), never RecursionError.
  MAX_DEPTH = 120

  def __init__(self, tokens:list[str], build:_Build, display:bool):
    self.toks = tokens
    self.pos = 0
    self.b = build
    self.display = display
    self.depth = 0

  #----------------------------------------------------------------------------------- token cursor

  # `" "` tokens are insignificant in math mode - the cursor transparently
  # skips them so parsing never sees whitespace. Text mode reads them back
  # via `_grab_group_raw_tokens` to preserve spaces inside `\text{...}`.

  def _skip_ws(self):
    while self.pos < len(self.toks) and self.toks[self.pos] == " ":
      self.pos += 1

  def _peek(self) -> str|None:
    i = self.pos
    while i < len(self.toks) and self.toks[i] == " ":
      i += 1
    return self.toks[i] if i < len(self.toks) else None

  def _next(self) -> str|None:
    self._skip_ws()
    t = self.toks[self.pos] if self.pos < len(self.toks) else None
    if t is not None:
      self.pos += 1
    return t

  def _expect(self, tok:str):
    if self._peek() != tok:
      raise MathConversionError(f"expected {tok!r}, got {self._peek()!r}")
    self._next()

  #-------------------------------------------------------------------------------------- sequences

  def parse_sequence(self, stops:set[str]) -> list:
    """Parse nodes until a stop token (not consumed) or EOF."""
    out: list = []
    while True:
      t = self._peek()
      if t is None or t in stops:
        break
      if t in ("&", "\\\\"):
        # Stray cell/row separators outside a matrix: treat as a space.
        self.pos += 1
        out.append(self.b.run(" "))
        continue
      out.extend(self.parse_one())
    return out

  def parse_one(self) -> list:
    """Parse one scripted atom (atom + trailing sub/sup, or a big operator
    with its limits + operand). Returns a node list."""
    # Big operators capture their own limits/operand.
    t = self._peek()
    if t is not None and t.startswith("\\") and t[1:] in BIG_OPERATORS:
      return [self._parse_nary(t[1:])]
    if t is not None and t.startswith("\\") and t[1:] in LIM_FUNCTIONS:
      return [self._parse_limop(t[1:])]
    base = self.parse_atom()
    return self._attach_scripts(base)

  def _attach_scripts(self, base:list) -> list:
    """Consume `^`/`_` following `base` and wrap into sSup/sSub/sSubSup."""
    sub = sup = None
    while self._peek() in ("^", "_"):
      op = self._next()
      script = self.parse_atom()
      if op == "_":
        if sub is not None:
          raise MathConversionError("double subscript")
        sub = script
      else:
        if sup is not None:
          raise MathConversionError("double superscript")
        sup = script
    if sub is None and sup is None:
      return base
    e = self.b.wrap("m:e", base)
    if sub is not None and sup is not None:
      node = _el("m:sSubSup")
      node.append(e)
      node.append(self.b.wrap("m:sub", sub))
      node.append(self.b.wrap("m:sup", sup))
    elif sub is not None:
      node = _el("m:sSub")
      node.append(e)
      node.append(self.b.wrap("m:sub", sub))
    else:
      node = _el("m:sSup")
      node.append(e)
      node.append(self.b.wrap("m:sup", sup))
    return [node]

  #------------------------------------------------------------------------------------------ atoms

  def parse_atom(self) -> list:
    """Parse a single base atom (no scripts). Returns a node list so a
    braced group `{...}` can expand to multiple runs under one base."""
    # Depth guard: pathological nesting must raise MathConversionError (caught
    # by the renderer → image fallback), never overflow into RecursionError
    # (which is uncaught and would abort the whole document).
    self.depth += 1
    if self.depth > self.MAX_DEPTH:
      raise MathConversionError("nesting too deep")
    try:
      return self._parse_atom_inner()
    finally:
      self.depth -= 1

  def _parse_atom_inner(self) -> list:
    t = self._next()
    if t is None:
      raise MathConversionError("unexpected end of formula")
    if t == "{":
      seq = self.parse_sequence({"}"})
      self._expect("}")
      return seq
    if t == "}":
      raise MathConversionError("unbalanced '}'")
    if t in ("^", "_"):
      # script with no base, e.g. leading `_x`: empty base.
      self.pos -= 1
      return [self.b.run("")]
    if t.startswith("\\"):
      return self._parse_command(t)
    # Ordinary character.
    return [self._char_run(t)]

  def _char_run(self, ch:str):
    """Single-character run. Digits/operators stay upright automatically in
    Word's math font; single ASCII letters render italic by default."""
    if ch == "-":
      ch = "−"  # proper minus sign
    return self.b.run(ch)

  #--------------------------------------------------------------------------------------- commands

  def _parse_command(self, cmd:str) -> list:
    name = cmd[1:]
    # Two-char escapes: \,  \;  \!  \{  \}  \%  \&  \#  \$  \_  \|  \\ handled elsewhere
    if cmd in DELIMS and name not in SYMBOLS:
      # bare \{ \} \| \langle etc. used outside \left → literal glyph
      if cmd in ("\\{", "\\}", "\\|", "\\langle", "\\rangle", "\\vert", "\\Vert",
        "\\lvert", "\\rvert", "\\lVert", "\\rVert",
        "\\lfloor", "\\rfloor", "\\lceil", "\\rceil", "\\backslash"):
        return [self.b.run(DELIMS[cmd])]
    if name in SPACES:
      glyph = SPACES[name]
      return [self.b.run(glyph)] if glyph else []
    if name in SYMBOLS:
      return [self.b.run(SYMBOLS[name])]
    if name in FUNCTIONS:
      return [self.b.run(name, sty="p"), self.b.run(" ")]
    if name == "bmod" or name == "mod":
      return [self.b.run(" "), self.b.run("mod", sty="p"), self.b.run(" ")]
    if name == "pmod":
      arg = self.parse_atom()
      return [self.b.run("  "),
        self._delimit("(", ")", [self.b.run("mod ", sty="p")] + arg)]
    if name == "frac" or name == "dfrac" or name == "tfrac" or name == "cfrac":
      return [self._parse_frac(bar=True)]
    if name == "binom" or name == "dbinom" or name == "tbinom":
      return [self._parse_binom()]
    if name == "sqrt":
      return [self._parse_sqrt()]
    if name in ACCENTS:
      return [self._parse_accent(ACCENTS[name])]
    if name in ("overline",):
      return [self._parse_bar(pos="top")]
    if name in ("underline",):
      return [self._parse_bar(pos="bot")]
    if name in ("overbrace", "underbrace"):
      return [self._parse_groupchr(name)]
    if name in FONT_CMDS:
      return self._parse_font(name)
    if name == "left":
      return [self._parse_delim()]
    if name == "begin":
      return [self._parse_environment()]
    if name in ("displaystyle", "textstyle", "scriptstyle", "limits", "nolimits"):
      return []  # styling hints - no structural effect here
    if name in ("right", "end"):
      raise MathConversionError(f"stray \\{name}")
    raise MathConversionError(f"unsupported command \\{name}")

  #------------------------------------------------------------------------------------- structures

  def _parse_frac(self, bar:bool):
    num = self.parse_atom()
    den = self.parse_atom()
    f = _el("m:f")
    if not bar:
      fpr = _el("m:fPr"); _prop(fpr, "m:type", val="noBar"); f.append(fpr)
    f.append(self.b.wrap("m:num", num))
    f.append(self.b.wrap("m:den", den))
    return f

  def _parse_binom(self):
    num = self.parse_atom()
    den = self.parse_atom()
    f = _el("m:f")
    fpr = _el("m:fPr"); _prop(fpr, "m:type", val="noBar"); f.append(fpr)
    f.append(self.b.wrap("m:num", num))
    f.append(self.b.wrap("m:den", den))
    return self._delimit("(", ")", [f])

  def _parse_sqrt(self):
    deg = None
    if self._peek() == "[":
      self._next()
      deg = self.parse_sequence({"]"})
      self._expect("]")
    rad_arg = self.parse_atom()
    rad = _el("m:rad")
    if deg is None:
      radpr = _el("m:radPr"); _prop(radpr, "m:degHide", val="1"); rad.append(radpr)
      rad.append(_el("m:deg"))
    else:
      rad.append(self.b.wrap("m:deg", deg))
    rad.append(self.b.wrap("m:e", rad_arg))
    return rad

  def _parse_accent(self, chr_:str):
    arg = self.parse_atom()
    acc = _el("m:acc")
    accpr = _el("m:accPr"); _prop(accpr, "m:chr", val=chr_); acc.append(accpr)
    acc.append(self.b.wrap("m:e", arg))
    return acc

  def _parse_bar(self, pos:str):
    arg = self.parse_atom()
    bar = _el("m:bar")
    barpr = _el("m:barPr"); _prop(barpr, "m:pos", val=pos); bar.append(barpr)
    bar.append(self.b.wrap("m:e", arg))
    return bar

  def _parse_groupchr(self, name:str):
    arg = self.parse_atom()
    is_over = name == "overbrace"
    g = _el("m:groupChr")
    gpr = _el("m:groupChrPr")
    _prop(gpr, "m:chr", val="⏞" if is_over else "⏟")
    _prop(gpr, "m:pos", val="top" if is_over else "bot")
    _prop(gpr, "m:vertJc", val="bot" if is_over else "top")
    g.append(gpr)
    g.append(self.b.wrap("m:e", arg))
    node = g
    # `^`/`_` after overbrace/underbrace become the brace label.
    if (is_over and self._peek() == "^") or (not is_over and self._peek() == "_"):
      self._next()
      label = self.parse_atom()
      limtag = "m:limUpp" if is_over else "m:limLow"
      lim = _el(limtag)
      lim.append(self.b.wrap("m:e", [node]))
      lim.append(self.b.wrap("m:lim", label))
      node = lim
    return node

  # Commands whose body is LITERAL TEXT: spaces preserved, `^`/`_` NOT parsed
  # as math. Style comes from FONT_CMDS (textbf→bold, textit→italic, ...).
  TEXT_MODE = {"text", "textrm", "textbf", "textit", "textsf", "texttt",
    "mathrm", "operatorname"}

  def _parse_font(self, name:str):
    scr, sty = FONT_CMDS[name]
    if name in self.TEXT_MODE:
      # Text mode: emit the literal characters as one styled run. Use the raw
      # token stream so spaces inside `\text{a b}` survive, and don't re-parse
      # `^`/`_` as math superscripts/subscripts.
      text = self._tokens_to_text(self._grab_group_raw_tokens())
      return [self.b.run(text, sty=sty, scr=scr)]
    # Math mode font (\mathbf, \mathbb, \mathcal, ...): parse the group as math
    # and tag each leaf run with the font's style/script.
    arg_tokens = self._grab_group_tokens()
    sub = _Parser(arg_tokens, self.b, self.display)
    nodes = sub.parse_sequence(set())
    _restyle_runs(nodes, sty, scr)
    return nodes

  def _parse_delim(self):
    open_tok = self._next()
    open_chr = self._delim_glyph(open_tok)
    inner = self.parse_sequence({"\\right"})
    self._expect("\\right")
    close_tok = self._next()
    close_chr = self._delim_glyph(close_tok)
    return self._delimit(open_chr, close_chr, inner)

  # Single chars accepted as bare \left/\right delimiters (letters are not).
  _DELIM_CHARS = set("()[]|/.<>")

  def _delim_glyph(self, tok:str|None) -> str:
    if tok is None:
      raise MathConversionError("missing delimiter")
    if tok in DELIMS:
      return DELIMS[tok]
    if len(tok) == 1 and tok in self._DELIM_CHARS:
      return tok
    raise MathConversionError(f"unsupported delimiter {tok!r}")

  def _delimit(self, open_chr:str, close_chr:str, nodes:list, sep:str|None=None):
    d = _el("m:d")
    need_pr = open_chr != "(" or close_chr != ")" or sep is not None
    if need_pr:
      dpr = _el("m:dPr")
      _prop(dpr, "m:begChr", val=open_chr)
      if sep is not None:
        _prop(dpr, "m:sepChr", val=sep)
      _prop(dpr, "m:endChr", val=close_chr)
      d.append(dpr)
    d.append(self.b.wrap("m:e", nodes))
    return d

  #---------------------------------------------------------------------------------- big operators

  def _parse_nary(self, name:str):
    self._next()  # consume the operator command
    sub = sup = None
    # Limits may appear in either order; skip \limits / \nolimits hints.
    while True:
      t = self._peek()
      if t == "_":
        self._next()
        if sub is not None:
          raise MathConversionError("double subscript")
        sub = self.parse_atom()
      elif t == "^":
        self._next()
        if sup is not None:
          raise MathConversionError("double superscript")
        sup = self.parse_atom()
      elif t in ("\\limits", "\\nolimits", "\\displaystyle"):
        self._next()
      else:
        break
    operand = self._parse_operand()
    nary = _el("m:nary")
    pr = _el("m:naryPr")
    glyph = BIG_OPERATORS[name]
    if name not in ("int",):  # int's default n-ary char is already ∫
      _prop(pr, "m:chr", val=glyph)
    side = (name in SIDE_LIMIT_OPS) or (not self.display)
    _prop(pr, "m:limLoc", val="subSup" if side else "undOvr")
    # CT_NaryPr child order is strict: chr, limLoc, grow, subHide, supHide.
    _prop(pr, "m:grow", val="1")
    if sub is None: _prop(pr, "m:subHide", val="1")
    if sup is None: _prop(pr, "m:supHide", val="1")
    nary.append(pr)
    nary.append(self.b.wrap("m:sub", sub or []))
    nary.append(self.b.wrap("m:sup", sup or []))
    nary.append(self.b.wrap("m:e", operand))
    return nary

  def _parse_operand(self) -> list:
    """One scripted atom as a big operator's operand. Stops at separators /
    relations would be ideal, but a single atom matches LaTeX binding and
    avoids over-greedy capture; the remainder trails as siblings."""
    t = self._peek()
    if t is None or t in ("}", "&", "\\\\", "\\right", "\\end"):
      return []
    if t in ("+", "-", "=", ")", "]"):
      return []  # operator/closer immediately after - empty operand
    return self.parse_one()

  def _parse_limop(self, name:str):
    self._next()
    # \max, \min, ... without a subscript are just upright function names.
    if self._peek() != "_":
      r = self.b.run(name, sty="p")
      return self._wrap_single(self._attach_scripts([r]))
    self._next()  # consume `_`
    lim = self.parse_atom()
    node = _el("m:limLow")
    node.append(self.b.wrap("m:e", [self.b.run(name, sty="p")]))
    node.append(self.b.wrap("m:lim", lim))
    # A trailing `^` (e.g. `\lim_x^2`) binds to the whole operator as a
    # superscript - wrap the limLow rather than orphaning `^` onto an empty base.
    if self._peek() == "^":
      self._next()
      sup = self.parse_atom()
      ssup = _el("m:sSup")
      ssup.append(self.b.wrap("m:e", [node]))
      ssup.append(self.b.wrap("m:sup", sup))
      return ssup
    return node

  @staticmethod
  def _wrap_single(nodes:list):
    if len(nodes) == 1:
      return nodes[0]
    box = _el("m:e")
    for n in nodes:
      box.append(n)
    return box

  #----------------------------------------------------------------------------------- environments

  def _parse_environment(self):
    env = self._read_env_name()
    if env not in MATRIX_ENVS:
      raise MathConversionError(f"unsupported environment {env!r}")
    # `array` (and friends) carry an optional `[pos]` then a mandatory
    # `{col-spec}` argument; consume both so the spec doesn't leak into cell 1.
    if env == "array":
      if self._peek() == "[":
        self._copy_bracket_group()
      if self._peek() == "{":
        self._grab_group_tokens()  # discard column spec
    # Collect raw tokens until the matching \end{env}, copying nested
    # \begin{..}/\end{..} verbatim (incl. braces) so the inner sub-parser can
    # re-read them. Both brace and begin/end nesting are tracked.
    depth = 1
    cell_tokens: list[str] = []
    while True:
      t = self._peek()
      if t is None:
        raise MathConversionError(f"unterminated \\begin{{{env}}}")
      if t == "\\begin":
        depth += 1
        self._next()
        cell_tokens.append("\\begin")
        cell_tokens.extend(self._copy_brace_group())
        continue
      if t == "\\end":
        self._next()
        grp = self._copy_brace_group()
        depth -= 1
        if depth == 0:
          break
        cell_tokens.append("\\end")
        cell_tokens.extend(grp)
        continue
      cell_tokens.append(self._next())
    return self._build_matrix(env, cell_tokens)

  def _read_env_name(self) -> str:
    self._expect("{")
    name = ""
    while self._peek() not in ("}", None):
      name += self._next()
    self._expect("}")
    return name.rstrip("*")

  def _copy_brace_group(self) -> list[str]:
    """Consume a `{...}` and return the literal tokens incl. both braces, so a
    nested `\\begin{matrix}` round-trips into the cell stream as separate
    `{`, name-chars..., `}` tokens (not a single re-serialized blob)."""
    self._expect("{")
    out = ["{"]
    depth = 1
    while True:
      t = self._next()
      if t is None:
        raise MathConversionError("unterminated environment name")
      out.append(t)
      if t == "{":
        depth += 1
      elif t == "}":
        depth -= 1
        if depth == 0:
          return out

  def _copy_bracket_group(self):
    """Consume and discard an optional `[...]` argument (e.g. array `[t]`)."""
    self._next()  # `[`
    while True:
      t = self._next()
      if t is None or t == "]":
        return

  def _build_matrix(self, env:str, tokens:list[str]):
    rows = _split_rows(tokens)
    mat = _el("m:m")
    if not rows:
      raise MathConversionError("empty matrix")
    ncols = max(len(r) for r in rows)
    mpr = _el("m:mPr")
    _prop(mpr, "m:baseJc", val="center")
    _prop(mpr, "m:plcHide", val="1")
    mcs = _el("m:mcs")
    mc = _el("m:mc")
    mcpr = _el("m:mcPr")
    _prop(mcpr, "m:count", val=str(ncols))
    _prop(mcpr, "m:mcJc", val="left" if env == "cases" else "center")
    mc.append(mcpr); mcs.append(mc); mpr.append(mcs)
    mat.append(mpr)
    for row in rows:
      mr = _el("m:mr")
      for c in range(ncols):
        cell = row[c] if c < len(row) else []
        sub = _Parser(cell, self.b, self.display)
        nodes = sub.parse_sequence(set())
        mr.append(self.b.wrap("m:e", nodes))
      mat.append(mr)
    beg, end = MATRIX_ENVS[env]
    if beg is None and end is None:
      return mat
    return self._delimit(beg or "", end or "", [mat])

  #----------------------------------------------------------------------------------- text helpers

  def _grab_group_tokens(self) -> list[str]:
    """Consume a `{...}` (or a single token) and return its inner tokens."""
    if self._peek() == "{":
      self._next()
      depth = 1
      buf: list[str] = []
      while True:
        t = self._next()
        if t is None:
          raise MathConversionError("unterminated group")
        if t == "{":
          depth += 1
        elif t == "}":
          depth -= 1
          if depth == 0:
            break
        buf.append(t)
      return buf
    t = self._next()
    if t is None:
      raise MathConversionError("missing argument")
    return [t]

  def _grab_group_raw_tokens(self) -> list[str]:
    """Like `_grab_group_tokens` but keeps `" "` (space) tokens, for text mode
    where `\\text{a b}` must preserve its spaces."""
    self._skip_ws()
    if self.pos < len(self.toks) and self.toks[self.pos] == "{":
      self.pos += 1
      depth = 1
      buf: list[str] = []
      while self.pos < len(self.toks):
        t = self.toks[self.pos]
        self.pos += 1
        if t == "{":
          depth += 1
        elif t == "}":
          depth -= 1
          if depth == 0:
            return buf
        buf.append(t)
      raise MathConversionError("unterminated group")
    # Single (non-space) token argument.
    t = self._next()
    if t is None:
      raise MathConversionError("missing argument")
    return [t]

  @staticmethod
  def _tokens_to_text(tokens:list[str]) -> str:
    """Flatten tokens back to literal text for `\\text{...}`. Known escapes
    map to their glyph; unknown commands keep their name."""
    out = []
    for t in tokens:
      if t == "\\\\":
        out.append(" ")
      elif t.startswith("\\"):
        name = t[1:]
        if name in SPACES:
          out.append(SPACES[name])
        elif name in SYMBOLS:
          out.append(SYMBOLS[name])
        elif t in DELIMS:
          out.append(DELIMS[t])
        else:
          out.append(name)
      else:
        out.append(t)
    return "".join(out)

#---------------------------------------------------------------------------------- post-processing

def _restyle_runs(nodes:list, sty:str|None, scr:str|None):
  """Apply a font style/script to every `<m:r>` descendant in `nodes`.

  CT_RPr has a strict child order (lit, nor, scr, sty, brk, aln), so when a
  scr-bearing font (`\\mathbb`) wraps a sty-bearing one (`\\mathbf`/`\\sin`)
  we must INSERT `m:scr` before the existing `m:sty`, not append after it."""
  for node in nodes:
    if node.tag == qn("m:r"):
      rpr = node.find(qn("m:rPr"))
      if rpr is None:
        rpr = _el("m:rPr")
        node.insert(0, rpr)
      if scr is not None and rpr.find(qn("m:scr")) is None:
        scr_el = _el("m:scr", val=scr)
        sty_el = rpr.find(qn("m:sty"))
        if sty_el is not None:
          sty_el.addprevious(scr_el)  # keep scr before sty
        else:
          rpr.append(scr_el)
      if sty is not None and rpr.find(qn("m:sty")) is None:
        _prop(rpr, "m:sty", val=sty)
    else:
      _restyle_runs(list(node), sty, scr)

def _coalesce_runs(el) -> None:
  """Merge adjacent `<m:r>` siblings with identical run properties into one
  run, concatenating their text. Mirrors how Word itself stores math (one run
  per styled span, not per character), shrinking output.

  Rebuilds the child list in one pass (detach-all then re-append survivors) so
  it stays O(n) - a naive per-merge `remove()` would be O(n^2) on big spans."""
  from lxml import etree
  def sig(r):
    parts = []
    for tag in ("m:rPr", "w:rPr"):
      e = r.find(qn(tag))
      parts.append(etree.tostring(e) if e is not None else b"")
    return tuple(parts)
  def is_simple_run(r):
    return r.tag == qn("m:r") and len(r.findall(qn("m:t"))) == 1
  children = list(el)
  for child in children:
    _coalesce_runs(child)
  merged = []
  merged_sig = None
  for child in children:
    if (is_simple_run(child) and merged and is_simple_run(merged[-1])
        and merged_sig == sig(child)):
      pt = merged[-1].find(qn("m:t"))
      ct = child.find(qn("m:t"))
      pt.text = (pt.text or "") + (ct.text or "")
      if pt.text != pt.text.strip() or pt.text == "":
        pt.set(qn("xml:space"), "preserve")
      continue
    merged.append(child)
    merged_sig = sig(child) if is_simple_run(child) else None
  if len(merged) != len(children):
    del el[:]
    el.extend(merged)

def _split_rows(tokens:list[str]) -> list[list[list[str]]]:
  """Split a matrix token stream into rows (`\\\\`) of cells (`&`)."""
  rows: list[list[list[str]]] = []
  cur_row: list[list[str]] = []
  cur_cell: list[str] = []
  depth = 0      # brace nesting
  env = 0        # \begin/\end nesting - inner separators must not split us
  for t in tokens:
    if t == "{":
      depth += 1; cur_cell.append(t)
    elif t == "}":
      depth -= 1; cur_cell.append(t)
    elif t == "\\begin":
      env += 1; cur_cell.append(t)
    elif t == "\\end":
      env -= 1; cur_cell.append(t)
    elif t == "&" and depth == 0 and env == 0:
      cur_row.append(cur_cell); cur_cell = []
    elif t == "\\\\" and depth == 0 and env == 0:
      cur_row.append(cur_cell); rows.append(cur_row)
      cur_row = []; cur_cell = []
    else:
      cur_cell.append(t)
  if cur_cell or cur_row:
    cur_row.append(cur_cell)
    rows.append(cur_row)
  # Drop a trailing empty row created by a final `\\`.
  rows = [r for r in rows if any(cell for cell in r)]
  return rows

#--------------------------------------------------------------------------------------- Public API

def latex_to_omath(latex:str, size_halfpt:int|None=None,
    color_hex:str|None=None, display:bool=False) -> "OxmlElement":
  """Convert LaTeX math to an `<m:oMath>` OMML element.

  Args:
    latex: Formula source without `$` delimiters.
    size_halfpt: Run size in half-points (e.g. 22 for 11pt). `None` inherits.
    color_hex: Run color as 6-hex (no `#`). `None` = default (black).
    display: `True` for block/display style (limits under/over big ops).

  Returns an `<m:oMath>` lxml element ready to append to a `<w:p>`.
  Raises `MathConversionError` for unsupported constructs.
  """
  if not latex or not latex.strip():
    raise MathConversionError("empty formula")
  if len(latex) > 5000:
    # Real formulas are never this long; route oversized input to the image
    # fallback rather than building a pathological OMML tree.
    raise MathConversionError("formula too long")
  build = _Build(size_halfpt, color_hex)
  parser = _Parser(_tokenize(latex), build, display)
  nodes = parser.parse_sequence(set())
  parser._skip_ws()  # a trailing space token is not "leftover" content
  if parser.pos < len(parser.toks):
    raise MathConversionError(f"trailing tokens at {parser.pos}")
  if not nodes:
    raise MathConversionError("no content")
  omath = _el("m:oMath")
  for n in nodes:
    omath.append(n)
  _coalesce_runs(omath)
  return omath

def build_omath_para(latex:str, size_halfpt:int|None=None,
    color_hex:str|None=None, align:str="center") -> "OxmlElement":
  """Convert LaTeX to a display `<m:oMathPara>` (centered block equation).

  Wraps `latex_to_omath(display=True)` in an `<m:oMathPara>` with the given
  justification. Raises `MathConversionError` on unsupported input.
  """
  omath = latex_to_omath(latex, size_halfpt, color_hex, display=True)
  para = _el("m:oMathPara")
  pr = _el("m:oMathParaPr")
  _prop(pr, "m:jc", val=align)
  para.append(pr)
  para.append(omath)
  return para

#----------------------------------------------------------------------------------- Image fallback

_BIG_DELIM_RE = _re.compile(r"\\(?:bigg?|Bigg?)[lrm]?\s*")
_CANCEL_RE = _re.compile(r"\\(?:cancel|bcancel|xcancel)\s*")
_TAG_RE = _re.compile(r"\\tag\s*\{[^{}]*\}")
_COLOR_RE = _re.compile(r"\\(?:text)?color\s*\{[^{}]*\}")

def _normalize_for_mathtext(latex:str) -> str:
  """Strip a few common constructs matplotlib mathtext can't parse so the
  image fallback succeeds instead of degrading to text. Conservative: only
  removes wrappers, never changes the math meaning."""
  s = latex
  s = _BIG_DELIM_RE.sub("", s)  # \big( → ( ; mathtext sizes via \left\right
  s = _CANCEL_RE.sub("", s)     # \cancel{x} → {x}
  s = _TAG_RE.sub("", s)        # drop \tag{1}
  s = _COLOR_RE.sub("", s)      # \color{red}{x} → {x}; \textcolor{red}{x} → {x}
  return s

def render_math_png(latex:str, fontsize_pt:float=11,
    color:tuple=(0, 0, 0), fontset:str="stix", dpi:int=300):
  """Render a formula to a transparent PNG via matplotlib mathtext.

  Returns `(BytesIO, width_mm, height_mm, baseline_from_bottom_mm)` or
  `None` when matplotlib is unavailable or the formula fails to parse.
  The baseline offset lets the caller drop an inline picture so its math
  baseline sits on the text baseline.
  """
  try:
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
  except ImportError:
    return None
  import io
  latex = _normalize_for_mathtext(latex)
  try:
    import matplotlib as mpl
    mpl.rcParams["mathtext.fontset"] = fontset if fontset in (
      "stix", "stixsans", "cm", "dejavusans", "dejavuserif") else "stix"
    fig = Figure(figsize=(10, 2), dpi=dpi)
    fig.patch.set_alpha(0)
    canvas = FigureCanvasAgg(fig)
    fig.text(0, 0, f"${latex}$", fontsize=fontsize_pt, color=color,
      ha="left", va="baseline")
    canvas.draw()
    renderer = canvas.get_renderer()
    bbox_in = fig.get_tightbbox(renderer)  # inches
    w_mm = bbox_in.width * 25.4
    h_mm = bbox_in.height * 25.4
    baseline_mm = (-bbox_in.y0) * 25.4  # baseline sits at fig y=0
    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True,
      bbox_inches="tight", pad_inches=0.01, dpi=dpi)
    buf.seek(0)
    return buf, w_mm, h_mm, baseline_mm
  except Exception:
    return None
