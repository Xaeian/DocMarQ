"""
Mermaid diagram rendering with hybrid backends.

Mermaid is a JavaScript-only library - no native Python implementation
exists. We try multiple rendering backends in priority order and use the
first one that succeeds:

  1. `mermaid-cli` (mmdc) - local subprocess, best quality. Requires Node.js
     plus `npm install -g @mermaid-js/mermaid-cli`. Honors
     `DOCMARQ_MMDC_PUPPETEER_CONFIG` env var pointing at a JSON file with
     `executablePath` if Chrome isn't on the default puppeteer cache path.
  2. `mermaid.ink` - public HTTP service, used when mmdc is missing or
     fails. **The diagram source is sent to a third-party server.** Warns
     once per process; `MarkdownStyle(mermaid_remote=False)` or
     `compile_to_png(remote=False)` keeps rendering local.
  3. `None` - no backend succeeded; caller falls back to a code block.

Output is always PNG since python-docx can't embed SVG. Results are cached
to `~/.cache/marq/mermaid/{hash}.png`, keyed on everything that affects the
pixels, so a diagram is rendered once per distinct look. `marq` rather than
`docmarq`: pdfmarq builds the same key and both packages read each other's
entries. Writes go through a temp file and an atomic rename; a cached PNG
that fails to decode is re-rendered.
"""
import base64
import hashlib
import os
import shutil
import subprocess
import tempfile
import urllib.request
import urllib.error
from pathlib import Path

#-------------------------------------------------------------------------------------------- Cache

# Keyed on source + theme + background + scale, so a different scale is a
# different entry.
_CACHE_DIR = Path.home() / ".cache" / "marq" / "mermaid"

def _cache_key(source:str, theme:str, background:str, scale:float,
    font_family:str="", font_dir:str="", cli:str="") -> str:
  """SHA1 over every input that affects rendering. Different theme, bg,
  scale, font (family *and* directory - the same family name can resolve to
  a different TTF) or cli must produce a different cache file.

  pdfmarq builds this payload identically so both packages share entries;
  the field order is part of that contract.
  """
  payload = (f"{source}\x00{theme}\x00{background}\x00{scale}\x00"
    f"{font_family}\x00{font_dir}\x00{cli}").encode("utf-8")
  return hashlib.sha1(payload).hexdigest()[:16]

#----------------------------------------------------------------------------------------- Font CSS

def _resolve_font_ttf(font_dir:str, family:str) -> Path|None:
  """Find `<family>-Regular.ttf` under `font_dir`."""
  base = Path(font_dir)
  for sub in (family.lower(), family):
    p = base / sub / f"{family}-Regular.ttf"
    if p.is_file(): return p
  p = base / f"{family}-Regular.ttf"
  if p.is_file(): return p
  return None

def _mmdc_css_with_font(ttf_path:Path, family:str) -> str:
  """CSS for mmdc: @font-face from local TTF + apply to all SVG text."""
  return (
    f"@font-face {{\n"
    f"  font-family: '{family}';\n"
    f"  src: url('file:///{ttf_path.as_posix()}');\n"
    f"}}\n"
    f"* {{ font-family: '{family}', sans-serif !important; }}\n"
  )

def _cache_path(key:str) -> Path:
  _CACHE_DIR.mkdir(parents=True, exist_ok=True)
  return _CACHE_DIR / f"{key}.png"

def _tmp_path(out_path:Path) -> Path:
  """Unique sibling temp path. Rendering straight into the cache name would
  leave a truncated PNG behind on a crash or a killed subprocess."""
  return out_path.with_name(f"{out_path.stem}.{os.getpid()}.tmp{out_path.suffix}")

def _valid_png(path:Path) -> bool:
  """True when `path` holds a PNG that decodes. A half-written leftover has
  to be re-rendered rather than served as a cache hit for the life of the
  installation."""
  if not path.is_file() or path.stat().st_size == 0:
    return False
  try:
    from PIL import Image
    with Image.open(path) as im:
      im.verify()
    return True
  except ImportError:
    return True # no Pillow to check with; trust the file
  except Exception:
    return False

_REMOTE_WARNED = False

def _warn_remote_once() -> None:
  """Warn once per process before any diagram source leaves the machine."""
  global _REMOTE_WARNED
  if _REMOTE_WARNED:
    return
  _REMOTE_WARNED = True
  import warnings
  warnings.warn(
    "mermaid: local `mmdc` not found, falling back to the mermaid.ink "
    "public HTTP service - the diagram source is sent to a third-party "
    "server. Install mermaid-cli (npm i -g @mermaid-js/mermaid-cli) to "
    "render locally, or pass mermaid_remote=False to stay offline.",
    RuntimeWarning, stacklevel=3,
  )

#---------------------------------------------------------------------------------------------- API

def compile_to_png(source:str, cli:str="mmdc", theme:str="default",
    background:str="transparent", scale:float=3,
    timeout:float=60,
    font_family:str|None=None, font_dir:str|None=None,
    remote:bool=True) -> str|None:
  """Render mermaid `source` to a PNG file. Returns a path the caller can
  embed; cache hits return the cached file, cache misses try mmdc first
  then mermaid.ink. Returns `None` when no backend succeeded.

  When `font_family`+`font_dir` are set and a matching TTF is found, the
  diagram text uses that font via mmdc's `--cssFile` (`@font-face` CSS).
  Ignored by mermaid.ink (no custom-font support).

  `remote=False` blocks the mermaid.ink fallback, keeping the diagram
  source on this machine; rendering then fails instead and the caller
  emits a code block.

  The returned path lives in the cache - caller MUST NOT delete it.
  """
  key = _cache_key(source, theme, background, scale,
    font_family or "", font_dir or "", cli)
  cached = _cache_path(key)
  if _valid_png(cached):
    return str(cached)
  cached.unlink(missing_ok=True) # drop a corrupt leftover
  tmp = _tmp_path(cached)
  ok = _try_mmdc(source, tmp, cli=cli, theme=theme,
    background=background, scale=scale, timeout=timeout,
    font_family=font_family, font_dir=font_dir)
  if not ok and remote:
    ok = _try_mermaid_ink(source, tmp, theme=theme, background=background)
  if not ok:
    _safe_remove(tmp)
    return None
  try:
    os.replace(tmp, cached) # atomic: readers never see a partial file
  except OSError:
    _safe_remove(tmp)
    return None
  return str(cached)

#----------------------------------------------------------------------------- Backend: mermaid-cli

def _try_mmdc(source:str, out_path:Path, *, cli:str, theme:str,
    background:str, scale:float, timeout:float,
    font_family:str|None=None, font_dir:str|None=None) -> bool:
  """Local mmdc subprocess. Returns `True` on success.
  When `font_family`+`font_dir` are set and a matching TTF is found, a
  temp CSS file with `@font-face` is injected via `--cssFile`."""
  mmdc = shutil.which(cli)
  if mmdc is None: return False
  try:
    with tempfile.NamedTemporaryFile("w", suffix=".mmd", delete=False,
        encoding="utf-8") as f:
      f.write(source)
      src_path = f.name
  except OSError:
    return False
  css_path = None
  cmd = [mmdc, "-i", src_path, "-o", str(out_path),
    "-t", theme, "-b", background, "-s", str(scale)]
  # Puppeteer config (custom Chrome path, sandbox flags, etc.).
  # Both vars are checked: XAEIAN_MMDC_PUPPETEER_CONFIG is the pdfmarq name;
  # DOCMARQ_MMDC_PUPPETEER_CONFIG is the docmarq-specific override.
  pp_config = (os.environ.get("DOCMARQ_MMDC_PUPPETEER_CONFIG")
    or os.environ.get("XAEIAN_MMDC_PUPPETEER_CONFIG"))
  if pp_config and Path(pp_config).is_file():
    cmd += ["-p", pp_config]
  if font_family and font_dir:
    ttf = _resolve_font_ttf(font_dir, font_family)
    if ttf is not None:
      try:
        with tempfile.NamedTemporaryFile("w", suffix=".css", delete=False,
            encoding="utf-8") as f:
          f.write(_mmdc_css_with_font(ttf, font_family))
          css_path = f.name
        cmd += ["--cssFile", css_path]
      except OSError:
        css_path = None
  try:
    result = subprocess.run(cmd, capture_output=True, timeout=timeout, text=True)
  except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as e:
    if os.environ.get("DOCMARQ_DEBUG"):
      print(f"[docmarq.mermaid] mmdc exception: {e}")
    _safe_remove(src_path)
    if css_path: _safe_remove(css_path)
    return False
  _safe_remove(src_path)
  if css_path: _safe_remove(css_path)
  if result.returncode != 0 or not out_path.is_file():
    if os.environ.get("DOCMARQ_DEBUG"):
      tail = (result.stderr or result.stdout or "").strip().splitlines()[-3:]
      print(f"[docmarq.mermaid] mmdc rc={result.returncode}: " + " | ".join(tail))
    return False
  return True

#----------------------------------------------------------------------------- Backend: mermaid.ink

def _try_mermaid_ink(source:str, out_path:Path, *, theme:str,
    background:str) -> bool:
  """HTTP fallback via mermaid.ink `/img/` endpoint. Returns `True` on
  success. Uses base64-encoded source in the URL path (the API's preferred
  encoding for direct GET requests).
  """
  _warn_remote_once()
  try:
    encoded = base64.urlsafe_b64encode(source.encode("utf-8")).decode("ascii")
    # `bgColor` may be `transparent`, a hex value (`!RRGGBB`), or named.
    bg = background.lstrip("#")
    if bg.lower() != "transparent" and all(c in "0123456789abcdefABCDEF" for c in bg):
      bg = f"!{bg}"
    url = f"https://mermaid.ink/img/{encoded}?type=png&theme={theme}&bgColor={bg}"
    req = urllib.request.Request(url, headers={"User-Agent": "docmarq/0.1"})
    with urllib.request.urlopen(req, timeout=30) as resp:
      data = resp.read()
    if not data or not data.startswith(b"\x89PNG"):
      return False
    out_path.write_bytes(data)
    return True
  except (urllib.error.URLError, urllib.error.HTTPError, OSError, ValueError) as e:
    if os.environ.get("DOCMARQ_DEBUG"):
      print(f"[docmarq.mermaid] mermaid.ink failed: {e}")
    return False

#------------------------------------------------------------------------------------------ Helpers

def _safe_remove(path):
  """`os.remove` that swallows missing-file errors."""
  try:
    os.remove(path)
  except OSError:
    pass
