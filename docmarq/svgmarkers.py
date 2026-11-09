# docmarq/svgmarkers.py

"""
Draw SVG `<marker>` arrowheads as ordinary shapes.

svglib drops `marker-start`, `marker-mid` and `marker-end` without a word,
so every arrow reaches the PDF as a bare line.
`expand` places a copy of each marker at its vertices, turned along the path as a browser does.

Not covered:
- marker properties set by a `<style>` sheet: only attributes and inline `style` are read
- the clip at the marker's own viewport
- shapes drawn through `<use>` or picked by `<switch>`

Example:
  >>> root = ElementTree.fromstring(svg_text)
  >>> expand(root)
"""

import copy, math, re

#---------------------------------------------------------------------------------------- Constants

_NUM = r"[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?"
_NUM_RE = re.compile(_NUM)
_URL_RE = re.compile(r"url\(\s*['\"]?#([^'\")]+)['\"]?\s*\)")

# Argument kinds per path command: `n` a number, `f` an arc flag.
# A flag is one digit, which minified data runs straight into the next number: `a5 5 0 015 5`.
_ARGS = {
  "M": "nn", "L": "nn", "H": "n", "V": "n",
  "C": "nnnnnn", "S": "nnnn", "Q": "nnnn", "T": "nn",
  "A": "nnnffnn",
}
_SCAN_RE = {
  "gap": re.compile(r"[\s,]*"),
  "letter": re.compile(r"[MLHVCSQTAZmlhvcsqtaz]"),
  "n": re.compile(_NUM),
  "f": re.compile(r"[01]"),
}

_MARKED = ("line", "path", "polyline", "polygon")
# Content under these is never drawn where it stands, so its markers never show.
_HIDDEN = ("defs", "marker", "symbol", "clipPath", "mask", "pattern", "switch")

# Inherited paint, at its initial values.
# Marker content inherits from the marker's ancestors, never from the shape that uses it.
# The copy sits beside the shape, so where the shape's ancestors set one of these,
# the group puts the marker's own value back.
_INITIAL = {
  "fill": "#000000", "fill-opacity": "1", "fill-rule": "nonzero",
  "stroke": "none", "stroke-width": "1", "stroke-opacity": "1",
  "stroke-dasharray": "none", "stroke-dashoffset": "0",
  "stroke-linecap": "butt", "stroke-linejoin": "miter", "stroke-miterlimit": "4",
}
_MARKER_PROPS = ("marker-start", "marker-mid", "marker-end")
_INHERITED = (*_INITIAL, *_MARKER_PROPS)

Point = tuple[float, float]
Dir = tuple[float, float]|None

#------------------------------------------------------------------------------------------- Styles

def _tag(node) -> str:
  return node.tag.split("}")[-1] if isinstance(node.tag, str) else ""

def _style(node) -> dict:
  """Declarations of the inline `style` of `node`."""
  out = {}
  for rule in (node.get("style") or "").split(";"):
    key, _, value = rule.partition(":")
    if value.strip(): out[key.strip()] = value.strip()
  return out

def _undrawn(node) -> bool:
  """`display: none` by attribute or inline style: nothing under `node` reaches the page."""
  return _style(node).get("display", node.get("display") or "").strip() == "none"

def _declared(node) -> dict:
  """Properties `node` sets itself: attributes, then inline `style` over them."""
  own = {key: node.get(key) for key in (*_INHERITED, "marker") if node.get(key)}
  own.update({key: value for key, value in _style(node).items()
    if key in _INHERITED or key == "marker"})
  shorthand = own.pop("marker", None) # sets all three, under any longhand given
  if shorthand:
    own = {**{key: shorthand for key in _MARKER_PROPS}, **own}
  return own

def _context(chain:list) -> dict:
  """Inherited properties in force at the end of `chain`, root first."""
  props: dict = {}
  for node in chain:
    props.update(_declared(node))
  return props

def _number(text:str|None, default:float=0.0) -> float:
  """Leading number of an attribute, `default` when there is none."""
  found = _NUM_RE.search(text or "")
  return float(found.group(0)) if found else default

def _degrees(text:str) -> float:
  """An `orient` angle, `grad` tested before the `rad` it ends with."""
  value = _number(text)
  for unit, scale in (("grad", 0.9), ("rad", 180 / math.pi), ("turn", 360.0)):
    if text.endswith(unit): return value * scale
  return value

#--------------------------------------------------------------------------------------- Directions

def _way(*points:Point) -> Dir:
  """Direction from the first point to the first later one that differs."""
  x0, y0 = points[0]
  for x, y in points[1:]:
    if (x, y) != (x0, y0): return x - x0, y - y0
  return None

def _back(*points:Point) -> Dir:
  """Direction arriving at the first point, from the nearest later one that differs."""
  way = _way(*points)
  return (-way[0], -way[1]) if way else None

def _arc_ways(
  p0:Point, rx:float, ry:float, phi:float,
  large:bool, sweep:bool, p1:Point,
) -> tuple[Dir, Dir]:
  """
  Tangents leaving `p0` and arriving at `p1` along an elliptical arc.
  Through the centre form of SVG implementation notes F.6.5.
  """
  rx, ry = abs(rx), abs(ry)
  if p0 == p1: return None, None # the spec omits the arc
  if not rx or not ry: return _way(p0, p1), _way(p0, p1) # the spec draws a line
  cos, sin = math.cos(math.radians(phi)), math.sin(math.radians(phi))
  dx, dy = (p0[0] - p1[0]) / 2, (p0[1] - p1[1]) / 2
  x1, y1 = cos * dx + sin * dy, -sin * dx + cos * dy
  reach = (x1 / rx) ** 2 + (y1 / ry) ** 2
  if reach > 1: # radii too small to span the ends: grow them just enough
    rx, ry = rx * math.sqrt(reach), ry * math.sqrt(reach)
  num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1
  den = rx * rx * y1 * y1 + ry * ry * x1 * x1
  k = math.sqrt(max(num, 0.0) / den) if den else 0.0
  if large == sweep: k = -k
  cx, cy = k * rx * y1 / ry, -k * ry * x1 / rx
  turn = 1 if sweep else -1
  def tangent(ux:float, uy:float) -> Dir:
    t = math.atan2((uy - cy) / ry, (ux - cx) / rx)
    tx, ty = -rx * math.sin(t) * turn, ry * math.cos(t) * turn
    return cos * tx - sin * ty, sin * tx + cos * ty
  return tangent(x1, y1), tangent(-x1, -y1)

#----------------------------------------------------------------------------------------- Vertices

class Vertex:
  """Where a marker may sit: the point, and the path's direction into and out of it."""
  __slots__ = ("x", "y", "into", "out")

  def __init__(self, x:float, y:float, into:Dir=None):
    self.x, self.y, self.into, self.out = x, y, into, None

  def angle(self) -> float:
    """`orient="auto"`, in degrees: along the path, halfway round a corner."""
    ways = [math.atan2(d[1], d[0]) for d in (self.into, self.out) if d]
    if len(ways) < 2: return math.degrees(ways[0]) if ways else 0.0
    a, b = ways
    mid = (a + b) / 2
    if abs(b - a) > math.pi: mid += math.pi
    return math.degrees(mid)

class _Scan:
  """Cursor over path data, reading one letter, number or flag at a time."""
  def __init__(self, text:str):
    self.text, self.at = text, 0

  def read(self, kind:str) -> str|None:
    self.at = _SCAN_RE["gap"].match(self.text, self.at).end()
    found = _SCAN_RE[kind].match(self.text, self.at)
    if not found: return None
    self.at = found.end()
    return found.group(0)

  def done(self) -> bool:
    self.at = _SCAN_RE["gap"].match(self.text, self.at).end()
    return self.at >= len(self.text)

def path_vertices(data:str) -> list[Vertex]:
  """Every vertex of SVG path data with its directions, `[]` for data it cannot read."""
  scan = _Scan(data)
  out: list[Vertex] = []
  x = y = 0.0
  start = 0 # the open subpath's first vertex
  curve: tuple[str, Point]|None = None # last `C`/`Q` control point, for `S`/`T`
  command = ""

  def segment(end:Point, into:Dir, leave:Dir) -> None:
    nonlocal x, y
    if out: out[-1].out = leave
    out.append(Vertex(*end, into))
    x, y = end

  def reflect(kind:str) -> Point:
    if curve and curve[0] == kind: return 2 * x - curve[1][0], 2 * y - curve[1][1]
    return x, y

  while not scan.done():
    command = scan.read("letter") or command
    if not command: return []
    op, rel = command.upper(), command.islower()
    if op == "Z":
      if out:
        first = out[start]
        way = _way((x, y), (first.x, first.y)) or out[-1].into
        segment((first.x, first.y), way, way)
        first.into, out[-1].out = way, first.out
      curve, command = None, ""
      continue
    nums = [scan.read(kind) for kind in _ARGS[op]]
    if None in nums: return []
    v = [float(n) for n in nums]
    ox, oy = (x, y) if rel else (0.0, 0.0)
    p0 = (x, y)
    if op == "M":
      x, y = v[0] + ox, v[1] + oy
      out.append(Vertex(x, y))
      start, curve = len(out) - 1, None
      command = "l" if rel else "L" # further pairs are lines
    elif op in ("L", "H", "V"):
      if op == "L": end = (v[0] + ox, v[1] + oy)
      elif op == "H": end = (v[0] + ox, y)
      else: end = (x, v[0] + oy)
      segment(end, _way(p0, end), _way(p0, end))
      curve = None
    elif op in ("C", "S"):
      c1 = (v.pop(0) + ox, v.pop(0) + oy) if op == "C" else reflect("C")
      c2, end = (v[0] + ox, v[1] + oy), (v[2] + ox, v[3] + oy)
      segment(end, _back(end, c2, c1, p0), _way(p0, c1, c2, end))
      curve = ("C", c2)
    elif op in ("Q", "T"):
      c = (v.pop(0) + ox, v.pop(0) + oy) if op == "Q" else reflect("Q")
      end = (v[0] + ox, v[1] + oy)
      segment(end, _back(end, c, p0), _way(p0, c, end))
      curve = ("Q", c)
    else: # `A`
      end = (v[5] + ox, v[6] + oy)
      leave, into = _arc_ways(p0, v[0], v[1], v[2], bool(v[3]), bool(v[4]), end)
      segment(end, into, leave)
      curve = None
  return out

def shape_vertices(node) -> list[Vertex]:
  """Vertices of a `line`, `path`, `polyline` or `polygon`."""
  name = _tag(node)
  if name == "path": return path_vertices(node.get("d") or "")
  if name == "line":
    points = [(_number(node.get("x1")), _number(node.get("y1"))),
      (_number(node.get("x2")), _number(node.get("y2")))]
  else:
    nums = [float(n) for n in _NUM_RE.findall(node.get("points") or "")]
    points = list(zip(nums[0::2], nums[1::2]))
  if len(points) < 2: return []
  data = "M" + " L".join(f"{px!r} {py!r}" for px, py in points)
  return path_vertices(data + (" Z" if name == "polygon" else ""))

#------------------------------------------------------------------------------------------- Expand

def _fmt(value:float) -> str:
  return f"{value:.4f}".rstrip("0").rstrip(".") or "0"

def _placement(marker, vertex:Vertex, stroke_width:float, at_start:bool) -> str:
  """`matrix(...)` taking marker content to `vertex`: turned, scaled, `refX`/`refY` on it."""
  orient = (marker.get("orient") or "0").strip()
  if orient == "auto": angle = vertex.angle()
  elif orient == "auto-start-reverse": angle = vertex.angle() + (180 if at_start else 0)
  else: angle = _degrees(orient)
  size = stroke_width if (marker.get("markerUnits") or "strokeWidth") == "strokeWidth" else 1
  sx = sy = 1.0
  box = [float(n) for n in _NUM_RE.findall(marker.get("viewBox") or "")]
  if len(box) == 4 and box[2] > 0 and box[3] > 0:
    sx = _number(marker.get("markerWidth"), 3) / box[2]
    sy = _number(marker.get("markerHeight"), 3) / box[3]
    fit = (marker.get("preserveAspectRatio") or "").split()
    if not fit or fit[0] != "none":
      sx = sy = max(sx, sy) if "slice" in fit else min(sx, sy)
  # The viewBox origin and alignment cancel out: `refX`/`refY` is pinned to the vertex.
  cos, sin = math.cos(math.radians(angle)), math.sin(math.radians(angle))
  a, b, c, d = cos * sx * size, sin * sx * size, -sin * sy * size, cos * sy * size
  rx, ry = _number(marker.get("refX")), _number(marker.get("refY"))
  e, f = vertex.x - (a * rx + c * ry), vertex.y - (b * rx + d * ry)
  return f"matrix({' '.join(_fmt(n) for n in (a, b, c, d, e, f))})"

def _repaint(node, paint:dict) -> None:
  """Resolve SVG 2 `context-stroke`/`context-fill` to the shape's own paint."""
  for part in node.iter():
    for key in ("fill", "stroke"):
      value = part.get(key)
      if value in ("context-stroke", "context-fill"):
        part.set(key, paint[value[len("context-"):]])
    style = part.get("style")
    if style and "context-" in style:
      part.set("style", style.replace("context-stroke", paint["stroke"])
        .replace("context-fill", paint["fill"]))

def _copy(marker, attrs:dict, paint:dict):
  """Group holding a copy of `marker`'s content under `attrs`."""
  group = marker.makeelement(marker.tag[:-len("marker")] + "g", attrs)
  for child in marker:
    if not _tag(child): continue
    part = copy.deepcopy(child)
    for node in part.iter():
      if _tag(node): node.attrib.pop("id", None) # a copy must not answer for the original
    group.append(part)
  _repaint(group, paint)
  return group

def _shapes(node, inherited:dict, found:list) -> None:
  """
  Collect `(parent, index, shape, inherited)` for every drawn shape that may carry markers.
  `inherited` is what the ancestors of `node` set, so each level adds its own once.
  """
  inherited = {**inherited, **_declared(node)}
  for at, child in enumerate(node):
    name = _tag(child)
    if not name or name in _HIDDEN or _undrawn(child): continue
    if name in _MARKED:
      found.append((node, at, child, inherited))
    elif len(child):
      _shapes(child, inherited, found)

def expand(root) -> int:
  """
  Add each shape's markers to `root`, as groups right after the shape.
  Works in place, on ElementTree and lxml trees alike. Returns how many groups it added.
  """
  markers = {m.get("id"): m for m in root.iter() if _tag(m) == "marker" and m.get("id")}
  if not markers: return 0
  parents = {child: parent for parent in root.iter() for child in parent}
  def ancestry(node) -> list:
    chain = []
    while node is not None:
      chain.append(node)
      node = parents.get(node)
    return chain[::-1]
  contexts = {key: _context(ancestry(marker)) for key, marker in markers.items()}
  found: list = []
  _shapes(root, {}, found)
  count = 0
  # Last first: groups inserted after a shape leave the index of every shape before it valid.
  for parent, at, shape, inherited in reversed(found):
    props = {**inherited, **_declared(shape)}
    wanted = {key: _URL_RE.search(props.get(key) or "") for key in _MARKER_PROPS}
    if not any(wanted.values()): continue
    vertices = shape_vertices(shape)
    if not vertices: continue
    paint = {key: props.get(key, _INITIAL[key]) for key in ("stroke", "fill")}
    stroke_width = _number(props.get("stroke-width"), 1)
    last = len(vertices) - 1
    spots = {"marker-start": [0], "marker-mid": range(1, last), "marker-end": [last]}
    groups = []
    for key in _MARKER_PROPS:
      marker_id = wanted[key].group(1) if wanted[key] else None
      marker = markers.get(marker_id)
      if marker is None: continue
      own = contexts[marker_id]
      reset = {prop: own.get(prop, _INITIAL[prop]) for prop in _INITIAL
        if prop in own or prop in inherited}
      if shape.get("opacity"): reset["opacity"] = shape.get("opacity")
      for spot in spots[key]:
        placement = _placement(marker, vertices[spot], stroke_width, key == "marker-start")
        if shape.get("transform"): placement = f"{shape.get('transform')} {placement}"
        groups.append(_copy(marker, {**reset, "transform": placement}, paint))
    for group in reversed(groups):
      parent.insert(at + 1, group)
    count += len(groups)
  return count
