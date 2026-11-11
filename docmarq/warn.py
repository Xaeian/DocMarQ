# docmarq/warn.py

"""Where a warning points: at the user's line, past every frame of the library."""
import sys

_LIBRARY = ("docmarq", "contextlib") # modules whose frames a warning looks past

def caller_level() -> int:
  """
  `stacklevel` of the first frame outside `docmarq`, seen from the function that warns,
  so the warning points at the user's line however deep the library called.
  `contextlib` counts as the library: a `with` block it closes warns at the user's `with`.
  """
  frame, level = sys._getframe(2), 2
  while frame is not None and frame.f_globals.get("__name__", "").split(".")[0] in _LIBRARY:
    frame, level = frame.f_back, level + 1
  return level
