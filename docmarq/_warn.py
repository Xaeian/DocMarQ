# docmarq/_warn.py

"""Where a warning points: at the user's line, past every frame of the library."""
import sys

def caller_level() -> int:
  """
  `stacklevel` of the first frame outside `docmarq`, seen from the function that warns,
  so the warning points at the user's line however deep the library called.
  """
  frame, level = sys._getframe(2), 2
  while frame is not None and frame.f_globals.get("__name__", "").split(".")[0] == "docmarq":
    frame, level = frame.f_back, level + 1
  return level
