# tests/conftest.py

"""With `python_functions = ["*"]`, collect only functions defined in the test module; plus the shared DOCX validator."""

import inspect
from pathlib import Path

def pytest_pycollect_makeitem(collector, name, obj):
  if inspect.isfunction(obj) and obj.__module__ != collector.obj.__name__:
    return [] # ignore library functions imported into the test file

#----------------------------------------------------------------------------------- Assertions

def assert_valid_docx(path:str|Path, min_size:int=2000):
  """Validate that `path` points to a real-looking DOCX file (ZIP magic `PK\\x03\\x04`)."""
  p = Path(path)
  assert p.exists(), f"DOCX not created: {p}"
  size = p.stat().st_size
  assert size >= min_size, f"DOCX suspiciously small ({size} bytes): {p}"
  with p.open("rb") as f:
    head = f.read(4)
  assert head == b"PK\x03\x04", f"Not a DOCX/ZIP (header={head!r}): {p}"
