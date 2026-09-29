# tests/test_table_widths.py

"""
Table columns are sized by content and fill the content area.

A short code keeps its column narrow and whole.
A long URL never squeezes the columns beside it.
"""

import pytest
from reportlab.pdfbase.pdfmetrics import stringWidth
from conftest import docx_grid_mm
from docmarq import DOCX, TableStyle
from docmarq.constants import Unit, Defaults
from docmarq.utils import smaller_size
from docmarq.md import md_to_docx
from docmarq.metrics import text_width_mm
from docmarq.tables import _fit_columns, _cap_widths

LONG = "opis ktory jest na tyle dlugi ze musi sie zawinac w komorce tabeli " * 2
URL = "https://example.com/very/long/path/to/some/resource/file.pdf"
PAD = TableStyle().cell_pad_h # a side
SIZE = smaller_size(Defaults.FONT_SIZE) # table text, one ladder step below the body
CONTENT = 170 # A4 less its 20 mm margins

@pytest.fixture
def grid(tmp_path):
  """Grid widths of the first table `md` renders to."""
  def render(md:str) -> list[float]:
    path = str(tmp_path / "table.docx")
    md_to_docx(md, path)
    return docx_grid_mm(path)[0]
  return render

#------------------------------------------------------------------------------------------ Content

def short_code_column_stays_narrow(grid):
  cols = grid(f"| Kod | Opis |\n|---|---|\n| PP-1 | {LONG} |\n")
  assert cols[0] < cols[1]
  assert cols[0] >= text_width_mm("PP-1", "Calibri", SIZE) + 2 * PAD
  assert sum(cols) == pytest.approx(CONTENT, abs=0.1)

def long_url_leaves_short_columns_whole(grid):
  cols = grid(f"| A | B | C | D |\n|---|---|---|---|\n| PP-1 | {URL} | {URL} | OK |\n")
  assert cols[0] >= text_width_mm("PP-1", "Calibri", SIZE) + 2 * PAD
  assert cols[3] >= text_width_mm("OK", "Calibri", SIZE) + 2 * PAD
  assert sum(cols) == pytest.approx(CONTENT, abs=0.1)

def explicit_widths_win(tmp_path):
  path = str(tmp_path / "fixed.docx")
  with DOCX(path) as doc:
    doc.table([["PP-1", LONG]], widths=[60, 40])
  assert docx_grid_mm(path)[0] == pytest.approx([60, 40], abs=0.1)

#------------------------------------------------------------------------------------------ Metrics

def unknown_family_measures_as_helvetica():
  expect = stringWidth("PP-1", "Helvetica", SIZE) * Unit.PT
  assert text_width_mm("PP-1", "NoSuchFamily", SIZE) == pytest.approx(expect)

def bold_measures_wider():
  regular = text_width_mm("Kolumna", "NoSuchFamily", SIZE)
  assert text_width_mm("Kolumna", "NoSuchFamily", SIZE, bold=True) > regular

#------------------------------------------------------------------------------------------- Solver

def spare_room_goes_to_growing_columns():
  assert _fit_columns([5, 5], [10, 20], 40, [1]) == pytest.approx([10, 30])
  assert _fit_columns([5, 5], [10, 20], 40, []) == pytest.approx([15, 25])

def squeeze_shares_out_the_gaps():
  # each column gives up the same share of its max - min gap
  assert _fit_columns([10, 10], [20, 60], 50, [0, 1]) == pytest.approx([15, 35])

def cap_cuts_only_the_widest():
  assert _cap_widths([10, 12, 150], 100) == pytest.approx([10, 12, 78])
  assert _cap_widths([50, 60, 70], 90) == pytest.approx([30, 30, 30])
