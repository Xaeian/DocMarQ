# tests/test_units.py

"""Pure helpers - units, colors, margins, size ladders, align. Cheap and fast, no docx."""

import pytest
from docmarq.utils import (
  to_mm, mm_to_emu, pt_to_emu, parse_color, rgb255, color_hex,
  parse_margin, smaller_size, tight_line_height,
)
from docmarq.constants import Align

#--------------------------------------------------------------------------------------- Units

@pytest.mark.parametrize("value, unit, expected", [
  (10, "mm", 10),
  (1, "cm", 10),
  (1, "in", 25.4),
])
def to_mm_converts_known_units(value, unit, expected):
  assert to_mm(value, unit) == pytest.approx(expected)

def to_mm_rejects_unknown_unit():
  with pytest.raises(ValueError):
    to_mm(1, "furlong")

def mm_to_emu_inch_is_914400():
  # 1 inch = 25.4 mm = 914400 EMU
  assert mm_to_emu(25.4) == 914400

def pt_to_emu_point_is_12700():
  # 1 pt = 12700 EMU
  assert pt_to_emu(1) == 12700

#-------------------------------------------------------------------------------------- Colors

def parse_color_returns_floats():
  # regression: previously returned ints 0-255. Canonical form is now floats
  # 0-1 (matching `pdfmarq.parse_color`). Conversion to ints happens at the
  # OOXML boundary via `rgb255`.
  r, g, b = parse_color("#FF0000")
  assert (r, g, b) == pytest.approx((1.0, 0.0, 0.0))
  assert all(isinstance(c, float) for c in (r, g, b))

@pytest.mark.parametrize("value, expected", [
  ("#FF0000", (1.0, 0.0, 0.0)),
  ("#F00", (1.0, 0.0, 0.0)),           # short hex
  ((0.2, 0.4, 0.8), (0.2, 0.4, 0.8)),  # tuple passthrough
  (None, (0.0, 0.0, 0.0)),
])
def parse_color_known_forms(value, expected):
  assert parse_color(value) == pytest.approx(expected)

@pytest.mark.parametrize("bad", ["#GGGGGG", "#12345"])
def parse_color_rejects_invalid_hex(bad):
  with pytest.raises(ValueError):
    parse_color(bad)

def rgb255_from_floats():
  assert rgb255((1.0, 0.0, 0.5)) == (255, 0, 128)

def rgb255_from_hex():
  assert rgb255("#1f2328") == (0x1F, 0x23, 0x28)

@pytest.mark.parametrize("value, expected", [
  ((1.0, 0.0, 0.0), "FF0000"),
  ((0.5, 0.5, 0.5), "808080"),
  ("#1f2328", "1F2328"),  # round-trip: hex in, hex out (uppercase)
])
def color_hex_normalizes_to_uppercase(value, expected):
  assert color_hex(value) == expected

#-------------------------------------------------------------------------------------- Margin

@pytest.mark.parametrize("value, expected", [
  (10, (10, 10, 10, 10)),            # scalar
  ((10, 20), (10, 20, 10, 20)),      # (v, h)
  ((10, 20, 30), (10, 20, 30, 20)),  # (t, h, b)
  ((1, 2, 3, 4), (1, 2, 3, 4)),      # full CSS form
])
def parse_margin_css_forms(value, expected):
  assert parse_margin(value) == expected

#------------------------------------------------------------------------------------- Helpers

@pytest.mark.parametrize("body, expected", [
  (11, 10),
  (12, 11),
  (14, 12),
  (16, 14),
])
def smaller_size_steps_down_ladder(body, expected):
  # per typographic ladder
  assert smaller_size(body) == expected

def smaller_size_clamps_to_floor():
  # body smaller than the floor returns the floor
  assert smaller_size(7, min_pt=7) == 7

@pytest.mark.parametrize("size, expected", [
  (11, 1.15),  # body
  (24, 1.0),   # display
])
def tight_line_height_by_size(size, expected):
  assert tight_line_height(size) == expected

#-------------------------------------------------------------------------------------- Align

@pytest.mark.parametrize("attr, code", [
  ("LEFT", "L"),
  ("RIGHT", "R"),
  ("CENTER", "C"),
  ("JUSTIFY", "J"),
])
def align_constants_match_pdfmarq(attr, code):
  # cross-lib reuse: both libs share these single-char codes
  assert getattr(Align, attr) == code
