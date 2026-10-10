# tests/test_merge.py

"""`docmarq.merge`: templates filled with records, the checks of rows and the sandbox."""

import csv, io, warnings, zipfile
import pytest
pytest.importorskip("jinja2")
from docmarq.merge import fill, merge, MergeError
from docmarq.md import split_frontmatter
from conftest import docx_text

LETTER = ('---\ntitle: Contract\ndate: "{{ signed }}"\n'
  'to: ["{{ name }} {{ surname }}", "{{ city }}"]\n---\n\n'
  "Contract with **{{ name }} {{ surname }}**.\n")
ROW = {"name": "Jan", "surname": "Kowalski", "city": "Gdynia", "signed": "2026-01-12"}
# values markdown or the preprocessing of both libraries would read as syntax
VALUES = ["Kowalski (junior)", "[A]", "(hereinafter: Contractor)", "1) Jan", "1. May Street 3",
  "note :smile:", "a\\b", "x_y_z", "<!-- pagebreak -->", "$5 and $6", "==mark==", "~sub~",
  "^sup^", "# h", "> q", "- item", "+ item", "a | b", "**b**", "`c`", "Kowalska-Nowak",
  "Sp. z o.o.", "12:30", "[x](http://e.pl)", "100%"]

#---------------------------------------------------------------------------------------- Variables

def frontmatter_and_body_are_filled():
  fm, body = split_frontmatter(fill(LETTER, ROW))
  assert fm == {"title": "Contract", "date": "2026-01-12", "to": ["Jan Kowalski", "Gdynia"]}
  assert body.strip() == "Contract with **Jan Kowalski**."

def frontmatter_value_keeps_its_colon_and_quotes():
  text = fill('---\ntitle: "{{ company }}"\n---\nx', {"company": 'Company: „Alpha” "B"'})
  assert split_frontmatter(text)[0] == {"title": 'Company: „Alpha” "B"'}

def frontmatter_written_back_reads_the_same():
  text = fill('---\ntitle: "{{ t }}"\nid: "{{ i }}"\nflag: "{{ f }}"\n---\nx',
    {"t": "Contract: A", "i": "0012", "f": "true"})
  assert split_frontmatter(text)[0] == {"title": "Contract: A", "id": "0012", "flag": "true"}

def unquoted_placeholder_names_the_fix():
  with pytest.raises(MergeError, match=r"line 2: invalid YAML frontmatter.*goes in quotes"):
    fill("---\ndate: {{ signed }}\n---\nx", ROW)

def vars_give_defaults_and_the_record_wins():
  template = '---\nvars:\n  rate: "30,50"\n---\n{{ rate }}'
  assert fill(template) == "30,50"
  assert fill(template, {"rate": "40"}) == "40"

def empty_vars_entry_is_required():
  template = "---\nvars:\n  id_card:\n---\n{{ id_card }}"
  with pytest.raises(MergeError, match="`id_card` is required and empty"):
    fill(template, {"id_card": " "})
  assert fill(template, {"id_card": "ABC"}) == "ABC"

def unknown_field_names_the_line_of_the_file():
  with pytest.raises(MergeError, match="line 4: 'missing' is undefined"):
    fill("---\ntitle: T\n---\n{{ missing }}\n")

def impossible_date_is_reported_once():
  with pytest.raises(MergeError) as caught:
    fill("---\ndate: 2026-02-30\n---\nx")
  assert str(caught.value).count("invalid YAML frontmatter") == 1

def same_record_gives_the_same_text():
  assert fill(LETTER, ROW) == fill(LETTER, ROW)

#------------------------------------------------------------------------------------------- Values

def values_print_literally(tmp_path):
  from docmarq.paper import md_to_docx, PaperStyle
  template = "\n\n".join(f"{{{{ v[{i}] }}}}" for i in range(len(VALUES))) + "\n"
  path = str(tmp_path / "values.docx")
  md_to_docx(fill(template, {"v": VALUES}), path, style=PaperStyle(page_number_label=None))
  text = docx_text(path)
  assert [v for v in VALUES if v not in text] == []
  assert "oMath" not in zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")

def pipe_stays_in_its_cell_and_replace_keeps_the_row():
  text = fill("| A | B |\n|---|---|\n| {{ a }} | {{ b | replace('\\n', ', ') }} |\n",
    {"a": "x | y", "b": "Main Street 5\nGdynia"})
  assert text.splitlines()[-1] == "| x \\| y | Main Street 5, Gdynia |"

def newline_in_a_paragraph_is_a_hard_break():
  assert fill("{{ a }}", {"a": "Main Street 5\nGdynia"}) == "Main Street 5\\\nGdynia"

def md_filter_inserts_markdown_as_written():
  assert fill("{{ a | md }} {{ a }}", {"a": "**b**"}) == "**b** \\*\\*b\\*\\*"

def none_prints_nothing():
  assert fill("[{{ a }}]", {"a": None}) == "[]"
  text = fill('---\ntitle: "x{{ a }}"\n---\nb', {"a": None})
  assert split_frontmatter(text)[0] == {"title": "x"}

@pytest.mark.parametrize("value, shown", [("", "F"), (None, "F"), (False, "F"), ("yes", "T"),
  ("no", "T"), ("0", "T")])
def if_takes_any_text_but_empty_for_true(value, shown):
  assert fill("{% if f %}T{% else %}F{% endif %}", {"f": value}) == shown

def default_with_true_covers_an_empty_value():
  assert fill("{{ a | default('-') }}|{{ a | default('-', true) }}", {"a": ""}) == "|\\-"

def empty_cell_takes_the_default():
  assert fill('---\nvars:\n  s: "30,50"\n---\n{{ s }}', {"s": " "}) == "30,50"

def value_loses_its_outer_whitespace():
  # a final newline would print a backslash, four leading spaces would start a code block
  assert fill("x\n\n{{ a }}", {"a": "    code\n"}) == "x\n\ncode"

@pytest.mark.parametrize("template", ["{% set x %}{{ v }}{% endset %}{{ x }}",
  "{% macro m(a) %}{{ a }}{% endmacro %}{{ m(v) }}", "{% filter upper %}{{ v }}{% endfilter %}"])
def captured_text_is_escaped_once(template):
  assert fill(template, {"v": "Kowalska-Nowak"}).lower() == "kowalska\\-nowak"

def escape_filters_escape_for_markdown():
  assert fill("{{ v | e }} {{ v | escape }}", {"v": "a*b*"}) == "a\\*b\\* a\\*b\\*"
  assert fill("{{ v | safe }}", {"v": "**b**"}) == "**b**" # as `md`

#--------------------------------------------------------------------------------------- Whitespace

def tags_alone_vanish_and_a_tag_ending_text_keeps_the_newline():
  template = "a {% if x %}1{% endif %}\nb\n{% if x %}\nc\n{% endif %}\nd\n"
  assert fill(template, {"x": True}) == "a 1\nb\nc\nd\n"

def raw_text_stays_as_written():
  assert fill("{% raw %}{{ a }} {% x %}{% endraw %}\nb\n") == "{{ a }} {% x %}\nb\n"

def loop_builds_a_whole_table():
  template = "| A |\n|---|\n{% for r in rows %}\n| {{ r }} |\n{% endfor %}\n\nEnd.\n"
  assert fill(template, {"rows": [1, 2]}) == "| A |\n|---|\n| 1 |\n| 2 |\n\nEnd.\n"

def include_in_a_loop_keeps_clauses_apart(tmp_path):
  (tmp_path / "k").mkdir()
  (tmp_path / "k" / "a.md").write_text("Text A.\n", encoding="utf-8")
  (tmp_path / "k" / "b.md").write_text("Text B.\n", encoding="utf-8")
  template = '{% for k in ks %}\n## {{ k }}\n\n{% include "k/" ~ k ~ ".md" %}\n\n{% endfor %}\n'
  assert fill(template, {"ks": ["a", "b"]}, base_dir=tmp_path) \
    == "## a\n\nText A.\n\n## b\n\nText B.\n\n"

def line_with_a_tag_vanishes_only_when_it_prints_nothing():
  template = "| A |\n|---|\n{% if x %}| {{ x }} |{% endif %}\n| B |{# note #}\n| C |\n"
  assert fill(template, {"x": ""}) == "| A |\n|---|\n| B |\n| C |\n"
  assert fill(template, {"x": "1"}) == "| A |\n|---|\n| 1 |\n| B |\n| C |\n"

def tag_keeps_the_indent_of_its_line():
  assert fill("- a\n  {% if x %}- b{% endif %}\n", {"x": True}) == "- a\n  - b\n"

def body_opening_with_a_rule_keeps_it():
  text = fill("---\nvars:\n  a: 1\n---\n---\nNote: text\n---\nB\n")
  assert split_frontmatter(text) == ({}, "---\nNote: text\n---\nB\n")

#----------------------------------------------------------------------------------------- Includes

def include_stays_under_base_dir(tmp_path):
  (tmp_path / "secret.md").write_text("secret text", encoding="utf-8")
  (tmp_path / "base").mkdir()
  with pytest.raises(MergeError, match="cannot include ../secret.md"):
    fill('{% include "../secret.md" %}', base_dir=tmp_path / "base")
  with pytest.raises(MergeError, match="needs base_dir"):
    fill('{% include "secret.md" %}')

def include_of_itself_is_a_problem(tmp_path):
  (tmp_path / "self.md").write_text('{% include "self.md" %}', encoding="utf-8")
  with pytest.raises(MergeError, match="includes itself"):
    fill('{% include "self.md" %}', base_dir=tmp_path)

def error_in_an_included_file_names_it(tmp_path):
  (tmp_path / "k").mkdir()
  (tmp_path / "k" / "a.md").write_text("x\n{{ missing }}\n", encoding="utf-8")
  with pytest.raises(MergeError, match="k/a.md:2: 'missing' is undefined"):
    fill('{% include "k/a.md" %}', base_dir=tmp_path)

def syntax_error_in_an_include_names_the_file(tmp_path):
  (tmp_path / "bad.md").write_text("ok\n{% if %}", encoding="utf-8")
  with pytest.raises(MergeError, match="bad.md:2"):
    fill('{% include "bad.md" %}', base_dir=tmp_path)

def include_and_template_drop_a_byte_order_mark(tmp_path):
  (tmp_path / "k.md").write_bytes("\ufeff## Clause\n".encode("utf-8"))
  assert fill('{% include "k.md" %}\n', base_dir=tmp_path) == "## Clause\n"
  assert split_frontmatter(fill("\ufeff---\ntitle: T\n---\nB"))[0] == {"title": "T"}

def deep_nesting_is_a_merge_error():
  with pytest.raises(MergeError, match="nests too deep"):
    fill("{{ " + "(" * 3000 + "1" + ")" * 3000 + " }}")

#------------------------------------------------------------------------------------------ Sandbox

@pytest.mark.parametrize("expression", ["{{ ''.__class__ }}", "{{ 7 ** 1000 }}",
  "{{ 'x' * 200000 }}", "{{ [1] * 200000 }}", "{{ [1, 2] | random }}", "{{ lipsum() }}"])
def sandbox_refuses(expression):
  with pytest.raises(MergeError):
    fill(expression)

def sandbox_still_calculates():
  assert fill("{{ 2 ** 10 }} {{ 'ab' * 3 }}") == "1024 ababab"

def template_cannot_change_the_data():
  template = "---\nvars:\n  items: [a]\n---\n{% set _ = items.append(k) %}{{ items | length }}"
  with pytest.raises(MergeError, match="unsafe"):
    merge(template, [{"k": "1"}, {"k": "2"}])

@pytest.mark.parametrize("name", ["tojson", "urlize", "xmlattr"])
def html_and_json_filters_are_gone(name):
  with pytest.raises(MergeError, match="No filter"):
    fill("{{ v | " + name + " }}", {"v": "x"})

#------------------------------------------------------------------------------------------ Filters

def date_takes_iso_text_and_leaves_other_text():
  assert fill("{{ d | date }}", {"d": "2026-01-12"}, date_format="%d.%m.%Y") \
    == "12\\.01\\.2026"
  assert fill("{{ d | date('%Y') }}", {"d": "2026-01-12"}) == "2026"
  assert fill("{{ d | date }}", {"d": "12.01.2026"}, date_format="%Y") == "12\\.01\\.2026"
  assert fill("{{ d | date }}", {"d": "2026-01-12"}) == "2026\\-01\\-12"

@pytest.mark.parametrize("value, lang, shown", [
  ("1234,5", "pl", "1\u00a0234,50"), (1234.5, "de", "1.234,50"), ("1234.5", "en", "1,234.50"),
  ("1 234,5", None, "1234.50"), ("2,675", "pl", "2,68"),
  ("1.234.567", "pl", "1\u00a0234\u00a0567,00"),
  ("1,234.5", "pl", "1\u00a0234,50"), ("-0,004", "pl", "0,00"), ("", "pl", ""),
])
def number_reads_and_writes_by_language(value, lang, shown):
  assert fill("{{ v | number | md }}", {"v": value}, lang=lang) == shown

def number_keeps_every_digit_of_a_large_amount():
  assert fill("{{ v | number }}", {"v": "1234567890123456789012345678.91"}) \
    == "1234567890123456789012345678\\.91"
  with pytest.raises(MergeError, match="100 places at most"):
    fill("{{ 1 | number(10 ** 9) }}")

def number_takes_places():
  assert fill("{{ v | number(0) }}", {"v": "41,5"}) == "42"

def number_refuses_text_that_is_no_number():
  with pytest.raises(MergeError, match="is no number"):
    fill("{{ v | number }}", {"v": "abc"})

def int_and_float_are_strict():
  assert fill("{{ (v | float) * 2 }}", {"v": "30,5"}) == "61\\.0"
  assert fill("{{ vs | map('int') | sum }}", {"vs": ["40", "12"]}) == "52"
  with pytest.raises(MergeError, match="is no number"):
    fill("{{ v | int }}", {"v": "abc"})

def unknown_language_fails_up_front():
  with pytest.raises(ValueError, match="number format"):
    fill("x", lang="nl")

def number_of_nothing_prints_nothing():
  rows = [{"k": "a", "v": None}, {"k": "b", "v": "1"}]
  assert [d.text for d in merge("{{ k }}[{{ v | number }}]", rows)] == ["a[]", "b[1\\.00]"]

def huge_number_is_no_number():
  with pytest.raises(MergeError, match="is no number"):
    fill("{{ v | int }}", {"v": "1e100000"})

def int_and_float_keep_an_explicit_default_and_base():
  assert fill("{{ v | int(5) }} {{ v | float(0.5) }} {{ 'ff' | int(base=16) }}", {"v": "abc"}) \
    == "5 0\\.5 255"

#------------------------------------------------------------------------------- One record or many

def fill_takes_one_record_and_merge_a_list():
  with pytest.raises(TypeError, match="a list of records goes to merge"):
    fill("{{ a }}", [{"a": "1"}])
  with pytest.raises(TypeError, match="one record, a mapping, goes to fill"):
    merge("{{ a }}", {"a": "1"})

def versions_fill_from_one_record_with_their_own_vars():
  first = "---\nvars:\n  word: first\n---\n{{ word }} {{ a }}"
  second = "---\nvars:\n  word: second\n---\n{{ word }} {{ a }}"
  assert fill([first, second], {"a": "1"}) == ["first 1", "second 1"]
  assert merge([first, second], [{"a": "1"}])[0].text == ["first 1", "second 1"]

def problem_names_its_version():
  with pytest.raises(MergeError, match="version 2: line 1: 'b' is undefined"):
    fill(["{{ a }}", "{{ b }}"], {"a": "1"})
  with pytest.raises(MergeError) as caught:
    merge(["{{ a }}", "{{ b }}"], [{"a": "1"}])
  assert caught.value.problems == ["record 1 (1): version 2: line 1: 'b' is undefined"]

def table_as_one_field_gives_one_document():
  template = ("{% for p in people %}\n## {{ p.name }}\n\nDone {{ p.date }}.\n"
    "{% if not loop.last %}\n\n<!-- pagebreak -->\n{% endif %}\n\n{% endfor %}\n")
  rows = list(csv.DictReader(io.StringIO("name,date\nJan,12\nAnna,13\n")))
  assert fill(template, {"people": rows}) \
    == "## Jan\n\nDone 12.\n\n<!-- pagebreak -->\n\n## Anna\n\nDone 13.\n\n"

def column_named_like_a_dict_method_reads_the_column():
  rows = [{"items": "3", "values": "a", "update": "b"}]
  template = "{% for p in rows %}{{ p.items }} {{ p.values }} {{ p.update }}{% endfor %}"
  assert fill(template, {"rows": rows}) == "3 a b"
  assert merge("{{ order.items }}", [{"order": {"items": "5"}}])[0].text == "5"
  # without such a key, the method
  assert fill("{{ d.items() | list | length }}", {"d": {"a": "1"}}) == "1"

#--------------------------------------------------------------------------------------------- Rows

def rows_of_a_dict_reader_fill_as_they_are():
  # reading the file is the caller's, and `csv.DictReader` gives mappings of text as most do
  rows = list(csv.DictReader(io.StringIO("name;surname\nJan;Kowalski\nAnna;Nowak\n"),
    delimiter=";"))
  assert [d.text for d in merge("{{ name }} {{ surname }}", rows)] \
    == ["Jan Kowalski", "Anna Nowak"]

def row_with_more_cells_than_names_is_refused():
  rows = list(csv.DictReader(io.StringIO("name\nJan;x\n"), delimiter=";"))
  with pytest.raises(MergeError, match="more cells than the header has names"):
    merge("{{ name }}", rows)

def empty_record_is_refused():
  with pytest.raises(MergeError) as caught:
    merge("{{ name }}", [{"name": "Jan"}, {"name": ""}, {"name": None}])
  assert caught.value.problems == ["record 2 is empty", "record 3 is empty"]

def byte_order_mark_in_the_header_asks_for_utf_8_sig():
  rows = list(csv.DictReader(io.StringIO("\ufeffname\nJan\n")))
  with pytest.raises(MergeError, match="byte-order mark, read the file as utf-8-sig"):
    merge("{{ name }}", rows)

def record_is_a_mapping():
  with pytest.raises(MergeError, match="no mapping"):
    merge("x", [["Jan", "Kowalski"]])

#-------------------------------------------------------------------------------------------- Merge

def merge_names_files_and_keeps_the_records():
  rows = [{"name": "Jan", "surname": "Kowalski"}, {"name": "Anna", "surname": "Nowak"}]
  docs = merge("{{ name }}", rows, name="{{ surname }}-{{ name }}")
  assert [(d.name, d.text, d.data) for d in docs] \
    == [("Kowalski-Jan", "Jan", rows[0]), ("Nowak-Anna", "Anna", rows[1])]

def merge_numbers_files_by_default():
  rows = [{"a": str(i)} for i in range(12)]
  assert [d.name for d in merge("{{ a }}", rows)] == [f"{i:02}" for i in range(1, 13)]

def merge_lists_every_problem_at_once():
  rows = [{"a": "1", "s": "Kowalski"}, {"a": "", "s": "Nowak"}, {"a": " ", "s": "Lis"}]
  with pytest.raises(MergeError) as caught:
    merge("---\nvars:\n  a:\n---\n{{ a }}", rows, name="{{ s }}")
  assert "2 of 3 records cannot be filled" in str(caught.value)
  assert caught.value.problems == ["record 2 (Nowak): `a` is required and empty",
    "record 3 (Lis): `a` is required and empty"]

def merge_refuses_no_records():
  with pytest.raises(MergeError, match="no records"):
    merge("x", [])

@pytest.mark.parametrize("field", ["First name", "none", "loop", "self", "range", "n"])
def merge_refuses_field_names_jinja_cannot_read(field):
  with pytest.raises(MergeError, match="field"):
    merge("x", [{field: "1"}])

#--------------------------------------------------------------------------------------- File names

def file_names_suit_windows():
  rows = [{"a": 'x/y:z*?"<>|'}, {"a": "Sp. z o.o. "}, {"a": "CON"}, {"a": "CON.txt"},
    {"a": "a" * 150}]
  assert [d.name for d in merge("{{ a }}", rows, name="{{ a }}")] \
    == ["xyz", "Sp. z o.o", "CON_", "CON_.txt", "a" * 100]

def file_names_repeat_whatever_the_case():
  with pytest.raises(MergeError, match="repeats record 1"):
    merge("{{ a }}", [{"a": "Kowalski"}, {"a": "kowalski"}], name="{{ a }}")

#------------------------------------------------------------------------------------ Unused fields

def unused_field_warns():
  with pytest.warns(RuntimeWarning, match="fields no template reads: emial"):
    merge("{{ name }}", [{"name": "Jan", "emial": "x"}])
  with pytest.warns(RuntimeWarning, match="fields no template reads: emial"):
    fill("{{ name }}", {"name": "Jan", "emial": "x"})

def fields_read_anywhere_count_as_used(tmp_path):
  (tmp_path / "a.md").write_text("{{ address }}", encoding="utf-8")
  template = '---\ntitle: "{{ heading }}"\n---\n{{ name }}{% include "a.md" %}'
  row = {"name": "Jan", "heading": "T", "address": "Gdynia", "file": "x"}
  with warnings.catch_warnings():
    warnings.simplefilter("error")
    merge(template, [row], name="{{ file }}", base_dir=tmp_path)
    merge('{% include f %}', [{"f": "a.md", "address": "x", "other": "y"}], base_dir=tmp_path)

#------------------------------------------------------------------------------------------- Render

def filled_contract_prints_its_record_literally(tmp_path):
  from docmarq.legal import md_to_docx, lang_style
  template = ('---\ntitle: Contract\nto: ["{{ company }}"]\n---\n\n| Task | Hours |\n|---|--:|\n'
    "{% for t in tasks %}\n| {{ t.name }} | {{ t.hours }} |\n{% endfor %}\n")
  path = str(tmp_path / "c.docx")
  record = {"company": "Alpha (Sp. z o.o.)", "tasks": [{"name": "A | B", "hours": 4}]}
  md_to_docx(fill(template, record), path, style=lang_style("pl"))
  text = docx_text(path)
  assert "Alpha (Sp. z o.o.)" in text and "A | B" in text
