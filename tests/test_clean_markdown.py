"""Tests for clean_page_markdown post-processing."""

from doc2md.transcribe import clean_page_markdown


def test_html_labelled_fence_unwrapped():
    md = 'Before\n\n```html\n<table>\n<tr><td>A</td></tr>\n</table>\n```\n\nAfter'
    out = clean_page_markdown(md)
    assert "```" not in out
    assert "<table>\n<tr><td>A</td></tr>\n</table>" in out
    assert out.startswith("Before") and out.endswith("After")


def test_unlabelled_fence_around_table_unwrapped():
    md = "```\n<table><tr><td>1</td></tr></table>\n```"
    out = clean_page_markdown(md)
    assert out == "<table><tr><td>1</td></tr></table>"


def test_other_code_fences_untouched():
    md = "```python\nprint('hi')\n```"
    assert clean_page_markdown(md) == md


def test_unlabelled_fence_without_table_untouched():
    md = "```\nNode A -> Node B\n```"
    assert clean_page_markdown(md) == md


def test_mermaid_block_structure_untouched():
    md = (
        "```mermaid\n"
        "flowchart LR\n"
        '    A["Input"] --> B["Output"]\n'
        "```"
    )
    assert clean_page_markdown(md) == md


def test_tilde_stripped_from_pie_and_xychart_values():
    pie = '```mermaid\npie title Share\n    "A" : ~42.5\n    "B" : 57.5\n```'
    out = clean_page_markdown(pie)
    assert '"A" : 42.5' in out and "~" not in out

    xy = "```mermaid\nxychart-beta\n    bar [~45, 62, ~58]\n```"
    out = clean_page_markdown(xy)
    assert "bar [45, 62, 58]" in out


def test_thousands_comma_stripped_in_pie_only():
    pie = '```mermaid\npie title Sales\n    "A" : 1,234\n```'
    assert '"A" : 1234' in clean_page_markdown(pie)

    # In xychart arrays commas delimit values and must never be collapsed.
    xy = "```mermaid\nxychart-beta\n    bar [45,620,58]\n```"
    assert "bar [45,620,58]" in clean_page_markdown(xy)


def test_tilde_untouched_outside_mermaid_and_in_gantt():
    md = "Values approx ~45 units.\n```mermaid\ngantt\n    dateFormat YYYY-MM-DD\n```"
    assert clean_page_markdown(md) == md


def test_plain_markdown_passthrough():
    md = "# Title\n\nSome text with a <table><tr><td>x</td></tr></table> inline."
    assert clean_page_markdown(md) == md


def test_multiple_html_fences_all_unwrapped():
    md = (
        "```html\n<table><tr><td>1</td></tr></table>\n```\n\n"
        "Text between.\n\n"
        "```HTML\n<table><tr><td>2</td></tr></table>\n```"
    )
    out = clean_page_markdown(md)
    assert "```" not in out
    assert "<td>1</td>" in out and "<td>2</td>" in out
