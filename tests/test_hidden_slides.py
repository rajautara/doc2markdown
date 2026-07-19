"""Tests for hidden-slide page planning and labelled PDF rendering.

Usage:
    .venv\\Scripts\\python tests\\test_hidden_slides.py

No Microsoft Office required — the COM code paths are not exercised here.
"""

import sys
import tempfile
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from doc2md.converter import (  # noqa: E402
    ConversionError,
    plan_ppt_pages,
    pdf_to_images,
    prepare_images,
)


def test_case_a_hidden_included_in_pdf():
    # PDF has all 5 slides; hidden ones are filtered out before rendering.
    render_map, skipped = plan_ppt_pages(5, {2, 4}, 5, None, True)
    assert render_map == {1: 1, 3: 3, 5: 5}, render_map
    assert skipped == [2, 4], skipped


def test_case_a_requested_subset():
    render_map, skipped = plan_ppt_pages(5, {2, 4}, 5, {2, 3, 4, 5}, True)
    assert render_map == {3: 3, 5: 5}, render_map
    assert skipped == [2, 4], skipped


def test_case_a_skip_disabled():
    render_map, skipped = plan_ppt_pages(5, {2, 4}, 5, None, False)
    assert render_map == {1: 1, 2: 2, 3: 3, 4: 4, 5: 5}, render_map
    assert skipped == [], skipped


def test_case_b_hidden_excluded_from_pdf():
    # PDF only contains the 3 visible slides; they compact onto pages 1-3
    # while labels keep the original slide numbers.
    render_map, skipped = plan_ppt_pages(5, {2, 4}, 3, None, True)
    assert render_map == {1: 1, 2: 3, 3: 5}, render_map
    assert skipped == [2, 4], skipped


def test_case_b_skip_disabled():
    # Hidden slides are not in the PDF, so they cannot be included even when
    # skipping is disabled; they are reported as skipped instead.
    render_map, skipped = plan_ppt_pages(5, {2, 4}, 3, None, False)
    assert render_map == {1: 1, 2: 3, 3: 5}, render_map
    assert skipped == [2, 4], skipped


def test_out_of_range_request():
    try:
        plan_ppt_pages(5, {2}, 5, {4, 9}, True)
    except ConversionError:
        return
    raise AssertionError("expected ConversionError for out-of-range page")


def test_all_requested_slides_hidden():
    render_map, skipped = plan_ppt_pages(2, {1, 2}, 2, None, True)
    assert render_map == {}, render_map
    assert skipped == [1, 2], skipped


def test_no_hidden_slides():
    render_map, skipped = plan_ppt_pages(3, set(), 3, None, True)
    assert render_map == {1: 1, 2: 2, 3: 3}, render_map
    assert skipped == [], skipped


def test_fallback_unexpected_page_count():
    # Weird export (page count matches neither case): identity within range.
    render_map, skipped = plan_ppt_pages(5, {2, 4}, 4, None, True)
    assert render_map == {1: 1, 3: 3}, render_map
    assert skipped == [2, 4, 5], skipped


def _make_pdf(path: Path, pages: int) -> None:
    doc = fitz.open()
    for _ in range(pages):
        doc.new_page(width=200, height=200)
    doc.save(path)
    doc.close()


def test_pdf_to_images_labels(tmp: Path):
    pdf = tmp / "deck.pdf"
    _make_pdf(pdf, 3)
    results = pdf_to_images(
        pdf, tmp / "img", dpi=50, pages={1, 2, 3}, labels={1: 1, 2: 3, 3: 5}
    )
    assert [page for page, _ in results] == [1, 3, 5], results
    for label in (1, 3, 5):
        assert (tmp / "img" / f"page-{label:04d}.png").is_file()


def test_pdf_to_images_no_labels(tmp: Path):
    pdf = tmp / "doc.pdf"
    _make_pdf(pdf, 3)
    results = pdf_to_images(pdf, tmp / "img", dpi=50, pages={2})
    assert [page for page, _ in results] == [2], results
    assert (tmp / "img" / "page-0002.png").is_file()


def test_prepare_images_pdf_source_unchanged(tmp: Path):
    # Regression: the PDF path must behave exactly as before.
    pdf = tmp / "sample.pdf"
    _make_pdf(pdf, 4)
    images, pdf_path, skipped = prepare_images(
        pdf, tmp / "work", dpi=50, image_format="png", pages={1, 3}
    )
    assert pdf_path == pdf
    assert skipped == [], skipped
    assert [page for page, _ in images] == [1, 3], images


def main() -> int:
    tests = [obj for name, obj in sorted(globals().items()) if name.startswith("test_")]
    failed = 0
    for test in tests:
        try:
            if "tmp" in test.__code__.co_varnames[: test.__code__.co_argcount]:
                with tempfile.TemporaryDirectory() as td:
                    test(Path(td))
            else:
                test()
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"FAIL {test.__name__}: {exc}")
        else:
            print(f"ok   {test.__name__}")
    print(f"\n{len(tests) - failed}/{len(tests)} tests passed.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
