"""Document -> image conversion.

Office documents (.ppt/.pptx/.doc/.docx) are converted to PDF through the
installed Microsoft Office applications via COM automation (pywin32), then
every PDF (converted or source) is rendered page-by-page to HD images with
PyMuPDF.
"""

from __future__ import annotations

from pathlib import Path

import fitz  # PyMuPDF

POWERPOINT_EXTS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm"}
WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
PDF_EXTS = {".pdf"}
SUPPORTED_EXTS = POWERPOINT_EXTS | WORD_EXTS | PDF_EXTS

# Office COM constants
WD_FORMAT_PDF = 17  # wdFormatPDF
PP_FORMAT_PDF = 32  # ppSaveAsPDF
WD_ALERTS_NONE = 0  # wdAlertsNone


class ConversionError(RuntimeError):
    """Raised when a source document cannot be converted or rendered."""


def is_office_file(path: Path) -> bool:
    return path.suffix.lower() in (POWERPOINT_EXTS | WORD_EXTS)


def document_kind(path: Path) -> str:
    """Return the prompt category for a source file: 'ppt' or 'document'."""

    if path.suffix.lower() in POWERPOINT_EXTS:
        return "ppt"
    return "document"


def office_to_pdf(source: Path, pdf_path: Path) -> tuple[set[int], int]:
    """Convert an Office document to PDF using the installed MS Office via COM.

    Returns ``(hidden_slides, slide_count)``: the 1-based indices of slides
    marked as hidden in the source presentation and the total slide count.
    Both are empty/zero for non-PowerPoint documents.
    """

    ext = source.suffix.lower()
    try:
        import pythoncom
    except ImportError as exc:  # pragma: no cover - platform dependent
        raise ConversionError(
            "pywin32 is required for Office conversion and MS Office must be "
            "installed. Install with: pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()
    try:
        if ext in POWERPOINT_EXTS:
            return _powerpoint_to_pdf(source, pdf_path)
        elif ext in WORD_EXTS:
            _word_to_pdf(source, pdf_path)
            return set(), 0
        else:
            raise ConversionError(f"Unsupported Office format: {ext}")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(
            f"Failed to convert {source.name} to PDF via Microsoft Office: {exc}"
        ) from exc
    finally:
        pythoncom.CoUninitialize()


def _powerpoint_to_pdf(source: Path, pdf_path: Path) -> tuple[set[int], int]:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        # WithWindow=False keeps the conversion headless (no visible window).
        deck = app.Presentations.Open(
            str(source.resolve()),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        try:
            slide_count = deck.Slides.Count
            # SlideShowTransition.Hidden is an MsoTriState: msoTrue (-1) when
            # the slide is hidden, msoFalse (0) otherwise.
            hidden = {
                i
                for i in range(1, slide_count + 1)
                if deck.Slides(i).SlideShowTransition.Hidden
            }
            deck.SaveAs(str(pdf_path.resolve()), PP_FORMAT_PDF)
        finally:
            deck.Close()
    finally:
        app.Quit()
    return hidden, slide_count


def _word_to_pdf(source: Path, pdf_path: Path) -> None:
    import win32com.client

    app = win32com.client.DispatchEx("Word.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = WD_ALERTS_NONE
        doc = app.Documents.Open(
            str(source.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        try:
            doc.SaveAs(str(pdf_path.resolve()), FileFormat=WD_FORMAT_PDF)
        finally:
            doc.Close(SaveChanges=False)
    finally:
        app.Quit()


def pdf_to_images(
    pdf_path: Path,
    out_dir: Path,
    dpi: int = 300,
    image_format: str = "png",
    pages: set[int] | None = None,
    labels: dict[int, int] | None = None,
) -> list[tuple[int, Path]]:
    """Render PDF pages to images.

    Returns a list of (1-based page number, image path) tuples. If ``pages``
    is given, only those 1-based page numbers are rendered. ``labels`` may map
    a PDF page number to a different output number — used when the PDF page
    order differs from the source document's (e.g. hidden PowerPoint slides
    excluded from the export); the label replaces the PDF page number in the
    returned tuples and the image file names.
    """

    out_dir.mkdir(parents=True, exist_ok=True)
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    results: list[tuple[int, Path]] = []

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        raise ConversionError(f"Cannot open PDF {pdf_path}: {exc}") from exc

    try:
        total = doc.page_count
        wanted = pages if pages is not None else set(range(1, total + 1))
        missing = sorted(p for p in wanted if p < 1 or p > total)
        if missing:
            raise ConversionError(
                f"Requested pages {missing} are outside the document range 1-{total}."
            )
        for page_no in sorted(wanted):
            label = labels.get(page_no, page_no) if labels else page_no
            page = doc.load_page(page_no - 1)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image_path = out_dir / f"page-{label:04d}.{image_format}"
            if image_format == "jpeg":
                pix.save(image_path, jpg_quality=95)
            else:
                pix.save(image_path)
            results.append((label, image_path))
    finally:
        doc.close()

    return results


def plan_ppt_pages(
    total_slides: int,
    hidden: set[int],
    pdf_page_count: int,
    requested: set[int] | None,
    skip_hidden: bool,
) -> tuple[dict[int, int], list[int]]:
    """Decide which PDF pages to render for a PowerPoint source.

    Returns ``(render_map, skipped)`` where ``render_map`` maps a 1-based PDF
    page number to the slide number it represents, and ``skipped`` lists the
    requested slide numbers that were excluded (hidden, or absent from the
    PDF).

    PowerPoint's PDF export may or may not include hidden slides depending on
    version and settings, so the slide -> page mapping is derived by comparing
    the PDF page count against the slide count:

    - ``pdf_page_count == total_slides``: every slide is in the PDF and page N
      is slide N; hidden slides are filtered out here instead.
    - ``pdf_page_count == total_slides - len(hidden)``: hidden slides were
      already excluded from the PDF; visible slides compact onto pages
      1..pdf_page_count.
    - Otherwise (unexpected): fall back to identity within the PDF's range.
    """

    if pdf_page_count == total_slides:
        slide_to_page = {s: s for s in range(1, total_slides + 1)}
    elif pdf_page_count == total_slides - len(hidden):
        visible = (s for s in range(1, total_slides + 1) if s not in hidden)
        slide_to_page = {s: i for i, s in enumerate(visible, start=1)}
    else:
        slide_to_page = {
            s: s for s in range(1, min(total_slides, pdf_page_count) + 1)
        }

    wanted = requested if requested is not None else set(range(1, total_slides + 1))
    out_of_range = sorted(s for s in wanted if s < 1 or s > total_slides)
    if out_of_range:
        raise ConversionError(
            f"Requested pages {out_of_range} are outside the document range "
            f"1-{total_slides}."
        )

    skipped = sorted(
        s for s in wanted if (skip_hidden and s in hidden) or s not in slide_to_page
    )
    render_map = {
        slide_to_page[s]: s
        for s in sorted(wanted)
        if s in slide_to_page and not (skip_hidden and s in hidden)
    }
    return render_map, skipped


def prepare_images(
    source: Path,
    work_dir: Path,
    dpi: int,
    image_format: str,
    pages: set[int] | None = None,
    skip_hidden: bool = True,
) -> tuple[list[tuple[int, Path]], Path, list[int]]:
    """Full pipeline step: source document -> (page images, PDF path, skipped).

    ``skipped`` lists PowerPoint slide numbers excluded because they are
    marked hidden (empty for non-PowerPoint sources, when none are hidden, or
    when ``skip_hidden`` is false and hidden slides are present in the PDF).
    """

    ext = source.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        raise ConversionError(
            f"Unsupported file type: {ext}. "
            f"Supported: {', '.join(sorted(SUPPORTED_EXTS))}"
        )
    if ext in PDF_EXTS:
        pdf_path = source
        images = pdf_to_images(pdf_path, work_dir, dpi=dpi, image_format=image_format, pages=pages)
        return images, pdf_path, []

    pdf_path = work_dir / f"{source.stem}.pdf"
    hidden, slide_count = office_to_pdf(source, pdf_path)

    if ext not in POWERPOINT_EXTS or not slide_count:
        images = pdf_to_images(pdf_path, work_dir, dpi=dpi, image_format=image_format, pages=pages)
        return images, pdf_path, []

    try:
        with fitz.open(pdf_path) as doc:
            pdf_page_count = doc.page_count
    except Exception as exc:
        raise ConversionError(f"Cannot open PDF {pdf_path}: {exc}") from exc

    render_map, skipped = plan_ppt_pages(
        slide_count, hidden, pdf_page_count, pages, skip_hidden
    )
    if not render_map:
        raise ConversionError(
            f"Nothing to convert: all requested slide(s) of {source.name} are "
            f"hidden ({', '.join(map(str, skipped))})."
        )
    images = pdf_to_images(
        pdf_path,
        work_dir,
        dpi=dpi,
        image_format=image_format,
        pages=set(render_map),
        labels=render_map,
    )
    return images, pdf_path, skipped


# Embedded images smaller than this on either side are treated as decorative
# (icons, rules, logos) and not extracted as figures.
MIN_FIGURE_SIZE_PT = 60.0


def extract_figures(
    pdf_path: Path,
    pages: set[int],
    out_dir: Path,
) -> dict[int, list[Path]]:
    """Extract embedded raster figures per page, in visual reading order.

    Returns ``{page_number: [image_path, ...]}`` with files named
    ``page-NNN-fig-K.<ext>`` so Markdown image references can point at them.
    """

    out_dir.mkdir(parents=True, exist_ok=True)
    figures: dict[int, list[Path]] = {}

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        raise ConversionError(f"Cannot open PDF {pdf_path}: {exc}") from exc

    try:
        for page_no in sorted(pages):
            if page_no < 1 or page_no > doc.page_count:
                continue
            page = doc.load_page(page_no - 1)
            candidates: list[tuple[fitz.Rect, int]] = []
            for info in page.get_image_info(xrefs=True):
                bbox = fitz.Rect(info["bbox"])
                if bbox.width < MIN_FIGURE_SIZE_PT or bbox.height < MIN_FIGURE_SIZE_PT:
                    continue
                candidates.append((bbox, int(info.get("xref", 0))))
            # Reading order: top-to-bottom, then left-to-right.
            candidates.sort(key=lambda item: (round(item[0].y0), round(item[0].x0)))

            saved: list[Path] = []
            for index, (bbox, xref) in enumerate(candidates, start=1):
                image_path = _save_figure(doc, page, bbox, xref, out_dir, page_no, index)
                if image_path is not None:
                    saved.append(image_path)
            if saved:
                figures[page_no] = saved
    finally:
        doc.close()

    return figures


def _save_figure(
    doc: fitz.Document,
    page: fitz.Page,
    bbox: fitz.Rect,
    xref: int,
    out_dir: Path,
    page_no: int,
    index: int,
) -> Path | None:
    base = out_dir / f"page-{page_no:03d}-fig-{index}"

    if xref > 0:
        try:
            extracted = doc.extract_image(xref)
            data: bytes = extracted["image"]
            ext = str(extracted["ext"]).lower()
        except Exception:
            data, ext = b"", ""
        if data and ext in ("png", "jpeg", "jpg"):
            path = base.with_suffix(".png" if ext == "png" else ".jpg")
            path.write_bytes(data)
            return path
        if data:
            # Exotic format (jbig2, jpx, ...) — convert to PNG via Pillow.
            try:
                import io

                from PIL import Image

                with Image.open(io.BytesIO(data)) as im:
                    path = base.with_suffix(".png")
                    im.save(path)
                return path
            except Exception:
                pass

    # Inline image (xref 0) or extraction failure: render the clipped region.
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=bbox, alpha=False)
        path = base.with_suffix(".png")
        pix.save(path)
        return path
    except Exception:
        return None
