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


def office_to_pdf(source: Path, pdf_path: Path) -> None:
    """Convert an Office document to PDF using the installed MS Office via COM."""

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
            _powerpoint_to_pdf(source, pdf_path)
        elif ext in WORD_EXTS:
            _word_to_pdf(source, pdf_path)
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


def _powerpoint_to_pdf(source: Path, pdf_path: Path) -> None:
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
            deck.SaveAs(str(pdf_path.resolve()), PP_FORMAT_PDF)
        finally:
            deck.Close()
    finally:
        app.Quit()


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
) -> list[tuple[int, Path]]:
    """Render PDF pages to images.

    Returns a list of (1-based page number, image path) tuples. If ``pages``
    is given, only those 1-based page numbers are rendered.
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
            page = doc.load_page(page_no - 1)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image_path = out_dir / f"page-{page_no:04d}.{image_format}"
            if image_format == "jpeg":
                pix.save(image_path, jpg_quality=95)
            else:
                pix.save(image_path)
            results.append((page_no, image_path))
    finally:
        doc.close()

    return results


def prepare_images(
    source: Path,
    work_dir: Path,
    dpi: int,
    image_format: str,
    pages: set[int] | None = None,
) -> tuple[list[tuple[int, Path]], Path]:
    """Full pipeline step: source document -> (page images, working PDF path)."""

    ext = source.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        raise ConversionError(
            f"Unsupported file type: {ext}. "
            f"Supported: {', '.join(sorted(SUPPORTED_EXTS))}"
        )
    if ext in PDF_EXTS:
        pdf_path = source
    else:
        pdf_path = work_dir / f"{source.stem}.pdf"
        office_to_pdf(source, pdf_path)
    images = pdf_to_images(pdf_path, work_dir, dpi=dpi, image_format=image_format, pages=pages)
    return images, pdf_path


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
