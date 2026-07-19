"""doc2md command line interface."""

from __future__ import annotations

import asyncio
import hashlib
import shutil
import tempfile
from importlib import resources
from pathlib import Path

import typer
from rich.console import Console
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
)

from .config import ConfigError, load_config
from .converter import (
    ConversionError,
    document_kind,
    extract_figures,
    prepare_images,
)
from .llm import VisionLLM
from .transcribe import (
    count_image_refs,
    rewrite_image_refs,
    stitch_markdown,
    transcribe_pages,
)

app = typer.Typer(
    name="doc2md",
    help="Convert PowerPoint, Word and PDF documents to Markdown using a vision LLM.",
    no_args_is_help=True,
)
console = Console()
err_console = Console(stderr=True)


def parse_pages(spec: str) -> set[int]:
    """Parse a page spec like '1-3,7,9-10' into a set of 1-based page numbers."""

    pages: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start, end = int(start_s), int(end_s)
            if start > end:
                start, end = end, start
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    if not pages:
        raise typer.BadParameter(f"Invalid --pages value: {spec!r}")
    return pages


def load_prompt(kind: str) -> str:
    """Load the bundled transcription prompt for 'ppt' or 'document'."""

    name = "ppt.txt" if kind == "ppt" else "document.txt"
    return (
        resources.files("doc2md.prompts")
        .joinpath(name)
        .read_text(encoding="utf-8")
    )


async def _run_conversion(
    source: Path,
    out_path: Path,
    pages: set[int] | None,
    cfg,
) -> int:
    kind = document_kind(source)
    prompt = load_prompt(kind)

    with tempfile.TemporaryDirectory(prefix="doc2md-") as tmp:
        work_dir = Path(tmp)
        with console.status(f"Rendering {source.name} to images..."):
            images, pdf_path, skipped = prepare_images(
                source,
                work_dir,
                dpi=cfg.render.dpi,
                image_format=cfg.render.image_format,
                pages=pages,
                skip_hidden=cfg.processing.skip_hidden_slides,
            )
        if skipped:
            console.print(
                f"[yellow]Skipped {len(skipped)} hidden slide(s): "
                f"{', '.join(map(str, skipped))}[/yellow]"
            )
        console.print(f"Rendered {len(images)} page(s) at {cfg.render.dpi} DPI.")

        llm = VisionLLM(cfg.llm)
        try:
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                MofNCompleteColumn(),
                TimeElapsedColumn(),
                console=console,
            ) as progress:
                task = progress.add_task(
                    f"Transcribing ({cfg.llm.model})", total=len(images)
                )

                def on_done(_result) -> None:
                    progress.advance(task)

                results = await transcribe_pages(
                    images,
                    prompt,
                    llm,
                    concurrency=cfg.processing.concurrency,
                    on_page_done=on_done,
                )
        finally:
            await llm.aclose()

        if cfg.output.extract_images:
            if kind == "ppt":
                _embed_slide_renders(images, results, out_path)
            else:
                _extract_and_link_figures(pdf_path, results, out_path)

    markdown = stitch_markdown(
        results, title=source.stem, page_marker=cfg.output.page_marker
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(markdown, encoding="utf-8")

    failed = [r for r in results if not r.ok]
    if failed:
        for r in failed:
            err_console.print(f"[red]Page {r.page} failed:[/red] {r.error}")
        err_console.print(
            f"[yellow]{len(failed)} page(s) failed; placeholders were inserted.[/yellow]"
        )
        return 1

    console.print(f"[green]Done:[/green] {out_path}")
    return 0


def _images_dir_for(out_path: Path) -> Path:
    """Images folder: <first 12 hex chars of sha256(doc name)>_images/."""

    digest = hashlib.sha256(out_path.stem.encode("utf-8")).hexdigest()[:12]
    return out_path.parent / f"{digest}_images"


def _embed_slide_renders(
    rendered: list[tuple[int, Path]], results, out_path: Path
) -> None:
    """Copy each rendered slide into the images folder and embed a reference
    to it directly above the slide's transcription."""

    images_dir = _images_dir_for(out_path)
    images_dir.mkdir(parents=True, exist_ok=True)
    rendered_by_page = dict(rendered)
    copied = 0

    for result in results:
        if not result.ok or not result.markdown:
            continue
        src = rendered_by_page.get(result.page)
        if src is None:
            continue
        dest = images_dir / src.name
        shutil.copyfile(src, dest)
        ref = f"![Slide {result.page}]({images_dir.name}/{dest.name})"
        result.markdown = f"{ref}\n\n{result.markdown}"
        copied += 1

    if copied:
        console.print(f"Embedded {copied} slide render(s) to {images_dir}")


def _extract_and_link_figures(pdf_path: Path, results, out_path: Path) -> None:
    """Extract page figures and point the page's image refs at the real files."""

    images_dir = _images_dir_for(out_path)
    figures = extract_figures(
        pdf_path, {r.page for r in results if r.ok}, images_dir
    )
    total = sum(len(paths) for paths in figures.values())

    for result in results:
        if not result.ok or not result.markdown:
            continue
        page_figures = figures.get(result.page, [])
        n_refs = count_image_refs(result.markdown)
        if page_figures:
            result.markdown = rewrite_image_refs(
                result.markdown, page_figures, images_dir.name
            )
        if n_refs != len(page_figures) and (n_refs or page_figures):
            console.print(
                f"[yellow]Page {result.page}: {n_refs} image ref(s) vs "
                f"{len(page_figures)} extracted figure(s) — please check the mapping.[/yellow]"
            )

    if total:
        console.print(f"Extracted {total} figure(s) to {images_dir}")


@app.command()
def convert(
    source: Path = typer.Argument(
        ..., exists=True, dir_okay=False, help="PowerPoint, Word or PDF file to convert."
    ),
    config: Path | None = typer.Option(
        None, "--config", "-c", help="Path to YAML config (default: ./config.yaml)."
    ),
    out: Path | None = typer.Option(
        None, "--out", "-o", help="Output .md path (default: <output.dir>/<name>.md)."
    ),
    pages: str | None = typer.Option(
        None, "--pages", help="Page subset, e.g. '1-3,7' (default: all pages)."
    ),
    model: str | None = typer.Option(
        None, "--model", help="Override the model from the config."
    ),
    dpi: int | None = typer.Option(
        None, "--dpi", help="Override the render DPI from the config."
    ),
) -> None:
    """Convert a document to Markdown via page images and a vision LLM."""

    try:
        cfg = load_config(config)
    except ConfigError as exc:
        err_console.print(f"[red]Config error:[/red] {exc}")
        raise typer.Exit(code=2) from exc

    if model:
        cfg.llm.model = model
    if dpi:
        cfg.render.dpi = dpi

    page_set: set[int] | None = None
    if pages:
        try:
            page_set = parse_pages(pages)
        except ValueError as exc:
            err_console.print(f"[red]Invalid --pages value:[/red] {exc}")
            raise typer.Exit(code=2) from exc

    out_path = out or (cfg.output.dir / f"{source.stem}.md")

    try:
        exit_code = asyncio.run(_run_conversion(source, out_path, page_set, cfg))
    except ConversionError as exc:
        err_console.print(f"[red]Conversion error:[/red] {exc}")
        raise typer.Exit(code=1) from exc
    raise typer.Exit(code=exit_code)


@app.command(name="init-config")
def init_config(
    dest: Path = typer.Option(
        Path("config.yaml"), "--dest", "-d", help="Where to write the config file."
    ),
    force: bool = typer.Option(
        False, "--force", "-f", help="Overwrite an existing file."
    ),
) -> None:
    """Write a ready-to-edit config.yaml based on the bundled example."""

    if dest.exists() and not force:
        err_console.print(
            f"[red]{dest} already exists.[/red] Use --force to overwrite."
        )
        raise typer.Exit(code=2)

    example = Path(__file__).resolve().parent.parent / "config.example.yaml"
    if not example.is_file():
        err_console.print("[red]Bundled config.example.yaml not found.[/red]")
        raise typer.Exit(code=1)
    shutil.copyfile(example, dest)
    console.print(f"[green]Config written to[/green] {dest}")


if __name__ == "__main__":
    app()
