"""Concurrent per-page transcription and Markdown stitching."""

from __future__ import annotations

import asyncio
import re
from dataclasses import dataclass
from pathlib import Path

from .llm import VisionLLM

# Markdown image references emitted by the model, e.g. ![alt](page-003-fig-1.png)
IMAGE_REF_RE = re.compile(r"!\[([^\]]*)\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")


@dataclass
class PageResult:
    page: int
    image: Path
    markdown: str | None = None
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None


def build_page_prompt(base_prompt: str, page_no: int) -> str:
    """Append the page-number context needed for canonical image references."""

    return (
        f"{base_prompt}\n\n"
        f"The image provided is page {page_no} of the document. "
        "Use this page number for any page-N-fig-K.png image references."
    )


async def transcribe_pages(
    pages: list[tuple[int, Path]],
    prompt: str,
    llm: VisionLLM,
    concurrency: int = 3,
    on_page_done: callable | None = None,
) -> list[PageResult]:
    """Transcribe every page image, bounded by a concurrency semaphore.

    Results are returned ordered by page number. A page that fails after all
    retries is captured with its error instead of aborting the whole document.
    """

    semaphore = asyncio.Semaphore(max(1, concurrency))
    results: list[PageResult] = []

    async def worker(page_no: int, image_path: Path) -> None:
        async with semaphore:
            result = PageResult(page=page_no, image=image_path)
            try:
                result.markdown = await llm.transcribe_image(
                    image_path, build_page_prompt(prompt, page_no)
                )
            except Exception as exc:  # noqa: BLE001 - record and continue
                result.error = str(exc)
            results.append(result)
            if on_page_done is not None:
                on_page_done(result)

    await asyncio.gather(*(worker(page_no, path) for page_no, path in pages))
    return sorted(results, key=lambda r: r.page)


def rewrite_image_refs(markdown: str, images: list[Path], images_dir_name: str) -> str:
    """Point model-generated image refs at extracted figure files, in order.

    Refs are matched to the page's extracted figures in reading order. Refs
    without a corresponding figure are left untouched.
    """

    remaining = iter(images)

    def repl(match: "re.Match[str]") -> str:
        try:
            image = next(remaining)
        except StopIteration:
            return match.group(0)
        return f"![{match.group(1)}]({images_dir_name}/{image.name})"

    return IMAGE_REF_RE.sub(repl, markdown)


def count_image_refs(markdown: str) -> int:
    return len(IMAGE_REF_RE.findall(markdown))


def stitch_markdown(
    results: list[PageResult],
    title: str,
    page_marker: bool = True,
) -> str:
    """Combine per-page Markdown into a single document."""

    parts: list[str] = [f"# {title}"]
    for result in results:
        if page_marker:
            parts.append(f"<!-- page {result.page} -->")
        if result.ok:
            parts.append(result.markdown or "")
        else:
            parts.append(f"[Page {result.page} transcription failed: {result.error}]")
    return "\n\n".join(parts).rstrip() + "\n"
