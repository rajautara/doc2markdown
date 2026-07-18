"""Probe: check whether the configured endpoint accepts image input.

Sends a 1x1 PNG through the exact same client code path as `doc2md convert`
and prints the outcome. Usage:  .venv\\Scripts\\python probe_vision.py
"""

import asyncio
import base64
import tempfile
from pathlib import Path

from doc2md.config import load_config
from doc2md.llm import VisionLLM

PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


async def main() -> None:
    cfg = load_config(None)
    print(f"endpoint : {cfg.llm.base_url}")
    print(f"model    : {cfg.llm.model}")
    print(f"api_mode : {cfg.llm.api_mode}")

    with tempfile.TemporaryDirectory() as tmp:
        img = Path(tmp) / "probe.png"
        img.write_bytes(PNG_1PX)
        llm = VisionLLM(cfg.llm)
        try:
            text = await llm.transcribe_image(
                img, "Describe this image in one short sentence."
            )
            print("OK - endpoint accepts images.")
            print("response:", text[:300])
        except Exception as exc:
            print("FAILED:", exc)
        finally:
            await llm.aclose()


if __name__ == "__main__":
    asyncio.run(main())
