"""Vision LLM client built on httpx (OpenAI-compatible endpoints).

Supports both the Chat Completions API (default) and the Responses API,
optional custom headers, optional temperature, and an SSL verification toggle.
"""

from __future__ import annotations

import asyncio
import base64
import random
from pathlib import Path
from typing import Any

import httpx

from .config import LLMConfig

RETRYABLE_STATUSES = {408, 409, 425, 429, 500, 502, 503, 504}
_MIME = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg"}


class TranscriptionError(RuntimeError):
    """Raised when the LLM request fails after all retries or returns bad data."""


def _image_data_url(image_path: Path) -> str:
    ext = image_path.suffix.lower().lstrip(".")
    mime = _MIME.get(ext, "image/png")
    encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


class VisionLLM:
    def __init__(self, cfg: LLMConfig):
        self._cfg = cfg
        headers: dict[str, str] = {}
        if cfg.api_key:
            headers["Authorization"] = f"Bearer {cfg.api_key}"
        if cfg.headers:
            headers.update(cfg.headers)
        self._client = httpx.AsyncClient(
            base_url=cfg.base_url.rstrip("/"),
            headers=headers,
            verify=cfg.ssl_verify,
            timeout=httpx.Timeout(cfg.timeout),
        )

    async def aclose(self) -> None:
        await self._client.aclose()

    async def transcribe_image(self, image_path: Path, prompt: str) -> str:
        """Send one page image to the vision model and return its Markdown."""

        data_url = _image_data_url(image_path)
        if self._cfg.api_mode == "response":
            url = "/responses"
            payload = self._build_response_payload(prompt, data_url)
        else:
            url = "/chat/completions"
            payload = self._build_chat_payload(prompt, data_url)
        body = await self._post_with_retry(url, payload)
        return self._extract_text(body)

    # ------------------------------------------------------------------ payloads

    def _build_chat_payload(self, prompt: str, data_url: str) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "model": self._cfg.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": data_url}},
                    ],
                }
            ],
        }
        # Temperature is only sent when explicitly configured.
        if self._cfg.temperature is not None:
            payload["temperature"] = self._cfg.temperature
        return payload

    def _build_response_payload(self, prompt: str, data_url: str) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "model": self._cfg.model,
            "input": [
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {"type": "input_image", "image_url": data_url},
                    ],
                }
            ],
        }
        if self._cfg.temperature is not None:
            payload["temperature"] = self._cfg.temperature
        return payload

    # ------------------------------------------------------------------- request

    async def _post_with_retry(self, url: str, payload: dict[str, Any]) -> dict[str, Any]:
        attempts = max(1, self._cfg.max_retries + 1)
        last_error: Exception | None = None

        for attempt in range(attempts):
            try:
                response = await self._client.post(url, json=payload)
                response.raise_for_status()
                return response.json()
            except httpx.HTTPStatusError as exc:
                last_error = exc
                if exc.response.status_code not in RETRYABLE_STATUSES:
                    detail = exc.response.text[:500]
                    raise TranscriptionError(
                        f"LLM returned HTTP {exc.response.status_code}: {detail}"
                    ) from exc
            except (httpx.ConnectError, httpx.ReadError, httpx.WriteError, httpx.TimeoutException) as exc:
                last_error = exc
            except ValueError as exc:  # invalid JSON body
                raise TranscriptionError(f"LLM returned invalid JSON: {exc}") from exc

            if attempt < attempts - 1:
                delay = min(2.0**attempt * 1.5, 60.0) + random.uniform(0, 0.5)
                await asyncio.sleep(delay)

        raise TranscriptionError(
            f"LLM request failed after {attempts} attempts: {last_error}"
        ) from last_error

    # ------------------------------------------------------------------ response

    def _extract_text(self, body: dict[str, Any]) -> str:
        if self._cfg.api_mode == "response":
            text = self._extract_response_text(body)
        else:
            text = self._extract_chat_text(body)
        text = (text or "").strip()
        if not text:
            raise TranscriptionError("LLM returned an empty transcription.")
        return text

    @staticmethod
    def _extract_chat_text(body: dict[str, Any]) -> str:
        try:
            return body["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError) as exc:
            raise TranscriptionError(
                f"Unexpected chat completion response shape: {str(body)[:500]}"
            ) from exc

    @staticmethod
    def _extract_response_text(body: dict[str, Any]) -> str:
        # Fast path: the SDK-style convenience field.
        output_text = body.get("output_text")
        if isinstance(output_text, str) and output_text.strip():
            return output_text
        # Otherwise walk output -> message -> content -> output_text items.
        parts: list[str] = []
        for item in body.get("output") or []:
            if not isinstance(item, dict) or item.get("type") != "message":
                continue
            for content in item.get("content") or []:
                if isinstance(content, dict) and content.get("type") == "output_text":
                    parts.append(content.get("text", ""))
        if parts:
            return "\n".join(parts)
        raise TranscriptionError(
            f"Unexpected responses API response shape: {str(body)[:500]}"
        )
