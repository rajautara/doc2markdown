"""YAML configuration loading and validation for doc2md."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Literal

import yaml
from pydantic import BaseModel, Field

DEFAULT_CONFIG_NAMES = ("config.yaml", "config.yml")

_ENV_PATTERN = re.compile(r"\$\{([^}]+)\}")


class LLMConfig(BaseModel):
    base_url: str
    api_key: str = ""
    model: str
    api_mode: Literal["chat_completion", "response"] = "chat_completion"
    temperature: float | None = None
    headers: dict[str, str] | None = None
    ssl_verify: bool = True
    timeout: float = 120.0
    max_retries: int = 3


class RenderConfig(BaseModel):
    dpi: int = 300
    image_format: Literal["png", "jpeg"] = "png"


class ProcessingConfig(BaseModel):
    concurrency: int = 3


class OutputConfig(BaseModel):
    dir: Path = Path("./output")
    page_marker: bool = True
    extract_images: bool = True


class Config(BaseModel):
    llm: LLMConfig
    render: RenderConfig = Field(default_factory=RenderConfig)
    processing: ProcessingConfig = Field(default_factory=ProcessingConfig)
    output: OutputConfig = Field(default_factory=OutputConfig)


class ConfigError(RuntimeError):
    """Raised when the configuration file is missing or invalid."""


def _expand_env(value: object) -> object:
    """Expand ${VAR} references inside string values (recursively)."""

    if isinstance(value, str):
        return _ENV_PATTERN.sub(lambda m: os.environ.get(m.group(1), ""), value)
    if isinstance(value, dict):
        return {k: _expand_env(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_expand_env(v) for v in value]
    return value


def find_config(explicit: Path | None) -> Path:
    """Resolve the config path: explicit option, else config.yaml/yml in cwd."""

    if explicit is not None:
        if not explicit.is_file():
            raise ConfigError(f"Config file not found: {explicit}")
        return explicit
    for name in DEFAULT_CONFIG_NAMES:
        candidate = Path.cwd() / name
        if candidate.is_file():
            return candidate
    raise ConfigError(
        "No config file found. Create one with `doc2md init-config` "
        "or pass -c <path>."
    )


def load_config(explicit: Path | None = None) -> Config:
    path = find_config(explicit)
    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    except yaml.YAMLError as exc:
        raise ConfigError(f"Invalid YAML in {path}: {exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigError(f"Config file {path} must contain a YAML mapping.")
    try:
        return Config.model_validate(_expand_env(raw))
    except Exception as exc:  # pydantic ValidationError and friends
        raise ConfigError(f"Invalid configuration in {path}:\n{exc}") from exc
