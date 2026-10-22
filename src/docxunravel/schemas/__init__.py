"""Versioned JSON Schemas distributed with DocXUnravel."""

from importlib.resources import files
from pathlib import Path


def schema_path(filename: str) -> Path:
    """Return the installed path for a bundled DocXUnravel schema."""

    return Path(str(files(__package__).joinpath(filename)))


__all__ = ["schema_path"]
