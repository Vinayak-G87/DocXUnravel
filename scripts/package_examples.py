"""Build the downloadable, self-contained form example for a GitHub release."""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_FILES = (
    "examples/form/form-template.docx",
    "examples/form/form-filled.docx",
    "examples/form/field-map.json",
    "examples/form/patch.json",
)


def build_example_archive(destination: Path, root: Path = ROOT) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w") as archive:
        for name in EXAMPLE_FILES:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            archive.writestr(info, (root / name).read_bytes())


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python scripts/package_examples.py OUTPUT.zip")
    build_example_archive(Path(sys.argv[1]))
