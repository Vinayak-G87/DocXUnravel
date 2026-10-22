"""Optional DOCX visual regression support for DocXUnravel.

This module is intentionally separate from the core renderer. It requires an
external office renderer (LibreOffice/soffice) plus PyMuPDF and Pillow.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any


def visual_dependencies() -> dict[str, Any]:
    renderer = shutil.which("libreoffice") or shutil.which("soffice")
    try:
        import pymupdf  # noqa: F401
        from PIL import Image  # noqa: F401
    except ImportError:
        imaging = False
    else:
        imaging = True
    return {
        "available": bool(renderer and imaging),
        "renderer": renderer,
        "imaging": imaging,
        "required_python_packages": ["PyMuPDF", "Pillow"],
    }


def _require_dependencies(renderer: str | None = None) -> str:
    dependencies = visual_dependencies()
    executable = renderer or dependencies["renderer"]
    if not executable:
        raise RuntimeError(
            "Visual verification requires LibreOffice (`libreoffice` or `soffice`) "
            "on PATH, or an explicit --renderer path."
        )
    if not dependencies["imaging"]:
        raise RuntimeError(
            "Visual verification requires optional Python packages PyMuPDF and Pillow. "
            "Install them with `pip install 'docxunravel[visual]'`."
        )
    return str(executable)


def _render_pdf(source: Path, output_dir: Path, renderer: str, timeout: int) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    user_profile = output_dir / "libreoffice-profile"
    user_profile.mkdir()
    command = [
        renderer,
        f"-env:UserInstallation={user_profile.resolve().as_uri()}",
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(source),
    ]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(
            f"DOCX-to-PDF rendering timed out after {timeout}s for {source}"
        ) from error
    pdf = output_dir / f"{source.stem}.pdf"
    if result.returncode != 0 or not pdf.exists():
        detail = (
            result.stderr or result.stdout
        ).strip() or "renderer produced no diagnostic output"
        raise RuntimeError(
            f"DOCX-to-PDF rendering failed for {source}: "
            f"exit code {result.returncode}; {detail}"
        )
    return pdf


def _page_image(page: Any, dpi: int) -> Any:
    import pymupdf
    from PIL import Image

    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False)
    return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def _normalized_pair(first: Any, second: Any) -> tuple[Any, Any]:
    from PIL import Image

    width = max(first.width, second.width)
    height = max(first.height, second.height)
    if first.size != (width, height):
        canvas = Image.new("RGB", (width, height), "white")
        canvas.paste(first, (0, 0))
        first = canvas
    if second.size != (width, height):
        canvas = Image.new("RGB", (width, height), "white")
        canvas.paste(second, (0, 0))
        second = canvas
    return first, second


def compare_pdfs(
    original_pdf: str | Path,
    generated_pdf: str | Path,
    output_dir: str | Path,
    *,
    difference_threshold: float = 0.001,
    expected_difference_threshold: float = 0.15,
    pixel_threshold: int = 12,
    dimension_tolerance_pt: float = 0.5,
    expected_changed_pages: set[int] | None = None,
    no_op: bool = False,
    dpi: int = 144,
) -> dict[str, Any]:
    import pymupdf
    from PIL import ImageChops, ImageEnhance

    if not 0 <= difference_threshold <= 1:
        raise ValueError("difference_threshold must be between 0 and 1")
    if not difference_threshold <= expected_difference_threshold <= 1:
        raise ValueError(
            "expected_difference_threshold must be between difference_threshold and 1"
        )
    if not 0 <= pixel_threshold <= 255:
        raise ValueError("pixel_threshold must be between 0 and 255")
    expected = expected_changed_pages or set()
    if no_op and expected:
        raise ValueError("No-op visual mode does not allow expected changed pages")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    original_document = pymupdf.open(str(original_pdf))
    generated_document = pymupdf.open(str(generated_pdf))
    page_count = max(len(original_document), len(generated_document))
    pages = []
    unexpected_pages = []

    for page_index in range(page_count):
        page_number = page_index + 1
        original_page = (
            original_document[page_index]
            if page_index < len(original_document)
            else None
        )
        generated_page = (
            generated_document[page_index]
            if page_index < len(generated_document)
            else None
        )
        if original_page is None or generated_page is None:
            status = "unexpected_change"
            unexpected_pages.append(page_number)
            pages.append(
                {
                    "page": page_number,
                    "status": status,
                    "difference_ratio": 1.0,
                    "dimension_match": False,
                    "reason": "page_missing_from_one_document",
                }
            )
            continue

        original_size = [
            round(original_page.rect.width, 3),
            round(original_page.rect.height, 3),
        ]
        generated_size = [
            round(generated_page.rect.width, 3),
            round(generated_page.rect.height, 3),
        ]
        dimension_match = all(
            abs(left - right) <= dimension_tolerance_pt
            for left, right in zip(original_size, generated_size)
        )
        original_image = _page_image(original_page, dpi)
        generated_image = _page_image(generated_page, dpi)
        original_image, generated_image = _normalized_pair(
            original_image, generated_image
        )
        difference = ImageChops.difference(original_image, generated_image)
        grayscale = difference.convert("L")
        changed_pixels = sum(
            count
            for value, count in enumerate(grayscale.histogram())
            if value > pixel_threshold
        )
        difference_ratio = changed_pixels / (grayscale.width * grayscale.height)
        visually_changed = (
            difference_ratio > difference_threshold or not dimension_match
        )
        visually_changed = (
            difference_ratio > difference_threshold or not dimension_match
        )
        expected_change = page_number in expected
        if not visually_changed:
            status = "no_change"
        elif (
            not no_op
            and expected_change
            and dimension_match
            and difference_ratio <= expected_difference_threshold
        ):
            status = "expected_change"
        else:
            status = "unexpected_change"
        if status == "unexpected_change":
            unexpected_pages.append(page_number)

        original_path = output_dir / f"page-{page_number:03d}-original.png"
        generated_path = output_dir / f"page-{page_number:03d}-generated.png"
        difference_path = output_dir / f"page-{page_number:03d}-diff.png"
        original_image.save(original_path)
        generated_image.save(generated_path)
        ImageEnhance.Contrast(difference).enhance(3).save(difference_path)
        pages.append(
            {
                "page": page_number,
                "status": status,
                "difference_ratio": round(difference_ratio, 8),
                "dimension_match": dimension_match,
                "original_size_pt": original_size,
                "generated_size_pt": generated_size,
                "images": {
                    "original": original_path.name,
                    "generated": generated_path.name,
                    "difference": difference_path.name,
                },
            }
        )

    report = {
        "passed": not unexpected_pages,
        "mode": "no-op" if no_op else "patch",
        "original_pdf": str(original_pdf),
        "generated_pdf": str(generated_pdf),
        "original_page_count": len(original_document),
        "generated_page_count": len(generated_document),
        "page_count_match": len(original_document) == len(generated_document),
        "thresholds": {
            "difference_ratio": difference_threshold,
            "expected_difference_ratio": expected_difference_threshold,
            "pixel_intensity": pixel_threshold,
            "dimension_tolerance_pt": dimension_tolerance_pt,
            "dpi": dpi,
        },
        "expected_changed_pages": sorted(expected),
        "unexpected_changed_pages": unexpected_pages,
        "pages": pages,
    }
    (output_dir / "visual-report.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8"
    )
    original_document.close()
    generated_document.close()
    return report


def verify_visuals(
    original_docx: str | Path,
    generated_docx: str | Path,
    output_dir: str | Path,
    *,
    renderer: str | None = None,
    timeout: int = 120,
    difference_threshold: float = 0.001,
    expected_difference_threshold: float = 0.15,
    pixel_threshold: int = 12,
    dimension_tolerance_pt: float = 0.5,
    expected_changed_pages: set[int] | None = None,
    no_op: bool = False,
    dpi: int = 144,
) -> dict[str, Any]:
    executable = _require_dependencies(renderer)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="docxunravel-visual-") as temporary:
        temporary = Path(temporary)
        rendered_original = _render_pdf(
            Path(original_docx), temporary / "original", executable, timeout
        )
        rendered_generated = _render_pdf(
            Path(generated_docx), temporary / "generated", executable, timeout
        )
        original_pdf = output_dir / "original.pdf"
        generated_pdf = output_dir / "generated.pdf"
        shutil.copyfile(rendered_original, original_pdf)
        shutil.copyfile(rendered_generated, generated_pdf)
        report = compare_pdfs(
            original_pdf,
            generated_pdf,
            output_dir,
            difference_threshold=difference_threshold,
            expected_difference_threshold=expected_difference_threshold,
            pixel_threshold=pixel_threshold,
            dimension_tolerance_pt=dimension_tolerance_pt,
            expected_changed_pages=expected_changed_pages,
            no_op=no_op,
            dpi=dpi,
        )
    report["renderer"] = executable
    report["original_docx"] = str(original_docx)
    report["generated_docx"] = str(generated_docx)
    (output_dir / "visual-report.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8"
    )
    return report


def _page_set(value: str) -> set[int]:
    pages: set[int] = set()
    if not value:
        return pages
    for item in value.split(","):
        item = item.strip()
        if "-" in item:
            start, end = (int(part) for part in item.split("-", 1))
            pages.update(range(start, end + 1))
        else:
            pages.add(int(item))
    if any(page < 1 for page in pages):
        raise argparse.ArgumentTypeError("Page numbers must be positive")
    return pages


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Optional DocXUnravel visual regression"
    )
    parser.add_argument("original")
    parser.add_argument("generated")
    parser.add_argument("output_dir")
    parser.add_argument("--renderer")
    parser.add_argument("--difference-threshold", type=float, default=0.001)
    parser.add_argument("--expected-difference-threshold", type=float, default=0.15)
    parser.add_argument("--pixel-threshold", type=int, default=12)
    parser.add_argument("--dimension-tolerance-pt", type=float, default=0.5)
    parser.add_argument("--expected-pages", type=_page_set, default=set())
    parser.add_argument("--no-op", action="store_true")
    parser.add_argument("--dpi", type=int, default=144)
    parser.add_argument("--timeout", type=int, default=120)
    args = parser.parse_args(argv)
    try:
        report = verify_visuals(
            args.original,
            args.generated,
            args.output_dir,
            renderer=args.renderer,
            timeout=args.timeout,
            difference_threshold=args.difference_threshold,
            expected_difference_threshold=args.expected_difference_threshold,
            pixel_threshold=args.pixel_threshold,
            dimension_tolerance_pt=args.dimension_tolerance_pt,
            expected_changed_pages=args.expected_pages,
            no_op=args.no_op,
            dpi=args.dpi,
        )
    except RuntimeError as error:
        print(
            json.dumps({"available": False, "error": str(error)}, indent=2),
            file=sys.stderr,
        )
        return 3
    print(json.dumps(report, indent=2))
    return 0 if report["passed"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
