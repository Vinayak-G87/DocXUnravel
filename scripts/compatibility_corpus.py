"""Baseline compatibility assessment for a public DOCX corpus.

This is a test harness, not part of the patch/render API. It never writes to the
source corpus and does not relax production validators or visual thresholds.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import shutil
import tempfile
import time
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

from lxml import etree

from docxunravel import (
    UNSAFE_EDIT_TAGS,
    create_fidelity_report,
    export_bundle,
    profile_document,
    render_bundle,
    validate_docx_package,
)
from docxunravel.visual_regression import verify_visuals, visual_dependencies

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

PACKAGE_FEATURE_TAGS = {
    f"{{{W_NS}}}bookmarkStart": "bookmark",
    f"{{{W_NS}}}commentReference": "comment_reference",
    f"{{{W_NS}}}ins": "tracked_insertion",
    f"{{{W_NS}}}del": "tracked_deletion",
    f"{{{W_NS}}}moveFrom": "tracked_move_from",
    f"{{{W_NS}}}moveTo": "tracked_move_to",
    f"{{{MC_NS}}}AlternateContent": "alternate_content",
    f"{{{M_NS}}}oMath": "office_math",
    f"{{{M_NS}}}oMathPara": "office_math",
}


def _paragraphs(profile: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []

    def visit(elements: list[dict[str, Any]]) -> None:
        for element in elements:
            if element.get("type") == "paragraph":
                result.append(element)
            elif element.get("type") == "table":
                for row in element.get("rows", []):
                    for cell in row.get("cells", []):
                        visit(cell.get("elements", []))

    visit(profile.get("body", []))
    for container in profile.get("headers", []) + profile.get("footers", []):
        visit(container.get("elements", []))
    return result


def _table_count(profile: dict[str, Any]) -> int:
    count = 0

    def visit(elements: list[dict[str, Any]]) -> None:
        nonlocal count
        for element in elements:
            if element.get("type") == "table":
                count += 1
                for row in element.get("rows", []):
                    for cell in row.get("cells", []):
                        visit(cell.get("elements", []))

    visit(profile.get("body", []))
    for container in profile.get("headers", []) + profile.get("footers", []):
        visit(container.get("elements", []))
    return count


def _package_features(bundle: Path, profile: dict[str, Any]) -> list[str]:
    features = {
        feature
        for paragraph in _paragraphs(profile)
        for feature in paragraph.get("unsupported_edit_structures", [])
    }
    package = bundle / "package"
    if (package / "word" / "comments.xml").exists():
        features.add("comments")
    if (package / "word" / "footnotes.xml").exists():
        features.add("footnotes")
    if (package / "word" / "endnotes.xml").exists():
        features.add("endnotes")
    if (package / "word" / "charts").exists():
        features.add("charts")
    if (package / "word" / "embeddings").exists():
        features.add("embedded_package")
    if (package / "customXml").exists():
        features.add("custom_xml")
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        huge_tree=False,
        recover=False,
    )
    for path in package.rglob("*"):
        if not path.is_file() or path.suffix not in {".xml", ".rels"}:
            continue
        try:
            root = etree.parse(str(path), parser).getroot()
        except Exception:
            continue
        for node in root.iter():
            if node.tag in PACKAGE_FEATURE_TAGS:
                features.add(PACKAGE_FEATURE_TAGS[node.tag])
            unsafe = UNSAFE_EDIT_TAGS.get(node.tag)
            if unsafe:
                features.add(unsafe)
    return sorted(features)


def _safe_paragraph(profile: dict[str, Any]) -> dict[str, Any] | None:
    for paragraph in _paragraphs(profile):
        text = paragraph.get("text", "")
        locator = paragraph.get("locator", {})
        if (
            locator.get("part") == "word/document.xml"
            and text.strip()
            and len(text) <= 500
            and not paragraph.get("has_revisions")
            and not paragraph.get("unsupported_edit_structures")
            and not paragraph.get("has_mixed_run_formatting")
            and paragraph.get("runs")
        ):
            return paragraph
    return None


def _stage_error(result: dict[str, Any], stage: str, error: BaseException) -> None:
    result["errors"].append(
        {
            "stage": stage,
            "type": type(error).__name__,
            "message": str(error),
        }
    )


def _preserve_failure(
    source: Path,
    result: dict[str, Any],
    working: Path,
    failures: Path,
) -> str:
    identity = result["id"]
    destination = failures / identity
    destination.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination / source.name)
    (destination / "failure.json").write_text(
        json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    for name in (
        "bundle",
        "noop.docx",
        "modified.docx",
        "noop-visual",
        "modified-visual",
    ):
        candidate = working / name
        if not candidate.exists():
            continue
        target = destination / name
        if candidate.is_dir():
            shutil.copytree(candidate, target, dirs_exist_ok=True)
        else:
            shutil.copy2(candidate, target)
    return str(destination)


def _status_counts(documents: list[dict[str, Any]], key: str) -> dict[str, int]:
    return dict(Counter(document.get(key, "not_run") for document in documents))


def assess_document(
    source: Path,
    corpus: Path,
    failures: Path,
    visual_available: bool,
) -> dict[str, Any]:
    relative = source.relative_to(corpus)
    identity = hashlib.sha256(str(relative).encode()).hexdigest()[:12]
    upstream = (
        source.name.split("__", 1)[0] if "__" in source.name else relative.parts[0]
    )
    result: dict[str, Any] = {
        "id": identity,
        "filename": source.name,
        "source": upstream,
        "relative_path": str(relative),
        "file_size": source.stat().st_size,
        "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "paragraph_count": None,
        "table_count": None,
        "extraction_status": "not_run",
        "profiling_status": "not_run",
        "noop_reconstruction_status": "not_run",
        "package_integrity_status": "not_run",
        "noop_visual_status": "not_run",
        "noop_visual_max_difference_ratio": None,
        "noop_visual_page_ratios": [],
        "text_edit_status": "not_run",
        "text_edit_target": None,
        "modified_visual_status": "not_run",
        "modified_visual_max_difference_ratio": None,
        "modified_visual_page_ratios": [],
        "unsupported_features": [],
        "warnings": [],
        "errors": [],
        "failure_classification": None,
        "failure_artifacts": None,
        "timings_seconds": {},
    }

    with tempfile.TemporaryDirectory(
        prefix=f"docxunravel-corpus-{identity}-"
    ) as temporary:
        working = Path(temporary)
        copied_source = working / "source.docx"
        shutil.copy2(source, copied_source)
        if (
            hashlib.sha256(copied_source.read_bytes()).hexdigest()
            != result["source_sha256"]
        ):
            raise RuntimeError("Isolated source copy hash mismatch")
        bundle = working / "bundle"
        profile = None

        try:
            started = time.perf_counter()
            profile = export_bundle(copied_source, bundle)
            result["timings_seconds"]["extract_and_profile"] = round(
                time.perf_counter() - started, 4
            )
            result["extraction_status"] = "pass"
            result["profiling_status"] = "pass"
            result["paragraph_count"] = profile["summary"]["paragraph_count"]
            result["table_count"] = _table_count(profile)
            result["unsupported_features"] = _package_features(bundle, profile)
        except Exception as error:
            result["extraction_status"] = "fail"
            result["profiling_status"] = "fail"
            result["failure_classification"] = "unsupported_or_malformed_source"
            _stage_error(result, "extract_and_profile", error)

        noop_output = working / "noop.docx"
        noop_result = None
        if profile is not None:
            try:
                started = time.perf_counter()
                noop_result = render_bundle(
                    bundle,
                    noop_output,
                    {"contract": profile["patch_contract"]},
                )
                result["timings_seconds"]["noop_reconstruction"] = round(
                    time.perf_counter() - started, 4
                )
                result["noop_reconstruction_status"] = "pass"
                result["warnings"].extend(noop_result.get("warnings", []))
            except Exception as error:
                result["noop_reconstruction_status"] = "fail"
                result["failure_classification"] = "implementation_failure"
                _stage_error(result, "noop_reconstruction", error)

        if noop_result is not None:
            try:
                integrity = validate_docx_package(noop_output)
                with zipfile.ZipFile(noop_output) as archive:
                    mismatches = [
                        entry["path"]
                        for entry in profile["package"]["entries"]
                        if hashlib.sha256(archive.read(entry["path"])).hexdigest()
                        != entry["sha256"]
                    ]
                if mismatches:
                    raise AssertionError(
                        "No-op changed package parts: " + ", ".join(mismatches)
                    )
                result["package_integrity_status"] = "pass"
                result["package_integrity"] = integrity
            except Exception as error:
                result["package_integrity_status"] = "fail"
                result["failure_classification"] = "implementation_failure"
                _stage_error(result, "noop_package_integrity", error)

        noop_visual = None
        if noop_result is not None and visual_available:
            try:
                started = time.perf_counter()
                noop_visual = verify_visuals(
                    copied_source,
                    noop_output,
                    working / "noop-visual",
                    difference_threshold=0.0,
                    expected_difference_threshold=0.0,
                    pixel_threshold=0,
                    dimension_tolerance_pt=0.0,
                    no_op=True,
                )
                result["timings_seconds"]["noop_visual"] = round(
                    time.perf_counter() - started, 4
                )
                ratios = [page["difference_ratio"] for page in noop_visual["pages"]]
                result["noop_visual_page_ratios"] = ratios
                result["noop_visual_max_difference_ratio"] = max(ratios, default=0.0)
                result["noop_visual_status"] = (
                    "pass" if noop_visual["passed"] else "fail"
                )
                if not noop_visual["passed"]:
                    result["failure_classification"] = "visual_compatibility_failure"
            except Exception as error:
                result["noop_visual_status"] = "error"
                result["failure_classification"] = "visual_renderer_failure"
                _stage_error(result, "noop_visual", error)
        elif not visual_available:
            result["noop_visual_status"] = "unavailable"

        if profile is not None and noop_result is not None:
            target = _safe_paragraph(profile)
            if target is None:
                result["text_edit_status"] = "unsupported"
                if result["failure_classification"] is None:
                    result["failure_classification"] = "unsupported_no_safe_paragraph"
                result["warnings"].append(
                    {
                        "code": "NO_SAFE_EDIT_TARGET",
                        "message": "No plain, non-revision, single-format body paragraph was available.",
                    }
                )
            else:
                original_text = target["text"]
                replacement = original_text + " [DocXUnravel compatibility edit]"
                patch = {
                    "contract": profile["patch_contract"],
                    "elements": {
                        target["id"]: {
                            "expected": original_text,
                            "value": replacement,
                            "format_policy": "reject_mixed_runs",
                        }
                    },
                }
                result["text_edit_target"] = {
                    "id": target["id"],
                    "original": original_text,
                    "replacement": replacement,
                }
                modified = working / "modified.docx"
                try:
                    started = time.perf_counter()
                    edit_result = render_bundle(bundle, modified, patch)
                    fidelity = create_fidelity_report(
                        copied_source,
                        modified,
                        bundle=bundle,
                        patch=patch,
                    )
                    result["timings_seconds"]["text_edit"] = round(
                        time.perf_counter() - started, 4
                    )
                    result["warnings"].extend(edit_result.get("warnings", []))
                    result["text_edit_fidelity"] = fidelity
                    if not fidelity["passed"]:
                        raise AssertionError("Modified-document fidelity report failed")
                    result["text_edit_status"] = "pass"
                except Exception as error:
                    result["text_edit_status"] = "fail"
                    result["failure_classification"] = "implementation_failure"
                    _stage_error(result, "controlled_text_edit", error)

                if result["text_edit_status"] == "pass" and visual_available:
                    try:
                        started = time.perf_counter()
                        modified_visual = verify_visuals(
                            copied_source,
                            modified,
                            working / "modified-visual",
                            difference_threshold=0.0,
                            expected_difference_threshold=0.0,
                            pixel_threshold=0,
                            dimension_tolerance_pt=0.0,
                            expected_changed_pages=set(),
                        )
                        result["timings_seconds"]["modified_visual"] = round(
                            time.perf_counter() - started, 4
                        )
                        ratios = [
                            page["difference_ratio"]
                            for page in modified_visual["pages"]
                        ]
                        result["modified_visual_page_ratios"] = ratios
                        result["modified_visual_max_difference_ratio"] = max(
                            ratios, default=0.0
                        )
                        result["modified_visual_status"] = (
                            "no_difference"
                            if modified_visual["passed"]
                            else "differences_detected"
                        )
                        result["modified_visual_unexpected_pages"] = modified_visual[
                            "unexpected_changed_pages"
                        ]
                    except Exception as error:
                        result["modified_visual_status"] = "error"
                        _stage_error(result, "modified_visual", error)
                elif not visual_available:
                    result["modified_visual_status"] = "unavailable"

        should_preserve = (
            bool(result["errors"]) or result["failure_classification"] is not None
        )
        if should_preserve:
            result["failure_artifacts"] = _preserve_failure(
                source, result, working, failures
            )
    return result


def write_reports(
    documents: list[dict[str, Any]],
    corpus: Path,
    output_json: Path,
    output_csv: Path,
    started: float,
) -> dict[str, Any]:
    stages = [
        "extraction_status",
        "profiling_status",
        "noop_reconstruction_status",
        "package_integrity_status",
        "noop_visual_status",
        "text_edit_status",
        "modified_visual_status",
    ]
    unsupported_counts = Counter(
        feature
        for document in documents
        for feature in document["unsupported_features"]
    )
    report = {
        "schema_version": "1.0",
        "corpus": corpus.name,
        "generated_at_epoch": time.time(),
        "duration_seconds": round(time.perf_counter() - started, 3),
        "visual_dependencies": visual_dependencies(),
        "document_count": len(documents),
        "stage_status_counts": {
            stage: _status_counts(documents, stage) for stage in stages
        },
        "unsupported_feature_counts": dict(unsupported_counts.most_common()),
        "failure_classification_counts": dict(
            Counter(
                document["failure_classification"] or "none" for document in documents
            )
        ),
        "documents": documents,
    }
    output_json.write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    columns = [
        "filename",
        "source",
        "relative_path",
        "file_size",
        "paragraph_count",
        "table_count",
        "extraction_status",
        "profiling_status",
        "noop_reconstruction_status",
        "package_integrity_status",
        "noop_visual_status",
        "noop_visual_max_difference_ratio",
        "text_edit_status",
        "modified_visual_status",
        "modified_visual_max_difference_ratio",
        "unsupported_features",
        "warnings",
        "errors",
        "failure_classification",
        "failure_artifacts",
    ]
    with output_csv.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for document in documents:
            row = {key: document.get(key) for key in columns}
            for key in ("unsupported_features", "warnings", "errors"):
                row[key] = json.dumps(row[key], ensure_ascii=False)
            writer.writerow(row)
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="DocXUnravel public-corpus baseline")
    parser.add_argument("corpus")
    parser.add_argument(
        "--json", default="reports/compatibility/compatibility-report.json"
    )
    parser.add_argument(
        "--csv", default="reports/compatibility/compatibility-report.csv"
    )
    parser.add_argument("--failures", default="compatibility-failures")
    args = parser.parse_args(argv)

    started = time.perf_counter()
    corpus = Path(args.corpus).expanduser().resolve()
    sources = sorted(
        path
        for path in corpus.rglob("*")
        if path.is_file() and path.suffix.lower() == ".docx"
    )
    failures = Path(args.failures)
    if failures.exists():
        shutil.rmtree(failures)
    failures.mkdir(parents=True)
    visual_available = visual_dependencies()["available"]
    documents = []
    for index, source in enumerate(sources, start=1):
        print(f"[{index}/{len(sources)}] {source.name}", flush=True)
        documents.append(assess_document(source, corpus, failures, visual_available))
    report = write_reports(
        documents,
        corpus,
        Path(args.json),
        Path(args.csv),
        started,
    )
    print(
        json.dumps(
            {
                "documents": report["document_count"],
                "duration_seconds": report["duration_seconds"],
                "stage_status_counts": report["stage_status_counts"],
                "failure_classification_counts": report[
                    "failure_classification_counts"
                ],
                "json": args.json,
                "csv": args.csv,
                "failures": args.failures,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
