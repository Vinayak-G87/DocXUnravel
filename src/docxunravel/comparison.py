"""Compact paragraph extraction and semantic document comparison."""

from __future__ import annotations

import json
from collections.abc import Iterable, Iterator
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

CONTENT_SCHEMA_VERSION = "1.0"


def _collect_paragraphs(elements: Iterable[dict[str, Any]]) -> Iterator[dict[str, str]]:
    for element in elements:
        element_type = element.get("type")
        if element_type == "paragraph":
            yield {
                "id": str(element["id"]),
                "text": str(element.get("text", "")),
            }
        elif element_type == "table":
            for row in element.get("rows", []):
                for cell in row.get("cells", []):
                    yield from _collect_paragraphs(cell.get("elements", []))


def compact_paragraph_content(profile: dict[str, Any]) -> dict[str, Any]:
    """Return an agent-oriented artifact containing paragraph IDs and text only."""
    all_paragraphs = list(_collect_paragraphs(profile.get("body", [])))
    for container in profile.get("headers", []) + profile.get("footers", []):
        all_paragraphs.extend(_collect_paragraphs(container.get("elements", [])))
    paragraphs = [
        paragraph for paragraph in all_paragraphs if paragraph["text"].strip()
    ]

    content: dict[str, Any] = {
        "schema_version": CONTENT_SCHEMA_VERSION,
        "artifact": "docxunravel.paragraph-content",
        "source": {
            key: value
            for key, value in profile.get("source", {}).items()
            if key in {"filename", "sha256"}
        },
        "summary": {
            "paragraph_count": len(paragraphs),
            "omitted_empty_paragraph_count": len(all_paragraphs) - len(paragraphs),
        },
        "paragraphs": paragraphs,
    }
    if "patch_contract" in profile:
        content["patch_contract"] = profile["patch_contract"]
    return content


def write_paragraph_content(
    profile: dict[str, Any], output: str | Path
) -> dict[str, Any]:
    """Write and return the compact paragraph-content artifact."""
    content = compact_paragraph_content(profile)
    Path(output).write_text(
        json.dumps(content, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return content


def _read_json(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as handle:
        value = json.load(handle)
    if not isinstance(value, dict):
        raise TypeError(f"JSON root must be an object: {path}")
    return value


def load_paragraph_content(source: str | Path) -> dict[str, Any]:
    """Load compact content from a DOCX, extracted bundle, profile, or content JSON."""
    path = Path(source)
    if path.is_dir():
        content_path = path / "content.json"
        profile_path = path / "profile.json"
        if content_path.is_file():
            value = _read_json(content_path)
        elif profile_path.is_file():
            value = _read_json(profile_path)
        else:
            raise ValueError(
                f"Bundle has neither content.json nor profile.json: {path}"
            )
    elif path.suffix.lower() == ".docx":
        from .core import profile_document

        value = profile_document(path)
    elif path.suffix.lower() in {".md", ".markdown"}:
        from .markdown_ingestion import parse_markdown

        markdown = parse_markdown(path)
        paragraphs: list[dict[str, str]] = []
        for block in markdown["blocks"]:
            block_id = f"markdown.{block['id']}"
            if block["type"] in {"heading", "paragraph"}:
                paragraphs.append({"id": block_id, "text": block["text"]})
            elif block["type"] == "checkbox":
                glyph = "☒" if block["checked"] else "☐"
                paragraphs.append(
                    {"id": block_id, "text": f"{glyph} {block['text']}".rstrip()}
                )
            elif block["type"] == "table":
                for row_index, row in enumerate([block["columns"], *block["rows"]]):
                    for column_index, text in enumerate(row):
                        if text.strip():
                            paragraphs.append(
                                {
                                    "id": f"{block_id}.r{row_index}.c{column_index}",
                                    "text": text,
                                }
                            )
        return {
            "schema_version": CONTENT_SCHEMA_VERSION,
            "artifact": "docxunravel.paragraph-content",
            "source": markdown["source"],
            "summary": {
                "paragraph_count": len(paragraphs),
                "omitted_empty_paragraph_count": 0,
            },
            "paragraphs": paragraphs,
        }
    elif path.is_file():
        value = _read_json(path)
    else:
        raise FileNotFoundError(path)

    if value.get("artifact") == "docxunravel.paragraph-content":
        paragraphs = value.get("paragraphs")
        if not isinstance(paragraphs, list) or any(
            not isinstance(paragraph, dict)
            or not isinstance(paragraph.get("id"), str)
            or not isinstance(paragraph.get("text"), str)
            for paragraph in paragraphs
        ):
            raise ValueError(f"Invalid paragraph-content JSON: {path}")
        return value
    if isinstance(value.get("body"), list):
        return compact_paragraph_content(value)
    raise ValueError(f"Expected a DOCX, bundle, profile.json, or content.json: {path}")


def compare_paragraph_content(
    original: dict[str, Any], revised: dict[str, Any]
) -> dict[str, Any]:
    """Compare two compact/profile artifacts by ordered paragraph text."""
    original_paragraphs = original.get("paragraphs", [])
    revised_paragraphs = revised.get("paragraphs", [])
    original_text = [paragraph["text"] for paragraph in original_paragraphs]
    revised_text = [paragraph["text"] for paragraph in revised_paragraphs]
    matcher = SequenceMatcher(None, original_text, revised_text, autojunk=False)

    changes: list[dict[str, Any]] = []
    unchanged = 0
    added = 0
    removed = 0
    modified = 0

    for (
        tag,
        original_start,
        original_end,
        revised_start,
        revised_end,
    ) in matcher.get_opcodes():
        if tag == "equal":
            unchanged += original_end - original_start
            continue
        if tag == "delete":
            for paragraph in original_paragraphs[original_start:original_end]:
                changes.append({"kind": "removed", "original": paragraph})
                removed += 1
            continue
        if tag == "insert":
            for paragraph in revised_paragraphs[revised_start:revised_end]:
                changes.append({"kind": "added", "revised": paragraph})
                added += 1
            continue

        original_block = original_paragraphs[original_start:original_end]
        revised_block = revised_paragraphs[revised_start:revised_end]
        paired = min(len(original_block), len(revised_block))
        for offset in range(paired):
            changes.append(
                {
                    "kind": "modified",
                    "original": original_block[offset],
                    "revised": revised_block[offset],
                }
            )
            modified += 1
        for paragraph in original_block[paired:]:
            changes.append({"kind": "removed", "original": paragraph})
            removed += 1
        for paragraph in revised_block[paired:]:
            changes.append({"kind": "added", "revised": paragraph})
            added += 1

    return {
        "schema_version": "1.0",
        "comparison": "ordered-paragraph-content",
        "equal": not changes,
        "original": original.get("source", {}),
        "revised": revised.get("source", {}),
        "summary": {
            "original_paragraph_count": len(original_paragraphs),
            "revised_paragraph_count": len(revised_paragraphs),
            "unchanged": unchanged,
            "modified": modified,
            "added": added,
            "removed": removed,
        },
        "changes": changes,
    }


def compare_documents(original: str | Path, revised: str | Path) -> dict[str, Any]:
    """Compare DOCX, bundle, profile, or compact-content inputs."""
    return compare_paragraph_content(
        load_paragraph_content(original), load_paragraph_content(revised)
    )
