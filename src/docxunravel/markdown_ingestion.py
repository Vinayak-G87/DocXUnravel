"""Bounded Markdown ingestion and deterministic JSON template rendering.

This module intentionally implements a small Markdown subset rather than relying
on a general-purpose Markdown renderer.  Supported constructs are ATX headings,
paragraphs, task-list checkboxes, GFM pipe tables, and standalone DocXUnravel
markers::

    <!-- docxunravel:field customer.name -->
    <!-- docxunravel:table line_items -->
    <!-- docxunravel:section-alias summary -->

``section`` is accepted as a shorter spelling of ``section-alias``.  A marker
names the content block immediately following it (blank lines are ignored).
Field markers may name headings, paragraphs, or checkboxes; table markers must
name tables; section aliases must name headings.

The parser rejects images and raw HTML, applies conservative size/count limits,
and returns only JSON-serializable values.  It does not execute or render any
Markdown content.
"""

from __future__ import annotations

import hashlib
import json
import math
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

MARKDOWN_SCHEMA_VERSION = "1.0"


@dataclass(frozen=True, slots=True)
class MarkdownLimits:
    """Resource limits used while reading, parsing, and rendering Markdown."""

    max_bytes: int = 1_048_576
    max_lines: int = 20_000
    max_line_length: int = 16_384
    max_blocks: int = 10_000
    max_table_rows: int = 5_000
    max_table_columns: int = 128
    max_placeholders: int = 10_000


DEFAULT_MARKDOWN_LIMITS = MarkdownLimits()

_MARKER_RE = re.compile(
    r"^\s*<!--\s*docxunravel:(field|table|section(?:-alias)?)\s+"
    r"([A-Za-z_][A-Za-z0-9_.-]*)\s*-->\s*$"
)
_HEADING_RE = re.compile(r"^ {0,3}(#{1,6})(?:[ \t]+(.*?))[ \t]*#*[ \t]*$")
_CHECKBOX_RE = re.compile(r"^ {0,3}[-+*][ \t]+\[([ xX])\](?:[ \t]+(.*))?$")
_SEPARATOR_CELL_RE = re.compile(r"^:?-{3,}:?$")
_RAW_HTML_RE = re.compile(r"<(?:[A-Za-z!/]|\?)")
_IMAGE_RE = re.compile(r"(?<!\\)!\[")
_PLACEHOLDER_RE = re.compile(
    r"{{\s*([A-Za-z_][A-Za-z0-9_-]*(?:\.[A-Za-z0-9_-]+)*)\s*}}"
)
_PLACEHOLDER_LIKE_RE = re.compile(r"{{|}}")


def _check_limits(limits: MarkdownLimits) -> None:
    for name in (
        "max_bytes",
        "max_lines",
        "max_line_length",
        "max_blocks",
        "max_table_rows",
        "max_table_columns",
        "max_placeholders",
    ):
        value = getattr(limits, name)
        if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
            raise ValueError(f"Markdown limit {name} must be a positive integer")


def _decode_markdown(raw: bytes, label: str, limits: MarkdownLimits) -> str:
    if len(raw) > limits.max_bytes:
        raise ValueError(f"Markdown input exceeds {limits.max_bytes} bytes: {label}")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError(f"Markdown input is not valid UTF-8: {label}") from exc
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    if len(lines) > limits.max_lines:
        raise ValueError(f"Markdown input exceeds {limits.max_lines} lines: {label}")
    for number, line in enumerate(lines, 1):
        if len(line) > limits.max_line_length:
            raise ValueError(
                f"Markdown line {number} exceeds {limits.max_line_length} characters"
            )
    return text


def _read_bounded(path: Path, limits: MarkdownLimits) -> bytes:
    if not path.is_file():
        raise FileNotFoundError(path)
    size = path.stat().st_size
    if size > limits.max_bytes:
        raise ValueError(f"Markdown input exceeds {limits.max_bytes} bytes: {path}")
    with path.open("rb") as handle:
        raw = handle.read(limits.max_bytes + 1)
    if len(raw) > limits.max_bytes:
        raise ValueError(f"Markdown input exceeds {limits.max_bytes} bytes: {path}")
    return raw


def _split_pipe_row(line: str) -> list[str] | None:
    """Split one GFM pipe-table row while preserving non-pipe escapes."""

    stripped = line.strip()
    if "|" not in stripped:
        return None
    cells: list[str] = []
    current: list[str] = []
    escaped = False
    for character in stripped:
        if escaped:
            if character == "|":
                current.append("|")
            else:
                current.extend(("\\", character))
            escaped = False
        elif character == "\\":
            escaped = True
        elif character == "|":
            cells.append("".join(current).strip())
            current = []
        else:
            current.append(character)
    if escaped:
        current.append("\\")
    cells.append("".join(current).strip())
    if stripped.startswith("|"):
        cells = cells[1:]
    if stripped.endswith("|") and cells:
        cells = cells[:-1]
    return cells if len(cells) >= 2 else None


def _table_start(lines: list[str], index: int) -> tuple[list[str], list[str]] | None:
    if index + 1 >= len(lines):
        return None
    header = _split_pipe_row(lines[index])
    separator = _split_pipe_row(lines[index + 1])
    if header is None or separator is None or len(header) != len(separator):
        return None
    if not all(_SEPARATOR_CELL_RE.fullmatch(cell) for cell in separator):
        return None
    return header, separator


def _validate_source_syntax(lines: list[str]) -> None:
    for number, line in enumerate(lines, 1):
        if _MARKER_RE.fullmatch(line):
            continue
        if _IMAGE_RE.search(line):
            raise ValueError(f"Markdown images are not allowed (line {number})")
        if _RAW_HTML_RE.search(line):
            raise ValueError(f"Raw HTML is not allowed (line {number})")


def _canonical_source(label: str, raw: bytes) -> dict[str, str]:
    return {
        "name": label,
        "sha256": hashlib.sha256(raw).hexdigest(),
    }


def _parse_text(
    text: str,
    *,
    source: dict[str, str],
    limits: MarkdownLimits,
) -> dict[str, Any]:
    lines = text.split("\n")
    _validate_source_syntax(lines)

    blocks: list[dict[str, Any]] = []
    fields: dict[str, dict[str, Any]] = {}
    tables: dict[str, dict[str, Any]] = {}
    checkboxes: dict[str, dict[str, Any]] = {}
    sections: dict[str, dict[str, Any]] = {}
    marker_keys: set[tuple[str, str]] = set()
    pending: tuple[str, str, int] | None = None
    index = 0

    def append_block(block: dict[str, Any], line_number: int) -> None:
        nonlocal pending
        if len(blocks) >= limits.max_blocks:
            raise ValueError(f"Markdown exceeds {limits.max_blocks} content blocks")
        block_id = f"block-{len(blocks) + 1:06d}"
        block["id"] = block_id
        block["line"] = line_number

        if pending is not None:
            kind, name, marker_line = pending
            block_type = block["type"]
            if kind == "table" and block_type != "table":
                raise ValueError(
                    f"Table marker {name!r} on line {marker_line} must precede a table"
                )
            if kind == "section" and block_type != "heading":
                raise ValueError(
                    f"Section alias {name!r} on line {marker_line} must precede a heading"
                )
            if kind == "field" and block_type not in {
                "heading",
                "paragraph",
                "checkbox",
            }:
                raise ValueError(
                    f"Field marker {name!r} on line {marker_line} cannot name a {block_type}"
                )
            block["name"] = name
            if kind == "table":
                tables[name] = {
                    "block_id": block_id,
                    "columns": list(block["columns"]),
                    "rows": [list(row) for row in block["rows"]],
                }
            elif kind == "section":
                sections[name] = {
                    "block_id": block_id,
                    "level": block["level"],
                    "title": block["text"],
                }
            else:
                if block_type == "checkbox":
                    value: str | bool = block["checked"]
                    checkboxes[name] = {
                        "block_id": block_id,
                        "checked": block["checked"],
                        "text": block["text"],
                    }
                else:
                    value = block["text"]
                fields[name] = {
                    "block_id": block_id,
                    "type": block_type,
                    "value": value,
                }
            pending = None
        blocks.append(block)

    while index < len(lines):
        line = lines[index]
        if not line.strip():
            index += 1
            continue

        marker = _MARKER_RE.fullmatch(line)
        if marker:
            marker_kind = marker.group(1)
            kind = "section" if marker_kind.startswith("section") else marker_kind
            name = marker.group(2)
            key = (kind, name)
            if key in marker_keys:
                raise ValueError(
                    f"Duplicate {kind} marker {name!r} on line {index + 1}"
                )
            if pending is not None:
                raise ValueError(
                    f"Marker on line {pending[2]} has no content before line {index + 1}"
                )
            marker_keys.add(key)
            pending = (kind, name, index + 1)
            index += 1
            continue

        heading = _HEADING_RE.fullmatch(line)
        if heading:
            append_block(
                {
                    "type": "heading",
                    "level": len(heading.group(1)),
                    "text": heading.group(2).strip(),
                },
                index + 1,
            )
            index += 1
            continue

        checkbox = _CHECKBOX_RE.fullmatch(line)
        if checkbox:
            append_block(
                {
                    "type": "checkbox",
                    "checked": checkbox.group(1).lower() == "x",
                    "text": (checkbox.group(2) or "").strip(),
                },
                index + 1,
            )
            index += 1
            continue

        table_start = _table_start(lines, index)
        if table_start is not None:
            columns, separator = table_start
            if len(columns) > limits.max_table_columns:
                raise ValueError(
                    f"Table on line {index + 1} exceeds {limits.max_table_columns} columns"
                )
            if any(not column for column in columns):
                raise ValueError(f"Table on line {index + 1} has an empty column name")
            if len(set(columns)) != len(columns):
                raise ValueError(
                    f"Table on line {index + 1} has duplicate column names"
                )
            alignments = [
                "center"
                if cell.startswith(":") and cell.endswith(":")
                else "right"
                if cell.endswith(":")
                else "left"
                if cell.startswith(":")
                else "default"
                for cell in separator
            ]
            rows: list[list[str]] = []
            cursor = index + 2
            while cursor < len(lines) and lines[cursor].strip():
                row = _split_pipe_row(lines[cursor])
                if row is None:
                    break
                if len(row) != len(columns):
                    raise ValueError(
                        f"Table row on line {cursor + 1} has {len(row)} cells; "
                        f"expected {len(columns)}"
                    )
                if len(rows) >= limits.max_table_rows:
                    raise ValueError(f"Table exceeds {limits.max_table_rows} data rows")
                rows.append(row)
                cursor += 1
            append_block(
                {
                    "type": "table",
                    "columns": columns,
                    "alignments": alignments,
                    "rows": rows,
                },
                index + 1,
            )
            index = cursor
            continue

        paragraph_start = index
        paragraph_lines: list[str] = []
        while index < len(lines) and lines[index].strip():
            if index != paragraph_start and (
                _MARKER_RE.fullmatch(lines[index])
                or _HEADING_RE.fullmatch(lines[index])
                or _CHECKBOX_RE.fullmatch(lines[index])
                or _table_start(lines, index) is not None
            ):
                break
            paragraph_lines.append(lines[index].strip())
            index += 1
        append_block(
            {"type": "paragraph", "text": "\n".join(paragraph_lines)},
            paragraph_start + 1,
        )

    if pending is not None:
        raise ValueError(f"Marker on line {pending[2]} has no following content")

    return {
        "schema_version": MARKDOWN_SCHEMA_VERSION,
        "artifact": "docxunravel.markdown",
        "source": source,
        "fields": fields,
        "tables": tables,
        "checkboxes": checkboxes,
        "sections": sections,
        "blocks": blocks,
    }


def parse_markdown_text(
    text: str,
    *,
    source_name: str = "<markdown>",
    limits: MarkdownLimits = DEFAULT_MARKDOWN_LIMITS,
) -> dict[str, Any]:
    """Parse Markdown text into the canonical bounded model."""

    _check_limits(limits)
    if not isinstance(text, str):
        raise TypeError("Markdown input must be text")
    raw = text.encode("utf-8")
    normalized = _decode_markdown(raw, source_name, limits)
    return _parse_text(
        normalized,
        source=_canonical_source(source_name, raw),
        limits=limits,
    )


def parse_markdown(
    path: str | Path,
    *,
    limits: MarkdownLimits = DEFAULT_MARKDOWN_LIMITS,
) -> dict[str, Any]:
    """Parse a UTF-8 Markdown file into a canonical JSON-serializable model.

    The source SHA-256 is calculated over the original bytes.  Line endings and
    an optional UTF-8 BOM are normalized only for parsing.
    """

    _check_limits(limits)
    source_path = Path(path)
    raw = _read_bounded(source_path, limits)
    text = _decode_markdown(raw, str(source_path), limits)
    return _parse_text(
        text,
        source=_canonical_source(source_path.name, raw),
        limits=limits,
    )


def markdown_to_data(
    path: str | Path,
    *,
    limits: MarkdownLimits = DEFAULT_MARKDOWN_LIMITS,
) -> dict[str, Any]:
    """Return named Markdown values in a compact, template-friendly form."""

    model = parse_markdown(path, limits=limits)
    return {
        "schema_version": MARKDOWN_SCHEMA_VERSION,
        "artifact": "docxunravel.markdown-data",
        "source": model["source"],
        "fields": {name: field["value"] for name, field in model["fields"].items()},
        "tables": {
            name: {
                "columns": table["columns"],
                "rows": [
                    dict(zip(table["columns"], row, strict=True))
                    for row in table["rows"]
                ],
            }
            for name, table in model["tables"].items()
        },
        "checkboxes": {
            name: checkbox["checked"] for name, checkbox in model["checkboxes"].items()
        },
        "sections": {
            name: section["title"] for name, section in model["sections"].items()
        },
    }


def _block_signature(block: Mapping[str, Any]) -> str:
    comparable = {
        key: value for key, value in block.items() if key not in {"id", "line"}
    }
    return json.dumps(
        comparable, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )


def compare_markdown(
    original: str | Path,
    revised: str | Path,
    *,
    limits: MarkdownLimits = DEFAULT_MARKDOWN_LIMITS,
) -> dict[str, Any]:
    """Compare two Markdown files by their ordered canonical content blocks."""

    original_model = parse_markdown(original, limits=limits)
    revised_model = parse_markdown(revised, limits=limits)
    original_blocks = original_model["blocks"]
    revised_blocks = revised_model["blocks"]
    matcher = SequenceMatcher(
        None,
        [_block_signature(block) for block in original_blocks],
        [_block_signature(block) for block in revised_blocks],
        autojunk=False,
    )
    changes: list[dict[str, Any]] = []
    counts = {"unchanged": 0, "modified": 0, "added": 0, "removed": 0}

    for tag, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        if tag == "equal":
            counts["unchanged"] += old_end - old_start
            continue
        old_group = original_blocks[old_start:old_end]
        new_group = revised_blocks[new_start:new_end]
        if tag == "replace":
            paired = min(len(old_group), len(new_group))
            for offset in range(paired):
                changes.append(
                    {
                        "kind": "modified",
                        "original": old_group[offset],
                        "revised": new_group[offset],
                    }
                )
                counts["modified"] += 1
            old_group = old_group[paired:]
            new_group = new_group[paired:]
        for block in old_group:
            changes.append({"kind": "removed", "original": block})
            counts["removed"] += 1
        for block in new_group:
            changes.append({"kind": "added", "revised": block})
            counts["added"] += 1

    return {
        "schema_version": MARKDOWN_SCHEMA_VERSION,
        "artifact": "docxunravel.markdown-comparison",
        "equal": not changes,
        "original": original_model["source"],
        "revised": revised_model["source"],
        "summary": {
            "original_block_count": len(original_blocks),
            "revised_block_count": len(revised_blocks),
            **counts,
        },
        "changes": changes,
    }


def _path_parts(path: str) -> tuple[str, ...]:
    return tuple(path.split("."))


def _resolve_path(value: Any, parts: Sequence[str], display_path: str) -> Any:
    current = value
    for part in parts:
        if isinstance(current, Mapping):
            if part not in current:
                raise KeyError(f"Missing template value {display_path!r}")
            current = current[part]
        elif (
            isinstance(current, Sequence)
            and not isinstance(current, (str, bytes, bytearray))
            and part.isdigit()
        ):
            position = int(part)
            if position >= len(current):
                raise KeyError(f"Missing template value {display_path!r}")
            current = current[position]
        else:
            raise KeyError(f"Missing template value {display_path!r}")
    return current


def _list_prefix(data: Mapping[str, Any], path: str) -> tuple[str, ...] | None:
    parts = _path_parts(path)
    current: Any = data
    for index, part in enumerate(parts):
        if not isinstance(current, Mapping) or part not in current:
            return None
        current = current[part]
        if isinstance(current, list):
            return parts[: index + 1]
    return None


def _scalar_text(value: Any, *, table_cell: bool) -> str:
    if value is None:
        text = ""
    elif value is True:
        text = "true"
    elif value is False:
        text = "false"
    elif isinstance(value, str):
        text = value
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("Template numbers must be finite")
        text = json.dumps(value, ensure_ascii=False, allow_nan=False)
    else:
        text = json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        )
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
    text = text.replace("\\", "\\\\")
    for character in ("`", "*", "_", "{", "}", "[", "]", "<", ">", "!"):
        text = text.replace(character, f"\\{character}")
    if table_cell:
        text = text.replace("|", "\\|")
    return text


def _substitute_line(
    line: str,
    data: Mapping[str, Any],
    *,
    table_cell: bool,
    repeated_prefix: tuple[str, ...] | None = None,
    repeated_item: Any = None,
) -> str:
    def replacement(match: re.Match[str]) -> str:
        path = match.group(1)
        parts = _path_parts(path)
        if (
            repeated_prefix is not None
            and parts[: len(repeated_prefix)] == repeated_prefix
        ):
            remainder = parts[len(repeated_prefix) :]
            value = (
                repeated_item
                if not remainder
                else _resolve_path(repeated_item, remainder, path)
            )
        else:
            value = _resolve_path(data, parts, path)
        if isinstance(value, list):
            raise TypeError(
                f"List template value {path!r} may only be used in a repeatable table row"
            )
        return _scalar_text(value, table_cell=table_cell)

    return _PLACEHOLDER_RE.sub(replacement, line)


def _table_body_lines(lines: list[str]) -> set[int]:
    body: set[int] = set()
    index = 0
    while index < len(lines):
        if _table_start(lines, index) is None:
            index += 1
            continue
        cursor = index + 2
        while cursor < len(lines) and lines[cursor].strip():
            if _split_pipe_row(lines[cursor]) is None:
                break
            body.add(cursor)
            cursor += 1
        index = cursor
    return body


def render_markdown_template_text(
    template: str,
    data: Mapping[str, Any],
    *,
    limits: MarkdownLimits = DEFAULT_MARKDOWN_LIMITS,
) -> str:
    """Render Markdown template text from a JSON-compatible mapping.

    A GFM table body row is repeated when one or more placeholders traverse the
    same JSON list.  For example, ``{{items.name}}`` and ``{{items.price}}`` in
    one prototype row render once per object in ``data["items"]``.  Missing
    values, multiple list sources in one row, malformed placeholders, and list
    placeholders outside table body rows are rejected.
    """

    _check_limits(limits)
    if not isinstance(template, str):
        raise TypeError("Markdown template must be text")
    if not isinstance(data, Mapping):
        raise TypeError("Markdown template data must be a JSON object")
    raw = template.encode("utf-8")
    normalized = _decode_markdown(raw, "<template>", limits)
    has_final_newline = normalized.endswith("\n")
    lines = normalized.split("\n")
    if has_final_newline:
        lines.pop()
    table_body = _table_body_lines(lines)
    placeholder_count = sum(len(_PLACEHOLDER_RE.findall(line)) for line in lines)
    if placeholder_count > limits.max_placeholders:
        raise ValueError(f"Template exceeds {limits.max_placeholders} placeholders")

    rendered: list[str] = []
    for index, line in enumerate(lines):
        matches = list(_PLACEHOLDER_RE.finditer(line))
        residue = _PLACEHOLDER_RE.sub("", line)
        if _PLACEHOLDER_LIKE_RE.search(residue):
            raise ValueError(f"Malformed template placeholder on line {index + 1}")
        prefixes = {
            prefix
            for match in matches
            if (prefix := _list_prefix(data, match.group(1))) is not None
        }
        if not prefixes:
            rendered.append(
                _substitute_line(line, data, table_cell=index in table_body)
            )
            continue
        if index not in table_body:
            raise ValueError(
                f"Repeatable list placeholder must be in a table body row (line {index + 1})"
            )
        if len(prefixes) != 1:
            raise ValueError(
                f"Table prototype row on line {index + 1} references multiple lists"
            )
        prefix = next(iter(prefixes))
        items = _resolve_path(data, prefix, ".".join(prefix))
        if not isinstance(items, list):
            raise TypeError(f"Repeatable value {'.'.join(prefix)!r} must be a list")
        if len(rendered) + len(items) > limits.max_lines:
            raise ValueError(f"Rendered Markdown exceeds {limits.max_lines} lines")
        for item in items:
            rendered.append(
                _substitute_line(
                    line,
                    data,
                    table_cell=True,
                    repeated_prefix=prefix,
                    repeated_item=item,
                )
            )

    result = "\n".join(rendered) + ("\n" if has_final_newline else "")
    result_raw = result.encode("utf-8")
    _decode_markdown(result_raw, "<rendered-template>", limits)
    _parse_text(
        result.replace("\r\n", "\n").replace("\r", "\n"),
        source=_canonical_source("<rendered-template>", result_raw),
        limits=limits,
    )
    return result


def _load_template(template: str | Path, limits: MarkdownLimits) -> str:
    if isinstance(template, Path):
        raw = _read_bounded(template, limits)
        return _decode_markdown(raw, str(template), limits)
    if not isinstance(template, str):
        raise TypeError("Markdown template must be text or a path")
    try:
        candidate = Path(template)
        if candidate.is_file():
            raw = _read_bounded(candidate, limits)
            return _decode_markdown(raw, str(candidate), limits)
    except OSError:
        pass
    return template


def _load_json_object(
    data: Mapping[str, Any] | str | Path,
    limits: MarkdownLimits,
) -> Mapping[str, Any]:
    if isinstance(data, Mapping):
        return data
    if isinstance(data, Path):
        raw = _read_bounded(data, limits)
        label = str(data)
    elif isinstance(data, str):
        try:
            candidate = Path(data)
            if candidate.is_file():
                raw = _read_bounded(candidate, limits)
                label = str(candidate)
            else:
                raw = data.encode("utf-8")
                label = "<json>"
        except OSError:
            raw = data.encode("utf-8")
            label = "<json>"
    else:
        raise TypeError("Template data must be a JSON object, JSON text, or a path")
    if len(raw) > limits.max_bytes:
        raise ValueError(f"JSON input exceeds {limits.max_bytes} bytes: {label}")
    try:
        value = json.loads(raw.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Invalid UTF-8 JSON object: {label}") from exc
    if not isinstance(value, dict):
        raise TypeError(f"JSON template data root must be an object: {label}")
    return value


def render_markdown_template(
    template: str | Path,
    data: Mapping[str, Any] | str | Path,
    *,
    limits: MarkdownLimits = DEFAULT_MARKDOWN_LIMITS,
) -> str:
    """Render Markdown text from template text/path and JSON object/text/path.

    The function returns text and performs no filesystem writes.  Existing string
    paths are read as files; other strings are treated as literal template or
    JSON text respectively.
    """

    _check_limits(limits)
    template_text = _load_template(template, limits)
    json_data = _load_json_object(data, limits)
    return render_markdown_template_text(template_text, json_data, limits=limits)


def populate_markdown_template(
    template: str | Path,
    data: Mapping[str, Any] | str | Path,
    *,
    limits: MarkdownLimits = DEFAULT_MARKDOWN_LIMITS,
) -> str:
    """Alias for :func:`render_markdown_template` for high-level API symmetry."""

    return render_markdown_template(template, data, limits=limits)


__all__ = [
    "DEFAULT_MARKDOWN_LIMITS",
    "MARKDOWN_SCHEMA_VERSION",
    "MarkdownLimits",
    "compare_markdown",
    "markdown_to_data",
    "parse_markdown",
    "parse_markdown_text",
    "populate_markdown_template",
    "render_markdown_template",
    "render_markdown_template_text",
]
