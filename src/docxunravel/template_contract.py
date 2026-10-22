"""Semantic template contracts for deterministic DocXUnravel population.

A template contract maps business-facing field names to semantic profile IDs.
It is intentionally separate from the schema 1.1 patch contract and never
contains OOXML paths or binary package content.
"""

from __future__ import annotations

import copy
import json
import re
import tempfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from .core import (
    FORMAT_POLICIES,
    SCHEMA_VERSION,
    TOKEN_PATTERN,
    _profile_index,
    _semantic_profile_hash,
    create_fidelity_report,
    preview_bundle,
    render_bundle,
    validate_patch,
)
from .migration_discovery import validate_migration_discovery

TEMPLATE_CONTRACT_VERSION = "0.1"
FIELD_TYPES = {"text", "table_cell", "table_rows", "placeholder"}
SEMANTIC_NAME_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.-]*$")
FIELD_KEYS = {
    "type",
    "target",
    "required",
    "description",
    "data_path",
    "format_policy",
    "token",
    "template_row",
    "columns",
}


def _require_object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise TypeError(f"{label} must be an object")
    return value


def _validate_binding(profile: dict[str, Any], binding: Any) -> None:
    binding = _require_object(binding, "Template contract binding")
    expected_keys = {"schema_version", "source_sha256", "profile_sha256"}
    unknown = set(binding) - expected_keys
    missing = expected_keys - set(binding)
    if unknown:
        raise ValueError(
            f"Unknown template contract binding fields: {', '.join(sorted(unknown))}"
        )
    if missing:
        raise ValueError(
            f"Missing template contract binding fields: {', '.join(sorted(missing))}"
        )
    stored = profile.get("patch_contract")
    if not isinstance(stored, dict):
        raise TypeError("Profile has no source/profile binding; export a bundle first")
    if stored.get("source_sha256") != profile.get("source", {}).get("sha256"):
        raise ValueError("Profile source hash and stored binding disagree")
    if stored.get("profile_sha256") != _semantic_profile_hash(profile):
        raise ValueError("Semantic profile is stale or has been modified")
    if binding.get("schema_version") != SCHEMA_VERSION:
        raise ValueError(f"Template contract requires profile schema {SCHEMA_VERSION}")
    if binding != stored:
        raise ValueError("Template contract binding does not match this bundle profile")


def _validate_data_path(value: Any, field_name: str) -> None:
    if value is None:
        return
    if not isinstance(value, str) or (value and not value.startswith("/")):
        raise ValueError(
            f"Field {field_name!r} data_path must be an RFC 6901 JSON pointer"
        )
    for segment in value.split("/")[1:]:
        if "~" in re.sub(r"~[01]", "", segment):
            raise ValueError(
                f"Field {field_name!r} data_path contains an invalid escape"
            )


def _validate_columns(
    field_name: str, field: dict[str, Any], target: dict[str, Any]
) -> None:
    columns = field.get("columns")
    if not isinstance(columns, dict) or not columns:
        raise ValueError(f"Table-row field {field_name!r} requires column mappings")
    seen_indexes: set[int] = set()
    for name, index in columns.items():
        if not isinstance(name, str) or not SEMANTIC_NAME_PATTERN.fullmatch(name):
            raise ValueError(
                f"Table-row field {field_name!r} has invalid column name {name!r}"
            )
        if not isinstance(index, int) or isinstance(index, bool) or index < 0:
            raise TypeError(
                f"Table-row field {field_name!r} column indexes must be non-negative integers"
            )
        if index in seen_indexes:
            raise ValueError(
                f"Table-row field {field_name!r} maps multiple values to column {index}"
            )
        if index >= target.get("column_count", 0):
            raise ValueError(
                f"Table-row field {field_name!r} column {index} is outside {target['id']}"
            )
        seen_indexes.add(index)
    if seen_indexes != set(range(target.get("column_count", 0))):
        raise ValueError(
            f"Table-row field {field_name!r} must map every target-table column"
        )
    template_row = field.get("template_row")
    if (
        not isinstance(template_row, int)
        or isinstance(template_row, bool)
        or template_row < 0
        or template_row >= target.get("row_count", 0)
    ):
        raise ValueError(
            f"Table-row field {field_name!r} template_row does not exist in {target['id']}"
        )


def validate_template_contract(
    profile: dict[str, Any], contract: dict[str, Any]
) -> dict[str, Any]:
    """Validate and return a copy with every target canonicalized.

    The function performs structural validation without modifying either input.
    It deliberately validates against the supplied semantic profile rather than
    accepting physical OOXML paths from callers.
    """

    profile = _require_object(profile, "Profile")
    contract = _require_object(contract, "Template contract")
    allowed_root = {"contract_version", "document_type", "binding", "fields"}
    unknown_root = set(contract) - allowed_root
    if unknown_root:
        raise ValueError(
            f"Unknown template contract fields: {', '.join(sorted(unknown_root))}"
        )
    missing_root = allowed_root - set(contract)
    if missing_root:
        raise ValueError(
            f"Missing template contract fields: {', '.join(sorted(missing_root))}"
        )
    if contract.get("contract_version") != TEMPLATE_CONTRACT_VERSION:
        raise ValueError(
            f"Unsupported template contract version: {contract.get('contract_version')!r}"
        )
    document_type = contract.get("document_type")
    if not isinstance(document_type, str) or not SEMANTIC_NAME_PATTERN.fullmatch(
        document_type
    ):
        raise ValueError("document_type must be a semantic identifier")
    _validate_binding(profile, contract.get("binding"))

    fields = _require_object(contract.get("fields"), "Template contract fields")
    if not fields:
        raise ValueError("Template contract must define at least one field")
    index = _profile_index(profile)
    normalized = copy.deepcopy(contract)
    seen_targets: dict[str, str] = {}
    seen_tokens: dict[str, str] = {}

    for field_name, field_value in fields.items():
        if not isinstance(field_name, str) or not SEMANTIC_NAME_PATTERN.fullmatch(
            field_name
        ):
            raise ValueError(f"Invalid semantic field name: {field_name!r}")
        field = _require_object(field_value, f"Field {field_name!r}")
        unknown = set(field) - FIELD_KEYS
        if unknown:
            raise ValueError(
                f"Field {field_name!r} has unknown properties: {', '.join(sorted(unknown))}"
            )
        required_keys = {"type", "target", "required"}
        missing = required_keys - set(field)
        if missing:
            raise ValueError(
                f"Field {field_name!r} is missing: {', '.join(sorted(missing))}"
            )
        field_type = field.get("type")
        if field_type not in FIELD_TYPES:
            raise ValueError(
                f"Field {field_name!r} has unsupported type {field_type!r}"
            )
        if not isinstance(field.get("required"), bool):
            raise TypeError(f"Field {field_name!r} required must be a boolean")
        if field.get("description") is not None and not isinstance(
            field["description"], str
        ):
            raise TypeError(f"Field {field_name!r} description must be a string")
        _validate_data_path(field.get("data_path"), field_name)

        target_id = field.get("target")
        if not isinstance(target_id, str):
            raise TypeError(f"Field {field_name!r} target must be a semantic ID")
        target = index.get(target_id)
        if target is None:
            raise KeyError(f"Field {field_name!r} has unknown target ID: {target_id}")
        canonical_target = target["id"]
        if canonical_target in seen_targets:
            raise ValueError(
                f"Fields {seen_targets[canonical_target]!r} and {field_name!r} "
                f"map to the same target {canonical_target}"
            )
        for other_target, other_field in seen_targets.items():
            table_target = canonical_target if target.get("type") == "table" else None
            other_entry = index[other_target]
            other_table = other_target if other_entry.get("type") == "table" else None
            if table_target and other_target.startswith(table_target + "."):
                raise ValueError(
                    f"Fields {field_name!r} and {other_field!r} conflict at {table_target}"
                )
            if other_table and canonical_target.startswith(other_table + "."):
                raise ValueError(
                    f"Fields {other_field!r} and {field_name!r} conflict at {other_table}"
                )
        seen_targets[canonical_target] = field_name
        normalized["fields"][field_name]["target"] = canonical_target

        expected_type = {
            "text": "paragraph",
            "table_cell": "cell",
            "table_rows": "table",
            "placeholder": "paragraph",
        }[field_type]
        if target.get("type") != expected_type:
            raise ValueError(
                f"Field {field_name!r} of type {field_type!r} requires a "
                f"{expected_type} target, not {target.get('type')!r}"
            )

        policy = field.get("format_policy", "preserve_first_run")
        if field_type in {"text", "table_cell"}:
            if policy not in FORMAT_POLICIES:
                raise ValueError(
                    f"Field {field_name!r} has unsupported format policy {policy!r}"
                )
        elif "format_policy" in field:
            raise ValueError(
                f"Field {field_name!r} cannot define format_policy for {field_type}"
            )

        if field_type == "placeholder":
            token = field.get("token")
            if not isinstance(token, str) or not TOKEN_PATTERN.fullmatch(
                "{{" + token + "}}"
            ):
                raise ValueError(f"Field {field_name!r} requires a valid token name")
            if token not in target.get("tokens", []):
                raise ValueError(
                    f"Placeholder token {token!r} is not present in {canonical_target}"
                )
            token_targets = {
                entry["id"]
                for entry in index.values()
                if entry.get("type") == "paragraph" and token in entry.get("tokens", [])
            }
            if token_targets != {canonical_target}:
                raise ValueError(
                    f"Placeholder token {token!r} is not unique to {canonical_target}"
                )
            if token in seen_tokens:
                raise ValueError(
                    f"Fields {seen_tokens[token]!r} and {field_name!r} map token {token!r}"
                )
            seen_tokens[token] = field_name
        elif "token" in field:
            raise ValueError(
                f"Field {field_name!r} token is only valid for placeholders"
            )

        if field_type == "table_rows":
            _validate_columns(field_name, field, target)
        elif "columns" in field or "template_row" in field:
            raise ValueError(
                f"Field {field_name!r} row settings are only valid for table_rows"
            )

    return normalized


def template_contract_binding(profile: Mapping[str, Any]) -> dict[str, str]:
    """Return a copy of the exact source/profile binding for contract creation."""

    stored = profile.get("patch_contract")
    if not isinstance(stored, dict):
        raise TypeError("Profile has no source/profile binding; export a bundle first")
    return {
        "schema_version": str(stored.get("schema_version", "")),
        "source_sha256": str(stored.get("source_sha256", "")),
        "profile_sha256": str(stored.get("profile_sha256", "")),
    }


_MISSING = object()


def _resolve_data_path(data: Any, path: str, field_name: str) -> Any:
    if not path:
        return data
    current = data
    for encoded_segment in path.split("/")[1:]:
        segment = encoded_segment.replace("~1", "/").replace("~0", "~")
        if isinstance(current, dict) and segment in current:
            current = current[segment]
        elif (
            isinstance(current, list)
            and segment.isdigit()
            and int(segment) < len(current)
        ):
            current = current[int(segment)]
        else:
            return _MISSING
    return current


def _field_value(data: dict[str, Any], field_name: str, field: dict[str, Any]) -> Any:
    if "data_path" in field:
        return _resolve_data_path(data, field["data_path"], field_name)
    return data.get(field_name, _MISSING)


def _table_cell_text(value: Any, field_name: str, row_index: int) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list) and all(isinstance(item, str) for item in value):
        return ", ".join(value)
    if value is None:
        return ""
    raise TypeError(
        f"Table-row field {field_name!r} row {row_index} contains an unsupported value"
    )


def _cell_paragraph(cell: dict[str, Any], field_name: str) -> dict[str, Any]:
    elements = cell.get("elements", [])
    if any(element.get("type") == "table" for element in elements):
        raise ValueError(
            f"Table-cell field {field_name!r} cannot target a nested table"
        )
    paragraphs = [element for element in elements if element.get("type") == "paragraph"]
    if len(paragraphs) != 1:
        raise ValueError(
            f"Table-cell field {field_name!r} requires exactly one cell paragraph"
        )
    return paragraphs[0]


def _compile_template_population(
    profile: dict[str, Any],
    contract: dict[str, Any],
    data: dict[str, Any],
    *,
    validate_patch_output: bool = True,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    if not isinstance(data, dict):
        raise TypeError("Template population data must be an object")
    normalized = validate_template_contract(profile, contract)
    index = _profile_index(profile)
    patch: dict[str, Any] = {"contract": copy.deepcopy(profile["patch_contract"])}
    elements: dict[str, Any] = {}
    tokens: dict[str, str] = {}
    operations: list[dict[str, Any]] = []
    changes: list[dict[str, Any]] = []

    for field_name, field in normalized["fields"].items():
        value = _field_value(data, field_name, field)
        if value is _MISSING:
            if field["required"]:
                raise ValueError(f"Missing required template field: {field_name}")
            continue
        if value is None:
            if field["required"]:
                raise ValueError(f"Required template field is null: {field_name}")
            value = "" if field["type"] != "table_rows" else []
        target = index[field["target"]]
        field_type = field["type"]
        original: Any
        proposed: Any
        operation_type: str

        if field_type == "text":
            if not isinstance(value, str):
                raise TypeError(f"Text field {field_name!r} must be a string")
            original = target.get("text", "")
            proposed = value
            operation_type = "replace_text"
            elements[target["id"]] = {
                "expected": original,
                "value": value,
                "format_policy": field.get("format_policy", "preserve_first_run"),
            }
        elif field_type == "table_cell":
            if not isinstance(value, str):
                raise TypeError(f"Table-cell field {field_name!r} must be a string")
            paragraph = _cell_paragraph(target, field_name)
            original = paragraph.get("text", "")
            proposed = value
            operation_type = "replace_table_cell_text"
            elements[paragraph["id"]] = {
                "expected": original,
                "value": value,
                "format_policy": field.get("format_policy", "preserve_first_run"),
            }
        elif field_type == "placeholder":
            if not isinstance(value, str):
                raise TypeError(f"Placeholder field {field_name!r} must be a string")
            original = target.get("text", "")
            proposed = value
            operation_type = "substitute_placeholder"
            tokens[field["token"]] = value
        else:
            if not isinstance(value, list):
                raise TypeError(f"Table-row field {field_name!r} must be an array")
            ordered_columns = sorted(field["columns"].items(), key=lambda item: item[1])
            compiled_rows: list[list[str]] = []
            for row_index, row in enumerate(value):
                if not isinstance(row, dict):
                    raise TypeError(
                        f"Table-row field {field_name!r} row {row_index} must be an object"
                    )
                missing = set(field["columns"]) - set(row)
                if missing:
                    raise ValueError(
                        f"Table-row field {field_name!r} row {row_index}: "
                        f"missing {', '.join(sorted(missing))}"
                    )
                values = [
                    _table_cell_text(row[name], field_name, row_index)
                    for name, _ in ordered_columns
                ]
                compiled_rows.append(values)
            template = target["rows"][field["template_row"]]
            original = [cell.get("text", "") for cell in template["cells"]]
            proposed = compiled_rows
            operation_type = "repeat_table_rows"
            operations.append(
                {
                    "op": "repeat_table_rows",
                    "target": target["id"],
                    "template_row": field["template_row"],
                    "rows": compiled_rows,
                }
            )

        changes.append(
            {
                "source_element": target["id"],
                "semantic_field": field_name,
                "original": original,
                "proposed": proposed,
                "operation_type": operation_type,
                "validation_warnings": [],
            }
        )

    if elements:
        patch["elements"] = elements
    if tokens:
        patch["tokens"] = tokens
    if operations:
        patch["operations"] = operations
    warnings = validate_patch(profile, patch) if validate_patch_output else []
    for warning in warnings:
        warning_element = warning.get("element")
        for change in changes:
            if change["source_element"] == warning_element:
                change["validation_warnings"].append(warning)
    return patch, changes


def compile_template_patch(
    profile: dict[str, Any], contract: dict[str, Any], data: dict[str, Any]
) -> dict[str, Any]:
    """Compile typed semantic values into a validated DocXUnravel patch."""

    patch, _ = _compile_template_population(profile, contract, data)
    return patch


def preview_template_population(
    bundle: str | Path, contract: dict[str, Any], data: dict[str, Any]
) -> dict[str, Any]:
    """Return a read-only field-oriented population plan for one bundle."""

    bundle = Path(bundle)
    profile = json.loads((bundle / "profile.json").read_text(encoding="utf-8"))
    patch, changes = _compile_template_population(
        profile, contract, data, validate_patch_output=False
    )
    try:
        patch_preview = preview_bundle(bundle, patch)
    except (TypeError, ValueError) as error:
        return {
            "valid": False,
            "source_sha256": profile["source"]["sha256"],
            "changes": changes,
            "validation_warnings": [],
            "unsupported_operations": [str(error)],
            "patch": patch,
        }
    return {
        "valid": True,
        "source_sha256": patch_preview["source_sha256"],
        "changes": changes,
        "validation_warnings": patch_preview["warnings"],
        "unsupported_operations": [],
        "patch": patch,
    }


def compile_migration_template_patch(
    profile: dict[str, Any], contract: dict[str, Any], discovery: dict[str, Any]
) -> dict[str, Any]:
    """Validate canonical discovery data and compile a schema 1.1 patch."""

    validate_migration_discovery(discovery)
    return compile_template_patch(profile, contract, discovery)


def preview_migration_template_population(
    bundle: str | Path, contract: dict[str, Any], discovery: dict[str, Any]
) -> dict[str, Any]:
    """Validate canonical discovery data and preview without writing a DOCX."""

    validate_migration_discovery(discovery)
    return preview_template_population(bundle, contract, discovery)


def populate_migration_template(
    bundle: str | Path,
    output: str | Path,
    contract: dict[str, Any],
    discovery: dict[str, Any],
    *,
    patch_output: str | Path | None = None,
    fidelity_output: str | Path | None = None,
) -> dict[str, Any]:
    """Run the validated contract-to-patch-to-render workflow atomically."""

    bundle = Path(bundle)
    output = Path(output)
    preview = preview_migration_template_population(bundle, contract, discovery)
    if not preview["valid"]:
        unsupported = "; ".join(preview["unsupported_operations"])
        raise ValueError(
            f"Template population contains unsupported operations: {unsupported}"
        )
    patch = preview["patch"]
    render_result = render_bundle(bundle, output, patch)
    if patch_output is not None:
        Path(patch_output).write_text(
            json.dumps(patch, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    fidelity = None
    if fidelity_output is not None:
        with tempfile.TemporaryDirectory(
            prefix="docxunravel-template-source-"
        ) as temporary:
            reconstructed_source = Path(temporary) / "source.docx"
            profile = json.loads((bundle / "profile.json").read_text(encoding="utf-8"))
            render_bundle(
                bundle,
                reconstructed_source,
                {"contract": profile["patch_contract"]},
            )
            fidelity = create_fidelity_report(
                reconstructed_source, output, bundle=bundle, patch=patch
            )
        Path(fidelity_output).write_text(
            json.dumps(fidelity, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    return {
        "preview": preview,
        "render": render_result,
        "patch_output": str(patch_output) if patch_output is not None else None,
        "fidelity_output": str(fidelity_output)
        if fidelity_output is not None
        else None,
        "fidelity": fidelity,
    }
