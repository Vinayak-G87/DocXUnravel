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
    FORMAT_POLICY_ALIASES,
    ROW_TARGET_PATTERN,
    SCHEMA_VERSION,
    TOKEN_PATTERN,
    _checkbox_index,
    _profile_index,
    _profile_row_target,
    _semantic_profile_hash,
    create_fidelity_report,
    preview_bundle,
    render_bundle,
    validate_patch,
)
from .migration_discovery import validate_migration_discovery

TEMPLATE_CONTRACT_VERSION = "0.1"
FIELD_TYPES = {
    "text",
    "table_cell",
    "table_rows",
    "table_keyed_rows",
    "placeholder",
    "checkbox_group",
}
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
    "row_targets",
    "empty_policy",
    "checkboxes",
    "key",
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


def _validate_row_targets(
    field_name: str, field: dict[str, Any], target: dict[str, Any]
) -> None:
    row_targets = field.get("row_targets")
    if not isinstance(row_targets, dict) or not row_targets:
        raise ValueError(f"Table-row field {field_name!r} requires row_targets")
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
    seen_targets: set[str] = set()
    for name, specification in row_targets.items():
        if not isinstance(name, str) or not SEMANTIC_NAME_PATTERN.fullmatch(name):
            raise ValueError(
                f"Table-row field {field_name!r} has invalid row-target name {name!r}"
            )
        specification = _require_object(
            specification, f"Table-row field {field_name!r} target {name!r}"
        )
        unknown = set(specification) - {"target", "data_path", "format_policy"}
        if unknown:
            raise ValueError(
                f"Table-row field {field_name!r} target {name!r} has unknown properties"
            )
        relative_target = specification.get("target")
        if not isinstance(relative_target, str) or not ROW_TARGET_PATTERN.fullmatch(
            relative_target
        ):
            raise ValueError(
                f"Table-row field {field_name!r} has invalid relative target {relative_target!r}"
            )
        if relative_target in seen_targets:
            raise ValueError(
                f"Table-row field {field_name!r} maps multiple values to {relative_target}"
            )
        seen_targets.add(relative_target)
        paragraph = _profile_row_target(target, template_row, relative_target)
        if paragraph is None or paragraph.get("type") != "paragraph":
            raise ValueError(
                f"Table-row field {field_name!r} target does not exist: {relative_target}"
            )
        _validate_data_path(specification.get("data_path"), f"{field_name}.{name}")
        policy = FORMAT_POLICY_ALIASES.get(
            specification.get("format_policy", "preserve_first_run"),
            specification.get("format_policy", "preserve_first_run"),
        )
        if policy not in FORMAT_POLICIES:
            raise ValueError(
                f"Table-row field {field_name!r} target {name!r} has unsupported format policy"
            )


def _validate_keyed_rows(
    field_name: str, field: dict[str, Any], target: dict[str, Any]
) -> None:
    key = _require_object(field.get("key"), f"Keyed-row field {field_name!r} key")
    unknown = set(key) - {"column", "data_path", "match"}
    if unknown:
        raise ValueError(f"Keyed-row field {field_name!r} key has unknown properties")
    column = key.get("column")
    if (
        not isinstance(column, int)
        or isinstance(column, bool)
        or not 0 <= column < target.get("column_count", 0)
    ):
        raise ValueError(
            f"Keyed-row field {field_name!r} key column is outside the table"
        )
    _validate_data_path(key.get("data_path"), field_name)
    if key.get("match", "exact") not in {"exact", "case_insensitive"}:
        raise ValueError(
            f"Keyed-row field {field_name!r} has unsupported key match policy"
        )
    row_targets = field.get("row_targets")
    if not isinstance(row_targets, dict) or not row_targets:
        raise ValueError(f"Keyed-row field {field_name!r} requires row_targets")
    for name, specification in row_targets.items():
        if not isinstance(name, str) or not SEMANTIC_NAME_PATTERN.fullmatch(name):
            raise ValueError(f"Keyed-row field {field_name!r} has invalid target name")
        specification = _require_object(
            specification, f"Keyed-row field {field_name!r} target {name!r}"
        )
        if set(specification) - {"target", "data_path", "format_policy"}:
            raise ValueError(
                f"Keyed-row field {field_name!r} target {name!r} has unknown properties"
            )
        relative = specification.get("target")
        if not isinstance(relative, str) or not ROW_TARGET_PATTERN.fullmatch(relative):
            raise ValueError(
                f"Keyed-row field {field_name!r} target {name!r} is invalid"
            )
        _validate_data_path(specification.get("data_path"), f"{field_name}.{name}")
        policy = specification.get("format_policy", "preserve_first_run")
        if policy not in FORMAT_POLICIES and policy not in FORMAT_POLICY_ALIASES:
            raise ValueError(
                f"Keyed-row field {field_name!r} target {name!r} has unsupported format policy"
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
    seen_checkboxes: dict[str, str] = {}

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
        required_keys = {"type", "required"}
        if field.get("type") != "checkbox_group":
            required_keys.add("target")
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

        field_type = field.get("type")
        if not isinstance(field_type, str):
            raise TypeError(f"Field {field_name!r} type must be a string")
        if field_type == "checkbox_group":
            checkbox_map = field.get("checkboxes")
            if not isinstance(checkbox_map, dict) or not checkbox_map:
                raise ValueError(
                    f"Checkbox-group field {field_name!r} requires checkbox mappings"
                )
            registry = _checkbox_index(profile)
            for selection, checkbox_id in checkbox_map.items():
                if not isinstance(
                    selection, str
                ) or not SEMANTIC_NAME_PATTERN.fullmatch(selection):
                    raise ValueError(
                        f"Checkbox-group field {field_name!r} has invalid selection {selection!r}"
                    )
                if not isinstance(checkbox_id, str) or checkbox_id not in registry:
                    raise KeyError(
                        f"Checkbox-group field {field_name!r} has unknown checkbox ID: {checkbox_id}"
                    )
                if checkbox_id in seen_checkboxes:
                    raise ValueError(
                        f"Checkbox fields {seen_checkboxes[checkbox_id]!r} and {field_name!r} conflict"
                    )
                seen_checkboxes[checkbox_id] = field_name
            forbidden = {
                "target",
                "token",
                "template_row",
                "columns",
                "row_targets",
                "format_policy",
                "empty_policy",
                "key",
            }
            present = forbidden.intersection(field)
            if present:
                raise ValueError(
                    f"Checkbox-group field {field_name!r} cannot define {', '.join(sorted(present))}"
                )
            continue

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
            "table_keyed_rows": "table",
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
            if "key" in field:
                raise ValueError(
                    f"Table-row field {field_name!r} cannot define a keyed-row selector"
                )
            has_columns = "columns" in field
            has_row_targets = "row_targets" in field
            if has_columns == has_row_targets:
                raise ValueError(
                    f"Table-row field {field_name!r} requires exactly one of columns or row_targets"
                )
            if has_columns:
                _validate_columns(field_name, field, target)
            else:
                _validate_row_targets(field_name, field, target)
            empty_policy = field.get("empty_policy", "remove_template_row")
            if empty_policy not in {
                "remove_template_row",
                "retain_empty_row",
                "error",
            }:
                raise ValueError(
                    f"Table-row field {field_name!r} has unsupported empty_policy"
                )
        elif field_type == "table_keyed_rows":
            if "columns" in field or "template_row" in field or "empty_policy" in field:
                raise ValueError(
                    f"Keyed-row field {field_name!r} cannot define repeat-row settings"
                )
            _validate_keyed_rows(field_name, field, target)
        elif (
            "columns" in field
            or "row_targets" in field
            or "template_row" in field
            or "empty_policy" in field
            or "key" in field
        ):
            raise ValueError(
                f"Field {field_name!r} row settings are only valid for table row fields"
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
            value = (
                []
                if field["type"] in {"table_rows", "table_keyed_rows", "checkbox_group"}
                else ""
            )
        field_type = field["type"]
        target: dict[str, Any] = (
            index[field["target"]] if field_type != "checkbox_group" else {}
        )
        original: Any
        proposed: Any
        operation_type: str

        if field_type == "checkbox_group":
            if not isinstance(value, list) or any(
                not isinstance(selection, str) for selection in value
            ):
                raise TypeError(
                    f"Checkbox-group field {field_name!r} must be an array of strings"
                )
            if len(set(value)) != len(value):
                raise ValueError(
                    f"Checkbox-group field {field_name!r} contains duplicate selections"
                )
            checkbox_map = field["checkboxes"]
            unknown_selections = set(value) - set(checkbox_map)
            if unknown_selections:
                raise ValueError(
                    f"Checkbox-group field {field_name!r} has unknown selections: "
                    + ", ".join(sorted(unknown_selections))
                )
            checkbox_registry = _checkbox_index(profile)
            original = [
                selection
                for selection, checkbox_id in checkbox_map.items()
                if checkbox_registry[checkbox_id]["checked"]
            ]
            proposed = value
            operation_type = "set_checkbox_group"
            for selection, checkbox_id in checkbox_map.items():
                operations.append(
                    {
                        "op": "set_checkbox",
                        "target": checkbox_id,
                        "expected": checkbox_registry[checkbox_id]["checked"],
                        "checked": selection in value,
                    }
                )
        elif field_type == "text":
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
        elif field_type == "table_keyed_rows":
            if not isinstance(value, list):
                raise TypeError(f"Keyed-row field {field_name!r} must be an array")
            key_specification = field["key"]
            key_column = key_specification["column"]
            match_policy = key_specification.get("match", "exact")

            def normalize_key(item: Any, policy: str = match_policy) -> str:
                text = str(item).strip()
                return text.casefold() if policy == "case_insensitive" else text

            row_lookup: dict[str, int] = {}
            for existing_row_index, existing_row in enumerate(target.get("rows", [])):
                cells = existing_row.get("cells", [])
                if key_column >= len(cells):
                    continue
                row_key = normalize_key(cells[key_column].get("text", ""))
                if row_key:
                    if row_key in row_lookup:
                        raise ValueError(
                            f"Keyed-row field {field_name!r} has duplicate template key {row_key!r}"
                        )
                    row_lookup[row_key] = existing_row_index

            original = []
            proposed = value
            operation_type = "populate_keyed_table_rows"
            used_rows: set[int] = set()
            for item_index, item in enumerate(value):
                if not isinstance(item, dict):
                    raise TypeError(
                        f"Keyed-row field {field_name!r} row {item_index} must be an object"
                    )
                key_value = _resolve_data_path(
                    item,
                    key_specification.get("data_path", "/key"),
                    f"{field_name}.key",
                )
                lookup_key = normalize_key(key_value)
                if lookup_key not in row_lookup:
                    raise ValueError(
                        f"Keyed-row field {field_name!r} has no template row for {key_value!r}"
                    )
                row_index = row_lookup[lookup_key]
                if row_index in used_rows:
                    raise ValueError(
                        f"Keyed-row field {field_name!r} selects template row {key_value!r} more than once"
                    )
                used_rows.add(row_index)
                row_original: dict[str, str] = {}
                for target_name, specification in field["row_targets"].items():
                    entry = _profile_row_target(
                        target, row_index, specification["target"]
                    )
                    if entry is None or entry.get("type") != "paragraph":
                        raise ValueError(
                            f"Keyed-row field {field_name!r} target {target_name!r} does not exist in row {row_index}"
                        )
                    item_value = _resolve_data_path(
                        item,
                        specification.get("data_path", "/" + target_name),
                        f"{field_name}.{target_name}",
                    )
                    replacement = _table_cell_text(item_value, field_name, item_index)
                    row_original[target_name] = entry.get("text", "")
                    if entry["id"] in elements:
                        raise ValueError(
                            f"Keyed-row field {field_name!r} targets {entry['id']} more than once"
                        )
                    elements[entry["id"]] = {
                        "expected": entry.get("text", ""),
                        "value": replacement,
                        "format_policy": specification.get(
                            "format_policy", "preserve_first_run"
                        ),
                    }
                original.append({"key": key_value, "values": row_original})
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
            compiled_rows: list[Any] = []
            operation_row_targets: dict[str, Any] | None = None
            if "columns" in field:
                ordered_columns = sorted(
                    field["columns"].items(), key=lambda item: item[1]
                )
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
                    compiled_rows.append(
                        [
                            _table_cell_text(row[name], field_name, row_index)
                            for name, _ in ordered_columns
                        ]
                    )
            else:
                operation_row_targets = {}
                for name, specification in field["row_targets"].items():
                    policy = FORMAT_POLICY_ALIASES.get(
                        specification.get("format_policy", "preserve_first_run"),
                        specification.get("format_policy", "preserve_first_run"),
                    )
                    operation_row_targets[name] = {
                        "target": specification["target"],
                        "format_policy": policy,
                    }
                for row_index, row in enumerate(value):
                    compiled: dict[str, str] = {}
                    for name, specification in field["row_targets"].items():
                        if "data_path" in specification:
                            item = _resolve_data_path(
                                row, specification["data_path"], f"{field_name}.{name}"
                            )
                        elif isinstance(row, dict):
                            item = row.get(name, _MISSING)
                        else:
                            item = _MISSING
                        if item is _MISSING:
                            raise ValueError(
                                f"Table-row field {field_name!r} row {row_index}: missing {name}"
                            )
                        compiled[name] = _table_cell_text(item, field_name, row_index)
                    compiled_rows.append(compiled)
            template = target["rows"][field["template_row"]]
            original = [cell.get("text", "") for cell in template["cells"]]
            proposed = compiled_rows
            operation_type = "repeat_table_rows"
            operation = {
                "op": "repeat_table_rows",
                "target": target["id"],
                "template_row": field["template_row"],
                "rows": compiled_rows,
            }
            if "empty_policy" in field:
                operation["empty_policy"] = field["empty_policy"]
            if operation_row_targets is not None:
                operation["row_targets"] = operation_row_targets
            operations.append(operation)

        changes.append(
            {
                "source_element": (
                    target["id"]
                    if field_type != "checkbox_group"
                    else list(field["checkboxes"].values())
                ),
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
