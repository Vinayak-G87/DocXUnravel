"""Validation for canonical migration-discovery JSON.

This domain schema is deliberately independent from document template contracts.
Validation is deterministic and has no model, network, or OOXML dependency.
"""

from __future__ import annotations

from datetime import date
from typing import Any

DISCOVERY_SCHEMA_VERSION = "0.1"
ROOT_FIELDS = {
    "schema_version",
    "fictional",
    "fictional_notice",
    "customer",
    "project",
    "source_infrastructure_inventory",
    "application_inventory",
    "application_dependencies",
    "current_state_observations",
    "proposed_azure_target_services",
    "migration_waves",
    "risks_and_assumptions",
    "security_and_networking_requirements",
    "outstanding_design_decisions",
}


def _object(value: Any, label: str, fields: set[str]) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise TypeError(f"{label} must be an object")
    missing = fields - set(value)
    unknown = set(value) - fields
    if missing:
        raise ValueError(f"{label} is missing: {', '.join(sorted(missing))}")
    if unknown:
        raise ValueError(f"{label} has unknown fields: {', '.join(sorted(unknown))}")
    return value


def _string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise TypeError(f"{label} must be a non-empty string")
    return value


def _choice(value: Any, label: str, choices: set[str]) -> str:
    value = _string(value, label)
    if value not in choices:
        raise ValueError(f"{label} must be one of: {', '.join(sorted(choices))}")
    return value


def _date(value: Any, label: str) -> str:
    value = _string(value, label)
    try:
        date.fromisoformat(value)
    except ValueError as error:
        raise ValueError(f"{label} must use ISO 8601 YYYY-MM-DD format") from error
    return value


def _array(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        raise TypeError(f"{label} must be an array")
    return value


def _unique_id(item: dict[str, Any], label: str, seen: set[str]) -> None:
    identifier = _string(item["id"], f"{label}.id")
    if identifier in seen:
        raise ValueError(f"Duplicate {label} id: {identifier}")
    seen.add(identifier)


def validate_migration_discovery(data: dict[str, Any]) -> dict[str, Any]:
    """Validate canonical v0.1 discovery data and return the original object."""

    data = _object(data, "Migration discovery", ROOT_FIELDS)
    if data["schema_version"] != DISCOVERY_SCHEMA_VERSION:
        raise ValueError(
            f"Unsupported migration discovery schema: {data['schema_version']!r}"
        )
    if data["fictional"] is not True:
        raise ValueError("Demonstration discovery data must be explicitly fictional")
    _string(data["fictional_notice"], "fictional_notice")

    customer_fields = {"name", "industry", "primary_region", "contact_name"}
    customer = _object(data["customer"], "customer", customer_fields)
    for field in customer_fields:
        _string(customer[field], f"customer.{field}")

    project_fields = {
        "name",
        "sponsor",
        "owner",
        "scope",
        "target_start",
        "target_completion",
    }
    project = _object(data["project"], "project", project_fields)
    for field in project_fields - {"target_start", "target_completion"}:
        _string(project[field], f"project.{field}")
    _date(project["target_start"], "project.target_start")
    _date(project["target_completion"], "project.target_completion")

    infrastructure_fields = {
        "id",
        "name",
        "category",
        "location",
        "quantity",
        "operating_system",
        "specification",
    }
    infrastructure_ids: set[str] = set()
    for index, value in enumerate(
        _array(data["source_infrastructure_inventory"], "source infrastructure")
    ):
        label = f"source_infrastructure_inventory[{index}]"
        item = _object(value, label, infrastructure_fields)
        _unique_id(item, label, infrastructure_ids)
        for field in ("name", "location", "operating_system", "specification"):
            _string(item[field], f"{label}.{field}")
        _choice(
            item["category"],
            f"{label}.category",
            {"compute", "database", "storage", "network", "platform"},
        )
        if (
            not isinstance(item["quantity"], int)
            or isinstance(item["quantity"], bool)
            or item["quantity"] < 0
        ):
            raise TypeError(f"{label}.quantity must be a non-negative integer")

    application_fields = {
        "id",
        "name",
        "owner",
        "business_criticality",
        "current_hosting",
        "data_classification",
        "migration_disposition",
    }
    application_ids: set[str] = set()
    for index, value in enumerate(
        _array(data["application_inventory"], "application inventory")
    ):
        label = f"application_inventory[{index}]"
        item = _object(value, label, application_fields)
        _unique_id(item, label, application_ids)
        for field in ("name", "owner", "current_hosting"):
            _string(item[field], f"{label}.{field}")
        _choice(
            item["business_criticality"],
            f"{label}.business_criticality",
            {"critical", "high", "medium", "low"},
        )
        _choice(
            item["data_classification"],
            f"{label}.data_classification",
            {"public", "internal", "confidential", "restricted"},
        )
        _choice(
            item["migration_disposition"],
            f"{label}.migration_disposition",
            {"rehost", "replatform", "refactor", "repurchase", "retain", "retire"},
        )

    dependency_fields = {
        "source_application_id",
        "target_application_id",
        "dependency_type",
        "notes",
    }
    for index, value in enumerate(
        _array(data["application_dependencies"], "application dependencies")
    ):
        label = f"application_dependencies[{index}]"
        item = _object(value, label, dependency_fields)
        source_id = _string(
            item["source_application_id"], f"{label}.source_application_id"
        )
        target_id = _string(
            item["target_application_id"], f"{label}.target_application_id"
        )
        if source_id not in application_ids or target_id not in application_ids:
            raise ValueError(f"{label} references an unknown application")
        _choice(
            item["dependency_type"],
            f"{label}.dependency_type",
            {"synchronous", "asynchronous", "data", "identity", "file"},
        )
        _string(item["notes"], f"{label}.notes")

    observation_fields = {"category", "observation", "impact"}
    for index, value in enumerate(
        _array(data["current_state_observations"], "current-state observations")
    ):
        label = f"current_state_observations[{index}]"
        item = _object(value, label, observation_fields)
        for field in observation_fields:
            _string(item[field], f"{label}.{field}")

    service_fields = {"workload", "azure_service", "region", "rationale"}
    for index, value in enumerate(
        _array(data["proposed_azure_target_services"], "Azure target services")
    ):
        label = f"proposed_azure_target_services[{index}]"
        item = _object(value, label, service_fields)
        for field in service_fields:
            _string(item[field], f"{label}.{field}")

    wave_fields = {
        "wave",
        "name",
        "application_ids",
        "target_start",
        "target_end",
        "exit_criteria",
    }
    wave_numbers: set[int] = set()
    for index, value in enumerate(_array(data["migration_waves"], "migration waves")):
        label = f"migration_waves[{index}]"
        item = _object(value, label, wave_fields)
        wave = item["wave"]
        if not isinstance(wave, int) or isinstance(wave, bool) or wave < 1:
            raise TypeError(f"{label}.wave must be a positive integer")
        if wave in wave_numbers:
            raise ValueError(f"Duplicate migration wave number: {wave}")
        wave_numbers.add(wave)
        _string(item["name"], f"{label}.name")
        _string(item["exit_criteria"], f"{label}.exit_criteria")
        _date(item["target_start"], f"{label}.target_start")
        _date(item["target_end"], f"{label}.target_end")
        wave_applications = _array(item["application_ids"], f"{label}.application_ids")
        if len(wave_applications) != len(set(wave_applications)):
            raise ValueError(f"{label}.application_ids contains duplicates")
        if any(app_id not in application_ids for app_id in wave_applications):
            raise ValueError(f"{label} references an unknown application")

    risks_assumptions = _object(
        data["risks_and_assumptions"],
        "risks_and_assumptions",
        {"risks", "assumptions"},
    )
    risk_ids: set[str] = set()
    risk_fields = {"id", "description", "likelihood", "impact", "mitigation"}
    for index, value in enumerate(_array(risks_assumptions["risks"], "risks")):
        label = f"risks[{index}]"
        item = _object(value, label, risk_fields)
        _unique_id(item, label, risk_ids)
        _string(item["description"], f"{label}.description")
        _string(item["mitigation"], f"{label}.mitigation")
        _choice(item["likelihood"], f"{label}.likelihood", {"low", "medium", "high"})
        _choice(item["impact"], f"{label}.impact", {"low", "medium", "high"})

    assumption_ids: set[str] = set()
    assumption_fields = {"id", "description", "validation_owner"}
    for index, value in enumerate(
        _array(risks_assumptions["assumptions"], "assumptions")
    ):
        label = f"assumptions[{index}]"
        item = _object(value, label, assumption_fields)
        _unique_id(item, label, assumption_ids)
        _string(item["description"], f"{label}.description")
        _string(item["validation_owner"], f"{label}.validation_owner")

    requirements = _object(
        data["security_and_networking_requirements"],
        "security_and_networking_requirements",
        {"security", "networking"},
    )
    requirement_ids: set[str] = set()
    requirement_fields = {"id", "requirement", "priority"}
    for category in ("security", "networking"):
        for index, value in enumerate(_array(requirements[category], category)):
            label = f"{category}[{index}]"
            item = _object(value, label, requirement_fields)
            _unique_id(item, label, requirement_ids)
            _string(item["requirement"], f"{label}.requirement")
            _choice(item["priority"], f"{label}.priority", {"must", "should", "could"})

    decision_ids: set[str] = set()
    decision_fields = {"id", "decision", "owner", "due_date", "status"}
    for index, value in enumerate(
        _array(data["outstanding_design_decisions"], "outstanding design decisions")
    ):
        label = f"outstanding_design_decisions[{index}]"
        item = _object(value, label, decision_fields)
        _unique_id(item, label, decision_ids)
        _string(item["decision"], f"{label}.decision")
        _string(item["owner"], f"{label}.owner")
        _date(item["due_date"], f"{label}.due_date")
        _choice(item["status"], f"{label}.status", {"open", "in_review", "decided"})

    return data
