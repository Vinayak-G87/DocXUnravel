"""Privacy-conscious structural inspection for tables and checkboxes."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import Any

from .core import profile_document


def _table_reports(
    elements: Iterable[dict[str, Any]], *, depth: int = 0
) -> list[dict[str, Any]]:
    reports: list[dict[str, Any]] = []
    for element in elements:
        if element.get("type") != "table":
            continue
        rows = element.get("rows", [])
        horizontal = [row["index"] for row in rows if row.get("has_horizontal_merge")]
        vertical = [row["index"] for row in rows if row.get("has_vertical_merge")]
        nested: list[dict[str, Any]] = []
        seen_cells: set[tuple[str, str]] = set()
        for row in rows:
            for cell in row.get("cells", []):
                locator = cell.get("locator", {})
                identity = (str(locator.get("part")), str(locator.get("xpath")))
                if identity in seen_cells:
                    continue
                seen_cells.add(identity)
                nested.extend(_table_reports(cell.get("elements", []), depth=depth + 1))
        reports.append(
            {
                "id": element["id"],
                "depth": depth,
                "rows": element.get("row_count", len(rows)),
                "columns": element.get("column_count", 0),
                "style": element.get("style"),
                "horizontal_merge_rows": horizontal,
                "vertical_merge_rows": vertical,
                "nested_table_count": len(nested),
            }
        )
        reports.extend(nested)
    return reports


def inspect_controls(source: str | Path) -> dict[str, Any]:
    """Report table geometry and checkbox controls without paragraph prose."""
    profile = profile_document(source)
    tables = _table_reports(profile.get("body", []))
    for container in profile.get("headers", []) + profile.get("footers", []):
        tables.extend(_table_reports(container.get("elements", [])))
    checkboxes = [
        {
            "id": item["id"],
            "representation": item["representation"],
            "checked": item["checked"],
            "glyph": item.get("glyph"),
            "owner_paragraph": item.get("owner_paragraph"),
            "locator": item.get("locator"),
        }
        for item in profile.get("checkboxes", {}).get("items", [])
    ]
    return {
        "schema_version": "1.0",
        "artifact": "docxunravel.control-inspection",
        "source": profile.get("source", {}),
        "summary": {
            "tables": len(tables),
            "nested_tables": sum(table["depth"] > 0 for table in tables),
            "tables_with_horizontal_merges": sum(
                bool(table["horizontal_merge_rows"]) for table in tables
            ),
            "tables_with_vertical_merges": sum(
                bool(table["vertical_merge_rows"]) for table in tables
            ),
            "checkboxes": len(checkboxes),
            "checked_checkboxes": sum(item["checked"] for item in checkboxes),
        },
        "tables": tables,
        "checkboxes": checkboxes,
    }


__all__ = ["inspect_controls"]
