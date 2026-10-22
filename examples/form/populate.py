"""Extract the form, write its reviewed mapping and patch, and render a filled copy."""

import json
from pathlib import Path
from tempfile import TemporaryDirectory

from generate import normalize_docx

from docxunravel import create_fidelity_report, export_bundle, render_bundle

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "form-template.docx"


def save_json(name, data):
    (HERE / name).write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def main():
    with TemporaryDirectory() as temporary:
        bundle = Path(temporary) / "bundle"
        profile = export_bundle(TEMPLATE, bundle, auto_placeholders=False)
        tables = [item for item in profile["body"] if item["type"] == "table"]
        assert len(tables) == 2

        def cell(table, row, column, expected):
            item = tables[table]["rows"][row]["cells"][column]
            assert item["text"] == expected, (table, row, column, item["text"])
            return item["elements"][0]["id"]

        fields = {
            "name": cell(0, 0, 1, "[...]"),
            "summary": cell(0, 1, 1, "[...]"),
        }
        assert cell(0, 0, 0, "Name") and cell(0, 1, 0, "Summary")
        assert all(
            cell(1, 0, column, heading)
            for column, heading in enumerate(("Item", "Goal", "Metric"))
        )
        goals = {
            label: {
                "goal": cell(1, row, 1, ""),
                "metric": cell(1, row, 2, ""),
            }
            for row, label in enumerate(("Item A", "Item B"), 1)
            if cell(1, row, 0, label)
        }
        controls = profile["checkboxes"]["items"]
        assert len(controls) == 2 and all(
            not control["checked"] for control in controls
        )
        checkboxes = {
            label: control["id"]
            for label, control in zip(("Standard", "Expedited"), controls)
        }

        field_map = {
            "source_sha256": profile["source"]["sha256"],
            "status": "reviewed_mapping_for_this_template_only",
            "fields": fields,
            "goals": goals,
            "checkboxes": checkboxes,
            "read_only": [
                "Form",
                "Guidance",
                "Name and Summary labels",
                "Goals",
                "Item/Goal/Metric headings",
                "Item A and Item B labels",
                "Service option",
            ],
        }
        patch = {
            "contract": profile["patch_contract"],
            "elements": {
                fields["name"]: {"expected": "[...]", "value": "Alex Example"},
                fields["summary"]: {
                    "expected": "[...]",
                    "value": "Prepare a short weekly update.",
                },
            },
            "operations": [
                {
                    "op": "set_checkbox",
                    "target": checkboxes["Standard"],
                    "expected": False,
                    "checked": True,
                }
            ],
        }
        for label, targets in goals.items():
            patch["elements"][targets["goal"]] = {
                "expected": "",
                "value": f"Complete {label.lower()}",
            }
            patch["elements"][targets["metric"]] = {
                "expected": "",
                "value": f"One {label.lower()} delivered",
            }

        output = HERE / "form-filled.docx"
        render_bundle(bundle, output, patch)
        normalize_docx(output)
        report = create_fidelity_report(TEMPLATE, output, bundle=bundle, patch=patch)
        if not report["passed"]:
            raise RuntimeError(f"Fidelity check failed: {report}")
        save_json("field-map.json", field_map)
        save_json("patch.json", patch)
        print(f"Saved {output}; fidelity passed: {report['passed']}")


if __name__ == "__main__":
    main()
