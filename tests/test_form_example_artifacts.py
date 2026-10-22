import json
import tempfile
import unittest
from pathlib import Path

from docx import Document

from docxunravel import (
    create_fidelity_report,
    export_bundle,
    inspect_controls,
    validate_docx_package,
)


class FormExampleArtifactTests(unittest.TestCase):
    def test_form_artifacts(self):
        example = Path(__file__).resolve().parents[1] / "examples" / "form"
        template = example / "form-template.docx"
        filled = example / "form-filled.docx"
        field_map = json.loads((example / "field-map.json").read_text(encoding="utf-8"))
        patch = json.loads((example / "patch.json").read_text(encoding="utf-8"))

        for path in (template, filled):
            self.assertTrue(validate_docx_package(path)["valid"])
        source = Document(str(template))
        result = Document(str(filled))
        self.assertEqual(
            [source.tables[0].cell(row, 1).text for row in (0, 1)], ["[...]"] * 2
        )
        self.assertNotEqual(field_map["fields"]["name"], field_map["fields"]["summary"])
        for row, key in enumerate(("name", "summary")):
            target = field_map["fields"][key]
            self.assertEqual(patch["elements"][target]["expected"], "[...]")
            self.assertEqual(
                result.tables[0].cell(row, 1).text, patch["elements"][target]["value"]
            )
        for row, label in enumerate(("Item A", "Item B"), 1):
            self.assertEqual(result.tables[1].cell(row, 0).text, label)
            for column, key in ((1, "goal"), (2, "metric")):
                target = field_map["goals"][label][key]
                self.assertEqual(
                    result.tables[1].cell(row, column).text,
                    patch["elements"][target]["value"],
                )

        before = {item["id"]: item for item in inspect_controls(template)["checkboxes"]}
        after = {item["id"]: item for item in inspect_controls(filled)["checkboxes"]}
        self.assertEqual(set(before), set(field_map["checkboxes"].values()))
        for label, target in field_map["checkboxes"].items():
            self.assertFalse(before[target]["checked"])
            self.assertEqual(after[target]["checked"], label == "Standard")
        self.assertEqual(
            patch["operations"][0]["target"], field_map["checkboxes"]["Standard"]
        )

        with tempfile.TemporaryDirectory() as temporary:
            bundle = Path(temporary) / "bundle"
            profile = export_bundle(template, bundle, auto_placeholders=False)
            self.assertEqual(field_map["source_sha256"], profile["source"]["sha256"])
            self.assertEqual(patch["contract"], profile["patch_contract"])
            report = create_fidelity_report(
                template, filled, bundle=bundle, patch=patch
            )
            self.assertTrue(report["passed"], report)


if __name__ == "__main__":
    unittest.main()
