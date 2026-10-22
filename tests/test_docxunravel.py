import hashlib
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Pt

from docxunravel import (
    create_fidelity_report,
    export_bundle,
    preview_bundle,
    profile_document,
    render_bundle,
    render_document,
    validate_docx_package,
)
from docxunravel.visual_regression import verify_visuals, visual_dependencies


class DocXUnravelTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.source = self.root / "template.docx"

        document = Document()
        title = document.add_paragraph(style="Title")
        title.add_run("Architecture for ").bold = True
        title.add_run("{{ CUST")
        title.add_run("OMER }}").italic = True

        summary = document.add_paragraph("Original summary")
        summary.runs[0].font.size = Pt(11)

        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Component"
        table.cell(0, 1).text = "Technology"
        table.cell(1, 0).text = "Compute"
        table.cell(1, 1).text = "{{TECH}}"

        header = document.sections[0].header
        header.paragraphs[0].text = "Confidential - {{CUSTOMER}}"
        document.save(self.source)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_profile_and_render(self):
        profile = profile_document(self.source)
        self.assertEqual(profile["schema_version"], "1.1")
        self.assertIn("CUSTOMER", profile["summary"]["tokens"])
        self.assertIn("TECH", profile["summary"]["tokens"])
        self.assertEqual(profile["body"][0]["style"], "Title")
        self.assertEqual(profile["body"][2]["type"], "table")
        self.assertIn("body.p1", profile["replacement_contract"]["elements"])

        output = self.root / "generated.docx"
        result = render_document(
            self.source,
            output,
            {
                "tokens": {"CUSTOMER": "Acme", "TECH": "Kubernetes"},
                "elements": {"body.p1": "Generated system summary"},
            },
            verify_sha256=profile["source"]["sha256"],
        )

        rendered = Document(output)
        self.assertEqual(rendered.paragraphs[0].text, "Architecture for Acme")
        self.assertEqual(rendered.paragraphs[1].text, "Generated system summary")
        self.assertEqual(rendered.tables[0].cell(1, 1).text, "Kubernetes")
        self.assertEqual(
            rendered.sections[0].header.paragraphs[0].text, "Confidential - Acme"
        )
        self.assertEqual(result["changed_elements"], 1)
        self.assertEqual(result["token_replacements"], 3)

    def test_lossless_bundle_and_structural_rendering(self):
        bundle = self.root / "bundle"
        profile = export_bundle(self.source, bundle)
        self.assertTrue((bundle / "package" / "word" / "document.xml").exists())
        self.assertGreater(profile["package"]["entry_count"], 5)
        self.assertIn("locator", profile["body"][0])

        output = self.root / "bundle-generated.docx"
        result = render_bundle(
            bundle,
            output,
            {
                "contract": profile["patch_contract"],
                "tokens": {"CUSTOMER": "Acme", "TECH": "EKS"},
                "elements": {"body.p1": "Bundle-generated summary"},
                "operations": [
                    {
                        "op": "insert_after",
                        "target": "body.p1",
                        "element": {
                            "type": "paragraph",
                            "style": "Normal",
                            "text": "Inserted architecture detail",
                        },
                    },
                    {
                        "op": "insert_after",
                        "target": "body.p0",
                        "element": {
                            "type": "table",
                            "style": "Table Grid",
                            "rows": [
                                ["Layer", "Details"],
                                [
                                    "Compute",
                                    {
                                        "elements": [
                                            {"type": "paragraph", "text": "Cluster"},
                                            {
                                                "type": "table",
                                                "style": "Table Grid",
                                                "rows": [["Nodes", "3"]],
                                            },
                                        ]
                                    },
                                ],
                            ],
                        },
                    },
                ],
            },
        )

        rendered = Document(output)
        self.assertEqual(rendered.paragraphs[0].text, "Architecture for Acme")
        self.assertEqual(rendered.paragraphs[1].text, "Bundle-generated summary")
        self.assertEqual(rendered.paragraphs[2].text, "Inserted architecture detail")
        self.assertEqual(rendered.tables[0].cell(1, 1).tables[0].cell(0, 1).text, "3")
        self.assertEqual(rendered.tables[1].cell(1, 1).text, "EKS")
        self.assertEqual(result["structural_changes"], 2)
        self.assertIn("word/document.xml", result["modified_parts"])
        self.assertTrue(result["integrity"]["valid"])
        self.assertTrue(validate_docx_package(output)["valid"])

    def test_patch_preview_conflicts_and_atomic_output(self):
        bundle = self.root / "preview-bundle"
        profile = export_bundle(self.source, bundle)
        patch = {
            "contract": profile["patch_contract"],
            "elements": {
                "body.p1": {
                    "expected": "Original summary",
                    "value": "New summary",
                    "format_policy": "preserve_first_run",
                }
            },
        }
        preview = preview_bundle(bundle, patch)
        self.assertTrue(preview["valid"])
        self.assertEqual(preview["changes"][0]["before"], "Original summary")

        with self.assertRaises(ValueError):
            render_bundle(
                bundle,
                self.root / "conflict.docx",
                {
                    **patch,
                    "contract": {**profile["patch_contract"], "source_sha256": "bad"},
                },
            )
        self.assertFalse((self.root / "conflict.docx").exists())

        with self.assertRaises(ValueError):
            preview_bundle(
                bundle,
                {
                    "contract": profile["patch_contract"],
                    "elements": {
                        "body.p1": {"expected": "stale", "value": "New summary"}
                    },
                },
            )

    def test_no_op_patch_preserves_every_package_part(self):
        bundle = self.root / "noop-bundle"
        profile = export_bundle(self.source, bundle)
        output = self.root / "noop.docx"
        result = render_bundle(
            bundle,
            output,
            {"contract": profile["patch_contract"]},
        )
        self.assertEqual(result["modified_parts"], [])
        with (
            zipfile.ZipFile(self.source) as original,
            zipfile.ZipFile(output) as generated,
        ):
            self.assertEqual(set(original.namelist()), set(generated.namelist()))
            for name in original.namelist():
                self.assertEqual(original.read(name), generated.read(name), name)

    def test_mixed_run_formatting_policies(self):
        bundle = self.root / "format-bundle"
        profile = export_bundle(self.source, bundle)
        contract = profile["patch_contract"]

        first_run_output = self.root / "first-run.docx"
        result = render_bundle(
            bundle,
            first_run_output,
            {
                "contract": contract,
                "elements": {
                    "body.p0": {
                        "value": "Replacement title",
                        "format_policy": "preserve_first_run",
                    }
                },
            },
        )
        first_run = Document(first_run_output).paragraphs[0]
        self.assertEqual(first_run.text, "Replacement title")
        self.assertTrue(first_run.runs[0].bold)
        self.assertTrue(result["warnings"])

        common_output = self.root / "common.docx"
        render_bundle(
            bundle,
            common_output,
            {
                "contract": contract,
                "elements": {
                    "body.p0": {
                        "value": "Common-format title",
                        "format_policy": "preserve_common_formatting",
                    }
                },
            },
        )
        common = Document(common_output).paragraphs[0]
        self.assertEqual(common.text, "Common-format title")
        self.assertIsNone(common.runs[0].bold)

        text_only_output = self.root / "text-only.docx"
        render_bundle(
            bundle,
            text_only_output,
            {
                "contract": contract,
                "elements": {
                    "body.p0": {
                        "value": "Text distributed over existing formatted runs",
                        "format_policy": "replace_text_only",
                    }
                },
            },
        )
        text_only = Document(text_only_output).paragraphs[0]
        self.assertEqual(
            text_only.text, "Text distributed over existing formatted runs"
        )
        self.assertTrue(text_only.runs[0].bold)
        self.assertTrue(text_only.runs[2].italic)

        with self.assertRaises(ValueError):
            render_bundle(
                bundle,
                self.root / "rejected.docx",
                {
                    "contract": contract,
                    "elements": {
                        "body.p0": {
                            "value": "Unsafe replacement",
                            "format_policy": "reject_mixed_runs",
                        }
                    },
                },
            )

    def test_run_level_replacement_preserves_other_runs(self):
        bundle = self.root / "run-bundle"
        profile = export_bundle(self.source, bundle)
        output = self.root / "run-output.docx"
        result = render_bundle(
            bundle,
            output,
            {
                "contract": profile["patch_contract"],
                "runs": {
                    "body.p0": [
                        {
                            "index": 0,
                            "expected": "Architecture for ",
                            "value": "Platform for ",
                        }
                    ]
                },
            },
        )
        paragraph = Document(output).paragraphs[0]
        self.assertEqual(paragraph.text, "Platform for {{ CUSTOMER }}")
        self.assertTrue(paragraph.runs[0].bold)
        self.assertEqual(paragraph.runs[1].text, "{{ CUST")
        self.assertTrue(paragraph.runs[2].italic)
        self.assertEqual(result["changed_runs"], 1)

    def test_stale_profile_and_structural_conflicts_are_rejected(self):
        source_bundle = self.root / "source-conflict-bundle"
        source_profile = export_bundle(self.source, source_bundle)
        document_xml = source_bundle / "package" / "word" / "document.xml"
        document_xml.write_bytes(document_xml.read_bytes() + b" ")
        with self.assertRaisesRegex(ValueError, "package part has changed"):
            preview_bundle(
                source_bundle,
                {"contract": source_profile["patch_contract"]},
            )
        source_profile_path = source_bundle / "profile.json"
        modified_profile = json.loads(source_profile_path.read_text(encoding="utf-8"))
        changed_bytes = document_xml.read_bytes()
        document_entry = next(
            entry
            for entry in modified_profile["package"]["entries"]
            if entry["path"] == "word/document.xml"
        )
        document_entry["sha256"] = hashlib.sha256(changed_bytes).hexdigest()
        document_entry["size"] = len(changed_bytes)
        source_profile_path.write_text(json.dumps(modified_profile), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "Semantic profile is stale"):
            preview_bundle(
                source_bundle,
                {"contract": source_profile["patch_contract"]},
            )

        conflict_bundle = self.root / "conflict-bundle"
        profile = export_bundle(self.source, conflict_bundle)
        contract = profile["patch_contract"]
        with self.assertRaises(ValueError):
            preview_bundle(
                conflict_bundle,
                {
                    "contract": contract,
                    "operations": [
                        {
                            "op": "insert_before",
                            "target": "body.p1",
                            "element": {"type": "paragraph", "text": "Before"},
                        },
                        {
                            "op": "delete",
                            "target": "body.p1",
                        },
                    ],
                },
            )
        with self.assertRaisesRegex(ValueError, "inside deleted element"):
            preview_bundle(
                conflict_bundle,
                {
                    "contract": contract,
                    "elements": {"body.t0.r1.c1.p0": "Conflicting cell edit"},
                    "operations": [{"op": "delete", "target": "body.t0"}],
                },
            )

        profile_path = conflict_bundle / "profile.json"
        stale_profile = json.loads(profile_path.read_text(encoding="utf-8"))
        stale_profile["body"][1]["text"] = "Tampered semantic profile"
        profile_path.write_text(json.dumps(stale_profile), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "Semantic profile is stale"):
            preview_bundle(conflict_bundle, {"contract": contract})

    def test_control_nodes_are_rejected_for_full_and_run_replacement(self):
        source = self.root / "breaks.docx"
        document = Document()
        run = document.add_paragraph().add_run("before")
        run.add_break()
        run.add_text("after")
        document.save(source)
        bundle = self.root / "break-bundle"
        profile = export_bundle(source, bundle)
        self.assertIn(
            "line_or_page_break", profile["body"][0]["unsupported_edit_structures"]
        )
        with self.assertRaisesRegex(ValueError, "Unsupported OOXML structures"):
            preview_bundle(
                bundle,
                {
                    "contract": profile["patch_contract"],
                    "elements": {"body.p0": "replacement"},
                },
            )
        with self.assertRaisesRegex(ValueError, "cannot be safely edited at run level"):
            preview_bundle(
                bundle,
                {
                    "contract": profile["patch_contract"],
                    "runs": {"body.p0": [{"index": 0, "value": "replacement"}]},
                },
            )

    @unittest.skipUnless(
        visual_dependencies()["available"],
        "LibreOffice, PyMuPDF, and Pillow are optional visual-test dependencies",
    )
    def test_optional_visual_regression_for_no_op_output(self):
        bundle = self.root / "visual-bundle"
        profile = export_bundle(self.source, bundle)
        output = self.root / "visual-noop.docx"
        render_bundle(bundle, output, {"contract": profile["patch_contract"]})
        report = verify_visuals(
            self.source,
            output,
            self.root / "visual-report",
            difference_threshold=0.0,
            expected_difference_threshold=0.0,
            no_op=True,
        )
        self.assertTrue(report["passed"])
        self.assertEqual(report["mode"], "no-op")
        self.assertTrue((self.root / "visual-report" / "visual-report.json").exists())
        fidelity = create_fidelity_report(
            self.source,
            output,
            bundle=bundle,
            patch={"contract": profile["patch_contract"]},
            visual_report=report,
        )
        self.assertTrue(fidelity["passed"])
        self.assertTrue(fidelity["visual_differences"]["performed"])

        changed_output = self.root / "visual-changed.docx"
        render_bundle(
            bundle,
            changed_output,
            {
                "contract": profile["patch_contract"],
                "elements": {"body.p1": "A substantially different generated summary."},
            },
        )
        bounded = verify_visuals(
            self.source,
            changed_output,
            self.root / "visual-bounded-report",
            difference_threshold=0.0,
            expected_difference_threshold=0.0,
            expected_changed_pages={1},
        )
        self.assertFalse(bounded["passed"])
        self.assertIn(1, bounded["unexpected_changed_pages"])

    def test_repository_regression_fixture_with_example_patch(self):
        project_root = Path(__file__).resolve().parents[1]
        fixture = (
            project_root
            / "tests"
            / "fixtures"
            / "DocXUnravel_Regression_Test_Fixture.docx"
        )
        patch_path = project_root / "examples" / "patches" / "schema-1.1-example.json"
        bundle = self.root / "regression-bundle"
        output = self.root / "regression-output.docx"

        profile = export_bundle(fixture, bundle)
        with patch_path.open(encoding="utf-8") as handle:
            patch = json.load(handle)
        self.assertEqual(patch["contract"], profile["patch_contract"])
        self.assertTrue(preview_bundle(bundle, patch)["valid"])
        with self.assertRaisesRegex(ValueError, "Unsupported OOXML structures"):
            preview_bundle(
                bundle,
                {
                    "contract": profile["patch_contract"],
                    "elements": {"body.p12": "Unsafe hyperlink replacement"},
                },
            )

        result = render_bundle(bundle, output, patch)
        rendered = Document(output)
        body_text = "\n".join(paragraph.text for paragraph in rendered.paragraphs)
        self.assertIn(
            "Customer: Acme Corporation  |  Platform: Order Platform", body_text
        )
        self.assertIn("Inserted regression test paragraph", body_text)
        self.assertIn("event-driven architecture", body_text)
        self.assertNotIn("TARGET_DELETE:", body_text)
        self.assertEqual(len(rendered.tables), 2)
        self.assertEqual(len(rendered.inline_shapes), 1)
        self.assertTrue(result["integrity"]["valid"])
        self.assertEqual(result["token_replacements"], 4)
        self.assertEqual(result["structural_changes"], 3)
        modified = set(result["modified_parts"])
        with zipfile.ZipFile(output) as generated:
            for entry in profile["package"]["entries"]:
                if entry["path"] not in modified:
                    self.assertEqual(
                        hashlib.sha256(generated.read(entry["path"])).hexdigest(),
                        entry["sha256"],
                        entry["path"],
                    )
            self.assertIn("word/media/image1.png", generated.namelist())
            relationships = generated.read("word/_rels/document.xml.rels")
            self.assertIn(b"relationships/hyperlink", relationships)
            self.assertIn(b"relationships/image", relationships)

    def test_hash_and_unknown_element_are_rejected(self):
        with self.assertRaises(ValueError):
            render_document(
                self.source, self.root / "bad.docx", {}, verify_sha256="bad"
            )

        with self.assertRaises(KeyError):
            render_document(
                self.source,
                self.root / "bad.docx",
                {"elements": {"body.p999": "no"}},
            )


if __name__ == "__main__":
    unittest.main()
