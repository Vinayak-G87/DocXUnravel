import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from docxunravel import export_bundle, inspect_controls, render_bundle

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


class CheckboxControlTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self):
        self.temporary.cleanup()

    def _native_checkbox_document(self, path: Path) -> None:
        plain = self.root / "plain.docx"
        document = Document()
        document.add_paragraph("☐ Azure")
        document.save(str(plain))

        with (
            zipfile.ZipFile(plain) as source,
            zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as destination,
        ):
            for info in source.infolist():
                content = source.read(info.filename)
                if info.filename == "word/document.xml":
                    root = etree.fromstring(content)
                    text = next(root.iter(qn("w:t")))
                    run = text.getparent()
                    paragraph = run.getparent()
                    position = paragraph.index(run)
                    paragraph.remove(run)

                    sdt = etree.Element(qn("w:sdt"))
                    properties = etree.SubElement(sdt, qn("w:sdtPr"))
                    identifier = etree.SubElement(properties, qn("w:id"))
                    identifier.set(qn("w:val"), "42")
                    checkbox = etree.SubElement(properties, f"{{{W14}}}checkbox")
                    checked = etree.SubElement(checkbox, f"{{{W14}}}checked")
                    checked.set(f"{{{W14}}}val", "0")
                    checked_state = etree.SubElement(checkbox, f"{{{W14}}}checkedState")
                    checked_state.set(f"{{{W14}}}val", "2612")
                    checked_state.set(f"{{{W14}}}font", "MS Gothic")
                    unchecked_state = etree.SubElement(
                        checkbox, f"{{{W14}}}uncheckedState"
                    )
                    unchecked_state.set(f"{{{W14}}}val", "2610")
                    unchecked_state.set(f"{{{W14}}}font", "MS Gothic")
                    content_node = etree.SubElement(sdt, qn("w:sdtContent"))
                    content_node.append(run)
                    paragraph.insert(position, sdt)
                    content = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                destination.writestr(info, content)

    def test_native_word_checkbox_is_discovered_once_and_toggled(self):
        source = self.root / "native-checkbox.docx"
        self._native_checkbox_document(source)
        bundle = self.root / "bundle"
        profile = export_bundle(source, bundle)

        self.assertEqual(profile["checkboxes"]["count"], 1)
        inspection = inspect_controls(source)
        self.assertEqual(inspection["summary"]["checkboxes"], 1)
        self.assertNotIn("body", inspection)
        checkbox = profile["checkboxes"]["items"][0]
        self.assertEqual(checkbox["representation"], "w14_checkbox_sdt")
        self.assertFalse(checkbox["checked"])

        output = self.root / "checked.docx"
        render_bundle(
            bundle,
            output,
            {
                "contract": profile["patch_contract"],
                "operations": [
                    {
                        "op": "set_checkbox",
                        "target": checkbox["id"],
                        "expected": False,
                        "checked": True,
                    }
                ],
            },
        )

        with zipfile.ZipFile(output) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))
        checked = next(root.iter(f"{{{W14}}}checked"))
        self.assertEqual(checked.get(f"{{{W14}}}val"), "1")
        display = "".join(node.text or "" for node in root.iter(qn("w:t")))
        self.assertIn("☒ Azure", display)


if __name__ == "__main__":
    unittest.main()
