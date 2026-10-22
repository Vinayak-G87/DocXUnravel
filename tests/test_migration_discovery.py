import copy
import json
import unittest
from pathlib import Path

from docxunravel.migration_discovery import validate_migration_discovery
from docxunravel.schemas import schema_path


class MigrationDiscoveryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = json.loads(
            (
                Path(__file__).resolve().parent
                / "fixtures"
                / "migration-discovery.json"
            ).read_text(encoding="utf-8")
        )

    def test_fictional_example_validates(self):
        validated = validate_migration_discovery(self.data)
        self.assertIs(validated, self.data)
        self.assertTrue(validated["fictional"])
        self.assertIn("ALL CUSTOMER", validated["fictional_notice"])
        self.assertEqual(len(validated["application_inventory"]), 3)

    def test_schema_is_separate_from_template_contract(self):
        schema = json.loads(
            schema_path("migration-discovery-0.1.schema.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(schema["properties"]["schema_version"]["const"], "0.1")
        self.assertIn("application_inventory", schema["properties"])
        self.assertNotIn("binding", schema["properties"])
        self.assertNotIn("fields", schema["properties"])

    def test_nonfictional_missing_and_cross_reference_errors_are_rejected(self):
        nonfictional = copy.deepcopy(self.data)
        nonfictional["fictional"] = False
        with self.assertRaisesRegex(ValueError, "explicitly fictional"):
            validate_migration_discovery(nonfictional)

        missing = copy.deepcopy(self.data)
        del missing["customer"]
        with self.assertRaisesRegex(ValueError, "missing"):
            validate_migration_discovery(missing)

        invalid_dependency = copy.deepcopy(self.data)
        invalid_dependency["application_dependencies"][0]["target_application_id"] = (
            "FICTIONAL-UNKNOWN"
        )
        with self.assertRaisesRegex(ValueError, "unknown application"):
            validate_migration_discovery(invalid_dependency)

    def test_duplicate_identifiers_and_bad_dates_are_rejected(self):
        duplicate = copy.deepcopy(self.data)
        duplicate["application_inventory"][1]["id"] = duplicate[
            "application_inventory"
        ][0]["id"]
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            validate_migration_discovery(duplicate)

        invalid_date = copy.deepcopy(self.data)
        invalid_date["project"]["target_start"] = "someday"
        with self.assertRaisesRegex(ValueError, "ISO 8601"):
            validate_migration_discovery(invalid_date)


if __name__ == "__main__":
    unittest.main()
