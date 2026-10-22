import json
import tempfile
import time
import unittest
from pathlib import Path

from docxunravel import export_bundle, profile_document, render_bundle
from tests.support.fixture_factory import create_performance_fixture


class PerformanceTests(unittest.TestCase):
    def test_document_sizes_and_complexity(self):
        cases = [
            ("small", 40, 1, 8.0),
            ("medium", 250, 6, 20.0),
            ("large", 750, 15, 45.0),
        ]
        observations = []
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for name, paragraph_count, table_count, maximum_seconds in cases:
                with self.subTest(size=name):
                    source = root / f"{name}.docx"
                    create_performance_fixture(
                        source,
                        paragraph_count=paragraph_count,
                        table_count=table_count,
                    )
                    started = time.perf_counter()
                    profile = profile_document(source)
                    profile_seconds = time.perf_counter() - started

                    bundle = root / f"{name}-bundle"
                    started = time.perf_counter()
                    exported = export_bundle(source, bundle)
                    extraction_seconds = time.perf_counter() - started

                    output = root / f"{name}-output.docx"
                    started = time.perf_counter()
                    render_bundle(
                        bundle,
                        output,
                        {"contract": exported["patch_contract"]},
                    )
                    render_seconds = time.perf_counter() - started
                    total_seconds = (
                        profile_seconds + extraction_seconds + render_seconds
                    )
                    observations.append(
                        {
                            "case": name,
                            "paragraphs": profile["summary"]["paragraph_count"],
                            "tables": profile["summary"]["table_count"],
                            "docx_bytes": source.stat().st_size,
                            "profile_seconds": round(profile_seconds, 4),
                            "extract_seconds": round(extraction_seconds, 4),
                            "noop_render_seconds": round(render_seconds, 4),
                            "measured_total_seconds": round(total_seconds, 4),
                        }
                    )
                    self.assertLess(total_seconds, maximum_seconds)
                    self.assertTrue(output.exists())
        print("PERFORMANCE_OBSERVATIONS=" + json.dumps(observations, sort_keys=True))


if __name__ == "__main__":
    unittest.main()
