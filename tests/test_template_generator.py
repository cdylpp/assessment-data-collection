from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from template_generator import (  # noqa: E402
    TemplateGenerationRequest,
    generate_template_workbook,
    load_generation_inputs,
)


class TemplateGeneratorSmokeTest(unittest.TestCase):
    def test_load_generation_inputs(self) -> None:
        request = TemplateGenerationRequest(
            config_path=(REPO_ROOT / "config" / "config.yaml").resolve()
        )

        loaded = load_generation_inputs(request)

        self.assertEqual(loaded.config_doc["registry_name"], "usna_screener_evos")
        self.assertTrue(loaded.metrics_doc["metrics"])
        self.assertTrue(loaded.evolutions_doc["evolutions"])
        self.assertTrue(loaded.roster_rows)

    def test_generate_template_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "phase1_smoke.xlsx"
            request = TemplateGenerationRequest(
                config_path=(REPO_ROOT / "config" / "config.yaml").resolve(),
                output_path=output_path,
                block_number="B01",
                fiscal_year="2026",
                entry_rows=10,
            )

            generated_path = generate_template_workbook(request)

            self.assertEqual(generated_path, output_path.resolve())
            self.assertTrue(generated_path.exists())

            workbook = load_workbook(generated_path, read_only=True)
            try:
                self.assertIn("META", workbook.sheetnames)
                self.assertIn("ROSTER", workbook.sheetnames)
                self.assertIn("LOOKUPS", workbook.sheetnames)
                self.assertIn("PST", workbook.sheetnames)
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
