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
    RosterUidConfig,
    TemplateGenerationRequest,
    build_candidate_uid_from_values,
    generate_template_workbook,
    load_generation_inputs,
    load_roster,
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
        self.assertEqual(
            loaded.roster_rows[0]["uid"],
            build_candidate_uid_from_values(["Andrew", "Lucas", "2004-12-16"]),
        )

    def test_load_roster_uses_existing_uid_column_from_config(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            roster_path = Path(temp_dir) / "roster.csv"
            roster_path.write_text(
                "Candidate ID,First,Last,DOB\nabc-123,Jane,Doe,2001-02-03\n",
                encoding="utf-8",
            )

            rows = load_roster(
                roster_path,
                uid_config=RosterUidConfig(
                    mode="existing",
                    source_column="Candidate ID",
                    key_columns=[],
                ),
            )

            self.assertEqual(rows[0]["uid"], "abc-123")
            self.assertEqual(rows[0]["first"], "Jane")
            self.assertEqual(rows[0]["last"], "Doe")

    def test_load_roster_generates_uid_from_configured_key_columns(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            roster_path = Path(temp_dir) / "roster.csv"
            roster_path.write_text(
                "First,Last,DOB\nJane,Doe,02/03/2001\n",
                encoding="utf-8",
            )

            rows = load_roster(
                roster_path,
                uid_config=RosterUidConfig(
                    mode="generated",
                    source_column=None,
                    key_columns=["last", "first", "dob"],
                ),
            )

            self.assertEqual(
                rows[0]["uid"],
                build_candidate_uid_from_values(["Doe", "Jane", "2001-02-03"]),
            )
            self.assertEqual(rows[0]["dob"], "2001-02-03")

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

            workbook = load_workbook(generated_path)
            try:
                loaded = load_generation_inputs(request)
                self.assertIn("META", workbook.sheetnames)
                self.assertIn("ROSTER", workbook.sheetnames)
                self.assertIn("LOOKUPS", workbook.sheetnames)
                self.assertIn("PST", workbook.sheetnames)
                pst = workbook["PST"]
                self.assertTrue(pst.row_dimensions[2].hidden)
                self.assertEqual(pst["A2"].value, "uid")
                self.assertEqual(pst["D2"].value, "m_push_ups")
                self.assertEqual(pst["A3"].value, workbook["ROSTER"]["A2"].value)

                ibs = workbook["IBS PT Land Portage"]
                ibs_config = next(
                    evolution
                    for evolution in loaded.evolutions_doc["evolutions"]
                    if evolution["evolution_id"] == "evo_ibs_pt_land_portage"
                )
                occurrence_count = ibs_config["metric_occurrences"][
                    "m_ibs_low_carry_physicality"
                ]
                for offset in range(occurrence_count):
                    col_idx = 4 + offset
                    self.assertEqual(
                        ibs.cell(row=1, column=col_idx).value,
                        "IBS Low Carry - Physicality {0}".format(offset + 1),
                    )
                    self.assertEqual(
                        ibs.cell(row=2, column=col_idx).value,
                        "m_ibs_low_carry_physicality",
                    )
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
