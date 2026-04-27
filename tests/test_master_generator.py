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

from master_generator import MasterGenerationRequest, generate_master_workbook  # noqa: E402
from template_generator import (  # noqa: E402
    TemplateGenerationRequest,
    generate_template_workbook,
    load_generation_inputs,
)


def set_sheet_value(path: Path, sheet_name: str, cell_ref: str, value: object) -> None:
    workbook = load_workbook(path)
    try:
        ws = workbook[sheet_name]
        ws[cell_ref] = value
        workbook.save(path)
    finally:
        workbook.close()


def rename_sheet(path: Path, old_name: str, new_name: str) -> None:
    workbook = load_workbook(path)
    try:
        workbook[old_name].title = new_name
        workbook.save(path)
    finally:
        workbook.close()


class MasterGeneratorTest(unittest.TestCase):
    def test_generate_master_workbook_from_single_node_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            node_path = temp_root / "node_a.xlsx"
            master_path = temp_root / "master.xlsx"

            generate_template_workbook(
                TemplateGenerationRequest(
                    config_path=(REPO_ROOT / "config" / "config.yaml").resolve(),
                    output_path=node_path,
                    block_number="101",
                    fiscal_year="2026",
                    entry_rows=10,
                )
            )

            set_sheet_value(node_path, "PST", "D1", "Push Ups Custom Label")
            rename_sheet(node_path, "PST", "PST Custom")
            set_sheet_value(node_path, "PST Custom", "D3", 61)
            set_sheet_value(node_path, "PST Custom", "E3", 62)
            set_sheet_value(node_path, "PST Custom", "F3", 13)
            set_sheet_value(node_path, "PST Custom", "G3", 600 / 86400)
            set_sheet_value(node_path, "PST Custom", "H3", 720 / 86400)
            set_sheet_value(node_path, "Pool 1", "H3", "Pass")

            generated_path = generate_master_workbook(
                MasterGenerationRequest(
                    config_path=(REPO_ROOT / "config" / "config.yaml").resolve(),
                    workbook_paths=[node_path],
                    output_path=master_path,
                )
            )

            self.assertEqual(generated_path, master_path.resolve())
            workbook = load_workbook(generated_path, data_only=True)
            try:
                ws = workbook["MASTER"]
                loaded = load_generation_inputs(
                    TemplateGenerationRequest(
                        config_path=(REPO_ROOT / "config" / "config.yaml").resolve()
                    )
                )
                first_roster_row = loaded.roster_rows[0]
                self.assertIsNotNone(ws["A2"].value)
                self.assertEqual(ws["B2"].value, first_roster_row["last"])
                self.assertEqual(ws["C2"].value, first_roster_row["first"])
                self.assertEqual(ws["D2"].value, "101")
                self.assertEqual(ws["E2"].value, "node_a.xlsx")
                self.assertEqual(ws["G2"].value, "Pass")
                self.assertEqual(ws["H2"].value, 61)
                self.assertEqual(ws["I2"].value, 62)
                self.assertEqual(ws["J2"].value, 13)
                self.assertEqual(ws["K2"].value, 600)
                self.assertEqual(ws["L2"].value, 720)
            finally:
                workbook.close()

    def test_aggregate_duplicate_metric_values_across_workbooks(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            node_a_path = temp_root / "node_a.xlsx"
            node_b_path = temp_root / "node_b.xlsx"
            master_path = temp_root / "master.xlsx"

            for path in (node_a_path, node_b_path):
                generate_template_workbook(
                    TemplateGenerationRequest(
                        config_path=(REPO_ROOT / "config" / "config.yaml").resolve(),
                        output_path=path,
                        block_number="101",
                        fiscal_year="2026",
                        entry_rows=10,
                    )
                )

            set_sheet_value(node_a_path, "PST", "D3", 45)
            set_sheet_value(node_b_path, "PST", "E3", 55)
            set_sheet_value(node_a_path, "Log PT 1", "D3", 3)
            set_sheet_value(node_b_path, "Log PT 2", "D3", 5)

            generated_path = generate_master_workbook(
                MasterGenerationRequest(
                    config_path=(REPO_ROOT / "config" / "config.yaml").resolve(),
                    workbook_paths=[node_a_path, node_b_path],
                    output_path=master_path,
                )
            )

            workbook = load_workbook(generated_path, data_only=True)
            try:
                ws = workbook["MASTER"]
                loaded = load_generation_inputs(
                    TemplateGenerationRequest(
                        config_path=(REPO_ROOT / "config" / "config.yaml").resolve()
                    )
                )
                self.assertEqual(ws.max_row, len(loaded.roster_rows) + 1)
                self.assertEqual(ws["H2"].value, 45)
                self.assertEqual(ws["I2"].value, 55)
                self.assertEqual(ws["T2"].value, 4)
                self.assertEqual(ws["E2"].value, "node_a.xlsx; node_b.xlsx")
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
