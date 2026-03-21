from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtWidgets import QApplication


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from gui_app import TemplateGeneratorMainWindow, create_application  # noqa: E402
from template_generator import TemplateGenerationRequest, load_generation_inputs  # noqa: E402


class GuiPrototypeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = create_application()

    def setUp(self) -> None:
        self.window = TemplateGeneratorMainWindow()

    def tearDown(self) -> None:
        self.window.close()

    def test_window_boots(self) -> None:
        self.assertEqual(self.window.windowTitle(), "Excel Template Generator Prototype")
        self.assertEqual(self.window.tabs.count(), 3)

    def test_load_config_populates_tables(self) -> None:
        self.window.load_config_from_path((REPO_ROOT / "config" / "config.yaml").resolve())

        self.assertGreater(self.window.metrics_table.rowCount(), 0)
        self.assertGreater(self.window.evolutions_table.rowCount(), 0)
        self.assertIn("config.yaml", self.window.config_path_label.text())

    def test_save_round_trip(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            shutil.copytree(REPO_ROOT / "config", temp_root / "config")

            config_path = (temp_root / "config" / "config.yaml").resolve()
            self.window.load_config_from_path(config_path)
            self.window.metrics_table.item(0, 1).setText("Push Ups Prototype")

            self.window.save_current_config()

            loaded = load_generation_inputs(
                TemplateGenerationRequest(config_path=config_path)
            )
            self.assertEqual(
                loaded.metrics_doc["metrics"][0]["display_name"],
                "Push Ups Prototype",
            )

    def test_generate_workbook_from_gui(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            shutil.copytree(REPO_ROOT / "config", temp_root / "config")

            config_path = (temp_root / "config" / "config.yaml").resolve()
            output_path = temp_root / "phase2_gui_smoke.xlsx"
            self.window.load_config_from_path(config_path)
            self.window.output_path_input.setText(str(output_path))
            self.window.block_number_input.setText("B02")
            self.window.fiscal_year_input.setText("2027")

            generated_path = self.window.generate_current_workbook()

            self.assertEqual(generated_path, output_path.resolve())
            self.assertTrue(generated_path.exists())


if __name__ == "__main__":
    unittest.main()
