from __future__ import annotations

import argparse
import copy
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSpinBox,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QVBoxLayout,
    QWidget,
)
import yaml

from template_generator import (
    LoadedTemplateConfig,
    TemplateGenerationRequest,
    build_metric_index,
    default_output_path,
    generate_template_workbook,
    load_generation_inputs,
    validate_generation_inputs,
)


TABLE_ROLE_ENTRY = Qt.UserRole + 1
METRIC_COLUMNS = [
    "metric_id",
    "display_name",
    "type",
    "input_kind",
    "unit",
    "domain_ref",
    "min_pass",
]
EVOLUTION_COLUMNS = [
    "evolution_id",
    "display_name",
    "sheet_name",
    "metric_ids",
]


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Template generator prototype GUI.")
    parser.add_argument(
        "--config",
        default=None,
        help="Optional path to config/config.yaml to load on startup.",
    )
    return parser.parse_args(argv)


def dump_yaml(path: Path, document: Dict[str, Any]) -> None:
    with path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(document, handle, sort_keys=False, allow_unicode=False)


def coerce_scalar(text: str, original: Any) -> Any:
    if text == "":
        return None
    if isinstance(original, int) and not isinstance(original, bool):
        try:
            return int(text)
        except ValueError:
            return text
    if isinstance(original, float):
        try:
            return float(text)
        except ValueError:
            return text
    return text


class ProjectConfigurationState(object):
    def __init__(self) -> None:
        self.loaded = None  # type: Optional[LoadedTemplateConfig]
        self.config_doc = None  # type: Optional[Dict[str, Any]]
        self.metrics_doc = None  # type: Optional[Dict[str, Any]]
        self.evolutions_doc = None  # type: Optional[Dict[str, Any]]

    def is_loaded(self) -> bool:
        return self.loaded is not None

    def load(self, config_path: Path) -> LoadedTemplateConfig:
        loaded = load_generation_inputs(
            TemplateGenerationRequest(config_path=config_path.resolve())
        )
        self.loaded = loaded
        self.config_doc = copy.deepcopy(loaded.config_doc)
        self.metrics_doc = copy.deepcopy(loaded.metrics_doc)
        self.evolutions_doc = copy.deepcopy(loaded.evolutions_doc)
        return loaded

    def build_loaded_config(self) -> LoadedTemplateConfig:
        if self.loaded is None or self.config_doc is None:
            raise ValueError("No configuration is loaded")
        if self.metrics_doc is None or self.evolutions_doc is None:
            raise ValueError("Configuration documents are incomplete")

        loaded = LoadedTemplateConfig(
            config_path=self.loaded.config_path,
            metrics_path=self.loaded.metrics_path,
            evolutions_path=self.loaded.evolutions_path,
            roster_path=self.loaded.roster_path,
            config_doc=copy.deepcopy(self.config_doc),
            metrics_doc=copy.deepcopy(self.metrics_doc),
            evolutions_doc=copy.deepcopy(self.evolutions_doc),
            roster_rows=copy.deepcopy(self.loaded.roster_rows),
            metrics_by_id=build_metric_index(self.metrics_doc),
        )
        validate_generation_inputs(loaded)
        return loaded

    def save(self) -> LoadedTemplateConfig:
        loaded = self.build_loaded_config()
        dump_yaml(loaded.metrics_path, loaded.metrics_doc)
        dump_yaml(loaded.evolutions_path, loaded.evolutions_doc)
        self.loaded = loaded
        self.metrics_doc = copy.deepcopy(loaded.metrics_doc)
        self.evolutions_doc = copy.deepcopy(loaded.evolutions_doc)
        return loaded


class TemplateGeneratorMainWindow(QMainWindow):
    def __init__(self) -> None:
        super(TemplateGeneratorMainWindow, self).__init__()
        self.state = ProjectConfigurationState()
        self.setWindowTitle("Excel Template Generator Prototype")
        self.resize(1200, 800)

        self.metrics_table = QTableWidget()
        self.evolutions_table = QTableWidget()
        self.output_path_input = QLineEdit()
        self.roster_override_input = QLineEdit()
        self.block_number_input = QLineEdit("TBD")
        self.fiscal_year_input = QLineEdit("TBD")
        self.entry_rows_input = QSpinBox()
        self.entry_rows_input.setMinimum(1)
        self.entry_rows_input.setMaximum(100000)
        self.entry_rows_input.setValue(200)
        self.config_path_label = QLabel("No config loaded")
        self.status_log = QPlainTextEdit()
        self.status_log.setReadOnly(True)

        self._build_ui()

    def _build_ui(self) -> None:
        self.setStatusBar(QStatusBar())
        self._build_toolbar()

        tabs = QTabWidget()
        tabs.addTab(self._build_metrics_tab(), "Metrics")
        tabs.addTab(self._build_evolutions_tab(), "Evolutions")
        tabs.addTab(self._build_generate_tab(), "Generate")
        self.tabs = tabs
        self.setCentralWidget(tabs)

    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        load_button = QPushButton("Load Config")
        load_button.clicked.connect(self.choose_and_load_config)
        toolbar.addWidget(load_button)

        save_button = QPushButton("Save YAML")
        save_button.clicked.connect(self.save_with_dialogs)
        toolbar.addWidget(save_button)

        generate_button = QPushButton("Generate Workbook")
        generate_button.clicked.connect(self.generate_with_dialogs)
        toolbar.addWidget(generate_button)

        toolbar.addSeparator()
        toolbar.addWidget(QLabel("Active Config:"))
        toolbar.addWidget(self.config_path_label)

    def _build_metrics_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)

        self.metrics_table.setColumnCount(len(METRIC_COLUMNS))
        self.metrics_table.setHorizontalHeaderLabels(METRIC_COLUMNS)
        self.metrics_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.metrics_table)

        button_row = QHBoxLayout()
        add_button = QPushButton("Add Metric")
        add_button.clicked.connect(self.add_metric_row)
        button_row.addWidget(add_button)

        delete_button = QPushButton("Delete Metric")
        delete_button.clicked.connect(self.delete_selected_metric_rows)
        button_row.addWidget(delete_button)
        button_row.addStretch(1)
        layout.addLayout(button_row)
        return container

    def _build_evolutions_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)

        self.evolutions_table.setColumnCount(len(EVOLUTION_COLUMNS))
        self.evolutions_table.setHorizontalHeaderLabels(EVOLUTION_COLUMNS)
        self.evolutions_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.evolutions_table)

        button_row = QHBoxLayout()
        add_button = QPushButton("Add Evolution")
        add_button.clicked.connect(self.add_evolution_row)
        button_row.addWidget(add_button)

        delete_button = QPushButton("Delete Evolution")
        delete_button.clicked.connect(self.delete_selected_evolution_rows)
        button_row.addWidget(delete_button)
        button_row.addStretch(1)
        layout.addLayout(button_row)
        return container

    def _build_generate_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)

        config_box = QGroupBox("Generation Inputs")
        form = QFormLayout(config_box)
        form.addRow("Output Workbook", self._build_path_row(self.output_path_input, self.choose_output_path))
        form.addRow(
            "Roster Override",
            self._build_path_row(self.roster_override_input, self.choose_roster_override),
        )
        form.addRow("Block Number", self.block_number_input)
        form.addRow("Fiscal Year", self.fiscal_year_input)
        form.addRow("Entry Rows", self.entry_rows_input)
        layout.addWidget(config_box)

        status_box = QGroupBox("Status")
        status_layout = QVBoxLayout(status_box)
        status_layout.addWidget(self.status_log)
        layout.addWidget(status_box)
        return container

    def _build_path_row(self, line_edit: QLineEdit, browse_callback: Any) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(line_edit)
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(browse_callback)
        layout.addWidget(browse_button)
        return container

    def append_status(self, message: str) -> None:
        self.statusBar().showMessage(message, 5000)
        self.status_log.appendPlainText(message)

    def choose_and_load_config(self) -> None:
        config_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select config.yaml",
            str(Path.cwd()),
            "YAML Files (*.yaml *.yml)",
        )
        if config_path:
            self.load_config_from_path(Path(config_path))

    def choose_output_path(self) -> None:
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Choose workbook output path",
            self.output_path_input.text() or str(Path.cwd() / "workbook.xlsx"),
            "Excel Workbook (*.xlsx)",
        )
        if output_path:
            self.output_path_input.setText(output_path)

    def choose_roster_override(self) -> None:
        roster_path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose roster override",
            str(Path.cwd()),
            "CSV Files (*.csv)",
        )
        if roster_path:
            self.roster_override_input.setText(roster_path)

    def load_config_from_path(self, config_path: Path) -> None:
        loaded = self.state.load(config_path)
        self.config_path_label.setText(str(loaded.config_path))
        self.populate_metrics_table()
        self.populate_evolutions_table()
        self.output_path_input.setText(str(default_output_path(loaded)))
        self.entry_rows_input.setValue(200)
        self.append_status("Loaded configuration from {0}".format(loaded.config_path))

    def populate_metrics_table(self) -> None:
        metrics = []
        if self.state.metrics_doc is not None:
            metrics = self.state.metrics_doc.get("metrics", [])

        self.metrics_table.setRowCount(0)
        for metric in metrics:
            row = self.metrics_table.rowCount()
            self.metrics_table.insertRow(row)
            for col, field_name in enumerate(METRIC_COLUMNS):
                value = metric.get(field_name, "")
                text = "" if value is None else str(value)
                item = QTableWidgetItem(text)
                if col == 0:
                    item.setData(TABLE_ROLE_ENTRY, copy.deepcopy(metric))
                self.metrics_table.setItem(row, col, item)

    def populate_evolutions_table(self) -> None:
        evolutions = []
        if self.state.evolutions_doc is not None:
            evolutions = self.state.evolutions_doc.get("evolutions", [])

        self.evolutions_table.setRowCount(0)
        for evolution in evolutions:
            row = self.evolutions_table.rowCount()
            self.evolutions_table.insertRow(row)
            for col, field_name in enumerate(EVOLUTION_COLUMNS):
                value = evolution.get(field_name, "")
                if field_name == "metric_ids" and isinstance(value, list):
                    text = ", ".join(str(metric_id) for metric_id in value)
                else:
                    text = "" if value is None else str(value)
                item = QTableWidgetItem(text)
                if col == 0:
                    item.setData(TABLE_ROLE_ENTRY, copy.deepcopy(evolution))
                self.evolutions_table.setItem(row, col, item)

    def add_metric_row(self) -> None:
        self._append_blank_row(self.metrics_table)

    def add_evolution_row(self) -> None:
        self._append_blank_row(self.evolutions_table)

    def delete_selected_metric_rows(self) -> None:
        self._delete_selected_rows(self.metrics_table)

    def delete_selected_evolution_rows(self) -> None:
        self._delete_selected_rows(self.evolutions_table)

    def _append_blank_row(self, table: QTableWidget) -> None:
        row = table.rowCount()
        table.insertRow(row)
        for col in range(table.columnCount()):
            item = QTableWidgetItem("")
            if col == 0:
                item.setData(TABLE_ROLE_ENTRY, {})
            table.setItem(row, col, item)

    def _delete_selected_rows(self, table: QTableWidget) -> None:
        rows = sorted(
            set(index.row() for index in table.selectionModel().selectedRows()),
            reverse=True,
        )
        for row in rows:
            table.removeRow(row)

    def _item_text(self, table: QTableWidget, row: int, col: int) -> str:
        item = table.item(row, col)
        if item is None:
            return ""
        return item.text().strip()

    def sync_tables_to_state(self) -> None:
        if not self.state.is_loaded():
            raise ValueError("Load a configuration before editing")

        metrics = []
        for row in range(self.metrics_table.rowCount()):
            first_item = self.metrics_table.item(row, 0)
            base_entry = {}
            if first_item is not None:
                data = first_item.data(TABLE_ROLE_ENTRY)
                if isinstance(data, dict):
                    base_entry = copy.deepcopy(data)

            for col, field_name in enumerate(METRIC_COLUMNS):
                text = self._item_text(self.metrics_table, row, col)
                original_value = base_entry.get(field_name)
                if text == "":
                    base_entry.pop(field_name, None)
                else:
                    base_entry[field_name] = coerce_scalar(text, original_value)

            if base_entry.get("metric_id"):
                metrics.append(base_entry)

        evolutions = []
        for row in range(self.evolutions_table.rowCount()):
            first_item = self.evolutions_table.item(row, 0)
            base_entry = {}
            if first_item is not None:
                data = first_item.data(TABLE_ROLE_ENTRY)
                if isinstance(data, dict):
                    base_entry = copy.deepcopy(data)

            for col, field_name in enumerate(EVOLUTION_COLUMNS):
                text = self._item_text(self.evolutions_table, row, col)
                if field_name == "metric_ids":
                    metric_ids = [value.strip() for value in text.split(",") if value.strip()]
                    if metric_ids:
                        base_entry[field_name] = metric_ids
                    else:
                        base_entry.pop(field_name, None)
                elif text == "":
                    base_entry.pop(field_name, None)
                else:
                    base_entry[field_name] = text

            if base_entry.get("evolution_id"):
                evolutions.append(base_entry)

        if self.state.metrics_doc is None or self.state.evolutions_doc is None:
            raise ValueError("Configuration documents are not available")

        self.state.metrics_doc["metrics"] = metrics
        self.state.evolutions_doc["evolutions"] = evolutions

    def save_current_config(self) -> LoadedTemplateConfig:
        self.sync_tables_to_state()
        loaded = self.state.save()
        self.append_status("Saved metrics and evolutions YAML")
        return loaded

    def save_with_dialogs(self) -> None:
        try:
            self.save_current_config()
        except Exception as exc:
            QMessageBox.critical(self, "Save Failed", str(exc))
            self.append_status("Save failed: {0}".format(exc))
            return
        QMessageBox.information(self, "Saved", "Configuration saved successfully.")

    def generate_current_workbook(self) -> Path:
        loaded = self.save_current_config()
        roster_override = self.roster_override_input.text().strip()
        output_path = self.output_path_input.text().strip()
        request = TemplateGenerationRequest(
            config_path=loaded.config_path,
            roster_path=Path(roster_override).resolve() if roster_override else None,
            output_path=Path(output_path).resolve() if output_path else None,
            block_number=self.block_number_input.text().strip() or "TBD",
            fiscal_year=self.fiscal_year_input.text().strip() or "TBD",
            entry_rows=int(self.entry_rows_input.value()),
        )
        generated_path = generate_template_workbook(request)
        self.output_path_input.setText(str(generated_path))
        self.append_status("Generated workbook at {0}".format(generated_path))
        return generated_path

    def generate_with_dialogs(self) -> None:
        try:
            generated_path = self.generate_current_workbook()
        except Exception as exc:
            QMessageBox.critical(self, "Generation Failed", str(exc))
            self.append_status("Generation failed: {0}".format(exc))
            return
        QMessageBox.information(
            self,
            "Workbook Generated",
            "Workbook created at {0}".format(generated_path),
        )


def create_application() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    app = create_application()
    window = TemplateGeneratorMainWindow()
    if args.config:
        try:
            window.load_config_from_path(Path(args.config))
        except Exception as exc:
            QMessageBox.critical(window, "Load Failed", str(exc))
            window.append_status("Load failed: {0}".format(exc))
    window.show()
    return app.exec_()


if __name__ == "__main__":
    raise SystemExit(main())
