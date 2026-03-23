SHELL := /bin/sh

PYTHON ?= python
APP_NAME ?= excel-template-generator
GUI_ENTRY := src/gui_app.py
CLI_ENTRY := src/excel-template.py
CONFIG_PATH := config/config.yaml
PHASE1_SMOKE := workbooks/phase1_make_smoke.xlsx
PHASE2_SMOKE := workbooks/phase2_make_smoke.xlsx
GUI_SPEC := $(APP_NAME).spec
PYINSTALLER_CONFIG_DIR := build/pyinstaller-config
PYINSTALLER_WORKPATH := build/pyinstaller-work
PYINSTALLER_DISTPATH := dist

.PHONY: help check-clean phase1-test phase1-run phase1-build phase2-test phase2-run phase2-build test run build clean

help:
	@printf "%s\n" \
	"Available targets:" \
	"  check-clean   Fail if the Git working tree is not clean." \
	"  phase1-test   Run the Phase 1 backend smoke tests." \
	"  phase1-run    Run the CLI workbook generator and write a smoke workbook." \
	"  phase1-build  Byte-compile project sources for an early build checkpoint." \
	"  phase2-test   Run the full Phase 1 + Phase 2 checkpoint suite." \
	"  phase2-run    Run the prototype PyQt GUI on the local machine." \
	"  phase2-build  Package the prototype GUI with PyInstaller." \
	"  test          Alias for phase2-test." \
	"  run           Alias for phase2-run." \
	"  build         Alias for phase2-build." \
	"  clean         Remove generated smoke workbooks, caches, and build artifacts."

check-clean:
	@if [ -n "$$(git status --porcelain)" ]; then \
		echo "Working tree is not clean."; \
		git status --short; \
		exit 1; \
	fi

phase1-test:
	$(PYTHON) -m unittest discover -s tests -p 'test_template_generator.py' -v

phase1-run:
	$(PYTHON) $(CLI_ENTRY) --config $(CONFIG_PATH) --output $(PHASE1_SMOKE)

phase1-build:
	$(PYTHON) -m compileall src tests

phase2-test:
	$(PYTHON) scripts/run_phase2_tests.py

phase2-run:
	$(PYTHON) $(GUI_ENTRY) --config $(CONFIG_PATH)

phase2-build:
	PYINSTALLER_CONFIG_DIR=$(PYINSTALLER_CONFIG_DIR) $(PYTHON) -m PyInstaller --noconfirm --clean --windowed --distpath $(PYINSTALLER_DISTPATH) --workpath $(PYINSTALLER_WORKPATH) --specpath . --name $(APP_NAME) $(GUI_ENTRY)

test: phase2-test

run: phase2-run

build: phase2-build

clean:
	rm -rf build
	rm -rf dist
	rm -f $(GUI_SPEC)
	rm -rf __pycache__
	rm -rf src/__pycache__
	rm -rf tests/__pycache__
	rm -rf .pytest_cache
	rm -f .coverage
	rm -f $(PHASE1_SMOKE)
	rm -f $(PHASE2_SMOKE)
