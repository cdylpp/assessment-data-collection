SHELL := /bin/sh

PYTHON ?= python
CONFIG ?= config/config.yaml
OUTPUT ?= workbooks/current_roster_workbook.xlsx

.PHONY: test run

test:
	$(PYTHON) -m unittest discover -s tests -v

run:
	$(PYTHON) src/excel-template.py --config $(CONFIG) --output $(OUTPUT)
