SHELL := /bin/sh

PYTHON ?= python
CONFIG ?= config/config.yaml
OUTPUT ?= workbooks/current_roster_workbook.xlsx
MASTER_OUTPUT ?= workbooks/current_master_workbook.xlsx
COMPRESSED_OUTPUT ?= workbooks/current_node_workbook.xlsx
DROPBOX ?= dropbox

.PHONY: test run master-dropbox compress-dropbox

test:
	$(PYTHON) -m unittest discover -s tests -v

run:
	$(PYTHON) src/excel-template.py --config $(CONFIG) --output $(OUTPUT)

master-dropbox:
	$(PYTHON) src/master_generator.py --config $(CONFIG) --dropbox $(DROPBOX) --output $(MASTER_OUTPUT)

compress-dropbox:
	$(PYTHON) src/template_compressor.py --config $(CONFIG) --dropbox $(DROPBOX) --output $(COMPRESSED_OUTPUT)
