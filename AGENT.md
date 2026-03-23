# AGENT Instructions

## Overview

You are software engineering developing a light-weight excel workbook generator based on configuration files. The entire method should take two inputs: (1) a roster of names and date of births, (2) a configuration file that specifies evolutions and metrics collected for each evolution (we often use evo for shorthand).

After completing the workbook generator, we are adding a simple GUI application for the end user to use the application without needing to run any command line scripts. 
We are use a simple MVC architecture with PyQt. 

The program outputs a data collection ready excel notebook pre-populated with a one to one mapping for evolutions and sheets. Each evo (sheet) displays candidate names locked on the left side (y-axis) and metrics on the x-axis with predefined data types (and dropdowns is specified in configs).

The program should also contain an "inverse" operation that takes an excel workbook and converts to a master excel file based on the data contract for the master excel file.

We want to prioritize flexibility and correctness. To be more specific, the users should be able to change the configuration values at will and run the program to re-generate a new excel sheet. Even better, the user can have different config files for different sets on evos. It should be correct in that no data should be lost when aggregating the scores from the data collection workbook to the master excel file.

### Structure

- `config` contains all the configuration files for creating workbooks including the roster for input.
- `src` contains the source code for main and excel operations
- `workbooks` is the drop folder for generating workbooks. The program should drop newly generated excels here.

Below is a phased plan for this project:

## Phase 0 — Define the offline “contract” (Node Workbook API)

Goal: Lock the interface between field capture and office ingestion.

### Deliverables

- A short “Node Workbook Contract” spec:
- required sheets (META, ROSTER, LOOKUPS, EXPORT optional, plus evolution sheets)
- required metadata fields (block_id, block_year, template_version, metrics_version, etc.)
- required candidate identifiers (candidate_id + display name)
- stable evolution identifiers (sheet keys)
- how time is represented (recommend: export as total_seconds)
- comments scoping rule (per evolution / per metric, never ambiguous “global comments”)

### Best practices

- Treat this as an API: breaking changes require version bump.
- Prefer stable IDs (e.g., evo_5mile_run, m_5mile_run_time_seconds) over display labels.

⸻

## Phase 1 — Configuration design (YAML becomes source of truth)

Goal: Centralize all template + validation configuration in YAML.

### Deliverables

- metrics.yaml (node capture contract + workbook layout + validation rules)
- evolutions catalog
- metrics registry: metric_id, type, unit, domain_ref, required/optional, derived rules
- domains: dropdown lists
- workbook layout hints: which metrics appear on which evolution sheet and in what order
- map.yaml (node capture → source system schema mapping)
- metric_id → source field name(s) / transform rules
- (Optional) thresholds.yaml (if pass/fail thresholds change by block/year)
- prevents churn in metrics.yaml

### Best practices

- Keep metrics.yaml stable; isolate frequently changing thresholds.
- Enforce uniqueness of IDs at build time.

⸻

## Phase 2 — Template generator (Option A)

Goal: Build the script that compiles YAML into a fully configured node workbook template.

### Deliverables

- A generator script (e.g., Python) that:
	1.	Reads metrics.yaml (and optionally thresholds.yaml)
	2.	Produces NodeTemplate_v{template_version}.xlsx
	3.	Writes:
- META sheet (locked): template_version, metrics_version, generated_on, config_hash
- ROSTER sheet: candidate_id, candidate_name (pre-populated later per block)
- LOOKUPS sheet: domains expanded from YAML + named ranges
- evolution sheets: candidate rows + metric columns + validation + formulas
- optional EXPORT sheet (a normalized table for ingestion convenience)
	4.	Applies protection rules:
- lock formulas, unlock entry cells
- protect sheet structures
	5.	Adds visual “guard rails”:
- conditional formatting for missing required values
- red highlights for invalid/unexpected entries (should be rare if validation is correct)

### Best practices

- Deterministic output: same YAML → identical structure every time.
- Sheet naming: use evolution_id or safe stable names, not display names.
- Put all dropdown sources in LOOKUPS with named ranges (domain_pass_fail, etc.).
- Embed a config_hash so ingestion can verify the workbook matches the expected config.

⸻

## Phase 3 — Per-block workbook instantiation (Roster injection)

Goal: Take the generated template and create the actual field workbook(s) for a specific block.

### Deliverables

- A script or step that:
- takes the block roster (source list of candidates)
- clones the template into Block_{block_id}_NodeWorkbook.xlsx
- pre-populates candidate rows on every evolution sheet
- fills META with block identifiers + roster version
- Optional: split by node/team (if multiple tablets/teams capture different evolutions)

### Best practices

- Don’t hand-edit rosters in the field workbook.
- Candidate identity should be candidate_id primary, name secondary.

⸻

## Phase 4 — Data ingestion (concat → validate → transform)

Goal: Safely convert field workbooks into the source system’s expected format.

### Deliverables

- Ingestion script that:
- loads all node workbooks for a block
- verifies:
- template_version / metrics_version / config_hash
- expected sheets and headers
- unpivots sheets to long format:
- block_id, candidate_id, evolution_id, metric_id, value, evaluator_id(optional), timestamp(optional)
- applies type normalization:
- timed conversions to seconds
- domain validation
- null/blank handling
- applies map.yaml transforms into source schema output
- produces:
- clean_long.csv (audit-friendly)
- source_payload.csv/xlsx/json (whatever source system expects)

### Best practices

- Fail fast: reject bad templates early.
- Keep an audit trail: raw workbook filename + row/column references.

⸻

## Phase 5 — Operationalization & governance

Goal: Keep this system maintainable and trustworthy.

### Deliverables

- Versioning policy:
- semantic versioning for templates/config
- Change control:
- metric additions/renames require a version bump
- deprecated metric support window
- A small QA harness:
- validate YAML schema
- generate template
- run ingestion on a synthetic test workbook
- Minimal documentation:
- “Field user instructions” (1 pager)
- “Operator runbook” for generation/ingestion

### Best practices
- One canonical repo/folder: /config/metrics.yaml, /config/map.yaml, /generator/, /ingest/.
- Store generated templates in a versioned folder; never overwrite without bumping version.

