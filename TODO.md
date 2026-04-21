# TODO

## Master File Aggregator

### Assertions that must be true before starting aggregator

- [ ] META must contain enough fields to validate workbook compatibility.
- [ ] We need at least one red test proving the desired repeated-metric rollup behavior.

## GUI: First iteration requirements for template configuration and generation

Build a desktop GUI for non-technical users to manage workbook configuration and generate template Excel files without editing YAML by hand.

### Platform and constraints

- GUI framework: PyQt 5.9.2 only.
- Target runtime: Windows in an offline DoD environment.
- Development environment: macOS.
- CI/CD enabled.
- Offline only: no network dependency, no telemetry, no remote services, no automatic update checks.

### First iteration goal

The first release should cover local configuration management and template generation only. It should sit on top of the existing YAML-driven workflow and current Excel template generator. As we build out the GUI, we will extend the Excel template generator and underlying YAML model as well.

### In scope for v1

- Load configuration from local project files, including `config/metrics.yaml`, `config/evolutions.yaml`, and related template inputs.
- Create, view, edit, and delete metrics.
- Create, view, edit, and delete evolutions.
- Allow users to assign metrics to evolutions and control display order.
- Support repeated metric usage where a metric may appear multiple times in an evolution. Repeated-use settings must preserve occurrence count and aggregation behavior needed by downstream ingestion.
- Provide forms for core metric fields, including stable identifier, display name, type, unit, domain reference, minimum/maximum or pass threshold fields, input kind, and any Excel input settings already supported by the generator. 
- Provide forms for core evolution fields, including stable identifier, display name, sheet name, and ordered metric membership.
- Validate edits before save or generation. At minimum, validate required fields, unique IDs, known metric references, supported data types, and basic versioning rules.
- Save validated changes back to local YAML files.
- Generate the template Excel workbook by invoking the existing generator flow from the GUI.
- Show clear local error messages when configuration is invalid or generation fails.

### Versioning requirements for v1

- Metrics and evolutions must carry tracked version metadata in configuration.
- Any edit must result in a version bump recorded in metadata.
- Minor version bump: display-only changes such as renaming a metric or evolution.
- Major version bump: structural or semantic changes such as data type changes, domain changes, range changes, aggregation changes, or changing the metric membership of an evolution.
- The GUI may recommend the bump type automatically, but the saved metadata must be explicit and auditable.

### Out of scope for v1

- Converting completed field workbooks into the master workbook.
- Full event builder workflow or higher-level event composition beyond direct evolution editing.
- Multi-user editing, approval workflows, or role-based access.
- Database storage; YAML remains the source of truth.
- Background sync, cloud storage, or any network-connected feature.

### Acceptance criteria for v1

- A user can open the GUI on a Windows machine, load the local config set, edit metrics and evolutions, save valid YAML, and generate a workbook template without using the command line.
- The generated workbook remains compatible with the existing template generation contract.
- Invalid edits are blocked before save or generation and reported with actionable messages.
- The application runs fully offline and does not require network access at runtime.

### Execution plan

Prototype speed is the priority. The plan should favor reuse of the current generator and defer non-critical polish until the basic workflow is working end to end.

#### Phase 1: Stabilize the generator as a reusable backend

- Extract the current workbook generation flow from `src/excel-template.py` into importable functions or a small service module that the GUI can call directly. Review the script for correctness, fix the script if required.
- Keep the existing CLI entry point working so the GUI and command line share the same generation logic.
- Add basic validation helpers for loading `config.yaml`, `metrics.yaml`, `evolutions.yaml`, and roster data.
- Add a small smoke test that loads the sample config and generates a workbook successfully.

##### Definition of Done:
 
 - [x] `src/excel-template.py` has been made into a callable module for the GUI can call when required to generate an excel workbook.
 - [x] Tests are put inplace to ensure any future changes at least passes the "phase 1" checkpoint, that is, can generate a workbook successfully.

#### Phase 2: Build the fastest useful GUI prototype

- Create a PyQt 5.9.2 desktop application with a single main window.
- Use a simple layout with three working areas: metrics editor, evolutions editor, and generate template panel.
- Implement load config, edit in memory, save to YAML, and generate workbook.
- Prefer standard Qt widgets and table/form views over custom components.
- Optimize for function over polish: basic usability is enough for the first prototype.

##### Phase 2 checkpoint tests

- [x] Offscreen GUI smoke test can create the main window successfully on the dev machine.
- [x] GUI can load the sample `config/config.yaml` and populate both metrics and evolutions working areas.
- [x] GUI can edit config values in memory and save the updated YAML successfully.
- [x] GUI can generate a workbook through the shared backend using the current configuration.

##### Definition of Done

- [x] GUI application can run on dev device (macOS).

#### Phase 3: Implement the metrics editor

- Show a searchable list of metrics on the left and a detail form on the right.
- Support create, edit, duplicate, and delete actions.
- Expose only fields already supported by the current YAML contract and generator.
- Validate required fields and block duplicate `metric_id` values.
- For prototype speed, advanced features such as rich diff views or bulk edit should be deferred.

#### Phase 4: Implement the evolutions editor

- Show a list of evolutions and an ordered metric membership editor.
- Support create, edit, duplicate, and delete actions.
- Allow users to add existing metrics to an evolution, remove them, and reorder them.
- Support repeated metric usage in a way that preserves the YAML contract used by generation and downstream ingestion.
- Validate unknown metric references and empty evolution definitions before save.

#### Phase 5: Add versioning workflow

- Add tracked version metadata to metric and evolution records if not already present in YAML.
- On save, detect whether a change is minor or major using the rules defined in the requirements.
- For the first prototype, it is acceptable to prompt the user with the recommended bump and let the user confirm before save.
- Record version updates in configuration metadata so changes are auditable.

#### Phase 6: Harden generation and error handling

- Connect the Generate action to the shared backend service.
- Allow the user to choose output path, block number, fiscal year, and optional roster file.
- Surface validation failures and generation failures in clear local dialogs.
- Confirm that a generated workbook opens correctly in Excel and matches the existing contract.

#### Phase 7: Package for offline Windows use

- Add a reproducible build path for the GUI application in CI/CD.
- Produce a Windows-friendly distributable that works in an offline environment.
- Confirm the packaged app does not depend on internet access at runtime.
- Document installation and local execution steps for operators.

### Prototype-first delivery order

1. Refactor generator into reusable backend.
2. Ship a thin GUI that can load config, edit a small subset of fields, save, and generate.
3. Expand editing coverage for all required metric and evolution fields.
4. Add versioning automation and validation hardening.
5. Package for Windows and connect to CI/CD.

### Explicit cuts to protect prototype schedule

- No event builder in the prototype.
- No master workbook conversion in the prototype.
- No database or non-YAML persistence layer.
- No collaboration or approval workflow.
- No advanced styling work until the end-to-end local workflow is stable.
