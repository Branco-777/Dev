# Master Workbook Orchestrator

**Last updated:** 8 September 2026

## Overview

This repository contains a standalone Python runner that creates
stress-specific workbooks from one read-only master workbook. The runner
validates the `Global Control` sheet, expands selected entity/category and
stress combinations, filters the workbook to the required sheets, copies the
selected sensitivity matrices, and writes one output workbook per combination.

The master workbook is never saved by the runner. The implementation edits the
XLSX package directly so validation does not require loading large data sheets
through `openpyxl`. Output packages are written to temporary files and moved
into place only after they are complete.

## Workbook contract

The master workbook must contain a worksheet named `Global Control` with these
Excel tables:

- `tbl_entity_mapping`: `Entity`, `Category`, `SourceSheet`, `Selected`
- `tbl_stress_selection`: `Stress`, `Transition Tab`, `Spread Tab`, `Suffix`,
  `OutputFolder`, `Selected`
- `tbl_run_settings`: `Setting`, `Value`

At least one entity mapping and one stress must be selected. The accepted
truth values for `Selected` are `yes`, `true`, `1`, and `selected`, ignoring
case and surrounding whitespace.

### Run settings

`tbl_run_settings` must contain exactly one value for each of:

- `ValuationDate`, supplied as an ISO date or Excel date
- `risk_factor_path`, which must be non-blank
- `bool_inv_exp`

The legacy global `OutputFolder` setting is not supported. Output folders are
configured per selected stress row.

### Entity and stress mappings

Every `SourceSheet` must name an existing worksheet. Entity/category pairs and
stress names must be unique. Each stress row must name existing `Transition
Tab` and `Spread Tab` worksheets and provide a non-blank `Suffix`.

A selected stress must also provide a non-blank `OutputFolder`. Relative paths
are resolved from the master workbook's directory; absolute paths are used as
provided. An unselected stress may leave `OutputFolder` blank.

### Required worksheets and named inputs

The master workbook must contain:

- A hidden `I. Sensitivity - Template` worksheet
- An `I. Control` worksheet
- The workbook-scoped named inputs `valuation_date`, `entity_name`,
  `risk_factor_path`, `output_folder_path`, `prm_LPIgeneratorWB`,
  `output_suffix`, and `bool_inv_exp`

The selected transition matrix is copied into every transition table on the
sensitivity template. The selected spread matrix is copied into every spread
and `_Orig` spread table. `prm_LPIgeneratorWB` is written as blank.

The eight static worksheets are derived from all worksheets except:

- `Global Control` and `I. Control`
- Entity source worksheets
- Configured transition and spread worksheets
- `I. Sensitivity - Template`
- Worksheets whose names start with `Sensitivities - ` or `Sensitivities_`

Exactly eight static worksheets must remain after these exclusions.

## Generated workbooks

Each output contains:

- One selected data worksheet, renamed to `I. Bonds Data` by default
- `I. Control`
- One populated sensitivity worksheet, renamed to `I. Sensitivity` by default
- The eight static worksheets

The data and sensitivity names can be changed with `--data-sheet-name` and
`--sensitivity-sheet-name`. Names must be unique, no longer than 31 characters,
and valid in Excel.

Output filenames use this format:

```text
YYYYMM_<entity>_<category>_<suffix>.xlsx
```

For example, a 30 June 2026 valuation produces a filename beginning with
`202606_`. The stress name is omitted because the stress-specific suffix
identifies the selected output. Invalid Windows filename characters are
replaced with underscores.

## Running the orchestrator

### Python

```text
python master_workbook_orchestrator_merged.py --master <master-workbook.xlsx>
```

Optional arguments are:

```text
--overwrite
--data-sheet-name <text>
--sensitivity-sheet-name <text>
```

Existing outputs are rejected by default. `--overwrite` allows an existing
output with the same deterministic name to be replaced.

### Windows launcher

Use `run_master_workbook_orchestrator_package.bat` with an optional master
workbook path:

```text
run_master_workbook_orchestrator_package.bat <master-workbook.xlsx>
```

When no path is supplied, the launcher looks for `Master Spreadsheet Template.xlsx`
beside the batch file. The launcher currently passes `--overwrite` and therefore
allows existing output workbooks to be replaced. It uses the Just virtual
environment under `%LOCALAPPDATA%` when available, then falls back to `python`.

## Processing flow

1. Read the master XLSX package and control metadata.
2. Validate named inputs, tables, worksheet names, dates, selections, and
	output destinations before creating outputs.
3. Expand selected mappings and stresses as a Cartesian product.
4. Remove unrequired worksheets and rewrite formulas and defined names that
	refer to renamed or removed worksheets.
5. Populate the sensitivity template and write the control values for the
	current combination.
6. Publish each workbook atomically and continue recording individual failures.
7. Write one JSON manifest in each distinct stress output folder.

Each manifest contains the master path, timestamp, valuation date, output
folder, and a result entry for every combination routed to that folder. Failed
combinations include their error message and do not hide successful outputs.

## Main files

- `master_workbook_orchestrator_merged.py`: consolidated implementation and CLI
- `run_master_workbook_orchestrator_package.bat`: Windows launcher
- `tests/`: workbook and orchestration tests
- `Archive/`: retained legacy scripts and run artefacts

## Development and testing

The main implementation is in `master_workbook_orchestrator_merged.py`.
Tests are in `tests/`, with workbook fixtures constructed in
`tests/test_master_workbook_orchestrator.py`.

Useful checks are:

```text
python -m py_compile master_workbook_orchestrator_merged.py
pytest -q
```

The current test module imports legacy modules that are not present at the
repository root in this standalone copy. Until those modules are restored or
the imports are migrated, `pytest` stops during collection with
`ModuleNotFoundError` before running tests.
