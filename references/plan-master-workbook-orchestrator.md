## Plan: Master Workbook Orchestrator

### Overview
Extend the existing Python workbook updater with a small orchestration layer. The
master workbook remains the user-facing configuration and data source, while
Python validates the Control tab, creates one output workbook per entity/stress
combination, and copies only the requested data, sensitivity, and static tabs.

The recommended distribution is a Python core launched by a `.bat` file. Keep
VBA optional as a convenience button only; it should not contain the business
logic or be required for execution.

### Recommended Architecture

1. **Master workbook contract**
  - Add a `Global Control` sheet with Excel tables for the entity selection,
   valuation date, and selected stresses. Each selected stress row carries
   its own output suffix and output folder.
   - Use an explicit mapping table for entity/category to source sheet:
     `JRL`, `JRL MAP`, `PLACL`, and `PLACL MAP`.
  - Use a stress table containing the stress name, exact worksheet name
   `Sensitivities_<stress name>`, output suffix, output folder, and a
   `Selected` flag. A selected row must have a non-blank output folder;
   unselected rows may leave it blank.
  - Define static sheets by rule: every sheet other than `Global Control`, the four
     data sheets, and sheets matching `Sensitivities_*` is copied as static
     content. Validate that this produces the expected eight sheets, while
     allowing the script to report a clear error if the template changes.
   - Treat the master workbook as read-only during processing. Users save and
     close it before launching the run.

2. **Python orchestration layer**
   - Add a dedicated orchestration module or entry-point beside the current
     updater, rather than putting all logic into a batch file or VBA.
   - Parse the Control sheet and validate all values before creating outputs.
   - Expand the Cartesian product of selected entities/categories and selected
     stresses. For two entities, four categories, and two stresses, this yields
     eight outputs.
   - For each combination, create a unique temporary working copy, retain the
     selected data sheet, selected `Sensitivities_<stress>` sheet, and the eight
     static sheets, then save the final output using a deterministic name.
   - First determine whether the existing updater's formula-rewrite workflow is
     required for these outputs. If it is, expose a reusable processing function;
     otherwise use a focused workbook-copy function rather than invoking logic
     designed for the current four-input summary workflow.
  - Write one run manifest in each distinct stress output folder containing
   the master path, timestamp, valuation date, selected combinations, output
   paths, and any failures for that folder.

3. **Batch distribution**
   - Keep `setup_summary_comparison_tool.bat` for one-time per-PC environment
     setup and dependency installation.
   - Add or adapt a run launcher that passes the master workbook path and uses
     the Python virtual environment under `%LOCALAPPDATA%`.
  - Use an explicit `--master` argument so the launcher is testable and does
   not depend on the current working directory. Output folders come only from
   selected stress rows.
   - Generate unique output names, for example including valuation date, stress,
    category, and stress-specific suffix. Reject an existing output unless overwrite is
     explicitly enabled.
   - Correct the existing launcher messages that refer to the non-existent
     `setup_summary_comparison_tool_user_folder.bat`.

4. **Optional Excel integration**
   - If desired, add a minimal macro or button that calls the run `.bat` with
     the current workbook path.
   - Keep the macro as a launcher only. Do not reproduce validation, sheet
     selection, copying, or formula logic in VBA.
   - Confirm macro-signing, trusted-location, and SharePoint policy requirements
     before making this the default user experience.

### Why This Approach

- Python is better suited to repeated workbook generation, validation, logging,
  command-line testing, and reuse of the existing `openpyxl`/`pywin32` code.
- A `.bat` file gives mixed-technicality users a double-click workflow while
  retaining a transparent fallback for support and automation.
- VBA alone would couple the implementation to Excel, complicate deployment and
  macro security, and make the Cartesian-product workflow harder to test.
- A separate executable packaged with PyInstaller could remove the Python
  prerequisite later, but should follow a working Python-based implementation
  because Excel COM and network synchronisation still need operational testing.

### Proposed Command-Line Interface

The Python entry point should support an explicit master path and optional
operational overrides:

```text
python master_workbook_orchestrator.py --master <path>
  [--output-folder <path>] [--overwrite]
```

Control-tab values should be the source of truth. The `.bat` launcher should
pass only the master workbook path explicitly or default to a documented shared
location. There is no global output-folder CLI override.

### Steps

1. Specify and freeze the `Global Control` sheet contract, worksheet-name rules, and
   output naming convention. Verification: invalid entities, stresses, dates,
   missing sheets, and duplicate names produce clear validation errors before
   any output is created.
2. Add unit tests for Control-sheet parsing and combination generation before
  implementation. Verification: the example produces exactly eight distinct
  combinations, while duplicate, unavailable, and unselected stresses are
  rejected appropriately. Also verify that the derived static-sheet rule finds
  exactly eight sheets and excludes `Control`, all data sheets, and all stress
  sheets.
3. Extract or introduce a reusable workbook-copy/process function around the
   existing updater. Verification: one combination preserves formatting and
   formulas, includes the selected data/sensitivity/static sheets, and excludes
   other combination-specific sheets.
4. Implement the orchestrator and per-folder run manifests. Verification: a
  multi-output run routes every combination to its stress folder, creates one
  manifest per distinct folder, is atomic per output, reports failed
  combinations individually, and a rerun cannot silently overwrite an
  existing result.
5. Adapt the batch setup/run launchers and document the supported double-click
   and command-line workflows. Verification: a clean test machine can install
   locally from the shared folder and run with an explicit master path.
6. Perform an end-to-end test with a copied master workbook and representative
   source data, including a SharePoint-synced path and a deliberately missing
   source. Verification: valid files open and calculate in Excel, and missing
   inputs fail before partial or misleading outputs are published.

### Test Strategy

- Unit tests: Excel-table Control parsing, entity/category mapping, stress
  worksheet resolution, static-sheet derivation, Cartesian-product expansion,
  filename sanitisation, per-stress output-folder resolution, and output
  collision handling.
- Workbook tests: static-sheet copying, selected-sheet filtering, formula and
  formatting preservation, and expected sheet order where relevant.
- Integration tests: one combination, all combinations across multiple stress
  folders, partial failure, per-folder manifests, and rerun into an existing
  output folder.
- Operational tests: shared-folder access, local OneDrive availability, Excel
  installation/COM behaviour, and simultaneous runs by two users.
- Use a small synthetic workbook fixture for repeatable tests, plus one approved
  real-world workbook for final regression testing. No automated test suite was
  found in the current workspace, so adding a focused `tests/` structure should
  be part of the implementation.

### Affected Files

- `summary_comparison_fixed_formula_updater.py` — expose reusable processing
  functionality or accept the orchestration inputs without duplicating logic.
- `summary_comparison_fixed_formula_updater_local_paths.py` — preserve the
  existing local-path compatibility wrapper if the shared processing entry point
  is changed.
- `run_summary_comparison_tool.bat` — launch the master-workbook orchestrator
  with explicit arguments and useful error handling.
- `setup_summary_comparison_tool.bat` — retain local environment setup and fix
  stale launcher references if needed.
- `README.md` — document the Control sheet contract, naming rules, setup, and
  run procedures.
- `README-master-workbook-orchestrator.md` — document the per-stress
  `OutputFolder` column and remove the global setting and CLI option.
- `run_master_workbook_orchestrator.bat` — accept only the master workbook and
  stop forwarding a global output folder.
- `tests/` — add parser, combination, workbook-copy, and integration tests.

### Decisions

- Use Python as the business-logic and workbook-processing layer.
- Use a `.bat` launcher as the default distribution interface for mixed users.
- Treat VBA as optional UI integration, not as the processing engine.
- Keep the master workbook read-only during a run and create isolated outputs.
- Use explicit sheet-name mappings and names rather than tab positions.
- Use a required suffix on each selected stress row; do not use a global suffix.
- Use a required `OutputFolder` on each selected stress row; do not use a
  global `tbl_run_settings.OutputFolder` value or `--output-folder` override.
- Resolve relative stress output folders against the master workbook directory.
- Write one manifest per distinct stress output folder, alongside that folder's
  generated workbooks.
- Start with per-machine virtual environments; evaluate PyInstaller only after
  the workflow is stable and operationally tested.

### Risks and Unknowns

- The current dependency list includes `pywin32`, so desktop Excel and COM
  behaviour remain deployment prerequisites even though the process is started
  from Python.
- Absolute external formula links may not work identically for every user's
  OneDrive or SharePoint sync root. A standardised shared path or a per-user
  path-resolution layer must be chosen.
- A master workbook being edited or synchronised during processing can produce
  inconsistent inputs. Saving, closing, and optionally copying locally before
  processing should be part of the operating procedure.
- The exact eight static sheets, stress naming, source-sheet formats, and whether
  outputs need formulas recalculated or only copied must be confirmed before
  implementation. The current decision is to preserve the existing updater's
  formula/link behaviour wherever that workflow applies.
  - Existing workbooks with only the global `tbl_run_settings.OutputFolder`
    setting will require migration to add `OutputFolder` to each selected stress
    row. Existing callers and launchers using `run(output_folder=...)` or
    `--output-folder` will require migration.

  ## Change Request: Per-Stress Output Folders

  ### Contract Change

  Add `OutputFolder` to `tbl_stress_selection`:

  ```text
  Stress | Worksheet | Suffix | OutputFolder | Selected
  ```

  `OutputFolder` is required and non-blank only when `Selected` is truthy. It is
  stored on each generated `Combination`. Remove `OutputFolder` from
  `tbl_run_settings`; that table retains only `ValuationDate`.

  ### Implementation Sequence

  1. Update the synthetic workbook fixture and add a red test for a selected
    stress with a blank folder. Add coverage proving blank folders on
    unselected stresses are accepted and that relative folders resolve against
    the master workbook directory.
  2. Change `Combination` and `parse_control()` to carry a stress-specific
    `Path`, remove the global-folder return value, and reject the legacy global
    setting according to the frozen contract.
  3. Change `run()` to group or route combinations by their resolved folder,
    perform collision checks within each folder, and write a manifest into every
    distinct folder. Serialise per-combination paths as strings.
  4. Remove `--output-folder` from `main()` and update the batch launcher to
    accept only the master workbook path.
  5. Update the README and examples, then run the focused parser and integration
    tests followed by the full test suite.

  ### Acceptance Criteria

  - Two selected stresses with different folders create each stress's outputs
    only in its assigned folder.
  - Each distinct output folder contains exactly one manifest for the run.
  - No global `OutputFolder` setting, fallback `outputs` directory, or
    `--output-folder` CLI option remains in the documented or executable
    contract.
  - A selected stress with a blank folder fails before any output is created.
  - An unselected stress may have a blank folder without failing validation.

## Addendum: Implementation Progress

- [✓ Implemented — see addendum] Per-stress `OutputFolder` is now required for
  selected stress rows and stored on each combination.
- [✓ Implemented — see addendum] Global `OutputFolder`, the default `outputs`
  fallback, and `--output-folder` have been removed from the executable
  contract.
- [✓ Implemented — see addendum] Relative folders resolve from the master
  workbook directory, and each distinct stress folder receives its own
  manifest.
- [✓ Implemented — see addendum] The launcher, README, parser tests, and
  multi-folder integration test have been updated.
- Verification: `pytest -q` passes with 17 tests.
- Remaining: end-to-end validation with an approved representative workbook,
  Excel recalculation, SharePoint-synced paths, and deliberately missing
  sources remain operational checks from Step 6.

## Addendum: Composed Sensitivity Outputs

### New Requirement

Replace the single stress `Worksheet` input with two columns in
`tbl_stress_selection`:

```text
Stress | Transition Tab | Spread Tab | Suffix | OutputFolder | Selected
```

Each selected row represents one transition/spread pair and produces one
output workbook. The output sensitivity worksheet is built from the fixed
`I. Sensitivity - Template` worksheet in the master workbook rather than
copied from one stress worksheet.

### Composition Contract

- The template workbook is authoritative for target table names, ranges,
  labels, styles, formulas, dimensions, and Excel table definitions.
- Each selected transition and spread source tab contains one matrix.
- Copy the transition matrix values into all four transition blocks:
  `GBP_FIN`, `GBP_NONFIN`, `USD_FIN`, and `USD_NONFIN`.
- Copy the spread matrix values into all four spread blocks and all four
  corresponding `_Orig` spread blocks.
- Preserve the template's formatting and table objects; write source values as
  values rather than source formulas.
- The source tabs must exist in the master workbook and must be validated before
  any output is published.

### Required Template Handoff

Before implementation, provide the canonical workbook containing
`I. Sensitivity - Template`, representative transition and spread source tabs,
and the relevant table definitions. The implementation should inspect that
workbook to derive and validate the source-to-target table mapping. The
currently observed ranges in a generated workbook are provisional and must not
be treated as the final mapping.

### Implementation Steps

1. Inspect the supplied template and source tabs, recording exact table names,
   ranges, dimensions, and value/formula behaviour. Add a minimal fixture that
   contains the template plus one transition and one spread source matrix.
2. Add failing parser tests for `Transition Tab` and `Spread Tab`, including
   missing source tabs, blank selected values, duplicate pair/output rows, and
   preservation of the existing stress-specific suffix and output-folder
   behaviour.
3. Add a range/table composition helper that copies values from each source
   matrix into every target block while retaining the template's table objects,
   styles, labels, dimensions, and sheet name `I. Sensitivity`.
4. Replace the whole-sheet stress copy in `copy_combination()` with template
   copying plus sensitivity composition. Keep the selected data sheet and
   static-sheet filtering unchanged.
5. Update the manifest to record both source tab names and the template name,
   then update README, launcher examples, and the persisted control contract.
6. Run focused unit/workbook tests, the full pytest suite, and an end-to-end
   test against the supplied representative workbook. Verify all four
   transition blocks, four spread blocks, and four `_Orig` blocks contain the
   expected values and that the master workbook is unchanged.

### Open Questions Before Coding

- The canonical template workbook and exact table names/ranges are still
  required.
- Confirm that source matrices have identical dimensions to their target
  blocks, or provide the intended row/column transformation if they do not.
- Confirm whether source values are available as cached values when the master
  is loaded with `data_only=True`; `openpyxl` does not calculate formulas.

### Implementation Progress

- [✓ Implemented] `tbl_stress_selection` now supplies `Transition Tab` and
  `Spread Tab` for each selected output.
- [✓ Implemented] The hidden `I. Sensitivity - Template` sheet is loaded from
  the master workbook itself; no separate template file or CLI override is
  required.
- [✓ Implemented] Template table layout, styles, and table metadata are
  retained while transition values populate transition tables and spread
  values populate spread and `_Orig` tables.
- [✓ Implemented] Focused and full test suites pass: 19 tests passed.
- [✓ Implemented] Add `I. Control` and its seven named-input values to every
  output; see `references/plan-i-control-output.md` for the agreed contract.
- Remaining: validate the supplied template with representative production
  transition/spread source tabs and confirm Excel recalculation/SharePoint
  behaviour.

## Skills consulted this session

- `questioning-patterns` — session opening and targeted requirements questions.
- `learn-more` — comparison of distribution and integration approaches.
- `developer-guides` — structure and clarity of the persistent implementation
  plan.