# Plan: Add I. Control to Generated Outputs

## Overview

Extend the master workbook orchestrator so every generated workbook contains a
copied `I. Control` sheet with output-specific values for the seven existing
workbook-scoped named inputs. Outputs remain self-contained: values are written
literally from `Global Control` and the selected entity/stress rows, while the
defined names continue to point to the corresponding cells on `I. Control`.

## Agreed Contract

- `I. Control` is required in the master workbook and is copied as an explicit
  output sheet named `I. Control`.
- The existing eight static sheets remain unchanged; `I. Control` is an
  additional copied sheet rather than one of those eight.
- The seven required workbook-scoped names are `valuation_date`,
  `entity_name`, `risk_factor_path`, `output_folder_path`,
  `prm_LPIgeneratorWB`, `output_suffix`, and `bool_inv_exp`.
- In the canonical workbook, these names point to `I. Control!D10`, `D13`,
  `D16`, `D19`, `D22`, `D25`, and `D28` respectively. The implementation should
  validate the names and retain their destinations when creating outputs.
- `valuation_date` comes from `ValuationDate` in `tbl_run_settings`.
- `entity_name` comes from `Entity`, so `JRL MAP` receives `JRL` and `PLACL
  MAP` receives `PLACL`.
- `risk_factor_path` comes from the new mandatory `risk_factor_path` row in
  `tbl_run_settings`.
- `output_folder_path` comes from the selected stress row's `OutputFolder`,
  preserving the exact control-table text.
- `output_suffix` comes from the selected stress row's `Suffix`.
- `prm_LPIgeneratorWB` is explicitly blank in every output.
- `bool_inv_exp` comes from the new mandatory `bool_inv_exp` row in
  `tbl_run_settings`; the recommended value is `TRUE`.
- Missing `I. Control`, names, cells, or required settings fail with a clear
  `OrchestratorError` before any output is published.

## Steps

1. **Extend control parsing** in `master_workbook_orchestrator.py` so
   `tbl_run_settings` accepts and validates exactly one `risk_factor_path` and
   one `bool_inv_exp` setting in addition to `ValuationDate`. Preserve the
   existing rejection of a global `OutputFolder` setting and per-stress folder
   behaviour.
2. **Validate the control sheet and names** before output generation. Confirm
   `I. Control` exists, each required defined name exists, each is
   workbook-scoped, and each resolves to a single cell on `I. Control`.
3. **Add an output-control helper** that copies `I. Control` using the existing
   sheet-copy behaviour, writes the seven named-cell values for one
   entity/stress combination, and explicitly blanks `prm_LPIgeneratorWB`.
   Preserve styles, merged cells, dimensions, and labels.
4. **Preserve named-input discoverability** in each new workbook by recreating
   the seven workbook-scoped defined names so they refer to the copied
   `I. Control` cells, rather than leaving names pointing at the master or
   dropping them during workbook creation.
5. **Integrate output ordering** in `copy_combination()` so `I. Control` is
   copied as a dedicated sheet, alongside the data sheet, composed sensitivity
   sheet, and eight static sheets. Ensure it is not a static-sheet candidate.
6. **Update documentation** in `README.md`, the
   existing master plan, and the decision record with the two new settings, the
   seven named inputs, output values, and explicit-sheet count.

All implementation steps are complete. The implementation validates the
control contract before output generation and recreates the seven names in each
saved workbook.

## Test Strategy

Follow the red-green-refactor sequence from the `testing-strategy` skill.

1. Extend the synthetic workbook fixture with `I. Control`, the seven defined
   names, and the two new settings rows.
2. Add parser tests for valid setting extraction, duplicate/missing setting
   failures, and preservation of existing valuation-date and stress-folder
   behaviour.
3. Add unit tests for missing `I. Control`, missing names, sheet-scoped names,
   and names resolving outside `I. Control`; each should fail clearly before
   output creation.
4. Add an integration test with one stress and four selected mappings. Open
   all four outputs and assert `I. Control` exists; the date, entity, risk path,
   output folder, suffix, blank LPI value, and boolean match the expected
   setting/mapping/stress values; and every defined name is present,
   workbook-scoped, and points to output `I. Control`.
5. Assert the original master workbook remains unchanged after a run, and the
   existing sensitivity/table composition tests continue to pass.
6. Run focused orchestrator tests, then the full `pytest -q` suite and
   diagnostics. A production workbook smoke test should verify compatibility
   with the canonical named ranges before release.

## Affected Files

- `master_workbook_orchestrator.py` — parse settings, validate names, populate
  and copy `I. Control`, and preserve defined names.
- `tests/test_master_workbook_orchestrator.py` — fixture, validation, and
  end-to-end output assertions.
- `README.md` — document the expanded settings
  and output contract.
- `references/plan-master-workbook-orchestrator.md` — record this feature.
- `references/decisions.md` — record the user-approved design choices.

## Decisions

- Use literal output values rather than formula links because generated
  workbooks do not currently include `Global Control`.
- Copy `I. Control` as an additional sheet, preserving the existing eight
  static sheets.
- Preserve path text exactly as entered in the control tables.
- Use the `Entity` column only for `entity_name`.
- Require both new settings rows and clear missing-contract errors.
- Force `prm_LPIgeneratorWB` blank in every output.

## Risks and Unknowns

- Defined-name APIs differ across `openpyxl` versions; test round-tripping
  through a saved workbook using the installed version.
- The canonical workbook exposes the seven names as workbook-scoped names. A
  future sheet-scoped variant should be rejected unless the contract changes.
- Excel formulas depending on these names may require recalculation when the
  output is opened; the orchestrator writes values and does not calculate
  formulas.

## Skills Consulted

- `questioning-patterns` — requirements and decision clarification.
- `testing-strategy` — red-green-refactor and integration coverage.
- `regression-testing` — representative workbook smoke/regression checks.
- `performance-optimisation` — no optimisation without measured evidence.
- `decision-logging` — recording user-selected contract decisions.
- `document-extraction` — inspecting workbook names and destinations.