## Review: Orchestrator Consolidation

### Summary
The two files can be merged into one supported entry point, but they are not currently interchangeable implementations. The package runner is a useful performance backend, while the openpyxl runner contains stronger validation and different failure semantics. Consolidate the shared orchestration contract first, then keep package-level XLSX mutation behind the same public runner.

### Test Status
- `pytest`: 34 passed, 0 failed, 0 skipped, 13 warnings
- New tests added: no
- `ruff`: unavailable in the environment
- Additional probe: renaming a worksheet with openpyxl leaves `='JRL'!A1` unchanged

### Findings

#### Blocking (must fix before commit)
- [master_workbook_orchestrator_package.py:284](../master_workbook_orchestrator_package.py#L284) **Package runner bypasses required named-input validation** — `_copy_output()` indexes `workbook.defined_names` directly and assumes every required name has one valid destination. Unlike [master_workbook_orchestrator.py:238](../master_workbook_orchestrator.py#L238), it does not validate `I. Control`, workbook scope, single-cell destinations, or missing names. A malformed workbook therefore raises `KeyError` or `IndexError`, neither of which is caught by [master_workbook_orchestrator_package.py:342](../master_workbook_orchestrator_package.py#L342), and the run can abort without a manifest. Move `_control_named_cells()` into the shared preflight and pass its validated coordinates to either backend.
- [master_workbook_orchestrator_package.py:276](../master_workbook_orchestrator_package.py#L276) **The package backend can create duplicate output worksheet names** — it checks only blank values, length, and equality between the two requested output names. It does not reject a data or sensitivity name that already exists in `static_names`, `I. Control`, or another retained sheet. The openpyxl path does reject conflicts with static sheets at [master_workbook_orchestrator.py:655](../master_workbook_orchestrator.py#L655). A single public runner must use one shared output-name validator before either backend mutates the workbook.
- [master_workbook_orchestrator_package.py:248](../master_workbook_orchestrator_package.py#L248) **Renaming sheets does not update worksheet formulas** — `_update_workbook()` rewrites selected defined names, but it never traverses formulas in retained worksheet XML. The openpyxl path also needs an explicit strategy; the direct probe confirmed that `worksheet.title = "I. Bonds Data"` leaves `='JRL'!A1` unchanged. Any retained formula referring to the source or sensitivity template can point at a deleted or renamed tab in the generated workbook. Add a regression fixture with cross-sheet formulas and either rewrite all affected formulas or preserve the source sheet names.
- [master_workbook_orchestrator.py:796](../master_workbook_orchestrator.py#L796) and [master_workbook_orchestrator_package.py:362](../master_workbook_orchestrator_package.py#L362) **The two CLI entry points report failures differently** — the openpyxl `main()` always returns `0` after writing manifests, even when `run()` recorded failed combinations, while the package `main()` returns `1` when a manifest contains a failure. Their collision behaviour also differs: the openpyxl runner preflights and aborts on any existing output at [master_workbook_orchestrator.py:731](../master_workbook_orchestrator.py#L731), while the package runner records per-combination failures and continues at [master_workbook_orchestrator_package.py:337](../master_workbook_orchestrator_package.py#L337). Choose one batch contract and implement it in the merged runner; otherwise automation will behave differently depending on which file is launched.

#### Suggestions (nice to have)
- [master_workbook_orchestrator.py:705](../master_workbook_orchestrator.py#L705) and [master_workbook_orchestrator_package.py:317](../master_workbook_orchestrator_package.py#L317) **Extract the shared run pipeline** — both implementations load and serialise the workbook, expand combinations, handle overwrite, collect results, and write per-folder manifests. Keep one `run()` and one `main()` in `master_workbook_orchestrator.py`; make the workbook-copy operation a backend function selected internally or by a narrowly scoped option. Reuse `RunResult` rather than the package runner's untyped dictionaries.
- [master_workbook_orchestrator.py:539](../master_workbook_orchestrator.py#L539) and [master_workbook_orchestrator_package.py:297](../master_workbook_orchestrator_package.py#L297) **Share control-value construction** — the same seven control values are assembled twice, once for openpyxl cells and once for XML cells. A shared `control_values()` helper would reduce drift while leaving the backend-specific cell writer separate.
- [master_workbook_orchestrator_package.py:236](../master_workbook_orchestrator_package.py#L236) **Remove dead intermediate state** — `worksheet_relationship_ids` is calculated but never used. The package module also imports `CONTROL_SHEET` without using it. These are small signs that the package implementation was evolved beside, rather than integrated with, the main path.
- [master_workbook_orchestrator_package.py:216](../master_workbook_orchestrator_package.py#L216) **Add performance evidence before making the package backend the default** — the package path is plausibly faster because it avoids workbook reconstruction, but no benchmark or memory budget exists. Capture runtime and output-size measurements for a representative workbook and retain an integration test that opens the generated file with openpyxl or Excel-compatible validation.

### Praise (what was done well)
- Both runners already share the domain model and core parsing through `Combination`, `parse_control()`, `static_sheet_names()`, and `output_name()`.
- The package path preserves original XLSX parts directly, which is a sensible performance direction for large workbooks.
- The current suite exercises both output paths and passes cleanly: 34 tests passed.
- The source-tab validation is now table-driven and is reused by static-sheet derivation.

### Decisions Check
- The existing `references/decisions.md` documents the control-sheet and sensitivity-composition decisions, but it does not record a choice between openpyxl copying and direct XLSX package editing.
- The user selected performance as the review priority. No implementation decision was made during this read-only review, so no decision-log entry was added.
- Repository team memory and lock files were absent; the review used the base Just standards only, as confirmed by the user.

### Recommended Next Steps
- Keep `master_workbook_orchestrator.py` as the single CLI/module surface.
- Centralise validation, output naming checks, control-value construction, result collection, manifest writing, and exit-code policy.
- Move the package implementation behind a private copy backend and make both backends satisfy the same contract.
- Add parity tests for missing named inputs, sheet-name conflicts, cross-sheet formulas, collision handling, and non-zero CLI exit status.
- Benchmark both backends on a representative workbook before selecting the package backend as the default.

## Skills consulted
- coding-standards
- testing-strategy
- regression-testing
- performance-optimisation
- decision-logging
- git-workflow
