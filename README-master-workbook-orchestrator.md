# Master Workbook Orchestrator

This standalone implementation leaves the original workbooks and legacy scripts unchanged.

## Control contract

The master workbook must contain a worksheet named `Global Control` with these Excel tables:

- `tbl_entity_mapping`: `Entity`, `Category`, `SourceSheet`, `Selected`
- `tbl_stress_selection`: `Stress`, `Transition Tab`, `Spread Tab`, `Suffix`, `OutputFolder`, `Selected`
- `tbl_run_settings`: `Setting`, `Value`

Run settings must include exactly one each of `ValuationDate`, `risk_factor_path`, and `bool_inv_exp`. `ValuationDate` may be in ISO format or an Excel date. `risk_factor_path` must be non-blank. Each selected stress row must provide existing `Transition Tab` and `Spread Tab` names, a non-blank `OutputFolder`, and a `Suffix`. Unselected rows may leave the folder blank. Relative folders are resolved from the master workbook's folder. The eight static sheets are every sheet other than `Global Control`, `I. Control`, the four data sheets, the configured transition/spread tabs, and the sensitivity template.

The master workbook must contain a hidden worksheet named `I. Sensitivity - Template`. Its transition tables, spread tables, and `_Orig` spread tables are retained; the selected transition matrix is copied into every transition table and the selected spread matrix into every spread table.

The master workbook must also contain `I. Control` and these workbook-scoped named inputs: `valuation_date`, `entity_name`, `risk_factor_path`, `output_folder_path`, `prm_LPIgeneratorWB`, `output_suffix`, and `bool_inv_exp`. Each generated workbook includes a copied `I. Control` sheet with those names pointing to its cells. The date comes from `ValuationDate`; `entity_name` uses the mapping's `Entity`; `risk_factor_path` and `bool_inv_exp` come from run settings; `output_folder_path` and `output_suffix` come from the selected stress; and `prm_LPIgeneratorWB` is blank.

Each output workbook has one data sheet, `I. Control`, and one sensitivity sheet, plus the eight static sheets. The data and sensitivity sheets are renamed to `I. Bonds Data` and `I. Sensitivity` by default. Override with `--data-sheet-name` and `--sensitivity-sheet-name`.

## Run

```text
python master_workbook_orchestrator.py --master <path> [--overwrite] [--data-sheet-name <text>] [--sensitivity-sheet-name <text>]
```

Or use `run_master_workbook_orchestrator.bat <master-workbook.xlsx>`.

The master workbook and its control/template sheets are loaded read-only in practice: this tool never saves them. Each output is written through a temporary file and a JSON manifest records successes and failures in each distinct stress output folder. Existing outputs are rejected unless `--overwrite` is supplied. The suffix and output folder from each selected stress row are used for its outputs. There is no global output folder or `--output-folder` override.