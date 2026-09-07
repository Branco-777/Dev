# Decisions

## Sensitivity Tab Composition

- **Decision:** Replace the selected stress worksheet input with two source-tab
  inputs: `Transition Tab` and `Spread Tab`.
- **Decision:** Each selected control row represents one transition/spread pair
  and generates one output workbook.
- **Decision:** Build `I. Sensitivity` from the fixed `I. Sensitivity - Template`
  worksheet in the master workbook. The template supplies the authoritative
  layout, ranges, styles, labels, and Excel table definitions.
- **Decision:** Each source tab contains one matrix. Replicate the transition
  matrix into all four transition blocks (`GBP_FIN`, `GBP_NONFIN`, `USD_FIN`,
  `USD_NONFIN`) and replicate the spread matrix into all four spread blocks and
  all four `_Orig` spread blocks.
- **Decision:** Populate the final sensitivity tab with source values while
  preserving the template's formatting and table objects.
- **Decision:** Derive target table names and ranges from the supplied template
  workbook rather than hard-coding the currently observed ranges.
- **Rationale:** The current implementation copies complete worksheets and has
  no safe range-level composition contract. The template is required before
  implementation to avoid silently placing matrices in incorrect cells.
- **Status:** Implemented against the hidden `I. Sensitivity - Template` sheet
  inside the master workbook; its table layout is retained and populated by
  source matrix values.

## Decision 1: Populate I. Control in Generated Outputs (26 Aug 2026)

**Question:** How should generated workbooks carry the seven named inputs used
by the Bonds Lite workflow?

**Decided by:** User

**Choice:** Copy `I. Control` as an explicit additional output sheet. Preserve
the seven workbook-scoped names and populate literal values from run settings,
the selected entity mapping, and the selected stress row. Add mandatory
`risk_factor_path` and `bool_inv_exp` rows to `tbl_run_settings`. Use the
`Entity` value only for `entity_name`, preserve path text exactly, and leave
`prm_LPIgeneratorWB` blank.

**Alternatives considered:**
- Formula links to `Global Control` — rejected because outputs do not currently
  include `Global Control` and should remain self-contained.
- Include `I. Control` within the eight static sheets — rejected because the
  existing eight-sheet contract is retained and `I. Control` is explicit.
- Create missing names or sheets automatically — rejected because required
  structure should fail clearly when absent.

**Rationale:** Each output must be directly usable by the existing Bonds Lite
workflow while retaining entity- and stress-specific values. Literal values
avoid external references, and explicit validation prevents malformed outputs.