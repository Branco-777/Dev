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

## Decision 2: Consolidated Package-Level Orchestrator (7 Sep 2026)

**Question:** How should the openpyxl and direct XLSX package orchestrators be
combined while prioritising performance?

**Decided by:** User

**Choice:** Add one new public runner using direct XLSX package editing as its
default backend. Centralise Control validation, output-sheet validation,
formula-reference rewriting, collision handling, result collection, manifest
writing, and CLI exit status in that runner. Rewrite formulas when source or
template sheets are renamed. Fail before writing any output when a
deterministic destination already exists and overwrite is disabled.

**Alternatives considered:**
- Use openpyxl as the only backend — rejected because the user prioritised
  performance for the consolidated implementation.
- Keep two public runners — rejected because it leaves validation and failure
  behaviour divergent.
- Preserve old sheet names rather than rewriting formulas — rejected because
  generated outputs use stable output sheet names and retained formulas must
  continue to refer to the renamed sheets.
- Record output collisions per combination — rejected because it can publish a
  partial run before the user has resolved a deterministic collision.

**Rationale:** One package-level entry point retains the performance benefit
  while presenting one validated operational contract. Preflight collision
  checks and atomic workbook writes prevent incomplete or misleading output
  sets.

### Decision 3: Selective XLSX Startup Loading (7 Sep 2026)

**Question:** Can the approximately 19-second initial workbook load be reduced
without weakening Control validation or changing generated outputs?

**Decided by:** Agent

**Choice:** Read and decompress the XLSX package once, then build a lightweight
  validation view from workbook metadata, defined names, and `Global Control`
  worksheet XML. Keep the original package bytes for output generation and
  retain openpyxl-compatible helper paths for existing callers.

**Alternatives considered:**
- Continue using full `openpyxl.load_workbook()` — rejected because it parses
  large data worksheets that are copied as package parts and not needed for
  startup validation.
- Use openpyxl `read_only=True` — rejected because read-only worksheets do not
  expose the Excel table API required by the Control parser.
- Parse every worksheet XML part — rejected because only Control cell values,
  workbook sheet metadata, and defined names are needed before output creation.

**Rationale:** The selective reader reduces the measured startup path from
  approximately 19 seconds to approximately 0.4 seconds on the shared master
  workbook while preserving the existing validation contract. The real
  end-to-end run completed four outputs successfully in 17.29 seconds.