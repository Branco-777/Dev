"""Generate stress-specific workbooks from a validated master workbook."""

from __future__ import annotations

import argparse
from copy import copy
from io import BytesIO
import itertools
import json
import re
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn
from openpyxl.utils.cell import range_boundaries


CONTROL_SHEET = "Global Control"
ENTITY_TABLE = "tbl_entity_mapping"
STRESS_TABLE = "tbl_stress_selection"
SETTINGS_TABLE = "tbl_run_settings"
DATA_SHEETS = {"JRL", "JRL MAP", "PLACL", "PLACL MAP"}
EXPECTED_STATIC_SHEETS = 8
DEFAULT_DATA_SHEET_NAME = "I. Bonds Data"
DEFAULT_SENSITIVITY_SHEET_NAME = "I. Sensitivity"
SENSITIVITY_TEMPLATE_SHEET = "I. Sensitivity - Template"
CONTROL_OUTPUT_SHEET = "I. Control"
CONTROL_INPUT_NAMES = (
    "valuation_date",
    "entity_name",
    "risk_factor_path",
    "output_folder_path",
    "prm_LPIgeneratorWB",
    "output_suffix",
    "bool_inv_exp",
)
SETTING_ALIASES = {
    "ValuationDate": {"valuationdate", "valuation date"},
    "risk_factor_path": {"risk_factor_path", "risk factor stresses path"},
    "bool_inv_exp": {"bool_inv_exp", "calculate investment expenses"},
}
INVALID_FILENAME_CHARS = re.compile(r"[<>:\"/\\|?*\x00-\x1f]")


class OrchestratorError(ValueError):
    """Raised when the master workbook does not meet the Control contract.

    This exception identifies validation and workbook-structure failures that
    can be reported directly to the command-line user.
    """


@dataclass(frozen=True)
class Combination:
    """Describe one entity, category, and stress output combination.

    Attributes:
        entity: Entity selected for the generated workbook.
        category: Entity category selected for the generated workbook.
        source_sheet: Master workbook data sheet for the entity/category pair.
        stress: Selected stress name.
        transition_tab: Worksheet containing the transition matrix.
        spread_tab: Worksheet containing the spread matrix.
        stress_suffix: Optional suffix appended to the output filename.
        output_folder: Resolved folder for the generated workbook.
        risk_factor_path: Path value passed to the generated control sheet.
        bool_inv_exp: Investment expense calculation setting.
        output_folder_value: Original output-folder text from the control table.
    """

    entity: str
    category: str
    source_sheet: str
    stress: str
    transition_tab: str
    spread_tab: str
    stress_suffix: str = ""
    output_folder: Path | None = None
    risk_factor_path: str = ""
    bool_inv_exp: Any = True
    output_folder_value: str | None = None


@dataclass
class RunResult:
    """Record the result of one attempted output workbook.

    Attributes:
        combination: Serialisable values identifying the attempted combination.
        output_path: Created workbook path, or ``None`` if creation failed.
        status: Attempt status, normally ``"success"`` or ``"failed"``.
        error: Failure message, or ``None`` for a successful attempt.
    """

    combination: dict[str, str]
    output_path: str | None
    status: str
    error: str | None = None


def _table_rows(workbook: Any, table_name: str) -> list[dict[str, Any]]:
    """Read the rows from a named table on the Global Control sheet.

    Args:
        workbook: Open workbook containing the control worksheet.
        table_name: Excel table name to read.

    Returns:
        A list of dictionaries keyed by the table's column headings.

    Raises:
        OrchestratorError: If the control sheet, table, or a table heading is
            missing or invalid.
    """
    if CONTROL_SHEET not in workbook.sheetnames:
        raise OrchestratorError(f"Missing required worksheet: {CONTROL_SHEET}")
    worksheet = workbook[CONTROL_SHEET]
    if table_name not in worksheet.tables:
        raise OrchestratorError(f"Missing required table: {table_name}")
    table = worksheet.tables[table_name]
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    headers = [worksheet.cell(min_row, column).value for column in range(min_col, max_col + 1)]
    if any(not isinstance(header, str) or not header.strip() for header in headers):
        raise OrchestratorError(f"Table {table_name} has a blank column heading")
    return [
        dict(zip(headers, (worksheet.cell(row, column).value for column in range(min_col, max_col + 1))))
        for row in range(min_row + 1, max_row + 1)
    ]


def _required_text(row: dict[str, Any], key: str, table_name: str) -> str:
    """Return a required non-blank text value from a control-table row.

    Args:
        row: Control-table row represented as a dictionary.
        key: Column heading whose value is required.
        table_name: Table name used in validation errors.

    Returns:
        The stripped text value.

    Raises:
        OrchestratorError: If the value is missing, not text, or blank.
    """
    value = row.get(key)
    if not isinstance(value, str) or not value.strip():
        raise OrchestratorError(f"{table_name} requires a non-blank {key}")
    return value.strip()


def _selected(value: Any) -> bool:
    """Return whether a control-table selection value is truthy.

    Args:
        value: Boolean or text value from a control-table row.

    Returns:
        ``True`` for supported selected values; otherwise ``False``.
    """
    return value is True or str(value).strip().lower() in {"yes", "true", "1", "selected"}


def _setting(rows: list[dict[str, Any]], name: str) -> Any:
    """Return the unique run-setting value matching a configured name.

    Args:
        rows: Rows from the run-settings table.
        name: Canonical setting name, including any accepted aliases.

    Returns:
        The value stored in the matching row.

    Raises:
        OrchestratorError: If zero or more than one matching setting exists.
    """
    accepted_names = SETTING_ALIASES.get(name, {name})
    accepted_names = {value.casefold() for value in accepted_names}
    values = [
        row.get("Value")
        for row in rows
        if str(row.get("Setting", "")).strip().casefold() in accepted_names
    ]
    if len(values) != 1:
        raise OrchestratorError(f"Run settings must contain exactly one {name} value")
    return values[0]


def _parse_date(value: Any) -> date:
    """Convert a supported Excel or ISO date value to a date.

    Args:
        value: ``date``, ``datetime``, or ISO-formatted date value.

    Returns:
        The normalised calendar date.

    Raises:
        OrchestratorError: If the value is not a supported or valid date.
    """
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value.strip())
        except ValueError as error:
            raise OrchestratorError(f"Valuation date is invalid: {value}") from error
    raise OrchestratorError(f"Valuation date is invalid: {value}")


def _control_named_cells(workbook: Any) -> dict[str, str]:
    """Validate and return the cells targeted by required workbook names.

    Args:
        workbook: Open workbook containing the output control sheet.

    Returns:
        A mapping from required input name to its single-cell coordinate.

    Raises:
        OrchestratorError: If the control sheet, a required name, or a valid
            workbook-scoped single-cell destination is missing.
    """
    if CONTROL_OUTPUT_SHEET not in workbook.sheetnames:
        raise OrchestratorError(f"Missing required worksheet: {CONTROL_OUTPUT_SHEET}")
    named_cells: dict[str, str] = {}
    for name in CONTROL_INPUT_NAMES:
        defined_name = workbook.defined_names.get(name)
        if defined_name is None:
            raise OrchestratorError(f"Missing required named input: {name}")
        if defined_name.localSheetId is not None:
            raise OrchestratorError(f"Named input must be workbook-scoped: {name}")
        destinations = list(defined_name.destinations)
        if len(destinations) != 1 or destinations[0][0] != CONTROL_OUTPUT_SHEET:
            raise OrchestratorError(f"Named input must refer to I. Control: {name}")
        destination = destinations[0][1]
        if ":" in destination or "!" in destination:
            raise OrchestratorError(f"Named input must refer to one cell on I. Control: {name}")
        named_cells[name] = destination
    return named_cells


def parse_control(workbook: Any, base_folder: Path | None = None) -> tuple[date, list[Combination]]:
    """Parse the Control sheet and expand selected combinations.

    Args:
        workbook: Open master workbook to validate.
        base_folder: Folder used to resolve relative stress output folders.
            Defaults to the current working directory when omitted.

    Returns:
        A tuple containing the valuation date and the Cartesian product of
        selected entity/category mappings and stresses.

    Raises:
        OrchestratorError: If the control tables contain invalid, incomplete,
            duplicate, or unsupported values.
    """
    mappings = _table_rows(workbook, ENTITY_TABLE)
    stresses = _table_rows(workbook, STRESS_TABLE)
    settings = _table_rows(workbook, SETTINGS_TABLE)

    selected_mappings: list[tuple[str, str, str]] = []
    seen_mappings: set[tuple[str, str]] = set()
    for row in mappings:
        entity = _required_text(row, "Entity", ENTITY_TABLE)
        category = _required_text(row, "Category", ENTITY_TABLE)
        source_sheet = _required_text(row, "SourceSheet", ENTITY_TABLE)
        key = (entity, category)
        if key in seen_mappings:
            raise OrchestratorError(f"Duplicate entity/category mapping: {entity}/{category}")
        seen_mappings.add(key)
        if source_sheet not in DATA_SHEETS:
            raise OrchestratorError(f"Unsupported source sheet: {source_sheet}")
        if _selected(row.get("Selected")):
            selected_mappings.append((entity, category, source_sheet))

    selected_stresses: list[tuple[str, str, str, str, Path, str]] = []
    seen_stresses: set[str] = set()
    for row in stresses:
        stress = _required_text(row, "Stress", STRESS_TABLE)
        transition_tab = _required_text(row, "Transition Tab", STRESS_TABLE)
        spread_tab = _required_text(row, "Spread Tab", STRESS_TABLE)
        stress_suffix = _required_text(row, "Suffix", STRESS_TABLE)
        if stress in seen_stresses:
            raise OrchestratorError(f"Duplicate stress: {stress}")
        seen_stresses.add(stress)
        for tab_name, label in ((transition_tab, "Transition Tab"), (spread_tab, "Spread Tab")):
            if tab_name not in workbook.sheetnames:
                raise OrchestratorError(f"Missing {label}: {tab_name}")
        if _selected(row.get("Selected")):
            output_value = row.get("OutputFolder")
            if not isinstance(output_value, str) or not output_value.strip():
                raise OrchestratorError(f"{STRESS_TABLE} requires a non-blank OutputFolder for selected stress: {stress}")
            output_folder = Path(output_value.strip()).expanduser()
            if not output_folder.is_absolute():
                output_folder = (base_folder or Path.cwd()) / output_folder
            selected_stresses.append(
                (stress, transition_tab, spread_tab, stress_suffix, output_folder, output_value.strip())
            )

    if not selected_mappings:
        raise OrchestratorError("At least one entity/category mapping must be selected")
    if not selected_stresses:
        raise OrchestratorError("At least one stress must be selected")
    valuation_date = _parse_date(_setting(settings, "ValuationDate"))
    control_tables = workbook[CONTROL_SHEET].tables
    has_global_output_folder = any(
        str(row.get("Setting", "")).strip() == "OutputFolder"
        for table_name in control_tables
        for row in _table_rows(workbook, table_name)
    )
    if has_global_output_folder:
        raise OrchestratorError("Run settings must not contain global OutputFolder")
    risk_factor_path = _setting(settings, "risk_factor_path")
    if not isinstance(risk_factor_path, str) or not risk_factor_path.strip():
        raise OrchestratorError("Run settings require a non-blank risk_factor_path value")
    bool_inv_exp = _setting(settings, "bool_inv_exp")
    combinations = [
        Combination(
            entity,
            category,
            source_sheet,
            stress,
            transition_tab,
            spread_tab,
            stress_suffix,
            output_folder,
            risk_factor_path.strip(),
            bool_inv_exp,
            output_folder_value,
        )
        for (entity, category, source_sheet), (stress, transition_tab, spread_tab, stress_suffix, output_folder, output_folder_value) in itertools.product(
            selected_mappings, selected_stresses
        )
    ]
    return valuation_date, combinations


def static_sheet_names(workbook: Any, excluded: set[str] | None = None) -> list[str]:
    """Return static sheets according to the master workbook rule.

    Args:
        workbook: Open master workbook whose sheet names are inspected.
        excluded: Additional sheet names that belong to the current
            combination and must not be treated as static sheets.

    Returns:
        The ordered list of static worksheet names.

    Raises:
        OrchestratorError: If the workbook does not contain the expected
            number of static sheets.
    """
    stress_tabs = {
        tab_name
        for row in _table_rows(workbook, STRESS_TABLE)
        for tab_name in (row.get("Transition Tab"), row.get("Spread Tab"))
        if isinstance(tab_name, str) and tab_name.strip()
    }
    static = [
        name
        for name in workbook.sheetnames
        if name != CONTROL_SHEET
        and name not in (excluded or set())
        and name not in DATA_SHEETS
        and name != CONTROL_OUTPUT_SHEET
        and name not in stress_tabs
        and not (name.startswith("Sensitivities - ") or name.startswith("Sensitivities_"))
    ]
    if len(static) != EXPECTED_STATIC_SHEETS:
        raise OrchestratorError(
            f"Expected {EXPECTED_STATIC_SHEETS} static sheets, found {len(static)}"
        )
    return static


def sanitise_filename(value: str) -> str:
    """Make a value suitable for use in a Windows filename.

    Args:
        value: Text to sanitise.

    Returns:
        The value with invalid Windows filename characters replaced, or
        ``"unnamed"`` when no usable text remains.
    """
    result = INVALID_FILENAME_CHARS.sub("_", value).strip().rstrip(".")
    return result or "unnamed"


def output_name(combination: Combination, valuation_date: date) -> str:
    """Build a deterministic output filename for a combination.

    Args:
        combination: Entity/category/stress values for the output.
        valuation_date: Valuation date included in the filename.

    Returns:
        A sanitised ``.xlsx`` filename derived from the combination.
    """
    parts = [
        valuation_date.isoformat(),
        combination.entity,
        combination.category,
        combination.stress,
    ]
    if combination.stress_suffix:
        parts.append(combination.stress_suffix)
    return "_".join(sanitise_filename(part) for part in parts) + ".xlsx"


def _copy_sheet(source: Any, target: Any) -> None:
    """Copy worksheet values and supported presentation metadata.

    Args:
        source: Worksheet to copy from.
        target: Worksheet to copy into.

    Raises:
        ValueError: If the target worksheet cannot accept copied worksheet
            structures such as merged cells or tables.
    """
    for row in source.iter_rows():
        for source_cell in row:
            if isinstance(source_cell, MergedCell):
                continue
            target_cell = target[source_cell.coordinate]
            target_cell.value = source_cell.value
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.number_format = source_cell.number_format
                target_cell.protection = copy(source_cell.protection)
            if source_cell.hyperlink:
                target_cell._hyperlink = source_cell.hyperlink
            if source_cell.comment:
                target_cell.comment = source_cell.comment
    for merged_range in source.merged_cells.ranges:
        target.merge_cells(str(merged_range))
    for key, dimension in source.column_dimensions.items():
        target.column_dimensions[key].width = dimension.width
        target.column_dimensions[key].hidden = dimension.hidden
    for key, dimension in source.row_dimensions.items():
        target.row_dimensions[key].height = dimension.height
        target.row_dimensions[key].hidden = dimension.hidden
    target.freeze_panes = source.freeze_panes
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    for table in dict.values(source.tables):
        copied_table = Table(displayName=table.displayName, ref=table.ref)
        copied_table.headerRowCount = table.headerRowCount
        copied_table.totalsRowCount = table.totalsRowCount
        copied_table.totalsRowShown = table.totalsRowShown
        if table.tableStyleInfo is not None:
            copied_table.tableStyleInfo = copy(table.tableStyleInfo)
        min_col, min_row, max_col, _ = range_boundaries(table.ref)
        source_columns = list(table.tableColumns)
        for offset, column in enumerate(range(min_col, max_col + 1), 1):
            source_column = source_columns[offset - 1] if offset <= len(source_columns) else None
            header = source_column.name if source_column is not None else target.cell(min_row, column).value
            copied_table.tableColumns.append(
                TableColumn(id=offset, name=str(header) if header is not None else f"Column{offset}")
            )
        target.add_table(copied_table)


def _copy_matrix(source: Any, target: Any, target_table: Any) -> None:
    """Copy one source matrix into a target table of matching dimensions.

    Args:
        source: Worksheet containing exactly one source matrix table.
        target: Worksheet containing the destination table.
        target_table: Destination table receiving the matrix values.

    Raises:
        OrchestratorError: If the source does not contain exactly one table or
            the source and destination matrix dimensions differ.
    """
    source_tables = list(dict.values(source.tables))
    if len(source_tables) != 1:
        raise OrchestratorError(
            f"Source matrix tab must contain exactly one table: {source.title}"
        )
    source_table = source_tables[0]
    source_min_col, source_min_row, source_max_col, source_max_row = range_boundaries(source_table.ref)
    target_min_col, target_min_row, target_max_col, target_max_row = range_boundaries(target_table.ref)
    source_shape = (source_max_row - source_min_row, source_max_col - source_min_col)
    target_shape = (target_max_row - target_min_row, target_max_col - target_min_col)
    if source_shape != target_shape:
        raise OrchestratorError(
            f"Matrix dimensions do not match: {source.title} {source_table.ref} and "
            f"target {target_table.name} {target_table.ref}"
        )
    for row_offset in range(source_max_row - source_min_row + 1):
        for column_offset in range(source_max_col - source_min_col + 1):
            target_cell = target.cell(target_min_row + row_offset, target_min_col + column_offset)
            target_cell.value = source.cell(
                source_min_row + row_offset, source_min_col + column_offset
            ).value


def _compose_sensitivity(template: Any, transition: Any, spread: Any, target: Any) -> None:
    """Create a sensitivity worksheet from a template and two matrices.

    Args:
        template: Worksheet providing the sensitivity layout and formatting.
        transition: Worksheet containing the transition matrix.
        spread: Worksheet containing the spread matrix.
        target: Worksheet to receive the copied template and matrix values.
    """
    _copy_sheet(template, target)
    _populate_sensitivity(target, transition, spread)


def _populate_sensitivity(target: Any, transition: Any, spread: Any) -> None:
    """Populate sensitivity tables with transition and spread matrices.

    Args:
        target: Sensitivity worksheet containing destination tables.
        transition: Worksheet containing the transition matrix.
        spread: Worksheet containing the spread matrix.

    Raises:
        OrchestratorError: If a destination matrix table has incompatible
            dimensions or a source matrix is malformed.
    """
    for table in dict.values(target.tables):
        if "Transitions" in table.name or "transition" in table.name.lower():
            _copy_matrix(transition, target, table)
        elif "Spreads" in table.name or "spread" in table.name.lower():
            _copy_matrix(spread, target, table)


def _copy_control(
    workbook: Any,
    output: Any,
    combination: Combination,
    valuation_date: date,
    named_cells: dict[str, str],
) -> None:
    """Copy and populate the output control worksheet.

    Args:
        workbook: Master workbook containing the source control worksheet.
        output: Output workbook receiving the copied worksheet.
        combination: Selected values to write to the control worksheet.
        valuation_date: Valuation date to write to the control worksheet.
        named_cells: Mapping of required input names to cell coordinates.
    """
    control = output.create_sheet(CONTROL_OUTPUT_SHEET)
    _copy_sheet(workbook[CONTROL_OUTPUT_SHEET], control)
    _populate_control(output, control, combination, valuation_date, named_cells)


def _populate_control(
    output: Any,
    control: Any,
    combination: Combination,
    valuation_date: date,
    named_cells: dict[str, str],
) -> None:
    """Write combination-specific values and names to an output control tab.

    Args:
        output: Output workbook whose defined names are updated.
        control: Output control worksheet receiving the values.
        combination: Selected values to write to the worksheet.
        valuation_date: Valuation date to write to the worksheet.
        named_cells: Mapping of required input names to cell coordinates.
    """
    values = {
        "valuation_date": valuation_date,
        "entity_name": combination.entity,
        "risk_factor_path": combination.risk_factor_path,
        "output_folder_path": combination.output_folder_value or str(combination.output_folder),
        "prm_LPIgeneratorWB": None,
        "output_suffix": combination.stress_suffix,
        "bool_inv_exp": combination.bool_inv_exp,
    }
    for name, value in values.items():
        control[named_cells[name]] = value
        output.defined_names.add(
            DefinedName(name, attr_text=f"'{CONTROL_OUTPUT_SHEET}'!{named_cells[name]}")
        )


def copy_combination(
    workbook: Any,
    combination: Combination,
    valuation_date: date,
    output_folder: Path,
    data_sheet_name: str = DEFAULT_DATA_SHEET_NAME,
    sensitivity_sheet_name: str = DEFAULT_SENSITIVITY_SHEET_NAME,
    workbook_bytes: bytes | None = None,
) -> Path:
    """Copy one validated combination to an atomic output workbook.

    Args:
        workbook: Validated master workbook containing source worksheets.
        combination: Entity/category/stress selection to generate.
        valuation_date: Valuation date used in the output control and name.
        output_folder: Destination folder for the generated workbook.
        data_sheet_name: Name assigned to the retained data worksheet.
        sensitivity_sheet_name: Name assigned to the generated sensitivity
            worksheet.
        workbook_bytes: Optional serialised master workbook to reuse when
            generating multiple outputs.

    Returns:
        The path of the newly written workbook.

    Raises:
        FileExistsError: If the deterministic output path already exists.
        OrchestratorError: If required template, control, source, or sheet
            structure is invalid.
    """
    if SENSITIVITY_TEMPLATE_SHEET not in workbook.sheetnames:
        raise OrchestratorError(f"Missing sensitivity template sheet: {SENSITIVITY_TEMPLATE_SHEET}")
    if workbook[SENSITIVITY_TEMPLATE_SHEET].sheet_state not in {"hidden", "veryHidden"}:
        raise OrchestratorError(
            f"Sensitivity template sheet must be hidden: {SENSITIVITY_TEMPLATE_SHEET}"
        )
    named_cells = _control_named_cells(workbook)
    source_names = {
        combination.source_sheet,
        combination.transition_tab,
        combination.spread_tab,
    }
    static_names = static_sheet_names(workbook, source_names | {SENSITIVITY_TEMPLATE_SHEET})
    output_folder.mkdir(parents=True, exist_ok=True)
    destination = output_folder / output_name(combination, valuation_date)
    if destination.exists():
        raise FileExistsError(f"Output already exists: {destination}")
    missing = [name for name in source_names if name not in workbook.sheetnames]
    if missing:
        raise OrchestratorError(f"Missing worksheets for combination: {', '.join(missing)}")
    if data_sheet_name in static_names:
        raise OrchestratorError(f"Data sheet name conflicts with an existing sheet: {data_sheet_name}")
    if sensitivity_sheet_name in {data_sheet_name, *static_names}:
        raise OrchestratorError(
            f"Sensitivity sheet name conflicts with an existing sheet: {sensitivity_sheet_name}"
        )
    if workbook_bytes is None:
        master_buffer = BytesIO()
        workbook.save(master_buffer)
        workbook_bytes = master_buffer.getvalue()
    output = load_workbook(BytesIO(workbook_bytes), data_only=False)
    retained_names = [
        combination.source_sheet,
        CONTROL_OUTPUT_SHEET,
        SENSITIVITY_TEMPLATE_SHEET,
        *static_names,
    ]
    for name in list(output.sheetnames):
        if name not in retained_names:
            del output[name]
    output[combination.source_sheet].title = data_sheet_name
    output[SENSITIVITY_TEMPLATE_SHEET].title = sensitivity_sheet_name
    output[sensitivity_sheet_name].sheet_state = "visible"
    _populate_control(output, output[CONTROL_OUTPUT_SHEET], combination, valuation_date, named_cells)
    _populate_sensitivity(
        output[sensitivity_sheet_name],
        output[combination.transition_tab] if combination.transition_tab in output.sheetnames else workbook[combination.transition_tab],
        output[combination.spread_tab] if combination.spread_tab in output.sheetnames else workbook[combination.spread_tab],
    )
    for name in (combination.transition_tab, combination.spread_tab):
        if name in output.sheetnames:
            del output[name]
    output._sheets = [output[name] for name in [data_sheet_name, CONTROL_OUTPUT_SHEET, sensitivity_sheet_name, *static_names]]
    buffer = BytesIO()
    output.save(buffer)
    destination.write_bytes(buffer.getvalue())
    return destination


def run(
    master_path: Path,
    overwrite: bool = False,
    data_sheet_name: str = DEFAULT_DATA_SHEET_NAME,
    sensitivity_sheet_name: str = DEFAULT_SENSITIVITY_SHEET_NAME,
) -> list[Path]:
    """Validate a master workbook, create outputs, and write run manifests.

    Args:
        master_path: Path to the master workbook.
        overwrite: Whether existing deterministic outputs may be replaced.
        data_sheet_name: Name assigned to each retained data worksheet.
        sensitivity_sheet_name: Name assigned to each generated sensitivity
            worksheet.

    Returns:
        Paths to the manifest written in each distinct stress output folder.

    Raises:
        FileExistsError: If an output already exists and overwrite is disabled.
        OrchestratorError: If the master workbook fails validation.
        OSError: If workbook or manifest file operations fail.
    """
    workbook = load_workbook(master_path, data_only=False)
    master_buffer = BytesIO()
    workbook.save(master_buffer)
    workbook_bytes = master_buffer.getvalue()
    _control_named_cells(workbook)
    valuation_date, combinations = parse_control(workbook, master_path.parent)
    if not overwrite:
        collisions = [
            item.output_folder / output_name(item, valuation_date)
            for item in combinations
            if (item.output_folder / output_name(item, valuation_date)).exists()
        ]
        if collisions:
            raise FileExistsError(f"Output already exists: {collisions[0]}")

    results: list[RunResult] = []
    for combination in combinations:
        destination_folder = combination.output_folder
        try:
            if destination_folder is None:
                raise OrchestratorError("Combination has no output folder")
            destination = destination_folder / output_name(combination, valuation_date)
            if overwrite:
                destination.unlink(missing_ok=True)
            created = copy_combination(
                workbook,
                combination,
                valuation_date,
                destination_folder,
                data_sheet_name=data_sheet_name,
                sensitivity_sheet_name=sensitivity_sheet_name,
                workbook_bytes=workbook_bytes,
            )
            result_combination = asdict(combination)
            result_combination["output_folder"] = str(destination_folder)
            results.append(RunResult(result_combination, str(created), "success"))
        except (FileExistsError, OSError, OrchestratorError) as error:
            result_combination = asdict(combination)
            result_combination["output_folder"] = str(destination_folder) if destination_folder else None
            results.append(RunResult(result_combination, None, "failed", str(error)))

    timestamp = datetime.now()
    manifests: list[Path] = []
    for destination_folder in sorted({item.output_folder for item in combinations if item.output_folder}, key=str):
        destination_folder.mkdir(parents=True, exist_ok=True)
        manifest = destination_folder / f"run_manifest_{timestamp.strftime('%Y%m%dT%H%M%S')}.json"
        folder_results = [
            asdict(result)
            for result in results
            if result.combination["output_folder"] == str(destination_folder)
        ]
        manifest.write_text(
            json.dumps(
                {
                    "master_path": str(master_path),
                    "timestamp": timestamp.isoformat(timespec="seconds"),
                    "valuation_date": valuation_date.isoformat(),
                    "output_folder": str(destination_folder),
                    "results": folder_results,
                },
                indent=2,
            ),
            encoding="utf-8",
        )
        manifests.append(manifest)
    return manifests


def main() -> int:
    """Parse command-line arguments and run the workbook orchestrator.

    Returns:
        Process exit code. Argument and orchestration errors are reported by
        ``argparse``.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--master", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--data-sheet-name", default=DEFAULT_DATA_SHEET_NAME)
    parser.add_argument("--sensitivity-sheet-name", default=DEFAULT_SENSITIVITY_SHEET_NAME)
    args = parser.parse_args()
    try:
        manifest = run(
            args.master,
            args.overwrite,
            args.data_sheet_name,
            args.sensitivity_sheet_name,
        )
    except (FileNotFoundError, OrchestratorError, FileExistsError, OSError) as error:
        parser.error(str(error))
    for manifest_path in manifest:
        print(f"Manifest written to {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())