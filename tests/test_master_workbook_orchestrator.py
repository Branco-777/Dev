import json
from datetime import date
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

from master_workbook_orchestrator import (
    DEFAULT_DATA_SHEET_NAME,
    DEFAULT_SENSITIVITY_SHEET_NAME,
    Combination,
    OrchestratorError,
    output_name,
    parse_control,
    copy_combination,
    run,
    sanitise_filename,
    static_sheet_names,
)
from master_workbook_orchestrator_merged import (
    OrchestratorError as MergedOrchestratorError,
    _copy_output as merged_copy_output,
    run as merged_run,
)
from master_workbook_orchestrator_package import _copy_output, run as package_run


def _add_table(worksheet, name, headers, rows, start_row):
    for column, value in enumerate(headers, 1):
        worksheet.cell(start_row, column, value)
    for row_index, row in enumerate(rows, start_row + 1):
        for column, value in enumerate(row, 1):
            worksheet.cell(row_index, column, value)
    end_row = start_row + len(rows)
    end_column = len(headers)
    table = Table(displayName=name, ref=f"A{start_row}:{chr(64 + end_column)}{end_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    worksheet.add_table(table)


def _workbook():
    workbook = Workbook()
    control = workbook.active
    control.title = "Global Control"
    _add_table(
        control,
        "tbl_entity_mapping",
        ["Entity", "Category", "SourceSheet", "Selected"],
        [
            ["JRL", "Bonds", "JRL", True],
            ["JRL", "MAP", "JRL MAP", True],
            ["PLACL", "Bonds", "PLACL", True],
            ["PLACL", "MAP", "PLACL MAP", True],
        ],
        1,
    )
    _add_table(
        control,
        "tbl_stress_selection",
        ["Stress", "Transition Tab", "Spread Tab", "Suffix", "OutputFolder", "Selected"],
        [
            ["1in20", "Transition Matrix", "Spread Matrix", "1in20Comb", "one", True],
            ["1in50", "Transition Matrix", "Spread Matrix", "1in50Comb", "two", True],
        ],
        6,
    )
    _add_table(
        control,
        "tbl_run_settings",
        ["Setting", "Value"],
        [
            ["ValuationDate", date(2026, 6, 30)],
            ["risk_factor_path", r"\\server015\risk factors"],
            ["bool_inv_exp", True],
        ],
        11,
    )
    for name in ["JRL", "JRL MAP", "PLACL", "PLACL MAP", "Transition Matrix", "Spread Matrix"] + [
        f"Static{i}" for i in range(8)
    ]:
        worksheet = workbook.create_sheet(name)
        worksheet["A1"] = "=1+1" if name == "JRL" else name
    for name, value in [("Transition Matrix", "transition"), ("Spread Matrix", "spread")]:
        worksheet = workbook[name]
        worksheet["A1"] = "Heading"
        worksheet["B1"] = "Value"
        worksheet["A2"] = "Row"
        worksheet["B2"] = value
        worksheet.add_table(Table(displayName=f"tbl_{name.replace(' ', '_')}", ref="A1:B2"))
    template = workbook.create_sheet("I. Sensitivity - Template")
    template.sheet_state = "hidden"
    target_ranges = ["B8:C9", "B19:C20", "B30:C31", "B41:C42"]
    target_ranges += ["L8:M9", "L19:M20", "L30:M31", "L41:M42"]
    target_ranges += ["DK8:DL9", "DK19:DL20", "DK30:DL31", "DK41:DL42"]
    target_names = [
        "I_tbl_Transitions_GBP_FIN",
        "I_tbl_Transitions_GBP_NONFIN",
        "I_tbl_Transitions_USD_FIN",
        "I_tbl_Transitions_USD_NONFIN",
        "I_tbl_Spreads_GBP_FIN",
        "I_tbl_Spreads_GBP_NONFIN",
        "I_tbl_Spreads_USD_FIN",
        "I_tbl_Spreads_USD_NONFIN",
        "I_tbl_Spreads_GBP_FIN_Orig",
        "I_tbl_Spreads_GBP_NONFIN_Orig",
        "I_tbl_Spreads_USD_FIN_Orig",
        "I_tbl_Spreads_USD_NONFIN_Orig",
    ]
    for name, cell_range in zip(target_names, target_ranges):
        template.add_table(Table(displayName=name, ref=cell_range))
    control_sheet = workbook.create_sheet("I. Control")
    control_sheet["D10"] = "template date"
    control_sheet["D13"] = "template entity"
    control_sheet["D16"] = "template risk path"
    control_sheet["D19"] = "template output path"
    control_sheet["D22"] = "template LPI workbook"
    control_sheet["D25"] = "template suffix"
    control_sheet["D28"] = False
    control_cells = {
        "valuation_date": "D10",
        "entity_name": "D13",
        "risk_factor_path": "D16",
        "output_folder_path": "D19",
        "prm_LPIgeneratorWB": "D22",
        "output_suffix": "D25",
        "bool_inv_exp": "D28",
    }
    for name, coordinate in control_cells.items():
        workbook.defined_names.add(
            DefinedName(name, attr_text=f"'I. Control'!${coordinate[0]}${coordinate[1:]}")
        )
    return workbook


def test_parse_control_generates_eight_combinations():
    _, combinations = parse_control(_workbook())
    assert len(combinations) == 8
    assert len({(item.entity, item.category, item.stress) for item in combinations}) == 8
    assert {item.stress_suffix for item in combinations} == {"1in20Comb", "1in50Comb"}


def test_parse_control_does_not_require_global_suffix():
    _, combinations = parse_control(_workbook())
    assert all(item.stress_suffix for item in combinations)


def test_parse_control_assigns_each_stress_output_folder(tmp_path: Path):
    _, combinations = parse_control(_workbook(), tmp_path)
    assert {item.output_folder for item in combinations} == {
        tmp_path / "one",
        tmp_path / "two",
    }


def test_parse_control_rejects_blank_folder_for_selected_stress():
    workbook = _workbook()
    workbook["Global Control"]["E7"] = ""
    with pytest.raises(OrchestratorError, match="OutputFolder"):
        parse_control(workbook)


def test_parse_control_accepts_blank_folder_for_unselected_stress():
    workbook = _workbook()
    workbook["Global Control"]["E8"] = ""
    workbook["Global Control"]["F8"] = False
    _, combinations = parse_control(workbook)
    assert {item.stress for item in combinations} == {"1in20"}


def test_parse_control_rejects_legacy_global_output_folder():
    workbook = _workbook()
    _add_table(
        workbook["Global Control"],
        "tbl_legacy_settings",
        ["Setting", "Value"],
        [["OutputFolder", "legacy"]],
        14,
    )
    with pytest.raises(OrchestratorError, match="global OutputFolder"):
        parse_control(workbook)


def test_parse_control_rejects_mismatched_stress_sheet():
    workbook = _workbook()
    workbook["Global Control"]["B7"] = "WrongSheet"
    with pytest.raises(OrchestratorError, match="Missing Transition Tab"):
        parse_control(workbook)


def test_parse_control_accepts_existing_source_sheet_from_entity_table():
    workbook = _workbook()
    workbook["Global Control"]["C2"] = "Custom Data"
    workbook["JRL"].title = "Custom Data"

    _, combinations = parse_control(workbook)

    assert combinations[0].source_sheet == "Custom Data"


def test_static_sheet_rule_excludes_custom_source_sheet_from_entity_table():
    workbook = _workbook()
    workbook["Global Control"]["C2"] = "Custom Data"
    workbook["JRL"].title = "Custom Data"

    assert static_sheet_names(
        workbook, {"Transition Matrix", "Spread Matrix", "I. Sensitivity - Template"}
    ) == [f"Static{i}" for i in range(8)]


def test_parse_control_rejects_missing_source_sheet_from_entity_table():
    workbook = _workbook()
    workbook["Global Control"]["C2"] = "Missing Data"

    with pytest.raises(OrchestratorError, match="SourceSheet.*Missing Data"):
        parse_control(workbook)


def test_static_sheet_rule_finds_eight_sheets():
    assert static_sheet_names(_workbook(), {"Transition Matrix", "Spread Matrix", "I. Sensitivity - Template"}) == [
        f"Static{i}" for i in range(8)
    ]


def test_static_sheet_rule_ignores_matrix_tabs_for_unselected_stresses():
    workbook = _workbook()
    control = workbook["Global Control"]
    control["A8"] = "New Stress"
    control["B8"] = "NewStress_Transition"
    control["C8"] = "NewStress_Spread"
    workbook.create_sheet("NewStress_Transition")
    workbook.create_sheet("NewStress_Spread")

    assert static_sheet_names(
        workbook, {"Transition Matrix", "Spread Matrix", "I. Sensitivity - Template"}
    ) == [f"Static{i}" for i in range(8)]


def test_sanitise_filename_removes_windows_characters():
    assert sanitise_filename('stress: a/b*?') == "stress_ a_b__"


def test_output_name_is_deterministic():
    combination = Combination("JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix", "1in20Comb")
    assert output_name(combination, date(2026, 6, 30)) == "202606_JRL_Bonds_1in20Comb.xlsx"


def test_output_name_uses_stress_suffix_when_no_override_is_provided():
    combination = Combination(
        "JRL", "Bonds", "JRL", "Credit Combined 1 in 20", "Transition Matrix", "Spread Matrix", "1in20Comb"
    )
    assert output_name(combination, date(2026, 6, 30)) == (
        "202606_JRL_Bonds_1in20Comb.xlsx"
    )


def test_copy_combination_renames_data_sheet_and_filters_others(tmp_path: Path):
    workbook = _workbook()
    combination = Combination("JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix")
    destination = copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert output.sheetnames == [DEFAULT_DATA_SHEET_NAME, "I. Control", DEFAULT_SENSITIVITY_SHEET_NAME] + [
        f"Static{i}" for i in range(8)
    ]
    assert output[DEFAULT_DATA_SHEET_NAME]["A1"].value == "=1+1"
    assert "JRL" not in output.sheetnames
    assert "PLACL" not in output.sheetnames
    assert output["I. Control"]["D13"].value == "JRL"


def test_copy_combination_copies_styles_into_output_style_tables(tmp_path: Path):
    workbook = _workbook()
    workbook["JRL"]["A1"].alignment = Alignment(horizontal="center", text_rotation=45)
    combination = Combination("JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix")

    destination = copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert output[DEFAULT_DATA_SHEET_NAME]["A1"].alignment.horizontal == "center"
    assert output[DEFAULT_DATA_SHEET_NAME]["A1"].alignment.textRotation == 45


def test_copy_combination_writes_valid_table_definitions(tmp_path: Path):
    workbook = _workbook()
    combination = Combination("JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix")

    destination = copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)

    with ZipFile(destination) as package:
        table_xml = [
            package.read(name).decode("utf-8")
            for name in package.namelist()
            if name.startswith("xl/tables/")
        ]
    assert table_xml
    assert all("dataDxfId=" not in xml for xml in table_xml)


def test_copy_combination_handles_deepcopied_sensitivity_tables(tmp_path: Path):
    workbook = _composition_workbook()
    combination = Combination(
        "JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix"
    )

    destination = copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert output[DEFAULT_SENSITIVITY_SHEET_NAME]["C9"].value == "transition"
    assert output[DEFAULT_SENSITIVITY_SHEET_NAME]["M9"].value == "spread"


def test_copy_combination_accepts_custom_sensitivity_sheet_name(tmp_path: Path):
    workbook = _workbook()
    combination = Combination("JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix")
    destination = copy_combination(
        workbook,
        combination,
        date(2026, 6, 30),
        tmp_path,
        sensitivity_sheet_name="Custom Sensitivity",
    )

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert "Custom Sensitivity" in output.sheetnames
    assert "Transition Matrix" not in output.sheetnames


def test_copy_combination_accepts_custom_data_sheet_name(tmp_path: Path):
    workbook = _workbook()
    combination = Combination("PLACL", "Bonds", "PLACL", "1in50", "Transition Matrix", "Spread Matrix")
    destination = copy_combination(
        workbook, combination, date(2026, 6, 30), tmp_path, data_sheet_name="Custom Bonds"
    )

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert "Custom Bonds" in output.sheetnames
    assert "PLACL" not in output.sheetnames


def test_copy_combination_rejects_data_sheet_name_conflict(tmp_path: Path):
    workbook = _workbook()
    combination = Combination("JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix")
    with pytest.raises(OrchestratorError, match="conflicts"):
        copy_combination(
            workbook,
            combination,
            date(2026, 6, 30),
            tmp_path,
            data_sheet_name="Static0",
        )


def test_run_routes_outputs_and_manifests_to_each_stress_folder(tmp_path: Path):
    master_path = tmp_path / "master.xlsx"
    workbook = _workbook()
    workbook.save(master_path)
    manifests = run(master_path)

    assert {path.parent.name for path in manifests} == {"one", "two"}
    assert len(list((tmp_path / "one").glob("*.xlsx"))) == 4
    assert len(list((tmp_path / "two").glob("*.xlsx"))) == 4
    for manifest_path in manifests:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        assert manifest["output_folder"] == str(manifest_path.parent)
        assert len(manifest["results"]) == 4
        assert all(result["status"] == "success" for result in manifest["results"])


def test_copy_combination_composes_transition_and_spread_tabs_from_template(tmp_path: Path):
    workbook = _composition_workbook()
    combination = Combination(
        "JRL",
        "Bonds",
        "JRL",
        "1in20",
        "Transition Matrix",
        "Spread Matrix",
        "1in20Comb",
        tmp_path,
    )
    destination = copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    sensitivity = output[DEFAULT_SENSITIVITY_SHEET_NAME]
    assert sensitivity["C9"].value == "transition"
    assert sensitivity["M9"].value == "spread"
    assert sensitivity["C20"].value == "transition"
    assert sensitivity["M20"].value == "spread"
    assert sensitivity["DL9"].value == "spread"
    assert list(sensitivity.tables) == list(workbook["I. Sensitivity - Template"].tables)


def test_package_copy_makes_sensitivity_sheet_visible(tmp_path: Path):
    workbook = _composition_workbook()
    combination = Combination(
        "JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix", output_folder=tmp_path
    )
    buffer = BytesIO()
    workbook.save(buffer)

    destination = _copy_output(
        buffer.getvalue(),
        workbook,
        combination,
        date(2026, 6, 30),
        tmp_path,
        DEFAULT_DATA_SHEET_NAME,
        DEFAULT_SENSITIVITY_SHEET_NAME,
    )

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert output[DEFAULT_SENSITIVITY_SHEET_NAME].sheet_state == "visible"


def test_package_copy_preserves_defined_names_for_retained_sheets(tmp_path: Path):
    workbook = _composition_workbook()
    workbook.defined_names.add(
        DefinedName("parameter_input", attr_text="'Static0'!$A$1")
    )
    combination = Combination(
        "JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix", output_folder=tmp_path
    )
    buffer = BytesIO()
    workbook.save(buffer)

    destination = _copy_output(
        buffer.getvalue(),
        workbook,
        combination,
        date(2026, 6, 30),
        tmp_path,
        DEFAULT_DATA_SHEET_NAME,
        DEFAULT_SENSITIVITY_SHEET_NAME,
    )

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert output.defined_names["parameter_input"].value == "'Static0'!$A$1"
    assert output.defined_names["valuation_date"].value == "'I. Control'!$D$10"


def test_package_run_records_results_in_each_stress_folder(tmp_path: Path):
    master_path = tmp_path / "master.xlsx"
    workbook = _composition_workbook()
    workbook.save(master_path)

    manifests = package_run(master_path)

    assert {path.parent.name for path in manifests} == {"one", "two"}
    for manifest_path in manifests:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        assert len(manifest["results"]) == 4
        assert all(result["status"] == "success" for result in manifest["results"])


def test_merged_copy_rewrites_retained_formula_references(tmp_path: Path):
    workbook = _workbook()
    workbook["Static0"]["A1"] = "='JRL'!A1"
    combination = Combination(
        "JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix"
    )
    buffer = BytesIO()
    workbook.save(buffer)

    destination = merged_copy_output(
        buffer.getvalue(),
        workbook,
        combination,
        date(2026, 6, 30),
        tmp_path,
        DEFAULT_DATA_SHEET_NAME,
        DEFAULT_SENSITIVITY_SHEET_NAME,
    )

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    assert output["Static0"]["A1"].value == "='I. Bonds Data'!A1"


def test_merged_copy_validates_required_named_inputs(tmp_path: Path):
    workbook = _workbook()
    del workbook.defined_names["entity_name"]
    combination = Combination(
        "JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix"
    )
    buffer = BytesIO()
    workbook.save(buffer)

    with pytest.raises(MergedOrchestratorError, match="entity_name"):
        merged_copy_output(
            buffer.getvalue(),
            workbook,
            combination,
            date(2026, 6, 30),
            tmp_path,
            DEFAULT_DATA_SHEET_NAME,
            DEFAULT_SENSITIVITY_SHEET_NAME,
        )


def test_merged_copy_rejects_output_sheet_conflict(tmp_path: Path):
    workbook = _workbook()
    combination = Combination(
        "JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix"
    )
    buffer = BytesIO()
    workbook.save(buffer)

    with pytest.raises(MergedOrchestratorError, match="conflicts"):
        merged_copy_output(
            buffer.getvalue(),
            workbook,
            combination,
            date(2026, 6, 30),
            tmp_path,
            "Static0",
            DEFAULT_SENSITIVITY_SHEET_NAME,
        )


def test_merged_run_rejects_collisions_before_writing_outputs(tmp_path: Path):
    master_path = tmp_path / "master.xlsx"
    workbook = _workbook()
    workbook.save(master_path)
    _, combinations = parse_control(workbook, tmp_path)
    existing = combinations[0].output_folder / output_name(combinations[0], date(2026, 6, 30))
    existing.parent.mkdir()
    existing.write_bytes(b"existing")

    with pytest.raises(FileExistsError):
        merged_run(master_path)

    assert not list((tmp_path / "two").glob("*.xlsx"))


def test_merged_run_creates_outputs_and_manifests(tmp_path: Path):
    master_path = tmp_path / "master.xlsx"
    workbook = _composition_workbook()
    workbook.save(master_path)

    manifests = merged_run(master_path)

    assert {path.parent.name for path in manifests} == {"one", "two"}
    assert len(list((tmp_path / "one").glob("*.xlsx"))) == 4
    assert len(list((tmp_path / "two").glob("*.xlsx"))) == 4
    for manifest_path in manifests:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        assert len(manifest["results"]) == 4
        assert all(result["status"] == "success" for result in manifest["results"])


def test_parse_control_requires_transition_and_spread_tabs():
    workbook = _composition_workbook()
    workbook["Global Control"]["B7"] = "Missing Transition"
    with pytest.raises(OrchestratorError, match="Transition Tab"):
        parse_control(workbook)


def test_copy_combination_rejects_visible_sensitivity_template(tmp_path: Path):
    workbook = _workbook()
    workbook["I. Sensitivity - Template"].sheet_state = "visible"
    combination = Combination(
        "JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix"
    )
    with pytest.raises(OrchestratorError, match="must be hidden"):
        copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)


def test_copy_combination_populates_i_control_and_defined_names(tmp_path: Path):
    workbook = _workbook()
    combination = Combination(
        "PLACL", "MAP", "PLACL MAP", "1in20", "Transition Matrix", "Spread Matrix",
        "1in20Comb", tmp_path, r"\\server015\risk factors", True,
    )
    destination = copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)

    from openpyxl import load_workbook

    output = load_workbook(destination, data_only=False)
    control = output["I. Control"]
    assert control["D10"].value.date() == date(2026, 6, 30)
    assert control["D13"].value == "PLACL"
    assert control["D16"].value == r"\\server015\risk factors"
    assert control["D19"].value == str(tmp_path)
    assert control["D22"].value is None
    assert control["D25"].value == "1in20Comb"
    assert control["D28"].value is True
    for name in (
        "valuation_date", "entity_name", "risk_factor_path", "output_folder_path",
        "prm_LPIgeneratorWB", "output_suffix", "bool_inv_exp",
    ):
        defined_name = output.defined_names[name]
        assert defined_name.localSheetId is None
        assert list(defined_name.destinations) == [("I. Control", defined_name.value.split("!")[1])]


def test_copy_combination_requires_i_control_and_named_inputs(tmp_path: Path):
    workbook = _workbook()
    del workbook["I. Control"]
    combination = Combination("JRL", "Bonds", "JRL", "1in20", "Transition Matrix", "Spread Matrix")
    with pytest.raises(OrchestratorError, match="I. Control"):
        copy_combination(workbook, combination, date(2026, 6, 30), tmp_path)


def test_parse_control_requires_new_run_settings():
    workbook = _workbook()
    workbook["Global Control"]["A13"] = "other"
    with pytest.raises(OrchestratorError, match="risk_factor_path"):
        parse_control(workbook)


def test_parse_control_accepts_human_readable_run_setting_labels():
    workbook = _workbook()
    workbook["Global Control"]["A12"] = "Valuation Date"
    workbook["Global Control"]["A13"] = "Risk Factor Stresses Path"
    workbook["Global Control"]["A14"] = "Calculate investment expenses"
    parse_control(workbook)


def test_run_populates_i_control_for_each_entity(tmp_path: Path):
    master_path = tmp_path / "master.xlsx"
    workbook = _workbook()
    workbook["Global Control"]["F7"] = True
    workbook["Global Control"]["F8"] = False
    workbook.save(master_path)

    run(master_path)

    from openpyxl import load_workbook

    outputs = sorted((tmp_path / "one").glob("*.xlsx"))
    assert len(outputs) == 4
    for output_path in outputs:
        output = load_workbook(output_path, data_only=False)
        assert output["I. Control"]["D10"].value.date() == date(2026, 6, 30)
        assert output["I. Control"]["D16"].value == r"\\server015\risk factors"
        assert output["I. Control"]["D19"].value == "one"
        assert output["I. Control"]["D25"].value == "1in20Comb"
        assert output["I. Control"]["D22"].value is None
        assert output["I. Control"]["D28"].value is True
    entities = {
        load_workbook(path, data_only=False)["I. Control"]["D13"].value
        for path in outputs
    }
    assert entities == {"JRL", "PLACL"}


def _composition_workbook():
    return _workbook()