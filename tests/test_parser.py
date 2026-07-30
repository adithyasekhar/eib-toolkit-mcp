"""Tests for the tolerant workbook parser.

Workbooks are hand-built with openpyxl to mirror the shapes Workday's EIB
templates actually take: an instructions sheet, a multi-row header band
(group row / field row / hint row), repeating-group sheets joined by a
spreadsheet key, blank padding, and filled loads with no hint rows at all.
"""

import datetime as dt
from pathlib import Path
from typing import Any

from openpyxl import Workbook as XlsxWorkbook

from eib_toolkit.model import ColumnType, SheetRole
from eib_toolkit.parser import parse_template, parse_workbook


def _write(path: Path, sheets: dict[str, list[list[Any]]]) -> Path:
    wb = XlsxWorkbook()
    assert wb.active is not None
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return path


def _template_sheets() -> dict[str, list[list[Any]]]:
    """A realistic two-data-sheet template plus an instructions sheet."""
    instructions = [
        [
            "This synthetic workbook mimics a Workday EIB inbound template. "
            "Fill the data sheets below and load via Enterprise Interface Builder. "
            "All names and IDs are fake."
        ],
    ]
    main = [
        [None, None, "Personal Data", None, "Compensation", None],
        [
            "Spreadsheet Key*",
            "Employee_Reference",
            "Legal Name",
            "Hire Date",
            "Annual Amount",
            "Active",
        ],
        [
            "Numeric",
            "Reference ID Type: Employee_ID",
            "Text. Required.",
            "Date (yyyy-mm-dd)",
            "Numeric. Optional.",
            "Boolean (Y/N)",
        ],
        [1, "EMP-001", "Avery Example", dt.date(2026, 1, 15), 55000, "Y"],
        [2, "EMP-002", "Blake Sample", dt.date(2026, 2, 1), 61000, "N"],
    ]
    plans = [
        ["Spreadsheet Key", "Plan Name", "Coverage Begin Date"],
        ["Numeric", "Text", "Date (yyyy-mm-dd)"],
        [1, "Medical - Synthetic PPO", dt.date(2026, 1, 15)],
        [1, "Dental - Synthetic", dt.date(2026, 1, 15)],
        [2, "Medical - Synthetic PPO", dt.date(2026, 2, 1)],
    ]
    return {"Instructions": instructions, "Submit Data": main, "Plan Assignments": plans}


class TestTemplateParsing:
    def test_sheet_roles(self, tmp_path: Path) -> None:
        spec = parse_template(_write(tmp_path / "t.xlsx", _template_sheets()))
        roles = {s.name: s.role for s in spec.sheets}
        assert roles == {
            "Instructions": SheetRole.INSTRUCTIONS,
            "Submit Data": SheetRole.PRIMARY,
            "Plan Assignments": SheetRole.REPEATING,
        }

    def test_header_band_geometry(self, tmp_path: Path) -> None:
        spec = parse_template(_write(tmp_path / "t.xlsx", _template_sheets()))
        main = spec.sheet("Submit Data")
        assert main is not None
        assert main.header_row == 2  # group row above, hint row below
        assert main.data_start_row == 4

    def test_column_classification_from_hints(self, tmp_path: Path) -> None:
        spec = parse_template(_write(tmp_path / "t.xlsx", _template_sheets()))
        main = spec.sheet("Submit Data")
        assert main is not None
        types = {c.header: c.col_type for c in main.columns}
        assert types == {
            "Spreadsheet Key": ColumnType.NUMERIC,
            "Employee_Reference": ColumnType.REFERENCE,
            "Legal Name": ColumnType.TEXT,
            "Hire Date": ColumnType.DATE,
            "Annual Amount": ColumnType.NUMERIC,
            "Active": ColumnType.BOOLEAN,
        }
        assert all(not c.type_inferred for c in main.columns)

    def test_key_required_and_reference_details(self, tmp_path: Path) -> None:
        spec = parse_template(_write(tmp_path / "t.xlsx", _template_sheets()))
        main = spec.sheet("Submit Data")
        assert main is not None

        key = main.key_column
        assert key is not None
        assert key.header == "Spreadsheet Key"  # trailing asterisk stripped
        assert key.required  # ...but remembered as a required marker

        ref = main.column_by_header("Employee_Reference")
        assert ref is not None and ref.ref_id_type == "Employee_ID"

        name = main.column_by_header("Legal Name")
        assert name is not None and name.required
        amount = main.column_by_header("Annual Amount")
        assert amount is not None and not amount.required  # "Optional." hint

    def test_group_row_forward_fills_merged_cells(self, tmp_path: Path) -> None:
        spec = parse_template(_write(tmp_path / "t.xlsx", _template_sheets()))
        main = spec.sheet("Submit Data")
        assert main is not None
        groups = [c.group for c in main.columns]
        assert groups == ["", "", "Personal Data", "Personal Data", "Compensation", "Compensation"]


class TestFilledWorkbookParsing:
    def test_data_rows_and_cell_addressing(self, tmp_path: Path) -> None:
        wb = parse_workbook(_write(tmp_path / "t.xlsx", _template_sheets()))
        main = wb.sheet_data("Submit Data")
        assert main is not None
        assert len(main.rows) == 2
        assert main.value(0, 2) == "Avery Example"
        assert str(main.cell_ref(0, 2)) == "Submit Data!C4"
        # openpyxl reads date cells back as naive datetimes; the day is intact.
        assert main.value(1, 3).date() == dt.date(2026, 2, 1)

        plans = wb.sheet_data("Plan Assignments")
        assert plans is not None
        assert plans.column_values(0) == [1, 1, 2]
        assert plans.spec.data_start_row == 3  # two-row band: header + hints

    def test_instruction_sheets_carry_no_data(self, tmp_path: Path) -> None:
        wb = parse_workbook(_write(tmp_path / "t.xlsx", _template_sheets()))
        assert wb.sheet_data("Instructions") is None
        assert set(wb.to_dict()["row_counts"]) == {"Submit Data", "Plan Assignments"}


class TestTolerance:
    def test_blank_padding_rows_and_columns(self, tmp_path: Path) -> None:
        sheets = {
            "Data": [
                [None, None, None],
                ["Spreadsheet Key", "Legal Name", None, None],
                ["Numeric", "Text. Required.", None],
                [None, None, None],  # padding between band and data
                [1, "Avery Example", None],
                [None, None, None],  # interior blank row is kept, addressable
                [2, "Blake Sample", None],
                [None, None, None],  # trailing padding is trimmed
                [None, None, None],
            ]
        }
        wb = parse_workbook(_write(tmp_path / "t.xlsx", sheets))
        sheet = wb.sheet_data("Data")
        assert sheet is not None
        assert sheet.spec.header_row == 2
        assert sheet.spec.data_start_row == 5
        assert len(sheet.spec.columns) == 2  # trailing blank column trimmed
        assert len(sheet.rows) == 3  # interior blank kept, trailing trimmed
        assert str(sheet.cell_ref(2, 1)) == "Data!B7"

    def test_types_inferred_from_data_when_no_hint_rows(self, tmp_path: Path) -> None:
        sheets = {
            "Data": [
                # Headers chosen so no declared-type token leaks in from the
                # header text itself — classification must come from the data.
                ["Spreadsheet Key", "Name", "Start", "Salary", "Active"],
                [1, "Avery Example", "2026-01-15", "55,000.00", "Y"],
                [2, "Blake Sample", "2026-02-01", "61000", "N"],
            ]
        }
        spec = parse_template(_write(tmp_path / "t.xlsx", sheets))
        sheet = spec.sheet("Data")
        assert sheet is not None
        assert sheet.data_start_row == 2  # no hint rows consumed
        types = {c.header: (c.col_type, c.type_inferred) for c in sheet.columns}
        assert types == {
            "Spreadsheet Key": (ColumnType.NUMERIC, True),
            "Name": (ColumnType.TEXT, True),
            "Start": (ColumnType.DATE, True),
            "Salary": (ColumnType.NUMERIC, True),
            "Active": (ColumnType.BOOLEAN, True),
        }

    def test_all_text_data_rows_do_not_steal_the_header(self, tmp_path: Path) -> None:
        sheets = {
            "Data": [
                ["Code", "Name"],
                ["A1", "Avery Example"],
                ["B2", "Blake Sample"],
            ]
        }
        spec = parse_template(_write(tmp_path / "t.xlsx", sheets))
        sheet = spec.sheet("Data")
        assert sheet is not None
        assert sheet.header_row == 1  # earliest row wins the tie on distinct strings
        assert [c.header for c in sheet.columns] == ["Code", "Name"]

    def test_prose_sheet_without_instruction_name_is_not_data(self, tmp_path: Path) -> None:
        sheets = {
            "Sheet3": [
                ["Loading notes: run validation before submitting anything."],
                ["Contact the integrations team with questions."],
            ],
            "Data": [
                ["Spreadsheet Key", "Name"],
                [1, "Avery Example"],
            ],
        }
        spec = parse_template(_write(tmp_path / "t.xlsx", sheets))
        prose = spec.sheet("Sheet3")
        assert prose is not None and prose.role == SheetRole.INSTRUCTIONS
        assert spec.primary_sheet() is not None
        assert spec.primary_sheet().name == "Data"  # type: ignore[union-attr]

    def test_empty_sheet_is_flagged_empty(self, tmp_path: Path) -> None:
        sheets = {"Blank": [], "Data": [["Spreadsheet Key", "Name"], [1, "Avery Example"]]}
        spec = parse_template(_write(tmp_path / "t.xlsx", sheets))
        blank = spec.sheet("Blank")
        assert blank is not None and blank.role == SheetRole.EMPTY

    def test_second_sheet_without_key_stays_primary(self, tmp_path: Path) -> None:
        sheets = {
            "Workers": [["Spreadsheet Key", "Name"], [1, "Avery Example"]],
            "Lookup": [["Code", "Meaning"], ["A", "Alpha"]],
        }
        spec = parse_template(_write(tmp_path / "t.xlsx", sheets))
        lookup = spec.sheet("Lookup")
        assert lookup is not None and lookup.role == SheetRole.PRIMARY
