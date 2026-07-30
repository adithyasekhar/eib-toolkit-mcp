"""Tests for the core data models."""

import pytest

from eib_toolkit.model import (
    CellRef,
    ColumnSpec,
    ColumnType,
    SheetData,
    SheetRole,
    SheetSpec,
    TemplateSpec,
    column_letter,
)


class TestColumnLetter:
    def test_single_letters(self) -> None:
        assert column_letter(1) == "A"
        assert column_letter(26) == "Z"

    def test_multi_letters(self) -> None:
        assert column_letter(27) == "AA"
        assert column_letter(52) == "AZ"
        assert column_letter(703) == "AAA"

    def test_rejects_non_positive(self) -> None:
        with pytest.raises(ValueError, match="column index"):
            column_letter(0)


class TestCellRef:
    def test_a1_and_str(self) -> None:
        ref = CellRef(sheet="Load Data", row=7, column=28)
        assert ref.a1 == "AB7"
        assert str(ref) == "Load Data!AB7"


def _sheet_spec() -> SheetSpec:
    return SheetSpec(
        name="Data",
        header_row=2,
        data_start_row=4,
        columns=[
            ColumnSpec(index=1, header="Spreadsheet Key", is_key=True),
            ColumnSpec(index=2, header="Legal Name", col_type=ColumnType.TEXT),
        ],
    )


class TestSheetSpec:
    def test_key_column(self) -> None:
        spec = _sheet_spec()
        assert spec.key_column is not None
        assert spec.key_column.index == 1
        assert SheetSpec(name="X").key_column is None

    def test_column_by_header_is_case_insensitive(self) -> None:
        spec = _sheet_spec()
        col = spec.column_by_header("  legal name ")
        assert col is not None and col.index == 2
        assert spec.column_by_header("No Such") is None


class TestTemplateSpec:
    def test_sheet_selection_helpers(self) -> None:
        spec = TemplateSpec(
            sheets=[
                SheetSpec(name="Instructions", role=SheetRole.INSTRUCTIONS),
                SheetSpec(name="Main", role=SheetRole.PRIMARY),
                SheetSpec(name="Children", role=SheetRole.REPEATING),
            ]
        )
        assert [s.name for s in spec.data_sheets()] == ["Main", "Children"]
        primary = spec.primary_sheet()
        assert primary is not None and primary.name == "Main"
        assert spec.sheet("Children") is not None
        assert spec.sheet("Nope") is None

    def test_to_dict_round_trips_enums_as_values(self) -> None:
        spec = TemplateSpec(sheets=[SheetSpec(name="Main", role=SheetRole.PRIMARY)])
        as_dict = spec.to_dict()
        assert as_dict["sheets"][0]["name"] == "Main"


class TestSheetData:
    def test_cell_addressing_maps_back_to_excel_rows(self) -> None:
        data = SheetData(spec=_sheet_spec(), rows=[[1, "Avery"], [2, "Blake"]])
        ref = data.cell_ref(1, 1)
        assert (ref.sheet, ref.row, ref.column) == ("Data", 5, 2)
        assert data.value(1, 1) == "Blake"

    def test_short_rows_pad_with_none(self) -> None:
        data = SheetData(spec=_sheet_spec(), rows=[[1, "Avery"], [2]])
        assert data.value(1, 1) is None
        assert data.column_values(1) == ["Avery", None]
