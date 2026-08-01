"""Tests for template drift and filled-workbook content diffing.

Template pairs mimic what a Workday release actually does to generated EIB
templates: add a column, drop a column, rename one in place, flip
requiredness, retype, change a reference-ID type. Workbook pairs exercise
key-based row matching (reordering is not a change) and normalization
(1 vs 1.0 vs "1" is not a change). All data is synthetic.
"""

import datetime as dt
from pathlib import Path
from typing import Any

from openpyxl import Workbook as XlsxWorkbook

from eib_toolkit.diff import ChangeKind, diff_templates, diff_workbooks
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


def _kinds(changes: list[Any]) -> list[ChangeKind]:
    return [c.kind for c in changes]


_OLD_HEADER = ["Spreadsheet Key*", "Employee_Reference", "Legal Name", "Grade"]
_OLD_HINTS = [
    "Numeric. Required.",
    "Reference ID Type: Employee_ID",
    "Text. Required.",
    "Text. Optional.",
]


class TestTemplateDiff:
    def test_identical_templates_diff_empty(self, tmp_path: Path) -> None:
        sheets = {"Workers": [_OLD_HEADER, _OLD_HINTS]}
        old = parse_template(_write(tmp_path / "a.xlsx", sheets))
        new = parse_template(_write(tmp_path / "b.xlsx", sheets))
        diff = diff_templates(old, new)
        assert diff.changes == []
        assert not diff.breaking

    def test_added_column_is_not_breaking(self, tmp_path: Path) -> None:
        old = parse_template(
            _write(tmp_path / "a.xlsx", {"Workers": [_OLD_HEADER, _OLD_HINTS]})
        )
        new = parse_template(
            _write(
                tmp_path / "b.xlsx",
                {
                    "Workers": [
                        [*_OLD_HEADER, "Cost Center"],
                        [*_OLD_HINTS, "Text. Optional."],
                    ]
                },
            )
        )
        diff = diff_templates(old, new)
        assert _kinds(diff.changes) == [ChangeKind.COLUMN_ADDED]
        assert diff.changes[0].column == "Cost Center"
        assert not diff.breaking

    def test_removed_column_is_breaking(self, tmp_path: Path) -> None:
        old = parse_template(
            _write(tmp_path / "a.xlsx", {"Workers": [_OLD_HEADER, _OLD_HINTS]})
        )
        new = parse_template(
            _write(
                tmp_path / "b.xlsx",
                {"Workers": [_OLD_HEADER[:3], _OLD_HINTS[:3]]},
            )
        )
        diff = diff_templates(old, new)
        assert _kinds(diff.changes) == [ChangeKind.COLUMN_REMOVED]
        assert diff.changes[0].old == "Grade"
        assert diff.breaking

    def test_rename_in_place_reported_as_heuristic_pair(self, tmp_path: Path) -> None:
        old = parse_template(
            _write(tmp_path / "a.xlsx", {"Workers": [_OLD_HEADER, _OLD_HINTS]})
        )
        renamed = [*_OLD_HEADER[:3], "Job Grade"]
        new = parse_template(
            _write(tmp_path / "b.xlsx", {"Workers": [renamed, _OLD_HINTS]})
        )
        diff = diff_templates(old, new)
        assert _kinds(diff.changes) == [ChangeKind.COLUMN_RENAMED]
        change = diff.changes[0]
        assert (change.old, change.new) == ("Grade", "Job Grade")
        assert "verify" in change.detail
        assert diff.breaking

    def test_requiredness_tightening_is_breaking_loosening_is_not(
        self, tmp_path: Path
    ) -> None:
        old = parse_template(
            _write(tmp_path / "a.xlsx", {"Workers": [_OLD_HEADER, _OLD_HINTS]})
        )
        hints = [*_OLD_HINTS[:3], "Text. Required."]  # Grade optional -> required
        new = parse_template(_write(tmp_path / "b.xlsx", {"Workers": [_OLD_HEADER, hints]}))
        tightened = diff_templates(old, new)
        assert _kinds(tightened.changes) == [ChangeKind.REQUIRED_CHANGED]
        assert tightened.changes[0].new == "required"
        assert tightened.breaking

        loosened = diff_templates(new, old)
        assert _kinds(loosened.changes) == [ChangeKind.REQUIRED_CHANGED]
        assert loosened.changes[0].new == "optional"
        assert not loosened.breaking

    def test_type_and_ref_type_changes(self, tmp_path: Path) -> None:
        old = parse_template(
            _write(tmp_path / "a.xlsx", {"Workers": [_OLD_HEADER, _OLD_HINTS]})
        )
        hints = [
            "Numeric. Required.",
            "Reference ID Type: Contingent_Worker_ID",  # ref type changed
            "Text. Required.",
            "Numeric. Optional.",  # Grade: text -> numeric
        ]
        new = parse_template(_write(tmp_path / "b.xlsx", {"Workers": [_OLD_HEADER, hints]}))
        diff = diff_templates(old, new)
        assert sorted(_kinds(diff.changes), key=lambda k: k.value) == [
            ChangeKind.REF_TYPE_CHANGED,
            ChangeKind.TYPE_CHANGED,
        ]
        ref = next(c for c in diff.changes if c.kind is ChangeKind.REF_TYPE_CHANGED)
        assert (ref.old, ref.new) == ("Employee_ID", "Contingent_Worker_ID")
        assert diff.breaking  # the type change breaks; ref change alone would not

    def test_sheet_added_and_removed(self, tmp_path: Path) -> None:
        old = parse_template(
            _write(
                tmp_path / "a.xlsx",
                {
                    "Workers": [_OLD_HEADER, _OLD_HINTS],
                    "Allowances": [["Spreadsheet Key*", "Plan"], ["Numeric.", "Text."]],
                },
            )
        )
        new = parse_template(
            _write(
                tmp_path / "b.xlsx",
                {
                    "Workers": [_OLD_HEADER, _OLD_HINTS],
                    "Deductions": [["Spreadsheet Key*", "Plan"], ["Numeric.", "Text."]],
                },
            )
        )
        diff = diff_templates(old, new)
        kinds = _kinds(diff.changes)
        assert ChangeKind.SHEET_REMOVED in kinds
        assert ChangeKind.SHEET_ADDED in kinds
        assert diff.breaking

    def test_to_dict_shape(self, tmp_path: Path) -> None:
        old = parse_template(
            _write(tmp_path / "a.xlsx", {"Workers": [_OLD_HEADER, _OLD_HINTS]})
        )
        new = parse_template(
            _write(tmp_path / "b.xlsx", {"Workers": [_OLD_HEADER[:3], _OLD_HINTS[:3]]})
        )
        d = diff_templates(old, new).to_dict()
        assert d["breaking"] is True
        assert d["summary"] == {"column_removed": 1}
        assert d["changes"][0]["kind"] == "column_removed"


_WB_HEADER = ["Spreadsheet Key", "Name", "Amount", "Start Date"]


def _load(path: Path, rows: list[list[Any]]) -> Any:
    return parse_workbook(_write(path, {"Data": [_WB_HEADER, *rows]}))


class TestWorkbookDiff:
    def test_identical_content_diff_empty(self, tmp_path: Path) -> None:
        rows = [[1, "Avery Example", 100, dt.date(2026, 1, 15)]]
        diff = diff_workbooks(
            _load(tmp_path / "a.xlsx", rows), _load(tmp_path / "b.xlsx", rows)
        )
        assert diff.changes == []
        assert diff.schema_drift is None

    def test_reordered_rows_are_not_changes(self, tmp_path: Path) -> None:
        old = _load(
            tmp_path / "a.xlsx",
            [[1, "Avery Example", 100, dt.date(2026, 1, 15)],
             [2, "Blake Sample", 200, dt.date(2026, 2, 1)]],
        )
        new = _load(
            tmp_path / "b.xlsx",
            [[2, "Blake Sample", 200, dt.date(2026, 2, 1)],
             [1, "Avery Example", 100, dt.date(2026, 1, 15)]],
        )
        assert diff_workbooks(old, new).changes == []

    def test_normalized_values_are_not_changes(self, tmp_path: Path) -> None:
        old = _load(tmp_path / "a.xlsx", [[1, "Avery Example", 100, "2026-01-15"]])
        # 1.0 vs 1 key, "100" text vs 100, real date vs ISO text: all equal.
        new = _load(tmp_path / "b.xlsx", [[1.0, "Avery Example", "100", dt.date(2026, 1, 15)]])
        assert diff_workbooks(old, new).changes == []

    def test_cell_change_is_addressed_and_keyed(self, tmp_path: Path) -> None:
        old = _load(tmp_path / "a.xlsx", [[1, "Avery Example", 100, dt.date(2026, 1, 15)]])
        new = _load(tmp_path / "b.xlsx", [[1, "Avery Example", 150, dt.date(2026, 1, 15)]])
        diff = diff_workbooks(old, new)
        assert _kinds(diff.changes) == [ChangeKind.CELL_CHANGED]
        change = diff.changes[0]
        assert change.key == "1"
        assert change.column == "Amount"
        assert (change.old, change.new) == ("100", "150")
        assert "Data!C2" in change.detail  # header row 1, data row 2, column C

    def test_row_added_and_removed_by_key(self, tmp_path: Path) -> None:
        old = _load(
            tmp_path / "a.xlsx",
            [[1, "Avery Example", 100, dt.date(2026, 1, 15)],
             [2, "Blake Sample", 200, dt.date(2026, 2, 1)]],
        )
        new = _load(
            tmp_path / "b.xlsx",
            [[1, "Avery Example", 100, dt.date(2026, 1, 15)],
             [3, "Casey Fixture", 300, dt.date(2026, 3, 1)]],
        )
        diff = diff_workbooks(old, new)
        kinds = _kinds(diff.changes)
        assert kinds.count(ChangeKind.ROW_REMOVED) == 1
        assert kinds.count(ChangeKind.ROW_ADDED) == 1
        removed = next(c for c in diff.changes if c.kind is ChangeKind.ROW_REMOVED)
        added = next(c for c in diff.changes if c.kind is ChangeKind.ROW_ADDED)
        assert removed.key == "2"
        assert added.key == "3"

    def test_keyless_sheet_falls_back_to_position(self, tmp_path: Path) -> None:
        header = ["Name", "Amount"]
        old = parse_workbook(
            _write(tmp_path / "a.xlsx", {"Data": [header, ["Avery Example", 100]]})
        )
        new = parse_workbook(
            _write(tmp_path / "b.xlsx", {"Data": [header, ["Avery Example", 150]]})
        )
        diff = diff_workbooks(old, new)
        assert _kinds(diff.changes) == [ChangeKind.CELL_CHANGED]
        assert diff.changes[0].key == ""  # positional match: no stable key to report
        assert "at position 0" in diff.changes[0].detail

    def test_schema_drift_attached_and_shared_columns_still_compared(
        self, tmp_path: Path
    ) -> None:
        old = parse_workbook(
            _write(
                tmp_path / "a.xlsx",
                {"Data": [["Spreadsheet Key", "Name", "Grade"], [1, "Avery Example", "A"]]},
            )
        )
        new = parse_workbook(
            _write(
                tmp_path / "b.xlsx",
                {"Data": [["Spreadsheet Key", "Name"], [1, "Avery Renamed"]]},
            )
        )
        diff = diff_workbooks(old, new)
        assert diff.schema_drift is not None
        assert ChangeKind.COLUMN_REMOVED in _kinds(diff.schema_drift.changes)
        cell = next(c for c in diff.changes if c.kind is ChangeKind.CELL_CHANGED)
        assert cell.column == "Name"
        assert (cell.old, cell.new) == ("Avery Example", "Avery Renamed")

    def test_changes_sorted_by_sheet_key_column(self, tmp_path: Path) -> None:
        old = _load(
            tmp_path / "a.xlsx",
            [[2, "Blake Sample", 200, dt.date(2026, 2, 1)],
             [10, "Jordan Ten", 1000, dt.date(2026, 3, 1)],
             [1, "Avery Example", 100, dt.date(2026, 1, 15)]],
        )
        new = _load(
            tmp_path / "b.xlsx",
            [[2, "Blake Changed", 200, dt.date(2026, 2, 1)],
             [10, "Jordan Changed", 1000, dt.date(2026, 3, 1)],
             [1, "Avery Changed", 100, dt.date(2026, 1, 15)]],
        )
        diff = diff_workbooks(old, new)
        # Numeric key order 1 < 2 < 10, not text order "1" < "10" < "2".
        assert [c.key for c in diff.changes] == ["1", "2", "10"]

    def test_to_dict_shape(self, tmp_path: Path) -> None:
        old = _load(tmp_path / "a.xlsx", [[1, "Avery Example", 100, dt.date(2026, 1, 15)]])
        new = _load(tmp_path / "b.xlsx", [[1, "Avery Example", 150, dt.date(2026, 1, 15)]])
        d = diff_workbooks(old, new).to_dict()
        assert d["summary"] == {"cell_changed": 1}
        assert d["schema_drift"] is None
        assert d["changes"][0]["old"] == "100"
