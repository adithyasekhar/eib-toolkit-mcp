"""Diff EIB templates (schema drift) and filled loads (content changes).

Two distinct questions, one evidence-backed answer shape:

* :func:`diff_templates` — "what changed between the template my load spec
  was built against and the one the tenant generates today?" Workday's
  twice-yearly releases quietly add, remove, rename, retype, and re-require
  template columns; this is the drift report to run before reusing last
  cycle's spreadsheets. Renames are reported as a heuristic pairing (same
  position, one removed + one added), flagged as such rather than asserted.

* :func:`diff_workbooks` — "what actually differs between these two filled
  loads?" Rows are matched by spreadsheet key when the sheet has one (so
  reordering isn't noise) and by position otherwise; cell values are
  compared after light normalization (dates to ISO text, int-valued floats
  to ints, text stripped) so `1` vs `1.0` vs `"1"` isn't a change.

Every change is a :class:`Change` with machine-usable fields and a
human-readable detail line; reports serialize with ``to_dict`` and order
deterministically.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any

from eib_toolkit.model import SheetSpec, TemplateSpec, Workbook

__all__ = [
    "Change",
    "ChangeKind",
    "TemplateDiff",
    "WorkbookDiff",
    "diff_templates",
    "diff_workbooks",
]


class ChangeKind(str, Enum):
    """Every kind of change either differ can report."""

    # Template drift
    SHEET_ADDED = "sheet_added"
    SHEET_REMOVED = "sheet_removed"
    COLUMN_ADDED = "column_added"
    COLUMN_REMOVED = "column_removed"
    COLUMN_RENAMED = "column_renamed"  # heuristic: same position, old out / new in
    TYPE_CHANGED = "type_changed"
    REQUIRED_CHANGED = "required_changed"
    REF_TYPE_CHANGED = "ref_type_changed"
    KEY_CHANGED = "key_changed"
    # Workbook content
    ROW_ADDED = "row_added"
    ROW_REMOVED = "row_removed"
    CELL_CHANGED = "cell_changed"


#: Drift kinds that break an existing load spec / filled workbook outright.
BREAKING_KINDS = frozenset(
    {
        ChangeKind.SHEET_REMOVED,
        ChangeKind.COLUMN_REMOVED,
        ChangeKind.COLUMN_RENAMED,
        ChangeKind.TYPE_CHANGED,
        ChangeKind.KEY_CHANGED,
    }
)


@dataclass
class Change:
    """One difference. ``key`` is the row's spreadsheet key ("" for schema
    changes and positional rows); ``old``/``new`` are display strings."""

    kind: ChangeKind
    sheet: str
    column: str = ""
    key: str = ""
    old: str = ""
    new: str = ""
    detail: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class TemplateDiff:
    """Schema drift between two templates."""

    old_source: str
    new_source: str
    changes: list[Change] = field(default_factory=list)

    @property
    def breaking(self) -> bool:
        """True when an existing load spec would need rework, not just review.

        A newly-required column also breaks loads built before it existed, so
        ``required_changed`` counts as breaking only in that direction.
        """
        return any(
            c.kind in BREAKING_KINDS
            or (c.kind is ChangeKind.REQUIRED_CHANGED and c.new == "required")
            for c in self.changes
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "old": self.old_source,
            "new": self.new_source,
            "breaking": self.breaking,
            "summary": _summarize(self.changes),
            "changes": [c.to_dict() for c in self.changes],
        }


@dataclass
class WorkbookDiff:
    """Content differences between two filled loads of the same template."""

    old_source: str
    new_source: str
    changes: list[Change] = field(default_factory=list)
    schema_drift: TemplateDiff | None = None  # set when the workbooks' schemas differ

    def to_dict(self) -> dict[str, Any]:
        return {
            "old": self.old_source,
            "new": self.new_source,
            "summary": _summarize(self.changes),
            "schema_drift": self.schema_drift.to_dict() if self.schema_drift else None,
            "changes": [c.to_dict() for c in self.changes],
        }


def _summarize(changes: list[Change]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for c in sorted(changes, key=lambda c: c.kind.value):
        counts[c.kind.value] = counts.get(c.kind.value, 0) + 1
    return counts


# ---------------------------------------------------------------------------
# Template drift


def diff_templates(old: TemplateSpec, new: TemplateSpec) -> TemplateDiff:
    """Report schema drift from ``old`` to ``new`` (sheets matched by name)."""
    diff = TemplateDiff(old_source=old.source, new_source=new.source)

    old_sheets = {s.name: s for s in old.data_sheets()}
    new_sheets = {s.name: s for s in new.data_sheets()}

    for name in old_sheets:
        if name not in new_sheets:
            diff.changes.append(
                Change(
                    kind=ChangeKind.SHEET_REMOVED,
                    sheet=name,
                    detail=f"Data sheet {name!r} no longer exists",
                )
            )
    for name in new_sheets:
        if name not in old_sheets:
            diff.changes.append(
                Change(
                    kind=ChangeKind.SHEET_ADDED,
                    sheet=name,
                    detail=f"New data sheet {name!r} "
                    f"({len(new_sheets[name].columns)} columns)",
                )
            )
    for name in old_sheets:
        if name in new_sheets:
            _diff_sheet(diff, old_sheets[name], new_sheets[name])

    diff.changes.sort(key=lambda c: (c.sheet, c.column, c.kind.value))
    return diff


def _diff_sheet(diff: TemplateDiff, old: SheetSpec, new: SheetSpec) -> None:
    def header_map(sheet: SheetSpec) -> dict[str, Any]:
        return {c.header.casefold(): c for c in sheet.columns if c.header}

    old_cols, new_cols = header_map(old), header_map(new)
    removed = [c for h, c in old_cols.items() if h not in new_cols]
    added = [c for h, c in new_cols.items() if h not in old_cols]

    # Rename heuristic: a removed and an added column at the same position is
    # far more likely a rename than a coincidence; report it as one paired
    # change, explicitly marked heuristic, instead of an add + a remove.
    renamed: list[tuple[Any, Any]] = []
    for gone in list(removed):
        match = next((c for c in added if c.index == gone.index), None)
        if match is not None:
            renamed.append((gone, match))
            removed.remove(gone)
            added.remove(match)

    for gone in removed:
        diff.changes.append(
            Change(
                kind=ChangeKind.COLUMN_REMOVED,
                sheet=old.name,
                column=gone.header,
                old=gone.header,
                detail=f"Column {gone.header!r} (was {gone.letter}) was removed",
            )
        )
    for came in added:
        required = "required" if came.required else "optional"
        diff.changes.append(
            Change(
                kind=ChangeKind.COLUMN_ADDED,
                sheet=new.name,
                column=came.header,
                new=came.header,
                detail=f"Column {came.header!r} ({came.letter}, {came.col_type.value}, "
                f"{required}) was added",
            )
        )
    for gone, came in renamed:
        diff.changes.append(
            Change(
                kind=ChangeKind.COLUMN_RENAMED,
                sheet=new.name,
                column=came.header,
                old=gone.header,
                new=came.header,
                detail=f"Column {came.letter}: {gone.header!r} -> {came.header!r} "
                "(rename inferred from matching position; verify)",
            )
        )
        _diff_column(diff, new.name, gone, came)

    for header, old_col in old_cols.items():
        new_col = new_cols.get(header)
        if new_col is not None:
            _diff_column(diff, new.name, old_col, new_col)


def _diff_column(diff: TemplateDiff, sheet: str, old_col: Any, new_col: Any) -> None:
    name = new_col.header
    if old_col.col_type is not new_col.col_type:
        diff.changes.append(
            Change(
                kind=ChangeKind.TYPE_CHANGED,
                sheet=sheet,
                column=name,
                old=old_col.col_type.value,
                new=new_col.col_type.value,
                detail=f"Column {name!r} changed type "
                f"{old_col.col_type.value} -> {new_col.col_type.value}",
            )
        )
    if old_col.required != new_col.required:
        old_s = "required" if old_col.required else "optional"
        new_s = "required" if new_col.required else "optional"
        diff.changes.append(
            Change(
                kind=ChangeKind.REQUIRED_CHANGED,
                sheet=sheet,
                column=name,
                old=old_s,
                new=new_s,
                detail=f"Column {name!r} is now {new_s} (was {old_s})",
            )
        )
    if old_col.ref_id_type != new_col.ref_id_type:
        diff.changes.append(
            Change(
                kind=ChangeKind.REF_TYPE_CHANGED,
                sheet=sheet,
                column=name,
                old=old_col.ref_id_type,
                new=new_col.ref_id_type,
                detail=f"Column {name!r} reference-ID type changed "
                f"{old_col.ref_id_type or '(none)'} -> {new_col.ref_id_type or '(none)'}",
            )
        )
    if old_col.is_key != new_col.is_key:
        diff.changes.append(
            Change(
                kind=ChangeKind.KEY_CHANGED,
                sheet=sheet,
                column=name,
                old=str(old_col.is_key).lower(),
                new=str(new_col.is_key).lower(),
                detail=f"Column {name!r} {'became' if new_col.is_key else 'is no longer'} "
                "the spreadsheet key",
            )
        )


# ---------------------------------------------------------------------------
# Workbook content diff


def diff_workbooks(old: Workbook, new: Workbook) -> WorkbookDiff:
    """Content diff between two filled loads, keyed on spreadsheet keys.

    When the two workbooks' schemas differ, the schema drift is attached as
    ``schema_drift`` and cells are still compared column-by-header for the
    columns both sides share.
    """
    diff = WorkbookDiff(old_source=old.path, new_source=new.path)

    drift = diff_templates(old.template, new.template)
    if drift.changes:
        diff.schema_drift = drift

    old_data = {d.name: d for d in old.data}
    new_data = {d.name: d for d in new.data}
    for name in old_data:
        if name in new_data:
            _diff_sheet_data(diff, old_data[name], new_data[name])
    # Sheets present on only one side are already reported via schema_drift.

    diff.changes.sort(key=lambda c: (c.sheet, _key_sort(c.key), c.column, c.kind.value))
    return diff


def _diff_sheet_data(diff: WorkbookDiff, old: Any, new: Any) -> None:
    shared = [
        (old_col, new_col)
        for old_col in old.spec.columns
        if old_col.header
        and (new_col := new.spec.column_by_header(old_col.header)) is not None
    ]

    old_rows = _index_rows(old)
    new_rows = _index_rows(new)

    for key in old_rows:
        if key not in new_rows:
            diff.changes.append(
                Change(
                    kind=ChangeKind.ROW_REMOVED,
                    sheet=old.name,
                    key=_display_key(key),
                    detail=f"Row {_row_label(key)} exists only in the old workbook",
                )
            )
    for key in new_rows:
        if key not in old_rows:
            diff.changes.append(
                Change(
                    kind=ChangeKind.ROW_ADDED,
                    sheet=new.name,
                    key=_display_key(key),
                    detail=f"Row {_row_label(key)} exists only in the new workbook",
                )
            )

    for key, old_r in old_rows.items():
        new_r = new_rows.get(key)
        if new_r is None:
            continue
        for old_col, new_col in shared:
            old_val = old.value(old_r, old_col.index - 1)
            new_val = new.value(new_r, new_col.index - 1)
            if _norm(old_val) != _norm(new_val):
                diff.changes.append(
                    Change(
                        kind=ChangeKind.CELL_CHANGED,
                        sheet=new.name,
                        column=new_col.header,
                        key=_display_key(key),
                        old=_display(old_val),
                        new=_display(new_val),
                        detail=f"{new.cell_ref(new_r, new_col.index - 1)}: "
                        f"{_display(old_val)} -> {_display(new_val)} "
                        f"(row {_row_label(key)}, column {new_col.header!r})",
                    )
                )


def _index_rows(sheet: Any) -> dict[str, int]:
    """Map row identity -> 0-based row index.

    Keyed sheets use the normalized spreadsheet key (duplicates keep the first
    occurrence — duplicate keys are the validator's finding, not the differ's).
    Keyless sheets fall back to position (``#0``, ``#1``, ...). All-blank
    padding rows are skipped either way.
    """
    key_col = sheet.spec.key_column
    index: dict[str, int] = {}
    for r, row in enumerate(sheet.rows):
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        if key_col is None:
            index[f"#{r}"] = r
        else:
            token = _norm(sheet.value(r, key_col.index - 1))
            if token and token not in index:
                index[token] = r
    return index


_INT_RE = re.compile(r"^-?\d+$")


def _norm(value: Any) -> str:
    """Normalized comparison text: 1 == 1.0 == "1", dates to ISO, text stripped."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).casefold()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, dt.datetime):
        if value.hour == value.minute == value.second == 0:
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, dt.date):
        return value.isoformat()
    text = str(value).strip()
    if _INT_RE.match(text):
        return str(int(text))
    return text


def _display(value: Any) -> str:
    return "(blank)" if value is None or str(value).strip() == "" else _norm(value)


def _display_key(key: str) -> str:
    return "" if key.startswith("#") else key


def _row_label(key: str) -> str:
    return f"at position {key[1:]}" if key.startswith("#") else f"with key {key!r}"


def _key_sort(key: str) -> tuple[int, Any]:
    """Numeric keys before text keys, each in natural order."""
    if _INT_RE.match(key):
        return (0, int(key))
    return (1, key)
