"""Data models for EIB templates and filled load workbooks.

Workday generates EIB inbound templates as .xlsx workbooks whose data sheets
carry a multi-row header band (column-group row(s), a field-name row, and
usually one or more hint rows with type/format/required annotations) followed
by data rows. Operations with repeating groups split across sheets joined by a
"Spreadsheet Key" column. These models capture that structure in a plain,
serializable form so the validator, differ, and generator (later slices) all
speak the same language.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any


class ColumnType(str, Enum):
    """Declared or inferred type of a template column."""

    TEXT = "text"
    NUMERIC = "numeric"
    DATE = "date"
    BOOLEAN = "boolean"
    REFERENCE = "reference"  # Workday reference-ID column (e.g. Employee_ID, WID)
    UNKNOWN = "unknown"


class SheetRole(str, Enum):
    """Role a sheet plays inside an EIB workbook."""

    PRIMARY = "primary"  # top-level records, one row per business object
    REPEATING = "repeating"  # one-to-many child sheet joined via spreadsheet key
    INSTRUCTIONS = "instructions"  # help/overview sheet emitted with the template
    EMPTY = "empty"  # no usable content


def column_letter(index: int) -> str:
    """1-based column index -> Excel letters (1 -> A, 27 -> AA)."""
    if index < 1:
        raise ValueError(f"column index must be >= 1, got {index}")
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


@dataclass(frozen=True)
class CellRef:
    """Address of a single cell, 1-based like Excel itself."""

    sheet: str
    row: int
    column: int

    @property
    def a1(self) -> str:
        return f"{column_letter(self.column)}{self.row}"

    def __str__(self) -> str:
        return f"{self.sheet}!{self.a1}"


@dataclass
class ColumnSpec:
    """One column of a data sheet, as declared by the template header band.

    ``header`` is the field-name row's text; ``group`` is the (forward-filled)
    column-group row above it, "" when the template has no group row. ``hints``
    keeps the raw hint-row strings so downstream reports can show the template's
    own words as evidence.
    """

    index: int  # 1-based column position
    header: str
    group: str = ""
    col_type: ColumnType = ColumnType.UNKNOWN
    required: bool = False
    is_key: bool = False  # spreadsheet-key column joining multi-sheet workbooks
    ref_id_type: str = ""  # e.g. "Employee_ID" when col_type is REFERENCE
    type_inferred: bool = False  # True when col_type came from data, not the band
    hints: list[str] = field(default_factory=list)

    @property
    def letter(self) -> str:
        return column_letter(self.index)


@dataclass
class SheetSpec:
    """Structure of one sheet: header geometry plus its columns."""

    name: str
    role: SheetRole = SheetRole.PRIMARY
    header_row: int = 0  # 1-based row holding field names; 0 = no header found
    data_start_row: int = 0  # 1-based first data row; 0 = no data region
    columns: list[ColumnSpec] = field(default_factory=list)

    @property
    def key_column(self) -> ColumnSpec | None:
        return next((c for c in self.columns if c.is_key), None)

    def column_by_header(self, header: str) -> ColumnSpec | None:
        want = header.strip().casefold()
        return next((c for c in self.columns if c.header.strip().casefold() == want), None)


@dataclass
class TemplateSpec:
    """The schema side of an EIB workbook: every sheet's structure."""

    source: str = ""  # path the spec was parsed from ("" if built in memory)
    sheets: list[SheetSpec] = field(default_factory=list)

    def data_sheets(self) -> list[SheetSpec]:
        return [s for s in self.sheets if s.role in (SheetRole.PRIMARY, SheetRole.REPEATING)]

    def primary_sheet(self) -> SheetSpec | None:
        return next((s for s in self.sheets if s.role is SheetRole.PRIMARY), None)

    def sheet(self, name: str) -> SheetSpec | None:
        return next((s for s in self.sheets if s.name == name), None)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class SheetData:
    """Data rows of one sheet, kept addressable back to real cells.

    ``rows`` is row-major raw cell values exactly as openpyxl read them
    (trailing all-blank rows/columns trimmed). Row *i* of ``rows`` lives at
    Excel row ``spec.data_start_row + i``; column *j* (0-based) is template
    column ``j + 1``.
    """

    spec: SheetSpec
    rows: list[list[Any]] = field(default_factory=list)

    @property
    def name(self) -> str:
        return self.spec.name

    def cell_ref(self, row_index: int, column_index: int) -> CellRef:
        """CellRef for 0-based (row_index, column_index) into ``rows``."""
        return CellRef(
            sheet=self.spec.name,
            row=self.spec.data_start_row + row_index,
            column=column_index + 1,
        )

    def value(self, row_index: int, column_index: int) -> Any:
        row = self.rows[row_index]
        return row[column_index] if column_index < len(row) else None

    def column_values(self, column_index: int) -> list[Any]:
        """All values of a 0-based column, padded with None for short rows."""
        return [r[column_index] if column_index < len(r) else None for r in self.rows]


@dataclass
class Workbook:
    """A parsed EIB workbook: template structure plus (possibly empty) data."""

    path: str
    template: TemplateSpec
    data: list[SheetData] = field(default_factory=list)

    def sheet_data(self, name: str) -> SheetData | None:
        return next((d for d in self.data if d.name == name), None)

    def to_dict(self) -> dict[str, Any]:
        """JSON-safe summary (structure + row counts, not full cell data)."""
        return {
            "path": self.path,
            "template": self.template.to_dict(),
            "row_counts": {d.name: len(d.rows) for d in self.data},
        }
