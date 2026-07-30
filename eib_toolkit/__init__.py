"""Workday EIB Toolkit — generate, validate, and diff EIB load spreadsheets.

Deterministic core, optional Claude assist, MCP server included. Unofficial:
not affiliated with Workday, Inc.
"""

__version__ = "0.1.0"

from eib_toolkit.model import (
    CellRef,
    ColumnSpec,
    ColumnType,
    SheetData,
    SheetRole,
    SheetSpec,
    TemplateSpec,
    Workbook,
)
from eib_toolkit.parser import parse_template, parse_workbook

__all__ = [
    "CellRef",
    "ColumnSpec",
    "ColumnType",
    "SheetData",
    "SheetRole",
    "SheetSpec",
    "TemplateSpec",
    "Workbook",
    "__version__",
    "parse_template",
    "parse_workbook",
]
