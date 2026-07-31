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
from eib_toolkit.validate import (
    RULES,
    Finding,
    Severity,
    ValidationConfig,
    ValidationReport,
    validate_workbook,
)

__all__ = [
    "RULES",
    "CellRef",
    "ColumnSpec",
    "ColumnType",
    "Finding",
    "Severity",
    "SheetData",
    "SheetRole",
    "SheetSpec",
    "TemplateSpec",
    "ValidationConfig",
    "ValidationReport",
    "Workbook",
    "__version__",
    "parse_template",
    "parse_workbook",
    "validate_workbook",
]
