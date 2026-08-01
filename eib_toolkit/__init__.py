"""Workday EIB Toolkit — generate, validate, and diff EIB load spreadsheets.

Deterministic core, optional Claude assist, MCP server included. Unofficial:
not affiliated with Workday, Inc.
"""

__version__ = "0.1.0"

from eib_toolkit.diff import (
    Change,
    ChangeKind,
    TemplateDiff,
    WorkbookDiff,
    diff_templates,
    diff_workbooks,
)
from eib_toolkit.generate import (
    ColumnMapping,
    GenerateError,
    GenerateReport,
    LoadSpec,
    SheetMapping,
    generate_workbook,
    load_spec,
)
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
    "Change",
    "ChangeKind",
    "ColumnMapping",
    "ColumnSpec",
    "ColumnType",
    "Finding",
    "GenerateError",
    "GenerateReport",
    "LoadSpec",
    "Severity",
    "SheetData",
    "SheetMapping",
    "SheetRole",
    "SheetSpec",
    "TemplateDiff",
    "TemplateSpec",
    "ValidationConfig",
    "ValidationReport",
    "Workbook",
    "WorkbookDiff",
    "__version__",
    "diff_templates",
    "diff_workbooks",
    "generate_workbook",
    "load_spec",
    "parse_template",
    "parse_workbook",
    "validate_workbook",
]
