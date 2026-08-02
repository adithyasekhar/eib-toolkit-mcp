"""Shared synthetic-workbook builders for the test suite and the CI smoke run.

Everything here is synthetic: fake sheet names, fake workers, fake IDs.
No real PII and no real tenant data anywhere.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook as XlsxWorkbook


def write_xlsx(path: Path, sheets: dict[str, list[list[Any]]]) -> Path:
    wb = XlsxWorkbook()
    assert wb.active is not None
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return path


def make_template(path: Path) -> Path:
    """Synthetic two-data-sheet EIB template: workers + repeating allowances."""
    return write_xlsx(
        path,
        {
            "Instructions": [
                ["Synthetic EIB template for tests. All names and IDs are fake."],
            ],
            "Workers": [
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
                    "Required",
                    "Required. Reference ID Type: Employee_ID",
                    "Required. Text",
                    "Required. Date (yyyy-mm-dd)",
                    "Numeric",
                    "Boolean (Y/N)",
                ],
            ],
            "Allowances": [
                [
                    "Spreadsheet Key*",
                    "Allowance Plan",
                    "Amount",
                    "Effective Date",
                ],
                [
                    "Required",
                    "Required. Text",
                    "Required. Numeric",
                    "Date (yyyy-mm-dd)",
                ],
            ],
        },
    )


def make_filled(path: Path, rows: dict[str, list[list[Any]]] | None = None) -> Path:
    """A filled load for :func:`make_template`'s schema (clean by default)."""
    default: dict[str, list[list[Any]]] = {
        "Workers": [
            [1, "EMP-1001", "Avery Example", "2026-02-01", 61000, "Y"],
            [2, "EMP-1002", "Blake Sample", "2026-03-15", 58000, "N"],
        ],
        "Allowances": [
            [1, "Transit", 120, "2026-02-01"],
            [1, "Meals", 80, "2026-02-01"],
            [2, "Transit", 120, "2026-03-15"],
        ],
    }
    data = rows if rows is not None else default
    return write_xlsx(
        path,
        {
            "Instructions": [
                ["Synthetic filled EIB load for tests. All names and IDs are fake."],
            ],
            "Workers": [
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
                    "Required",
                    "Required. Reference ID Type: Employee_ID",
                    "Required. Text",
                    "Required. Date (yyyy-mm-dd)",
                    "Numeric",
                    "Boolean (Y/N)",
                ],
                *data.get("Workers", []),
            ],
            "Allowances": [
                [
                    "Spreadsheet Key*",
                    "Allowance Plan",
                    "Amount",
                    "Effective Date",
                ],
                [
                    "Required",
                    "Required. Text",
                    "Required. Numeric",
                    "Date (yyyy-mm-dd)",
                ],
                *data.get("Allowances", []),
            ],
        },
    )


def make_spec_and_csvs(directory: Path) -> Path:
    """A YAML load spec + matching CSVs for :func:`make_template`'s schema."""
    (directory / "workers.csv").write_text(
        "employee_id,name,hired,salary,active\n"
        "EMP-1001,Avery Example,2026-02-01,61000,yes\n"
        "EMP-1002,Blake Sample,2026-03-15,58000,no\n",
        encoding="utf-8",
    )
    (directory / "allowances.csv").write_text(
        "employee_id,plan,amount,effective\n"
        "EMP-1001,Transit,120,2026-02-01\n"
        "EMP-1001,Meals,80,2026-02-01\n"
        "EMP-1002,Transit,120,2026-03-15\n",
        encoding="utf-8",
    )
    spec = directory / "spec.yaml"
    spec.write_text(
        """\
key_strategy: sequential
sheets:
  - sheet: Workers
    source: workers.csv
    key_from: employee_id
    columns:
      - column: Employee_Reference
        source: employee_id
      - column: Legal Name
        source: name
      - column: Hire Date
        source: hired
      - column: Annual Amount
        source: salary
      - column: Active
        source: active
  - sheet: Allowances
    source: allowances.csv
    key_from: employee_id
    columns:
      - column: Allowance Plan
        source: plan
      - column: Amount
        source: amount
      - column: Effective Date
        source: effective
""",
        encoding="utf-8",
    )
    return spec
