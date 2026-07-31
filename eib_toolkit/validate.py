"""Rule-based validation of filled EIB load workbooks.

Every check is deterministic and produces a :class:`Finding` addressed to a
real cell (or sheet) with the offending value quoted as evidence — the same
"evidence-backed, auditable output" contract the rest of the toolkit follows.
The rules encode the failure modes that actually sink EIB loads in practice:
blank required cells, locale-formatted dates, numbers stored as text,
unnormalizable booleans, WIDs pasted into named reference-ID columns, broken
spreadsheet-key joins across sheets, oversized loads, and characters that
break the XML the EIB transform generates. Severity is fixed per rule; an
optional Claude layer (later slice) may annotate findings but never decides.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any

from eib_toolkit.model import (
    CellRef,
    ColumnSpec,
    ColumnType,
    SheetData,
    Workbook,
)

# ---------------------------------------------------------------------------
# Findings


class Severity(str, Enum):
    """Fixed, rule-assigned severity. ERROR blocks a load; WARNING is likely
    to load but with mangled or ambiguous data; INFO is awareness only."""

    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


#: Every rule this module can emit: code -> (severity, one-line description).
RULES: dict[str, tuple[Severity, str]] = {
    "EIB001": (Severity.ERROR, "Required cell is blank"),
    "EIB010": (Severity.ERROR, "Unparseable value in a date column"),
    "EIB011": (Severity.WARNING, "Locale-formatted date text (not ISO 8601 yyyy-mm-dd)"),
    "EIB020": (Severity.ERROR, "Non-numeric value in a numeric column"),
    "EIB021": (Severity.WARNING, "Number stored as text (needs coercion)"),
    "EIB030": (Severity.ERROR, "Unrecognized boolean token"),
    "EIB040": (Severity.WARNING, "WID-shaped value in a named reference-ID column"),
    "EIB041": (Severity.ERROR, "Value in a WID column is not a WID"),
    "EIB050": (Severity.ERROR, "Duplicate spreadsheet key on the primary sheet"),
    "EIB051": (Severity.ERROR, "Missing spreadsheet key on a repeating-group row"),
    "EIB052": (Severity.ERROR, "Orphaned spreadsheet key (no matching primary row)"),
    "EIB060": (Severity.WARNING, "Sheet exceeds the configured row ceiling"),
    "EIB061": (Severity.ERROR, "Cell exceeds the maximum cell length"),
    "EIB070": (Severity.ERROR, "Character not safe for EIB XML generation"),
    "EIB071": (Severity.WARNING, "Leading or trailing whitespace"),
    "EIB080": (Severity.WARNING, "Value in a column with no header"),
}


@dataclass
class Finding:
    """One rule violation, addressed to a cell (row/column 0 = sheet-level)."""

    severity: Severity
    code: str
    sheet: str
    row: int  # 1-based Excel row; 0 for sheet-level findings
    column: int  # 1-based column index; 0 for sheet-level findings
    message: str
    evidence: str = ""

    @property
    def location(self) -> str:
        if self.row and self.column:
            return str(CellRef(sheet=self.sheet, row=self.row, column=self.column))
        return self.sheet

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["location"] = self.location
        return d


@dataclass
class ValidationConfig:
    """Tunable ceilings; every default is deliberately conservative."""

    max_rows_per_sheet: int = 30_000  # practical EIB batch guidance, not a hard limit
    max_cell_length: int = 32_767  # the .xlsx hard ceiling
    max_findings_per_rule_per_sheet: int = 200  # keep reports readable on pathological files


@dataclass
class ValidationReport:
    """All findings for one workbook plus deterministic rollups."""

    source: str
    findings: list[Finding] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(f.severity is Severity.ERROR for f in self.findings)

    def by_severity(self) -> dict[str, int]:
        counts = {s.value: 0 for s in Severity}
        for f in self.findings:
            counts[f.severity.value] += 1
        return counts

    def by_code(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for f in sorted(self.findings, key=lambda f: f.code):
            counts[f.code] = counts.get(f.code, 0) + 1
        return counts

    def by_sheet(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for f in self.findings:
            counts[f.sheet] = counts.get(f.sheet, 0) + 1
        return counts

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "ok": self.ok,
            "summary": {
                "total": len(self.findings),
                "by_severity": self.by_severity(),
                "by_code": self.by_code(),
                "by_sheet": self.by_sheet(),
            },
            "findings": [f.to_dict() for f in self.findings],
        }


# ---------------------------------------------------------------------------
# Value-shape patterns

_WID_RE = re.compile(r"^[0-9a-f]{32}$", re.IGNORECASE)
_WID_TYPE_RE = re.compile(r"^wid$", re.IGNORECASE)

_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2})?)?$")
#: Recognizable-but-locale-ambiguous date shapes: 1/15/2026, 15.01.2026,
#: 15-Jan-2026, Jan 15 2026, 2026/01/15 ...
_LOCALE_DATE_RE = re.compile(
    r"^(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}"
    r"|\d{4}[/.]\d{1,2}[/.]\d{1,2}"
    r"|\d{1,2}[ \-][A-Za-z]{3,9}[ \-,]+\d{2,4}"
    r"|[A-Za-z]{3,9}[ \-]\d{1,2}[ \-,]+\d{2,4})$"
)

_PLAIN_NUMBER_RE = re.compile(r"^-?\d+(\.\d+)?$")
#: Coercible numeric text: thousands separators and/or a decimal comma.
_COERCIBLE_NUMBER_RE = re.compile(
    r"^-?(\d{1,3}(,\d{3})+(\.\d+)?"  # 1,234,567.89
    r"|\d{1,3}(\.\d{3})+(,\d+)?"  # 1.234.567,89
    r"|\d+,\d+)$"  # 1234,56
)

_BOOL_TOKENS = frozenset({"y", "n", "yes", "no", "true", "false", "1", "0"})

#: Characters invalid in XML 1.0 — these break the EIB's spreadsheet-to-XML
#: transform outright. Everything else printable is fine in UTF-8.
_XML_UNSAFE_RE = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f"
    "\ufdd0-\ufdef\ufffe\uffff]"
)


def _describe_char(ch: str) -> str:
    name = unicodedata.name(ch, "") or f"control character 0x{ord(ch):02X}"
    return f"U+{ord(ch):04X} ({name})"


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _quote(value: Any, limit: int = 60) -> str:
    text = repr(value) if isinstance(value, str) else str(value)
    return text if len(text) <= limit else text[: limit - 3] + "..."


# ---------------------------------------------------------------------------
# The validator


class _Collector:
    """Accumulates findings, capping per (rule, sheet) so a pathological file
    yields a readable report instead of a million identical rows."""

    def __init__(self, config: ValidationConfig) -> None:
        self.config = config
        self.findings: list[Finding] = []
        self._counts: dict[tuple[str, str], int] = {}
        self.suppressed: dict[tuple[str, str], int] = {}

    def add(
        self,
        code: str,
        sheet: str,
        row: int,
        column: int,
        message: str,
        evidence: str = "",
    ) -> None:
        key = (code, sheet)
        n = self._counts.get(key, 0)
        if n >= self.config.max_findings_per_rule_per_sheet:
            self.suppressed[key] = self.suppressed.get(key, 0) + 1
            return
        self._counts[key] = n + 1
        severity, _ = RULES[code]
        self.findings.append(
            Finding(
                severity=severity,
                code=code,
                sheet=sheet,
                row=row,
                column=column,
                message=message,
                evidence=evidence,
            )
        )

    def finalize(self) -> list[Finding]:
        for (code, sheet), extra in sorted(self.suppressed.items()):
            severity, description = RULES[code]
            self.findings.append(
                Finding(
                    severity=severity,
                    code=code,
                    sheet=sheet,
                    row=0,
                    column=0,
                    message=(
                        f"{extra} further finding(s) of {code} ({description}) suppressed "
                        f"after the first {self.config.max_findings_per_rule_per_sheet}"
                    ),
                )
            )
        return self.findings


def validate_workbook(
    workbook: Workbook, config: ValidationConfig | None = None
) -> ValidationReport:
    """Run every rule against a parsed workbook and return the full report.

    Findings come back in reading order (sheet as ordered in the workbook,
    then row, column, code) so reports are stable across runs.
    """
    config = config or ValidationConfig()
    collector = _Collector(config)

    primary = workbook.template.primary_sheet()
    primary_keys = _check_primary_keys(collector, workbook, primary)

    for sheet in workbook.data:
        _check_sheet(collector, sheet, config)
        if sheet.spec.role.value == "repeating":
            _check_repeating_keys(collector, sheet, primary_keys)

    findings = collector.finalize()
    sheet_order = {d.name: i for i, d in enumerate(workbook.data)}
    findings.sort(
        key=lambda f: (
            sheet_order.get(f.sheet, len(sheet_order)),
            f.row,
            f.column,
            f.code,
        )
    )
    return ValidationReport(source=workbook.path, findings=findings)


# ---------------------------------------------------------------------------
# Per-sheet cell rules


def _check_sheet(collector: _Collector, sheet: SheetData, config: ValidationConfig) -> None:
    if len(sheet.rows) > config.max_rows_per_sheet:
        collector.add(
            "EIB060",
            sheet.name,
            0,
            0,
            f"Sheet has {len(sheet.rows)} data rows, over the configured ceiling of "
            f"{config.max_rows_per_sheet}; consider splitting the load",
            evidence=f"{len(sheet.rows)} rows",
        )

    columns = sheet.spec.columns
    for r, row in enumerate(sheet.rows):
        if all(_is_blank(v) for v in row):
            continue  # interior padding row: not data, nothing to validate
        for c, column in enumerate(columns):
            value = row[c] if c < len(row) else None
            if not column.header and not _is_blank(value):
                # The parser builds a headerless ColumnSpec for any used column
                # the template's field-name row does not cover.
                ref = sheet.cell_ref(r, c)
                collector.add(
                    "EIB080",
                    sheet.name,
                    ref.row,
                    ref.column,
                    "Value present in a column the template declares no header for",
                    evidence=_quote(value),
                )
                continue
            _check_cell(collector, sheet, r, c, column, value, config)
        for c in range(len(columns), len(row)):
            if not _is_blank(row[c]):
                ref = sheet.cell_ref(r, c)
                collector.add(
                    "EIB080",
                    sheet.name,
                    ref.row,
                    ref.column,
                    "Value present in a column the template declares no header for",
                    evidence=_quote(row[c]),
                )


def _check_cell(
    collector: _Collector,
    sheet: SheetData,
    r: int,
    c: int,
    column: ColumnSpec,
    value: Any,
    config: ValidationConfig,
) -> None:
    ref = sheet.cell_ref(r, c)

    if _is_blank(value):
        if column.required:
            collector.add(
                "EIB001",
                sheet.name,
                ref.row,
                ref.column,
                f"Required column {column.letter} ({column.header!r}) is blank",
            )
        return

    if isinstance(value, str):
        text = value
        if len(text) > config.max_cell_length:
            collector.add(
                "EIB061",
                sheet.name,
                ref.row,
                ref.column,
                f"Cell is {len(text)} characters, over the {config.max_cell_length} limit",
                evidence=_quote(text, 40),
            )
        for match in {m.group(0) for m in _XML_UNSAFE_RE.finditer(text)}:
            collector.add(
                "EIB070",
                sheet.name,
                ref.row,
                ref.column,
                f"Cell contains {_describe_char(match)}, which is invalid in EIB XML",
                evidence=_quote(text),
            )
        if text != text.strip():
            collector.add(
                "EIB071",
                sheet.name,
                ref.row,
                ref.column,
                f"Value in column {column.letter} ({column.header!r}) has leading or "
                "trailing whitespace; Workday will treat it as a distinct value",
                evidence=_quote(text),
            )

    checker = _TYPE_CHECKS.get(column.col_type)
    if checker is not None:
        checker(collector, sheet.name, ref, column, value)


def _check_date(
    collector: _Collector, sheet: str, ref: CellRef, column: ColumnSpec, value: Any
) -> None:
    if not isinstance(value, str):
        return  # real date/datetime cells serialize cleanly
    text = value.strip()
    if _ISO_DATE_RE.match(text):
        return
    if _LOCALE_DATE_RE.match(text):
        collector.add(
            "EIB011",
            sheet,
            ref.row,
            ref.column,
            f"Date in column {column.letter} ({column.header!r}) is locale-formatted; "
            "use ISO 8601 (yyyy-mm-dd) to avoid day/month ambiguity",
            evidence=_quote(text),
        )
        return
    collector.add(
        "EIB010",
        sheet,
        ref.row,
        ref.column,
        f"Value in date column {column.letter} ({column.header!r}) is not a "
        "recognizable date",
        evidence=_quote(text),
    )


def _check_numeric(
    collector: _Collector, sheet: str, ref: CellRef, column: ColumnSpec, value: Any
) -> None:
    if isinstance(value, bool):
        collector.add(
            "EIB020",
            sheet,
            ref.row,
            ref.column,
            f"Boolean in numeric column {column.letter} ({column.header!r})",
            evidence=_quote(value),
        )
        return
    if isinstance(value, (int, float)):
        return
    text = str(value).strip()
    if _PLAIN_NUMBER_RE.match(text) or _COERCIBLE_NUMBER_RE.match(text):
        canonical = _canonical_number(text)
        collector.add(
            "EIB021",
            sheet,
            ref.row,
            ref.column,
            f"Number stored as text in column {column.letter} ({column.header!r}); "
            f"coerces to {canonical}",
            evidence=_quote(value),
        )
        return
    collector.add(
        "EIB020",
        sheet,
        ref.row,
        ref.column,
        f"Value in numeric column {column.letter} ({column.header!r}) is not a number",
        evidence=_quote(value),
    )


def _canonical_number(text: str) -> str:
    """Best-effort canonical form for coercible numeric text (evidence only)."""
    t = text
    if _COERCIBLE_NUMBER_RE.match(t):
        if re.match(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$", t) or re.match(r"^-?\d+,\d+$", t):
            t = t.replace(".", "").replace(",", ".")  # European style
        else:
            t = t.replace(",", "")  # US thousands separators
    return t


def _check_boolean(
    collector: _Collector, sheet: str, ref: CellRef, column: ColumnSpec, value: Any
) -> None:
    if isinstance(value, bool):
        return
    text = str(value).strip().casefold()
    if text in _BOOL_TOKENS:
        return
    collector.add(
        "EIB030",
        sheet,
        ref.row,
        ref.column,
        f"Value in boolean column {column.letter} ({column.header!r}) is not one of "
        "Y/N, Yes/No, True/False, 1/0",
        evidence=_quote(value),
    )


def _check_reference(
    collector: _Collector, sheet: str, ref: CellRef, column: ColumnSpec, value: Any
) -> None:
    text = str(value).strip()
    declared_wid = bool(_WID_TYPE_RE.match(column.ref_id_type))
    looks_wid = bool(_WID_RE.match(text))
    if declared_wid and not looks_wid:
        collector.add(
            "EIB041",
            sheet,
            ref.row,
            ref.column,
            f"Column {column.letter} ({column.header!r}) declares reference-ID type WID "
            "but the value is not a 32-character hex WID",
            evidence=_quote(text),
        )
    elif not declared_wid and looks_wid:
        declared = column.ref_id_type or "a named reference-ID type"
        collector.add(
            "EIB040",
            sheet,
            ref.row,
            ref.column,
            f"Column {column.letter} ({column.header!r}) expects {declared} values but "
            "this looks like a WID; WIDs are tenant-specific and fragile across "
            "environments",
            evidence=_quote(text),
        )


_TYPE_CHECKS = {
    ColumnType.DATE: _check_date,
    ColumnType.NUMERIC: _check_numeric,
    ColumnType.BOOLEAN: _check_boolean,
    ColumnType.REFERENCE: _check_reference,
}


# ---------------------------------------------------------------------------
# Cross-sheet spreadsheet-key rules


def _key_token(value: Any) -> str:
    """Normalize a key for joining: 1, 1.0 and "1" are the same key."""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if re.match(r"^-?\d+$", text):
        return str(int(text))
    return text.casefold()


def _check_primary_keys(
    collector: _Collector, workbook: Workbook, primary_spec: Any
) -> set[str]:
    """Emit duplicate-key findings for the primary sheet; return its key set."""
    if primary_spec is None or primary_spec.key_column is None:
        return set()
    sheet = workbook.sheet_data(primary_spec.name)
    if sheet is None:
        return set()

    key_col = primary_spec.key_column
    seen: dict[str, int] = {}  # token -> first Excel row
    for r, value in enumerate(sheet.column_values(key_col.index - 1)):
        if _is_blank(value):
            continue  # blank rows / blank keys on primary are the required-cell rule's job
        token = _key_token(value)
        ref = sheet.cell_ref(r, key_col.index - 1)
        first = seen.get(token)
        if first is None:
            seen[token] = ref.row
        else:
            collector.add(
                "EIB050",
                sheet.name,
                ref.row,
                ref.column,
                f"Spreadsheet key {_quote(value)} duplicates row {first}; keys must be "
                "unique on the primary sheet",
                evidence=_quote(value),
            )
    return set(seen)


def _check_repeating_keys(
    collector: _Collector, sheet: SheetData, primary_keys: set[str]
) -> None:
    key_col = sheet.spec.key_column
    if key_col is None:
        return
    for r, row in enumerate(sheet.rows):
        if all(_is_blank(v) for v in row):
            continue
        value = sheet.value(r, key_col.index - 1)
        ref = sheet.cell_ref(r, key_col.index - 1)
        if _is_blank(value):
            collector.add(
                "EIB051",
                sheet.name,
                ref.row,
                ref.column,
                "Repeating-group row has no spreadsheet key, so it cannot be joined to "
                "a primary row",
            )
        elif _key_token(value) not in primary_keys:
            collector.add(
                "EIB052",
                sheet.name,
                ref.row,
                ref.column,
                f"Spreadsheet key {_quote(value)} has no matching row on the primary "
                "sheet; this row would be silently dropped or rejected",
                evidence=_quote(value),
            )
