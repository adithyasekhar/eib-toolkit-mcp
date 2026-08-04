# eib-toolkit-mcp

**Workday EIB Toolkit** — generate, validate, and diff Workday EIB load spreadsheets, locally
and deterministically. Ships as a Python library, a CLI, and an MCP server so an AI assistant
can do EIB prep with files-in/files-out — no tenant credentials involved.

> **Unofficial** — not affiliated with, endorsed by, or supported by Workday, Inc. Works
> entirely on files you already have (EIB templates and load workbooks exported from your
> tenant). Nothing here connects to a tenant.

## Why

Inbound EIB loads live or die by a spreadsheet. Practitioners lose hours to
upload→fail→fix loops because Workday's staging errors are vague and row-anonymous, templates
drift between releases, and multi-sheet spreadsheet keys are maintained by hand
(see [docs/RESEARCH.md](docs/RESEARCH.md) for the survey with sources). This toolkit is the
pre-flight layer that should have existed:

- **Validate** a filled workbook against its template *before* upload — typed, evidence-backed
  findings with sheet/row/column coordinates (bad dates, wrong reference-ID types, missing
  required fields, broken spreadsheet keys, encoding hazards).
- **Diff** two template generations to see exactly what a Workday release changed, or two
  filled workbooks to review a load like a code review.
- **Generate** a correctly-typed, correctly-keyed load workbook from a declarative load spec
  plus your CSV data — with an optional Claude assist that drafts the spec from natural
  language (Claude drafts and annotates; generation itself stays deterministic).

## Quick Start

Install from source:

```bash
pip install -e ".[dev,mcp]"
```

Get started in seconds:

```bash
# Inspect a template
eib-toolkit inspect your_template.xlsx

# Validate a filled load
eib-toolkit validate your_load.xlsx

# Generate from a spec
eib-toolkit generate --spec load.yaml -o output.xlsx template.xlsx

# Start the MCP server (for Claude)
eib-toolkit mcp
```

Exit codes: **0** clean | **1** findings found | **2** operational error.

## Examples

Bundled fixtures live in `tests/fixtures/` — all **synthetic data** (no real PII).

**Inspect a template:**
```
$ eib-toolkit inspect tests/fixtures/template.xlsx
tests/fixtures/template.xlsx
  [primary] Employees, 0 data row(s)
      A  Employee ID (WID) [unknown]
      B  First Name [unknown]
      C  Last Name [unknown]
      D  Date of Birth [date]
      E  Hire Date [date]
      F  Annual Salary [unknown]
      G  Eligible for Benefits [unknown]
  [primary] Dependents, 0 data row(s)
      A  Employee WID (link) [unknown]
      B  Dependent Name [unknown]
      C  Relationship [unknown]
      D  Date of Birth [date]
```

**Validate a clean load:**
```
$ eib-toolkit validate tests/fixtures/load_clean.xlsx
WARNING  EIB021  Employees!F3  Number stored as text in column F ('Annual Salary'); coerces to 85000.00
WARNING  EIB021  Employees!F4  Number stored as text in column F ('Annual Salary'); coerces to 92000.00
WARNING  EIB021  Employees!F5  Number stored as text in column F ('Annual Salary'); coerces to 78000.00
OK: 3 finding(s) (0 error, 3 warning, 0 info)
```

**Diff two template versions:**
```
$ eib-toolkit diff tests/fixtures/template.xlsx tests/fixtures/template_next.xlsx
schema drift (non-breaking):
      column_added  Column 'Department (NEW)' (H, unknown, optional) was added
0 content change(s)
```

**Generate from a declarative spec:**
```
$ eib-toolkit generate --spec tests/fixtures/spec.yaml -o load.xlsx tests/fixtures/template.xlsx
Wrote 2 row(s) to load.xlsx (Employees: 2)
Post-write validation: clean (2 finding(s) total)
```

## Fixtures

Bundled example workbooks in `tests/fixtures/`:

- `template.xlsx` — Baseline EIB template (Employees + Dependents sheets)
- `load_clean.xlsx` — Properly-filled load (passes validation)
- `load_broken.xlsx` — Deliberately-broken load (exercises validation rules)
- `template_next.xlsx` — Next-release template (for drift diffing)
- `spec.yaml` — Load spec (declarative mapping of template columns to CSV fields)
- `data.csv` — Sample input data (pairs with `spec.yaml` for generate examples)

All fixtures contain **clearly-marked synthetic data**: no real tenant names, employee IDs, or PII.

## How It Works

### Validation

Typed, evidence-backed findings with sheet/row/column coordinates:

- Required values
- Date parsing (ISO 8601 vs. locale formats)
- Numeric coercion (decimal commas, thousands separators)
- Boolean normalization (Yes/No/Y/N/1/0/True/False)
- Reference ID shape checks
- Spreadsheet key integrity (duplicates, broken joins, orphaned keys)
- Cell-length ceilings
- XML-unsafe characters

Every finding includes a **code** (e.g., `EIB021` for coercible numbers), **severity**
(ERROR / WARNING / INFO), and **evidence** (the actual value that triggered the finding).

### Generation

Generate a filled .xlsx from:

1. A template (Workday-generated EIB or your own)
2. A load spec (YAML/JSON: sheet mappings, column mappings, key strategy)
3. Input data (one CSV per target sheet)

The generator:

- Fills a **copy of the template**, preserving multi-row headers byte-for-byte
- Auto-assigns spreadsheet keys (sequential for primary rows, join-keyed for repeating groups)
- Writes value-level problems verbatim and warns, so the validator shows what a loader sees
- Validates its own output by default (use `--no-check` to skip)

### Diffing

Compare two templates or two filled workbooks:

- **Templates**: schema changes (columns added/removed/renamed, type changes, requiredness changes)
- **Workbooks**: content changes (normalized: `1 == 1.0 == "1"`), schema drift attached

### Claude Integration

Optional Claude assist layer (enabled if `ANTHROPIC_API_KEY` is set):

- `draft-spec` — Claude drafts a load spec from natural language
- `validate` — Claude annotates findings (but never adds, removes, or reclassifies them)

Generation itself is always deterministic — Claude only augments the human-readable input and
output, never the core logic.

## Status

**v0.1.0** — Core functionality complete and tested.

- ✅ Parser (template + filled workbook)
- ✅ Validator (146 test cases, ~30 distinct rules)
- ✅ Generator (load spec + CSV → filled workbook)
- ✅ Differ (templates and workbooks)
- ✅ CLI (6 commands with text/JSON/markdown output)
- ✅ MCP server (6 tools, error-as-JSON contract)
- ✅ Synthetic fixtures + full test suite (146 tests, all passing)

See [docs/PLAN.md](docs/PLAN.md) for the build roadmap and
[docs/RESEARCH.md](docs/RESEARCH.md) for the landscape research underlying this design.

## Roadmap

- [x] Research + scaffold
- [x] Core models + template/workbook parser
- [x] Validation engine (evidence-backed findings)
- [x] Generator + template/workbook diff
- [x] CLI + MCP server (FastMCP, optional Claude assist)
- [x] Synthetic fixtures + full test suite
- [x] v0.1.0

## License

MIT © Adithya Sekhar Gummadi. Workday, EIB, and related marks are trademarks of Workday, Inc.
