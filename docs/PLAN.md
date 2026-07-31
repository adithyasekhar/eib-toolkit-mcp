# Build plan — eib-toolkit-mcp

One slice per day. Each slice ships complete: real code **and** its tests, quality gate green
(pytest, ruff, CLI smoke once a CLI exists) before push. Design rationale in
[RESEARCH.md](RESEARCH.md).

## Slices

- [x] **Slice 0 — Research + scaffold** (2026-07-29)
  Landscape research with sources; pyproject/LICENSE/CI/README skeleton; package + placeholder
  tests.

- [x] **Slice 1 — Core models + workbook parser** (2026-07-30) (`eib_toolkit/model.py`, `eib_toolkit/parser.py`)
  Dataclasses: `TemplateSpec` (sheets, columns with header text, required flag, declared type,
  reference-ID type, spreadsheet-key role), `Workbook`/`SheetData`/`CellRef`. Parser reads an
  EIB template or filled load workbook (.xlsx via openpyxl, read-only mode): detect header
  rows across the multi-row header band EIB templates use, classify columns (date, numeric,
  boolean, text, reference-ID), detect sheet roles (primary vs. repeating-group sheets keyed
  by spreadsheet key). Tolerant of messy real-world exports: blank padding rows/columns,
  merged header cells, stray instruction sheets. Tests on hand-built minimal workbooks.

- [x] **Slice 2 — Validation engine** (2026-07-31) (`eib_toolkit/validate.py`)
  Rule-based checks producing `Finding(severity, code, sheet, row, column, message, evidence)`:
  missing required values; date parse/format (ISO 8601 vs. locale text); numeric coercion
  (decimal commas, thousands separators); boolean normalization; reference-ID shape checks
  (WID-looking vs. named reference-ID columns); duplicate/broken/orphaned spreadsheet keys
  across sheets; row-count and cell-length ceilings; non-UTF-8-safe characters. Every finding
  row/column-addressed and evidence-backed. Deterministic severity; summary rollup. Full test
  coverage per rule.

- [ ] **Slice 3 — Generator + diff** (`eib_toolkit/generate.py`, `eib_toolkit/diff.py`)
  Generate: declarative load spec (YAML/JSON: target sheets, column mappings, constants,
  key strategy) + input CSV → filled .xlsx matching a template, with correct types and
  auto-assigned spreadsheet keys for one-to-many sheets. Diff: (a) template vs. template —
  release-drift report (added/removed/renamed columns, requiredness/type changes); (b) filled
  workbook vs. filled workbook — content diff keyed on spreadsheet keys. Tests round-trip
  generate→parse→validate.

- [ ] **Slice 4 — CLI + MCP server** (`eib_toolkit/cli.py`, `eib_toolkit/mcp_server.py`)
  CLI subcommands: `inspect`, `validate`, `diff`, `generate` (+ `--json`/markdown report
  output). Restore `[project.scripts]` in pyproject; add CLI smoke step to CI. MCP server via
  FastMCP exposing 5–7 files-in/files-out tools (inspect_template, validate_workbook,
  diff_templates, diff_workbooks, generate_workbook, draft_load_spec). Optional Claude layer
  (`ANTHROPIC_API_KEY`): draft a load spec from natural language and annotate findings —
  Claude drafts/annotates, never decides. Tests for CLI exit codes and tool schemas.

- [ ] **Slice 5 — Synthetic fixtures + README + release** (`tests/fixtures/`, `examples/`)
  Realistic synthetic fixture set (clearly marked synthetic; fake tenant/workers, no real
  PII): a template workbook, a clean filled load, a deliberately-broken load exercising every
  validation rule, a "next release" template for drift diff, a load spec + CSV pair.
  End-to-end tests over fixtures; README finalized with real captured CLI output; tag v0.1.0.

## Scope guards

No tenant connectivity or EIB launching; no XSLT authoring; natural language only in the
optional Claude layer. If a slice runs long, cut scope here honestly rather than pushing
broken work.
