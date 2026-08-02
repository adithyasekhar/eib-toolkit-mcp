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

## Status

In active development — parser, validator, generator, differ, CLI, and MCP server are in
place; bundled example fixtures and the finalized README land next. See
[docs/PLAN.md](docs/PLAN.md) for the build roadmap and [docs/RESEARCH.md](docs/RESEARCH.md)
for the landscape research this design is based on.

## Roadmap

- [x] Research + scaffold
- [x] Core models + template/workbook parser
- [x] Validation engine (evidence-backed findings)
- [x] Generator + template/workbook diff
- [x] CLI + MCP server (FastMCP, optional Claude assist)
- [ ] Synthetic fixtures + full test suite
- [ ] v0.1.0

## License

MIT © Adithya Sekhar Gummadi. Workday, EIB, and related marks are trademarks of Workday, Inc.
