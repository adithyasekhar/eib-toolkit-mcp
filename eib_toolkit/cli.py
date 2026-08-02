"""eib-toolkit CLI.

Commands:
  inspect     Parse an EIB template or filled workbook and show its structure
  validate    Validate a filled load workbook -> evidence-backed findings
  diff        Diff two filled workbooks (default) or two templates (--templates)
  generate    Fill a template from a load spec + CSVs (output is validated)
  draft-spec  Draft a load spec (deterministic skeleton; Claude adapts it
              to --instruction when ANTHROPIC_API_KEY is configured)
  mcp         Start the MCP server (stdio) for Claude Desktop / Claude Code

Every command reports as text (default), ``--format json`` (full machine
output), or ``--format markdown`` (paste-ready report).

Exit codes: 0 clean; 1 findings or differences present (validate, diff, and
generate's post-write validation); 2 operational error (missing file, bad
spec, unreadable workbook — argparse usage errors are also 2).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from eib_toolkit.diff import TemplateDiff, WorkbookDiff, diff_templates, diff_workbooks
from eib_toolkit.generate import GenerateError, generate_workbook, load_spec
from eib_toolkit.model import Workbook
from eib_toolkit.parser import parse_template, parse_workbook
from eib_toolkit.validate import ValidationReport, validate_workbook


class CliError(Exception):
    """Operational failure — reported to stderr, exit code 2."""


def _require_file(path: str) -> Path:
    p = Path(path)
    if not p.is_file():
        raise CliError(f"not a file: {p}")
    return p


def _parse_workbook(path: str) -> Workbook:
    try:
        return parse_workbook(_require_file(path))
    except CliError:
        raise
    except Exception as exc:  # zip errors, not-an-xlsx, truncated files
        raise CliError(f"cannot read {path} as an .xlsx workbook: {exc}") from exc


# ---------------------------------------------------------------------------
# inspect


def cmd_inspect(args: argparse.Namespace) -> int:
    wb = _parse_workbook(args.workbook)
    if args.format == "json":
        print(json.dumps(wb.to_dict(), indent=2))
    elif args.format == "markdown":
        print(_md_inspect(wb))
    else:
        print(f"{wb.path}")
        for sheet in wb.template.sheets:
            data = wb.sheet_data(sheet.name)
            rows = f", {len(data.rows)} data row(s)" if data is not None else ""
            print(f"  [{sheet.role.value}] {sheet.name}{rows}")
            for col in sheet.columns:
                if not col.header and sheet.role.value not in ("primary", "repeating"):
                    continue
                print(f"    {col.letter:>3}  {_col_line(col)}")
    return 0


def _col_line(col: Any) -> str:
    flags = []
    if col.is_key:
        flags.append("key")
    if col.required:
        flags.append("required")
    if col.ref_id_type:
        flags.append(f"ref:{col.ref_id_type}")
    if col.type_inferred:
        flags.append("inferred")
    suffix = f" ({', '.join(flags)})" if flags else ""
    return f"{col.header or '(no header)'} [{col.col_type.value}]{suffix}"


def _md_inspect(wb: Workbook) -> str:
    lines = [f"# Workbook structure — `{wb.path}`", ""]
    for sheet in wb.template.data_sheets():
        data = wb.sheet_data(sheet.name)
        rows = len(data.rows) if data is not None else 0
        lines += [
            f"## {sheet.name} ({sheet.role.value}, {rows} data row(s))",
            "",
            "| Col | Header | Type | Required | Notes |",
            "| --- | --- | --- | --- | --- |",
        ]
        for col in sheet.columns:
            notes = []
            if col.is_key:
                notes.append("spreadsheet key")
            if col.ref_id_type:
                notes.append(f"ref-ID: {col.ref_id_type}")
            if col.type_inferred:
                notes.append("type inferred from data")
            lines.append(
                f"| {col.letter} | {col.header or '(no header)'} | {col.col_type.value} "
                f"| {'yes' if col.required else ''} | {', '.join(notes)} |"
            )
        lines.append("")
    return "\n".join(lines).rstrip()


# ---------------------------------------------------------------------------
# validate


def cmd_validate(args: argparse.Namespace) -> int:
    wb = _parse_workbook(args.workbook)
    report = validate_workbook(wb)
    annotations: dict[str, Any] = {"annotated_by": "none", "notes": {}}
    if args.annotate:
        from eib_toolkit.claude import annotate_findings

        annotations = annotate_findings(report)

    if args.format == "json":
        out = report.to_dict()
        out["annotations"] = annotations
        print(json.dumps(out, indent=2))
    elif args.format == "markdown":
        print(_md_validate(report, annotations))
    else:
        for f in report.findings:
            evidence = f" | {f.evidence}" if f.evidence else ""
            print(f"{f.severity.value.upper():>7}  {f.code}  {f.location}  "
                  f"{f.message}{evidence}")
        counts = report.by_severity()
        print(
            f"{'OK' if report.ok else 'FAIL'}: {len(report.findings)} finding(s) "
            f"({counts['error']} error, {counts['warning']} warning, {counts['info']} info)"
        )
        for code, note in annotations["notes"].items():
            print(f"  note [{code}]: {note}")
    return 0 if report.ok else 1


def _md_validate(report: ValidationReport, annotations: dict[str, Any]) -> str:
    counts = report.by_severity()
    lines = [
        f"# Validation report — `{report.source}`",
        "",
        f"**{'OK' if report.ok else 'FAIL'}** — {len(report.findings)} finding(s): "
        f"{counts['error']} error, {counts['warning']} warning, {counts['info']} info.",
        "",
    ]
    if report.findings:
        lines += [
            "| Severity | Code | Location | Message | Evidence |",
            "| --- | --- | --- | --- | --- |",
        ]
        lines += [
            f"| {f.severity.value} | {f.code} | {f.location} | {_md_cell(f.message)} "
            f"| {_md_cell(f.evidence)} |"
            for f in report.findings
        ]
        lines.append("")
    if annotations["notes"]:
        lines.append(f"## Notes (advisory, by {annotations['annotated_by']})")
        lines.append("")
        lines += [f"- **{code}** — {note}" for code, note in annotations["notes"].items()]
    return "\n".join(lines).rstrip()


def _md_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


# ---------------------------------------------------------------------------
# diff


def cmd_diff(args: argparse.Namespace) -> int:
    if args.templates:
        old = parse_template(_require_file(args.old))
        new = parse_template(_require_file(args.new))
        diff: TemplateDiff | WorkbookDiff = diff_templates(old, new)
        changed = bool(diff.changes)
    else:
        diff = diff_workbooks(_parse_workbook(args.old), _parse_workbook(args.new))
        changed = bool(diff.changes) or diff.schema_drift is not None

    if args.format == "json":
        print(json.dumps(diff.to_dict(), indent=2))
    elif args.format == "markdown":
        print(_md_diff(diff))
    else:
        if isinstance(diff, WorkbookDiff) and diff.schema_drift is not None:
            drift = "BREAKING" if diff.schema_drift.breaking else "non-breaking"
            print(f"schema drift ({drift}):")
            for c in diff.schema_drift.changes:
                print(f"  {c.kind.value:>16}  {c.detail}")
        for c in diff.changes:
            print(f"{c.kind.value:>16}  {c.detail}")
        if isinstance(diff, TemplateDiff):
            verdict = "BREAKING" if diff.breaking else ("changed" if changed else "identical")
            print(f"{len(diff.changes)} change(s) — {verdict}")
        else:
            print(f"{len(diff.changes)} content change(s)")
    return 1 if changed else 0


def _md_diff(diff: TemplateDiff | WorkbookDiff) -> str:
    kind = "Template diff" if isinstance(diff, TemplateDiff) else "Workbook diff"
    lines = [f"# {kind} — `{diff.old_source}` vs `{diff.new_source}`", ""]
    if isinstance(diff, TemplateDiff):
        lines += [f"**Verdict:** {'BREAKING' if diff.breaking else 'non-breaking'}", ""]
    elif diff.schema_drift is not None:
        drift = "BREAKING" if diff.schema_drift.breaking else "non-breaking"
        lines += [f"**Schema drift:** {drift} "
                  f"({len(diff.schema_drift.changes)} change(s), see below)", ""]
    changes = list(diff.changes)
    if isinstance(diff, WorkbookDiff) and diff.schema_drift is not None:
        changes = list(diff.schema_drift.changes) + changes
    if not changes:
        lines.append("No differences.")
    else:
        lines += ["| Kind | Sheet | Column | Key | Old | New |", "| --- | --- | --- | --- | --- | --- |"]
        lines += [
            f"| {c.kind.value} | {c.sheet} | {_md_cell(c.column)} | {_md_cell(c.key)} "
            f"| {_md_cell(c.old)} | {_md_cell(c.new)} |"
            for c in changes
        ]
    return "\n".join(lines).rstrip()


# ---------------------------------------------------------------------------
# generate


def cmd_generate(args: argparse.Namespace) -> int:
    template = _require_file(args.template)
    try:
        spec = load_spec(_require_file(args.spec))
        report = generate_workbook(template, spec, args.output)
    except GenerateError as exc:
        raise CliError(str(exc)) from exc

    validation: ValidationReport | None = None
    if not args.no_check:
        validation = validate_workbook(parse_workbook(args.output))

    if args.format == "json":
        out = report.to_dict()
        out["validation"] = validation.to_dict() if validation else None
        print(json.dumps(out, indent=2))
    else:
        total = sum(report.rows_written.values())
        sheets = ", ".join(f"{name}: {n}" for name, n in report.rows_written.items())
        print(f"Wrote {total} row(s) to {args.output} ({sheets})")
        for warning in report.warnings:
            print(f"  warning: {warning}")
        if validation is not None:
            errors = validation.by_severity()["error"]
            status = "clean" if validation.ok else f"{errors} error(s)"
            print(f"Post-write validation: {status} "
                  f"({len(validation.findings)} finding(s) total)")
    return 0 if validation is None or validation.ok else 1


# ---------------------------------------------------------------------------
# draft-spec


def cmd_draft_spec(args: argparse.Namespace) -> int:
    import yaml

    from eib_toolkit.claude import draft_load_spec

    template = parse_template(_require_file(args.template))
    if not template.data_sheets():
        raise CliError(f"{args.template} has no data sheets to draft a spec for")
    spec, notes = draft_load_spec(template, instruction=args.instruction)
    text = yaml.safe_dump(spec, sort_keys=False, allow_unicode=True)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
        print(f"Load spec written to {args.output}")
    else:
        print(text, end="")
    for note in notes:
        print(f"note: {note}", file=sys.stderr)
    return 0


# ---------------------------------------------------------------------------
# mcp


def cmd_mcp(_args: argparse.Namespace) -> int:
    from eib_toolkit.mcp_server import main as mcp_main

    mcp_main()
    return 0


# ---------------------------------------------------------------------------
# parser / entry point


def _add_format(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        "--format",
        choices=["text", "json", "markdown"],
        default="text",
        help="output format (default: text)",
    )


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="eib-toolkit",
        description="Workday EIB Toolkit — generate, validate, and diff EIB load spreadsheets",
    )
    sub = p.add_subparsers(dest="command", required=True)

    i = sub.add_parser("inspect", help="Show the structure of a template or workbook")
    i.add_argument("workbook", help="EIB template or filled load (.xlsx)")
    _add_format(i)
    i.set_defaults(func=cmd_inspect)

    v = sub.add_parser("validate", help="Validate a filled load workbook")
    v.add_argument("workbook", help="Filled EIB load workbook (.xlsx)")
    v.add_argument(
        "--annotate",
        action="store_true",
        help="Add advisory fix-it notes via Claude (needs ANTHROPIC_API_KEY; "
        "notes never change the findings)",
    )
    _add_format(v)
    v.set_defaults(func=cmd_validate)

    d = sub.add_parser("diff", help="Diff two workbooks (or two templates)")
    d.add_argument("old")
    d.add_argument("new")
    d.add_argument(
        "--templates",
        action="store_true",
        help="Schema-only drift report (release-to-release template comparison)",
    )
    _add_format(d)
    d.set_defaults(func=cmd_diff)

    g = sub.add_parser("generate", help="Fill a template from a load spec + CSVs")
    g.add_argument("template", help="EIB template (.xlsx)")
    g.add_argument("--spec", required=True, help="Load spec (.yaml/.json)")
    g.add_argument("-o", "--output", required=True, help="Output workbook path (.xlsx)")
    g.add_argument(
        "--no-check",
        action="store_true",
        help="Skip validating the generated workbook (validation failures exit 1)",
    )
    g.add_argument("--format", choices=["text", "json"], default="text")
    g.set_defaults(func=cmd_generate)

    ds = sub.add_parser("draft-spec", help="Draft a load spec for a template")
    ds.add_argument("template", help="EIB template (.xlsx)")
    ds.add_argument(
        "--instruction",
        default="",
        help="Natural-language description of the load; adapted by Claude when "
        "ANTHROPIC_API_KEY is set, otherwise the deterministic skeleton is returned",
    )
    ds.add_argument("-o", "--output", default="", help="Write the spec here instead of stdout")
    ds.set_defaults(func=cmd_draft_spec)

    m = sub.add_parser("mcp", help="Start the MCP server (stdio)")
    m.set_defaults(func=cmd_mcp)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except CliError as exc:
        print(f"eib-toolkit: error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
