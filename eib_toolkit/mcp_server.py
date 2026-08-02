"""MCP server mode — drive the EIB toolkit from Claude Desktop or Claude Code.

Requires the optional dependency:  pip install "eib-toolkit-mcp[mcp]"
Register with:  claude mcp add eib-toolkit -- eib-toolkit mcp

Every tool is stateless and files-in/files-out: it takes paths to workbooks,
specs, and CSVs the user already has, and returns JSON (or writes an .xlsx /
.yaml next to them). Nothing here connects to a Workday tenant and no
credentials are involved. Errors come back as ``{"error": ...}`` rather than
raising, so a conversation can recover.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "MCP mode needs the 'mcp' package: pip install 'eib-toolkit-mcp[mcp]'"
    ) from exc

from eib_toolkit import diff as diff_mod
from eib_toolkit.generate import GenerateError, load_spec
from eib_toolkit.generate import generate_workbook as _generate
from eib_toolkit.parser import parse_template, parse_workbook
from eib_toolkit.validate import validate_workbook as _validate

mcp = FastMCP(
    "eib-toolkit",
    instructions=(
        "Workday EIB Toolkit. Typical flows: inspect_template on an EIB "
        "template the user exported, draft_load_spec to map their CSV data "
        "onto it, generate_workbook to fill it, validate_workbook before "
        "upload. diff_templates shows release-to-release schema drift; "
        "diff_workbooks reviews two filled loads. All outputs are "
        "deterministic and evidence-backed; advisory Claude notes are "
        "clearly marked and never change findings."
    ),
)


def _err(message: str) -> str:
    return json.dumps({"error": message})


def _checked_file(path: str) -> Path | None:
    p = Path(path).expanduser()
    return p if p.is_file() else None


@mcp.tool()
def inspect_template(path: str) -> str:
    """Parse an EIB template or filled load workbook and return its structure.

    Returns JSON: every sheet's role (primary/repeating/instructions), header
    geometry, and per-column header, type, requiredness, reference-ID type,
    and spreadsheet-key role, plus data row counts.
    """
    p = _checked_file(path)
    if p is None:
        return _err(f"Not a file: {path}")
    try:
        return json.dumps(parse_workbook(p).to_dict(), indent=2)
    except Exception as exc:
        return _err(f"Cannot read {path} as an .xlsx workbook: {exc}")


@mcp.tool()
def validate_workbook(path: str, annotate: bool = False) -> str:
    """Validate a filled EIB load workbook before upload.

    Returns JSON findings, each addressed to a real sheet/row/column with the
    offending value as evidence, plus severity rollups and an overall ok flag.
    Set annotate=true to add advisory fix-it notes via the Anthropic API
    (needs ANTHROPIC_API_KEY; notes never change the findings).
    """
    p = _checked_file(path)
    if p is None:
        return _err(f"Not a file: {path}")
    try:
        report = _validate(parse_workbook(p))
    except Exception as exc:
        return _err(f"Cannot read {path} as an .xlsx workbook: {exc}")
    out = report.to_dict()
    if annotate:
        from eib_toolkit.claude import annotate_findings

        out["annotations"] = annotate_findings(report)
    return json.dumps(out, indent=2)


@mcp.tool()
def diff_templates(old_path: str, new_path: str) -> str:
    """Schema-drift report between two EIB template generations.

    Run this before reusing last cycle's spreadsheets: returns JSON changes
    (columns added/removed/renamed, type/requiredness/key changes) with an
    overall breaking verdict. Renames are heuristic and flagged as such.
    """
    for path in (old_path, new_path):
        if _checked_file(path) is None:
            return _err(f"Not a file: {path}")
    try:
        d = diff_mod.diff_templates(parse_template(old_path), parse_template(new_path))
    except Exception as exc:
        return _err(f"Cannot diff: {exc}")
    return json.dumps(d.to_dict(), indent=2)


@mcp.tool()
def diff_workbooks(old_path: str, new_path: str) -> str:
    """Content diff between two filled EIB loads of the same template.

    Rows are matched by spreadsheet key (so reordering is not noise), values
    compared after normalization (1 == 1.0 == "1"). Any schema drift between
    the two files is attached to the JSON result as schema_drift.
    """
    for path in (old_path, new_path):
        if _checked_file(path) is None:
            return _err(f"Not a file: {path}")
    try:
        d = diff_mod.diff_workbooks(parse_workbook(old_path), parse_workbook(new_path))
    except Exception as exc:
        return _err(f"Cannot diff: {exc}")
    return json.dumps(d.to_dict(), indent=2)


@mcp.tool()
def generate_workbook(template_path: str, spec_path: str, output_path: str) -> str:
    """Fill an EIB template from a load spec (.yaml/.json) + its CSVs.

    Writes a filled .xlsx to output_path (header band preserved, values
    typed, spreadsheet keys auto-assigned) and immediately validates it.
    Returns JSON: rows written, key assignments, generation warnings, and
    the post-write validation report.
    """
    for path in (template_path, spec_path):
        if _checked_file(path) is None:
            return _err(f"Not a file: {path}")
    try:
        spec = load_spec(spec_path)
        report = _generate(template_path, spec, output_path)
        validation = _validate(parse_workbook(output_path))
    except GenerateError as exc:
        return _err(str(exc))
    except Exception as exc:
        return _err(f"Generation failed: {exc}")
    out = report.to_dict()
    out["validation"] = validation.to_dict()
    return json.dumps(out, indent=2)


@mcp.tool()
def draft_load_spec(template_path: str, instruction: str = "", output_path: str = "") -> str:
    """Draft a load spec for an EIB template.

    The skeleton (every column mapped to a suggested snake_case CSV field) is
    always deterministic. With an instruction and ANTHROPIC_API_KEY, Claude
    adapts it — the draft is then re-validated against the template and
    rejected back to the skeleton if structurally wrong. Returns JSON with
    the spec, its YAML rendering, and notes on how it was produced; set
    output_path to also write the .yaml next to the user's files.
    """
    p = _checked_file(template_path)
    if p is None:
        return _err(f"Not a file: {template_path}")
    import yaml

    from eib_toolkit.claude import draft_load_spec as _draft

    try:
        template = parse_template(p)
    except Exception as exc:
        return _err(f"Cannot read {template_path} as an .xlsx workbook: {exc}")
    if not template.data_sheets():
        return _err(f"{template_path} has no data sheets to draft a spec for")
    spec, notes = _draft(template, instruction=instruction)
    text = yaml.safe_dump(spec, sort_keys=False, allow_unicode=True)
    result: dict[str, Any] = {"spec": spec, "yaml": text, "notes": notes}
    if output_path:
        Path(output_path).expanduser().write_text(text, encoding="utf-8")
        result["written"] = output_path
    return json.dumps(result, indent=2)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
