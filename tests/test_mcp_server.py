"""MCP server tests — registered tool schemas plus files-in/files-out behavior.

Skipped cleanly when the optional 'mcp' dependency is not installed. The
tools are plain functions after registration, so behavior is tested by
calling them directly; schemas are read from the FastMCP registry.
"""

import asyncio
import json
from pathlib import Path

import pytest

pytest.importorskip("mcp")

from eib_toolkit import mcp_server
from tests.helpers import make_filled, make_spec_and_csvs, make_template

EXPECTED_TOOLS = {
    "inspect_template": {"path"},
    "validate_workbook": {"path", "annotate"},
    "diff_templates": {"old_path", "new_path"},
    "diff_workbooks": {"old_path", "new_path"},
    "generate_workbook": {"template_path", "spec_path", "output_path"},
    "draft_load_spec": {"template_path", "instruction", "output_path"},
}


def test_registered_tool_schemas() -> None:
    tools = {t.name: t for t in asyncio.run(mcp_server.mcp.list_tools())}
    assert set(tools) == set(EXPECTED_TOOLS)
    for name, params in EXPECTED_TOOLS.items():
        tool = tools[name]
        assert tool.description, f"{name} has no description"
        assert set(tool.inputSchema["properties"]) == params, name


def test_inspect_and_validate_tools(tmp_path: Path) -> None:
    template = make_template(tmp_path / "template.xlsx")
    filled = make_filled(tmp_path / "filled.xlsx")

    structure = json.loads(mcp_server.inspect_template(str(template)))
    assert [s["name"] for s in structure["template"]["sheets"]] == [
        "Instructions", "Workers", "Allowances",
    ]

    report = json.loads(mcp_server.validate_workbook(str(filled)))
    assert report["ok"] is True and "annotations" not in report

    broken = make_filled(
        tmp_path / "broken.xlsx",
        rows={"Workers": [[1, "EMP-1001", None, "bad-date", 1, "Y"]]},
    )
    report = json.loads(mcp_server.validate_workbook(str(broken)))
    assert report["ok"] is False
    assert {f["code"] for f in report["findings"]} >= {"EIB001", "EIB010"}


def test_diff_and_generate_tools(tmp_path: Path) -> None:
    template = make_template(tmp_path / "template.xlsx")
    spec = make_spec_and_csvs(tmp_path)
    out = tmp_path / "load.xlsx"

    result = json.loads(
        mcp_server.generate_workbook(str(template), str(spec), str(out))
    )
    assert result["rows_written"] == {"Workers": 2, "Allowances": 3}
    assert result["validation"]["ok"] is True
    assert out.is_file()

    # Template-vs-template: identical schema, no drift. (Diffing a blank
    # template against a *filled* workbook would report inferred-type drift
    # on data-bearing columns — that is intended behavior, not schema identity.)
    template2 = make_template(tmp_path / "template2.xlsx")
    drift = json.loads(mcp_server.diff_templates(str(template), str(template2)))
    assert drift["breaking"] is False and drift["changes"] == []

    content = json.loads(mcp_server.diff_workbooks(str(out), str(out)))
    assert content["changes"] == [] and content["schema_drift"] is None


def test_draft_tool_writes_yaml(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    template = make_template(tmp_path / "template.xlsx")
    out = tmp_path / "spec.yaml"
    result = json.loads(mcp_server.draft_load_spec(str(template), output_path=str(out)))
    assert result["written"] == str(out)
    assert result["spec"]["sheets"][0]["sheet"] == "Workers"
    assert "sheet: Workers" in out.read_text(encoding="utf-8")
    assert any("Deterministic skeleton" in n for n in result["notes"])


def test_tools_return_json_errors_not_exceptions(tmp_path: Path) -> None:
    missing = str(tmp_path / "nope.xlsx")
    for call in (
        lambda: mcp_server.inspect_template(missing),
        lambda: mcp_server.validate_workbook(missing),
        lambda: mcp_server.diff_templates(missing, missing),
        lambda: mcp_server.diff_workbooks(missing, missing),
        lambda: mcp_server.generate_workbook(missing, missing, str(tmp_path / "o.xlsx")),
        lambda: mcp_server.draft_load_spec(missing),
    ):
        assert "error" in json.loads(call())

    bogus = tmp_path / "bogus.xlsx"
    bogus.write_text("not a workbook")
    assert "error" in json.loads(mcp_server.inspect_template(str(bogus)))
