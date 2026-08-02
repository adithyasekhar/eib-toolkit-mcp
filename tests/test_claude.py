"""Tests for the optional Claude layer — drafts and annotates, never decides.

No network anywhere: the Anthropic client is replaced by fakes. What is
under test is the deterministic machinery around Claude — the skeleton
drafting, the structural gate that re-validates every draft, and the
guarantee that annotation leaves the validation report untouched.
"""

import copy
import json
from pathlib import Path

import pytest

from eib_toolkit.claude import (
    annotate_findings,
    draft_load_spec,
    draft_spec_skeleton,
    spec_problems,
    suggest_field_name,
)
from eib_toolkit.model import TemplateSpec
from eib_toolkit.parser import parse_template, parse_workbook
from eib_toolkit.validate import validate_workbook
from tests.helpers import make_filled, make_template


@pytest.fixture
def template(tmp_path: Path) -> TemplateSpec:
    return parse_template(make_template(tmp_path / "template.xlsx"))


class FakeClient:
    """Stands in for anthropic.Anthropic; returns a canned response text."""

    def __init__(self, text: str = "", error: Exception | None = None) -> None:
        self.requests: list[dict] = []
        self._text = text
        self._error = error
        outer = self

        class _Messages:
            def create(self, **kwargs):
                outer.requests.append(kwargs)
                if outer._error is not None:
                    raise outer._error
                block = type("Block", (), {"text": outer._text})()
                return type("Resp", (), {"content": [block]})()

        self.messages = _Messages()


# ---------------------------------------------------------------------------
# Field-name suggestion + skeleton


def test_suggest_field_name() -> None:
    assert suggest_field_name("Legal Name") == "legal_name"
    assert suggest_field_name("Employee_Reference") == "employee_reference"
    assert suggest_field_name("  Amount (USD)  ") == "amount_usd"
    assert suggest_field_name("***") == "field"


def test_skeleton_covers_template(template: TemplateSpec) -> None:
    skeleton = draft_spec_skeleton(template)
    assert skeleton["key_strategy"] == "sequential"
    workers, allowances = skeleton["sheets"]
    assert workers["sheet"] == "Workers"
    assert workers["source"] == "workers.csv"
    assert workers["key_from"] == "business_key"
    headers = [c["column"] for c in workers["columns"]]
    assert headers == [
        "Employee_Reference", "Legal Name", "Hire Date", "Annual Amount", "Active",
    ]
    assert all(c["source"] == suggest_field_name(c["column"]) for c in workers["columns"])
    assert "Spreadsheet Key" not in headers  # the key is never mapped
    assert allowances["sheet"] == "Allowances"
    # The skeleton must always pass its own structural gate.
    assert spec_problems(skeleton, template) == []


# ---------------------------------------------------------------------------
# The structural gate


def test_spec_problems_catches_bad_drafts(template: TemplateSpec) -> None:
    good = draft_spec_skeleton(template)

    bad = copy.deepcopy(good)
    bad["sheets"][0]["sheet"] = "Nowhere"
    assert any("no data sheet named 'Nowhere'" in p for p in spec_problems(bad, template))

    bad = copy.deepcopy(good)
    bad["sheets"][0]["columns"][0]["column"] = "Imaginary"
    problems = spec_problems(bad, template)
    assert any("no column 'Imaginary'" in p for p in problems)
    assert any("required column(s) not mapped: Employee_Reference" in p for p in problems)

    bad = copy.deepcopy(good)
    bad["sheets"][0]["columns"].append({"column": "Spreadsheet Key", "source": "k"})
    assert any("spreadsheet key" in p for p in spec_problems(bad, template))

    bad = copy.deepcopy(good)
    bad["sheets"][0]["columns"].append({"column": "Legal Name", "source": "again"})
    assert any("mapped twice" in p for p in spec_problems(bad, template))

    bad = copy.deepcopy(good)
    del bad["sheets"][0]["key_from"]
    assert any("needs 'key_from'" in p for p in spec_problems(bad, template))

    assert spec_problems(["not", "a", "mapping"], template) != []


# ---------------------------------------------------------------------------
# Drafting


def test_draft_without_instruction_is_skeleton(template: TemplateSpec) -> None:
    spec, notes = draft_load_spec(template)
    assert spec == draft_spec_skeleton(template)
    assert notes == ["Deterministic skeleton (no instruction given)."]


def test_draft_without_client_falls_back(template: TemplateSpec, monkeypatch) -> None:
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    spec, notes = draft_load_spec(template, instruction="load workers from hr.csv")
    assert spec == draft_spec_skeleton(template)
    assert "Claude is not available" in notes[0]


def test_draft_accepts_valid_claude_output(template: TemplateSpec) -> None:
    draft = draft_spec_skeleton(template)
    draft["sheets"][0]["source"] = "hr_export.csv"
    draft["sheets"][0]["columns"][1]["source"] = "full_name"
    client = FakeClient(text=json.dumps(draft))

    spec, notes = draft_load_spec(template, instruction="workers come from hr_export.csv", client=client)
    assert spec == draft
    assert notes == ["Drafted by Claude from your instruction; review mappings before generating."]
    # The instruction and template summary actually reached the model.
    payload = json.loads(client.requests[0]["messages"][0]["content"])
    assert payload["instruction"] == "workers come from hr_export.csv"
    assert payload["template"][0]["sheet"] == "Workers"


def test_draft_strips_code_fences(template: TemplateSpec) -> None:
    draft = draft_spec_skeleton(template)
    client = FakeClient(text=f"```json\n{json.dumps(draft)}\n```")
    spec, _notes = draft_load_spec(template, instruction="anything", client=client)
    assert spec == draft


def test_draft_rejects_structurally_bad_claude_output(template: TemplateSpec) -> None:
    """Claude drafts, the deterministic gate decides: bad drafts never escape."""
    draft = draft_spec_skeleton(template)
    draft["sheets"][0]["columns"][0]["column"] = "Hallucinated Column"
    client = FakeClient(text=json.dumps(draft))

    spec, notes = draft_load_spec(template, instruction="anything", client=client)
    assert spec == draft_spec_skeleton(template)
    assert "rejected by structural validation" in notes[0]
    assert any("no column 'Hallucinated Column'" in n for n in notes[1:])


def test_draft_survives_client_errors(template: TemplateSpec) -> None:
    client = FakeClient(error=RuntimeError("network down"))
    spec, notes = draft_load_spec(template, instruction="anything", client=client)
    assert spec == draft_spec_skeleton(template)
    assert "Claude draft failed (RuntimeError)" in notes[0]

    client = FakeClient(text="I would rather write prose than JSON.")
    spec, notes = draft_load_spec(template, instruction="anything", client=client)
    assert spec == draft_spec_skeleton(template)
    assert "Claude draft failed" in notes[0]


# ---------------------------------------------------------------------------
# Annotation


@pytest.fixture
def broken_report(tmp_path: Path):
    broken = make_filled(
        tmp_path / "broken.xlsx",
        rows={"Workers": [[1, "EMP-1001", None, "01/02/2026", "sixty", "Y"]]},
    )
    return validate_workbook(parse_workbook(broken))


def test_annotate_without_client_is_none(broken_report, monkeypatch) -> None:
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    assert annotate_findings(broken_report) == {"annotated_by": "none", "notes": {}}


def test_annotate_empty_report_never_calls_claude(tmp_path: Path) -> None:
    clean = make_filled(tmp_path / "clean.xlsx")
    report = validate_workbook(parse_workbook(clean))
    client = FakeClient(error=AssertionError("must not be called"))
    assert annotate_findings(report, client=client) == {"annotated_by": "none", "notes": {}}


def test_annotate_adds_notes_but_never_touches_findings(broken_report) -> None:
    codes = sorted({f.code for f in broken_report.findings})
    notes = {code: f"Fix the {code} rows in your source data." for code in codes}
    notes["EIB999"] = "An invented code that must be dropped."
    client = FakeClient(text=json.dumps(notes))

    before = broken_report.to_dict()
    result = annotate_findings(broken_report, client=client)
    assert result["annotated_by"] == "claude"
    assert sorted(result["notes"]) == codes  # invented code dropped
    assert broken_report.to_dict() == before  # findings byte-for-byte untouched

    # The payload grouped by code with rule descriptions and samples.
    payload = json.loads(client.requests[0]["messages"][0]["content"])
    assert sorted(payload) == codes
    assert all({"rule", "severity", "count", "samples"} <= set(v) for v in payload.values())


def test_annotate_survives_bad_output(broken_report) -> None:
    for text in ("not json", json.dumps(["a", "list"]), json.dumps({})):
        result = annotate_findings(broken_report, client=FakeClient(text=text))
        assert result == {"annotated_by": "none", "notes": {}}
    result = annotate_findings(broken_report, client=FakeClient(error=RuntimeError("boom")))
    assert result == {"annotated_by": "none", "notes": {}}
