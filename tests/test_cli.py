"""CLI tests — exit codes, output formats, and the end-to-end command flows.

Exit-code contract under test: 0 clean, 1 findings/differences present,
2 operational error. All workbooks are synthetic (see helpers.py).
"""

import json
from pathlib import Path

import pytest
import yaml

from eib_toolkit.cli import main
from tests.helpers import make_filled, make_spec_and_csvs, make_template


@pytest.fixture
def template(tmp_path: Path) -> Path:
    return make_template(tmp_path / "template.xlsx")


@pytest.fixture
def filled(tmp_path: Path) -> Path:
    return make_filled(tmp_path / "filled.xlsx")


# ---------------------------------------------------------------------------
# inspect


def test_inspect_text(template: Path, capsys) -> None:
    assert main(["inspect", str(template)]) == 0
    out = capsys.readouterr().out
    assert "[primary] Workers" in out
    assert "[repeating] Allowances" in out
    assert "Spreadsheet Key [numeric" in out or "Spreadsheet Key" in out


def test_inspect_json(template: Path, capsys) -> None:
    assert main(["inspect", str(template), "--format", "json"]) == 0
    data = json.loads(capsys.readouterr().out)
    names = [s["name"] for s in data["template"]["sheets"]]
    assert names == ["Instructions", "Workers", "Allowances"]


def test_inspect_markdown(template: Path, capsys) -> None:
    assert main(["inspect", str(template), "--format", "markdown"]) == 0
    out = capsys.readouterr().out
    assert "## Workers (primary" in out
    assert "| Col | Header |" in out


def test_inspect_missing_file_exits_2(tmp_path: Path, capsys) -> None:
    assert main(["inspect", str(tmp_path / "nope.xlsx")]) == 2
    assert "not a file" in capsys.readouterr().err


def test_inspect_not_xlsx_exits_2(tmp_path: Path, capsys) -> None:
    bogus = tmp_path / "bogus.xlsx"
    bogus.write_text("this is not a zip")
    assert main(["inspect", str(bogus)]) == 2
    assert "cannot read" in capsys.readouterr().err


# ---------------------------------------------------------------------------
# validate


def test_validate_clean_exits_0(filled: Path, capsys) -> None:
    assert main(["validate", str(filled)]) == 0
    assert "OK: 0 finding(s)" in capsys.readouterr().out


def test_validate_broken_exits_1_with_findings(tmp_path: Path, capsys) -> None:
    broken = make_filled(
        tmp_path / "broken.xlsx",
        rows={
            "Workers": [
                [1, "EMP-1001", None, "01/02/2026", "sixty", "Y"],  # blank req, bad date/number
            ],
            "Allowances": [
                [9, "Transit", 120, "2026-02-01"],  # orphaned key
            ],
        },
    )
    assert main(["validate", str(broken), "--format", "json"]) == 1
    report = json.loads(capsys.readouterr().out)
    assert report["ok"] is False
    codes = {f["code"] for f in report["findings"]}
    assert {"EIB001", "EIB011", "EIB020", "EIB052"} <= codes
    assert report["annotations"] == {"annotated_by": "none", "notes": {}}


def test_validate_markdown(tmp_path: Path, capsys) -> None:
    broken = make_filled(
        tmp_path / "broken.xlsx",
        rows={"Workers": [[1, "EMP-1001", "Avery Example", "not-a-date", 1, "Y"]]},
    )
    assert main(["validate", str(broken), "--format", "markdown"]) == 1
    out = capsys.readouterr().out
    assert "# Validation report" in out
    assert "| error | EIB010 |" in out


# ---------------------------------------------------------------------------
# diff


def test_diff_identical_templates_exits_0(template: Path, tmp_path: Path, capsys) -> None:
    other = make_template(tmp_path / "other.xlsx")
    assert main(["diff", str(template), str(other), "--templates"]) == 0
    assert "identical" in capsys.readouterr().out


def test_diff_changed_workbooks_exits_1(filled: Path, tmp_path: Path, capsys) -> None:
    changed = make_filled(
        tmp_path / "changed.xlsx",
        rows={
            "Workers": [
                [1, "EMP-1001", "Avery Example", "2026-02-01", 65000, "Y"],  # raise
                [2, "EMP-1002", "Blake Sample", "2026-03-15", 58000, "N"],
            ],
            "Allowances": [
                [1, "Transit", 120, "2026-02-01"],
                [1, "Meals", 80, "2026-02-01"],
                [2, "Transit", 120, "2026-03-15"],
            ],
        },
    )
    assert main(["diff", str(filled), str(changed), "--format", "json"]) == 1
    data = json.loads(capsys.readouterr().out)
    assert data["summary"] == {"cell_changed": 1}
    change = data["changes"][0]
    assert (change["column"], change["old"], change["new"]) == ("Annual Amount", "61000", "65000")


def test_diff_markdown(filled: Path, tmp_path: Path, capsys) -> None:
    other = make_filled(tmp_path / "same.xlsx")
    assert main(["diff", str(filled), str(other), "--format", "markdown"]) == 0
    assert "No differences." in capsys.readouterr().out


# ---------------------------------------------------------------------------
# generate


def test_generate_end_to_end(template: Path, tmp_path: Path, capsys) -> None:
    spec = make_spec_and_csvs(tmp_path)
    out = tmp_path / "load.xlsx"
    assert main(["generate", str(template), "--spec", str(spec), "-o", str(out)]) == 0
    assert out.is_file()
    text = capsys.readouterr().out
    assert "Wrote 5 row(s)" in text
    assert "Post-write validation: clean" in text
    # And the generated load validates clean through the CLI too.
    assert main(["validate", str(out)]) == 0


def test_generate_json_includes_validation(template: Path, tmp_path: Path, capsys) -> None:
    spec = make_spec_and_csvs(tmp_path)
    out = tmp_path / "load.xlsx"
    args = ["generate", str(template), "--spec", str(spec), "-o", str(out), "--format", "json"]
    assert main(args) == 0
    data = json.loads(capsys.readouterr().out)
    assert data["rows_written"] == {"Workers": 2, "Allowances": 3}
    assert data["validation"]["ok"] is True


def test_generate_bad_spec_exits_2(template: Path, tmp_path: Path, capsys) -> None:
    make_spec_and_csvs(tmp_path)
    bad = tmp_path / "bad.yaml"
    bad.write_text(
        "sheets:\n  - sheet: Nowhere\n    source: workers.csv\n", encoding="utf-8"
    )
    out = tmp_path / "load.xlsx"
    assert main(["generate", str(template), "--spec", str(bad), "-o", str(out)]) == 2
    assert "no data sheet named 'Nowhere'" in capsys.readouterr().err
    assert not out.exists()


# ---------------------------------------------------------------------------
# draft-spec (deterministic path; the Claude path is tested in test_claude.py)


def test_draft_spec_round_trip(template: Path, tmp_path: Path, capsys, monkeypatch) -> None:
    """Drafted skeleton -> rename CSVs to the suggested names -> generate cleanly."""
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    spec_path = tmp_path / "drafted.yaml"
    assert main(["draft-spec", str(template), "-o", str(spec_path)]) == 0

    drafted = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    assert [s["sheet"] for s in drafted["sheets"]] == ["Workers", "Allowances"]
    assert all(s["key_from"] == "business_key" for s in drafted["sheets"])

    # Provide CSVs matching the skeleton's suggested field names, then generate.
    (tmp_path / "workers.csv").write_text(
        "business_key,employee_reference,legal_name,hire_date,annual_amount,active\n"
        "EMP-1001,EMP-1001,Avery Example,2026-02-01,61000,Y\n",
        encoding="utf-8",
    )
    (tmp_path / "allowances.csv").write_text(
        "business_key,allowance_plan,amount,effective_date\n"
        "EMP-1001,Transit,120,2026-02-01\n",
        encoding="utf-8",
    )
    out = tmp_path / "load.xlsx"
    assert main(["generate", str(template), "--spec", str(spec_path), "-o", str(out)]) == 0
    assert "Post-write validation: clean" in capsys.readouterr().out


def test_draft_spec_stdout_and_notes(template: Path, capsys, monkeypatch) -> None:
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    assert main(["draft-spec", str(template)]) == 0
    captured = capsys.readouterr()
    assert yaml.safe_load(captured.out)["key_strategy"] == "sequential"
    assert "note: Deterministic skeleton" in captured.err


def test_draft_spec_no_data_sheets_exits_2(tmp_path: Path, capsys) -> None:
    from tests.helpers import write_xlsx

    empty = write_xlsx(tmp_path / "empty.xlsx", {"Instructions": [["Nothing here."]]})
    assert main(["draft-spec", str(empty)]) == 2
    assert "no data sheets" in capsys.readouterr().err
