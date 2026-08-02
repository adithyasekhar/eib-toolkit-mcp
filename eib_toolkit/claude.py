"""Optional Claude assist — drafts load specs and annotates findings.

The suite's hard rule applies throughout: **the deterministic core decides,
Claude only drafts and annotates.**

* Drafting: a load-spec skeleton is always derived deterministically from the
  parsed template. When an ``ANTHROPIC_API_KEY`` is configured and the caller
  gives a natural-language instruction, Claude adapts that skeleton — and the
  result is re-validated structurally against the template before it is
  returned. A draft that references unknown sheets or columns, drops a
  required mapping, or malforms the spec is rejected and the deterministic
  skeleton comes back instead, with a note saying why.

* Annotating: Claude may add advisory remediation notes to a validation
  report, one per rule code present. It can never add, remove, re-address,
  or reclassify a finding — annotations ride alongside the report, the
  findings themselves are untouched.

Everything degrades gracefully: no API key, no ``anthropic`` package, a
network failure, or unparseable model output all fall back to the
deterministic behavior with an explanatory note.
"""

from __future__ import annotations

import json
import os
import re
from typing import Any

from eib_toolkit.generate import GenerateError, LoadSpec, spec_from_dict
from eib_toolkit.model import TemplateSpec
from eib_toolkit.validate import RULES, ValidationReport

__all__ = [
    "annotate_findings",
    "claude_client",
    "draft_load_spec",
    "draft_spec_skeleton",
    "spec_from_draft",
    "spec_problems",
    "suggest_field_name",
]

DEFAULT_MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-5")


def claude_client() -> Any | None:
    """An Anthropic client, or None when the key or package is absent."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None
    try:
        import anthropic
    except ImportError:
        return None
    return anthropic.Anthropic(api_key=api_key)


# ---------------------------------------------------------------------------
# Deterministic skeleton


_NON_WORD_RE = re.compile(r"[^0-9a-zA-Z]+")


def suggest_field_name(header: str) -> str:
    """Deterministic CSV field-name suggestion for a template header.

    ``"Legal Name"`` -> ``"legal_name"``, ``"Employee_Reference"`` ->
    ``"employee_reference"``. Never empty for a non-empty header.
    """
    name = _NON_WORD_RE.sub("_", header.strip()).strip("_").lower()
    return name or "field"


def draft_spec_skeleton(template: TemplateSpec) -> dict[str, Any]:
    """A complete, deterministic load-spec draft for ``template``.

    Every headered non-key column of every data sheet is mapped to a
    suggested snake_case CSV field; keyed sheets get ``key_from`` pointing at
    a shared ``business_key`` field. The result always passes
    :func:`spec_problems` against its own template.
    """
    sheets: list[dict[str, Any]] = []
    for sheet in template.data_sheets():
        entry: dict[str, Any] = {
            "sheet": sheet.name,
            "source": f"{suggest_field_name(sheet.name)}.csv",
        }
        if sheet.key_column is not None:
            entry["key_from"] = "business_key"
        entry["columns"] = [
            {"column": col.header, "source": suggest_field_name(col.header)}
            for col in sheet.columns
            if col.header and not col.is_key
        ]
        sheets.append(entry)
    return {"key_strategy": "sequential", "sheets": sheets}


# ---------------------------------------------------------------------------
# Structural re-validation — the deterministic gate a Claude draft must pass


def spec_problems(raw: Any, template: TemplateSpec) -> list[str]:
    """Every structural reason ``raw`` cannot drive ``template``; [] when clean.

    Checks the spec's own shape (via the same gate ``load_spec`` uses) and
    then its fit against the template: sheets exist, columns exist and are
    not the spreadsheet key, nothing is mapped twice, required columns are
    all mapped, keyed sheets carry ``key_from``. CSV existence is *not*
    checked — a draft is a plan, not a run.
    """
    try:
        spec = spec_from_dict(raw, origin="draft")
    except GenerateError as exc:
        return [str(exc)]

    problems: list[str] = []
    for mapping in spec.sheets:
        sheet = template.sheet(mapping.sheet)
        if sheet is None or not sheet.columns:
            known = ", ".join(s.name for s in template.data_sheets())
            problems.append(f"no data sheet named {mapping.sheet!r} (data sheets: {known})")
            continue
        seen: set[int] = set()
        for col_map in mapping.columns:
            target = sheet.column_by_header(col_map.column)
            if target is None:
                problems.append(f"sheet {mapping.sheet!r}: no column {col_map.column!r}")
            elif target.is_key:
                problems.append(
                    f"sheet {mapping.sheet!r}: {target.header!r} is the spreadsheet key "
                    "and cannot be mapped"
                )
            elif target.index in seen:
                problems.append(f"sheet {mapping.sheet!r}: {target.header!r} is mapped twice")
            else:
                seen.add(target.index)
        unmapped = [
            c.header
            for c in sheet.columns
            if c.required and not c.is_key and c.index not in seen
        ]
        if unmapped:
            problems.append(
                f"sheet {mapping.sheet!r}: required column(s) not mapped: {', '.join(unmapped)}"
            )
        if sheet.key_column is not None and not mapping.key_from:
            problems.append(f"sheet {mapping.sheet!r}: keyed sheet needs 'key_from'")
    return problems


# ---------------------------------------------------------------------------
# Drafting


_DRAFT_SYSTEM = """You draft Workday EIB load specs for the eib-toolkit generator.
You receive a template summary, a deterministic skeleton spec, and the user's
instruction. Adapt the skeleton to the instruction: rename 'source' CSV fields
to what the user's data calls them, set constants via 'const' (a column takes
exactly one of 'source' or 'const'), drop optional columns the user does not
want, and adjust CSV filenames and 'key_from'. Never invent sheets or columns
that are not in the template summary, never map the spreadsheet-key column,
and keep every required column mapped. Respond with ONLY the JSON spec object
— no prose, no code fences."""


def draft_load_spec(
    template: TemplateSpec,
    instruction: str = "",
    client: Any | None = None,
    model: str = DEFAULT_MODEL,
) -> tuple[dict[str, Any], list[str]]:
    """Draft a load spec for ``template``; returns ``(spec_dict, notes)``.

    With no instruction or no usable Claude client, the deterministic
    skeleton is returned. With both, Claude adapts the skeleton and the
    result is re-validated via :func:`spec_problems`; any problem rejects
    the draft in favor of the skeleton. ``client`` is injectable for tests.
    """
    skeleton = draft_spec_skeleton(template)
    if not instruction:
        return skeleton, ["Deterministic skeleton (no instruction given)."]
    client = client or claude_client()
    if client is None:
        return skeleton, [
            "Claude is not available (no ANTHROPIC_API_KEY or 'anthropic' package); "
            "returned the deterministic skeleton instead."
        ]

    payload = {
        "template": _template_summary(template),
        "skeleton": skeleton,
        "instruction": instruction,
    }
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=4096,
            system=_DRAFT_SYSTEM,
            messages=[{"role": "user", "content": json.dumps(payload, indent=1)}],
        )
        raw = json.loads(_strip_fences(resp.content[0].text))
    except Exception as exc:  # network, refusal, unparseable output
        return skeleton, [
            f"Claude draft failed ({type(exc).__name__}); "
            "returned the deterministic skeleton instead."
        ]

    problems = spec_problems(raw, template)
    if problems:
        return skeleton, [
            "Claude draft rejected by structural validation; "
            "returned the deterministic skeleton instead.",
            *problems,
        ]
    return raw, ["Drafted by Claude from your instruction; review mappings before generating."]


def _template_summary(template: TemplateSpec) -> list[dict[str, Any]]:
    return [
        {
            "sheet": s.name,
            "role": s.role.value,
            "key_column": s.key_column.header if s.key_column else None,
            "columns": [
                {
                    "header": c.header,
                    "type": c.col_type.value,
                    "required": c.required,
                    "reference_id_type": c.ref_id_type or None,
                }
                for c in s.columns
                if c.header and not c.is_key
            ],
        }
        for s in template.data_sheets()
    ]


def _strip_fences(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    return text.strip()


# ---------------------------------------------------------------------------
# Annotation


_ANNOTATE_SYSTEM = """You are a Workday EIB load reviewer. You receive rule
codes from a deterministic validator with each rule's description, how often
it fired, and sample findings (message + evidence). For each code write ONE
concrete, imperative fix-it note (<= 35 words) a practitioner can act on in
Excel or their source data. Do not dispute the findings or their severity.
Respond with a JSON object mapping each code to its note, nothing else."""


def annotate_findings(
    report: ValidationReport,
    client: Any | None = None,
    model: str = DEFAULT_MODEL,
    samples_per_code: int = 3,
) -> dict[str, Any]:
    """Advisory notes for ``report``, one per rule code present.

    Returns ``{"annotated_by": "claude"|"none", "notes": {code: note}}``.
    The report itself is never modified — findings, severities, and counts
    stay exactly as the validator produced them. ``client`` is injectable
    for tests.
    """
    none: dict[str, Any] = {"annotated_by": "none", "notes": {}}
    if not report.findings:
        return none
    client = client or claude_client()
    if client is None:
        return none

    by_code: dict[str, list[dict[str, str]]] = {}
    for f in report.findings:
        samples = by_code.setdefault(f.code, [])
        if len(samples) < samples_per_code:
            samples.append({"location": f.location, "message": f.message, "evidence": f.evidence})
    payload = {
        code: {
            "rule": RULES[code][1],
            "severity": RULES[code][0].value,
            "count": report.by_code().get(code, 0),
            "samples": samples,
        }
        for code, samples in sorted(by_code.items())
    }
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=2048,
            system=_ANNOTATE_SYSTEM,
            messages=[{"role": "user", "content": json.dumps(payload, indent=1)}],
        )
        raw = json.loads(_strip_fences(resp.content[0].text))
    except Exception:  # network, refusal, unparseable output
        return none
    if not isinstance(raw, dict):
        return none
    notes = {
        code: str(note).strip()
        for code, note in raw.items()
        if code in by_code and str(note).strip()
    }
    return {"annotated_by": "claude", "notes": notes} if notes else none


# Re-exported for callers that draft and then immediately want a LoadSpec.
def spec_from_draft(raw: dict[str, Any], base_dir: str = ".") -> LoadSpec:
    """Turn a (validated) drafted spec dict into a runnable :class:`LoadSpec`."""
    return spec_from_dict(raw, base_dir=base_dir, origin="draft")
