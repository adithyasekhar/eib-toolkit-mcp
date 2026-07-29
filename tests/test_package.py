"""Scaffold sanity checks — replaced by real suites as slices land (docs/PLAN.md)."""

import eib_toolkit


def test_version_is_semver() -> None:
    parts = eib_toolkit.__version__.split(".")
    assert len(parts) == 3
    assert all(p.isdigit() for p in parts)


def test_openpyxl_dependency_importable() -> None:
    import openpyxl  # noqa: F401  # core dependency for all workbook I/O
