"""Snapshot tests for Excel ↔ YAML ↔ Markdown conversions.

These tests use syrupy for snapshot testing to ensure conversion outputs
remain stable. Only YAML and Markdown outputs are snapshotted (not Excel
files, as openpyxl output is not byte-stable).
"""

import io
import pathlib

import pytest

from yxf import (
    read_xlsform,
    write_xlsform,
    read_yaml,
    write_yaml,
    read_markdown,
    write_markdown,
)


# Test data directory
TESTDATA_DIR = pathlib.Path(__file__).parent / "testdata"


@pytest.fixture(
    params=[
        "favorite-color.xlsx",
        "simple-repeat.xlsx",
        "cascading-select.xlsx",
        "xlsform-dot-org-template.xlsx",
    ]
)
def excel_file(request):
    """Fixture providing all test Excel files."""
    return TESTDATA_DIR / request.param


def test_excel_to_yaml_snapshot(excel_file, snapshot):
    """Test Excel → YAML conversion produces stable output."""
    with open(excel_file, "rb") as f:
        form = read_xlsform(f)

    # Skip xlsform.org template - it has empty sheets that strictyaml can't serialize
    if "xlsform-dot-org-template" in excel_file.name:
        pytest.skip(
            "xlsform.org template has empty sheets that strictyaml cannot serialize"
        )

    yaml_output = write_yaml(form)
    assert yaml_output == snapshot


def test_excel_to_markdown_snapshot(excel_file, snapshot):
    """Test Excel → Markdown conversion produces stable output."""
    with open(excel_file, "rb") as f:
        form = read_xlsform(f)
    markdown_output = write_markdown(form, excel_file.name)
    assert markdown_output == snapshot


def test_yaml_roundtrip_stability(excel_file):
    """Test Excel → YAML → Excel → YAML produces identical YAML.

    This verifies that converting through Excel doesn't lose or alter data.
    """
    # Skip xlsform.org template - it has empty sheets that strictyaml can't serialize
    if "xlsform-dot-org-template" in excel_file.name:
        pytest.skip(
            "xlsform.org template has empty sheets that strictyaml cannot serialize"
        )

    # Excel → YAML
    with open(excel_file, "rb") as f:
        form1 = read_xlsform(f)
    yaml1 = write_yaml(form1)

    # YAML → Excel (in-memory) → YAML
    form2 = read_yaml(yaml1)
    excel_bytes = io.BytesIO()
    write_xlsform(form2, excel_bytes)
    excel_bytes.seek(0)
    form3 = read_xlsform(excel_bytes)
    yaml2 = write_yaml(form3)

    assert yaml1 == yaml2, "YAML should be identical after round-trip through Excel"


def test_markdown_roundtrip_via_yaml(excel_file):
    """Test Excel → Markdown → Excel → YAML produces same result as Excel → YAML.

    This verifies that Markdown conversion preserves all data needed for YAML.
    Note: We compare via YAML since that's our canonical format.
    """
    # Skip xlsform.org template - it has empty sheets that strictyaml can't serialize
    if "xlsform-dot-org-template" in excel_file.name:
        pytest.skip(
            "xlsform.org template has empty sheets that strictyaml cannot serialize"
        )

    # Excel → YAML (baseline)
    with open(excel_file, "rb") as f:
        form1 = read_xlsform(f)
    yaml_baseline = write_yaml(form1)

    # Excel → Markdown → Excel → YAML
    markdown = write_markdown(form1, excel_file.name)
    form2 = read_markdown(markdown, excel_file.name)
    excel_bytes = io.BytesIO()
    write_xlsform(form2, excel_bytes)
    excel_bytes.seek(0)
    form3 = read_xlsform(excel_bytes)
    yaml_after_markdown = write_yaml(form3)

    # Compare the YAMLs (they should be very similar, though comments may differ)
    assert yaml_baseline == yaml_after_markdown

    # We check that the structure is preserved
    assert "survey" in yaml_after_markdown
    assert "yxf" in yaml_after_markdown
