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


@pytest.fixture(
    params=[
        "fruit-nested-groups.yaml",
        "fruit-html-labels.yaml",
    ]
)
def yaml_file(request):
    """Fixture providing the YAML forms with a complex group structure."""
    return TESTDATA_DIR / request.param


def test_excel_to_yaml_snapshot(excel_file, snapshot):
    """Test Excel → YAML conversion produces stable output."""
    with open(excel_file, "rb") as f:
        form = read_xlsform(f)

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


def test_yaml_form_roundtrip_stability(yaml_file):
    """Test YAML → Excel → YAML on forms with a complex group structure.

    These forms nest groups and repeats several levels deep, mix all the
    spellings of the group markers, and contain "type: end" metadata rows.
    """
    yaml1 = yaml_file.read_text(encoding="utf-8")

    excel_bytes = io.BytesIO()
    write_xlsform(read_yaml(yaml1), excel_bytes)
    excel_bytes.seek(0)
    yaml2 = write_yaml(read_xlsform(excel_bytes))

    assert yaml1 == yaml2, "YAML should be identical after round-trip through Excel"


def test_yaml_form_group_structure_is_preserved(yaml_file):
    """Test that the nesting of groups survives a round-trip through Excel."""
    from yxf.xlsform import nesting_levels

    form1 = read_yaml(yaml_file.read_text(encoding="utf-8"))
    levels1 = nesting_levels([row.get("type", "") for row in form1["survey"]])
    assert max(levels1) >= 3, "the fixture should nest at least three levels deep"

    excel_bytes = io.BytesIO()
    write_xlsform(form1, excel_bytes)
    excel_bytes.seek(0)
    form2 = read_xlsform(excel_bytes)
    levels2 = nesting_levels([row.get("type", "") for row in form2["survey"]])

    assert levels1 == levels2


def test_yaml_form_markdown_preserves_structure(yaml_file):
    """Test that Markdown keeps the question structure of the complex forms.

    Markdown is a lossy format for these forms: it cannot represent multi-line
    values (yxf warns about that), and a comment column with no comments in it
    is not restored when reading back. The order, types and names of the
    questions must survive regardless, since that is what defines the form.
    """
    form1 = read_yaml(yaml_file.read_text(encoding="utf-8"))
    structure = [(row.get("type"), row.get("name")) for row in form1["survey"]]

    markdown = write_markdown(form1, yaml_file.name)
    form2 = read_markdown(markdown, yaml_file.name)

    assert [(row.get("type"), row.get("name")) for row in form2["survey"]] == structure


def test_write_markdown_does_not_modify_the_form(yaml_file):
    """Test that writing Markdown leaves the caller's form untouched.

    write_markdown used to delete the comment column from the form it was
    given, so a caller that wrote Markdown and then YAML lost their comments.
    """
    form = read_yaml(yaml_file.read_text(encoding="utf-8"))
    before = write_yaml(read_yaml(yaml_file.read_text(encoding="utf-8")))

    write_markdown(form, yaml_file.name)

    assert write_yaml(form) == before
