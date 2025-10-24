"""Unit tests for YAML parsing and generation."""

import pytest
import strictyaml

from yxf.yaml import read_yaml, write_yaml


class TestReadYaml:
    """Tests for read_yaml function."""

    def test_missing_survey_key_raises_error(self):
        """Test that missing survey key raises ValueError."""
        yaml_content = """
yxf:
  headers:
    survey:
      - name
      - type
      - label
choices:
  - name: yes
"""
        with pytest.raises(ValueError, match='must have a "survey" entry'):
            read_yaml(yaml_content)

    def test_missing_yxf_key_raises_error(self):
        """Test that missing yxf key raises ValueError."""
        yaml_content = """
survey:
  - name: q1
    type: text
    label: Question 1
"""
        with pytest.raises(ValueError, match='must have a "yxf" entry'):
            read_yaml(yaml_content)

    def test_malformed_yaml_raises_error(self):
        """Test that malformed YAML raises an error."""
        yaml_content = """
survey:
  - name: q1
    type: text
  label: invalid indentation
"""
        with pytest.raises(strictyaml.YAMLError):
            read_yaml(yaml_content)

    def test_valid_yaml_parses_correctly(self):
        """Test that valid YAML parses correctly."""
        yaml_content = """
yxf:
  headers:
    survey:
      - name
      - type
      - label
survey:
  - name: q1
    type: text
    label: Question 1
"""
        form = read_yaml(yaml_content)
        assert "survey" in form
        assert "yxf" in form
        assert len(form["survey"]) == 1
        assert form["survey"][0]["name"] == "q1"


class TestWriteYaml:
    """Tests for write_yaml function."""

    def test_write_minimal_form(self):
        """Test writing a minimal form to YAML."""
        form = {
            "yxf": {"headers": {"survey": ["name", "type", "label"]}},
            "survey": [{"name": "q1", "type": "text", "label": "Question 1"}],
        }
        yaml_output = write_yaml(form)

        # Verify it's valid YAML that can be parsed back
        form_parsed = read_yaml(yaml_output)
        assert form_parsed["survey"][0]["name"] == "q1"

    def test_write_preserves_order(self):
        """Test that write_yaml preserves key order."""
        form = {
            "yxf": {"headers": {"survey": ["name", "type"]}},
            "survey": [{"name": "q1", "type": "text"}],
        }
        yaml_output = write_yaml(form)

        # Check that yxf comes before survey in output
        lines = yaml_output.split("\n")
        yxf_line = next(i for i, line in enumerate(lines) if line.startswith("yxf:"))
        survey_line = next(
            i for i, line in enumerate(lines) if line.startswith("survey:")
        )
        assert yxf_line < survey_line

    def test_roundtrip_preserves_data(self):
        """Test that reading and writing YAML preserves data."""
        yaml_content = """
yxf:
  headers:
    survey:
      - name
      - type
      - label
    choices:
      - list_name
      - name
      - label
survey:
  - name: color
    type: select_one colors
    label: What is your favorite color?
choices:
  - list_name: colors
    name: red
    label: Red
  - list_name: colors
    name: blue
    label: Blue
"""
        form = read_yaml(yaml_content)
        yaml_output = write_yaml(form)
        form_roundtrip = read_yaml(yaml_output)

        # Verify key data is preserved
        assert form_roundtrip["survey"][0]["name"] == "color"
        assert len(form_roundtrip["choices"]) == 2
