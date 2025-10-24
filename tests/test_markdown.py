"""Unit tests for Markdown parsing and generation."""

import pytest

from yxf.markdown import read_markdown, write_markdown


class TestReadMarkdown:
    """Tests for read_markdown function."""

    def test_invalid_sheet_name_raises_error(self):
        """Test that invalid sheet name raises ValueError."""
        md_content = """
## invalid_sheet

| type | name | label |
| ---- | ---- | ----- |
| text | q1   | Q1    |
"""
        with pytest.raises(ValueError, match="Invalid sheet name"):
            read_markdown(md_content, "test.md")

    def test_table_without_sheet_name_raises_error(self):
        """Test that table without sheet name raises ValueError."""
        md_content = """
| type | name | label |
| ---- | ---- | ----- |
| text | q1   | Q1    |
"""
        with pytest.raises(ValueError, match="No sheet name specified"):
            read_markdown(md_content, "test.md")

    def test_basic_markdown_parsing(self):
        """Test parsing basic markdown with survey sheet."""
        md_content = """
## survey

| type               | name | label      |
| ------------------ | ---- | ---------- |
| text               | q1   | Question 1 |
| integer            | q2   | Question 2 |
| select_one choices | q3   | Question 3 |
"""
        form = read_markdown(md_content, "test.md")
        assert "survey" in form
        assert len(form["survey"]) == 3

        # Check first row
        assert form["survey"][0]["type"] == "text"
        assert form["survey"][0]["name"] == "q1"
        assert form["survey"][0]["label"] == "Question 1"

        # Check second row
        assert form["survey"][1]["type"] == "integer"
        assert form["survey"][1]["name"] == "q2"
        assert form["survey"][1]["label"] == "Question 2"

        # Check third row
        assert form["survey"][2]["type"] == "select_one choices"
        assert form["survey"][2]["name"] == "q3"
        assert form["survey"][2]["label"] == "Question 3"

    def test_comment_paragraphs_added_to_rows(self):
        """Test that paragraphs are treated as comments."""
        md_content = """
## survey

This is a comment paragraph.

| type | name | label      |
| ---- | ---- | ---------- |
| text | q1   | Question 1 |
| text | q2   | Question 2 |
"""
        form = read_markdown(md_content, "test.md")
        # Comments should be added to the beginning of the sheet
        assert form["survey"][0]["#"] == "This is a comment paragraph."
        assert form["survey"][1]["name"] == "q1"
        assert form["survey"][1]["label"] == "Question 1"
        assert form["survey"][2]["name"] == "q2"
        assert form["survey"][2]["label"] == "Question 2"

    def test_escaped_pipe_in_markdown(self):
        """Test that escaped pipes in markdown are handled."""
        md_content = """
## survey

| type | name | label               |
| ---- | ---- | ------------------- |
| text | q1   | Question with \\| pipe |
"""
        form = read_markdown(md_content, "test.md")
        # The pipe should be unescaped when parsing
        assert form["survey"][0]["label"] == "Question with | pipe"

    def test_multiple_rows_with_empty_cells(self):
        """Test parsing multiple rows where some cells are empty."""
        md_content = """
## survey

| type                | name   | label | required |
| ------------------- | ------ | ----- | -------- |
| text                | q1     | Q1    | yes      |
| select_one choices  | q2     | Q2    |          |
| integer             | q3     | Q3    | yes      |
"""
        form = read_markdown(md_content, "test.md")
        assert len(form["survey"]) == 3

        assert form["survey"][0] == {
            "type": "text",
            "name": "q1",
            "label": "Q1",
            "required": "yes",
        }
        assert form["survey"][1] == {
            "type": "select_one choices",
            "name": "q2",
            "label": "Q2",
        }
        assert form["survey"][2] == {
            "type": "integer",
            "name": "q3",
            "label": "Q3",
            "required": "yes",
        }


class TestWriteMarkdown:
    """Tests for write_markdown function."""

    def test_write_basic_form(self):
        """Test writing a basic form to markdown."""
        form = {
            "survey": [{"type": "text", "name": "q1", "label": "Question 1"}],
            "yxf": {"headers": {"survey": ["type", "name", "label"]}},
        }
        md_output = write_markdown(form)

        assert "## survey" in md_output
        assert "| type | name | label" in md_output
        assert "| text | q1   | Question 1" in md_output

    def test_write_removes_comment_column(self):
        """Test that comment column is removed from markdown tables."""
        form = {
            "survey": [{"#": "A comment", "type": "text", "name": "q1"}],
            "yxf": {"headers": {"survey": ["#", "type", "name"]}},
        }
        md_output = write_markdown(form)

        # Comment should appear as paragraph, not in table
        assert "\nA comment\n\n" in md_output
        # Comment column should not be in table header
        assert "| type | name |" in md_output
        assert "| # |" not in md_output

    def test_write_escapes_pipes(self):
        """Test that pipes are escaped in markdown tables."""
        form = {
            "survey": [{"type": "text", "name": "q1", "label": "A | B"}],
            "yxf": {"headers": {"survey": ["type", "name", "label"]}},
        }
        md_output = write_markdown(form)

        # Pipe should be escaped
        assert "A \\| B" in md_output

    def test_write_escapes_backslashes(self):
        """Test that backslashes are escaped in markdown tables."""
        form = {
            "survey": [{"type": "text", "name": "q1", "label": "A \\ B"}],
            "yxf": {"headers": {"survey": ["type", "name", "label"]}},
        }
        md_output = write_markdown(form)

        # Backslash should be escaped
        assert "A \\\\ B" in md_output

    def test_multiline_warning(self, caplog):
        """Test that multiline values generate a warning."""
        form = {
            "survey": [{"type": "text", "name": "q1", "label": "Line 1\nLine 2"}],
            "yxf": {"headers": {"survey": ["type", "name", "label"]}},
        }
        md_output = write_markdown(form)

        # Should generate warning and replace newline with space
        assert "Line 1 Line 2" in md_output

        # Check that a warning was logged
        assert len(caplog.records) == 1
        assert caplog.records[0].levelname == "WARNING"
        assert "Multi-line value for column label" in caplog.records[0].message
        assert (
            "Markdown does not support multi-line values" in caplog.records[0].message
        )
