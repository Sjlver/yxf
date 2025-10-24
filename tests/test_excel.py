"""Unit tests for Excel-specific logic."""

import io

import pytest
import openpyxl

from yxf.excel import row_to_dict, validate_sheet_name, read_xlsform, write_xlsform


class TestRowToDict:
    """Tests for row_to_dict function."""

    def test_basic_conversion(self):
        """Test basic row to dict conversion."""
        headers = ["name", "type", "label"]
        values = ["q1", "text", "Question 1"]
        result = row_to_dict(headers, values)
        assert result == {"name": "q1", "type": "text", "label": "Question 1"}

    def test_empty_values_skipped(self):
        """Test that empty values are skipped."""
        headers = ["name", "type", "label"]
        values = ["q1", "", None]
        result = row_to_dict(headers, values)
        assert result == {"name": "q1"}

    def test_value_without_header_raises_error(self):
        """Test that a value without a header raises ValueError."""
        headers = ["name", None, "label"]
        values = ["q1", "text", "Question 1"]
        with pytest.raises(ValueError, match="Cell with no column header"):
            row_to_dict(headers, values)

    def test_mismatched_lengths(self):
        """Test handling when headers and values have different lengths."""
        headers = ["name", "type", "label"]
        values = ["q1", "text"]  # Shorter than headers
        result = row_to_dict(headers, values)
        assert result == {"name": "q1", "type": "text"}


class TestValidateSheetName:
    """Tests for validate_sheet_name function."""

    def test_valid_sheet_names(self):
        """Test that valid sheet names pass validation."""
        validate_sheet_name("survey", "test.yaml", 1)
        validate_sheet_name("choices", "test.yaml", 1)
        validate_sheet_name("settings", "test.yaml", 1)

    def test_invalid_sheet_name_raises_error(self):
        """Test that invalid sheet names raise ValueError."""
        with pytest.raises(ValueError, match="Invalid sheet name"):
            validate_sheet_name("invalid", "test.yaml", 1)

    def test_error_includes_location(self):
        """Test that error message includes source and line number."""
        with pytest.raises(ValueError, match="test.yaml:42"):
            validate_sheet_name("invalid", "test.yaml", 42)


class TestReadXlsform:
    """Tests for read_xlsform function."""

    def test_missing_survey_sheet_raises_error(self):
        """Test that missing survey sheet raises ValueError."""
        # Create a workbook without a survey sheet (but with valid content)
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "choices"
        sheet.append(["list_name", "name", "label"])
        sheet.append(["colors", "red", "Red"])
        excel_bytes = io.BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        with pytest.raises(ValueError, match='must have a "survey" sheet'):
            read_xlsform(excel_bytes)

    def test_comment_column_not_first_raises_error(self):
        """Test that comment column not in first position raises error."""
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "survey"
        sheet.append(["type", "#", "name"])  # Comment column not first
        sheet.append(["text", "A comment", "q1"])
        excel_bytes = io.BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        with pytest.raises(ValueError, match="comment column must come first"):
            read_xlsform(excel_bytes)


class TestWriteXlsform:
    """Tests for write_xlsform function."""

    def test_invalid_key_in_row_raises_error(self):
        """Test that a row with an invalid key raises ValueError."""
        form = {
            "survey": [{"name": "q1", "invalid_key": "value"}],
            "yxf": {"headers": {"survey": ["name", "type"]}},
        }

        excel_bytes = io.BytesIO()
        with pytest.raises(ValueError, match='Invalid key "invalid_key"'):
            write_xlsform(form, excel_bytes)

    def test_minimal_form_writes_successfully(self):
        """Test that a minimal valid form can be written."""
        form = {
            "survey": [{"name": "q1", "type": "text", "label": "Question 1"}],
            "yxf": {"headers": {"survey": ["name", "type", "label"]}},
        }

        excel_bytes = io.BytesIO()
        write_xlsform(form, excel_bytes)
        excel_bytes.seek(0)

        # Verify it can be read back
        form_read = read_xlsform(excel_bytes)
        assert "survey" in form_read
        assert len(form_read["survey"]) == 1
        assert form_read["survey"][0]["name"] == "q1"
