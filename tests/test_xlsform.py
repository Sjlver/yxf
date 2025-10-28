import pytest

from yxf.xlsform import row_to_dict, validate_sheet_name


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
