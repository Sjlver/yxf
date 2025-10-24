"""Unit tests for CLI argument parsing and file handling."""

import pytest

from yxf import cli, write_xlsform


# Sample data for tests
SAMPLE_YAML_CONTENT = """
yxf:
  headers:
    survey:
      - type
      - name
survey:
  - type: text
    name: q1
"""

SAMPLE_MARKDOWN_CONTENT = """
## survey

| type | name |
| ---- | ---- |
| text | q1   |
"""

SAMPLE_FORM = {
    "survey": [{"type": "text", "name": "q1"}],
    "yxf": {"headers": {"survey": ["type", "name"]}},
}


# Helper functions for test file creation
def create_xlsx_file(path):
    """Create a minimal Excel file at the given path."""
    with open(path, "wb") as f:
        write_xlsform(SAMPLE_FORM, f)


def create_yaml_file(path):
    """Create a minimal YAML file at the given path."""
    path.write_text(SAMPLE_YAML_CONTENT)


def create_markdown_file(path):
    """Create a minimal Markdown file at the given path."""
    path.write_text(SAMPLE_MARKDOWN_CONTENT)


class TestFileExtensionDetection:
    """Tests for automatic file type detection based on extension."""

    def test_xlsx_to_yaml_by_default(self, tmp_path, monkeypatch):
        """Test that .xlsx converts to .yaml by default."""
        test_xlsx = tmp_path / "test.xlsx"
        test_yaml = tmp_path / "test.yaml"

        create_xlsx_file(test_xlsx)

        # Run CLI
        monkeypatch.setattr("sys.argv", ["yxf", str(test_xlsx)])
        cli.main()

        assert test_yaml.exists()

    def test_yaml_to_xlsx(self, tmp_path, monkeypatch):
        """Test that .yaml converts to .xlsx."""
        test_yaml = tmp_path / "test.yaml"
        test_xlsx = tmp_path / "test.xlsx"

        create_yaml_file(test_yaml)

        # Run CLI
        monkeypatch.setattr("sys.argv", ["yxf", str(test_yaml)])
        cli.main()

        assert test_xlsx.exists()

    def test_markdown_to_xlsx(self, tmp_path, monkeypatch):
        """Test that .md converts to .xlsx."""
        test_md = tmp_path / "test.md"
        test_xlsx = tmp_path / "test.xlsx"

        create_markdown_file(test_md)

        # Run CLI
        monkeypatch.setattr("sys.argv", ["yxf", str(test_md)])
        cli.main()

        assert test_xlsx.exists()

    def test_xlsx_to_markdown_with_flag(self, tmp_path, monkeypatch):
        """Test that --markdown flag converts .xlsx to .md."""
        test_xlsx = tmp_path / "test.xlsx"
        test_md = tmp_path / "test.md"

        create_xlsx_file(test_xlsx)

        # Run CLI with --markdown
        monkeypatch.setattr("sys.argv", ["yxf", "--markdown", str(test_xlsx)])
        cli.main()

        assert test_md.exists()


class TestOutputFileNameGeneration:
    """Tests for -o/--output option."""

    def test_custom_output_name(self, tmp_path, monkeypatch):
        """Test specifying custom output file name."""
        test_yaml = tmp_path / "input.yaml"
        custom_xlsx = tmp_path / "custom_output.xlsx"

        create_yaml_file(test_yaml)

        # Run CLI with custom output
        monkeypatch.setattr("sys.argv", ["yxf", "-o", str(custom_xlsx), str(test_yaml)])
        cli.main()

        assert custom_xlsx.exists()

    def test_output_extension_overrides_markdown_flag(self, tmp_path, monkeypatch):
        """Test that explicit .md output extension works even without --markdown."""
        test_xlsx = tmp_path / "test.xlsx"
        test_md = tmp_path / "output.md"

        create_xlsx_file(test_xlsx)

        # Run CLI with .md output (no --markdown flag needed)
        monkeypatch.setattr("sys.argv", ["yxf", "-o", str(test_md), str(test_xlsx)])
        cli.main()

        assert test_md.exists()


class TestForceOverwrite:
    """Tests for -f/--force option."""

    def test_existing_file_without_force_raises_error(self, tmp_path, monkeypatch):
        """Test that existing output file raises error without --force."""
        test_yaml = tmp_path / "test.yaml"
        test_xlsx = tmp_path / "test.xlsx"

        create_yaml_file(test_yaml)
        test_xlsx.write_text("existing content")

        # Run CLI without --force
        monkeypatch.setattr("sys.argv", ["yxf", str(test_yaml)])

        with pytest.raises(ValueError, match="already exists"):
            cli.main()

    def test_force_overwrites_existing_file(self, tmp_path, monkeypatch):
        """Test that --force allows overwriting existing files."""
        test_yaml = tmp_path / "test.yaml"
        test_xlsx = tmp_path / "test.xlsx"

        create_yaml_file(test_yaml)
        test_xlsx.write_text("existing content")

        # Run CLI with --force
        monkeypatch.setattr("sys.argv", ["yxf", "--force", str(test_yaml)])
        cli.main()

        # Output file should be overwritten (and be a valid Excel file)
        assert test_xlsx.exists()
        # The content should have changed (no longer "existing content")
        content = test_xlsx.read_bytes()
        assert b"existing content" not in content


class TestInvalidInput:
    """Tests for error handling with invalid inputs."""

    def test_unrecognized_extension_raises_error(self, tmp_path, monkeypatch):
        """Test that unrecognized file extension raises error."""
        test_file = tmp_path / "test.txt"
        test_file.write_text("some content")

        monkeypatch.setattr("sys.argv", ["yxf", str(test_file)])

        with pytest.raises(ValueError, match="Unrecognized file extension"):
            cli.main()


class TestCheckExistingOutput:
    """Tests for _check_existing_output helper function."""

    def test_nonexistent_file_passes(self, tmp_path):
        """Test that nonexistent file passes check."""
        nonexistent = tmp_path / "does_not_exist.xlsx"
        cli._check_existing_output(nonexistent, force=False)  # Should not raise

    def test_existing_file_without_force_raises(self, tmp_path):
        """Test that existing file without force raises error."""
        existing = tmp_path / "exists.xlsx"
        existing.write_text("content")

        with pytest.raises(ValueError, match="already exists"):
            cli._check_existing_output(existing, force=False)

    def test_existing_file_with_force_passes(self, tmp_path):
        """Test that existing file with force passes check."""
        existing = tmp_path / "exists.xlsx"
        existing.write_text("content")

        cli._check_existing_output(existing, force=True)  # Should not raise
