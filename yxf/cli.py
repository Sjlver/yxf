"""Command-line interface for yxf.

This module handles file I/O and CLI argument parsing, delegating the actual
conversion work to the library modules (excel, yaml, markdown).
"""

import argparse
import logging
import pathlib

from . import excel, markdown, xlsform, yaml

log = logging.getLogger(__name__)


def _check_existing_output(filename: pathlib.Path, force: bool) -> None:
    """Check if output file exists and raise error if it does (unless force=True).

    Args:
        filename: Path to output file
        force: If True, allow overwriting existing files

    Raises:
        ValueError: If file exists and force is False
    """
    if filename.exists() and not force:
        raise ValueError(f"File already exists (use --force to override): {filename}")


def read_xlsform_file(filename: pathlib.Path) -> dict:
    """Read an XLSForm file and return form dictionary.

    Args:
        filename: Path to input Excel file

    Returns:
        Form dictionary
    """
    with open(filename, "rb") as f:
        return excel.read_xlsform(f)


def read_yaml_file(filename: pathlib.Path) -> dict:
    """Read a YAML file and return form dictionary.

    Args:
        filename: Path to input YAML file

    Returns:
        Form dictionary
    """
    with open(filename, encoding="utf-8") as f:
        return yaml.read_yaml(f.read())


def read_markdown_file(filename: pathlib.Path) -> dict:
    """Read a Markdown file and return form dictionary.

    Args:
        filename: Path to input Markdown file

    Returns:
        Form dictionary
    """
    with open(filename, encoding="utf-8") as f:
        return markdown.read_markdown(f.read(), filename.name)


def write_xlsform_file(form: dict, target: pathlib.Path):
    """Write form dictionary to an XLSForm file.

    Args:
        form: Form dictionary
        target: Path to output Excel file
    """
    with open(target, "wb") as f:
        excel.write_xlsform(form, f)


def write_yaml_file(form: dict, target: pathlib.Path):
    """Write form dictionary to a YAML file.

    Args:
        form: Form dictionary
        target: Path to output YAML file
    """
    yaml_content = yaml.write_yaml(form)
    with open(target, "w", encoding="utf-8") as f:
        f.write(yaml_content)


def write_markdown_file(form: dict, target: pathlib.Path, source_name: str):
    """Write form dictionary to a Markdown file.

    Args:
        form: Form dictionary
        target: Path to output Markdown file
        source_name: Name of source file (for metadata)
    """
    md_content = markdown.write_markdown(form, source_name)
    with open(target, "w", encoding="utf-8") as f:
        f.write(md_content)


def main():
    """yxf: Convert from XLSForm to YAML and back."""

    logging.basicConfig(level=logging.DEBUG)
    logging.getLogger("markdown_it").setLevel(logging.INFO)

    parser = argparse.ArgumentParser(
        description="Convert from XLSForm to YAML and back"
    )
    parser.add_argument("file", type=pathlib.Path, help="a file to be converted")
    parser.add_argument(
        "--markdown",
        action="store_true",
        help="use Markdown instead of YAML",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=pathlib.Path,
        help="output file name (default: same as input, with extension changed)",
    )
    parser.add_argument(
        "-f",
        "--force",
        action="store_true",
        help="allow overwriting existing output files",
    )
    args = parser.parse_args()

    # Step 1: Determine input format
    if args.file.suffix not in [".xlsx", ".yaml", ".md"]:
        raise ValueError(f"Unrecognized file extension: {args.file}")
    input_format = args.file.suffix.lstrip(".")

    # Step 2: Determine output format
    if input_format == "xlsx":
        # XLSForm can go to either YAML or Markdown
        if args.markdown or (args.output and args.output.suffix == ".md"):
            output_format = "md"
        else:
            output_format = "yaml"
    else:
        # YAML and Markdown both go to XLSForm
        output_format = "xlsx"
    args.output = args.output or args.file.with_suffix(f".{output_format}")

    _check_existing_output(args.output, args.force)
    log.info("Converting: %s -> %s", args.file, args.output)

    # Step 3: Read input
    if input_format == "xlsx":
        form = read_xlsform_file(args.file)
    elif input_format == "yaml":
        form = read_yaml_file(args.file)
    elif input_format == "md":
        form = read_markdown_file(args.file)

    # Step 4: Add yxf comment
    if output_format == "xlsx":
        canonical_format = "YAML" if input_format == "yaml" else "Markdown"
    else:
        canonical_format = "YAML" if output_format == "yaml" else "Markdown"
    xlsform.ensure_yxf_comment(form, args.file.name, canonical_format)

    # Step 5: Write output
    if output_format == "xlsx":
        write_xlsform_file(form, args.output)
    elif output_format == "yaml":
        write_yaml_file(form, args.output)
    elif output_format == "md":
        write_markdown_file(form, args.output, args.file.name)
