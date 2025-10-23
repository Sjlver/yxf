"""Command-line interface for yxf.

This module handles file I/O and CLI argument parsing, delegating the actual
conversion work to the library modules (excel, yaml, markdown).
"""

import argparse
import logging
import pathlib

from . import excel, markdown, yaml

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


def xlsform_to_yaml(filename: pathlib.Path, target: pathlib.Path):
    """Convert XLSForm file to YAML file.

    Args:
        filename: Path to input Excel file
        target: Path to output YAML file
    """
    log.info("xlsform_to_yaml: %s -> %s", filename, target)

    with open(filename, "rb") as f:
        form = excel.read_xlsform(f)

    excel.ensure_yxf_comment(form, filename.name, "YAML")
    yaml_content = yaml.write_yaml(form)

    with open(target, "w", encoding="utf-8") as f:
        f.write(yaml_content)


def xlsform_to_markdown(filename: pathlib.Path, target: pathlib.Path):
    """Convert XLSForm file to Markdown file.

    Args:
        filename: Path to input Excel file
        target: Path to output Markdown file
    """
    log.info("xlsform_to_markdown: %s -> %s", filename, target)

    with open(filename, "rb") as f:
        form = excel.read_xlsform(f)

    excel.ensure_yxf_comment(form, filename.name, "Markdown")
    md_content = markdown.write_markdown(form, filename.name)

    with open(target, "w", encoding="utf-8") as f:
        f.write(md_content)


def yaml_to_xlsform(filename: pathlib.Path, target: pathlib.Path):
    """Convert YAML file to XLSForm file.

    Args:
        filename: Path to input YAML file
        target: Path to output Excel file
    """
    log.info("yaml_to_xlsform: %s -> %s", filename, target)

    with open(filename, encoding="utf-8") as f:
        form = yaml.read_yaml(f.read())

    excel.ensure_yxf_comment(form, filename.name, "YAML")

    with open(target, "wb") as f:
        excel.write_xlsform(form, f)


def markdown_to_xlsform(filename: pathlib.Path, target: pathlib.Path):
    """Convert Markdown file to XLSForm file.

    Args:
        filename: Path to input Markdown file
        target: Path to output Excel file
    """
    log.info("markdown_to_xlsform: %s -> %s", filename, target)

    with open(filename, encoding="utf-8") as f:
        form = markdown.read_markdown(f.read(), filename.name)

    excel.ensure_yxf_comment(form, filename.name, "Markdown")

    with open(target, "wb") as f:
        excel.write_xlsform(form, f)


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

    if args.file.suffix == ".xlsx":
        if args.markdown or (args.output and args.output.suffix == ".md"):
            args.output = args.output or args.file.with_suffix(".md")
            _check_existing_output(args.output, args.force)
            xlsform_to_markdown(args.file, args.output)
        else:
            args.output = args.output or args.file.with_suffix(".yaml")
            _check_existing_output(args.output, args.force)
            xlsform_to_yaml(args.file, args.output)
    elif args.file.suffix == ".yaml":
        args.output = args.output or args.file.with_suffix(".xlsx")
        _check_existing_output(args.output, args.force)
        yaml_to_xlsform(args.file, args.output)
    elif args.file.suffix == ".md":
        args.output = args.output or args.file.with_suffix(".xlsx")
        _check_existing_output(args.output, args.force)
        markdown_to_xlsform(args.file, args.output)
    else:
        raise ValueError(f"Unrecognized file extension: {args.file}")
