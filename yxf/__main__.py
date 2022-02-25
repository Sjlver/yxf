"""
yxf: Convert from XLSForm to YAML and back.

To convert an XLSForm to a YAML file: `python -m yxf form.xlsx`.

By default, the result will be called `form.yaml`, in other words, the same name
as the input file with the extension changed to `.yaml`. You can specify a
different output file name using the `--output othername.yaml` option.

To convert a YAML file to an XLSForm: `python -m yxf form.yaml`.
"""

import argparse
import collections
import logging
import pathlib
import re

import markdown_it
import markdown_it.tree
import openpyxl
import openpyxl.styles
import openpyxl.utils
import strictyaml

from . import xlsform

log = logging.getLogger("yxf.__main__")


def _row_to_dict(headers, values):
    row_dict = collections.OrderedDict()
    for h, v in zip(headers, values):
        if v is None:
            continue
        if h is None:
            raise ValueError(f"Cell with no column header: {v}")
        row_dict[h] = v
    return row_dict


def _convert_sheet(sheet):
    headers = xlsform.headers(sheet)
    result = []
    for row in xlsform.content_rows(sheet, values_only=True):
        values = xlsform.truncate_row(row)
        row_dict = _row_to_dict(headers, values)
        if row_dict:
            result.append(row_dict)
    return result


def _convert_to_sheet(sheet, rows, keys):
    key_set = set(keys)

    for i, key in enumerate(keys):
        sheet.cell(row=1, column=i + 1, value=key)

    next_row = 2
    for row in rows:
        if row.get("type") == "begin_group":
            next_row += 1

        if not all(k in key_set for k in row.keys()):
            missing_key = next(k for k in row.keys() if k not in key_set)
            raise ValueError(
                f'Invalid key "{missing_key}" in row "{row.get("name", "(unnamed)")}". '
                f"Add it to yxf.headers.{sheet.title} in the YAML file."
            )

        for i, key in enumerate(keys):
            if key in row:
                sheet.cell(row=next_row, column=i + 1, value=row[key])

        next_row += 1

    return sheet


def _write_to_xlsform(form, target):
    wb = openpyxl.Workbook()
    for sheet_name in form:
        if sheet_name == "yxf":
            continue
        _convert_to_sheet(
            wb.create_sheet(sheet_name),
            form[sheet_name],
            form["yxf"]["headers"][sheet_name],
        )
    wb.remove(wb.active)
    xlsform.make_pretty(wb)
    wb.save(target)


def _check_existing_output(filename, force):
    if filename.exists() and not force:
        raise ValueError(f"File already exists (use --force to override): {filename}")


def _ensure_yxf_comment(form, name, file_format):
    desired_comment = f"Converted by yxf, from {name}. Edit the {file_format} file instead of the Excel file."

    first_line = form["survey"][0]
    if "#" not in first_line or not first_line["#"].startswith("Converted by yxf,"):
        form["survey"].insert(0, {"#": desired_comment})
    else:
        form["survey"][0]["#"] = desired_comment

    if "#" not in form["yxf"]["headers"]["survey"]:
        form["yxf"]["headers"]["survey"].insert(0, "#")


def _validate_sheet_name(sheet_name, filename, line):
    if sheet_name not in ["survey", "choices", "settings"]:
        raise ValueError(
            f"{filename}:{line}: Invalid sheet name (must be survey, choices, or settings): {sheet_name}"
        )


def xlsform_to_yaml(filename: pathlib.Path, target: pathlib.Path):
    """Convert XLSForm file `filename` to YAML file `target`."""

    log.info("xlsform_to_yaml: %s -> %s", filename, target)

    wb = openpyxl.load_workbook(filename, read_only=True)
    result = collections.OrderedDict()
    headers = collections.OrderedDict()
    for sheet_name in ["survey", "choices", "settings"]:
        if sheet_name in wb:
            result[sheet_name] = _convert_sheet(wb[sheet_name])
            headers[sheet_name] = xlsform.headers(wb[sheet_name])
            if headers[sheet_name] and headers[sheet_name][0] != "#":
                if "#" in headers[sheet_name]:
                    raise ValueError(
                        f"The comment column must come first in sheet {sheet_name}."
                    )

    if "survey" not in result:
        raise ValueError('An XLSForm must have a "survey" sheet.')

    result["yxf"] = {"headers": headers}
    _ensure_yxf_comment(result, filename.name, "YAML")

    with open(target, "w", encoding="utf-8") as f:
        f.write(strictyaml.as_document(result).as_yaml())


def xlsform_to_markdown(filename: pathlib.Path, target: pathlib.Path):
    """Convert XLSForm file `filename` to Markdown file `target`."""

    log.info("xlsform_to_markdown: %s -> %s", filename, target)


def yaml_to_xlsform(filename: pathlib.Path, target: pathlib.Path):
    """Convert YAML file `filename` to XLSForm file `target`."""

    log.info("yaml_to_xlsform: %s -> %s", filename, target)

    with open(filename, encoding="utf-8") as f:
        form = strictyaml.load(f.read()).data

    if "yxf" not in form:
        raise ValueError('YAML file must have a "yxf" entry.')
    if "survey" not in form:
        raise ValueError('YAML file must have a "survey" entry.')
    _ensure_yxf_comment(form, filename.name, "YAML")
    _write_to_xlsform(form, target)


def markdown_to_xlsform(filename: pathlib.Path, target: pathlib.Path):
    """Convert Markdown file `filename` to XLSForm file `target`."""

    log.info("markdown_to_xlsform: %s -> %s", filename, target)

    with open(filename, encoding="utf-8") as f:
        md = f.read()

    parser = markdown_it.MarkdownIt("js-default")
    ast = markdown_it.tree.SyntaxTreeNode(parser.parse(md))
    form = collections.OrderedDict()
    form_headers = collections.OrderedDict()
    sheet_name = None
    for node in ast:
        if node.tag == "h2":
            sheet_name = node.children[0].content
            _validate_sheet_name(sheet_name, filename.name, node.map[0])
        elif node.tag == "p":
            content = node.children[0].content
            match = re.match(r"%%\s*(.*)", content)
            if match:
                sheet_name = match.group(1)
                _validate_sheet_name(sheet_name, filename.name, node.map[0])
        elif node.tag == "table":
            if not sheet_name:
                raise ValueError(
                    f"{filename.name}:{node.map[0]}: No sheet name specified for table."
                )
            thead, tbody = node.children
            headers = [c.children[0].content for c in thead.children[0].children]
            rows = tbody.children
            rows = [[c.children[0].content for c in row.children] for row in rows]
            result = []
            for values in rows:
                row_dict = _row_to_dict(headers, values)
                if row_dict:
                    result.append(row_dict)
            form[sheet_name] = result
            form_headers[sheet_name] = headers
    form["yxf"] = {"headers": form_headers}

    _ensure_yxf_comment(form, filename.name, "Markdown")
    _write_to_xlsform(form, target)


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
        if args.markdown:
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


if __name__ == "__main__":
    main()
