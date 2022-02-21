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

import openpyxl
import openpyxl.styles
import openpyxl.utils
import strictyaml

from . import xlsform

log = logging.getLogger("yxf.__main__")


def _convert_sheet(sheet):
    headers = xlsform.headers(sheet)
    result = []
    for row in xlsform.content_rows(sheet, values_only=True):
        values = xlsform.truncate_row(row)
        row_dict = collections.OrderedDict()
        for h, v in zip(headers, values):
            if v is None:
                continue
            if h is None:
                raise ValueError(f"Cell with no column header: {v}")
            row_dict[h] = v
        if row_dict:
            result.append(row_dict)
    return result


def _convert_to_sheet(sheet, rows):
    keys = []
    keys_set = set()
    for row in rows:
        for key in row:
            if key not in keys_set:
                keys.append(key)
                keys_set.add(key)

    for i, key in enumerate(keys):
        sheet.cell(row=1, column=i + 1, value=key)

    next_row = 2
    for row in rows:
        if row.get("type") == "begin_group":
            next_row += 1

        for i, key in enumerate(keys):
            if key in row:
                sheet.cell(row=next_row, column=i + 1, value=row[key])

        next_row += 1

    return sheet


def _check_existing_output(filename, force):
    if filename.exists() and not force:
        raise ValueError(f"File already exists (use --force to override): {filename}")


def xlsform_to_yaml(filename: pathlib.Path, target: pathlib.Path):
    """Convert XLSForm file `filename` to YAML file `target`."""

    log.info("xlsform_to_yaml: %s -> %s", filename, target)

    wb = openpyxl.load_workbook(filename, read_only=True)
    result = collections.OrderedDict()
    for sheet_name in ["survey", "choices", "settings"]:
        if sheet_name in wb:
            result[sheet_name] = _convert_sheet(wb[sheet_name])

    first_line = result["survey"][0]
    if not "#" in first_line or not first_line["#"].startswith("Converted by yxf,"):
        first_line = {
            "#": f"Converted by yxf, from {filename.name}. Edit the YAML file instead of the Excel file."
        }
        result["survey"].insert(0, first_line)
    elif first_line.get("#").startswith("Converted by yxf,"):
        first_line[
            "#"
        ] = f"Converted by yxf, from {filename.name}. Edit the YAML file instead of the Excel file."

    with open(target, "w", encoding="utf-8") as f:
        f.write(strictyaml.as_document(result).as_yaml())


def yaml_to_xlsform(filename: pathlib.Path, target: pathlib.Path):
    """Convert YAML file `filename` to XLSForm file `target`."""

    log.info("yaml_to_xlsform: %s -> %s", filename, target)

    with open(filename, encoding="utf-8") as f:
        form = strictyaml.load(f.read()).data

    wb = openpyxl.Workbook()
    for sheet_name in form:
        _convert_to_sheet(wb.create_sheet(sheet_name), form[sheet_name])
    wb.remove(wb.active)
    xlsform.make_pretty(wb)
    wb.save(target)


def main():
    logging.basicConfig(level=logging.DEBUG)

    parser = argparse.ArgumentParser(
        description="Convert from XLSForm to YAML and back"
    )
    parser.add_argument("file", type=pathlib.Path, help="a file to be converted")
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
        args.output = args.output or args.file.with_suffix(".yaml")
        _check_existing_output(args.output, args.force)
        xlsform_to_yaml(args.file, args.output)
    elif args.file.suffix == ".yaml":
        args.output = args.output or args.file.with_suffix(".xlsx")
        _check_existing_output(args.output, args.force)
        yaml_to_xlsform(args.file, args.output)
    else:
        raise ValueError(f"Unrecognized file extension: {args.file}")


if __name__ == "__main__":
    main()
