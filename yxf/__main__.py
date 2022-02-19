import argparse
import collections
import grapefruit
import logging
import openpyxl
import openpyxl.styles
import openpyxl.utils
import pathlib
import strictyaml

GROUP_COLOR = grapefruit.Color.NewFromHtml("#c7ffdb")

log = logging.getLogger("yxf.__main__")


def _truncate_row(row):
    row = list(row)
    while row and row[-1] is None:
        row.pop()
    return row


def _convert_sheet(sheet):
    headers = None
    result = []
    for row in sheet.values:
        if headers is None:
            headers = _truncate_row(row)
            continue

        values = _truncate_row(row)
        row_dict = collections.OrderedDict()
        for h, v in zip(headers, values):
            if v is None:
                continue
            if h is None:
                raise ValueError("Cell with no column header: {}".format(v))
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


def _make_pretty_spreadsheet(wb):
    header_style = openpyxl.styles.NamedStyle(name="header")
    header_style.font = openpyxl.styles.Font(bold=True)

    code_style = openpyxl.styles.NamedStyle(name="code")
    code_style.font = openpyxl.styles.Font(name="Courier New", color="ff19007d")

    name_style = openpyxl.styles.NamedStyle(name="name")
    name_style.font = openpyxl.styles.Font(name="Courier New", color="ffa13b16")

    comment_style = openpyxl.styles.NamedStyle(name="comment")
    comment_style.font = openpyxl.styles.Font(name="Courier New", color="ff009c5d")

    note_style = openpyxl.styles.NamedStyle(name="note")
    note_style.font = openpyxl.styles.Font(color="ff555555")

    for sheet in wb:
        for cell in sheet[1]:
            cell.style = header_style
        sheet.freeze_panes = sheet["A2"]

        # Set column widths to reasonable values
        widths = [[] for _ in sheet[1]]
        headers = None
        for row in sheet:
            if headers is None:
                headers = [c.value for c in row]
                comment_column = headers.index("#") if "#" in headers else -1
                continue
            for i, cell in enumerate(row):
                if cell.value:
                    width = max(len(w) for w in cell.value.splitlines())
                    widths[i].append(width)
        # We take the 75th percentile width plus 10
        for i in range(len(widths)):
            col_widths = sorted(widths[i])
            num_rows = len(col_widths)
            percentile_index = num_rows * 3 // 4
            estimated_width = col_widths[percentile_index] + 10
            if i == comment_column:
                estimated_width = 2
            if estimated_width <= 60:
                sheet.column_dimensions[
                    openpyxl.utils.get_column_letter(i + 1)
                ].width = estimated_width
            else:
                sheet.column_dimensions[
                    openpyxl.utils.get_column_letter(i + 1)
                ].width = 60
                for row_index, _ in enumerate(sheet):
                    sheet.cell(
                        row=row_index + 1, column=i + 1
                    ).alignment = openpyxl.styles.Alignment(wrap_text=True)

        # Apply specific styles to known special columns or rows
        headers = None
        code_columns = set(
            ["calculation", "relevant", "constraint", "repeat_count", "instance_name"]
        )
        for row in sheet:
            if headers is None:
                headers = [c.value for c in row]
                type_column = headers.index("type") if "type" in headers else -1
                continue
            for i, cell in enumerate(row):
                if headers[i] in code_columns:
                    cell.style = code_style
                elif headers[i] == "name":
                    cell.style = name_style
                elif headers[i] == "#":
                    cell.style = comment_style
                elif type_column >= 0 and row[type_column].value == "note":
                    cell.style = note_style

        # Highlight groups and nesting
        group_number = 0
        color_stack = []
        headers = None
        for row in sheet:
            if headers is None:
                headers = [c.value for c in row]
                if "type" not in headers:
                    break
                type_column = headers.index("type")
                continue
            if str(row[type_column].value).startswith("begin_"):
                if not color_stack:
                    h, s, v = GROUP_COLOR.hsv
                    h += 20 * group_number
                    group_number += 1
                    color_stack.append(grapefruit.Color.NewFromHsv(h, s, v))
                else:
                    color_stack.append(color_stack[-1].DarkerColor(0.05))

            if color_stack:
                for cell in row[:1]:
                    cell.fill = openpyxl.styles.PatternFill(
                        fgColor="ff" + color_stack[-1].html[1:], fill_type="solid"
                    )

            if str(row[type_column].value).startswith("end_"):
                color_stack.pop()


def _check_existing_output(filename, force):
    if filename.exists() and not force:
        raise ValueError(
            "File already exists (use --force to override): {}".format(filename)
        )


def xlsform_to_yaml(filename: pathlib.Path, target: pathlib.Path):
    log.info("xlsform_to_yaml: %s -> %s", filename, target)

    wb = openpyxl.load_workbook(filename, read_only=True)
    result = collections.OrderedDict()
    for sheet_name in ["survey", "choices", "settings"]:
        if sheet_name in wb:
            result[sheet_name] = _convert_sheet(wb[sheet_name])

    first_line = result["survey"][0]
    if not "#" in first_line or not first_line["#"].startswith("Converted by yxf,"):
        first_line = {
            "#": "Converted by yxf, from {}. Edit the YAML file instead of the Excel file.".format(
                filename.name
            )
        }
        result["survey"].insert(0, first_line)
    elif first_line.get("#").startswith("Converted by yxf,"):
        first_line[
            "#"
        ] = "Converted by yxf, from {}. Edit the YAML file instead of the Excel file.".format(
            filename.name
        )

    with open(target, "w") as f:
        f.write(strictyaml.as_document(result).as_yaml())


def yaml_to_xlsform(filename: pathlib.Path, target: pathlib.Path):
    log.info("yaml_to_xlsform: %s -> %s", filename, target)

    with open(filename) as f:
        form = strictyaml.load(f.read()).data

    wb = openpyxl.Workbook()
    for sheet_name in form:
        _convert_to_sheet(wb.create_sheet(sheet_name), form[sheet_name])
    wb.remove(wb.active)
    _make_pretty_spreadsheet(wb)
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
        raise ValueError("Unrecognized file extension: {}".format(args.file))


if __name__ == "__main__":
    main()
