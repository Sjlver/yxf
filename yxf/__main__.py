import argparse
import logging
import pathlib
import openpyxl
import openpyxl.styles
import strictyaml

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
        row_dict = {}
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
    for sheet in wb:
        for cell in sheet[1]:
            cell.style = header_style
        sheet.freeze_panes = sheet["A2"]

        # TODO(Jonas): set column widths and overflow
        # TODO(Jonas): format type and name column with a fixed-width font
        # TODO(Jonas): add support for comments


def xlsform_to_yaml(filename: pathlib.Path):
    target_filename = filename.with_suffix(".yaml")
    log.info("xlsform_to_yaml: %s -> %s", filename, target_filename)

    wb = openpyxl.load_workbook(filename, read_only=True)
    result = {}
    for sheet_name in ["survey", "choices", "settings"]:
        if sheet_name in wb:
            result[sheet_name] = _convert_sheet(wb[sheet_name])

    with open(target_filename, "w") as f:
        f.write(strictyaml.as_document(result).as_yaml())


def yaml_to_xlsform(filename: pathlib.Path):
    target_filename = filename.with_suffix(".xlsx")
    log.info("yaml_to_xlsform: %s -> %s", filename, target_filename)

    with open(filename) as f:
        form = strictyaml.load(f.read()).data

    wb = openpyxl.Workbook()
    for sheet_name in form:
        _convert_to_sheet(wb.create_sheet(sheet_name), form[sheet_name])
    wb.remove(wb.active)
    _make_pretty_spreadsheet(wb)
    wb.save(target_filename)


def main():
    logging.basicConfig(level=logging.DEBUG)

    parser = argparse.ArgumentParser(
        description="Convert from XLSForm to YAML and back"
    )
    parser.add_argument(
        "files", metavar="file", nargs="+", type=str, help="a file to be converted"
    )
    args = parser.parse_args()

    for filename in args.files:
        filename = pathlib.Path(filename)
        if filename.suffix == ".xlsx":
            xlsform_to_yaml(filename)
        elif filename.suffix == ".yaml":
            yaml_to_xlsform(filename)
        else:
            raise ValueError("Unrecognized file extension: {}".format(filename))


if __name__ == "__main__":
    main()
