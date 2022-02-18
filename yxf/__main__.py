import argparse
import logging
import pathlib
import openpyxl
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
            log.debug("headers: %s", headers)
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


def xlsform_to_yaml(filename: pathlib.Path):
    target_filename = filename.with_suffix(".yaml")
    log.info("xlsform_to_yaml: %s -> %s", filename, target_filename)

    wb = openpyxl.load_workbook(filename, read_only=True)
    log.debug("Workbook: %s", wb)
    result = {}
    for sheet_name in ["survey", "choices", "settings"]:
        if sheet_name in wb:
            result[sheet_name] = _convert_sheet(wb[sheet_name])

    with open(target_filename, "w") as f:
        f.write(strictyaml.as_document(result).as_yaml())


def yaml_to_xlsform(filename: pathlib.Path):
    log.info("yaml_to_xlsform: %s", filename)


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
