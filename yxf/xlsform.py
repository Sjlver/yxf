"""Functions to add XLSForm-specific logic to openpyxl Worksheets."""

import grapefruit
import openpyxl.styles
import openpyxl.utils

# Openpyxl styles for various parts of an XLSForm.
HEADER_STYLE = openpyxl.styles.NamedStyle(name="header")
HEADER_STYLE.font = openpyxl.styles.Font(bold=True)

CODE_STYLE = openpyxl.styles.NamedStyle(name="code")
CODE_STYLE.font = openpyxl.styles.Font(name="Courier New", color="ff19007d")

NAME_STYLE = openpyxl.styles.NamedStyle(name="name")
NAME_STYLE.font = openpyxl.styles.Font(name="Courier New", color="ffa13b16")

COMMENT_STYLE = openpyxl.styles.NamedStyle(name="comment")
COMMENT_STYLE.font = openpyxl.styles.Font(name="Courier New", color="ff009c5d")

NOTE_STYLE = openpyxl.styles.NamedStyle(name="note")
NOTE_STYLE.font = openpyxl.styles.Font(color="ff555555")

GROUP_COLOR = grapefruit.Color.NewFromHtml("#c7ffdb")


def truncate_row(row):
    """Returns the row without any empty cells at the end."""
    row = list(row)
    while row and row[-1] is None:
        row.pop()
    return row


def headers(sheet):
    """Returns the values of the sheet's header row (i.e., the first row)."""

    return truncate_row(next(sheet.iter_rows(values_only=True)))


def content_rows(sheet, **kwargs):
    """Returns an iterator over the sheet's content rows.

    These are the rows below the header row)."""

    rows_iter = sheet.iter_rows(**kwargs)
    next(rows_iter)
    return rows_iter


def make_pretty(wb: openpyxl.Workbook):
    """Applies styles to the given workbook to make it prettier.

    This function knows about some XLSForm column names and row types, and
    formats them appropriately. It also adds color to highlight the group
    structure of the file.
    """
    for sheet in wb:
        for cell in sheet[1]:
            cell.style = HEADER_STYLE
        sheet.freeze_panes = sheet["A2"]

        sheet_headers = headers(sheet)
        comment_column = sheet_headers.index("#") if "#" in sheet_headers else -1
        type_column = sheet_headers.index("type") if "type" in sheet_headers else -1

        # Set column widths to reasonable values. First, get all widths.
        widths = [[] for _ in sheet[1]]
        for row in content_rows(sheet):
            for i, cell in enumerate(row):
                if cell.value:
                    width = max(len(w) for w in cell.value.splitlines())
                    widths[i].append(width)

        # We take the 75th percentile width plus 10.
        for i, ws in enumerate(widths):
            col_widths = sorted(ws)
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
        code_columns = set(
            ["calculation", "relevant", "constraint", "repeat_count", "instance_name"]
        )
        for row in content_rows(sheet):
            for i, cell in enumerate(row):
                if sheet_headers[i] in code_columns:
                    cell.style = CODE_STYLE
                elif sheet_headers[i] == "name":
                    cell.style = NAME_STYLE
                elif sheet_headers[i] == "#":
                    cell.style = COMMENT_STYLE
                elif type_column >= 0 and row[type_column].value == "note":
                    cell.style = NOTE_STYLE

        # Highlight groups and nesting
        group_number = 0
        color_stack = []
        if type_column >= 0:
            for row in content_rows(sheet):
                if str(row[type_column].value).startswith("begin_"):
                    if not color_stack:
                        h, s, v = GROUP_COLOR.hsv
                        h += 20 * group_number
                        group_number += 1
                        color_stack.append(grapefruit.Color.NewFromHsv(h, s, v))
                    else:
                        color_stack.append(color_stack[-1].DarkerColor(0.05))

                if color_stack:
                    row[comment_column].fill = openpyxl.styles.PatternFill(
                        fgColor="ff" + color_stack[-1].html[1:], fill_type="solid"
                    )

                if str(row[type_column].value).startswith("end_"):
                    color_stack.pop()
