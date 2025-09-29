"""Functions to add XLSForm-specific logic to openpyxl Worksheets."""

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

# Colors for groups. These are essentially "oklch(0.8 - j*0.07, 0.25, 30*i)".
# Each row has a different hue, and values get darker with increasing column.
GROUP_COLORS = [
    ["#ff88d7", "#ff6ec1", "#ff54ab", "#ff3795"],
    ["#ff8e72", "#ff745b", "#ff5a43", "#ff3d29"],
    ["#ffab00", "#ff9300", "#ff7a00", "#ff6200"],
    ["#ffd100", "#ffb900", "#eea200", "#d78b00"],
    ["#cbf300", "#b5db00", "#9fc400", "#8bad00"],
    ["#00ff7c", "#00f165", "#00d94d", "#00c233"],
    ["#00ffe4", "#00f7ce", "#00dfb7", "#00c8a1"],
    ["#00ffff", "#00edff", "#00d5ff", "#00bdef"],
    ["#00edff", "#00d4ff", "#00bcff", "#00a5ff"],
    ["#a0cdff", "#8bb5ff", "#769dff", "#6386ff"],
    ["#faafff", "#e397ff", "#cc80ff", "#b668ff"],
    ["#ff96ff", "#ff7eff", "#ff66fa", "#eb4de3"],
    ["#ff88d7", "#ff6ec1", "#ff54ab", "#ff3795"],
]


def truncate_row(row):
    """Returns the row without any empty cells at the end.

    >>> truncate_row([1, 2, 3, None, None, None])
    [1, 2, 3]
    """
    row = list(row)
    while row and row[-1] is None:
        row.pop()
    return row


def stringify_value(v):
    """Converts a value to string in a way that's meaningful for values read from Excel."""
    return str(v) if v else ""


def headers(sheet):
    """Returns the values of the sheet's header row (i.e., the first row)."""

    for row in sheet.iter_rows(values_only=True):
        return [stringify_value(h) for h in truncate_row(row)]

    # If we get here, the sheet is empty.
    return []


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
        if sheet.max_row >= 1:
            for cell in sheet[1]:
                cell.style = HEADER_STYLE
        sheet.freeze_panes = sheet["A2"]

        sheet_headers = headers(sheet)
        comment_column = sheet_headers.index("#") if "#" in sheet_headers else -1
        type_column = sheet_headers.index("type") if "type" in sheet_headers else -1

        # Set column widths to reasonable values. First, get all widths.
        widths = [[] for _ in headers(sheet)]
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
            if i == comment_column:
                estimated_width = 2
            elif col_widths:
                estimated_width = col_widths[percentile_index] + 10
            else:
                estimated_width = 10

            if estimated_width <= 60:
                sheet.column_dimensions[
                    openpyxl.utils.get_column_letter(i + 1)
                ].width = estimated_width
            else:
                sheet.column_dimensions[
                    openpyxl.utils.get_column_letter(i + 1)
                ].width = 60
                for row_index, _ in enumerate(sheet):
                    sheet.cell(row=row_index + 1, column=i + 1).alignment = (
                        openpyxl.styles.Alignment(wrap_text=True)
                    )

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
        nesting_depth = 0
        if type_column >= 0:
            for row in content_rows(sheet):
                if str(row[type_column].value).startswith("begin_"):
                    if nesting_depth == 0:
                        group_number += 1
                    nesting_depth += 1

                if nesting_depth > 0:
                    group_colors = GROUP_COLORS[group_number % len(GROUP_COLORS)]
                    color_index = min(nesting_depth - 1, len(group_colors) - 1)
                    cell_color = group_colors[color_index]
                    row[comment_column].fill = openpyxl.styles.PatternFill(
                        fgColor="ff" + cell_color[1:], fill_type="solid"
                    )

                if str(row[type_column].value).startswith("end_"):
                    nesting_depth -= 1
