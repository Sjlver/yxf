"""Functions for reading and writing Markdown format XLSForms.

This module provides functions for converting between Markdown strings and
XLSForm dictionaries. Functions work with strings rather than file objects
for maximum flexibility.
"""

import logging
import re

import markdown_it
import markdown_it.tree

from .xlsform import row_to_dict, validate_sheet_name

log = logging.getLogger(__name__)


def read_markdown(content: str, source_name: str = "<input>") -> dict:
    """Parse a Markdown string into a form dictionary.

    Args:
        content: Markdown string content
        source_name: Name of source (for error messages), defaults to "<input>"

    Returns:
        Dictionary with sheet names as keys and lists of row dicts as values.
        Includes a "yxf" key with headers metadata.

    Raises:
        ValueError: If Markdown structure is invalid
    """
    parser = markdown_it.MarkdownIt("js-default")
    ast = markdown_it.tree.SyntaxTreeNode(parser.parse(content))
    form = {}
    form_headers = {}
    sheet_name = None
    result = []

    for node in ast:
        if node.tag == "h2":
            sheet_name = node.children[0].content
            validate_sheet_name(sheet_name, source_name, node.map[0])
            result = []
        elif node.tag == "p":
            content_text = node.children[0].content
            match = re.match(r"%%\s*(.*)", content_text)
            if match:
                sheet_name = match.group(1)
                validate_sheet_name(sheet_name, source_name, node.map[0])
            else:
                # Other paragraphs are treated as comments and added to the
                # beginning of the current sheet.
                result.append({"#": content_text})
        elif node.tag == "table":
            if not sheet_name:
                raise ValueError(
                    f"{source_name}:{node.map[0]}: No sheet name specified for table."
                )
            thead = node.children[0]
            headers = [c.children[0].content for c in thead.children[0].children]
            add_comment_column = (
                headers[0] != "#" and len(result) > 0 and "#" in result[0]
            )
            if add_comment_column:
                headers.insert(0, "#")
            if len(node.children) > 1:
                tbody = node.children[1]
                rows = tbody.children
                rows = [[c.children[0].content for c in row.children] for row in rows]
                for values in rows:
                    if add_comment_column:
                        values.insert(0, "")
                    row_dict = row_to_dict(headers, values)
                    if row_dict:
                        result.append(row_dict)
            form[sheet_name] = result
            form_headers[sheet_name] = headers

    form["yxf"] = {"headers": form_headers}
    return form


def write_markdown(form: dict, source_name: str = "<source>") -> str:
    """Convert a form dictionary to a Markdown string.

    Args:
        form: Dictionary with sheet names as keys and lists of row dicts as values.
              Must include a "yxf" key with headers metadata.
        source_name: Name of source file (for warning messages), defaults to "<source>"

    Returns:
        Markdown string representation of the form

    The form is not modified; comment removal and escaping happen on copies.
    """
    md = []
    for sheet_name in form["yxf"]["headers"]:
        md.append(f"## {sheet_name}")
        md.append("")

        # Copy the sheet and its headers, so that rendering Markdown does not
        # change the caller's form.
        sheet = [dict(row) for row in form.get(sheet_name, [])]
        headers = list(form["yxf"]["headers"][sheet_name])

        # Before we render the table, look for comments and render those.
        # We simply put them as paragraphs in the Markdown file.
        for row in sheet:
            if "#" in row:
                if row["#"]:
                    md.append(row["#"])
                    md.append("")
                del row["#"]

        if headers and headers[0] == "#":
            headers.pop(0)
        header_indices = dict(zip(headers, range(len(headers))))

        for i, row in enumerate(sheet):
            for k, v in row.items():
                # Markdown does not support multi-line entries in cells. Check
                # and complain if needed.
                if "\n" in v:
                    log.warning(
                        "%s:%d Multi-line value for column %s.\n"
                        "Markdown does not support multi-line values. Use YAML instead.",
                        source_name,
                        i + 2,
                        k,
                    )
                    v = v.replace("\n", " ")
                # Markdown uses "|" as a table cell separator. Escape it if it
                # occurs in one of the values. And duplicate each escape character.
                row[k] = v.replace("\\", "\\\\").replace("|", "\\|")

        # Find column widths
        widths = [len(h) for h in headers]
        for row in sheet:
            for k, v in row.items():
                i = header_indices[k]
                widths[i] = max(widths[i], len(v))

        # Render the table
        header_row = [h.ljust(w) for (h, w) in zip(headers, widths)]
        md.append(f"| {' | '.join(header_row)} |")
        separator_row = ["-" * w for w in widths]
        md.append(f"| {' | '.join(separator_row)} |")
        for row in sheet:
            if not row:
                continue
            formatted_row = [row.get(h, "").ljust(w) for (h, w) in zip(headers, widths)]
            md.append(f"| {' | '.join(formatted_row)} |")
        md.append("")

    return "\n".join(md)
