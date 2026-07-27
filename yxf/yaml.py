"""Functions for reading and writing YAML format XLSForms.

This module provides functions for converting between YAML strings and
XLSForm dictionaries. Functions work with strings rather than file objects
for maximum flexibility.
"""

import strictyaml

from .xlsform import KNOWN_SHEETS


def read_yaml(content: str) -> dict:
    """Parse a YAML string into a form dictionary.

    Args:
        content: YAML string content

    Returns:
        Dictionary with sheet names as keys and lists of row dicts as values.
        Includes a "yxf" key with metadata.

    Raises:
        ValueError: If YAML is missing required keys
        strictyaml.YAMLError: If YAML is malformed
    """
    form = strictyaml.load(content).data

    if "yxf" not in form or "headers" not in form["yxf"]:
        raise ValueError('YAML file must have a "yxf" entry.')
    if "survey" not in form["yxf"]["headers"]:
        raise ValueError('YAML file must have a "survey" entry.')

    return form


def write_yaml(form: dict) -> str:
    """Convert a form dictionary to a YAML string.

    Args:
        form: Dictionary with sheet names as keys and lists of row dicts as values.
              Must include a "yxf" key with headers metadata.

    Returns:
        YAML string representation of the form
    """

    if "yxf" not in form or not form["yxf"]:
        raise ValueError('YAML file must have a "yxf" entry.')

    # Since StrictYAML can't store empty lists, we remove empty sheets from the
    # form. Their headers are still stored in the yxf entry. We work on a copy,
    # so that callers keep the form they passed in.
    form = {
        key: value for (key, value) in form.items() if value or key not in KNOWN_SHEETS
    }

    return strictyaml.as_document(form).as_yaml()
