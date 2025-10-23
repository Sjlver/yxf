"""yxf: Convert from XLSForm to YAML and back.

This module provides library functions for converting between XLSForm Excel
files, YAML, and Markdown formats. Functions work with file-like objects (Excel)
or strings (Markdown, YAML) for maximum flexibility.

Example usage:
    import yxf

    # Read an XLSForm from a file object
    with open("form.xlsx", "rb") as f:
        form = yxf.read_xlsform(f)

    # Convert to YAML
    yaml_str = yxf.write_yaml(form)
"""

from .excel import read_xlsform, write_xlsform
from .markdown import read_markdown, write_markdown
from .yaml import read_yaml, write_yaml

__all__ = [
    "read_xlsform",
    "write_xlsform",
    "read_yaml",
    "write_yaml",
    "read_markdown",
    "write_markdown",
]
