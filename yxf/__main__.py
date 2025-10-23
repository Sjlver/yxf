"""
yxf: Convert from XLSForm to YAML and back.

To convert an XLSForm to a YAML file: `python -m yxf form.xlsx`.

By default, the result will be called `form.yaml`, in other words, the same name
as the input file with the extension changed to `.yaml`. You can specify a
different output file name using the `--output othername.yaml` option.

To convert a YAML file to an XLSForm: `python -m yxf form.yaml`.
"""

from .cli import main

if __name__ == "__main__":
    main()
