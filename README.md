# Convert from XLSForm to YAML and back

yxf (short for **Y**AML **X**LS**F**orms) is a converter between XLSForms and
YAML files. With yxf, you can store forms as text-files. This brings a number of
advantages to managing many forms. For example, you can store forms in version
control, view differences between two forms, and easily share updates between
multiple forms.

## Usage

To convert an XLSForm to a YAML file: `python -m yxf form.xlsx`.

By default, the result will be called `form.yaml`, in other words, the same name
as the input file with the extension changed to `.yaml`. You can specify a
different output file name using the `--output othername.yaml` option.

To convert a YAML file to an XLSForm: `python -m yxf form.yaml`.

Here's a screenshot showing the YAML and XLSForm version of a form side-by-side:

![YAML and XLSForm version of a form](https://github.com/Sjlver/yxf/blob/main/docs/yxf-yaml-and-xlsx-side-by-side.png)

### Using Markdown

yxf can generate and read Markdown instead of YAML. The format is taken from
[md2xlsform](https://github.com/joshuaberetta/md2xlsform). This can be useful if
you would like something more compact than YAML, e.g., to paste into a community
forum.

To use Markdown, add a `--markdown` argument to the yxf invocation:

```shell
# Will generate form.md
python -m yxf --markdown form.xlsx
```

## Installation

Get the latest version from the GitHub repo:

```
python -m pip install yxf
```

## Features

### Comments in forms

yxf encourages adding comments to XLSForms. It uses a special column labeled `#`
to that end. Other tools ignore this column, so that it can be used for
explanations that are useful to the readers of the `.xlsx` or `.yaml` files.

### Pretty spreadsheets

yxf tries hard to make the XLSForm files look pretty. It is possible to use yxf
just for that purpose, without intending to store forms as YAML files. To do so,
simply convert a form to YAML and back:

```
python -m yxf form.xlsx
python -m yxf -o form-pretty.xlsx form.yaml
```

## Development

yxf is still being developed. Feedback and contributions are welcome.
Please open issues in the [issue tracker](https://github.com/Sjlver/yxf/issues).

Please run all of the following before committing:

- Format the code: `black .`
- Run unit tests: `pytest`
- See lint warnings (and please fix them :)): `pylint yxf`

To publish on PyPI:

- Increment the version number in `setup.cfg`.
- Run `python -m build`.
- Upload using `python -m twine upload dist/*-version-*`.