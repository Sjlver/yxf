# Convert from XLSForm to YAML and back

yxf (short for **Y**AML **X**LS**F**orms) is a converter between XLSForms and
YAML files. With yxf, you can store forms as text-files. This brings a number of
advantages to managing many forms. For example, users can store forms in version
control or view differences between two forms.

## Usage

To convert an XLSForm to a YAML file: `python -m yxf form.xlsx`.

By default, the result will be called `form.yaml`, in other words, the same name
as the input file with the extension changed to `.yaml`. You can specify a
differnt output file name using the `--output othername.yaml` option.

To convert a YAML file to an XLSForm: `python -m yxf form.yaml`.

Here's a screenshot showing the YAML and XLSForm version of a form side-by-side:

![YAML and XLSForm version of a form](docs/yxf-yaml-and-xlsx-side-by-side.png)

## Comments in forms

yxf encourages adding comments to XLSForms. It uses a special column labeled `#`
to that end. Other tools ignore this column, so that it can be used for whatever
explanations that are useful to the readers of the `.xlsx` or `.yaml` files.

## Development

yxf is in an early stage of development. Feedback and contributions are welcome.

Please run all of the following before committing:

- Format the code: `black .`
- Run unit tests: `pytest`