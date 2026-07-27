# yxf Tests

This directory contains tests for the yxf project, covering conversion workflows, edge cases, and CLI functionality.

## Test Data

The `testdata/` directory contains sample Excel files for testing:

1. **favorite-color.xlsx**: Simple form with translations
2. **simple-repeat.xlsx**: Tests repeat functionality
3. **cascading-select.xlsx**: Tests cascading selects with choices
4. **xlsform-dot-org-template.xlsx**: Comprehensive template, including an
   `entities` sheet

It also contains two YAML forms. These are derived from real, moderately
complex forms: the structure is kept as it was, while all the names and all the
human-readable text have been replaced with innocuous questions about fruit.
They are stored as YAML rather than Excel so that they can be reviewed and
diffed; the tests convert them to Excel in memory.

5. **fruit-nested-groups.yaml**: Two languages, groups and repeats nested three
   levels deep, all four spellings of the group markers, `type: end` metadata at
   the top level, calculates, `choice_filter`, `select_one_from_file`, and a
   comment column.
6. **fruit-html-labels.yaml**: Four languages, 48 groups nested three levels
   deep, `type: end` metadata *inside* a group, and labels full of HTML
   (`<div style=...>`, `<b>`, `<br>`), including multi-line values.

`fruit-nested-groups.yaml` converts cleanly with `xls2xform`. `fruit-html-labels.yaml`
does not, because it keeps the SurveyCTO-only types (`text audit`, `audio audit`,
`calculate_here`) of the form it came from; pyxform only knows the ODK types.

## Running Tests

Run all tests (including doctests):

```bash
uv run pytest
```

## Snapshot Testing

This project uses [syrupy](https://github.com/tophat/syrupy) for snapshot testing. Snapshots are stored in `tests/__snapshots__/`.

To update snapshots after intentional changes:

```bash
uv run pytest tests/ --snapshot-update
```

Only YAML and Markdown outputs are snapshotted (not Excel files, as openpyxl output might not be byte-stable).

The formatting that yxf applies to Excel files is covered by `test_pretty.py`,
which writes a workbook in memory and inspects the cells directly rather than
snapshotting bytes.
