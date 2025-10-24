# yxf Tests

This directory contains tests for the yxf project, covering conversion workflows, edge cases, and CLI functionality.

## Test Data

The `testdata/` directory contains sample Excel files for testing:

1. **favorite-color.xlsx**: Simple form with translations
2. **simple-repeat.xlsx**: Tests repeat functionality
3. **cascading-select.xlsx**: Tests cascading selects with choices
4. **xlsform-dot-org-template.xlsx**: Comprehensive template

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
