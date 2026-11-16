# Development

## Tests

The tests call `sqlookup()` directly with nested lists, the way Excel passes ranges, so they don't need Excel:

```text
uv run pytest
```

## Build

```text
uv build
```

The version comes from the git tag via `hatch-vcs`: a build of the commit tagged `0.1.0` is version `0.1.0`, other commits get a development version such as `0.1.1.dev3+g1a2b3c4`.

To try the wheel in xlwings Lite, upload it under **Files** > **Import/Export** and add the path to the wheel, e.g., `/data/xlwings_sqlookup-0.1.0-py3-none-any.whl`, to the workbook's `requirements.txt`.

## Release

Create a GitHub release with the version as the tag, without a `v` prefix, e.g., `gh release create 0.1.0 --generate-notes`. The release workflow then builds the package and publishes it to PyPI.
