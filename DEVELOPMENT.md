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

To try the wheel in xlwings Lite, upload it under **Files** > **Import/Export** and add `/data/xlwings_sqlookup-0.1.0-py3-none-any.whl` to the workbook's `requirements.txt`.
