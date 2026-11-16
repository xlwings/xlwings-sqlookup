import numbers

import xlwings as xw
from xlwings import arg, func
from xlwings.ext.sql import _sql

# No __xlwings_func_namespace__: the function must be called exactly =SQLOOKUP(...)


@func
@arg("query", doc="The SQL query.")
@arg(
    "*arg",
    ndim=2,
    doc="A table alias, an Excel range/table, or a comma-separated list of date columns as the last argument.",
)
def sqlookup(query, *arg):
    # The docstring is shown verbatim in Excel's formula help, so keep it plain text
    """Use SQL to query your data in Excel."""
    # arg is singular so it appears as "arg" in Excel
    tables = list(arg)
    date_columns = set()
    if tables and _is_single_string(tables[-1]):
        date_columns = {name.strip().lower() for name in tables.pop()[0][0].split(",")}
        date_columns.discard("")
    return _sql(query, *(_parse_dates(table, date_columns) for table in tables))


def _is_single_string(table):
    # Same rule as _sql uses to detect an alias
    return len(table) == 1 and len(table[0]) == 1 and isinstance(table[0][0], str)


def _parse_dates(table, date_columns):
    """Returns a new table with Excel date serials in `date_columns` converted to datetimes.

    Column names match exactly, ignoring case. Aliases, header-only tables, and tables without a matching column are returned unchanged.
    """
    if not date_columns or _is_single_string(table) or len(table) < 2:
        return table
    header = table[0]
    indices = {
        i
        for i, name in enumerate(header)
        if isinstance(name, str) and name.lower() in date_columns
    }
    if not indices:
        return table
    return [header] + [
        [_to_datetime(value) if i in indices else value for i, value in enumerate(row)]
        for row in table[1:]
    ]


def _to_datetime(value):
    # Empty cells, text, and booleans stay as they are
    if isinstance(value, numbers.Number) and not isinstance(value, bool):
        return xw.to_datetime(value)
    return value
