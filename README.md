# xlwings-sqlookup

An xlwings Lite plugin that adds the `SQLOOKUP` custom function: query your Excel ranges and tables with SQL, right in a cell.

```text
=SQLOOKUP("SELECT name, SUM(amount) AS total FROM a JOIN b ON a.id = b.customer_id GROUP BY name", Customers[#All], Orders[#All])
```

The SQL engine is [SQLite](https://www.sqlite.org/lang_select.html), running in memory inside xlwings Lite. Your data never leaves Excel.

## Installation

1. In xlwings Lite, open `requirements.txt` and add:

   ```text
   xlwings-sqlookup
   ```

2. Restart xlwings Lite.

The `SQLOOKUP` function is now available in every cell of this workbook. You don't need to write any Python code. Since plugins are activated per workbook, repeat these steps for each workbook that should use `SQLOOKUP`.

## Usage

```text
=SQLOOKUP(query, [alias1], range1, [alias2], range2, ..., [parse_dates])
```

- `query`: the SQL query as text, or a reference to a cell that contains it.
- `range1`, `range2`, ...: the ranges or Excel tables to query. The first row must contain the column names. With Excel tables, use `Table1[#All]` to include the header row.
- `alias1`, `alias2`, ...: optional table names. Without an alias, the ranges are available as `a`, `b`, `c`, ... in the order you pass them.
- `parse_dates`: optional, see [Dates](#dates).

The result spills into the neighboring cells, with the column names in the first row.

### Examples

Filter and sort a single range, which is available as `a`:

```text
=SQLOOKUP("SELECT * FROM a WHERE amount > 100 ORDER BY amount DESC", A1:C100)
```

Join two ranges, available as `a` and `b`:

```text
=SQLOOKUP("SELECT b.name, a.amount FROM a JOIN b ON a.customer_id = b.id", Orders[#All], Customers[#All])
```

Use aliases to give the ranges meaningful names:

```text
=SQLOOKUP("SELECT c.name, SUM(o.amount) AS total FROM orders o JOIN customers c ON o.customer_id = c.id GROUP BY c.name", "orders", Orders[#All], "customers", Customers[#All])
```

Write the query in a separate cell, e.g., `F1`, to keep the formula short:

```text
=SQLOOKUP(F1, "orders", Orders[#All], "customers", Customers[#All])
```

Put column names and aliases that contain spaces or other special characters in square brackets:

```text
=SQLOOKUP("SELECT [Order ID], Amount FROM a WHERE [Order ID] > 100", A1:B100)
```

Alternatively, use standard SQL double quotes, which you need to double inside an Excel formula, e.g., `"SELECT ""Order ID"" FROM a"`. Double quotes also work for names that contain a `]`.

## Dates

Excel passes dates to custom functions as numbers (date serials, e.g., `45306` for 2024-01-15). To work with them as dates in SQL, list the date columns as the last argument, separated by commas:

```text
=SQLOOKUP("SELECT * FROM a WHERE order_date >= '2024-02-01'", Orders[#All], "order_date")
=SQLOOKUP("SELECT * FROM orders o JOIN customers c ON o.customer_id = c.id", "orders", Orders[#All], "customers", Customers[#All], "order_date, signup_date")
```

- The names must match the column names exactly, but case doesn't matter: `"order_date"` matches `Order_Date`, but `"date"` doesn't match `order_date`.
- The conversion applies to the matching columns of all ranges.
- Empty cells stay empty. Cells with text or `TRUE`/`FALSE` are left as they are.
- Within the query, dates are text in ISO format, e.g., `2024-01-15 00:00:00`, the same format as SQLite's `datetime()`. Comparisons such as `order_date >= '2024-02-01'` work, but since `2024-01-31 00:00:00` sorts after `2024-01-31`, `order_date <= '2024-01-31'` and `BETWEEN '2024-01-01' AND '2024-01-31'` leave out the last day. For an end date, use `order_date < '2024-02-01'` or `date(order_date) <= '2024-01-31'`, and to match a single day, use `date(order_date) = '2024-01-15'`. See also SQLite's [date and time functions](https://www.sqlite.org/lang_datefunc.html).
- In the result, the date columns are written back to Excel as dates, unless they also contain text, in which case their dates are returned as text. Values that SQL functions compute from dates, such as `date(order_date)`, are text.

Since the last argument is a list of date columns whenever it's a single cell with text, an alias always needs a range after it.

## Moving from the SQLookup add-in

xlwings-sqlookup is an alternative to the [SQLookup add-in](https://sqlookup.com): if you already use xlwings Lite, it gives you the same `SQLOOKUP` function without a separate add-in. Both define a function called `SQLOOKUP`, so you can't have the SQLookup add-in and the xlwings-sqlookup plugin installed at the same time.

The function works the same way as in the SQLookup add-in, with one difference: the date columns in the last argument must match the column names exactly (ignoring case), see [Dates](#dates). The SQLookup add-in also converts columns that merely contain the given text, e.g., `"date"` converts `order_date` and `update_count`. If your formulas rely on this, list the full column names instead, e.g., `"order_date, ship_date"`.

To switch from the SQLookup add-in to the plugin:

1. Remove the SQLookup add-in from Excel.
2. Clear the Office cache, see [Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache). This also deletes data that xlwings Lite keeps in the cache, such as imported files and local environment variables, so make a backup first.
3. Install the plugin as described under [Installation](#installation).
