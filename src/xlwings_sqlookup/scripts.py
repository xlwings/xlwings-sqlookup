import datetime as dt
from textwrap import dedent

import xlwings as xw
from xlwings import script

CUSTOMERS = [
    ["customer_id", "name", "email", "city"],
    [1, "Alice Brown", "alice.brown@example.com", "New York"],
    [2, "Bob Smith", "bob.smith@example.com", None],
    [3, "Charlie Davis", "charlie.davis@example.com", "Los Angeles"],
]

ORDERS = [
    ["order id", "customer_id", "product", "amount", "is_shipped", "order_date"],
    [101, 1, "Laptop", 1200.50, True, dt.datetime(2024, 1, 10)],
    [102, 1, "Mouse", 25.95, False, dt.datetime(2024, 1, 12)],
    [103, 2, "Keyboard", 75.20, True, dt.datetime(2024, 1, 15)],
    [104, 3, "Monitor", 300.15, False, dt.datetime(2024, 1, 18)],
]

# (row, title, formula), with the formula written into the cell below the title
EXAMPLES = [
    (
        12,
        "Example 1: Orders after 2024-01-12",
        """\
        =SQLOOKUP(
            "SELECT *
             FROM orders
             WHERE order_date > datetime('2024-01-12')",
            "orders", Orders[#All],
            "order_date"
        )""",
    ),
    (
        17,
        "Example 2: Inner join",
        """\
        =SQLOOKUP(
            "SELECT o.[order id], c.name, o.product, o.amount
             FROM c
             INNER JOIN o ON c.customer_id = o.customer_id
             WHERE c.city IS NOT NULL",
            "c", Customers[#All],
            "o", Orders[#All]
        )""",
    ),
    (
        23,
        "Example 3: Total amount spent per customer",
        """\
        =SQLOOKUP(
            "SELECT customer_id, SUM(amount) AS total_spent
             FROM orders
             GROUP BY customer_id",
            "orders", Orders[#All]
        )""",
    ),
]


@script
async def insert_sample_sheet(book: xw.BookAsync):
    """Inserts a new sheet with sample tables and SQLOOKUP formulas."""
    for sheet in book.sheets:
        for table in sheet.tables:
            # Excel table names are case-insensitive
            if table.name.lower() in ["orders", "customers"]:
                raise xw.XlwingsError(
                    "Excel table names must be unique: delete the sheet with the "
                    "Customers and Orders tables before inserting the examples again."
                )
    sheet = book.sheets.add()
    sheet["A1"].value = CUSTOMERS
    sheet["A6"].value = ORDERS

    customers_range = sheet[0 : len(CUSTOMERS), 0 : len(CUSTOMERS[0])]
    orders_range = sheet[5 : 5 + len(ORDERS), 0 : len(ORDERS[0])]
    sheet.tables.add(customers_range, name="Customers")
    sheet.tables.add(orders_range, name="Orders")
    customers_range.columns.autofit()
    orders_range.columns.autofit()

    for row, title, formula in EXAMPLES:
        sheet[f"A{row}"].value = title
        sheet[f"A{row + 1}"].value = dedent(formula)

    sheet.activate()
