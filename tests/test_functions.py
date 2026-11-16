"""
Tests for the `SQLOOKUP` custom function.

These call the function directly with the nested lists that Excel passes for each range, so they don't need Excel. The SQL engine itself (`xlwings.ext.sql._sql`) is tested in xlwings.
"""

import copy
import datetime as dt

import pytest
from xlwings.pro.udfs_officejs import custom_functions_meta

from xlwings_sqlookup import functions
from xlwings_sqlookup.functions import sqlookup

# Excel date serials
JAN_15 = 45306.0  # 2024-01-15
JAN_15_NOON = 45306.5  # 2024-01-15 12:00
FEB_01 = 45323.0  # 2024-02-01

ORDERS = [
    ["order_id", "order_date", "amount"],
    [1.0, JAN_15, 100.0],
    [2.0, FEB_01, 50.0],
]


# Registration


def test_registered_as_sqlookup_without_namespace():
    (meta,) = custom_functions_meta(functions)["functions"]
    assert meta["id"] == "SQLOOKUP"
    assert meta["name"] == "SQLOOKUP"


def test_registration_metadata():
    xlfunc = sqlookup.__xlfunc__
    assert xlfunc["help_url"] is None
    assert xlfunc["ret"]["doc"] == "Use SQL to query your data in Excel."
    query, varargs = xlfunc["args"]
    assert query["name"] == "query"
    assert query["doc"] == "The SQL query."
    assert varargs["name"] == "arg"
    assert varargs["vararg"]
    assert varargs["options"]["ndim"] == 2
    assert "expand" not in varargs["options"]
    assert not varargs["doc"].startswith("Positional argument")


# Without parse_dates


def test_query_without_parse_dates():
    assert sqlookup("SELECT order_date FROM a WHERE order_id = 1", ORDERS) == [
        ["order_date"],
        [JAN_15],
    ]


def test_aliases_are_passed_through():
    customers = [["id", "name"], [1.0, "Alice"]]
    assert sqlookup(
        "SELECT c.name, o.amount FROM c JOIN o ON c.id = o.order_id",
        [["c"]],
        customers,
        [["o"]],
        ORDERS,
    ) == [["name", "amount"], ["Alice", 100.0]]


# parse_dates


def test_parse_dates_converts_serials():
    assert sqlookup("SELECT order_date FROM a", ORDERS, [["order_date"]]) == [
        ["order_date"],
        [dt.datetime(2024, 1, 15)],
        [dt.datetime(2024, 2, 1)],
    ]


def test_parse_dates_keeps_time():
    table = [["ts"], [JAN_15_NOON]]
    assert sqlookup("SELECT ts FROM a", table, [["ts"]]) == [
        ["ts"],
        [dt.datetime(2024, 1, 15, 12)],
    ]


def test_parse_dates_allows_date_comparisons():
    assert sqlookup(
        "SELECT order_id FROM a WHERE order_date >= '2024-02-01'",
        ORDERS,
        [["order_date"]],
    ) == [["order_id"], [2.0]]


@pytest.mark.parametrize("parse_dates", ["ORDER_DATE", "Order_Date", "order_date"])
def test_parse_dates_ignores_case(parse_dates):
    table = [["Order_Date"], [JAN_15]]
    assert sqlookup("SELECT * FROM a", table, [[parse_dates]]) == [
        ["Order_Date"],
        [dt.datetime(2024, 1, 15)],
    ]


def test_parse_dates_matches_whole_names_only():
    table = [
        ["date", "update_count", "date_2"],
        [JAN_15, 3.0, FEB_01],
    ]
    assert sqlookup("SELECT * FROM a", table, [["date"]]) == [
        ["date", "update_count", "date_2"],
        [dt.datetime(2024, 1, 15), 3.0, FEB_01],
    ]


def test_parse_dates_does_not_match_part_of_a_name():
    assert sqlookup("SELECT order_date FROM a", ORDERS, [["date"]]) == [
        ["order_date"],
        [JAN_15],
        [FEB_01],
    ]


@pytest.mark.parametrize(
    "parse_dates",
    ["order_date,ship_date", "order_date, ship_date", " order_date ,ship_date ,"],
)
def test_parse_dates_with_several_columns(parse_dates):
    table = [["order_date", "ship_date", "amount"], [JAN_15, FEB_01, 100.0]]
    assert sqlookup("SELECT * FROM a", table, [[parse_dates]]) == [
        ["order_date", "ship_date", "amount"],
        [dt.datetime(2024, 1, 15), dt.datetime(2024, 2, 1), 100.0],
    ]


def test_parse_dates_applies_to_every_table():
    customers = [["id", "signup_date"], [1.0, FEB_01]]
    assert sqlookup(
        "SELECT a.order_date, b.signup_date FROM a JOIN b ON a.order_id = b.id",
        ORDERS,
        customers,
        [["order_date, signup_date"]],
    ) == [
        ["order_date", "signup_date"],
        [dt.datetime(2024, 1, 15), dt.datetime(2024, 2, 1)],
    ]


def test_parse_dates_leaves_empty_cells_empty():
    table = [["order_date", "n"], [JAN_15, 1.0], ["", 2.0], [None, 3.0]]
    assert sqlookup("SELECT order_date FROM a ORDER BY n", table, [["order_date"]]) == [
        ["order_date"],
        [dt.datetime(2024, 1, 15)],
        [None],
        [None],
    ]


@pytest.mark.parametrize("value", ["n/a", True, False])
def test_parse_dates_leaves_text_and_booleans(value):
    table = [["d"], [value]]
    assert sqlookup("SELECT d FROM a", table, [["d"]]) == [["d"], [value]]


def test_parse_dates_with_unknown_column_is_ignored():
    assert sqlookup("SELECT order_date FROM a", ORDERS, [["ship_date"]]) == [
        ["order_date"],
        [JAN_15],
        [FEB_01],
    ]


@pytest.mark.parametrize("parse_dates", ["", " ", ","])
def test_empty_parse_dates_converts_nothing(parse_dates):
    assert sqlookup("SELECT order_date FROM a", ORDERS, [[parse_dates]]) == [
        ["order_date"],
        [JAN_15],
        [FEB_01],
    ]


def test_parse_dates_with_header_only_table():
    table = [["order_date", "amount"]]
    assert sqlookup("SELECT * FROM a", table, [["order_date"]]) == [
        ["order_date", "amount"]
    ]


# Detecting the trailing parse_dates argument


def test_parse_dates_after_aliased_tables():
    customers = [["id", "signup_date"], [1.0, FEB_01]]
    assert sqlookup(
        "SELECT o.order_date, c.signup_date FROM orders o JOIN customers c ON o.order_id = c.id",
        [["orders"]],
        ORDERS,
        [["customers"]],
        customers,
        [["order_date,signup_date"]],
    ) == [
        ["order_date", "signup_date"],
        [dt.datetime(2024, 1, 15), dt.datetime(2024, 2, 1)],
    ]


def test_alias_followed_by_table_is_not_parse_dates():
    # The last argument is a table, so "order_date" is an alias, not parse_dates
    assert sqlookup(
        "SELECT amount FROM order_date WHERE order_id = 1", [["order_date"]], ORDERS
    ) == [["amount"], [100.0]]


def test_trailing_alias_without_table_is_parse_dates():
    # A trailing text cell is always parse_dates, even if it was meant as an alias
    assert sqlookup("SELECT order_date FROM a", ORDERS, [["orders"]]) == [
        ["order_date"],
        [JAN_15],
        [FEB_01],
    ]


def test_trailing_number_is_a_table():
    assert sqlookup("SELECT * FROM b", ORDERS, [[1.0]]) == [["1.0"]]


def test_only_parse_dates():
    assert sqlookup("SELECT 1 AS x", [["order_date"]]) == [["x"], [1]]


# The caller's data


def test_does_not_modify_arguments():
    orders = copy.deepcopy(ORDERS)
    customers = [["id", "signup_date"], [1.0, FEB_01], [2.0, ""]]
    expected = copy.deepcopy([orders, customers])
    sqlookup(
        "SELECT * FROM o JOIN c ON o.order_id = c.id",
        [["o"]],
        orders,
        [["c"]],
        customers,
        [["order_date, signup_date"]],
    )
    assert [orders, customers] == expected
