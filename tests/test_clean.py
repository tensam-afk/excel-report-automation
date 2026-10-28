from datetime import datetime

import pandas as pd
import pytest

from clean import clean, load, map_columns, normalize_header, parse_date, parse_number, read_file

HEADER = "Bill No.,Bill Date,Outlet,Item Name,Qty,Rate,Amount (Rs)"


def write_csv(path, lines):
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def issue_counts(issues):
    return issues.groupby(["issue", "action"]).size().to_dict()


@pytest.mark.parametrize("name,expected", [
    ("Bill No.", "bill_no"),
    ("  Invoice Date ", "invoice_date"),
    ("Amount (Rs)", "amount"),
    ("Unit Price (INR)", "unit_price"),
    ("Price (₹)", "price"),
])
def test_normalize_header(name, expected):
    assert normalize_header(name) == expected


def test_map_columns():
    headers = ["Bill No.", " QTY ", "Customer Phone", "Amount (Rs)", "Store", "Net Amount", None]
    assert map_columns(headers) == {"order_id": 0, "quantity": 1, "total": 3, "branch": 4}


@pytest.mark.parametrize("value,expected", [
    ("₹1,234.50", 1234.5),
    ("₹ 899", 899.0),
    ("Rs 1234", 1234.0),
    ("Rs. 1,234.00", 1234.0),
    ("INR 1,23,456.78", 123456.78),
    ("1234/-", 1234.0),
    ("2 pcs", 2.0),
    ("(150.00)", -150.0),
    ("-75", -75.0),
    (799, 799.0),
    (12.5, 12.5),
    ("", None),
    ("-", None),
    ("N/A", None),
    (None, None),
    (float("nan"), None),
])
def test_parse_number(value, expected):
    assert parse_number(value) == expected


SEP_3 = pd.Timestamp("2026-09-03")


@pytest.mark.parametrize("value", [
    "03/09/2026",
    "3/9/2026",
    "03-09-2026",
    "03/09/26",
    "2026-09-03",
    "2026/09/03",
    "3 Sep 2026",
    "3 sept 2026",
    "03-Sep-26",
    "Sep 3, 2026",
    "3 September 2026",
    " 2026-09-03 14:22:05 ",
    "03/09/2026 6:15 PM",
    datetime(2026, 9, 3, 18, 45),
    pd.Timestamp("2026-09-03 10:00"),
    46268,
])
def test_parse_date_day_first(value):
    assert parse_date(value) == SEP_3


@pytest.mark.parametrize("value", ["31/09/2026", "2026-09-31", "13/13/2026", "soon", "", None, 12])
def test_parse_date_rejects_bad_values(value):
    assert parse_date(value) is None


def test_parse_date_month_first():
    assert parse_date("09/03/2026", dayfirst=False) == SEP_3
    assert parse_date("2026-09-03", dayfirst=False) == SEP_3
    assert parse_date("03/09/2026", dayfirst=False) == pd.Timestamp("2026-03-09")


def test_header_below_title_rows(tmp_path):
    path = write_csv(tmp_path / "export.csv", [
        "Sales register - Pune",
        "",
        HEADER,
        "PUN-1,01/09/2026,Pune,Phone Case,2,349,698",
    ])
    df = read_file(path)
    assert len(df) == 1
    assert df.iloc[0]["order_id"] == "PUN-1"
    assert df.iloc[0]["source_row"] == 4


def test_missing_columns_are_reported(tmp_path):
    path = write_csv(tmp_path / "export.csv", ["Bill No,Bill Date,Outlet,Customer,Amount", "1,01/09/2026,Pune,A,349"])
    with pytest.raises(ValueError, match="no column for product, two of quantity/unit_price/total"):
        read_file(path)


def test_no_header_row(tmp_path):
    path = write_csv(tmp_path / "export.csv", ["Customer,Phone", "A,98765"])
    with pytest.raises(ValueError, match="header row"):
        read_file(path)


def test_cleaning_drops_and_fixes(tmp_path):
    write_csv(tmp_path / "pune.csv", [
        HEADER,
        "PUN-1,01/09/2026,Pune,Phone Case,2,₹349,₹698",
        "PUN-1,01/09/2026,Pune,Phone Case,2,₹349,₹698",           # exact duplicate
        ",,,,,,",                                                  # blank
        "pun-1 ,2026-09-01,PUNE,phone case,2,Rs 349,698/-",       # same line exported again
        "PUN-2,02/09/2026, pune,Wireless Mouse,,599,1198",        # quantity from amount
        "PUN-3,02/09/2026,Pune,Wireless Mouse,,599,",             # nothing to work quantity out from
        "PUN-4,31/09/2026,Pune,Phone Case,1,349,349",             # no such date
        "PUN-5,03/09/2026,Pune,Phone Case,0,349,0",               # zero quantity
        "PUN-6,03/09/2026,Pune,Wireless Mouse,1,599,",            # total missing
        ",03/09/2026,Pune,Phone Case,1,349,349",                  # no order id
    ])
    data, issues = clean(load(tmp_path))

    assert issue_counts(issues) == {
        ("blank row", "dropped"): 1,
        ("exact duplicate", "dropped"): 1,
        ("repeated order line", "dropped"): 1,
        ("inconsistent branch name", "fixed"): 2,
        ("inconsistent product name", "fixed"): 1,
        ("missing quantity", "fixed"): 1,
        ("missing quantity", "dropped"): 1,
        ("unreadable date", "dropped"): 1,
        ("zero or negative quantity", "dropped"): 1,
        ("missing total", "fixed"): 1,
        ("missing order id", "fixed"): 1,
        ("no category column", "fixed"): 1,
    }
    assert list(data["order_id"]) == ["PUN-1", "PUN-2", "NOID-pune-11", "PUN-6"]
    assert set(data["branch"]) == {"Pune"}
    assert set(data["product"]) == {"Phone Case", "Wireless Mouse"}
    assert list(data["quantity"]) == [2, 2, 1, 1]
    assert data["total"].sum() == 698 + 1198 + 349 + 599

    dup = issues[issues["issue"] == "repeated order line"].iloc[0]
    assert dup["source_row"] == 5
    assert dup["detail"] == "same order and product as pune.csv row 2"


def test_category_filled_from_other_files(tmp_path):
    write_csv(tmp_path / "a.csv", [
        "Order ID,Date,Branch,Product,Category,Qty,Price",
        "A1,01/09/2026,Pune,Phone Case,Mobile Accessories,1,349",
    ])
    write_csv(tmp_path / "b.csv", [
        "Order ID,Date,Branch,Product,Qty,Price",
        "B1,01/09/2026,Nashik,Phone Case,1,349",
        "B2,01/09/2026,Nashik,Smart Watch,1,3499",
    ])
    data, issues = clean(load(tmp_path))

    categories = dict(zip(data["order_id"], data["category"]))
    assert categories == {"A1": "Mobile Accessories", "B1": "Mobile Accessories", "B2": "Uncategorised"}
    # a missing column is one note per file, not one per row
    assert issue_counts(issues) == {("no category column", "fixed"): 1, ("no total column", "fixed"): 2}


def test_same_order_number_in_two_branches_is_kept(tmp_path):
    write_csv(tmp_path / "sales.csv", [
        "Order ID,Date,Branch,Product,Qty,Price",
        "1001,01/09/2026,Pune,Phone Case,1,349",
        "1001,01/09/2026,Nashik,Phone Case,1,349",
    ])
    data, issues = clean(load(tmp_path))
    assert len(data) == 2
    assert "repeated order line" not in set(issues["issue"])


def test_month_first_dates(tmp_path):
    write_csv(tmp_path / "us.csv", [
        "Order ID,Date,Branch,Product,Qty,Price",
        "1,09/13/2026,Pune,Phone Case,1,349",
    ])
    data, issues = clean(load(tmp_path))
    assert data.empty
    bad = issues[issues["issue"] == "unreadable date"]
    assert "--month-first" in bad.iloc[0]["detail"]

    data, issues = clean(load(tmp_path), dayfirst=False)
    assert data.iloc[0]["date"] == pd.Timestamp("2026-09-13")
