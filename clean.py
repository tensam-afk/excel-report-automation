import csv
import io
import logging
import numbers
import re
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

log = logging.getLogger("report")

STANDARD_COLUMNS = ["date", "order_id", "branch", "product", "category", "quantity", "unit_price", "total"]
ISSUE_COLUMNS = ["source_file", "source_row", "order_id", "issue", "action", "detail"]

# every billing system names these differently. Headers are lowercased and
# punctuation becomes "_" before lookup, so "Bill No." matches "bill_no".
COLUMN_ALIASES = {
    "date": ["date", "order_date", "bill_date", "invoice_date", "sale_date", "txn_date", "transaction_date"],
    "order_id": ["order_id", "order_no", "order_number", "bill_no", "bill_number", "invoice_no",
                 "invoice_number", "receipt_no", "txn_id"],
    "branch": ["branch", "branch_name", "store", "store_name", "outlet", "outlet_name", "location", "shop"],
    "product": ["product", "product_name", "product_description", "item", "item_name", "item_description",
                "description"],
    "category": ["category", "product_category", "item_category", "department", "dept"],
    "quantity": ["quantity", "qty", "units", "pcs", "no_of_units"],
    "unit_price": ["unit_price", "price", "rate", "selling_price", "mrp", "price_per_unit"],
    "total": ["total", "total_amount", "amount", "net_amount", "line_total", "bill_amount", "sales", "value"],
}
ALIAS_LOOKUP = {alias: col for col, aliases in COLUMN_ALIASES.items() for alias in aliases}

# Branch is part of the key because each branch's billing software numbers
# its orders on its own, so the same order number can show up in two branches.
ORDER_KEY = ["branch", "order_id"]
LINE_KEY = ORDER_KEY + ["product"]

ISO_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%b-%y", "%b %d %Y", "%B %d %Y"]
DAY_FIRST = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y"]
MONTH_FIRST = ["%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m/%d/%y", "%m-%d-%y"]
EXCEL_EPOCH = datetime(1899, 12, 30)


def is_missing(value):
    return value is None or (not isinstance(value, str) and pd.isna(value))


def normalize_header(name):
    key = re.sub(r"[^a-z0-9]+", "_", str(name).lower()).strip("_")
    return re.sub(r"_(rs|inr)$", "", key)


def map_columns(headers):
    found = {}
    for pos, name in enumerate(headers):
        col = None if is_missing(name) else ALIAS_LOOKUP.get(normalize_header(name))
        if col and col not in found:
            found[col] = pos
    return found


def parse_number(value):
    if is_missing(value) or isinstance(value, bool):
        return None
    if isinstance(value, numbers.Number):
        return float(value)
    text = str(value).strip()
    m = re.search(r"\d[\d,]*(?:\.\d+)?|\.\d+", text)
    if not m:
        return None
    number = float(m.group().replace(",", ""))
    if "-" in text[:m.start()] or (text.startswith("(") and text.endswith(")")):
        number = -number
    return number


def parse_date(value, dayfirst=True):
    if is_missing(value):
        return None
    if isinstance(value, datetime):
        return pd.Timestamp(value).normalize()
    if isinstance(value, date):
        return pd.Timestamp(value)
    if isinstance(value, numbers.Number):
        # a date column that lost its formatting shows up as Excel's day count
        if 20000 < value < 80000:
            return pd.Timestamp(EXCEL_EPOCH + timedelta(days=int(value)))
        return None

    text = re.sub(r"[,\s]+", " ", str(value)).strip()
    text = re.sub(r"[ T]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?( ?[AaPp][Mm])?$", "", text)
    text = re.sub(r"\bsept\b", "Sep", text, flags=re.IGNORECASE)
    for fmt in ISO_FORMATS + (DAY_FIRST if dayfirst else MONTH_FIRST):
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return None


def clean_text(value):
    if is_missing(value):
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def clean_id(value):
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = clean_text(value)
    return text.upper() if text else None


def canonical_names(values):
    # use whichever spelling turns up most often, so "PUNE" and "pune " become
    # "Pune" without breaking names like "USB-C Cable" the way .title() would
    best = {}
    for name, _ in Counter(values.dropna()).most_common():
        best.setdefault(name.casefold(), name)
    return best


def read_table(path):
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        return pd.read_excel(path, header=None, dtype=object)
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        # older Windows billing software still exports in the local code page
        text = raw.decode("cp1252")
    rows = list(csv.reader(io.StringIO(text)))
    # exports often have short rows (blank lines, a title line), pad them so every row is the same width
    width = max((len(r) for r in rows), default=0)
    return pd.DataFrame([r + [None] * (width - len(r)) for r in rows])


def find_header_row(table, max_rows=10):
    # exports often start with a title or a blank line before the real header
    for i in range(min(max_rows, len(table))):
        if len(map_columns(table.iloc[i].tolist())) >= 3:
            return i
    return None


def read_file(path):
    path = Path(path)
    table = read_table(path)
    header = find_header_row(table)
    if header is None:
        raise ValueError(f"{path.name}: no header row in the first 10 rows matches the names in COLUMN_ALIASES")

    headers = table.iloc[header].tolist()
    found = map_columns(headers)
    missing = [c for c in ("date", "branch", "product") if c not in found]
    if len({"quantity", "unit_price", "total"} & found.keys()) < 2:
        missing.append("two of quantity/unit_price/total")
    if missing:
        raise ValueError(
            f"{path.name}: no column for {', '.join(missing)}. Headers are {headers}. "
            "Add the missing names to COLUMN_ALIASES in clean.py."
        )

    ignored = [str(h) for pos, h in enumerate(headers) if clean_text(h) and pos not in found.values()]
    if ignored:
        log.info("%s: ignoring columns %s", path.name, ", ".join(ignored))

    body = table.iloc[header + 1:]
    df = pd.DataFrame(index=body.index)
    for col in STANDARD_COLUMNS:
        df[col] = body[found[col]].astype(object) if col in found else None
    df = df.map(lambda v: None if isinstance(v, str) and not v.strip() else v)
    df["source_file"] = path.name
    # table rows are 0-based, people look for the row number Excel shows
    df["source_row"] = body.index + 1
    log.info("%s: %d rows, header on row %d", path.name, len(df), header + 1)
    return df


def find_files(source):
    source = Path(source)
    if source.is_file():
        return [source]
    # "~$" files are the lock files Excel leaves next to an open workbook
    return sorted(
        p for p in source.iterdir()
        if p.suffix.lower() in (".csv", ".xlsx", ".xlsm") and not p.name.startswith("~$")
    )


def load(source):
    files = find_files(source)
    if not files:
        raise ValueError(f"no .csv or .xlsx files found in {source}")
    return pd.concat([read_file(p) for p in files], ignore_index=True)


def flag(df, mask, issue, action, detail=""):
    return [
        {
            "source_file": df.at[i, "source_file"],
            "source_row": int(df.at[i, "source_row"]),
            "order_id": clean_id(df.at[i, "order_id"]),
            "issue": issue,
            "action": action,
            "detail": detail if isinstance(detail, str) else detail.get(i, ""),
        }
        for i in df.index[mask.to_numpy()]
    ]


def drop(df, mask, issues, issue, detail=""):
    issues += flag(df, mask, issue, "dropped", detail)
    return df[~mask].copy()


def file_note(name, issue, detail):
    return {"source_file": name, "source_row": None, "order_id": None,
            "issue": issue, "action": "fixed", "detail": detail}


def empty_columns(df):
    return {(name, col) for name, part in df.groupby("source_file") for col in STANDARD_COLUMNS
            if part[col].isna().all()}


def clean(raw, dayfirst=True):
    df = raw.copy()
    df["where"] = df["source_file"] + " row " + df["source_row"].astype(str)
    issues = []

    # a column that is empty for a whole file was never exported, which gets
    # one note per file instead of one per row
    no_column = empty_columns(df)
    lacking = {col: [name for name, c in sorted(no_column) if c == col] for col in STANDARD_COLUMNS}

    df = drop(df, df[STANDARD_COLUMNS].isna().all(axis=1), issues, "blank row")

    first = df.groupby(STANDARD_COLUMNS, dropna=False, sort=False)["where"].transform("first")
    df = drop(df, df.duplicated(STANDARD_COLUMNS), issues, "exact duplicate", "same as " + first)

    df["order_id"] = df["order_id"].map(clean_id)
    no_id = df["order_id"].isna()
    placeholder = "NOID-" + df["source_file"].str.rsplit(".", n=1).str[0] + "-" + df["source_row"].astype(str)
    df.loc[no_id, "order_id"] = placeholder
    issues += flag(df, no_id & ~df["source_file"].isin(lacking["order_id"]),
                   "missing order id", "fixed", "gave it a placeholder id")
    for name in lacking["order_id"]:
        issues.append(file_note(name, "no order id column", "every row got a placeholder id"))

    for col in ["branch", "product", "category"]:
        before = df[col]
        tidy = before.map(clean_text)
        names = canonical_names(tidy)
        df[col] = tidy.map(lambda v: names[v.casefold()] if isinstance(v, str) else None)
        changed = before.notna() & (before.astype(str) != df[col].astype(str))
        detail = {i: f"'{before[i]}' -> '{df.at[i, col]}'" for i in df.index[changed.to_numpy()]}
        issues += flag(df, changed, f"inconsistent {col} name", "fixed", detail)

    for col in ["branch", "product"]:
        df = drop(df, df[col].isna(), issues, f"missing {col}")

    parsed = df["date"].map(lambda v: parse_date(v, dayfirst))
    no_date = df["date"].isna()
    unreadable = parsed.isna() & ~no_date
    detail = {}
    for i in df.index[unreadable.to_numpy()]:
        value = df.at[i, "date"]
        hint = ""
        if parse_date(value, not dayfirst) is not None:
            hint = " (valid if read month-first, see --month-first)" if dayfirst else " (valid if read day-first)"
        detail[i] = f"'{value}'{hint}"
    df["date"] = pd.to_datetime(parsed)
    df = drop(df, no_date, issues, "missing date")
    df = drop(df, unreadable[df.index], issues, "unreadable date", detail)

    for col in ["quantity", "unit_price", "total"]:
        df[col] = df[col].map(parse_number).astype(float)

    qty, price, total = df["quantity"], df["unit_price"], df["total"]
    fix = qty.isna() & price.gt(0) & total.notna()
    detail = {i: f"{total[i]:g} / {price[i]:g}" for i in df.index[fix.to_numpy()]}
    df.loc[fix, "quantity"] = (total / price)[fix].round(3)
    issues += flag(df, fix, "missing quantity", "fixed", detail)

    qty = df["quantity"]
    fix = price.isna() & qty.gt(0) & total.notna()
    detail = {i: f"{total[i]:g} / {qty[i]:g}" for i in df.index[fix.to_numpy()]}
    df.loc[fix, "unit_price"] = (total / qty)[fix].round(2)
    issues += flag(df, fix, "missing unit price", "fixed", detail)

    df = drop(df, df["quantity"].isna(), issues, "missing quantity", "no total and unit price to work it out from")
    df = drop(df, df["unit_price"].isna(), issues, "missing unit price", "no total and quantity to work it out from")

    bad_qty = df["quantity"].le(0)
    detail = {i: f"quantity {df.at[i, 'quantity']:g}" for i in df.index[bad_qty.to_numpy()]}
    df = drop(df, bad_qty, issues, "zero or negative quantity", detail)

    no_total = df["total"].isna()
    df.loc[no_total, "total"] = df["quantity"] * df["unit_price"]
    issues += flag(df, no_total & ~df["source_file"].isin(lacking["total"]),
                   "missing total", "fixed", "calculated as quantity x unit price")
    for name in lacking["total"]:
        issues.append(file_note(name, "no total column", "calculated as quantity x unit price"))

    no_cat = df["category"].isna()
    known = df[~no_cat].groupby("product")["category"].agg(lambda s: s.mode().iloc[0])
    df.loc[no_cat, "category"] = df.loc[no_cat, "product"].map(lambda p: known.get(p, "Uncategorised"))
    detail = {i: f"set to '{df.at[i, 'category']}' from other rows for this product"
              for i in df.index[no_cat.to_numpy()]}
    issues += flag(df, no_cat & ~df["source_file"].isin(lacking["category"]), "missing category", "fixed", detail)
    for name in lacking["category"]:
        issues.append(file_note(name, "no category column", "filled in from the same product in other files"))

    first = df.groupby(LINE_KEY, sort=False)["where"].transform("first")
    df = drop(df, df.duplicated(LINE_KEY), issues, "repeated order line", "same order and product as " + first)

    df["total"] = df["total"].round(2)
    if (df["quantity"] % 1 == 0).all():
        df["quantity"] = df["quantity"].astype(int)
    clean_df = (
        df.sort_values(["date", "branch", "order_id"], kind="stable")
        [STANDARD_COLUMNS + ["source_file", "source_row"]]
        .reset_index(drop=True)
    )
    issues_df = pd.DataFrame(issues, columns=ISSUE_COLUMNS)
    issues_df = issues_df.sort_values(["source_file", "source_row"], na_position="first", kind="stable")
    return clean_df, issues_df.reset_index(drop=True)


def filter_month(df, month):
    return df[df["date"].dt.strftime("%Y-%m") == month].reset_index(drop=True)
