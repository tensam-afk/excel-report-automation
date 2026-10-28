import argparse
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

import mailer
from clean import clean, filter_month, load
from report import kpis, write_report

log = logging.getLogger("report")


def build(source, month=None, dayfirst=True):
    raw = load(source)
    data, issues = clean(raw, dayfirst=dayfirst)
    dropped = issues[issues["action"] == "dropped"]
    fixed = issues[(issues["action"] == "fixed") & issues["source_row"].notna()]
    # a row can be fixed and later dropped as a duplicate, count it as dropped only
    fixed_rows = set(zip(fixed["source_file"], fixed["source_row"]))
    fixed_rows -= set(zip(dropped["source_file"], dropped["source_row"]))
    stats = {
        "files": raw["source_file"].value_counts(sort=False).to_dict(),
        "rows_read": len(raw),
        "dropped": len(dropped),
        "fixed": len(fixed_rows),
        "outside_period": 0,
    }

    if month:
        inside = filter_month(data, month)
        stats["outside_period"] = len(data) - len(inside)
        data = inside
        start = pd.Timestamp(month + "-01")
        stats["start"], stats["end"] = start, start + pd.offsets.MonthEnd(0)
        stats["period"] = f"{start:%B %Y}"
    elif not data.empty:
        stats["start"], stats["end"] = data["date"].min(), data["date"].max()
        stats["period"] = f"{stats['start']:%d %b %Y} to {stats['end']:%d %b %Y}"
    return data, issues, stats


def email_body(data, stats, currency):
    k = kpis(data)
    return "\n".join([
        f"The sales report for {stats['period']} is attached.",
        "",
        f"Revenue:             {currency}{k['revenue']:,.2f}",
        f"Orders:              {k['orders']:,}",
        f"Average order value: {currency}{k['avg_order']:,.2f}",
        f"Top branch:          {k['top_branch'][0]}",
        f"Top product:         {k['top_product'][0]}",
        "",
        f"{stats['dropped']} of {stats['rows_read']} rows were dropped while cleaning the data. "
        "The Data Issues sheet lists each one and why.",
    ])


def month_arg(text):
    try:
        datetime.strptime(text, "%Y-%m")
    except ValueError:
        raise argparse.ArgumentTypeError("use YYYY-MM, for example 2026-09")
    return text


def main(argv=None):
    p = argparse.ArgumentParser(description="Clean messy sales exports and build an Excel report.")
    p.add_argument("source", help="folder of .csv/.xlsx exports, or a single file")
    p.add_argument("-o", "--out", default="output/report.xlsx", help="where to save the report (default: %(default)s)")
    p.add_argument("--month", type=month_arg, help="only report on this month, as YYYY-MM")
    p.add_argument("--currency", default="₹", help="symbol for money cells (default: rupee sign)")
    p.add_argument("--month-first", action="store_true",
                   help="read 03/09/2026 as 9 March (US style) instead of 3 September")
    p.add_argument("--email", help="send the report to this address (comma-separate several); needs SMTP_* in .env")
    args = p.parse_args(argv)

    load_dotenv()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", datefmt="%H:%M:%S")

    started = time.time()
    try:
        data, issues, stats = build(args.source, args.month, dayfirst=not args.month_first)
    except (OSError, ValueError) as e:
        log.error("%s", e)
        return 1
    if data.empty:
        log.error("no rows left to report on%s", f" for {args.month}" if args.month else "")
        return 1

    log.info("read %d rows from %d files: %d dropped, %d fixed", stats["rows_read"], len(stats["files"]),
             stats["dropped"], stats["fixed"])
    if args.month:
        log.info("%d rows fall outside %s, %d left for the report", stats["outside_period"], args.month, len(data))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        write_report(out, data, issues, stats, args.currency)
    except PermissionError:
        log.error("can't write %s, is it open in Excel?", out)
        return 1
    log.info("saved %s in %.1fs", out, time.time() - started)

    if args.email:
        if not mailer.is_configured():
            log.warning("not emailing: set SMTP_USER and SMTP_PASSWORD in .env first")
            return 0
        try:
            mailer.send_report(out, args.email, f"Sales report: {stats['period']}",
                               email_body(data, stats, args.currency))
        except Exception as e:
            log.error("email failed: %s", e)
            return 1
        log.info("emailed the report to %s", args.email)
    return 0


if __name__ == "__main__":
    sys.exit(main())
