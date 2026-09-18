#!/usr/bin/env python3
"""
One-time migration: splits the DASHBOARD_DATA blob currently embedded in
kck_sales_dashboard_V11.html into the per-month JSON files the new backend
reads/writes (backend/data/...).

Usage:
    python migrate_existing.py "<path to kck_sales_dashboard_V11.html>"

Safe to re-run: it always overwrites backend/data/ from whatever is in the
HTML right now, it never reads backend/data/ back in.
"""

import sys
import os
import json
from collections import defaultdict

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")


def load_dashboard_data(html_path):
    with open(html_path, encoding="utf-8") as f:
        for line in f:
            if line.startswith("const DASHBOARD_DATA"):
                raw = line.strip()[len("const DASHBOARD_DATA = "):]
                if raw.endswith(";"):
                    raw = raw[:-1]
                return json.loads(raw)
    raise RuntimeError("Could not find 'const DASHBOARD_DATA = ...;' in " + html_path)


def month_key(date_str):
    return date_str[:7]  # "2026-07-15" -> "2026-07"


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def split_by_month(rows, subdir):
    by_month = defaultdict(list)
    for row in rows:
        by_month[month_key(row["date"])].append(row)
    for month, month_rows in by_month.items():
        write_json(os.path.join(DATA_DIR, subdir, f"{month}.json"), month_rows)
    return {m: len(r) for m, r in by_month.items()}


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)

    html_path = sys.argv[1]
    data = load_dashboard_data(html_path)

    write_json(os.path.join(DATA_DIR, "meta.json"), data["meta"])

    bills_by_month = split_by_month(data["allBills"], "bills")
    sales_by_month = split_by_month(data["salesData"], "sales")
    discounts_by_month = split_by_month(data["discountData"], "discounts")
    cancellations_by_month = split_by_month(data.get("cancelData", []), "cancellations")

    print(f"meta.json written")
    print(f"bills:         {dict(sorted(bills_by_month.items()))}")
    print(f"sales:         {dict(sorted(sales_by_month.items()))}")
    print(f"discounts:     {dict(sorted(discounts_by_month.items()))}")
    print(f"cancellations: {dict(sorted(cancellations_by_month.items()))} (carried over as-is, no source report yet)")
    print(f"\nWrote to: {DATA_DIR}")


if __name__ == "__main__":
    main()
