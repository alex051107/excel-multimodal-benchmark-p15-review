#!/usr/bin/env python3
"""Independent held-out normalization, join, and exception replay."""
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "input_files"


def recompute():
    products = {row["product_code"]: row for row in csv.DictReader((ROOT / "product_master.csv").open())}
    locations = {row["location_code"]: row for row in csv.DictReader((ROOT / "location_master.csv").open())}
    seen, clean, joined, exceptions = set(), [], [], []
    for raw in csv.DictReader((ROOT / "raw_orders.csv").open()):
        product = raw["product_code"].strip().upper()
        location = raw["location_code"].strip().upper()
        try:
            units = int(raw["units"])
        except ValueError:
            units = None
        fingerprint = (product, location, units, raw["order_date"])
        duplicate = fingerprint in seen if units is not None else False
        if units is not None:
            seen.add(fingerprint)
        product_ok, location_ok = product in products, location in locations
        if units is None:
            disposition, issue = "Exception: invalid units", "INVALID_UNITS"
        elif duplicate:
            disposition, issue = "Exception: exact duplicate", "DUPLICATE"
        elif not product_ok:
            disposition, issue = "Exception: unmatched product", "UNMATCHED_PRODUCT"
        elif not location_ok:
            disposition, issue = "Exception: unmatched location", "UNMATCHED_LOCATION"
        else:
            disposition, issue = "Join", None
        clean.append((raw["order_id"], product, location, units, "Yes" if duplicate else "No", "Yes" if product_ok else "No", "Yes" if location_ok else "No", disposition))
        if issue:
            exceptions.append((raw["order_id"], issue))
        else:
            product_row, location_row = products[product], locations[location]
            joined.append((raw["order_id"], product_row["product"], location_row["region"], units, int(product_row["unit_cost"]), units * int(product_row["unit_cost"])))
    return {
        "clean": clean,
        "joined": joined,
        "exceptions": exceptions,
        "total": sum(row[5] for row in joined),
        "summary_region": "West",
        "summary_region_total": sum(row[5] for row in joined if row[2] == "West"),
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
