"""Generates examples/example.xlsx — a small workbook exercising the engine.

Run it with openpyxl installed:

    python examples/build_example.py

The file is committed, so you only need this if you want to change the example.
"""

import pathlib

from openpyxl import Workbook

HERE = pathlib.Path(__file__).parent


def build():
    wb = Workbook()

    items = wb.active
    items.title = "Items"
    items["A1"], items["B1"], items["C1"], items["D1"] = "Item", "Qty", "Unit", "Total"
    rows = [("Widget", 4, 25.0), ("Gadget", 10, 7.5), ("Doohickey", 3, 42.0)]
    for i, (name, qty, unit) in enumerate(rows, start=2):
        items[f"A{i}"], items[f"B{i}"], items[f"C{i}"] = name, qty, unit
        items[f"D{i}"] = f"=B{i}*C{i}"          # binary op over two references
    items["D5"] = "=SUM(D2:D4)"                  # range + aggregate
    items["D6"] = "=ROUNDUP(D5*0.21,2)"          # nested call over a reference
    items["D7"] = "=D5+D6"

    summary = wb.create_sheet("Summary")
    summary["A1"] = "Subtotal"
    summary["B1"] = "=Items!D5"                  # cross-sheet reference
    summary["A2"] = "Tax"
    summary["B2"] = "=Items!D6"
    summary["A3"] = "Total"
    summary["B3"] = "=B1+B2"
    summary["A4"] = "Free shipping"
    summary["B4"] = '=IF(B3>100,"yes","no")'     # conditional
    summary["A5"] = "Line count"
    summary["B5"] = "=SUM(Items!B2:B4)"          # aggregate across sheets

    # a sheet name with a space, to exercise normalize_sheet_name()
    report = wb.create_sheet("Cost Report")
    report["A1"] = "Grand total"
    report["B1"] = "=Summary!B3"
    report["A2"] = "Per item"
    report["B2"] = "=ROUNDDOWN(B1/Summary!B5,2)"

    out = HERE / "example.xlsx"
    wb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    build()
