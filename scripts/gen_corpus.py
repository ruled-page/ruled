"""
Builds the workbooks in tests/corpus, which the round-trip checks save and
validate (npm run validate:saves). They cover what templates a host hands
out tend to use and the samples don't: a protected sheet with unlocked input
cells, an Excel table with a calculated column, a totals row and structured
references, dynamic array formulas, LET and LAMBDA, defined names, links,
hidden and grouped rows, and the 1904 date system.

XlsxWriter writes files the way Excel does, but doesn't calculate, so each
formula is given the value Excel would have saved, worked out here.

    python3 -m venv .venv && .venv/bin/pip install xlsxwriter
    .venv/bin/python scripts/gen_corpus.py
"""

import datetime as dt
from pathlib import Path

import xlsxwriter

OUT = Path(__file__).resolve().parent.parent / "tests" / "corpus"
ORDERS = [
    ("Bolts", "Hardware", 120, 0.25),
    ("Nuts", "Hardware", 300, 0.1),
    ("Cable", "Electrical", 15, 4.5),
    ("Switch", "Electrical", 8, 12.0),
    ("Gloves", "Safety", 40, 3.75),
]


def templates():
    wb = xlsxwriter.Workbook(OUT / "templates.xlsx", {"use_future_functions": True})
    bold = wb.add_format({"bold": True})
    money = wb.add_format({"num_format": "#,##0.00"})
    unlocked = wb.add_format({"locked": False, "bg_color": "#FFF2CC"})
    date = wb.add_format({"num_format": "yyyy-mm-dd", "locked": False, "bg_color": "#FFF2CC"})

    # A form to fill in: everything is locked except the yellow input cells.
    form = wb.add_worksheet("Request")
    form.protect()
    form.set_column("A:A", 22)
    form.set_column("B:B", 18)
    form.write("A1", "Purchase request", bold)
    form.write("A3", "Requested by")
    form.write("B3", "", unlocked)
    form.write("A4", "Needed by")
    form.write_datetime("B4", dt.datetime(2026, 11, 2), date)
    form.write("A5", "Budget")
    form.write_number("B5", 2500, wb.add_format({"num_format": "#,##0", "locked": False, "bg_color": "#FFF2CC"}))
    form.write("A6", "Order total")
    total = sum(q * p for _, _, q, p in ORDERS)
    # Its formula is hidden from the formula bar while the sheet is protected.
    form.write_formula("B6", "=SUM(Orders[Total])", wb.add_format({"num_format": "#,##0.00", "hidden": True}), total)
    form.write("A7", "Within budget")
    form.write_formula("B7", "=IF(B6<=Budget,\"Yes\",\"No\")", None, "Yes" if total <= 2500 else "No")
    form.data_validation("B5", {"validate": "integer", "criteria": ">", "value": 0})
    wb.define_name("Budget", "=Request!$B$5")

    # An Excel table with a calculated column, a totals row and structured references.
    lines = wb.add_worksheet("Lines")
    lines.freeze_panes(1, 0)
    rows = [[n, c, q, p, q * p] for n, c, q, p in ORDERS]
    lines.add_table(
        0, 0, len(rows) + 1, 4,
        {
            "name": "Orders",
            "total_row": True,
            "data": [r[:4] for r in rows],
            "columns": [
                {"header": "Item", "total_string": "Total"},
                {"header": "Category"},
                {"header": "Qty", "total_function": "sum", "total_value": sum(r[2] for r in rows)},
                {"header": "Price", "format": money},
                {"header": "Total", "formula": "=[@Qty]*[@Price]", "format": money, "total_function": "sum", "total_value": total},
            ],
        },
    )
    for i, r in enumerate(rows):
        lines.write_formula(i + 1, 4, "=[@Qty]*[@Price]", money, r[4])
    lines.write_url("G1", "https://example.com/catalogue", string="Catalogue")
    lines.write_url("G2", "internal:Request!B3", string="Back to the request")
    # Notes rows, grouped and collapsed.
    lines.write("A10", "Notes")
    for r in range(10, 13):
        lines.write(r, 0, f"Note {r - 9}")
        lines.set_row(r, None, None, {"level": 1, "hidden": True})
    lines.set_row(13, None, None, {"collapsed": True})

    # Dynamic arrays, LET and LAMBDA, each with Excel's saved results.
    arrays = wb.add_worksheet("Arrays")
    arrays.write("A1", "Categories", bold)
    cats = sorted({c for _, c, _, _ in ORDERS})
    arrays.write_dynamic_array_formula(f"A2:A{1 + len(cats)}", "=SORT(UNIQUE(Orders[Category]))", None, cats[0])
    for i, c in enumerate(cats[1:], start=2):
        arrays.write_string(i, 0, c)
    arrays.write("B1", "Spend", bold)
    spend = {c: sum(q * p for _, cc, q, p in ORDERS if cc == c) for c in cats}
    arrays.write_dynamic_array_formula(f"B2:B{1 + len(cats)}", "=SUMIFS(Orders[Total],Orders[Category],A2#)", None, spend[cats[0]])
    for i, c in enumerate(cats[1:], start=2):
        arrays.write_number(i, 1, spend[c])
    arrays.write("D1", "Biggest line", bold)
    biggest = max(rows, key=lambda r: r[4])
    arrays.write_formula("D2", "=XLOOKUP(MAX(Orders[Total]),Orders[Total],Orders[Item])", None, biggest[0])
    arrays.write("D3", "With tax")
    arrays.write_formula("E3", "=LET(net,SUM(Orders[Total]),rate,0.2,net*(1+rate))", money, total * 1.2)
    arrays.write("D4", "Doubled")
    arrays.write_formula("E4", "=LAMBDA(x,x*2)(Budget)", None, 5000)
    arrays.write("D5", "Hardware lines")
    arrays.write_dynamic_array_formula("E5:E6", '=FILTER(Orders[Item],Orders[Category]="Hardware")', None, "Bolts")
    arrays.write_string(5, 4, "Nuts")
    wb.close()


def dates_1904():
    wb = xlsxwriter.Workbook(OUT / "dates-1904.xlsx", {"date_1904": True})
    ws = wb.add_worksheet("Schedule")
    fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
    ws.set_column("A:C", 14)
    ws.write_row("A1", ["Start", "Days", "End"])
    start = dt.datetime(2026, 3, 1)
    for i, days in enumerate([10, 45, 90]):
        ws.write_datetime(i + 1, 0, start + dt.timedelta(days=30 * i), fmt)
        ws.write_number(i + 1, 1, days)
        end = start + dt.timedelta(days=30 * i + days)
        serial = (end - dt.datetime(1904, 1, 1)).days
        ws.write_formula(i + 1, 2, f"=A{i + 2}+B{i + 2}", fmt, serial)
    wb.close()


OUT.mkdir(parents=True, exist_ok=True)
templates()
dates_1904()
print(f"Wrote {OUT}/templates.xlsx and dates-1904.xlsx")
