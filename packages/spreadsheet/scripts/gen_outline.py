"""
Builds tests/corpus/outline.xlsx: rows and columns grouped as Excel groups
them. "Report" has row groups two levels deep with their totals below (one
collapsed), and column groups with their totals to the right (one
collapsed); "Above" has its summaries above and left of its groups, as
<outlinePr> can say; "Locked" is protected, letting rows be formatted (and
so grouped, shown and hidden) but not columns. The unit tests and the
round-trip checks read it.

    python3 -m venv .venv && .venv/bin/pip install xlsxwriter
    .venv/bin/python scripts/gen_outline.py
"""

import datetime as dt
from pathlib import Path

import xlsxwriter

OUT = Path(__file__).resolve().parent.parent / "tests" / "corpus" / "outline.xlsx"
MONTHS = ["Jan", "Feb", "Mar", "Q1", "Apr", "May", "Jun", "Q2", "Half year"]
LINES = {"North A": [10, 12, 14, 9, 8, 11], "North B": [5, 6, 7, 8, 9, 10], "South A": [3, 4, 5, 6, 7, 8], "South B": [20, 18, 16, 14, 12, 10]}

wb = xlsxwriter.Workbook(str(OUT))
wb.set_properties({"created": dt.datetime(2026, 10, 1, 9, 0, 0)})
bold = wb.add_format({"bold": True})


def months(ws, row, values, fmt=None):
    """Jan to Jun in B:D and F:H, with each quarter's total after it and the half year's in J."""
    ws.write_row(row, 1, values[:3], fmt)
    ws.write_row(row, 5, values[3:], fmt)
    r = row + 1
    ws.write_formula(row, 4, f"=SUM(B{r}:D{r})", fmt, sum(values[:3]))
    ws.write_formula(row, 8, f"=SUM(F{r}:H{r})", fmt, sum(values[3:]))
    ws.write_formula(row, 9, f"=E{r}+I{r}", fmt, sum(values))


report = wb.add_worksheet("Report")
report.write_row(0, 0, ["Region", *MONTHS], bold)
# Rows 2-3 and 5-6 are each region's lines (level 2), under its total (level 1); row 8 is the grand total.
for i, (name, values) in enumerate(LINES.items()):
    row = 1 + i + i // 2
    report.write(row, 0, name)
    months(report, row, values)
    report.set_row(row, None, None, {"level": 2, "hidden": row >= 4})
for row, region, lines in [(3, "North", ["North A", "North B"]), (6, "South", ["South A", "South B"])]:
    report.write(row, 0, region, bold)
    months(report, row, [sum(LINES[n][k] for n in lines) for k in range(6)], bold)
    report.set_row(row, None, None, {"level": 1, "collapsed": row == 6})
report.write(7, 0, "Total", bold)
months(report, 7, [sum(v[k] for v in LINES.values()) for k in range(6)], bold)
# Each quarter's months are grouped beside its total; the second quarter's are collapsed.
report.set_column("A:A", 12)
report.set_column("B:D", 8, None, {"level": 1})
report.set_column("E:E", 8)
report.set_column("F:H", 8, None, {"level": 1, "hidden": True})
report.set_column("I:I", 8, None, {"collapsed": True})

above = wb.add_worksheet("Above")
above.outline_settings(True, False, False, False)
above.write_row(0, 0, ["Team", "Ann", "Bo"], bold)
above.write(1, 0, "Staff", bold)
for row, name in enumerate(["Ann", "Bo", "Cy"], start=2):
    above.write(row, 0, name)
    above.set_row(row, None, None, {"level": 1})
above.set_column("B:C", 8, None, {"level": 1})

locked = wb.add_worksheet("Locked")
locked.write(0, 0, "Checked", bold)
for row in range(1, 4):
    locked.write(row, 0, f"Step {row}")
    locked.set_row(row, None, None, {"level": 1})
locked.protect("", {"format_rows": True})
wb.close()
print(f"Wrote {OUT}")
