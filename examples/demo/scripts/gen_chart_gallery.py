"""
Builds public/samples/chart-gallery.xlsx: one sheet of data with a chart of
each kind the editor draws beyond the everyday ones (radar, filled radar,
bubble, two stock charts, and 3-D column, line, area and pie charts drawn
flat), plus a pie of pie, drawn as one pie, and a surface chart, which stays
a placeholder and is kept as it is.

XlsxWriter writes neither bubble nor 3-D charts, so this one uses openpyxl.

    python3 -m venv .venv && .venv/bin/pip install openpyxl
    .venv/bin/python scripts/gen_chart_gallery.py
"""

import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart3D,
    BarChart3D,
    BubbleChart,
    LineChart3D,
    PieChart3D,
    ProjectedPieChart,
    RadarChart,
    Reference,
    Series,
    StockChart,
    SurfaceChart,
)
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.updown_bars import UpDownBars
from openpyxl.styles import Font, PatternFill

OUT = Path(__file__).resolve().parent.parent / "public" / "samples" / "chart-gallery.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Gallery"
head = Font(bold=True)
band = PatternFill("solid", fgColor="E2EFDA")


def heading(row: int, col: int, values: list) -> None:
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = head
        c.fill = band


# --- Radar: two players' skills -------------------------------------------------
heading(1, 1, ["Skill", "Ana", "Ben"])
for i, row in enumerate([("Speed", 8, 6), ("Strength", 5, 9), ("Stamina", 7, 7), ("Skill", 9, 5), ("Teamwork", 6, 8)], start=2):
    for j, v in enumerate(row):
        ws.cell(row=i, column=1 + j, value=v)

# --- Bubble: products by price, margin and units sold ----------------------------
heading(1, 5, ["Product", "Price", "Margin", "Units"])
for i, row in enumerate([("Widget", 12.5, 0.32, 1200), ("Gadget", 30, 0.41, 450), ("Gizmo", 7.25, 0.18, 2600), ("Doohickey", 18, 0.27, 900)], start=2):
    for j, v in enumerate(row):
        ws.cell(row=i, column=5 + j, value=v)

# --- Stock: ten days of a share's open, high, low and close -----------------------
heading(1, 10, ["Day", "Open", "High", "Low", "Close"])
prices = [(41.2, 42.9, 40.8, 42.5), (42.5, 43.1, 41.6, 41.9), (41.9, 42.4, 40.2, 40.6), (40.6, 41.8, 40.1, 41.5), (41.5, 43.6, 41.2, 43.3),
          (43.3, 44.0, 42.7, 42.9), (42.9, 43.4, 41.9, 43.2), (43.2, 45.1, 43.0, 44.8), (44.8, 45.0, 43.5, 43.9), (43.9, 44.6, 43.1, 44.4)]
start = dt.date(2026, 10, 1)
for i, (o, h, lo, c) in enumerate(prices, start=2):
    d = ws.cell(row=i, column=10, value=start + dt.timedelta(days=i - 2))
    d.number_format = "d mmm"
    for j, v in enumerate((o, h, lo, c)):
        ws.cell(row=i, column=11 + j, value=v)

# --- Quarterly sales, for the 3-D charts --------------------------------------------
heading(14, 1, ["Quarter", "North", "South"])
for i, (q, n, s) in enumerate([("Q1", 120, 95), ("Q2", 135, 110), ("Q3", 128, 140), ("Q4", 160, 150)], start=15):
    for j, v in enumerate((q, n, s)):
        ws.cell(row=i, column=1 + j, value=v)

# --- Costs, for the pie of pie and the surface -----------------------------------------
heading(14, 5, ["Cost", "Amount"])
for i, (k, v) in enumerate([("Rent", 4200), ("Staff", 9800), ("Stock", 6100), ("Power", 640), ("Phones", 320), ("Post", 180)], start=15):
    ws.cell(row=i, column=5, value=k)
    ws.cell(row=i, column=6, value=v)

ws.column_dimensions["A"].width = 11
ws.column_dimensions["E"].width = 11
ws.column_dimensions["J"].width = 9


def place(chart, anchor: str, title: str) -> None:
    chart.title = title
    chart.width = 12
    chart.height = 7
    ws.add_chart(chart, anchor)


skills = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=6)
skill_names = Reference(ws, min_col=1, min_row=2, max_row=6)

radar = RadarChart()
radar.type = "marker"
radar.add_data(skills, titles_from_data=True)
radar.set_categories(skill_names)
# openpyxl turns each series' markers off; a radar with markers shows them.
for s in radar.series:
    s.marker.symbol = "circle"
place(radar, "A22", "Skills (radar)")

filled = RadarChart()
filled.type = "filled"
filled.add_data(skills, titles_from_data=True)
filled.set_categories(skill_names)
place(filled, "H22", "Skills (filled radar)")

bubble = BubbleChart()
bubble.style = 18
bubble.append(Series(values=Reference(ws, min_col=7, min_row=2, max_row=5), xvalues=Reference(ws, min_col=6, min_row=2, max_row=5),
                     zvalues=Reference(ws, min_col=8, min_row=2, max_row=5), title="Products"))
bubble.x_axis.title = "Price"
bubble.y_axis.title = "Margin"
place(bubble, "O22", "Price, margin and units (bubble)")

days = Reference(ws, min_col=10, min_row=2, max_row=11)


def stock(first_col: int, anchor: str, title: str, up_down: bool) -> None:
    chart = StockChart()
    chart.add_data(Reference(ws, min_col=first_col, min_row=1, max_col=14, max_row=11), titles_from_data=True)
    chart.set_categories(days)
    for s in chart.series:
        s.graphicalProperties.line.noFill = True
    # The close is marked on high-low-close; open-high-low-close shows it as the bar's end.
    if not up_down:
        chart.series[-1].marker.symbol = "dash"
        chart.series[-1].marker.size = 7
    chart.hiLowLines = ChartLines()
    if up_down:
        chart.upDownBars = UpDownBars()
    place(chart, anchor, title)


stock(12, "A37", "High, low, close (stock)", False)
stock(11, "H37", "Open, high, low, close (stock)", True)

quarters = Reference(ws, min_col=1, min_row=15, max_row=18)
sales = Reference(ws, min_col=2, min_row=14, max_col=3, max_row=18)
for cls, anchor, title in [(BarChart3D, "O37", "Sales (3-D column)"), (LineChart3D, "A52", "Sales (3-D line)"), (AreaChart3D, "H52", "Sales (3-D area)")]:
    chart = cls()
    chart.add_data(sales, titles_from_data=True)
    chart.set_categories(quarters)
    place(chart, anchor, title)

costs = Reference(ws, min_col=6, min_row=14, max_row=20)
cost_names = Reference(ws, min_col=5, min_row=15, max_row=20)
pie3d = PieChart3D()
pie3d.add_data(costs, titles_from_data=True)
pie3d.set_categories(cost_names)
place(pie3d, "O52", "Costs (3-D pie)")

of_pie = ProjectedPieChart()
of_pie.type = "pie"
of_pie.splitType = "pos"
of_pie.splitPos = 3
of_pie.add_data(costs, titles_from_data=True)
of_pie.set_categories(cost_names)
place(of_pie, "A67", "Costs (pie of pie)")

surface = SurfaceChart()
surface.add_data(sales, titles_from_data=True)
surface.set_categories(quarters)
place(surface, "H67", "Sales (surface, kept as it is)")

OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print(f"Wrote {OUT}")
