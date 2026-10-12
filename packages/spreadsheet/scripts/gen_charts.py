"""
Builds tests/corpus/charts.xlsx: one chart of each kind the editor draws
(column, bar, line, area, pie, doughnut and scatter, clustered, stacked and
100% stacked, a column and line combination on two axes), one it doesn't
(radar), charts that read another sheet and a defined name, a line over a
range with a hidden row, and a chart sheet. The unit tests read each chart's
definition from it.

    python3 -m venv .venv && .venv/bin/pip install xlsxwriter
    .venv/bin/python scripts/gen_charts.py
"""

import datetime as dt
from pathlib import Path

import xlsxwriter

OUT = Path(__file__).resolve().parent.parent / "tests" / "corpus" / "charts.xlsx"
QUARTERS = ["Q1", "Q2", "Q3", "Q4"]
SALES = {"North": [120, 150, 130, 170], "South": [90, 110, 100, 120], "East": [60, 80, 70, 90]}
MARGIN = [0.12, 0.15, 0.1, 0.18]
VISITS = [410, 380, 520, 470, 610, 590]
SCATTER = [(12, 210), (15, 260), (18, 300), (22, 390), (26, 420), (29, 480)]

wb = xlsxwriter.Workbook(str(OUT))
bold = wb.add_format({"bold": True})
pct = wb.add_format({"num_format": "0%"})
money = wb.add_format({"num_format": '"$"#,##0'})
month = wb.add_format({"num_format": "mmm yyyy"})

data = wb.add_worksheet("Data")
charts = wb.add_worksheet("Charts")
other = wb.add_worksheet("Sales Data")
trend = wb.add_chartsheet("Trend")

# Quarterly sales by region, with a margin to plot on a second axis.
data.write_row("A1", ["Quarter", *SALES, "Margin"], bold)
for i, q in enumerate(QUARTERS):
    data.write_string(1 + i, 0, q)
    for j, values in enumerate(SALES.values()):
        data.write_number(1 + i, 1 + j, values[i], money)
    data.write_number(1 + i, 4, MARGIN[i], pct)

# Monthly visits, with April hidden: Excel leaves hidden rows out of charts.
data.write_row("A10", ["Month", "Visits"], bold)
for i, v in enumerate(VISITS):
    data.write_datetime(10 + i, 0, dt.datetime(2026, 1 + i, 1), month)
    data.write_number(10 + i, 1, v)
data.set_row(13, None, None, {"hidden": True})

# Temperature against sales, for the scatter charts.
data.write_row("H1", ["Temperature", "Ice cream"], bold)
for i, (x, y) in enumerate(SCATTER):
    data.write_number(1 + i, 7, x)
    data.write_number(1 + i, 8, y)
wb.define_name("Temperatures", "=Data!$H$2:$H$7")

other.write_row("A1", ["Product", "Units"], bold)
for i, (p, u) in enumerate([("Widget", 40), ("Gadget", 25), ("Gizmo", 35)]):
    other.write_string(1 + i, 0, p)
    other.write_number(1 + i, 1, u)

cats = "=Data!$A$2:$A$5"


def regional(chart, colors=False):
    for j, region in enumerate(SALES):
        col = chr(ord("B") + j)
        series = {"name": f"=Data!${col}$1", "categories": cats, "values": f"=Data!${col}$2:${col}$5"}
        if colors and j == 0:
            series["fill"] = {"color": "#1F77B4"}
        chart.add_series(series)
    return chart


placed = 0


def place(chart, title):
    global placed
    chart.set_title({"name": title})
    charts.insert_chart(1 + (placed // 3) * 16, 1 + (placed % 3) * 8, chart)
    placed += 1


column = regional(wb.add_chart({"type": "column"}), colors=True)
column.set_x_axis({"name": "Quarter"})
column.set_y_axis({"name": "Units", "min": 0, "max": 200, "major_unit": 50})
column.set_legend({"position": "bottom"})
place(column, "Quarterly sales")

place(regional(wb.add_chart({"type": "column", "subtype": "stacked"})), "Stacked columns")
place(regional(wb.add_chart({"type": "column", "subtype": "percent_stacked"})), "Share by quarter")
bar = regional(wb.add_chart({"type": "bar"}))
bar.set_legend({"position": "top"})
place(bar, "Horizontal bars")
place(regional(wb.add_chart({"type": "bar", "subtype": "stacked"})), "Stacked bars")

line = wb.add_chart({"type": "line"})
line.add_series({"name": "=Data!$B$1", "categories": cats, "values": "=Data!$B$2:$B$5", "marker": {"type": "circle"}, "line": {"color": "#C00000", "width": 2.5}})
line.add_series({"name": "=Data!$C$1", "categories": cats, "values": "=Data!$C$2:$C$5", "smooth": True, "marker": {"type": "none"}, "line": {"dash_type": "dash"}})
line.set_legend({"position": "right"})
place(line, "Sales trend")

visits = wb.add_chart({"type": "line"})
visits.add_series({"name": "Visits", "categories": "=Data!$A$11:$A$16", "values": "=Data!$B$11:$B$16"})
visits.show_blanks_as("span")
place(visits, "Visits by month")

place(regional(wb.add_chart({"type": "area"})), "Area")
place(regional(wb.add_chart({"type": "area", "subtype": "stacked"})), "Stacked area")

pie = wb.add_chart({"type": "pie"})
pie.add_series({"name": "North", "categories": cats, "values": "=Data!$B$2:$B$5", "points": [{"fill": {"color": "#FF0000"}}, None, {"fill": {"color": "#00B050"}}]})
pie.set_rotation(90)
place(pie, "North by quarter")

doughnut = wb.add_chart({"type": "doughnut"})
doughnut.add_series({"name": "Units", "categories": "='Sales Data'!$A$2:$A$4", "values": "='Sales Data'!$B$2:$B$4"})
doughnut.set_hole_size(40)
doughnut.set_legend({"position": "left"})
place(doughnut, "Units by product")

scatter = wb.add_chart({"type": "scatter"})
scatter.add_series({"name": "Ice cream", "categories": "=Data!$H$2:$H$7", "values": "=Data!$I$2:$I$7"})
scatter.set_x_axis({"name": "Temperature", "min": 10})
place(scatter, "Ice cream sales")

lines = wb.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
lines.add_series({"name": "Ice cream", "categories": "=Temperatures", "values": "=Data!$I$2:$I$7"})
place(lines, "Sales by temperature")

combo = wb.add_chart({"type": "column"})
combo.add_series({"name": "=Data!$B$1", "categories": cats, "values": "=Data!$B$2:$B$5"})
margin = wb.add_chart({"type": "line"})
margin.add_series({"name": "=Data!$E$1", "categories": cats, "values": "=Data!$E$2:$E$5", "y2_axis": True})
margin.set_y2_axis({"name": "Margin", "num_format": "0%"})
combo.combine(margin)
place(combo, "Sales and margin")

radar = regional(wb.add_chart({"type": "radar"}))
place(radar, "Radar")

untitled = wb.add_chart({"type": "column"})
untitled.add_series({"name": "=Data!$D$1", "categories": cats, "values": "=Data!$D$2:$D$5"})
charts.insert_chart(1 + (placed // 3) * 16, 1 + (placed % 3) * 8, untitled)

trend.set_chart(regional(wb.add_chart({"type": "line"})))

wb.close()
print(f"Wrote {OUT}")
