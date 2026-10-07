"""
Builds public/samples/kitchen-sink.xlsx: a workbook that exercises most of
the features the editor has to preserve or warn about (charts, a chart
sheet, a table, conditional formatting, validation, comments, sparklines,
pictures, array and spill formulas, functions the in-app engine lacks).

XlsxWriter does not calculate formulas, so every formula is written with the
value Excel would have saved, computed here in Python.

    python3 -m venv .venv && .venv/bin/pip install xlsxwriter
    .venv/bin/python scripts/gen_kitchen_sink.py
"""

import datetime as dt
import io
import random
import struct
import zlib
from pathlib import Path

import xlsxwriter

OUT = Path(__file__).resolve().parent.parent / "public" / "samples" / "kitchen-sink.xlsx"
EPOCH = dt.date(1899, 12, 30)
REGIONS = ["North", "South", "East", "West"]
PRODUCTS = {"Widget": 12.5, "Gadget": 30.0, "Gizmo": 7.25}
TAX_RATE = 0.2
ROWS = 200


def serial(d: dt.date) -> int:
    return (d - EPOCH).days


def edate(d: dt.date, months: int) -> dt.date:
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    last = [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return dt.date(y, m, min(d.day, last))


def tiny_png(w=48, h=48, rgb=(33, 115, 70)) -> bytes:
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))

    def chunk(kind, data):
        c = struct.pack(">I", len(data)) + kind + data
        return c + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


rng = random.Random(42)
start = dt.date(2026, 1, 1)
data = []
for i in range(ROWS):
    d = start + dt.timedelta(days=int(i * 364 / ROWS))
    region = rng.choice(REGIONS)
    product = rng.choice(list(PRODUCTS))
    units = rng.randint(1, 99)
    price = PRODUCTS[product]
    data.append((d, region, product, units, price, units * price))

OUT.parent.mkdir(parents=True, exist_ok=True)
wb = xlsxwriter.Workbook(str(OUT), {"use_future_functions": True})

title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#FFFFFF", "bg_color": "#217346", "align": "center"})
header = wb.add_format({"bold": True, "bg_color": "#E2EFDA", "bottom": 1})
money = wb.add_format({"num_format": '"$"#,##0.00'})
money_bold = wb.add_format({"num_format": '"$"#,##0.00', "bold": True, "top": 1})
pct = wb.add_format({"num_format": "0.0%"})
pct_bold = wb.add_format({"num_format": "0.0%", "bold": True, "top": 1})
bold = wb.add_format({"bold": True})
date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
month_fmt = wb.add_format({"num_format": "mmm yyyy"})
wrap = wb.add_format({"text_wrap": True, "valign": "top"})
italic = wb.add_format({"italic": True})
red = wb.add_format({"font_color": "#C00000"})

summary = wb.add_worksheet("Summary")
data_ws = wb.add_worksheet("Data")
lookups = wb.add_worksheet("Lookups")
notes = wb.add_worksheet("Notes")
reference = wb.add_worksheet("Reference")
settings = wb.add_worksheet("Settings")
chartsheet = wb.add_chartsheet("Monthly Chart")

# --- Settings (hidden) -------------------------------------------------------
settings.write("A1", "Tax rate")
settings.write_number("B1", TAX_RATE, pct)
settings.hide()
wb.define_name("TaxRate", "=Settings!$B$1")

# --- Data ---------------------------------------------------------------------
data_ws.set_column("A:A", 12)
data_ws.set_column("B:C", 11)
data_ws.set_column("D:E", 9)
data_ws.set_column("F:F", 12)
data_ws.freeze_panes(1, 0)
data_ws.add_table(
    0, 0, ROWS, 5,
    {
        "name": "SalesData",
        "style": "Table Style Medium 2",
        "columns": [{"header": h} for h in ["Date", "Region", "Product", "Units", "Price", "Revenue"]],
    },
)
for i, (d, region, product, units, price, revenue) in enumerate(data, start=1):
    data_ws.write_datetime(i, 0, dt.datetime.combine(d, dt.time()), date_fmt)
    data_ws.write_string(i, 1, region)
    data_ws.write_string(i, 2, product)
    data_ws.write_number(i, 3, units)
    data_ws.write_number(i, 4, price, money)
    data_ws.write_formula(i, 5, f"=D{i + 1}*E{i + 1}", money, revenue)
data_ws.conditional_format(1, 3, ROWS, 3, {"type": "data_bar", "bar_color": "#63BE7B"})
data_ws.data_validation(1, 1, ROWS, 1, {"validate": "list", "source": REGIONS})

# --- Summary ------------------------------------------------------------------
summary.set_column("A:A", 18)
summary.set_column("B:F", 13)
summary.set_column("G:G", 10)
summary.merge_range("A1:G1", "Regional Sales Summary 2026", title)
summary.write("A2", "Report date")
report_date = dt.date(2026, 10, 1)
summary.write_formula("B2", "=DATE(2026,10,1)", date_fmt, serial(report_date))
summary.write("D2", "Opened")
summary.write_formula("E2", "=TODAY()", date_fmt, serial(dt.date.today()))

summary.write_row("A4", ["Region", "Q1", "Q2", "Q3", "Q4", "Total", "Share"], header)
summary.write_comment("A4", "Totals are calculated from the Data sheet.")
quarters = [(dt.date(2026, m, 1), edate(dt.date(2026, m, 1), 3)) for m in (1, 4, 7, 10)]
totals = {}
for r, region in enumerate(REGIONS):
    row = 4 + r
    summary.write_string(row, 0, region)
    region_total = 0.0
    for q, (q_start, q_end) in enumerate(quarters):
        v = sum(x[5] for x in data if x[1] == region and q_start <= x[0] < q_end)
        region_total += v
        col = chr(ord("B") + q)
        summary.write_formula(
            row, 1 + q,
            f'=SUMIFS(Data!$F$2:$F${ROWS + 1},Data!$B$2:$B${ROWS + 1},$A{row + 1},'
            f'Data!$A$2:$A${ROWS + 1},">="&DATE({q_start.year},{q_start.month},1),'
            f'Data!$A$2:$A${ROWS + 1},"<"&DATE({q_end.year},{q_end.month},1))',
            money, v,
        )
        _ = col
    totals[region] = region_total
    summary.write_formula(row, 5, f"=SUM(B{row + 1}:E{row + 1})", money, region_total)
grand = sum(totals.values())
for r, region in enumerate(REGIONS):
    summary.write_formula(4 + r, 6, f"=F{5 + r}/$F$9", pct, totals[region] / grand)
summary.write("A9", "Total", bold)
for q in range(4):
    col = chr(ord("B") + q)
    summary.write_formula(8, 1 + q, f"=SUM({col}5:{col}8)", money_bold,
                          sum(x[5] for x in data if quarters[q][0] <= x[0] < quarters[q][1]))
summary.write_formula("F9", "=SUM(F5:F8)", money_bold, grand)
summary.write_formula("G9", "=SUM(G5:G8)", pct_bold, 1.0)

summary.write("A11", "Tax rate")
summary.write_formula("B11", "=TaxRate", pct, TAX_RATE)
summary.write("A12", "Tax due")
summary.write_formula("B12", "=F9*TaxRate", money, grand * TAX_RATE)
top = max(totals, key=totals.get)
summary.write("A13", "Top region")
summary.write_formula("B13", "=INDEX(A5:A8,MATCH(MAX(F5:F8),F5:F8,0))", None, top)
summary.write("A14", "Regions")
summary.write_formula("B14", '=TEXTJOIN(", ",TRUE,A5:A8)', None, ", ".join(REGIONS))
summary.write("A15", "Regions (CONCAT)")
summary.write_formula("B15", "=CONCAT(A5:A8)", None, "".join(REGIONS))
summary.write("A16", "Margin (LET)")
summary.write_formula("B16", "=LET(revenue,F9,cost,F9*0.6,revenue-cost)", money, grand - grand * 0.6)
summary.write("A17", "Margin after tax")
summary.write_formula("B17", "=B16*(1-TaxRate)", money, (grand - grand * 0.6) * (1 - TAX_RATE))
summary.write_url("A19", "https://example.com/sales-handbook", string="Sales handbook")

summary.write_row("A21", ["Month", "Revenue"], header)
monthly = []
for m in range(12):
    m_start = dt.date(2026, m + 1, 1)
    m_end = edate(m_start, 1)
    v = sum(x[5] for x in data if m_start <= x[0] < m_end)
    monthly.append(v)
    summary.write_datetime(21 + m, 0, dt.datetime.combine(m_start, dt.time()), month_fmt)
    summary.write_formula(
        21 + m, 1,
        f'=SUMIFS(Data!$F$2:$F${ROWS + 1},Data!$A$2:$A${ROWS + 1},">="&A{22 + m},'
        f'Data!$A$2:$A${ROWS + 1},"<"&EDATE(A{22 + m},1))',
        money, v,
    )
summary.freeze_panes(4, 1)

col_chart = wb.add_chart({"type": "column"})
col_chart.add_series({"name": "Total", "categories": "=Summary!$A$5:$A$8", "values": "=Summary!$F$5:$F$8", "fill": {"color": "#217346"}})
col_chart.set_title({"name": "Revenue by region"})
col_chart.set_legend({"none": True})
summary.insert_chart("I3", col_chart, {"x_scale": 1.0, "y_scale": 1.0})

pie = wb.add_chart({"type": "pie"})
pie.add_series({"name": "Share", "categories": "=Summary!$A$5:$A$8", "values": "=Summary!$G$5:$G$8"})
pie.set_title({"name": "Share of revenue"})
summary.insert_chart("I19", pie)

line = wb.add_chart({"type": "line"})
line.add_series({"name": "Monthly revenue", "categories": "=Summary!$A$22:$A$33", "values": "=Summary!$B$22:$B$33", "line": {"color": "#217346", "width": 2.25}})
line.set_title({"name": "Monthly revenue 2026"})
chartsheet.set_chart(line)

# --- Lookups -----------------------------------------------------------------
lookups.set_column("A:A", 26)
lookups.set_column("B:D", 16)
lookups.set_column("F:G", 12)
lookups.write("A1", "Lookup and formula examples", bold)
lookups.write_row("F3", ["Product", "Price"], header)
for i, (p, price) in enumerate(PRODUCTS.items()):
    lookups.write_string(3 + i, 5, p)
    lookups.write_number(3 + i, 6, price, money)
lookups.write_row("A3", ["Product", "Price (XLOOKUP)", "Price (VLOOKUP)", "Units sold (SUMIF)"], header)
for i, (p, price) in enumerate(PRODUCTS.items()):
    r = 3 + i
    lookups.write_string(r, 0, p)
    lookups.write_formula(r, 1, f"=XLOOKUP(A{r + 1},$F$4:$F$6,$G$4:$G$6)", money, price)
    lookups.write_formula(r, 2, f"=VLOOKUP(A{r + 1},$F$4:$G$6,2,FALSE)", money, price)
    lookups.write_formula(r, 3, f"=SUMIF(Data!$C$2:$C${ROWS + 1},A{r + 1},Data!$D$2:$D${ROWS + 1})", None,
                          sum(x[3] for x in data if x[2] == p))

lookups.write("A8", "Revenue (table reference)")
lookups.write_formula("B8", "=SUM(SalesData[Revenue])", money, grand)
lookups.write("A9", "Distinct regions (spill)")
lookups.write_dynamic_array_formula("B9:B12", f"=SORT(UNIQUE(Data!B2:B{ROWS + 1}))", None, sorted(REGIONS)[0])
for i, region in enumerate(sorted(REGIONS)[1:], start=1):
    lookups.write_string(8 + i, 1, region)
lookups.write("A14", "Weighted units (array)")
lookups.write_array_formula("B14:B14", f"{{=SUM(Data!D2:D{ROWS + 1}*Data!E2:E{ROWS + 1})}}", money, grand)
lookups.write("A15", "Doubled prices (array)")
doubled = [price * 2 for price in PRODUCTS.values()]
lookups.write_array_formula("B15:D15", "{=TRANSPOSE(G4:G6*2)}", money, doubled[0])
lookups.write_number("C15", doubled[1], money)
lookups.write_number("D15", doubled[2], money)
lookups.write("A17", "Rich text")
lookups.write_rich_string("B17", bold, "Bold", " and ", italic, "italic", " text")
lookups.write("A18", "Price band (IFS)")
lookups.write_formula("B18", '=IFS(B4>20,"High",B4>10,"Mid",TRUE,"Low")', None, "Mid")
lookups.write("A19", "Month end (EDATE)")
lookups.write_formula("B19", "=EDATE(DATE(2026,1,31),1)", date_fmt, serial(edate(dt.date(2026, 1, 31), 1)))
lookups.write("A20", "Label (text functions)")
lookups.write_formula("B20", '=UPPER(LEFT(A4,3))&"-"&TEXT(G4,"0.00")', None, "WID-12.50")
lookups.write("A21", "Negative check")
lookups.write_formula("B21", "=IF(SUM(D4:D6)<0,\"Negative\",\"OK\")", red, "OK")

# --- Notes ---------------------------------------------------------------------
notes.set_column("A:F", 14)
notes.write("A1", "Notes", bold)
notes.merge_range("A2:F4", "This sheet holds a picture, a text box, hidden rows and sparklines. "
                  "The editor keeps all of them when it saves, even though it only shows placeholders.", wrap)
notes.write_row("A6", ["Hidden", "row", "one"])
notes.write_row("A7", ["Hidden", "row", "two"])
notes.set_row(5, None, None, {"hidden": True})
notes.set_row(6, None, None, {"hidden": True})
notes.write_row("A9", ["Trend", "Jan", "Feb", "Mar", "Apr", "May"], header)
notes.write("A10", "Units")
for i, v in enumerate([12, 18, 9, 24, 30]):
    notes.write_number(9, 1 + i, v)
notes.add_sparkline("G10", {"range": "Notes!B10:F10", "markers": True})
notes.insert_image("H2", "logo.png", {"image_data": io.BytesIO(tiny_png())})
notes.insert_textbox("H8", "Text boxes are kept too.", {"width": 220, "height": 60})

# --- Reference (protected) -------------------------------------------------------
reference.write_row("A1", ["Code", "Meaning"], header)
for i, (code, meaning) in enumerate([("N", "North"), ("S", "South"), ("E", "East"), ("W", "West")], start=1):
    reference.write_row(i, 0, [code, meaning])
reference.protect()

wb.close()
print(f"Wrote {OUT}")
