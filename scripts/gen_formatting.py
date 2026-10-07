"""
Builds public/samples/formatting.xlsx: a workbook for checking how the editor
draws Excel formatting — every border style, text spilling into empty cells,
notes, row and column formats, each kind of conditional formatting, table
styles and frozen rows.

    python3 -m venv .venv && .venv/bin/pip install xlsxwriter
    .venv/bin/python scripts/gen_formatting.py
"""

from pathlib import Path

import xlsxwriter

OUT = Path(__file__).resolve().parent.parent / "public" / "samples" / "formatting.xlsx"

wb = xlsxwriter.Workbook(str(OUT))
wb.set_properties({"title": "Formatting sample", "author": "Workbook editor"})
bold = wb.add_format({"bold": True})
heading = wb.add_format({"bold": True, "font_size": 13})

# --- Borders & text -----------------------------------------------------------
ws = wb.add_worksheet("Borders & text")
ws.set_column("A:A", 18)
ws.set_column("B:B", 12)
ws.set_column("C:G", 11)
ws.write("A1", "Border styles", heading)
styles = [
    (1, "thin"), (2, "medium"), (5, "thick"), (6, "double"), (3, "dashed"), (4, "dotted"),
    (7, "hair"), (8, "medium dashed"), (9, "dash dot"), (13, "slant dash dot"),
]
for i, (index, name) in enumerate(styles):
    row = 2 + i * 2
    ws.write(row, 0, name)
    ws.write_blank(row, 1, None, wb.add_format({"border": index}))
ws.write("D3", "Coloured", bold)
ws.write_blank("D5", None, wb.add_format({"border": 2, "border_color": "#C00000"}))
ws.write_blank("E5", None, wb.add_format({"border": 2, "border_color": "#2F75B5"}))
ws.write("D7", "Outline", bold)
box = {"top": 2, "bottom": 2, "left": 2, "right": 2}
for r in range(8, 11):
    for c in range(3, 6):
        f = {"border": 1}
        if r == 8:
            f["top"] = 2
        if r == 10:
            f["bottom"] = 6
        if c == 3:
            f["left"] = 2
        if c == 5:
            f["right"] = 2
        ws.write_number(r, c, (r - 7) * (c - 2), wb.add_format(f))
ws.write("D13", "Bottom only on the row above", bold)
ws.write_blank("D14", None, wb.add_format({"bottom": 5}))
ws.write_blank("E14", None, wb.add_format({"bottom": 5}))

ws.write("A24", "Overflow", heading)
ws.write("A25", "Left-aligned text runs on into the empty cells to its right.")
ws.write("F26", "Right-aligned text runs left.", wb.add_format({"align": "right"}))
ws.write("C27", "Centred text spreads both ways", wb.add_format({"align": "center"}))
ws.write("A28", "Stopped by the next cell")
ws.write("B28", "here")
ws.write("A29", "Wrapped text stays inside its own cell", wb.add_format({"text_wrap": True}))
ws.set_row(28, 45)
ws.write_number("A30", 123456789012345, wb.add_format({"num_format": "0"}))
ws.write("A31", "Through a styled empty cell")
ws.write_blank("B31", None, wb.add_format({"bg_color": "#FFF2CC"}))

ws.write("A33", "Notes", heading)
ws.write("A34", "Hover me")
ws.write_comment("A34", "This is a note on A34.\nIt has two lines.", {"author": "Alex Analyst"})
ws.write_number("B34", 42)
ws.write_comment("B34", "Checked against the ledger.", {"author": "Sam Reviewer"})

ws.write("A36", "Row and column formats", heading)
ws.set_column("H:H", 12, wb.add_format({"bg_color": "#E2EFDA"}))
ws.write("H1", "Column H has a fill")
ws.set_row(37, None, wb.add_format({"bg_color": "#DDEBF7"}))
ws.write("A38", "Row 38 has a fill")

# --- Conditional formatting -----------------------------------------------------
cf = wb.add_worksheet("Conditional")
cf.freeze_panes(1, 1)
cf.set_column("A:A", 12)
cf.set_column("B:J", 11)
headers = ["Item", "Score", "Status", "Colour scale", "Data bar", "Arrows", "Text", "Duplicates", "Top 3", "Above avg"]
cf.write_row(0, 0, headers, bold)
statuses = ["On time", "Late", "On time", "On time", "Late"]
texts = ["ok", "error: missing", "", "ok", "fine", "error", "", "ok"]
for i in range(30):
    row = i + 1
    cf.write(row, 0, f"Item {i + 1}")
    cf.write_number(row, 1, (i * 37) % 100)
    cf.write(row, 2, statuses[i % len(statuses)])
    cf.write_number(row, 3, (i * 13) % 50)
    cf.write_number(row, 4, (i * 29) % 100)
    cf.write_number(row, 5, (i * 17) % 30)
    text = texts[i % len(texts)]
    if text:
        cf.write(row, 6, text)
    cf.write_number(row, 7, (i * 7) % 12)
    cf.write_number(row, 8, (i * 23) % 61)
    cf.write_number(row, 9, (i * 11) % 40)
last = 30
green = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "bold": True})
red_text = wb.add_format({"font_color": "#9C0006", "italic": True})
late = wb.add_format({"bg_color": "#FCE4D6"})
yellow = wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C5700"})
blue = wb.add_format({"bold": True, "font_color": "#1F4E78", "bg_color": "#DDEBF7"})
grey = wb.add_format({"bg_color": "#EDEDED"})
strike = wb.add_format({"font_strikeout": True})
cf.conditional_format(1, 1, last, 1, {"type": "cell", "criteria": ">=", "value": 80, "format": green})
cf.conditional_format(1, 1, last, 1, {"type": "cell", "criteria": "<", "value": 30, "format": red_text})
cf.conditional_format(1, 0, last, 9, {"type": "formula", "criteria": '=$C2="Late"', "format": late})
cf.conditional_format(1, 3, last, 3, {"type": "3_color_scale"})
cf.conditional_format(1, 4, last, 4, {"type": "data_bar", "bar_color": "#638EC6"})
cf.conditional_format(1, 5, last, 5, {"type": "icon_set", "icon_style": "3_arrows"})
cf.conditional_format(1, 6, last, 6, {"type": "text", "criteria": "containing", "value": "error", "format": yellow})
cf.conditional_format(1, 6, last, 6, {"type": "blanks", "format": grey})
cf.conditional_format(1, 7, last, 7, {"type": "duplicate", "format": strike})
cf.conditional_format(1, 8, last, 8, {"type": "top", "value": 3, "format": blue})
cf.conditional_format(1, 9, last, 9, {"type": "average", "criteria": "above", "format": green})

# --- Tables -------------------------------------------------------------------
tables = wb.add_worksheet("Tables")
tables.set_column("A:J", 11)
rows = [["North", 120, 95, 30], ["South", 80, 110, 22], ["East", 140, 70, 41], ["West", 60, 130, 18]]
specs = [
    (0, 0, "Table Style Medium 2", {"total_row": True}),
    (0, 5, "Table Style Light 9", {}),
    (8, 0, "Table Style Dark 2", {"banded_columns": True, "first_column": True}),
    (8, 5, "Table Style Medium 14", {}),
    (16, 0, "Table Style Light 1", {}),
    (16, 5, "Table Style Light 15", {"banded_rows": False, "banded_columns": True}),
]
for n, (r, c, style, extra) in enumerate(specs):
    columns = [{"header": "Region"}, {"header": "Plan"}, {"header": "Actual"}, {"header": "Margin"}]
    if extra.get("total_row"):
        columns[0]["total_string"] = "Total"
        for col in columns[1:]:
            col["total_function"] = "sum"
    tables.write(r + 6 if extra.get("total_row") else r + 5, c, style)
    tables.add_table(
        r, c, r + len(rows) + (1 if extra.get("total_row") else 0), c + 3,
        {"name": f"Styled{n + 1}", "style": style, "data": rows, "columns": columns, **extra},
    )

wb.close()
print(f"Wrote {OUT}")
