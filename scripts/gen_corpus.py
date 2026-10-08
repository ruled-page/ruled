"""
Builds the workbooks in tests/corpus, which the round-trip checks save and
validate (npm run validate:saves). They cover what templates a host hands
out tend to use and the samples don't: a protected sheet with unlocked input
cells, an Excel table with a calculated column, a totals row and structured
references, dynamic array formulas, LET and LAMBDA, defined names, links,
hidden and grouped rows, the 1904 date system, and notes and threaded
comments.

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


PEOPLE = {
    "Ann Analyst": "{6B1F2D3A-0C4E-4F7B-9A21-3E5D7C8B9A10}",
    "Bob Reviewer": "{C2D4E6F8-1A3B-4C5D-8E9F-0A1B2C3D4E5F}",
}
# Each thread: its cell, whether it is resolved, and its comments (who, when, what), first one first.
THREADS = [
    ("C3", False, [
        ("Ann Analyst", "2026-09-14T09:12:31.52", "{2F4A6C8E-1B3D-4E5F-8A9B-0C1D2E3F4A5B}", "Does this include the Q3 adjustment?"),
        ("Bob Reviewer", "2026-09-14T11:40:02.17", "{7A8B9C0D-2E3F-4A5B-9C6D-1E2F3A4B5C6D}", "Yes, since Monday."),
    ]),
    ("A6", True, [
        ("Bob Reviewer", "2026-09-15T08:03:44.90", "{9E8D7C6B-5A4F-4E3D-8C2B-1A0F9E8D7C6B}", "Rename this row?"),
    ]),
]
LEGACY = (
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it "
    "will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n"
)


def comments():
    """
    Notes and threaded comments as Excel writes them. XlsxWriter writes notes
    only, so each thread's placeholder note (what versions of Excel before
    threads show) is written as a note, and the threads, the people in them
    and the links Excel adds are put into the package afterwards.
    """
    import re
    import zipfile
    from xml.sax.saxutils import escape

    path = OUT / "comments.xlsx"
    wb = xlsxwriter.Workbook(path)
    wb.set_properties({"created": dt.datetime(2026, 9, 14, 9, 0, 0)})
    bold = wb.add_format({"bold": True})

    review = wb.add_worksheet("Review")
    review.set_column("A:A", 18)
    review.write_row("A1", ["Line", "Budget", "Actual"], bold)
    for i, (name, budget, actual) in enumerate([("Salaries", 1200, 1180), ("Travel", 300, 412), ("Training", 150, 90), ("Software", 400, 400)]):
        review.write_row(i + 1, 0, [name, budget, actual])
    review.write_comment("B2", "Agreed with finance.\nRevisit in January.", {"author": "Ann Analyst"})
    for ref, _done, entries in THREADS:
        text = LEGACY + "Comment:\n    " + entries[0][3] + "".join(f"\nReply:\n    {e[3]}" for e in entries[1:])
        review.write_comment(ref, text, {"author": f"tc={entries[0][2]}"})

    plain = wb.add_worksheet("Plain")
    plain.write_row("A1", ["Item", "Qty"], bold)
    plain.write_row("A2", ["Bolts", 12])

    # Protected with Excel's defaults, which keep notes and comments as they are…
    locked = wb.add_worksheet("Locked")
    locked.protect()
    locked.write("A1", "Locked", bold)
    locked.write_comment("A1", "Ask before changing.", {"author": "Ann Analyst"})
    # …and with "Edit objects" allowed, which lets them be changed.
    edit = wb.add_worksheet("Objects")
    edit.protect("", {"objects": True})
    edit.write("A1", "Notes allowed", bold)
    wb.close()

    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    ns = 'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    persons = "".join(f'<person displayName="{n}" id="{i}" userId="{n}" providerId="None"/>' for n, i in PEOPLE.items())
    parts["xl/persons/person.xml"] = f"{head}<personList {ns}>{persons}</personList>".encode()
    threads = ""
    for ref, done, entries in THREADS:
        root = entries[0][2]
        for k, (who, when, id_, text) in enumerate(entries):
            parent = f' parentId="{root}"' if k else ""
            resolved = ' done="1"' if done and not k else ""
            threads += f'<threadedComment ref="{ref}" dT="{when}" personId="{PEOPLE[who]}" id="{id_}"{parent}{resolved}><text>{escape(text)}</text></threadedComment>'
    parts["xl/threadedComments/threadedComment1.xml"] = f"{head}<ThreadedComments {ns}>{threads}</ThreadedComments>".encode()

    # Excel marks each placeholder with its thread's id.
    xml = parts["xl/comments1.xml"].decode()
    xml = xml.replace(
        '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">',
    )
    authors = re.findall(r"<author>([^<]*)</author>", xml)
    for ref, _done, entries in THREADS:
        author = authors.index(f"tc={entries[0][2]}")
        xml = xml.replace(f'<comment ref="{ref}" authorId="{author}">', f'<comment ref="{ref}" authorId="{author}" shapeId="0" xr:uid="{entries[0][2]}">')
    parts["xl/comments1.xml"] = xml.encode()

    rels = parts["xl/worksheets/_rels/sheet1.xml.rels"].decode()
    rels = rels.replace(
        "</Relationships>",
        '<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>',
    )
    parts["xl/worksheets/_rels/sheet1.xml.rels"] = rels.encode()
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    rels = rels.replace(
        "</Relationships>",
        '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>',
    )
    parts["xl/_rels/workbook.xml.rels"] = rels.encode()
    types = parts["[Content_Types].xml"].decode()
    types = types.replace(
        "</Types>",
        '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>'
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>',
    )
    parts["[Content_Types].xml"] = types.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(zipfile.ZipInfo(name, (2026, 9, 14, 9, 0, 0)), data, zipfile.ZIP_DEFLATED)


OUT.mkdir(parents=True, exist_ok=True)
templates()
dates_1904()
comments()
print(f"Wrote {OUT}/templates.xlsx, dates-1904.xlsx and comments.xlsx")
