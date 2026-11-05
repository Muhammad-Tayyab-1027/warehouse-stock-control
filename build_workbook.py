"""Builds warehouse_stock_control.xlsx. All data is made up (fixed random seed)."""
import random
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

random.seed(1027)
FONT = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
INPUT = PatternFill("solid", fgColor="FFF2CC")
RED = PatternFill("solid", bgColor="F8CBAD")
AMBER = PatternFill("solid", bgColor="FFE699")
GREEN = PatternFill("solid", bgColor="C6E0B4")

ITEMS = {
    "Plumbing": ["Copper Pipe 15mm 3m", "Copper Pipe 22mm 3m", "Compression Elbow 15mm", "PTFE Tape 12m",
                 "Isolation Valve 15mm", "Radiator Valve Set", "Flexi Hose 300mm"],
    "Bathroom": ["Basin Mixer Tap", "Shower Head Chrome", "Toilet Seat Soft Close", "Bath Waste Kit",
                 "Towel Rail 500mm", "Silicone Sealant White"],
    "Electrical": ["LED Bulb E27 9W", "Extension Lead 4 Gang", "Cable Clips 100pk", "Junction Box IP65",
                   "Twin Socket White", "Light Switch 1 Gang"],
    "Packaging": ["Carton Small", "Carton Medium", "Carton Large", "Parcel Tape 66m", "Bubble Wrap 50m",
                  "Pallet Wrap 400m", "Shipping Labels 500pk"],
    "Safety": ["Hi-Vis Vest L", "Work Gloves Pair", "Safety Boots 43", "Dust Mask 10pk"],
}
START = date(2026, 7, 1)
END = date(2026, 9, 30)
DAYS = (END - START).days + 1


def header(ws, row, names, widths):
    for i, name in enumerate(names, 1):
        c = ws.cell(row=row, column=i, value=name)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[col(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 32


wb = Workbook()

# ---------- Guide ----------
guide = wb.active
guide.title = "Guide"
guide.column_dimensions["A"].width = 100
lines = [
    "Warehouse Stock Control Workbook", "",
    "Portfolio project. All products, quantities and prices are made up.",
    "", "How it works:",
    "Dashboard: headline figures, assumptions you can edit, and stock value by category.",
    "Products: the product list. Yellow cells are inputs.",
    "Movements: every receipt (IN), pick (OUT) and adjustment (ADJ). Adjustments may be + or -.",
    "Stock Status: formulas work out stock on hand, reorder point, status and the quantity to order.",
    "ABC Analysis: ranks products by value moved. Class A = first 80%, B = next 15%, C = last 5%.",
    "Cycle Count: type a physical count in the yellow column to see the variance and count accuracy.",
    "", "Reorder point = average daily use x supplier lead time (days) + safety stock.",
]
for i, t in enumerate(lines, 1):
    guide.cell(row=i, column=1, value=t)
guide["A1"].font = Font(name=FONT, bold=True, size=14)
guide["A5"].font = Font(name=FONT, bold=True)

# ---------- Products ----------
pr = wb.create_sheet("Products")
header(pr, 1, ["SKU", "Product", "Category", "Location", "Unit Cost (EUR)", "Lead Time (days)",
               "Safety Stock", "Opening Stock"], [10, 28, 13, 10, 11, 10, 9, 10])
products = []
n = 1001
for cat, names in ITEMS.items():
    for name in names:
        products.append(dict(sku=f"SKU{n}", name=name, cat=cat,
                             loc=f"{random.choice('ABCD')}{random.randint(1, 12):02d}-{random.randint(1, 4)}",
                             cost=round(random.uniform(0.8, 48), 2), lead=random.choice([2, 3, 5, 7, 10, 14]),
                             safety=random.choice([10, 15, 20, 30, 40, 60]), rate=random.uniform(1, 14)))
        n += 1
for r, p in enumerate(products, 2):
    p["open"] = int(p["rate"] * random.uniform(4, 24)) + p["safety"]
    for c, v in enumerate([p["sku"], p["name"], p["cat"], p["loc"], p["cost"], p["lead"], p["safety"], p["open"]], 1):
        cell = pr.cell(row=r, column=c, value=v)
        if c >= 5:
            cell.fill = INPUT
    pr.cell(row=r, column=5).number_format = "#,##0.00"
pr.freeze_panes = "A2"
PN = len(products) + 1

# ---------- Movements ----------
mv = wb.create_sheet("Movements")
header(mv, 1, ["Date", "SKU", "Type", "Quantity", "Reference"], [12, 10, 8, 10, 14])
rows = []
for p in products:
    # Some products are deliberately under-replenished so that every status appears.
    level = random.choice([0.7, 0.85, 0.95, 1.0, 1.05])
    d = START
    while d <= END:
        if d.weekday() < 5:
            q = max(1, int(random.gauss(p["rate"] * 1.4, p["rate"] * 0.4)))
            rows.append((d, p["sku"], "OUT", q))
        d += timedelta(days=1)
    received = 0
    target = p["rate"] * DAYS * level
    while received < target:
        q = max(10, int(p["rate"] * random.uniform(12, 25)))
        rows.append((START + timedelta(days=random.randrange(DAYS)), p["sku"], "IN", q))
        received += q
    if random.random() < 0.25:
        rows.append((START + timedelta(days=random.randrange(DAYS)), p["sku"], "ADJ", random.choice([-3, -2, -1, 1, 2])))
rows.sort(key=lambda x: (x[0], x[1]))
counters = {"IN": 0, "OUT": 0, "ADJ": 0}
prefix = {"IN": "GRN", "OUT": "ORD", "ADJ": "ADJ"}
for r, (d, sku, typ, q) in enumerate(rows, 2):
    counters[typ] += 1
    mv.cell(row=r, column=1, value=d).number_format = "yyyy-mm-dd"
    mv.cell(row=r, column=2, value=sku)
    mv.cell(row=r, column=3, value=typ)
    mv.cell(row=r, column=4, value=q)
    mv.cell(row=r, column=5, value=f"{prefix[typ]}-{counters[typ]:05d}")
MN = len(rows) + 1
LIMIT = MN + 2000  # room for new rows
dv = DataValidation(type="list", formula1='"IN,OUT,ADJ"', allow_blank=False)
mv.add_data_validation(dv)
dv.add(f"C2:C{LIMIT}")
mv.freeze_panes = "A2"
mv.auto_filter.ref = f"A1:E{MN}"
M_SKU, M_TYPE, M_QTY = (f"Movements!${c}$2:${c}${LIMIT}" for c in "BCD")

# ---------- Dashboard is built last, but other sheets refer to its assumption cells ----------
db = wb.create_sheet("Dashboard", 1)
DAYS_CELL, A_CELL, B_CELL, COVER_CELL = "Dashboard!$B$11", "Dashboard!$B$12", "Dashboard!$B$13", "Dashboard!$B$14"

# ---------- Stock Status ----------
st = wb.create_sheet("Stock Status")
header(st, 1, ["SKU", "Product", "Category", "Opening", "Received", "Picked", "Adjusted", "On Hand", "Avg Daily Use",
               "Reorder Point", "Status", "Suggested Order", "Stock Value (EUR)"],
       [10, 28, 13, 9, 9, 9, 9, 9, 8, 9, 14, 10, 12])
for r in range(2, PN + 1):
    st[f"A{r}"] = f"=Products!A{r}"
    st[f"B{r}"] = f"=Products!B{r}"
    st[f"C{r}"] = f"=Products!C{r}"
    st[f"D{r}"] = f"=Products!H{r}"
    st[f"E{r}"] = f'=SUMIFS({M_QTY},{M_SKU},$A{r},{M_TYPE},"IN")'
    st[f"F{r}"] = f'=SUMIFS({M_QTY},{M_SKU},$A{r},{M_TYPE},"OUT")'
    st[f"G{r}"] = f'=SUMIFS({M_QTY},{M_SKU},$A{r},{M_TYPE},"ADJ")'
    st[f"H{r}"] = f"=D{r}+E{r}-F{r}+G{r}"
    st[f"I{r}"] = f"=F{r}/{DAYS_CELL}"
    st[f"I{r}"].number_format = "0.0"
    st[f"J{r}"] = f"=ROUNDUP(I{r}*Products!F{r}+Products!G{r},0)"
    st[f"K{r}"] = f'=IF(H{r}<=0,"OUT OF STOCK",IF(H{r}<=J{r},"REORDER",IF(H{r}<=J{r}*1.25,"WATCH","OK")))'
    st[f"L{r}"] = f'=IF(OR(K{r}="REORDER",K{r}="OUT OF STOCK"),ROUNDUP(J{r}-H{r}+I{r}*{COVER_CELL},0),0)'
    st[f"M{r}"] = f"=MAX(H{r},0)*Products!E{r}"
    st[f"M{r}"].number_format = "#,##0.00"
st.freeze_panes = "C2"
st.auto_filter.ref = f"A1:M{PN}"
rng = f"K2:K{PN}"
st.conditional_formatting.add(rng, FormulaRule(formula=['OR($K2="REORDER",$K2="OUT OF STOCK")'], fill=RED))
st.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"WATCH"'], fill=AMBER))
st.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], fill=GREEN))

# ---------- ABC ----------
ab = wb.create_sheet("ABC Analysis")
header(ab, 1, ["SKU", "Product", "Usage Value (EUR)", "Rank", "Share", "Cumulative Share", "Class"],
       [10, 28, 13, 7, 9, 12, 7])
for r in range(2, PN + 1):
    ab[f"A{r}"] = f"=Products!A{r}"
    ab[f"B{r}"] = f"=Products!B{r}"
    ab[f"C{r}"] = f"='Stock Status'!F{r}*Products!E{r}"
    ab[f"D{r}"] = f"=RANK(C{r},$C$2:$C${PN})"
    ab[f"E{r}"] = f"=C{r}/SUM($C$2:$C${PN})"
    ab[f"F{r}"] = f'=SUMIF($C$2:$C${PN},">="&C{r})/SUM($C$2:$C${PN})'
    # A product belongs to a class according to the cumulative share reached BEFORE it.
    ab[f"G{r}"] = f'=IF(F{r}-E{r}<{A_CELL},"A",IF(F{r}-E{r}<{B_CELL},"B","C"))'
    ab[f"C{r}"].number_format = "#,##0.00"
    ab[f"E{r}"].number_format = ab[f"F{r}"].number_format = "0.0%"
ab.freeze_panes = "C2"
ab.auto_filter.ref = f"A1:G{PN}"

# ---------- Cycle Count ----------
cc = wb.create_sheet("Cycle Count")
header(cc, 1, ["SKU", "Product", "Location", "System Qty", "Counted Qty", "Variance", "Variance Value (EUR)", "Match?"],
       [16, 28, 10, 9, 10, 9, 11, 8])
sample = random.sample(range(2, PN + 1), 12)
for i, src in enumerate(sorted(sample), 2):
    cc[f"A{i}"] = f"=Products!A{src}"
    cc[f"B{i}"] = f"=Products!B{src}"
    cc[f"C{i}"] = f"=Products!D{src}"
    cc[f"D{i}"] = f"='Stock Status'!H{src}"
    cc[f"E{i}"].fill = INPUT
    cc[f"F{i}"] = f'=IF(E{i}="","",E{i}-D{i})'
    cc[f"G{i}"] = f'=IF(E{i}="","",F{i}*Products!E{src})'
    cc[f"G{i}"].number_format = "#,##0.00"
    cc[f"H{i}"] = f'=IF(E{i}="","",IF(F{i}=0,"Yes","No"))'
CN = len(sample) + 1
cc[f"A{CN + 2}"] = "Items counted"
cc[f"D{CN + 2}"] = f"=COUNT(E2:E{CN})"
cc[f"A{CN + 3}"] = "Count accuracy"
cc[f"D{CN + 3}"] = f'=IF(D{CN + 2}=0,"Enter counts",COUNTIF(H2:H{CN},"Yes")/D{CN + 2})'
cc[f"D{CN + 3}"].number_format = "0.0%"
cc[f"A{CN + 2}"].font = cc[f"A{CN + 3}"].font = Font(name=FONT, bold=True)
cc[f"A{CN + 5}"] = "Type a physical count in the yellow column. The counts are left blank on purpose."

# ---------- Dashboard ----------
db.column_dimensions["A"].width = 36
db.column_dimensions["B"].width = 16
db["A1"] = "Stock Dashboard (sample data)"
db["A1"].font = Font(name=FONT, bold=True, size=14)
kpis = [
    ("Products", f"=COUNTA(Products!A2:A{PN})", "0"),
    ("Total stock value (EUR)", f"=SUM('Stock Status'!M2:M{PN})", "#,##0"),
    ("Units on hand", f"=SUM('Stock Status'!H2:H{PN})", "#,##0"),
    ("Products to reorder", f"=COUNTIF('Stock Status'!K2:K{PN},\"REORDER\")", "0"),
    ("Products out of stock", f"=COUNTIF('Stock Status'!K2:K{PN},\"OUT OF STOCK\")", "0"),
    ("Class A products", f"=COUNTIF('ABC Analysis'!G2:G{PN},\"A\")", "0"),
]
for i, (label, f, fmt) in enumerate(kpis, 3):
    db[f"A{i}"] = label
    db[f"B{i}"] = f
    db[f"B{i}"].number_format = fmt
    db[f"B{i}"].font = Font(name=FONT, bold=True)
db["A10"] = "Assumptions (edit these)"
db["A10"].font = Font(name=FONT, bold=True)
for r, (label, v, fmt) in enumerate([("Days in period (1 Jul to 30 Sep 2026)", DAYS, "0"), ("Class A cut-off", 0.8, "0%"),
                                     ("Class B cut-off", 0.95, "0%"), ("Days of cover to order", 30, "0")], 11):
    db[f"A{r}"] = label
    db[f"B{r}"] = v
    db[f"B{r}"].fill = INPUT
    db[f"B{r}"].font = Font(name=FONT, color="0000FF")
    db[f"B{r}"].number_format = fmt
db["A16"] = "Category"
db["B16"] = "Stock Value (EUR)"
for c in ("A16", "B16"):
    db[c].font = Font(name=FONT, bold=True, color="FFFFFF")
    db[c].fill = HEAD
for r, cat in enumerate(ITEMS, 17):
    db[f"A{r}"] = cat
    db[f"B{r}"] = f"=SUMIF('Stock Status'!$C$2:$C${PN},A{r},'Stock Status'!$M$2:$M${PN})"
    db[f"B{r}"].number_format = "#,##0"
chart = BarChart()
chart.type = "bar"
chart.title = "Stock value by category (EUR)"
chart.legend = None
chart.add_data(Reference(db, min_col=2, min_row=16, max_row=16 + len(ITEMS)), titles_from_data=True)
chart.set_categories(Reference(db, min_col=1, min_row=17, max_row=16 + len(ITEMS)))
chart.series[0].graphicalProperties.solidFill = "1F3864"
chart.height, chart.width = 8, 16
db.add_chart(chart, "D3")

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != FONT:
                c.font = Font(name=FONT, bold=c.font.bold, color=c.font.color)
wb.save("warehouse_stock_control.xlsx")
print("movements", len(rows), "products", len(products))
