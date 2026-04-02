"""
FMCG Supplier Dashboard Builder
================================
Step 1 : Simulate a realistic supply-chain dataset (mirrors Kaggle "Supply Chain Analysis" datasets).
Step 2 : Write Raw_Data, KPI_Scorecard, and Dashboard sheets to an Excel workbook.
Step 3 : Apply professional formatting, Excel formulas, and embedded charts.
"""

import random, math
from datetime import date, timedelta

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage

# ─────────────────────────────────────────────────────────────
# SECTION 1 ▸ SIMULATE DATA
# Mirrors the structure of Kaggle "Supply Chain Analysis" dataset
# ─────────────────────────────────────────────────────────────

random.seed(42)
np.random.seed(42)

SUPPLIERS = {
    "NatureFresh Ltd":    {"base_cost": 4.20,  "otd_rate": 0.94, "quality": 88},
    "GlobalGrain Co":     {"base_cost": 3.80,  "otd_rate": 0.78, "quality": 72},
    "PrimePack GmbH":     {"base_cost": 6.50,  "otd_rate": 0.91, "quality": 91},
    "SunSource Agri":     {"base_cost": 2.90,  "otd_rate": 0.85, "quality": 81},
    "EcoPalm Traders":    {"base_cost": 5.10,  "otd_rate": 0.70, "quality": 68},
    "FastFoil Industries":{"base_cost": 8.30,  "otd_rate": 0.96, "quality": 94},
    "HarvestHub Inc":     {"base_cost": 3.60,  "otd_rate": 0.88, "quality": 85},
    "BioBase Supplies":   {"base_cost": 7.20,  "otd_rate": 0.82, "quality": 79},
}

CATEGORIES = {
    "Oils & Fats":        ["Palm Oil", "Sunflower Oil", "Coconut Oil"],
    "Grains & Cereals":   ["Wheat Flour", "Corn Starch", "Oat Bran"],
    "Packaging":          ["PET Bottles", "Aluminium Foil", "Cardboard Boxes"],
    "Sweeteners":         ["Refined Sugar", "HFCS", "Honey"],
    "Proteins":           ["Soy Isolate", "Whey Powder", "Pea Protein"],
}

SUPPLIER_CATEGORY = {
    "NatureFresh Ltd":     "Oils & Fats",
    "GlobalGrain Co":      "Grains & Cereals",
    "PrimePack GmbH":      "Packaging",
    "SunSource Agri":      "Sweeteners",
    "EcoPalm Traders":     "Oils & Fats",
    "FastFoil Industries": "Packaging",
    "HarvestHub Inc":      "Grains & Cereals",
    "BioBase Supplies":    "Proteins",
}

PAYMENT_TERMS = ["Net 30", "Net 45", "Net 60", "2/10 Net 30"]
WAREHOUSES    = ["London", "Manchester", "Birmingham", "Bristol", "Leeds"]

start_date = date(2023, 1, 1)
end_date   = date(2024, 12, 31)

rows = []
order_id = 1001

d = start_date
while d <= end_date:
    # 3–7 orders per day across suppliers
    for _ in range(random.randint(3, 7)):
        supplier = random.choice(list(SUPPLIERS.keys()))
        cfg      = SUPPLIERS[supplier]
        category = SUPPLIER_CATEGORY[supplier]
        product  = random.choice(CATEGORIES[category])

        # Cost drifts ± seasonally + random noise
        month_factor = 1 + 0.04 * math.sin((d.month - 1) * math.pi / 6)
        noise        = random.uniform(0.90, 1.12)
        unit_cost    = round(cfg["base_cost"] * month_factor * noise, 2)

        qty   = random.randint(200, 5000)
        total = round(unit_cost * qty, 2)

        # Lead time: base 7 days, varies by supplier quality
        lead_base = 7 + round((100 - cfg["otd_rate"] * 100) / 5)
        lead_time = max(1, int(np.random.normal(lead_base, 2.5)))

        # On-Time Delivery
        otd = "Yes" if random.random() < cfg["otd_rate"] else "No"

        # Quality score with noise
        q_score = min(100, max(40, int(np.random.normal(cfg["quality"], 4))))

        # Defect rate inversely related to quality
        defect_rate = round(max(0, (100 - q_score) / 400 + random.uniform(0, 0.02)), 4)

        rows.append({
            "Order_ID":         f"PO-{order_id}",
            "Date":             d,
            "Supplier":         supplier,
            "Category":         category,
            "Product":          product,
            "Warehouse":        random.choice(WAREHOUSES),
            "Payment_Terms":    random.choice(PAYMENT_TERMS),
            "Unit_Cost_GBP":    unit_cost,
            "Order_Qty_kg":     qty,
            "Total_Cost_GBP":   total,
            "Lead_Time_Days":   lead_time,
            "On_Time_Delivery": otd,
            "Quality_Score":    q_score,
            "Defect_Rate_Pct":  round(defect_rate * 100, 2),
        })
        order_id += 1

    d += timedelta(days=1)

df = pd.DataFrame(rows)
print(f"✅ Generated {len(df):,} order records across {df['Supplier'].nunique()} suppliers")

# ─────────────────────────────────────────────────────────────
# SECTION 2 ▸ AGGREGATE KPIs PER SUPPLIER
# ─────────────────────────────────────────────────────────────

kpi = (
    df.groupby("Supplier")
    .agg(
        Category        =("Category",         "first"),
        Total_Orders    =("Order_ID",         "count"),
        Total_Spend_GBP =("Total_Cost_GBP",   "sum"),
        Avg_Unit_Cost   =("Unit_Cost_GBP",    "mean"),
        Avg_Lead_Time   =("Lead_Time_Days",   "mean"),
        OTD_Rate_Pct    =("On_Time_Delivery", lambda x: round((x == "Yes").mean() * 100, 1)),
        Avg_Quality     =("Quality_Score",    "mean"),
        Avg_Defect_Rate =("Defect_Rate_Pct",  "mean"),
    )
    .reset_index()
)
kpi["Total_Spend_GBP"] = kpi["Total_Spend_GBP"].round(0)
kpi["Avg_Unit_Cost"]   = kpi["Avg_Unit_Cost"].round(2)
kpi["Avg_Lead_Time"]   = kpi["Avg_Lead_Time"].round(1)
kpi["Avg_Quality"]     = kpi["Avg_Quality"].round(1)
kpi["Avg_Defect_Rate"] = kpi["Avg_Defect_Rate"].round(2)

# Supplier Performance Score (0-100)  
# Weighted: OTD 40% + Quality 40% + (inverse lead time) 20%
max_lead = kpi["Avg_Lead_Time"].max()
kpi["Perf_Score"] = (
    0.40 * kpi["OTD_Rate_Pct"] +
    0.40 * kpi["Avg_Quality"] +
    0.20 * ((max_lead - kpi["Avg_Lead_Time"]) / max_lead * 100)
).round(1)

kpi = kpi.sort_values("Perf_Score", ascending=False).reset_index(drop=True)

# Monthly trend for line chart
monthly = (
    df.assign(Month=df["Date"].apply(lambda d: d.strftime("%b %Y")),
              MonthNum=df["Date"].apply(lambda d: d.year * 100 + d.month))
    .groupby(["MonthNum", "Month"])
    .agg(
        Total_Spend       =("Total_Cost_GBP", "sum"),
        Avg_Unit_Cost     =("Unit_Cost_GBP",  "mean"),
        OTD_Rate          =("On_Time_Delivery", lambda x: round((x == "Yes").mean() * 100, 1)),
    )
    .reset_index()
    .sort_values("MonthNum")
    .drop("MonthNum", axis=1)
)
monthly["Total_Spend"]   = monthly["Total_Spend"].round(0)
monthly["Avg_Unit_Cost"] = monthly["Avg_Unit_Cost"].round(2)

print("✅ KPI aggregation complete")

# ─────────────────────────────────────────────────────────────
# SECTION 3 ▸ BUILD EXCEL WORKBOOK
# ─────────────────────────────────────────────────────────────

wb = Workbook()

# ── Colour palette ──────────────────────────────────────────
C_DARK_BLUE  = "1F3864"   # header backgrounds
C_MID_BLUE   = "2E75B6"   # sub-headers
C_LIGHT_BLUE = "D6E4F0"   # alternating rows
C_WHITE      = "FFFFFF"
C_ACCENT     = "ED7D31"   # KPI highlight
C_GREEN      = "70AD47"
C_RED        = "FF0000"
C_YELLOW     = "FFD966"
C_TITLE_FONT = "FFFFFF"
C_DARK_TEXT  = "1A1A1A"

def col_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def bold_font(size=11, color=C_DARK_TEXT, italic=False):
    return Font(name="Arial", bold=True, size=size, color=color, italic=italic)

def reg_font(size=10, color=C_DARK_TEXT):
    return Font(name="Arial", size=size, color=color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center")

def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def set_row_height(ws, row, height):
    ws.row_dimensions[row].height = height

# ════════════════════════════════════════════════════════════
# SHEET A ▸ Raw_Data
# ════════════════════════════════════════════════════════════
ws_raw = wb.active
ws_raw.title = "Raw_Data"
ws_raw.sheet_view.showGridLines = True
ws_raw.freeze_panes = "A2"

RAW_COLS = list(df.columns)
COL_WIDTHS_RAW = [12, 13, 20, 18, 20, 14, 16, 15, 14, 16, 15, 17, 14, 16]

# Header row
for c, (col, w) in enumerate(zip(RAW_COLS, COL_WIDTHS_RAW), start=1):
    cell = ws_raw.cell(row=1, column=c, value=col.replace("_", " "))
    cell.font      = bold_font(10, C_TITLE_FONT)
    cell.fill      = col_fill(C_DARK_BLUE)
    cell.alignment = center()
    cell.border    = thin_border()
    ws_raw.column_dimensions[get_column_letter(c)].width = w

set_row_height(ws_raw, 1, 32)

# Data rows
for r_idx, row in enumerate(df.itertuples(index=False), start=2):
    fill = col_fill(C_LIGHT_BLUE) if r_idx % 2 == 0 else col_fill(C_WHITE)
    for c_idx, val in enumerate(row, start=1):
        cell = ws_raw.cell(row=r_idx, column=c_idx, value=val)
        cell.font      = reg_font(9)
        cell.fill      = fill
        cell.alignment = center()
        cell.border    = thin_border()
        if c_idx in (8, 10):   # cost columns
            cell.number_format = "£#,##0.00"
        if c_idx == 9:         # qty
            cell.number_format = "#,##0"
        if c_idx == 14:        # defect %
            cell.number_format = "0.00%"

print("✅ Raw_Data sheet written")

# ════════════════════════════════════════════════════════════
# SHEET B ▸ KPI_Scorecard
# ════════════════════════════════════════════════════════════
ws_kpi = wb.create_sheet("KPI_Scorecard")
ws_kpi.sheet_view.showGridLines = False
ws_kpi.freeze_panes = "A3"

# ── Title banner ─────────────────────────────────────────
ws_kpi.merge_cells("A1:K1")
t = ws_kpi["A1"]
t.value      = "FMCG SUPPLIER PERFORMANCE SCORECARD  |  FY 2023–2024"
t.font       = bold_font(15, C_TITLE_FONT)
t.fill       = col_fill(C_DARK_BLUE)
t.alignment  = center()
set_row_height(ws_kpi, 1, 40)

# Column headers
KPI_HEADERS = [
    "Rank", "Supplier", "Category", "Total Orders",
    "Total Spend (£)", "Avg Unit Cost (£)", "Avg Lead Time (Days)",
    "OTD Rate (%)", "Avg Quality Score", "Avg Defect Rate (%)",
    "Performance Score"
]
KPI_WIDTHS = [7, 22, 18, 14, 18, 18, 20, 13, 18, 20, 18]

for c, (h, w) in enumerate(zip(KPI_HEADERS, KPI_WIDTHS), start=1):
    cell = ws_kpi.cell(row=2, column=c, value=h)
    cell.font      = bold_font(10, C_TITLE_FONT)
    cell.fill      = col_fill(C_MID_BLUE)
    cell.alignment = center()
    cell.border    = thin_border()
    ws_kpi.column_dimensions[get_column_letter(c)].width = w

set_row_height(ws_kpi, 2, 38)

# KPI data rows
for r_idx, row in enumerate(kpi.itertuples(index=False), start=3):
    rank   = r_idx - 2
    fill   = col_fill(C_LIGHT_BLUE) if rank % 2 == 0 else col_fill(C_WHITE)
    score  = row.Perf_Score

    vals = [
        rank, row.Supplier, row.Category, row.Total_Orders,
        row.Total_Spend_GBP, row.Avg_Unit_Cost, row.Avg_Lead_Time,
        row.OTD_Rate_Pct, row.Avg_Quality, row.Avg_Defect_Rate,
        score
    ]
    fmts = [
        "0", "@", "@", "#,##0",
        "£#,##0", "£#,##0.00", "0.0",
        "0.0", "0.0", "0.00",
        "0.0"
    ]

    for c_idx, (val, fmt) in enumerate(zip(vals, fmts), start=1):
        cell = ws_kpi.cell(row=r_idx, column=c_idx, value=val)
        cell.font          = reg_font(10)
        cell.fill          = fill
        cell.alignment     = center()
        cell.border        = thin_border()
        cell.number_format = fmt

    # Colour-code OTD (col 8)
    otd_cell = ws_kpi.cell(row=r_idx, column=8)
    if row.OTD_Rate_Pct >= 90:
        otd_cell.font = bold_font(10, C_GREEN)
    elif row.OTD_Rate_Pct >= 80:
        otd_cell.font = bold_font(10, C_DARK_TEXT)
    else:
        otd_cell.font = bold_font(10, "C00000")

    # Colour-code Performance Score (col 11)
    ps_cell = ws_kpi.cell(row=r_idx, column=11)
    if score >= 85:
        ps_cell.fill = col_fill(C_GREEN); ps_cell.font = bold_font(10, C_WHITE)
    elif score >= 75:
        ps_cell.fill = col_fill(C_YELLOW); ps_cell.font = bold_font(10, C_DARK_TEXT)
    else:
        ps_cell.fill = col_fill("C00000"); ps_cell.font = bold_font(10, C_WHITE)

    set_row_height(ws_kpi, r_idx, 22)

# ── Summary totals row ───────────────────────────────────
total_row = len(kpi) + 3
ws_kpi.merge_cells(f"A{total_row}:D{total_row}")
tot_label = ws_kpi[f"A{total_row}"]
tot_label.value     = "PORTFOLIO TOTALS / AVERAGES"
tot_label.font      = bold_font(10, C_TITLE_FONT)
tot_label.fill      = col_fill(C_DARK_BLUE)
tot_label.alignment = center()

n = len(kpi)
summary_vals = {
    5:  f"=SUM(E3:E{total_row-1})",       # Total Spend
    6:  f"=AVERAGE(F3:F{total_row-1})",   # Avg Unit Cost
    7:  f"=AVERAGE(G3:G{total_row-1})",   # Avg Lead Time
    8:  f"=AVERAGE(H3:H{total_row-1})",   # Avg OTD
    9:  f"=AVERAGE(I3:I{total_row-1})",   # Avg Quality
    10: f"=AVERAGE(J3:J{total_row-1})",   # Avg Defect
    11: f"=AVERAGE(K3:K{total_row-1})",   # Avg Perf Score
}
fmts_tot = {5:"£#,##0", 6:"£#,##0.00", 7:"0.0", 8:"0.0", 9:"0.0", 10:"0.00", 11:"0.0"}

for col, formula in summary_vals.items():
    c = ws_kpi.cell(row=total_row, column=col, value=formula)
    c.font          = bold_font(10, C_TITLE_FONT)
    c.fill          = col_fill(C_DARK_BLUE)
    c.alignment     = center()
    c.border        = thin_border()
    c.number_format = fmts_tot[col]

set_row_height(ws_kpi, total_row, 28)

# ── Monthly trend table (used by charts) ─────────────────
trend_start_row = total_row + 3
ws_kpi.cell(row=trend_start_row, column=1, value="MONTHLY TREND DATA").font = bold_font(11, C_DARK_BLUE)

trend_headers = ["Month", "Total Spend (£)", "Avg Unit Cost (£)", "OTD Rate (%)"]
for c, h in enumerate(trend_headers, start=1):
    cell = ws_kpi.cell(row=trend_start_row+1, column=c, value=h)
    cell.font      = bold_font(10, C_TITLE_FONT)
    cell.fill      = col_fill(C_MID_BLUE)
    cell.alignment = center()
    cell.border    = thin_border()

for r_off, row in enumerate(monthly.itertuples(index=False), start=2):
    rr = trend_start_row + r_off
    vals = [row.Month, row.Total_Spend, row.Avg_Unit_Cost, row.OTD_Rate]
    fmts = ["@", "£#,##0", "£#,##0.00", "0.0"]
    fill = col_fill(C_LIGHT_BLUE) if r_off % 2 == 0 else col_fill(C_WHITE)
    for c, (v, fmt) in enumerate(zip(vals, fmts), start=1):
        cell = ws_kpi.cell(row=rr, column=c, value=v)
        cell.font          = reg_font(10)
        cell.fill          = fill
        cell.alignment     = center()
        cell.border        = thin_border()
        cell.number_format = fmt

n_months = len(monthly)
print("✅ KPI_Scorecard sheet written")

# ════════════════════════════════════════════════════════════
# SHEET C ▸ Dashboard  (charts + KPI cards)
# ════════════════════════════════════════════════════════════
ws_dash = wb.create_sheet("Dashboard")
ws_dash.sheet_view.showGridLines = False

# Set all rows/cols to a neutral look
for col_idx in range(1, 30):
    ws_dash.column_dimensions[get_column_letter(col_idx)].width = 4.5
for row_idx in range(1, 80):
    ws_dash.row_dimensions[row_idx].height = 15

# ── Dashboard Title ──────────────────────────────────────
ws_dash.merge_cells("A1:AC2")
dt = ws_dash["A1"]
dt.value     = "⬛  FMCG SUPPLIER PERFORMANCE DASHBOARD  |  FY 2023 – 2024"
dt.font      = bold_font(18, C_TITLE_FONT)
dt.fill      = col_fill(C_DARK_BLUE)
dt.alignment = center()
ws_dash.row_dimensions[1].height = 45
ws_dash.row_dimensions[2].height = 10

# ── KPI Summary Cards (row 3-6) ──────────────────────────
def make_card(ws, row, col_start, col_end, label, formula, fmt, accent_color):
    cl = get_column_letter

    # Row 1: label
    ws.merge_cells(f"{cl(col_start)}{row}:{cl(col_end)}{row}")
    c = ws.cell(row=row, column=col_start)
    c.value     = label
    c.font      = bold_font(9, "808080")
    c.fill      = col_fill(C_WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Row 2-3: big value
    ws.merge_cells(f"{cl(col_start)}{row+1}:{cl(col_end)}{row+2}")
    v = ws.cell(row=row+1, column=col_start)
    v.value         = formula
    v.font          = bold_font(22, accent_color)
    v.fill          = col_fill(C_WHITE)
    v.alignment     = center()
    v.number_format = fmt

    # Row 4: accent stripe
    ws.merge_cells(f"{cl(col_start)}{row+3}:{cl(col_end)}{row+3}")
    b = ws.cell(row=row+3, column=col_start)
    b.fill = col_fill(accent_color)

n_kpi_rows = len(kpi)
spend_formula   = f"=KPI_Scorecard!E{n_kpi_rows+3}"
orders_formula  = f"=SUM(KPI_Scorecard!D3:D{n_kpi_rows+2})"
otd_formula     = f"=KPI_Scorecard!H{n_kpi_rows+3}"
quality_formula = f"=KPI_Scorecard!I{n_kpi_rows+3}"
lead_formula    = f"=KPI_Scorecard!G{n_kpi_rows+3}"

make_card(ws_dash, 4,  1,  5,  "TOTAL SPEND",           spend_formula,   "£#,##0",  C_MID_BLUE)
make_card(ws_dash, 4,  6,  10, "TOTAL ORDERS",          orders_formula,  "#,##0",   C_ACCENT)
make_card(ws_dash, 4,  11, 15, "AVG OTD RATE",          otd_formula,     "0.0\"%\"", C_GREEN)
make_card(ws_dash, 4,  16, 20, "AVG QUALITY SCORE",     quality_formula, "0.0",     C_DARK_BLUE)
make_card(ws_dash, 4,  21, 25, "AVG LEAD TIME (DAYS)",  lead_formula,    "0.0",     "7030A0")

# ── CHART 1: Supplier Performance Score (Horizontal Bar) ─
chart1 = BarChart()
chart1.type    = "bar"   # horizontal
chart1.title   = "Supplier Performance Score (0–100)"
chart1.style   = 10
chart1.y_axis.title = "Supplier"
chart1.x_axis.title = "Score"
chart1.height  = 13
chart1.width   = 18
chart1.x_axis.scaling.min = 0
chart1.x_axis.scaling.max = 100

# Suppliers: KPI_Scorecard col B (2), Performance Score col K (11)
sup_names = Reference(ws_kpi, min_col=2, min_row=3, max_row=n_kpi_rows+2)
sup_scores = Reference(ws_kpi, min_col=11, min_row=2, max_row=n_kpi_rows+2)
chart1.add_data(sup_scores, titles_from_data=True)
chart1.set_categories(sup_names)
chart1.series[0].graphicalProperties.solidFill    = C_MID_BLUE
chart1.series[0].graphicalProperties.line.solidFill = C_MID_BLUE
chart1.dataLabels = DataLabelList()
chart1.dataLabels.showVal = True
ws_dash.add_chart(chart1, "A9")

# ── CHART 2: OTD Rate by Supplier (Column Bar) ───────────
chart2 = BarChart()
chart2.type    = "col"
chart2.title   = "On-Time Delivery Rate by Supplier (%)"
chart2.style   = 10
chart2.y_axis.title = "OTD Rate (%)"
chart2.y_axis.scaling.min = 0
chart2.y_axis.scaling.max = 100
chart2.height  = 13
chart2.width   = 18

otd_data = Reference(ws_kpi, min_col=8, min_row=2, max_row=n_kpi_rows+2)
chart2.add_data(otd_data, titles_from_data=True)
chart2.set_categories(sup_names)
chart2.series[0].graphicalProperties.solidFill     = C_ACCENT
chart2.series[0].graphicalProperties.line.solidFill = C_ACCENT
ws_dash.add_chart(chart2, "N9")

# ── CHART 3: Monthly Spend Trend (Line) ──────────────────
trend_data_start = total_row + 1   # header row of monthly table in KPI_Scorecard
trend_data_end   = trend_data_start + n_months

chart3 = LineChart()
chart3.title   = "Monthly Total Spend (£)"
chart3.style   = 10
chart3.y_axis.title = "Total Spend (£)"
chart3.height  = 12
chart3.width   = 22

spend_series = Reference(ws_kpi, min_col=2, min_row=trend_data_start,
                         max_row=trend_data_end)
month_labels = Reference(ws_kpi, min_col=1, min_row=trend_data_start+1,
                         max_row=trend_data_end)
chart3.add_data(spend_series, titles_from_data=True)
chart3.set_categories(month_labels)
chart3.series[0].graphicalProperties.line.solidFill = C_MID_BLUE
chart3.series[0].graphicalProperties.line.width     = 20000
chart3.series[0].smooth = True
ws_dash.add_chart(chart3, "A32")

# ── CHART 4: Monthly Avg Unit Cost Trend (Line) ──────────
chart4 = LineChart()
chart4.title   = "Monthly Avg Unit Cost (£/kg)"
chart4.style   = 10
chart4.y_axis.title = "Avg Unit Cost (£)"
chart4.height  = 12
chart4.width   = 22

cost_series = Reference(ws_kpi, min_col=3, min_row=trend_data_start,
                        max_row=trend_data_end)
chart4.add_data(cost_series, titles_from_data=True)
chart4.set_categories(month_labels)
chart4.series[0].graphicalProperties.line.solidFill = C_ACCENT
chart4.series[0].graphicalProperties.line.width     = 20000
chart4.series[0].smooth = True
ws_dash.add_chart(chart4, "N32")

# ── CHART 5: Quality vs Defect Rate (Dual column) ────────
chart5 = BarChart()
chart5.type    = "col"
chart5.title   = "Avg Quality Score vs Defect Rate by Supplier"
chart5.style   = 10
chart5.y_axis.title = "Score / Rate"
chart5.height  = 12
chart5.width   = 22
chart5.grouping = "clustered"

qual_data   = Reference(ws_kpi, min_col=9,  min_row=2, max_row=n_kpi_rows+2)
defect_data = Reference(ws_kpi, min_col=10, min_row=2, max_row=n_kpi_rows+2)
chart5.add_data(qual_data,   titles_from_data=True)
chart5.add_data(defect_data, titles_from_data=True)
chart5.set_categories(sup_names)
chart5.series[0].graphicalProperties.solidFill     = C_GREEN
chart5.series[0].graphicalProperties.line.solidFill = C_GREEN
chart5.series[1].graphicalProperties.solidFill     = "C00000"
chart5.series[1].graphicalProperties.line.solidFill = "C00000"
ws_dash.add_chart(chart5, "A50")

# ── CHART 6: Spend by Category (column bar) ──────────────
# Build a small helper table for category spend
cat_spend = df.groupby("Category")["Total_Cost_GBP"].sum().reset_index()
cat_spend.columns = ["Category", "Spend"]
cat_spend = cat_spend.sort_values("Spend", ascending=False)

ws_helper = wb.create_sheet("_Helper")   # hidden helper sheet
ws_helper["A1"] = "Category"
ws_helper["B1"] = "Total Spend (£)"
for i, (_, row) in enumerate(cat_spend.iterrows(), start=2):
    ws_helper[f"A{i}"] = row["Category"]
    ws_helper[f"B{i}"] = row["Spend"]

n_cats = len(cat_spend)
chart6 = BarChart()
chart6.type    = "col"
chart6.title   = "Total Spend by Category (£)"
chart6.style   = 10
chart6.y_axis.title = "Total Spend (£)"
chart6.height  = 12
chart6.width   = 22

cat_vals   = Reference(ws_helper, min_col=2, min_row=1, max_row=n_cats+1)
cat_labels = Reference(ws_helper, min_col=1, min_row=2, max_row=n_cats+1)
chart6.add_data(cat_vals, titles_from_data=True)
chart6.set_categories(cat_labels)
chart6.series[0].graphicalProperties.solidFill     = "7030A0"
chart6.series[0].graphicalProperties.line.solidFill = "7030A0"
ws_dash.add_chart(chart6, "N50")

# Hide helper sheet
ws_helper.sheet_state = "hidden"

# ── Footer ───────────────────────────────────────────────
ws_dash.merge_cells("A67:AC67")
footer = ws_dash["A67"]
footer.value     = "Data Source: Simulated FMCG Supply Chain Dataset (2023–2024)  |  Built with Python + openpyxl  |  For Power BI import, connect to 'Raw_Data' tab"
footer.font      = Font(name="Arial", size=8, italic=True, color="808080")
footer.alignment = center()

# ── Tab colours ──────────────────────────────────────────
ws_raw.sheet_properties.tabColor  = "2E75B6"
ws_kpi.sheet_properties.tabColor  = "ED7D31"
ws_dash.sheet_properties.tabColor = "70AD47"

# Reorder sheets: Dashboard first
wb.move_sheet("Dashboard", offset=-2)

# ── Save ──────────────────────────────────────────────────
out_path = "/home/claude/FMCG_Supplier_Dashboard.xlsx"
wb.save(out_path)
print(f"✅ Workbook saved → {out_path}")
