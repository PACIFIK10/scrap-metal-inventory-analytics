"""
Build the ABC/XYZ Classification & Price Volatility workbook (Project 2).

ABC = grades ranked by revenue contribution (cumulative % of total shipment value).
  A = top 80%, B = next 15%, C = bottom 5%.
XYZ = grades ranked by demand variability (coefficient of variation of monthly shipments).
  X = CoV <= 0.5 (stable), Y = 0.5 < CoV <= 1.0 (variable), Z = CoV > 1.0 (erratic).

Sheets:
  README          — purpose, methodology, sheet guide
  ABC_Analysis    — grades ranked by cumulative revenue share
  XYZ_Analysis    — grades ranked by demand coefficient of variation
  ABC_XYZ_Matrix  — combined 3x3 classification with strategy recommendations
  Price_Volatility — grade-level price stats, 30/90-day trends, volatility ranking
  Strategy_Summary — actionable recommendations by ABC-XYZ cell
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import sys

# ── paths ──────────────────────────────────────────────────────────────
DATA = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("/mnt/user-data/uploads/emr-inventory-project")
OUT  = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("/home/claude/project2/abc_xyz_classification.xlsx")

# ── load data ──────────────────────────────────────────────────────────
txn    = pd.read_csv(DATA / "data/raw/transactions.csv", parse_dates=["date"])
grades = pd.read_csv(DATA / "data/reference/grades.csv")
yards  = pd.read_csv(DATA / "data/reference/yards.csv")
prices = pd.read_csv(DATA / "data/raw/grade_prices_daily.csv", parse_dates=["date"])
me     = pd.read_csv(DATA / "data/system_snapshots/inventory_month_end.csv", parse_dates=["month_end"])

# ── reference maps ─────────────────────────────────────────────────────
grade_map  = grades.set_index("grade_id")["metal_group"].to_dict()
grade_name = grades.set_index("grade_id")["grade_name"].to_dict()
proc_cost  = grades.set_index("grade_id")["processing_cost_per_ton"].to_dict()

# ── filter transactions to reference grades only ──────────────────────
valid_grades = set(grades["grade_id"])
txn = txn[txn["grade_id"].isin(valid_grades)].copy()
prices = prices[prices["grade_id"].isin(valid_grades)].copy()
me = me[me["grade_id"].isin(valid_grades)].copy()
print(f"Filtered to {len(valid_grades)} reference grades")

# ── monthly period ─────────────────────────────────────────────────────
txn["month"] = txn["date"].dt.to_period("M")
n_months = txn["month"].nunique()
print(f"Data spans {n_months} months")

# ══════════════════════════════════════════════════════════════════════
# ABC ANALYSIS — by shipment revenue (tons * unit_price for SHIPMENT txns)
# ══════════════════════════════════════════════════════════════════════
ships = txn[txn["txn_type"] == "SHIPMENT"].copy()
ships["revenue"] = ships["tons"] * ships["unit_price"]

abc_data = (ships.groupby("grade_id")
    .agg(
        total_tons_shipped=("tons", "sum"),
        total_revenue=("revenue", "sum"),
        shipment_count=("txn_id", "count"),
        avg_price=("unit_price", "mean"),
    ).reset_index()
    .sort_values("total_revenue", ascending=False))

total_revenue = abc_data["total_revenue"].sum()
abc_data["revenue_pct"] = abc_data["total_revenue"] / total_revenue
abc_data["cumulative_pct"] = abc_data["revenue_pct"].cumsum()

# ABC classification
def abc_class(cum_pct):
    if cum_pct <= 0.80:
        return "A"
    elif cum_pct <= 0.95:
        return "B"
    else:
        return "C"

abc_data["abc_class"] = abc_data["cumulative_pct"].apply(abc_class)
abc_data["grade_name"] = abc_data["grade_id"].map(grade_name)
abc_data["metal_group"] = abc_data["grade_id"].map(grade_map)
abc_data["rank"] = range(1, len(abc_data) + 1)

print(f"ABC: A={sum(abc_data['abc_class']=='A')}, B={sum(abc_data['abc_class']=='B')}, C={sum(abc_data['abc_class']=='C')}")

# ══════════════════════════════════════════════════════════════════════
# XYZ ANALYSIS — by coefficient of variation of monthly shipment tons
# ══════════════════════════════════════════════════════════════════════
monthly_ships = (ships.groupby(["grade_id", "month"])["tons"]
    .sum().reset_index())

# Ensure every grade has an entry for every month (0 if no shipments)
all_months = sorted(txn["month"].unique())
all_grades = sorted(grades["grade_id"])
full_index = pd.MultiIndex.from_product([all_grades, all_months], names=["grade_id", "month"])
monthly_full = (monthly_ships.set_index(["grade_id", "month"])
    .reindex(full_index, fill_value=0)
    .reset_index())

xyz_data = (monthly_full.groupby("grade_id")["tons"]
    .agg(["mean", "std", "sum"])
    .reset_index()
    .rename(columns={"mean": "avg_monthly_tons", "std": "std_monthly_tons", "sum": "total_tons"}))

xyz_data["cov"] = xyz_data["std_monthly_tons"] / xyz_data["avg_monthly_tons"]
xyz_data["cov"] = xyz_data["cov"].replace([np.inf, -np.inf], 0).fillna(0)

# XYZ classification
def xyz_class(cov):
    if cov <= 0.50:
        return "X"
    elif cov <= 1.00:
        return "Y"
    else:
        return "Z"

xyz_data["xyz_class"] = xyz_data["cov"].apply(xyz_class)
xyz_data["grade_name"] = xyz_data["grade_id"].map(grade_name)
xyz_data["metal_group"] = xyz_data["grade_id"].map(grade_map)
xyz_data = xyz_data.sort_values("cov")

print(f"XYZ: X={sum(xyz_data['xyz_class']=='X')}, Y={sum(xyz_data['xyz_class']=='Y')}, Z={sum(xyz_data['xyz_class']=='Z')}")

# ══════════════════════════════════════════════════════════════════════
# ABC-XYZ MATRIX — combine both classifications
# ══════════════════════════════════════════════════════════════════════
matrix = abc_data[["grade_id", "abc_class", "total_revenue", "revenue_pct", "total_tons_shipped", "avg_price"]].merge(
    xyz_data[["grade_id", "xyz_class", "cov", "avg_monthly_tons"]],
    on="grade_id"
)
matrix["grade_name"] = matrix["grade_id"].map(grade_name)
matrix["metal_group"] = matrix["grade_id"].map(grade_map)
matrix["abc_xyz"] = matrix["abc_class"] + matrix["xyz_class"]

# Current inventory
latest_month = me["month_end"].max()
inv_latest = me[me["month_end"] == latest_month].groupby("grade_id").agg(
    tons_on_hand=("tons_on_hand", "sum"),
    book_value=("book_value", "sum"),
    market_value=("market_value", "sum"),
    tons_over_90=("tons_over_90", "sum"),
).reset_index()
matrix = matrix.merge(inv_latest, on="grade_id", how="left")
matrix = matrix.sort_values(["abc_class", "xyz_class", "total_revenue"], ascending=[True, True, False])

# Strategy map
strategy_map = {
    "AX": "Core stock: high value, stable demand. Maintain safety stock, negotiate volume contracts, JIT replenishment.",
    "AY": "Plan carefully: high value, variable demand. Forecast closely, maintain moderate buffer, review weekly.",
    "AZ": "Critical watch: high value, erratic demand. Hold strategic buffer, source from multiple suppliers, frequent price reviews.",
    "BX": "Reliable secondary: moderate value, stable. Automate reorder, standard safety stock, quarterly review.",
    "BY": "Active management: moderate value, variable. Monthly forecast updates, flexible supplier terms.",
    "BZ": "Reduce exposure: moderate value, erratic. Minimize holdings, buy on demand, consider spot market.",
    "CX": "Low-touch: low value, stable. Min-max reorder, bulk buy for discounts, infrequent review.",
    "CY": "Simplify: low value, variable. Consolidate suppliers, consider dropping marginal grades.",
    "CZ": "Exit candidates: low value, erratic. Phase out or make-to-order only, no safety stock.",
}
matrix["strategy"] = matrix["abc_xyz"].map(strategy_map)

# Count by cell
cell_counts = matrix.groupby("abc_xyz").agg(
    count=("grade_id", "count"),
    total_revenue=("total_revenue", "sum"),
    total_tons=("tons_on_hand", "sum"),
).reset_index()

print(f"\nABC-XYZ matrix cells:")
for _, r in cell_counts.iterrows():
    print(f"  {r['abc_xyz']}: {r['count']} grades, ${r['total_revenue']:,.0f} revenue, {r['total_tons']:,.0f} tons on hand")

# ══════════════════════════════════════════════════════════════════════
# PRICE VOLATILITY ANALYSIS
# ══════════════════════════════════════════════════════════════════════
# Monthly average prices
prices["month"] = prices["date"].dt.to_period("M")
monthly_prices = prices.groupby(["grade_id", "month"]).agg(
    avg_buy=("buy_price", "mean"),
    avg_sell=("sell_price", "mean"),
).reset_index()
monthly_prices["spread"] = monthly_prices["avg_sell"] - monthly_prices["avg_buy"]

# Price volatility stats
price_vol = monthly_prices.groupby("grade_id").agg(
    mean_buy=("avg_buy", "mean"),
    std_buy=("avg_buy", "std"),
    min_buy=("avg_buy", "min"),
    max_buy=("avg_buy", "max"),
    mean_spread=("spread", "mean"),
    std_spread=("spread", "std"),
).reset_index()

price_vol["buy_cov"] = price_vol["std_buy"] / price_vol["mean_buy"]
price_vol["buy_range_pct"] = (price_vol["max_buy"] - price_vol["min_buy"]) / price_vol["mean_buy"]
price_vol["spread_cov"] = price_vol["std_spread"] / price_vol["mean_spread"]
price_vol = price_vol.replace([np.inf, -np.inf], 0).fillna(0)

# Latest prices and short-term trends
latest_date = prices["date"].max()
prices_latest = prices[prices["date"] == latest_date].set_index("grade_id")

# 30-day and 90-day trends
for days, label in [(30, "30d"), (90, "90d")]:
    cutoff = latest_date - pd.Timedelta(days=days)
    period_prices = prices[prices["date"] > cutoff].groupby("grade_id")["buy_price"].agg(["first", "last"])
    period_prices[f"trend_{label}"] = (period_prices["last"] - period_prices["first"]) / period_prices["first"]
    price_vol = price_vol.merge(period_prices[[f"trend_{label}"]], left_on="grade_id", right_index=True, how="left")

price_vol["grade_name"] = price_vol["grade_id"].map(grade_name)
price_vol["metal_group"] = price_vol["grade_id"].map(grade_map)

# Add current price
price_vol["current_buy"] = price_vol["grade_id"].map(
    lambda g: prices_latest.loc[g, "buy_price"] if g in prices_latest.index else 0)
price_vol["current_sell"] = price_vol["grade_id"].map(
    lambda g: prices_latest.loc[g, "sell_price"] if g in prices_latest.index else 0)

# Volatility rank (high CoV = high volatility)
price_vol = price_vol.sort_values("buy_cov", ascending=False)
price_vol["vol_rank"] = range(1, len(price_vol) + 1)

# Add ABC-XYZ class for cross-reference
abc_xyz_map = matrix.set_index("grade_id")["abc_xyz"].to_dict()
price_vol["abc_xyz"] = price_vol["grade_id"].map(abc_xyz_map)

print(f"\nPrice volatility: top 5 most volatile grades:")
for _, r in price_vol.head(5).iterrows():
    print(f"  {r['grade_id']} ({r['grade_name']}): buy CoV={r['buy_cov']:.3f}, range={r['buy_range_pct']:.1%}")

# ══════════════════════════════════════════════════════════════════════
# BUILD WORKBOOK
# ══════════════════════════════════════════════════════════════════════
wb = Workbook()

# ── styles ─────────────────────────────────────────────────────────────
FONT_TITLE  = Font(name="Arial", size=16, bold=True, color="1F3864")
FONT_SECT   = Font(name="Arial", size=13, bold=True, color="1F3864")
FONT_HEAD   = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FONT_BODY   = Font(name="Arial", size=10)
FONT_BODY_B = Font(name="Arial", size=10, bold=True)
FONT_KPI_V  = Font(name="Arial", size=14, bold=True, color="1F3864")
FONT_KPI_L  = Font(name="Arial", size=10, color="595959")
FONT_RED    = Font(name="Arial", size=10, color="C00000")
FONT_GREEN  = Font(name="Arial", size=10, color="006100")
FONT_LINK   = Font(name="Arial", size=10, italic=True, color="808080")

FILL_HEADER = PatternFill("solid", fgColor="1F3864")
FILL_LIGHT  = PatternFill("solid", fgColor="F2F7FB")
FILL_A      = PatternFill("solid", fgColor="C6EFCE")  # green — high value
FILL_B      = PatternFill("solid", fgColor="FFEB9C")  # yellow — moderate
FILL_C      = PatternFill("solid", fgColor="FFC7CE")  # red — low value
FILL_X      = PatternFill("solid", fgColor="C6EFCE")  # green — stable
FILL_Y      = PatternFill("solid", fgColor="FFEB9C")  # yellow — variable
FILL_Z      = PatternFill("solid", fgColor="FFC7CE")  # red — erratic
FILL_KPI    = PatternFill("solid", fgColor="E8EEF7")

# ABC-XYZ cell colors (green=good combo, yellow=needs attention, red=risky)
FILL_ABCXYZ = {
    "AX": PatternFill("solid", fgColor="C6EFCE"),
    "AY": PatternFill("solid", fgColor="FFEB9C"),
    "AZ": PatternFill("solid", fgColor="FFC7CE"),
    "BX": PatternFill("solid", fgColor="D6E4F0"),
    "BY": PatternFill("solid", fgColor="FFEB9C"),
    "BZ": PatternFill("solid", fgColor="FFC7CE"),
    "CX": PatternFill("solid", fgColor="D6E4F0"),
    "CY": PatternFill("solid", fgColor="FFEB9C"),
    "CZ": PatternFill("solid", fgColor="FFC7CE"),
}

ALIGN_C = Alignment(horizontal="center", vertical="center")
ALIGN_R = Alignment(horizontal="right", vertical="center")
ALIGN_L = Alignment(horizontal="left", vertical="center", wrap_text=True)

BORDER = Border(
    bottom=Side(style="thin", color="B4C6E7"),
    top=Side(style="thin", color="B4C6E7"),
    left=Side(style="thin", color="B4C6E7"),
    right=Side(style="thin", color="B4C6E7"),
)

FMT_NUM    = '#,##0'
FMT_DEC1   = '#,##0.0'
FMT_DEC2   = '#,##0.00'
FMT_DEC3   = '0.000'
FMT_DOLLAR = '$#,##0'
FMT_DOLLAR2 = '$#,##0.00'
FMT_PCT    = '0.0%'
FMT_PCT1   = '0.0%'


def write_header(ws, row, headers, col_start=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col_start + i, value=h)
        c.font = FONT_HEAD
        c.fill = FILL_HEADER
        c.alignment = ALIGN_C
        c.border = BORDER


def write_row(ws, row, vals, fmts=None, col_start=1, alt=False):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=col_start + i, value=v)
        c.font = FONT_BODY
        c.border = BORDER
        if alt:
            c.fill = FILL_LIGHT
        if fmts and i < len(fmts) and fmts[i]:
            c.number_format = fmts[i]
            if fmts[i] in (FMT_DOLLAR, FMT_DOLLAR2, FMT_NUM, FMT_DEC1, FMT_DEC2, FMT_DEC3, FMT_PCT):
                c.alignment = ALIGN_R
            else:
                c.alignment = ALIGN_L
        else:
            c.alignment = ALIGN_L


def auto_width(ws, min_w=10, max_w=32):
    for col in ws.columns:
        mx = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                mx = max(mx, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max(mx + 2, min_w), max_w)


# ═══════════════════════════════════════════════════════════════════════
# SHEET 1: README
# ═══════════════════════════════════════════════════════════════════════
ws_r = wb.active
ws_r.title = "README"
ws_r.sheet_properties.tabColor = "808080"

readme = [
    ("ABC/XYZ CLASSIFICATION & PRICE VOLATILITY ANALYSIS", FONT_TITLE),
    ("", None),
    ("PURPOSE", FONT_SECT),
    ("Classifies 40 scrap-metal grades into a 3x3 matrix by value contribution (ABC) and demand stability (XYZ), then overlays price volatility to guide purchasing, stocking, and disposal decisions. Designed for a Commercial Planning & Inventory Analyst deciding which grades to stock aggressively, manage carefully, or phase out.", FONT_BODY),
    ("", None),
    ("METHODOLOGY", FONT_SECT),
    ("ABC Analysis: grades ranked by cumulative share of total shipment revenue (24 months). A = top 80%, B = next 15%, C = bottom 5%.", FONT_BODY),
    ("XYZ Analysis: grades ranked by coefficient of variation (std / mean) of monthly shipment tons. X = CoV <= 0.5 (stable), Y = 0.5-1.0 (variable), Z = CoV > 1.0 (erratic).", FONT_BODY),
    ("ABC-XYZ Matrix: combines both dimensions into 9 cells (AX, AY, AZ, BX, BY, BZ, CX, CY, CZ), each with a distinct inventory strategy.", FONT_BODY),
    ("Price Volatility: buy-price coefficient of variation, 24-month range, and 30/90-day trends per grade.", FONT_BODY),
    ("", None),
    ("SHEET GUIDE", FONT_SECT),
    ("README — this page", FONT_BODY),
    ("ABC_Analysis — grades ranked by cumulative revenue share with A/B/C classification", FONT_BODY),
    ("XYZ_Analysis — grades ranked by demand coefficient of variation with X/Y/Z classification", FONT_BODY),
    ("ABC_XYZ_Matrix — combined classification with current inventory position and strategy", FONT_BODY),
    ("Price_Volatility — price statistics, trends, and volatility ranking for every grade", FONT_BODY),
    ("Strategy_Summary — 3x3 matrix of recommended actions by ABC-XYZ cell", FONT_BODY),
    ("", None),
    ("KEY FINDINGS", FONT_SECT),
    (f"A grades: {sum(abc_data['abc_class']=='A')} grades drive 80% of revenue (${abc_data[abc_data['abc_class']=='A']['total_revenue'].sum():,.0f})", FONT_BODY),
    (f"C grades: {sum(abc_data['abc_class']=='C')} grades contribute only 5% — candidates for rationalization", FONT_BODY),
    (f"X (stable demand): {sum(xyz_data['xyz_class']=='X')} grades — safe for JIT and volume contracts", FONT_BODY),
    (f"Z (erratic demand): {sum(xyz_data['xyz_class']=='Z')} grades — buy on demand, minimize stock", FONT_BODY),
    ("", None),
    ("DATA NOTE", FONT_SECT),
    ("All data is synthetic, simulating a scrap-metal yard network (8 yards, 40 grades, 24 months).", FONT_BODY),
]
for i, (text, font) in enumerate(readme):
    c = ws_r.cell(row=i + 1, column=1, value=text)
    if font:
        c.font = font
ws_r.column_dimensions["A"].width = 110

# ═══════════════════════════════════════════════════════════════════════
# SHEET 2: ABC_Analysis
# ═══════════════════════════════════════════════════════════════════════
ws_abc = wb.create_sheet("ABC_Analysis")
ws_abc.sheet_properties.tabColor = "548235"

ws_abc.cell(row=1, column=1, value="ABC Analysis — Grades by Revenue Contribution").font = FONT_TITLE
ws_abc.cell(row=2, column=1, value=f"Based on 24 months of shipment data  |  Total revenue: ${total_revenue:,.0f}").font = FONT_KPI_L
ws_abc.merge_cells("A1:J1")

headers_abc = ["Rank", "Grade ID", "Grade Name", "Metal Group", "ABC Class",
               "Total Revenue ($)", "Revenue (%)", "Cumulative (%)",
               "Tons Shipped", "Shipments", "Avg Price ($/ton)"]
write_header(ws_abc, 4, headers_abc)

fmts_abc = [FMT_NUM, None, None, None, None, FMT_DOLLAR, FMT_PCT, FMT_PCT,
            FMT_DEC1, FMT_NUM, FMT_DOLLAR2]

abc_fill = {"A": FILL_A, "B": FILL_B, "C": FILL_C}

for i, (_, row) in enumerate(abc_data.iterrows()):
    vals = [row["rank"], row["grade_id"], row["grade_name"], row["metal_group"],
            row["abc_class"], row["total_revenue"], row["revenue_pct"],
            row["cumulative_pct"], row["total_tons_shipped"],
            row["shipment_count"], row["avg_price"]]
    write_row(ws_abc, 5 + i, vals, fmts_abc)

    # Color the ABC class cell
    cls_cell = ws_abc.cell(row=5 + i, column=5)
    cls_cell.fill = abc_fill.get(row["abc_class"], FILL_LIGHT)
    cls_cell.alignment = ALIGN_C
    cls_cell.font = FONT_BODY_B

# Summary below
sr = 5 + len(abc_data) + 2
ws_abc.cell(row=sr, column=1, value="Classification Summary").font = FONT_SECT
write_header(ws_abc, sr + 1, ["Class", "# Grades", "Total Revenue ($)", "Revenue Share (%)", "Avg Tons/Grade"])
for ci, cls in enumerate(["A", "B", "C"]):
    d = abc_data[abc_data["abc_class"] == cls]
    vals = [cls, len(d), d["total_revenue"].sum(), d["revenue_pct"].sum(),
            d["total_tons_shipped"].mean()]
    write_row(ws_abc, sr + 2 + ci, vals, [None, FMT_NUM, FMT_DOLLAR, FMT_PCT, FMT_DEC1])
    ws_abc.cell(row=sr + 2 + ci, column=1).fill = abc_fill[cls]
    ws_abc.cell(row=sr + 2 + ci, column=1).font = FONT_BODY_B

auto_width(ws_abc)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 3: XYZ_Analysis
# ═══════════════════════════════════════════════════════════════════════
ws_xyz = wb.create_sheet("XYZ_Analysis")
ws_xyz.sheet_properties.tabColor = "BF8F00"

ws_xyz.cell(row=1, column=1, value="XYZ Analysis — Grades by Demand Variability").font = FONT_TITLE
ws_xyz.cell(row=2, column=1, value="Coefficient of Variation (std / mean) of monthly shipment tons over 24 months").font = FONT_KPI_L
ws_xyz.merge_cells("A1:H1")

headers_xyz = ["Grade ID", "Grade Name", "Metal Group", "XYZ Class",
               "Avg Monthly Tons", "Std Dev", "Coeff of Variation",
               "Total Tons Shipped"]
write_header(ws_xyz, 4, headers_xyz)

fmts_xyz = [None, None, None, None, FMT_DEC1, FMT_DEC1, FMT_DEC3, FMT_DEC1]

xyz_fill = {"X": FILL_X, "Y": FILL_Y, "Z": FILL_Z}

for i, (_, row) in enumerate(xyz_data.iterrows()):
    vals = [row["grade_id"], row["grade_name"], row["metal_group"],
            row["xyz_class"], row["avg_monthly_tons"], row["std_monthly_tons"],
            row["cov"], row["total_tons"]]
    write_row(ws_xyz, 5 + i, vals, fmts_xyz)

    cls_cell = ws_xyz.cell(row=5 + i, column=4)
    cls_cell.fill = xyz_fill.get(row["xyz_class"], FILL_LIGHT)
    cls_cell.alignment = ALIGN_C
    cls_cell.font = FONT_BODY_B

# Summary
sr = 5 + len(xyz_data) + 2
ws_xyz.cell(row=sr, column=1, value="Classification Summary").font = FONT_SECT
write_header(ws_xyz, sr + 1, ["Class", "Description", "# Grades", "CoV Range", "Avg Monthly Tons"])
xyz_desc = {"X": "Stable demand (CoV <= 0.5)", "Y": "Variable demand (0.5 < CoV <= 1.0)", "Z": "Erratic demand (CoV > 1.0)"}
for ci, cls in enumerate(["X", "Y", "Z"]):
    d = xyz_data[xyz_data["xyz_class"] == cls]
    cov_range = f"{d['cov'].min():.3f} - {d['cov'].max():.3f}" if len(d) > 0 else "N/A"
    vals = [cls, xyz_desc[cls], len(d), cov_range, d["avg_monthly_tons"].mean() if len(d) > 0 else 0]
    write_row(ws_xyz, sr + 2 + ci, vals, [None, None, FMT_NUM, None, FMT_DEC1])
    ws_xyz.cell(row=sr + 2 + ci, column=1).fill = xyz_fill[cls]
    ws_xyz.cell(row=sr + 2 + ci, column=1).font = FONT_BODY_B

auto_width(ws_xyz)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 4: ABC_XYZ_Matrix
# ═══════════════════════════════════════════════════════════════════════
ws_mx = wb.create_sheet("ABC_XYZ_Matrix")
ws_mx.sheet_properties.tabColor = "2E75B6"

ws_mx.cell(row=1, column=1, value="ABC-XYZ Classification Matrix").font = FONT_TITLE
ws_mx.cell(row=2, column=1, value=f"40 grades classified into 9 cells  |  As of {latest_month.strftime('%B %Y')}").font = FONT_KPI_L
ws_mx.merge_cells("A1:N1")

headers_mx = ["Grade ID", "Grade Name", "Metal Group", "ABC", "XYZ", "ABC-XYZ",
              "Revenue ($)", "Revenue (%)", "CoV", "Avg Monthly Tons",
              "Tons on Hand", "Book Value ($)", "Market Value ($)",
              "Tons 90+ Days", "Strategy"]
write_header(ws_mx, 4, headers_mx)

fmts_mx = [None, None, None, None, None, None, FMT_DOLLAR, FMT_PCT, FMT_DEC3,
           FMT_DEC1, FMT_DEC1, FMT_DOLLAR, FMT_DOLLAR, FMT_DEC1, None]

for i, (_, row) in enumerate(matrix.iterrows()):
    vals = [row["grade_id"], row["grade_name"], row["metal_group"],
            row["abc_class"], row["xyz_class"], row["abc_xyz"],
            row["total_revenue"], row["revenue_pct"], row["cov"],
            row["avg_monthly_tons"],
            row.get("tons_on_hand", 0), row.get("book_value", 0),
            row.get("market_value", 0), row.get("tons_over_90", 0),
            row["strategy"]]
    write_row(ws_mx, 5 + i, vals, fmts_mx)

    # Color the ABC-XYZ cell
    abcxyz_cell = ws_mx.cell(row=5 + i, column=6)
    abcxyz_cell.fill = FILL_ABCXYZ.get(row["abc_xyz"], FILL_LIGHT)
    abcxyz_cell.alignment = ALIGN_C
    abcxyz_cell.font = FONT_BODY_B

    # ABC and XYZ cells
    ws_mx.cell(row=5 + i, column=4).fill = abc_fill.get(row["abc_class"], FILL_LIGHT)
    ws_mx.cell(row=5 + i, column=4).alignment = ALIGN_C
    ws_mx.cell(row=5 + i, column=5).fill = xyz_fill.get(row["xyz_class"], FILL_LIGHT)
    ws_mx.cell(row=5 + i, column=5).alignment = ALIGN_C

auto_width(ws_mx)
ws_mx.column_dimensions["O"].width = 60  # Strategy column

# ═══════════════════════════════════════════════════════════════════════
# SHEET 5: Price_Volatility
# ═══════════════════════════════════════════════════════════════════════
ws_pv = wb.create_sheet("Price_Volatility")
ws_pv.sheet_properties.tabColor = "7030A0"

ws_pv.cell(row=1, column=1, value="Price Volatility Analysis by Grade").font = FONT_TITLE
ws_pv.cell(row=2, column=1, value=f"24-month price statistics  |  Prices as of {latest_date.strftime('%Y-%m-%d')}").font = FONT_KPI_L
ws_pv.merge_cells("A1:N1")

headers_pv = ["Vol Rank", "Grade ID", "Grade Name", "Metal Group", "ABC-XYZ",
              "Current Buy ($/t)", "Current Sell ($/t)", "Current Spread ($/t)",
              "Mean Buy ($/t)", "Buy Std Dev", "Buy CoV",
              "24m Range (%)", "30-Day Trend", "90-Day Trend"]
write_header(ws_pv, 4, headers_pv)

fmts_pv = [FMT_NUM, None, None, None, None, FMT_DOLLAR2, FMT_DOLLAR2, FMT_DOLLAR2,
           FMT_DOLLAR2, FMT_DOLLAR2, FMT_DEC3, FMT_PCT, FMT_PCT, FMT_PCT]

for i, (_, row) in enumerate(price_vol.iterrows()):
    spread = row["current_sell"] - row["current_buy"]
    vals = [row["vol_rank"], row["grade_id"], row["grade_name"], row["metal_group"],
            row.get("abc_xyz", ""), row["current_buy"], row["current_sell"], spread,
            row["mean_buy"], row["std_buy"], row["buy_cov"],
            row["buy_range_pct"],
            row.get("trend_30d", 0), row.get("trend_90d", 0)]
    write_row(ws_pv, 5 + i, vals, fmts_pv, alt=i % 2 == 0)

    # Color trends
    for col_idx in (13, 14):
        trend_cell = ws_pv.cell(row=5 + i, column=col_idx)
        val = vals[col_idx - 1]
        if isinstance(val, (int, float)):
            if val > 0.02:
                trend_cell.font = FONT_GREEN
            elif val < -0.02:
                trend_cell.font = FONT_RED

    # Color ABC-XYZ
    abcxyz_cell = ws_pv.cell(row=5 + i, column=5)
    abcxyz_cell.fill = FILL_ABCXYZ.get(row.get("abc_xyz", ""), FILL_LIGHT)
    abcxyz_cell.alignment = ALIGN_C

auto_width(ws_pv)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 6: Strategy_Summary (3x3 grid)
# ═══════════════════════════════════════════════════════════════════════
ws_st = wb.create_sheet("Strategy_Summary")
ws_st.sheet_properties.tabColor = "C00000"

ws_st.cell(row=1, column=1, value="ABC-XYZ Strategy Matrix").font = FONT_TITLE
ws_st.cell(row=2, column=1, value="Inventory management strategy by classification cell").font = FONT_KPI_L
ws_st.merge_cells("A1:E1")

# Build a visual 3x3 matrix
r = 4
ws_st.cell(row=r, column=1, value="").font = FONT_BODY
ws_st.cell(row=r, column=2, value="X (Stable)").font = FONT_BODY_B
ws_st.cell(row=r, column=2).alignment = ALIGN_C
ws_st.cell(row=r, column=3, value="Y (Variable)").font = FONT_BODY_B
ws_st.cell(row=r, column=3).alignment = ALIGN_C
ws_st.cell(row=r, column=4, value="Z (Erratic)").font = FONT_BODY_B
ws_st.cell(row=r, column=4).alignment = ALIGN_C

for ri, abc in enumerate(["A", "B", "C"]):
    row_n = r + 1 + ri * 3  # 3 rows per cell (label + count + strategy)
    abc_labels = {"A": "A (High Value)", "B": "B (Medium Value)", "C": "C (Low Value)"}
    ws_st.cell(row=row_n, column=1, value=abc_labels[abc]).font = FONT_BODY_B
    ws_st.cell(row=row_n, column=1).alignment = ALIGN_C

    for ci, xyz in enumerate(["X", "Y", "Z"]):
        cell_code = f"{abc}{xyz}"
        col = 2 + ci
        cell_data = matrix[matrix["abc_xyz"] == cell_code]
        count = len(cell_data)
        rev = cell_data["total_revenue"].sum()

        # Cell code and count
        c = ws_st.cell(row=row_n, column=col, value=f"{cell_code}: {count} grades")
        c.font = FONT_BODY_B
        c.fill = FILL_ABCXYZ.get(cell_code, FILL_LIGHT)
        c.alignment = ALIGN_C
        c.border = BORDER

        # Revenue
        c2 = ws_st.cell(row=row_n + 1, column=col, value=f"${rev:,.0f}")
        c2.font = FONT_BODY
        c2.fill = FILL_ABCXYZ.get(cell_code, FILL_LIGHT)
        c2.alignment = ALIGN_C
        c2.border = BORDER

        # Strategy (short)
        short_strat = strategy_map.get(cell_code, "").split(".")[0] + "."
        c3 = ws_st.cell(row=row_n + 2, column=col, value=short_strat)
        c3.font = FONT_BODY
        c3.fill = FILL_ABCXYZ.get(cell_code, FILL_LIGHT)
        c3.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c3.border = BORDER

# Detailed strategies below the matrix
sr = r + 1 + 9 + 2
ws_st.cell(row=sr, column=1, value="Detailed Strategies").font = FONT_SECT
write_header(ws_st, sr + 1, ["Cell", "# Grades", "Revenue ($)", "Strategy"])

sorted_cells = ["AX", "AY", "AZ", "BX", "BY", "BZ", "CX", "CY", "CZ"]
for ci, cell_code in enumerate(sorted_cells):
    cell_data = matrix[matrix["abc_xyz"] == cell_code]
    vals = [cell_code, len(cell_data), cell_data["total_revenue"].sum(),
            strategy_map.get(cell_code, "")]
    write_row(ws_st, sr + 2 + ci, vals, [None, FMT_NUM, FMT_DOLLAR, None])
    ws_st.cell(row=sr + 2 + ci, column=1).fill = FILL_ABCXYZ.get(cell_code, FILL_LIGHT)
    ws_st.cell(row=sr + 2 + ci, column=1).font = FONT_BODY_B
    ws_st.cell(row=sr + 2 + ci, column=1).alignment = ALIGN_C

# Grade lists per cell
sr2 = sr + 2 + len(sorted_cells) + 2
ws_st.cell(row=sr2, column=1, value="Grades by Cell").font = FONT_SECT
write_header(ws_st, sr2 + 1, ["Cell", "Grade ID", "Grade Name", "Revenue ($)", "CoV"])
row_n = sr2 + 2
for cell_code in sorted_cells:
    cell_data = matrix[matrix["abc_xyz"] == cell_code].sort_values("total_revenue", ascending=False)
    for _, gd in cell_data.iterrows():
        vals = [cell_code, gd["grade_id"], gd["grade_name"], gd["total_revenue"], gd["cov"]]
        write_row(ws_st, row_n, vals, [None, None, None, FMT_DOLLAR, FMT_DEC3])
        ws_st.cell(row=row_n, column=1).fill = FILL_ABCXYZ.get(cell_code, FILL_LIGHT)
        ws_st.cell(row=row_n, column=1).alignment = ALIGN_C
        row_n += 1

auto_width(ws_st)
ws_st.column_dimensions["B"].width = 18
ws_st.column_dimensions["C"].width = 22
ws_st.column_dimensions["D"].width = 60

# ── hide gridlines ─────────────────────────────────────────────────────
for ws_sheet in wb.worksheets:
    ws_sheet.sheet_view.showGridLines = False

# ── save ───────────────────────────────────────────────────────────────
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print(f"\nWorkbook saved to {OUT}")
print(f"Sheets: {wb.sheetnames}")
