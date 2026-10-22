"""
Build the Inventory Positioning & Movement Plan workbook (Project 1).

This is the prescriptive/action-oriented workbook:
  - Where should each grade be stocked (positioning)?
  - How much safety stock per yard-grade (target levels)?
  - Which yards are overstocked or understocked relative to demand?
  - What inter-yard transfers would rebalance the network?

Sheets:
  README              — purpose, methodology, sheet guide
  Positioning_Summary — target vs actual stock by yard-grade, gap analysis
  Rebalance_Plan      — recommended inter-yard transfers with priority
  Demand_Analysis     — monthly demand trends and seasonality by grade
  Yard_Capacity       — utilization, headroom, risk flags per yard
  Grade_Strategy      — grade-level positioning recommendations informed by ABC-XYZ
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
OUT  = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("/home/claude/project1/inventory_positioning.xlsx")

# ── load data ──────────────────────────────────────────────────────────
txn    = pd.read_csv(DATA / "data/raw/transactions.csv", parse_dates=["date"])
grades = pd.read_csv(DATA / "data/reference/grades.csv")
yards  = pd.read_csv(DATA / "data/reference/yards.csv")
prices = pd.read_csv(DATA / "data/raw/grade_prices_daily.csv", parse_dates=["date"])
me     = pd.read_csv(DATA / "data/system_snapshots/inventory_month_end.csv", parse_dates=["month_end"])
yard_grades = pd.read_csv(DATA / "data/reference/yard_grades.csv")

# ── filter to reference grades ─────────────────────────────────────────
valid_grades = set(grades["grade_id"])
txn = txn[txn["grade_id"].isin(valid_grades)].copy()
me  = me[me["grade_id"].isin(valid_grades)].copy()

# ── reference maps ─────────────────────────────────────────────────────
grade_map   = grades.set_index("grade_id")["metal_group"].to_dict()
grade_name  = grades.set_index("grade_id")["grade_name"].to_dict()
yard_name   = yards.set_index("yard_id")["yard_name"].to_dict()
yard_cap    = yards.set_index("yard_id")["capacity_tons"].to_dict()

# ── time setup ─────────────────────────────────────────────────────────
txn["month"] = txn["date"].dt.to_period("M")
latest_month = me["month_end"].max()
n_months = txn["month"].nunique()
print(f"Data: {n_months} months, latest snapshot: {latest_month.strftime('%Y-%m-%d')}")

# ══════════════════════════════════════════════════════════════════════
# ABC-XYZ CLASSIFICATION (re-derive for positioning context)
# ══════════════════════════════════════════════════════════════════════
ships = txn[txn["txn_type"] == "SHIPMENT"].copy()
ships["revenue"] = ships["tons"] * ships["unit_price"]

# ABC
abc = (ships.groupby("grade_id")
    .agg(total_revenue=("revenue", "sum"))
    .reset_index()
    .sort_values("total_revenue", ascending=False))
total_rev = abc["total_revenue"].sum()
abc["cum_pct"] = abc["total_revenue"].cumsum() / total_rev
abc["abc"] = abc["cum_pct"].apply(lambda x: "A" if x <= 0.80 else ("B" if x <= 0.95 else "C"))

# XYZ
all_months_list = sorted(txn["month"].unique())
monthly_ships = ships.groupby(["grade_id", "month"])["tons"].sum().reset_index()
full_idx = pd.MultiIndex.from_product([sorted(valid_grades), all_months_list], names=["grade_id", "month"])
monthly_full = monthly_ships.set_index(["grade_id", "month"]).reindex(full_idx, fill_value=0).reset_index()

xyz = (monthly_full.groupby("grade_id")["tons"]
    .agg(["mean", "std"]).reset_index()
    .rename(columns={"mean": "avg_m", "std": "std_m"}))
xyz["cov"] = (xyz["std_m"] / xyz["avg_m"]).replace([np.inf, -np.inf], 0).fillna(0)
xyz["xyz"] = xyz["cov"].apply(lambda c: "X" if c <= 0.5 else ("Y" if c <= 1.0 else "Z"))

abc_xyz = abc[["grade_id", "abc"]].merge(xyz[["grade_id", "xyz", "cov", "avg_m"]], on="grade_id")
abc_xyz["abc_xyz"] = abc_xyz["abc"] + abc_xyz["xyz"]
abc_xyz_map = abc_xyz.set_index("grade_id")["abc_xyz"].to_dict()
abc_map_only = abc_xyz.set_index("grade_id")["abc"].to_dict()

# ══════════════════════════════════════════════════════════════════════
# DEMAND ANALYSIS — monthly shipments by yard-grade
# ══════════════════════════════════════════════════════════════════════
# Recent 6-month average demand as the demand signal
recent_months = sorted(all_months_list)[-6:]
recent_ships = ships[ships["month"].isin(recent_months)]

demand_yg = (recent_ships.groupby(["yard_id", "grade_id"])["tons"]
    .agg(["sum", "mean", "std", "count"])
    .reset_index()
    .rename(columns={"sum": "tons_6m", "mean": "avg_monthly", "std": "std_monthly", "count": "shipment_count"}))
demand_yg["std_monthly"] = demand_yg["std_monthly"].fillna(0)

# Full 24-month demand for trend context
demand_24m = (ships.groupby(["yard_id", "grade_id"])["tons"]
    .agg(["sum", "mean"]).reset_index()
    .rename(columns={"sum": "tons_24m", "mean": "avg_monthly_24m"}))

demand_yg = demand_yg.merge(demand_24m, on=["yard_id", "grade_id"], how="outer")
demand_yg = demand_yg.fillna(0)

# ══════════════════════════════════════════════════════════════════════
# POSITIONING — target stock vs actual, gap analysis
# ══════════════════════════════════════════════════════════════════════
# Latest inventory snapshot
inv = me[me["month_end"] == latest_month].copy()

# Target stock = safety factor * avg monthly demand
# Safety factor varies by ABC-XYZ class
safety_factors = {
    "AX": 1.5,  # high value, stable: lean JIT
    "AY": 2.0,  # high value, variable: moderate buffer
    "AZ": 2.5,  # high value, erratic: strategic buffer
    "BX": 1.2,  # moderate, stable: automated reorder
    "BY": 1.8,  # moderate, variable: flexible
    "BZ": 2.0,  # moderate, erratic: minimize
    "CX": 1.0,  # low value, stable: min-max
    "CY": 1.5,  # low value, variable: consolidate
    "CZ": 0.5,  # low value, erratic: exit
}

# Build positioning table
pos = inv[["yard_id", "grade_id", "tons_on_hand", "book_value", "market_value",
           "tons_0_30", "tons_31_60", "tons_61_90", "tons_over_90"]].copy()

pos = pos.merge(demand_yg[["yard_id", "grade_id", "avg_monthly", "std_monthly", "tons_6m"]],
                on=["yard_id", "grade_id"], how="left")
pos = pos.fillna(0)

pos["grade_name"] = pos["grade_id"].map(grade_name)
pos["metal_group"] = pos["grade_id"].map(grade_map)
pos["yard_name"] = pos["yard_id"].map(yard_name)
pos["abc_xyz"] = pos["grade_id"].map(abc_xyz_map)
pos["abc"] = pos["grade_id"].map(abc_map_only)

# Target tons = safety_factor * avg_monthly_demand
pos["safety_factor"] = pos["abc_xyz"].map(safety_factors).fillna(1.5)
pos["target_tons"] = pos["safety_factor"] * pos["avg_monthly"]

# Gap analysis
pos["gap_tons"] = pos["tons_on_hand"] - pos["target_tons"]
pos["gap_pct"] = np.where(pos["target_tons"] > 0,
                          pos["gap_tons"] / pos["target_tons"], 0)

# Position status
def position_status(row):
    if row["target_tons"] == 0 and row["tons_on_hand"] > 0:
        return "EXCESS — No demand"
    if row["target_tons"] == 0:
        return "INACTIVE"
    ratio = row["tons_on_hand"] / row["target_tons"]
    if ratio < 0.5:
        return "CRITICALLY LOW"
    elif ratio < 0.8:
        return "UNDERSTOCKED"
    elif ratio <= 1.3:
        return "ON TARGET"
    elif ratio <= 2.0:
        return "OVERSTOCKED"
    else:
        return "EXCESS"

pos["status"] = pos.apply(position_status, axis=1)

# Days of supply
pos["days_of_supply"] = np.where(pos["avg_monthly"] > 0,
                                  (pos["tons_on_hand"] / pos["avg_monthly"]) * 30, 999)

# Sort by priority: A grades first, then by gap
pos = pos.sort_values(["abc", "status", "gap_tons"], ascending=[True, True, True])

print(f"\nPositioning summary:")
for s in ["CRITICALLY LOW", "UNDERSTOCKED", "ON TARGET", "OVERSTOCKED", "EXCESS", "EXCESS — No demand", "INACTIVE"]:
    cnt = sum(pos["status"] == s)
    if cnt > 0:
        print(f"  {s}: {cnt} positions")

# ══════════════════════════════════════════════════════════════════════
# REBALANCE PLAN — identify inter-yard transfer opportunities
# ══════════════════════════════════════════════════════════════════════
rebalance = []

for gid in sorted(valid_grades):
    grade_pos = pos[pos["grade_id"] == gid].copy()
    if len(grade_pos) < 2:
        continue

    over = grade_pos[grade_pos["status"].isin(["OVERSTOCKED", "EXCESS", "EXCESS — No demand"])]
    under = grade_pos[grade_pos["status"].isin(["CRITICALLY LOW", "UNDERSTOCKED"])]

    if len(over) == 0 or len(under) == 0:
        continue

    for _, o_row in over.iterrows():
        available = o_row["tons_on_hand"] - o_row["target_tons"]
        if available <= 0:
            continue

        for _, u_row in under.iterrows():
            needed = u_row["target_tons"] - u_row["tons_on_hand"]
            if needed <= 0:
                continue

            transfer = min(available, needed)
            if transfer < 1:  # skip trivial transfers
                continue

            abc_class = abc_map_only.get(gid, "C")
            priority = {"A": "HIGH", "B": "MEDIUM", "C": "LOW"}.get(abc_class, "LOW")

            rebalance.append({
                "grade_id": gid,
                "grade_name": grade_name.get(gid, gid),
                "metal_group": grade_map.get(gid, ""),
                "abc_xyz": abc_xyz_map.get(gid, ""),
                "from_yard": o_row["yard_id"],
                "from_yard_name": o_row["yard_name"],
                "to_yard": u_row["yard_id"],
                "to_yard_name": u_row["yard_name"],
                "transfer_tons": round(transfer, 1),
                "from_current": round(o_row["tons_on_hand"], 1),
                "from_target": round(o_row["target_tons"], 1),
                "to_current": round(u_row["tons_on_hand"], 1),
                "to_target": round(u_row["target_tons"], 1),
                "priority": priority,
            })
            available -= transfer

rebal_df = pd.DataFrame(rebalance)
if len(rebal_df) > 0:
    priority_order = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}
    rebal_df["_sort"] = rebal_df["priority"].map(priority_order)
    rebal_df = rebal_df.sort_values(["_sort", "transfer_tons"], ascending=[True, False]).drop(columns="_sort")

print(f"\nRebalance plan: {len(rebal_df)} transfers recommended")
if len(rebal_df) > 0:
    print(f"  Total tons to move: {rebal_df['transfer_tons'].sum():,.0f}")
    print(f"  HIGH priority: {sum(rebal_df['priority']=='HIGH')}")
    print(f"  MEDIUM priority: {sum(rebal_df['priority']=='MEDIUM')}")

# ══════════════════════════════════════════════════════════════════════
# DEMAND ANALYSIS — monthly trends by grade (network-wide)
# ══════════════════════════════════════════════════════════════════════
demand_monthly = (ships.groupby(["grade_id", "month"])["tons"]
    .sum().reset_index())

# Compute trend: compare last 6 months vs first 6 months
first_6 = sorted(all_months_list)[:6]
last_6 = sorted(all_months_list)[-6:]

demand_trend = []
for gid in sorted(valid_grades):
    gd = demand_monthly[demand_monthly["grade_id"] == gid]
    early = gd[gd["month"].isin(first_6)]["tons"].mean()
    late = gd[gd["month"].isin(last_6)]["tons"].mean()
    overall_avg = gd["tons"].mean()
    overall_std = gd["tons"].std()

    trend_pct = (late - early) / early if early > 0 else 0
    if trend_pct > 0.10:
        trend_dir = "GROWING"
    elif trend_pct < -0.10:
        trend_dir = "DECLINING"
    else:
        trend_dir = "STABLE"

    # Month with highest demand
    if len(gd) > 0:
        peak_month = gd.loc[gd["tons"].idxmax(), "month"]
        peak_tons = gd["tons"].max()
    else:
        peak_month = ""
        peak_tons = 0

    demand_trend.append({
        "grade_id": gid,
        "grade_name": grade_name.get(gid, gid),
        "metal_group": grade_map.get(gid, ""),
        "abc_xyz": abc_xyz_map.get(gid, ""),
        "avg_monthly_tons": overall_avg if not pd.isna(overall_avg) else 0,
        "std_monthly": overall_std if not pd.isna(overall_std) else 0,
        "early_avg": early if not pd.isna(early) else 0,
        "late_avg": late if not pd.isna(late) else 0,
        "trend_pct": trend_pct,
        "trend_direction": trend_dir,
        "peak_month": str(peak_month),
        "peak_tons": peak_tons,
    })

demand_df = pd.DataFrame(demand_trend)
demand_df = demand_df.sort_values("avg_monthly_tons", ascending=False)

print(f"\nDemand trends:")
for d in ["GROWING", "STABLE", "DECLINING"]:
    cnt = sum(demand_df["trend_direction"] == d)
    print(f"  {d}: {cnt} grades")

# ══════════════════════════════════════════════════════════════════════
# YARD CAPACITY ANALYSIS
# ══════════════════════════════════════════════════════════════════════
yard_inv = inv.groupby("yard_id").agg(
    tons_on_hand=("tons_on_hand", "sum"),
    book_value=("book_value", "sum"),
    market_value=("market_value", "sum"),
    grades_stocked=("grade_id", "nunique"),
    tons_over_90=("tons_over_90", "sum"),
).reset_index()

yard_inv["yard_name"] = yard_inv["yard_id"].map(yard_name)
yard_inv["capacity"] = yard_inv["yard_id"].map(yard_cap)
yard_inv["utilization"] = yard_inv["tons_on_hand"] / yard_inv["capacity"]
yard_inv["headroom_tons"] = yard_inv["capacity"] - yard_inv["tons_on_hand"]
yard_inv["aged_pct"] = yard_inv["tons_over_90"] / yard_inv["tons_on_hand"]

# Yard shipment throughput
yard_ships = (ships.groupby("yard_id")["tons"]
    .agg(["sum", "mean"]).reset_index()
    .rename(columns={"sum": "total_shipped_24m", "mean": "avg_monthly_shipped"}))
yard_inv = yard_inv.merge(yard_ships, on="yard_id", how="left")
yard_inv["turnover"] = (yard_inv["total_shipped_24m"] / 24) / yard_inv["tons_on_hand"]
yard_inv["turnover"] = yard_inv["turnover"].replace([np.inf, -np.inf], 0).fillna(0)

# Receipts
recpts = txn[txn["txn_type"] == "RECEIPT"]
yard_recpts = (recpts.groupby("yard_id")["tons"]
    .agg(["sum", "mean"]).reset_index()
    .rename(columns={"sum": "total_received_24m", "mean": "avg_monthly_received"}))
yard_inv = yard_inv.merge(yard_recpts, on="yard_id", how="left")

# Net flow
yard_inv["net_monthly_flow"] = yard_inv["avg_monthly_received"].fillna(0) - yard_inv["avg_monthly_shipped"].fillna(0)

# Status
def yard_status(row):
    if row["utilization"] > 0.90:
        return "AT CAPACITY"
    elif row["utilization"] > 0.75:
        return "HIGH UTILIZATION"
    elif row["utilization"] > 0.50:
        return "BALANCED"
    else:
        return "UNDERUTILIZED"

yard_inv["status"] = yard_inv.apply(yard_status, axis=1)
yard_inv = yard_inv.sort_values("utilization", ascending=False)

print(f"\nYard capacity:")
for _, y in yard_inv.iterrows():
    print(f"  {y['yard_id']} ({y['yard_name']}): {y['utilization']:.0%} utilized, {y['headroom_tons']:,.0f}t headroom, {y['status']}")


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
FONT_RED    = Font(name="Arial", size=10, bold=True, color="C00000")
FONT_GREEN  = Font(name="Arial", size=10, bold=True, color="006100")
FONT_ORANGE = Font(name="Arial", size=10, bold=True, color="BF6900")

FILL_HEADER = PatternFill("solid", fgColor="1F3864")
FILL_LIGHT  = PatternFill("solid", fgColor="F2F7FB")
FILL_GREEN  = PatternFill("solid", fgColor="C6EFCE")
FILL_YELLOW = PatternFill("solid", fgColor="FFEB9C")
FILL_RED    = PatternFill("solid", fgColor="FFC7CE")
FILL_BLUE   = PatternFill("solid", fgColor="D6E4F0")
FILL_KPI    = PatternFill("solid", fgColor="E8EEF7")

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
FMT_DOLLAR = '$#,##0'
FMT_PCT    = '0.0%'


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
            if fmts[i] in (FMT_DOLLAR, FMT_NUM, FMT_DEC1, FMT_DEC2, FMT_PCT):
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


status_fill = {
    "CRITICALLY LOW": FILL_RED,
    "UNDERSTOCKED": FILL_YELLOW,
    "ON TARGET": FILL_GREEN,
    "OVERSTOCKED": FILL_BLUE,
    "EXCESS": FILL_RED,
    "EXCESS — No demand": FILL_RED,
    "INACTIVE": FILL_LIGHT,
}

status_font = {
    "CRITICALLY LOW": FONT_RED,
    "UNDERSTOCKED": FONT_ORANGE,
    "ON TARGET": FONT_GREEN,
    "OVERSTOCKED": FONT_BODY_B,
    "EXCESS": FONT_RED,
    "EXCESS — No demand": FONT_RED,
    "INACTIVE": FONT_BODY,
}

# ═══════════════════════════════════════════════════════════════════════
# SHEET 1: README
# ═══════════════════════════════════════════════════════════════════════
ws_r = wb.active
ws_r.title = "README"
ws_r.sheet_properties.tabColor = "808080"

# KPI cards at top
kpi_data = [
    ("Positions Analyzed", f"{len(pos):,}"),
    ("Critically Low", f"{sum(pos['status']=='CRITICALLY LOW')}"),
    ("On Target", f"{sum(pos['status']=='ON TARGET')}"),
    ("Overstocked/Excess", f"{sum(pos['status'].isin(['OVERSTOCKED','EXCESS','EXCESS — No demand']))}"),
    ("Transfers Recommended", f"{len(rebal_df)}"),
    ("Tons to Rebalance", f"{rebal_df['transfer_tons'].sum():,.0f}" if len(rebal_df) > 0 else "0"),
]
for ci, (label, val) in enumerate(kpi_data):
    col = 1 + ci * 2
    c1 = ws_r.cell(row=1, column=col, value=val)
    c1.font = FONT_KPI_V
    c1.fill = FILL_KPI
    c1.alignment = ALIGN_C
    c2 = ws_r.cell(row=2, column=col, value=label)
    c2.font = FONT_KPI_L
    c2.fill = FILL_KPI
    c2.alignment = ALIGN_C
    # Merge across 2 columns for spacing
    ws_r.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
    ws_r.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 1)

readme = [
    ("", None),
    ("INVENTORY POSITIONING & MOVEMENT PLAN", FONT_TITLE),
    ("", None),
    ("PURPOSE", FONT_SECT),
    ("Translates KPI metrics and ABC-XYZ classifications into actionable inventory positioning decisions: where to stock each grade, how much safety stock to hold, which yards need rebalancing, and where to direct purchasing and sales effort. Designed for a Commercial Planning & Inventory Analyst managing a multi-yard scrap-metal network.", FONT_BODY),
    ("", None),
    ("METHODOLOGY", FONT_SECT),
    ("Target stock levels are calculated as: Safety Factor x Average Monthly Demand (last 6 months).", FONT_BODY),
    ("Safety factors vary by ABC-XYZ class: AX=1.5 months (lean JIT), AZ=2.5 months (strategic buffer), CZ=0.5 months (exit candidate).", FONT_BODY),
    ("Rebalance transfers move excess inventory from overstocked yards to understocked yards of the same grade.", FONT_BODY),
    ("Demand trends compare the first 6 months vs the last 6 months of the 24-month window.", FONT_BODY),
    ("", None),
    ("SHEET GUIDE", FONT_SECT),
    ("README — this page with KPI summary", FONT_BODY),
    ("Positioning_Summary — target vs actual stock by yard-grade with gap analysis and status flags", FONT_BODY),
    ("Rebalance_Plan — recommended inter-yard transfers prioritized by grade value class", FONT_BODY),
    ("Demand_Analysis — monthly demand trends, seasonality, and growth direction per grade", FONT_BODY),
    ("Yard_Capacity — capacity utilization, throughput, and headroom per yard", FONT_BODY),
    ("Grade_Strategy — grade-level positioning recommendations driven by ABC-XYZ and demand trends", FONT_BODY),
    ("", None),
    ("DATA NOTE", FONT_SECT),
    ("All data is synthetic, simulating a scrap-metal yard network (8 yards, 40 grades, 24 months Oct 2024 – Sep 2026).", FONT_BODY),
]
for i, (text, font) in enumerate(readme):
    c = ws_r.cell(row=4 + i, column=1, value=text)
    if font:
        c.font = font
ws_r.column_dimensions["A"].width = 110

# ═══════════════════════════════════════════════════════════════════════
# SHEET 2: Positioning_Summary
# ═══════════════════════════════════════════════════════════════════════
ws_ps = wb.create_sheet("Positioning_Summary")
ws_ps.sheet_properties.tabColor = "2E75B6"

ws_ps.cell(row=1, column=1, value="Inventory Positioning — Target vs Actual").font = FONT_TITLE
ws_ps.cell(row=2, column=1, value=f"As of {latest_month.strftime('%B %Y')}  |  Safety stock targets based on 6-month avg demand x ABC-XYZ factor").font = FONT_KPI_L
ws_ps.merge_cells("A1:P1")

headers_ps = ["Yard", "Grade ID", "Grade Name", "Metal Group", "ABC-XYZ",
              "Tons on Hand", "Target Tons", "Gap (tons)", "Gap (%)",
              "Status", "Days of Supply", "Safety Factor",
              "Avg Monthly Demand", "Book Value ($)", "Market Value ($)",
              "Tons 90+ Days"]
write_header(ws_ps, 4, headers_ps)

fmts_ps = [None, None, None, None, None, FMT_DEC1, FMT_DEC1, FMT_DEC1,
           FMT_PCT, None, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DOLLAR, FMT_DOLLAR, FMT_DEC1]

for i, (_, row) in enumerate(pos.iterrows()):
    vals = [row["yard_id"], row["grade_id"], row["grade_name"], row["metal_group"],
            row["abc_xyz"], row["tons_on_hand"], row["target_tons"],
            row["gap_tons"], row["gap_pct"], row["status"],
            min(row["days_of_supply"], 999), row["safety_factor"],
            row["avg_monthly"], row["book_value"], row["market_value"],
            row["tons_over_90"]]
    write_row(ws_ps, 5 + i, vals, fmts_ps, alt=i % 2 == 0)

    # Color status
    st_cell = ws_ps.cell(row=5 + i, column=10)
    st_cell.fill = status_fill.get(row["status"], FILL_LIGHT)
    st_cell.font = status_font.get(row["status"], FONT_BODY)
    st_cell.alignment = ALIGN_C

# Summary row
sr = 5 + len(pos) + 1
ws_ps.cell(row=sr, column=1, value="TOTAL").font = FONT_BODY_B
# Add SUM formulas
last_data_row = 4 + len(pos)
for col_idx, col_letter in [(6, "F"), (7, "G"), (8, "H"), (14, "N"), (15, "O"), (16, "P")]:
    c = ws_ps.cell(row=sr, column=col_idx, value=f"=SUM({col_letter}5:{col_letter}{last_data_row})")
    c.font = FONT_BODY_B
    c.border = BORDER
    if col_idx in (14, 15):
        c.number_format = FMT_DOLLAR
    else:
        c.number_format = FMT_DEC1

auto_width(ws_ps)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 3: Rebalance_Plan
# ═══════════════════════════════════════════════════════════════════════
ws_rb = wb.create_sheet("Rebalance_Plan")
ws_rb.sheet_properties.tabColor = "C00000"

ws_rb.cell(row=1, column=1, value="Inter-Yard Rebalance Plan").font = FONT_TITLE
total_transfer = rebal_df["transfer_tons"].sum() if len(rebal_df) > 0 else 0
ws_rb.cell(row=2, column=1, value=f"{len(rebal_df)} transfers recommended  |  {total_transfer:,.0f} tons to move").font = FONT_KPI_L
ws_rb.merge_cells("A1:M1")

headers_rb = ["Priority", "Grade ID", "Grade Name", "Metal Group", "ABC-XYZ",
              "From Yard", "To Yard", "Transfer Tons",
              "From Current", "From Target", "To Current", "To Target"]
write_header(ws_rb, 4, headers_rb)

fmts_rb = [None, None, None, None, None, None, None, FMT_DEC1,
           FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC1]

priority_fill = {"HIGH": FILL_RED, "MEDIUM": FILL_YELLOW, "LOW": FILL_BLUE}

if len(rebal_df) > 0:
    for i, (_, row) in enumerate(rebal_df.iterrows()):
        vals = [row["priority"], row["grade_id"], row["grade_name"],
                row["metal_group"], row["abc_xyz"],
                row["from_yard"], row["to_yard"], row["transfer_tons"],
                row["from_current"], row["from_target"],
                row["to_current"], row["to_target"]]
        write_row(ws_rb, 5 + i, vals, fmts_rb, alt=i % 2 == 0)

        # Color priority
        p_cell = ws_rb.cell(row=5 + i, column=1)
        p_cell.fill = priority_fill.get(row["priority"], FILL_LIGHT)
        p_cell.font = FONT_BODY_B
        p_cell.alignment = ALIGN_C

    # Totals
    sr = 5 + len(rebal_df) + 1
    ws_rb.cell(row=sr, column=7, value="TOTAL").font = FONT_BODY_B
    last_rb = 4 + len(rebal_df)
    c = ws_rb.cell(row=sr, column=8, value=f"=SUM(H5:H{last_rb})")
    c.font = FONT_BODY_B
    c.number_format = FMT_DEC1
    c.border = BORDER
else:
    ws_rb.cell(row=5, column=1, value="No rebalance transfers needed — all positions are within target ranges.").font = FONT_BODY

auto_width(ws_rb)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 4: Demand_Analysis
# ═══════════════════════════════════════════════════════════════════════
ws_da = wb.create_sheet("Demand_Analysis")
ws_da.sheet_properties.tabColor = "548235"

ws_da.cell(row=1, column=1, value="Demand Analysis by Grade").font = FONT_TITLE
ws_da.cell(row=2, column=1, value="24-month demand patterns — first 6 months vs last 6 months trend comparison").font = FONT_KPI_L
ws_da.merge_cells("A1:L1")

headers_da = ["Grade ID", "Grade Name", "Metal Group", "ABC-XYZ",
              "Avg Monthly Tons", "Std Dev", "Early 6m Avg", "Late 6m Avg",
              "Trend (%)", "Direction", "Peak Month", "Peak Tons"]
write_header(ws_da, 4, headers_da)

fmts_da = [None, None, None, None, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC1,
           FMT_PCT, None, None, FMT_DEC1]

trend_fill = {"GROWING": FILL_GREEN, "STABLE": FILL_BLUE, "DECLINING": FILL_RED}
trend_font = {"GROWING": FONT_GREEN, "STABLE": FONT_BODY_B, "DECLINING": FONT_RED}

for i, (_, row) in enumerate(demand_df.iterrows()):
    vals = [row["grade_id"], row["grade_name"], row["metal_group"],
            row["abc_xyz"], row["avg_monthly_tons"], row["std_monthly"],
            row["early_avg"], row["late_avg"], row["trend_pct"],
            row["trend_direction"], row["peak_month"], row["peak_tons"]]
    write_row(ws_da, 5 + i, vals, fmts_da, alt=i % 2 == 0)

    # Color direction
    d_cell = ws_da.cell(row=5 + i, column=10)
    d_cell.fill = trend_fill.get(row["trend_direction"], FILL_LIGHT)
    d_cell.font = trend_font.get(row["trend_direction"], FONT_BODY)
    d_cell.alignment = ALIGN_C

auto_width(ws_da)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 5: Yard_Capacity
# ═══════════════════════════════════════════════════════════════════════
ws_yc = wb.create_sheet("Yard_Capacity")
ws_yc.sheet_properties.tabColor = "7030A0"

ws_yc.cell(row=1, column=1, value="Yard Capacity & Throughput Analysis").font = FONT_TITLE
ws_yc.cell(row=2, column=1, value=f"As of {latest_month.strftime('%B %Y')}  |  8 yards across the network").font = FONT_KPI_L
ws_yc.merge_cells("A1:N1")

headers_yc = ["Yard ID", "Yard Name", "Capacity (tons)", "Tons on Hand",
              "Utilization (%)", "Headroom (tons)", "Status",
              "Grades Stocked", "Avg Monthly Shipped", "Avg Monthly Received",
              "Net Monthly Flow", "Turnover (x/mo)", "Tons 90+ Days", "Aged (%)"]
write_header(ws_yc, 4, headers_yc)

fmts_yc = [None, None, FMT_NUM, FMT_DEC1, FMT_PCT, FMT_DEC1, None,
           FMT_NUM, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC2, FMT_DEC1, FMT_PCT]

yard_status_fill = {
    "AT CAPACITY": FILL_RED,
    "HIGH UTILIZATION": FILL_YELLOW,
    "BALANCED": FILL_GREEN,
    "UNDERUTILIZED": FILL_BLUE,
}

for i, (_, row) in enumerate(yard_inv.iterrows()):
    vals = [row["yard_id"], row["yard_name"], row["capacity"],
            row["tons_on_hand"], row["utilization"], row["headroom_tons"],
            row["status"], row["grades_stocked"],
            row.get("avg_monthly_shipped", 0), row.get("avg_monthly_received", 0),
            row.get("net_monthly_flow", 0), row["turnover"],
            row["tons_over_90"], row["aged_pct"]]
    write_row(ws_yc, 5 + i, vals, fmts_yc, alt=i % 2 == 0)

    st_cell = ws_yc.cell(row=5 + i, column=7)
    st_cell.fill = yard_status_fill.get(row["status"], FILL_LIGHT)
    st_cell.font = FONT_BODY_B
    st_cell.alignment = ALIGN_C

# Totals
sr = 5 + len(yard_inv) + 1
ws_yc.cell(row=sr, column=1, value="TOTAL").font = FONT_BODY_B
last_yc = 4 + len(yard_inv)
for col_idx, col_letter in [(3, "C"), (4, "D"), (6, "F"), (13, "M")]:
    c = ws_yc.cell(row=sr, column=col_idx, value=f"=SUM({col_letter}5:{col_letter}{last_yc})")
    c.font = FONT_BODY_B
    c.border = BORDER
    c.number_format = FMT_NUM if col_idx == 3 else FMT_DEC1

# Network avg utilization
c = ws_yc.cell(row=sr, column=5, value=f"=D{sr}/C{sr}")
c.font = FONT_BODY_B
c.number_format = FMT_PCT
c.border = BORDER

auto_width(ws_yc)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 6: Grade_Strategy
# ═══════════════════════════════════════════════════════════════════════
ws_gs = wb.create_sheet("Grade_Strategy")
ws_gs.sheet_properties.tabColor = "BF8F00"

ws_gs.cell(row=1, column=1, value="Grade-Level Positioning Strategy").font = FONT_TITLE
ws_gs.cell(row=2, column=1, value="Combines ABC-XYZ classification, demand trend, and current position into actionable recommendations").font = FONT_KPI_L
ws_gs.merge_cells("A1:K1")

# Combine data for strategy
grade_strat = abc_xyz[["grade_id", "abc", "xyz", "abc_xyz", "avg_m"]].copy()
grade_strat["grade_name"] = grade_strat["grade_id"].map(grade_name)
grade_strat["metal_group"] = grade_strat["grade_id"].map(grade_map)

# Current network inventory
inv_net = inv.groupby("grade_id").agg(
    tons_on_hand=("tons_on_hand", "sum"),
    market_value=("market_value", "sum"),
    tons_over_90=("tons_over_90", "sum"),
).reset_index()
grade_strat = grade_strat.merge(inv_net, on="grade_id", how="left").fillna(0)

# Demand trend
grade_strat = grade_strat.merge(demand_df[["grade_id", "trend_direction", "trend_pct"]], on="grade_id", how="left")

# Positioning status count
pos_status = pos.groupby("grade_id")["status"].value_counts().unstack(fill_value=0).reset_index()
grade_strat = grade_strat.merge(pos_status[["grade_id"]], on="grade_id", how="left")

# Generate recommendation
def make_recommendation(row):
    abc = row["abc"]
    xyz = row["xyz"]
    trend = row.get("trend_direction", "STABLE")
    aged = row["tons_over_90"]
    on_hand = row["tons_on_hand"]
    aged_pct = aged / on_hand if on_hand > 0 else 0

    parts = []

    # Stock strategy by ABC-XYZ
    strat = {
        "AX": "STOCK AGGRESSIVELY — negotiate volume contracts, maintain lean safety stock",
        "AY": "STOCK WITH BUFFER — forecast-driven purchasing, weekly review cycle",
        "AZ": "STRATEGIC HOLD — diversify suppliers, hold 2.5x buffer, price-triggered buying",
        "BX": "AUTOMATED REORDER — set min/max levels, quarterly review",
        "BY": "ACTIVE MANAGEMENT — monthly forecast updates, flexible terms",
        "BZ": "MINIMIZE EXPOSURE — buy on demand, spot-market sourcing",
        "CX": "LOW-TOUCH — min/max with bulk discounts, semi-annual review",
        "CY": "CONSOLIDATE — reduce supplier count, consider dropping",
        "CZ": "EXIT CANDIDATE — phase out, make-to-order only",
    }
    parts.append(strat.get(row["abc_xyz"], "REVIEW"))

    # Trend overlay
    if trend == "GROWING" and abc in ("A", "B"):
        parts.append("Demand growing — consider increasing targets")
    elif trend == "DECLINING" and abc == "C":
        parts.append("Demand declining — accelerate exit")
    elif trend == "DECLINING" and abc == "A":
        parts.append("Demand declining — monitor closely, adjust targets down")

    # Aging flag
    if aged_pct > 0.20:
        parts.append(f"WARNING: {aged_pct:.0%} of stock >90 days — prioritize movement")

    return " | ".join(parts)

grade_strat["recommendation"] = grade_strat.apply(make_recommendation, axis=1)
grade_strat = grade_strat.sort_values(["abc", "xyz"])

headers_gs = ["Grade ID", "Grade Name", "Metal Group", "ABC", "XYZ", "ABC-XYZ",
              "Avg Monthly Demand", "Tons on Hand", "Market Value ($)",
              "Demand Trend", "Recommendation"]
write_header(ws_gs, 4, headers_gs)

fmts_gs = [None, None, None, None, None, None, FMT_DEC1, FMT_DEC1, FMT_DOLLAR, None, None]

abc_fill_map = {"A": FILL_GREEN, "B": FILL_YELLOW, "C": FILL_RED}
xyz_fill_map = {"X": FILL_GREEN, "Y": FILL_YELLOW, "Z": FILL_RED}

for i, (_, row) in enumerate(grade_strat.iterrows()):
    vals = [row["grade_id"], row["grade_name"], row["metal_group"],
            row["abc"], row["xyz"], row["abc_xyz"],
            row["avg_m"], row["tons_on_hand"], row["market_value"],
            row.get("trend_direction", ""), row["recommendation"]]
    write_row(ws_gs, 5 + i, vals, fmts_gs, alt=i % 2 == 0)

    # Color ABC and XYZ
    ws_gs.cell(row=5 + i, column=4).fill = abc_fill_map.get(row["abc"], FILL_LIGHT)
    ws_gs.cell(row=5 + i, column=4).alignment = ALIGN_C
    ws_gs.cell(row=5 + i, column=5).fill = xyz_fill_map.get(row["xyz"], FILL_LIGHT)
    ws_gs.cell(row=5 + i, column=5).alignment = ALIGN_C

    # Color trend
    t_cell = ws_gs.cell(row=5 + i, column=10)
    t_cell.fill = trend_fill.get(row.get("trend_direction", ""), FILL_LIGHT)
    t_cell.font = trend_font.get(row.get("trend_direction", ""), FONT_BODY)
    t_cell.alignment = ALIGN_C

auto_width(ws_gs)
ws_gs.column_dimensions["K"].width = 80  # Recommendation column

# ── hide gridlines ─────────────────────────────────────────────────────
for ws_sheet in wb.worksheets:
    ws_sheet.sheet_view.showGridLines = False

# ── save ───────────────────────────────────────────────────────────────
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print(f"\nWorkbook saved to {OUT}")
print(f"Sheets: {wb.sheetnames}")
