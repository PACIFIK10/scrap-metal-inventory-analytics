"""
Build the Inventory KPI & Aging Analysis workbook for the scrap-metal project.

Reads the generated CSVs and produces a formatted Excel file with:
  - KPI_Summary: network-wide headline metrics (latest month)
  - KPI_By_Yard: per-yard KPIs with sparkline-style trends
  - KPI_By_Metal: per-metal-group KPIs
  - Aging_Detail: lot-level aging with 30/60/90+ buckets by yard and grade
  - Margin_Analysis: buy/sell spread and margin per grade
  - At_Risk_Inventory: flagged slow-moving and excess stock
  - Monthly_Trends: 24-month time series of key KPIs
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter
from pathlib import Path
import sys

# ── paths ──────────────────────────────────────────────────────────────
DATA = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("/mnt/user-data/uploads/emr-inventory-project")
OUT  = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("/home/claude/kpi_workbook/inventory_kpi_aging.xlsx")

# ── load data ──────────────────────────────────────────────────────────
txn   = pd.read_csv(DATA / "data/raw/transactions.csv", parse_dates=["date"])
me    = pd.read_csv(DATA / "data/system_snapshots/inventory_month_end.csv", parse_dates=["month_end"])
lots  = pd.read_csv(DATA / "data/system_snapshots/inventory_open_lots.csv", parse_dates=["receipt_date"])
grades = pd.read_csv(DATA / "data/reference/grades.csv")
yards  = pd.read_csv(DATA / "data/reference/yards.csv")
prices = pd.read_csv(DATA / "data/raw/grade_prices_daily.csv", parse_dates=["date"])

# ── enrich with metal group ────────────────────────────────────────────
grade_map = grades.set_index("grade_id")["metal_group"].to_dict()
grade_name_map = grades.set_index("grade_id")["grade_name"].to_dict()
yard_name_map = yards.set_index("yard_id")["yard_name"].to_dict()
yard_cap_map  = yards.set_index("yard_id")["capacity_tons"].to_dict()

me["metal_group"] = me["grade_id"].map(grade_map)
txn["metal_group"] = txn["grade_id"].map(grade_map)
lots["metal_group"] = lots["grade_id"].map(grade_map)
lots["grade_name"]  = lots["grade_id"].map(grade_name_map)

# ── latest month ───────────────────────────────────────────────────────
latest_month = me["month_end"].max()
prev_month   = (latest_month - pd.DateOffset(months=1)) + pd.offsets.MonthEnd(0)
me_latest    = me[me["month_end"] == latest_month].copy()
me_prev      = me[me["month_end"] == prev_month].copy()

print(f"Latest month: {latest_month.strftime('%Y-%m-%d')}")
print(f"Rows in latest snapshot: {len(me_latest)}")

# ── latest prices (last available day) ─────────────────────────────────
latest_price_date = prices["date"].max()
prices_latest = prices[prices["date"] == latest_price_date].set_index("grade_id")

# ── compute KPIs ───────────────────────────────────────────────────────

# 1. Monthly shipments and receipts by yard
txn_monthly = txn.copy()
txn_monthly["month_end"] = txn_monthly["date"].dt.to_period("M").dt.to_timestamp() + pd.offsets.MonthEnd(0)

ship_monthly = (txn_monthly[txn_monthly["txn_type"] == "SHIPMENT"]
                .groupby(["month_end", "yard_id", "grade_id"])["tons"].sum().reset_index()
                .rename(columns={"tons": "tons_shipped"}))

recv_monthly = (txn_monthly[txn_monthly["txn_type"] == "RECEIPT"]
                .groupby(["month_end", "yard_id", "grade_id"])["tons"].sum().reset_index()
                .rename(columns={"tons": "tons_received"}))

# 2. Merge shipments into month-end for turnover calc (match on yard+grade to avoid fan-out)
me_full = me.merge(ship_monthly, on=["month_end", "yard_id", "grade_id"], how="left")
me_full = me_full.merge(recv_monthly, on=["month_end", "yard_id", "grade_id"], how="left")
me_full["tons_shipped"]  = me_full["tons_shipped"].fillna(0)
me_full["tons_received"] = me_full["tons_received"].fillna(0)

# ── SHEET DATA PREP ────────────────────────────────────────────────────

# === KPI_Summary (network-wide, latest month) ===
total_tons   = me_latest["tons_on_hand"].sum()
total_book   = me_latest["book_value"].sum()
total_market = me_latest["market_value"].sum()
unrealized   = total_market - total_book

# Aging
total_0_30  = me_latest["tons_0_30"].sum()
total_31_60 = me_latest["tons_31_60"].sum()
total_61_90 = me_latest["tons_61_90"].sum()
total_90p   = me_latest["tons_over_90"].sum()

# Turnover (annualized): last 12 months shipments / avg inventory
last_12 = me_full[me_full["month_end"] > latest_month - pd.DateOffset(months=12)]
avg_tons_12m = last_12.groupby("month_end")["tons_on_hand"].sum().mean()
total_shipped_12m = last_12.groupby("month_end")["tons_shipped"].sum().sum()
turnover_annual = total_shipped_12m / avg_tons_12m if avg_tons_12m > 0 else 0
days_of_inv = 365 / turnover_annual if turnover_annual > 0 else 999

# Capacity
total_capacity = yards["capacity_tons"].sum()
utilization = total_tons / total_capacity

# Margin (latest prices, weighted avg across grades)
margin_data = []
for _, row in me_latest.iterrows():
    gid = row["grade_id"]
    if gid in prices_latest.index:
        buy  = prices_latest.loc[gid, "buy_price"]
        sell = prices_latest.loc[gid, "sell_price"]
        margin_data.append({
            "grade_id": gid,
            "tons": row["tons_on_hand"],
            "buy": buy,
            "sell": sell,
            "margin": sell - buy
        })

margin_df = pd.DataFrame(margin_data)
if len(margin_df) > 0:
    avg_margin = (margin_df["margin"] * margin_df["tons"]).sum() / margin_df["tons"].sum()
else:
    avg_margin = 0

# Previous month comparison
prev_total_tons = me_prev["tons_on_hand"].sum() if len(me_prev) > 0 else total_tons
tons_change = total_tons - prev_total_tons
tons_change_pct = tons_change / prev_total_tons if prev_total_tons > 0 else 0

summary_kpis = {
    "Reporting Period": latest_month.strftime("%B %Y"),
    "Total Tons on Hand": total_tons,
    "Month-over-Month Change (tons)": tons_change,
    "Month-over-Month Change (%)": tons_change_pct,
    "Book Value ($)": total_book,
    "Market Value ($)": total_market,
    "Unrealized Gain/Loss ($)": unrealized,
    "Inventory Turnover (annualized)": turnover_annual,
    "Days of Inventory": days_of_inv,
    "Capacity Utilization (%)": utilization,
    "Weighted Avg Margin ($/ton)": avg_margin,
    "Tons Aged 0-30 Days": total_0_30,
    "Tons Aged 31-60 Days": total_31_60,
    "Tons Aged 61-90 Days": total_61_90,
    "Tons Aged 90+ Days": total_90p,
    "Aged 90+ as % of Total": total_90p / total_tons if total_tons > 0 else 0,
}

# === KPI_By_Yard ===
yard_kpis = []
for yid in sorted(me_latest["yard_id"].unique()):
    yd = me_latest[me_latest["yard_id"] == yid]
    yd_last12 = last_12[last_12["yard_id"] == yid]

    tons = yd["tons_on_hand"].sum()
    book = yd["book_value"].sum()
    market = yd["market_value"].sum()
    cap = yard_cap_map.get(yid, 1)

    shipped_12m = yd_last12["tons_shipped"].sum()
    avg_inv_12m = yd_last12.groupby("month_end")["tons_on_hand"].sum().mean()
    turn = shipped_12m / avg_inv_12m if avg_inv_12m > 0 else 0
    doi  = 365 / turn if turn > 0 else 999

    t_0_30  = yd["tons_0_30"].sum()
    t_31_60 = yd["tons_31_60"].sum()
    t_61_90 = yd["tons_61_90"].sum()
    t_90p   = yd["tons_over_90"].sum()

    # Previous month
    yd_prev = me_prev[me_prev["yard_id"] == yid] if len(me_prev) > 0 else pd.DataFrame()
    prev_tons = yd_prev["tons_on_hand"].sum() if len(yd_prev) > 0 else tons
    mom_change = tons - prev_tons

    yard_kpis.append({
        "Yard": yard_name_map.get(yid, yid),
        "Yard ID": yid,
        "Tons on Hand": tons,
        "MoM Change": mom_change,
        "Book Value ($)": book,
        "Market Value ($)": market,
        "Unrealized ($)": market - book,
        "Turnover": turn,
        "Days of Inv": doi,
        "Capacity (tons)": cap,
        "Utilization (%)": tons / cap if cap > 0 else 0,
        "0-30 Days": t_0_30,
        "31-60 Days": t_31_60,
        "61-90 Days": t_61_90,
        "90+ Days": t_90p,
        "Aged 90+ (%)": t_90p / tons if tons > 0 else 0,
    })

yard_kpi_df = pd.DataFrame(yard_kpis)

# === KPI_By_Metal ===
metal_kpis = []
for mg in ["COPPER", "ALUMINUM", "BRASS", "STAINLESS", "FERROUS"]:
    mgd = me_latest[me_latest["metal_group"] == mg]
    mg_last12 = last_12[last_12["grade_id"].map(grade_map) == mg]

    tons = mgd["tons_on_hand"].sum()
    book = mgd["book_value"].sum()
    market = mgd["market_value"].sum()

    shipped_12m = mg_last12["tons_shipped"].sum()
    avg_inv_12m = mg_last12.groupby("month_end")["tons_on_hand"].sum().mean()
    turn = shipped_12m / avg_inv_12m if avg_inv_12m > 0 else 0
    doi  = 365 / turn if turn > 0 else 999

    t_0_30  = mgd["tons_0_30"].sum()
    t_31_60 = mgd["tons_31_60"].sum()
    t_61_90 = mgd["tons_61_90"].sum()
    t_90p   = mgd["tons_over_90"].sum()

    # Margin
    mg_margins = margin_df[margin_df["grade_id"].map(grade_map) == mg] if len(margin_df) > 0 else pd.DataFrame()
    if len(mg_margins) > 0 and mg_margins["tons"].sum() > 0:
        mg_avg_margin = (mg_margins["margin"] * mg_margins["tons"]).sum() / mg_margins["tons"].sum()
    else:
        mg_avg_margin = 0

    metal_kpis.append({
        "Metal Group": mg.title(),
        "Tons on Hand": tons,
        "# Grades": mgd["grade_id"].nunique(),
        "# Yards": mgd["yard_id"].nunique(),
        "Book Value ($)": book,
        "Market Value ($)": market,
        "Unrealized ($)": market - book,
        "Avg Margin ($/ton)": mg_avg_margin,
        "Turnover": turn,
        "Days of Inv": doi,
        "0-30 Days": t_0_30,
        "31-60 Days": t_31_60,
        "61-90 Days": t_61_90,
        "90+ Days": t_90p,
        "Aged 90+ (%)": t_90p / tons if tons > 0 else 0,
    })

metal_kpi_df = pd.DataFrame(metal_kpis)

# === Aging_Detail (lot-level, from open_lots) ===
aging_detail = lots.copy()
aging_detail["yard_name"] = aging_detail["yard_id"].map(yard_name_map)
aging_detail["aging_bucket"] = pd.cut(
    aging_detail["age_days"],
    bins=[-1, 30, 60, 90, 9999],
    labels=["0-30", "31-60", "61-90", "90+"]
)
aging_detail["lot_value"] = aging_detail["tons_remaining"] * aging_detail["cost_per_ton"]
aging_detail = aging_detail.sort_values(["yard_id", "grade_id", "age_days"], ascending=[True, True, False])

# Aging summary by yard + grade (for the aging pivot)
aging_summary = (me_latest.groupby(["yard_id", "grade_id"])
    .agg({
        "tons_on_hand": "sum",
        "book_value": "sum",
        "market_value": "sum",
        "tons_0_30": "sum",
        "tons_31_60": "sum",
        "tons_61_90": "sum",
        "tons_over_90": "sum",
    }).reset_index())
aging_summary["yard_name"]  = aging_summary["yard_id"].map(yard_name_map)
aging_summary["grade_name"] = aging_summary["grade_id"].map(grade_name_map)
aging_summary["metal_group"] = aging_summary["grade_id"].map(grade_map)
aging_summary["aged_pct"] = (aging_summary["tons_over_90"] / aging_summary["tons_on_hand"]).fillna(0)
aging_summary = aging_summary.sort_values(["yard_id", "metal_group", "grade_id"])

# === Margin_Analysis ===
margin_analysis = []
for gid in sorted(grades["grade_id"]):
    gname = grade_name_map[gid]
    mg = grade_map[gid]
    tons_held = me_latest[me_latest["grade_id"] == gid]["tons_on_hand"].sum()

    if gid in prices_latest.index:
        buy  = prices_latest.loc[gid, "buy_price"]
        sell = prices_latest.loc[gid, "sell_price"]
        margin = sell - buy
        margin_pct = margin / buy if buy > 0 else 0
    else:
        buy = sell = margin = margin_pct = 0

    # 30-day price trend
    price_30d = prices[(prices["grade_id"] == gid) &
                       (prices["date"] > latest_price_date - pd.Timedelta(days=30))]
    if len(price_30d) > 1:
        first_buy = price_30d.iloc[0]["buy_price"]
        last_buy  = price_30d.iloc[-1]["buy_price"]
        price_trend = (last_buy - first_buy) / first_buy if first_buy > 0 else 0
    else:
        price_trend = 0

    margin_analysis.append({
        "Grade ID": gid,
        "Grade Name": gname,
        "Metal Group": mg.title(),
        "Tons Held": tons_held,
        "Buy Price ($/ton)": buy,
        "Sell Price ($/ton)": sell,
        "Spread ($/ton)": margin,
        "Spread (%)": margin_pct,
        "30-Day Price Trend": price_trend,
        "Exposure ($)": tons_held * buy,
    })

margin_analysis_df = pd.DataFrame(margin_analysis)

# === At_Risk_Inventory ===
at_risk = []
for _, row in aging_summary.iterrows():
    tons = row["tons_on_hand"]
    if tons == 0:
        continue

    risks = []
    risk_score = 0

    # Flag 1: >20% aged 90+ days
    if row["aged_pct"] > 0.20:
        risks.append("High aging (>20% over 90 days)")
        risk_score += 3

    # Flag 2: any tons over 90 days
    if row["tons_over_90"] > 0:
        risks.append(f"{row['tons_over_90']:.1f} tons over 90 days")
        risk_score += 1

    # Flag 3: tons 61-90 days (pre-aging warning)
    if row["tons_61_90"] > 5:
        risks.append(f"{row['tons_61_90']:.1f} tons at 61-90 days (approaching aged)")
        risk_score += 1

    # Flag 4: market below book (underwater)
    if row["market_value"] < row["book_value"]:
        loss = row["book_value"] - row["market_value"]
        risks.append(f"Market below book by ${loss:,.0f}")
        risk_score += 2

    if risk_score >= 2:
        # Recommendation
        if risk_score >= 4:
            action = "Priority: move to Sales for immediate disposal"
        elif row["market_value"] < row["book_value"]:
            action = "Review: sell or write down to market"
        else:
            action = "Monitor: schedule for next shipment cycle"

        at_risk.append({
            "Yard": row["yard_name"],
            "Yard ID": row["yard_id"],
            "Grade ID": row["grade_id"],
            "Grade Name": row["grade_name"],
            "Metal Group": row["metal_group"].title(),
            "Tons on Hand": tons,
            "Book Value ($)": row["book_value"],
            "Market Value ($)": row["market_value"],
            "Gain/Loss ($)": row["market_value"] - row["book_value"],
            "Tons 61-90d": row["tons_61_90"],
            "Tons 90+ d": row["tons_over_90"],
            "Aged %": row["aged_pct"],
            "Risk Score": risk_score,
            "Risk Flags": "; ".join(risks),
            "Recommended Action": action,
        })

at_risk_df = pd.DataFrame(at_risk).sort_values("Risk Score", ascending=False)
print(f"At-risk items: {len(at_risk_df)}")

# === Monthly_Trends ===
monthly = (me_full.groupby("month_end")
    .agg({
        "tons_on_hand": "sum",
        "book_value": "sum",
        "market_value": "sum",
        "tons_0_30": "sum",
        "tons_31_60": "sum",
        "tons_61_90": "sum",
        "tons_over_90": "sum",
        "tons_shipped": "sum",
        "tons_received": "sum",
    }).reset_index()
    .sort_values("month_end"))

# Add rolling turnover
monthly["avg_tons_3m"] = monthly["tons_on_hand"].rolling(3, min_periods=1).mean()
monthly["shipped_3m"]  = monthly["tons_shipped"].rolling(3, min_periods=1).sum()
monthly["turnover_3m"] = (monthly["shipped_3m"] / monthly["avg_tons_3m"]).replace([np.inf, -np.inf], 0)
monthly["aged_pct"]    = monthly["tons_over_90"] / monthly["tons_on_hand"]
monthly["utilization"]  = monthly["tons_on_hand"] / total_capacity

# ── BUILD WORKBOOK ─────────────────────────────────────────────────────

wb = Workbook()

# ── styles ─────────────────────────────────────────────────────────────
FONT_TITLE  = Font(name="Arial", size=16, bold=True, color="1F3864")
FONT_HEAD   = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FONT_BODY   = Font(name="Arial", size=10)
FONT_BODY_B = Font(name="Arial", size=10, bold=True)
FONT_KPI_VAL = Font(name="Arial", size=14, bold=True, color="1F3864")
FONT_KPI_LBL = Font(name="Arial", size=10, color="595959")
FONT_RED     = Font(name="Arial", size=10, color="C00000")
FONT_GREEN   = Font(name="Arial", size=10, color="006100")
FONT_LINK    = Font(name="Arial", size=10, italic=True, color="808080")

FILL_HEADER  = PatternFill("solid", fgColor="1F3864")
FILL_ALT     = PatternFill("solid", fgColor="D6E4F0")
FILL_LIGHT   = PatternFill("solid", fgColor="F2F7FB")
FILL_KPI_BOX = PatternFill("solid", fgColor="E8EEF7")
FILL_RED_LT  = PatternFill("solid", fgColor="FDE8E8")
FILL_YEL_LT  = PatternFill("solid", fgColor="FFF3CD")
FILL_GRN_LT  = PatternFill("solid", fgColor="D4EDDA")

ALIGN_C = Alignment(horizontal="center", vertical="center")
ALIGN_R = Alignment(horizontal="right", vertical="center")
ALIGN_L = Alignment(horizontal="left", vertical="center", wrap_text=True)

BORDER_THIN = Border(
    bottom=Side(style="thin", color="B4C6E7"),
    top=Side(style="thin", color="B4C6E7"),
    left=Side(style="thin", color="B4C6E7"),
    right=Side(style="thin", color="B4C6E7"),
)

FMT_NUM    = '#,##0'
FMT_DEC1   = '#,##0.0'
FMT_DEC2   = '#,##0.00'
FMT_DOLLAR = '$#,##0'
FMT_DOLLAR2 = '$#,##0.00'
FMT_PCT    = '0.0%'
FMT_PCT0   = '0%'


def write_header_row(ws, row, headers, col_start=1):
    """Write a formatted header row."""
    for i, h in enumerate(headers):
        cell = ws.cell(row=row, column=col_start + i, value=h)
        cell.font = FONT_HEAD
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_C
        cell.border = BORDER_THIN


def write_data_row(ws, row, values, formats=None, col_start=1, alt=False):
    """Write a data row with optional alternating fill and number formats."""
    for i, v in enumerate(values):
        cell = ws.cell(row=row, column=col_start + i, value=v)
        cell.font = FONT_BODY
        cell.border = BORDER_THIN
        if alt:
            cell.fill = FILL_LIGHT
        if formats and i < len(formats) and formats[i]:
            cell.number_format = formats[i]
            if formats[i] in (FMT_DOLLAR, FMT_DOLLAR2, FMT_NUM, FMT_DEC1, FMT_DEC2, FMT_PCT, FMT_PCT0):
                cell.alignment = ALIGN_R
            else:
                cell.alignment = ALIGN_L
        else:
            cell.alignment = ALIGN_L


def auto_width(ws, min_width=10, max_width=30):
    """Auto-fit column widths based on content."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        width = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = width


# ═══════════════════════════════════════════════════════════════════════
# SHEET 1: KPI_Summary
# ═══════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "KPI_Summary"
ws.sheet_properties.tabColor = "1F3864"

ws.cell(row=1, column=1, value="Scrap Metal Inventory KPI Summary").font = FONT_TITLE
ws.cell(row=2, column=1, value=f"Reporting Period: {latest_month.strftime('%B %Y')}").font = FONT_KPI_LBL
ws.merge_cells("A1:F1")
ws.merge_cells("A2:F2")

# KPI cards in a 3x2 grid
kpi_cards = [
    ("Total Tons on Hand", total_tons, FMT_NUM),
    ("Book Value", total_book, FMT_DOLLAR),
    ("Market Value", total_market, FMT_DOLLAR),
    ("Inventory Turnover", turnover_annual, FMT_DEC1),
    ("Days of Inventory", days_of_inv, FMT_DEC1),
    ("Capacity Utilization", utilization, FMT_PCT),
    ("Avg Margin ($/ton)", avg_margin, FMT_DOLLAR2),
    ("Unrealized Gain/Loss", unrealized, FMT_DOLLAR),
    ("MoM Change", tons_change, FMT_NUM),
]

r = 4
for idx, (label, val, fmt) in enumerate(kpi_cards):
    col = 1 + (idx % 3) * 2
    row = r + (idx // 3) * 3

    lbl_cell = ws.cell(row=row, column=col, value=label)
    lbl_cell.font = FONT_KPI_LBL

    val_cell = ws.cell(row=row + 1, column=col, value=val)
    val_cell.font = FONT_KPI_VAL
    val_cell.number_format = fmt
    val_cell.fill = FILL_KPI_BOX
    # Merge value cell across 2 columns for width
    ws.merge_cells(start_row=row+1, start_column=col, end_row=row+1, end_column=col+1)

# Aging breakdown section
r_aging = 14
ws.cell(row=r_aging, column=1, value="Aging Breakdown").font = Font(name="Arial", size=13, bold=True, color="1F3864")
aging_headers = ["Bucket", "Tons", "% of Total"]
write_header_row(ws, r_aging + 1, aging_headers)
aging_rows = [
    ("0-30 Days", total_0_30, total_0_30/total_tons),
    ("31-60 Days", total_31_60, total_31_60/total_tons),
    ("61-90 Days", total_61_90, total_61_90/total_tons),
    ("90+ Days", total_90p, total_90p/total_tons),
    ("Total", total_tons, 1.0),
]
for i, (bucket, t, pct) in enumerate(aging_rows):
    fmt_list = [None, FMT_DEC1, FMT_PCT]
    is_total = (bucket == "Total")
    row_n = r_aging + 2 + i
    write_data_row(ws, row_n, [bucket, t, pct], fmt_list, alt=i % 2 == 0)
    if is_total:
        for c in range(1, 4):
            ws.cell(row=row_n, column=c).font = FONT_BODY_B

# Data note
ws.cell(row=r_aging + 8, column=1,
        value="All data is synthetic. Turnover is annualized from trailing 12-month shipments / average on-hand.").font = FONT_LINK

auto_width(ws)
ws.column_dimensions["A"].width = 22
ws.column_dimensions["C"].width = 22
ws.column_dimensions["E"].width = 22

# ═══════════════════════════════════════════════════════════════════════
# SHEET 2: KPI_By_Yard
# ═══════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("KPI_By_Yard")
ws2.sheet_properties.tabColor = "2E75B6"

ws2.cell(row=1, column=1, value="Inventory KPIs by Yard").font = FONT_TITLE
ws2.cell(row=2, column=1, value=f"As of {latest_month.strftime('%B %Y')}").font = FONT_KPI_LBL
ws2.merge_cells("A1:P1")

headers = list(yard_kpi_df.columns)
write_header_row(ws2, 4, headers)

formats_yard = [None, None, FMT_DEC1, FMT_DEC1, FMT_DOLLAR, FMT_DOLLAR, FMT_DOLLAR,
                FMT_DEC1, FMT_DEC1, FMT_NUM, FMT_PCT, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_PCT]

for i, (_, row) in enumerate(yard_kpi_df.iterrows()):
    write_data_row(ws2, 5 + i, list(row.values), formats_yard, alt=i % 2 == 0)

# Totals row
r_total = 5 + len(yard_kpi_df)
ws2.cell(row=r_total, column=1, value="TOTAL").font = FONT_BODY_B
for col_idx in range(3, len(headers) + 1):
    col_letter = get_column_letter(col_idx)
    cell = ws2.cell(row=r_total, column=col_idx)
    if col_idx in (8, 9, 11, 16):  # Turnover, Days, Utilization, Aged% → average
        cell.value = f"=AVERAGE({col_letter}5:{col_letter}{r_total-1})"
    else:
        cell.value = f"=SUM({col_letter}5:{col_letter}{r_total-1})"
    cell.font = FONT_BODY_B
    cell.border = BORDER_THIN
    if col_idx < len(formats_yard) + 1 and formats_yard[col_idx - 1]:
        cell.number_format = formats_yard[col_idx - 1]

auto_width(ws2)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 3: KPI_By_Metal
# ═══════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("KPI_By_Metal")
ws3.sheet_properties.tabColor = "548235"

ws3.cell(row=1, column=1, value="Inventory KPIs by Metal Group").font = FONT_TITLE
ws3.cell(row=2, column=1, value=f"As of {latest_month.strftime('%B %Y')}").font = FONT_KPI_LBL
ws3.merge_cells("A1:O1")

headers_m = list(metal_kpi_df.columns)
write_header_row(ws3, 4, headers_m)

formats_metal = [None, FMT_DEC1, FMT_NUM, FMT_NUM, FMT_DOLLAR, FMT_DOLLAR, FMT_DOLLAR,
                 FMT_DOLLAR2, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_PCT]

for i, (_, row) in enumerate(metal_kpi_df.iterrows()):
    write_data_row(ws3, 5 + i, list(row.values), formats_metal, alt=i % 2 == 0)

# Totals
r_total_m = 5 + len(metal_kpi_df)
ws3.cell(row=r_total_m, column=1, value="TOTAL").font = FONT_BODY_B
for col_idx in range(2, len(headers_m) + 1):
    col_letter = get_column_letter(col_idx)
    cell = ws3.cell(row=r_total_m, column=col_idx)
    if col_idx in (8, 9, 10, 15):  # Margin, Turnover, Days, Aged% → average
        cell.value = f"=AVERAGE({col_letter}5:{col_letter}{r_total_m-1})"
    else:
        cell.value = f"=SUM({col_letter}5:{col_letter}{r_total_m-1})"
    cell.font = FONT_BODY_B
    cell.border = BORDER_THIN
    if col_idx < len(formats_metal) + 1 and formats_metal[col_idx - 1]:
        cell.number_format = formats_metal[col_idx - 1]

auto_width(ws3)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 4: Aging_Detail
# ═══════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Aging_Detail")
ws4.sheet_properties.tabColor = "BF8F00"

ws4.cell(row=1, column=1, value="Inventory Aging by Yard and Grade").font = FONT_TITLE
ws4.cell(row=2, column=1, value=f"As of {latest_month.strftime('%B %Y')} — system month-end snapshot").font = FONT_KPI_LBL
ws4.merge_cells("A1:K1")

aging_headers = ["Yard", "Yard ID", "Grade ID", "Grade Name", "Metal Group",
                 "Tons on Hand", "0-30 Days", "31-60 Days", "61-90 Days", "90+ Days",
                 "Aged 90+ (%)", "Book Value ($)", "Market Value ($)"]
write_header_row(ws4, 4, aging_headers)

formats_aging = [None, None, None, None, None, FMT_DEC1, FMT_DEC1, FMT_DEC1, FMT_DEC1,
                 FMT_DEC1, FMT_PCT, FMT_DOLLAR, FMT_DOLLAR]

for i, (_, row) in enumerate(aging_summary.iterrows()):
    vals = [
        row["yard_name"], row["yard_id"], row["grade_id"], row["grade_name"],
        row["metal_group"].title(), row["tons_on_hand"],
        row["tons_0_30"], row["tons_31_60"], row["tons_61_90"], row["tons_over_90"],
        row["aged_pct"], row["book_value"], row["market_value"]
    ]
    write_data_row(ws4, 5 + i, vals, formats_aging, alt=i % 2 == 0)

    # Conditional: highlight rows with high aging
    if row["aged_pct"] > 0.20:
        for c in range(1, len(aging_headers) + 1):
            ws4.cell(row=5 + i, column=c).fill = FILL_RED_LT
    elif row["tons_61_90"] > 5 or row["tons_over_90"] > 0:
        for c in range(1, len(aging_headers) + 1):
            ws4.cell(row=5 + i, column=c).fill = FILL_YEL_LT

# Legend
legend_row = 5 + len(aging_summary) + 2
ws4.cell(row=legend_row, column=1, value="Legend:").font = FONT_BODY_B
ws4.cell(row=legend_row + 1, column=1, value="  Red = >20% of tons aged 90+ days").font = FONT_BODY
ws4.cell(row=legend_row + 1, column=1).fill = FILL_RED_LT
ws4.cell(row=legend_row + 2, column=1, value="  Yellow = has tons aged 61+ days").font = FONT_BODY
ws4.cell(row=legend_row + 2, column=1).fill = FILL_YEL_LT

auto_width(ws4)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 5: Margin_Analysis
# ═══════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Margin_Analysis")
ws5.sheet_properties.tabColor = "7030A0"

ws5.cell(row=1, column=1, value="Buy/Sell Margin Analysis by Grade").font = FONT_TITLE
ws5.cell(row=2, column=1, value=f"Prices as of {latest_price_date.strftime('%Y-%m-%d')}").font = FONT_KPI_LBL
ws5.merge_cells("A1:J1")

headers_mg = list(margin_analysis_df.columns)
write_header_row(ws5, 4, headers_mg)

formats_mg = [None, None, None, FMT_DEC1, FMT_DOLLAR2, FMT_DOLLAR2, FMT_DOLLAR2,
              FMT_PCT, FMT_PCT, FMT_DOLLAR]

for i, (_, row) in enumerate(margin_analysis_df.iterrows()):
    write_data_row(ws5, 5 + i, list(row.values), formats_mg, alt=i % 2 == 0)

    # Color the 30-day trend
    trend_cell = ws5.cell(row=5 + i, column=9)
    if row["30-Day Price Trend"] > 0.02:
        trend_cell.font = FONT_GREEN
    elif row["30-Day Price Trend"] < -0.02:
        trend_cell.font = FONT_RED

auto_width(ws5)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 6: At_Risk_Inventory
# ═══════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("At_Risk_Inventory")
ws6.sheet_properties.tabColor = "C00000"

ws6.cell(row=1, column=1, value="At-Risk Inventory — Material to Move").font = FONT_TITLE
ws6.cell(row=2, column=1, value=f"As of {latest_month.strftime('%B %Y')} — items scoring 2+ on risk factors").font = FONT_KPI_LBL
ws6.merge_cells("A1:O1")

if len(at_risk_df) > 0:
    headers_ar = list(at_risk_df.columns)
    write_header_row(ws6, 4, headers_ar)

    formats_ar = [None, None, None, None, None, FMT_DEC1, FMT_DOLLAR, FMT_DOLLAR,
                  FMT_DOLLAR, FMT_DEC1, FMT_DEC1, FMT_PCT, FMT_NUM, None, None]

    for i, (_, row) in enumerate(at_risk_df.iterrows()):
        write_data_row(ws6, 5 + i, list(row.values), formats_ar, alt=i % 2 == 0)

        # Priority rows in red
        if row["Risk Score"] >= 4:
            for c in range(1, len(headers_ar) + 1):
                ws6.cell(row=5 + i, column=c).fill = FILL_RED_LT
        elif row["Risk Score"] >= 3:
            for c in range(1, len(headers_ar) + 1):
                ws6.cell(row=5 + i, column=c).fill = FILL_YEL_LT

    # Summary
    sum_row = 5 + len(at_risk_df) + 2
    ws6.cell(row=sum_row, column=1, value="Summary:").font = FONT_BODY_B
    ws6.cell(row=sum_row + 1, column=1,
             value=f"  {len(at_risk_df)} yard-grade positions flagged").font = FONT_BODY
    ws6.cell(row=sum_row + 2, column=1,
             value=f"  {at_risk_df['Tons on Hand'].sum():,.0f} total tons at risk").font = FONT_BODY
    ws6.cell(row=sum_row + 3, column=1,
             value=f"  {at_risk_df[at_risk_df['Risk Score'] >= 4].shape[0]} priority items for immediate action").font = FONT_RED

    # Legend
    ws6.cell(row=sum_row + 5, column=1, value="Risk scoring:").font = FONT_BODY_B
    ws6.cell(row=sum_row + 6, column=1, value="  +3 = >20% of tons aged 90+ days").font = FONT_BODY
    ws6.cell(row=sum_row + 7, column=1, value="  +2 = market value below book value").font = FONT_BODY
    ws6.cell(row=sum_row + 8, column=1, value="  +1 = any tons over 90 days").font = FONT_BODY
    ws6.cell(row=sum_row + 9, column=1, value="  +1 = >5 tons at 61-90 days").font = FONT_BODY

else:
    ws6.cell(row=4, column=1, value="No at-risk inventory items found.").font = FONT_BODY

auto_width(ws6, max_width=40)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 7: Monthly_Trends
# ═══════════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("Monthly_Trends")
ws7.sheet_properties.tabColor = "4472C4"

ws7.cell(row=1, column=1, value="Monthly KPI Trends").font = FONT_TITLE
ws7.cell(row=2, column=1, value="October 2024 – September 2026 (24 months)").font = FONT_KPI_LBL
ws7.merge_cells("A1:L1")

trend_headers = ["Month", "Tons on Hand", "Tons Received", "Tons Shipped",
                 "Book Value ($)", "Market Value ($)", "Unrealized ($)",
                 "0-30 Days", "31-60 Days", "61-90 Days", "90+ Days",
                 "Aged 90+ (%)", "Utilization (%)"]
write_header_row(ws7, 4, trend_headers)

formats_trend = [None, FMT_NUM, FMT_NUM, FMT_NUM, FMT_DOLLAR, FMT_DOLLAR, FMT_DOLLAR,
                 FMT_NUM, FMT_NUM, FMT_NUM, FMT_NUM, FMT_PCT, FMT_PCT]

for i, (_, row) in enumerate(monthly.iterrows()):
    vals = [
        row["month_end"].strftime("%b %Y"),
        row["tons_on_hand"], row["tons_received"], row["tons_shipped"],
        row["book_value"], row["market_value"],
        row["market_value"] - row["book_value"],
        row["tons_0_30"], row["tons_31_60"], row["tons_61_90"], row["tons_over_90"],
        row["aged_pct"], row["utilization"],
    ]
    write_data_row(ws7, 5 + i, vals, formats_trend, alt=i % 2 == 0)

auto_width(ws7)

# ═══════════════════════════════════════════════════════════════════════
# SHEET 8: README
# ═══════════════════════════════════════════════════════════════════════
ws_readme = wb.create_sheet("README")
ws_readme.sheet_properties.tabColor = "808080"

# Move README to first position
wb.move_sheet("README", offset=-7)

readme_text = [
    ("SCRAP METAL INVENTORY KPI & AGING ANALYSIS", FONT_TITLE),
    ("", None),
    ("PURPOSE", Font(name="Arial", size=12, bold=True, color="1F3864")),
    ("Tracks inventory position, turnover, aging, margin, and at-risk material across 8 scrap-metal yards (40 grades, 24 months). Designed for a Commercial Planning & Inventory Analyst reviewing inventory health and deciding what to move.", FONT_BODY),
    ("", None),
    ("SHEET GUIDE", Font(name="Arial", size=12, bold=True, color="1F3864")),
    ("README — this page", FONT_BODY),
    ("KPI_Summary — network-wide headline metrics: tons, value, turnover, days of inventory, aging, capacity utilization, margin", FONT_BODY),
    ("KPI_By_Yard — same KPIs broken down by yard, with month-over-month change", FONT_BODY),
    ("KPI_By_Metal — same KPIs broken down by metal group (Copper, Aluminum, Brass, Stainless, Ferrous)", FONT_BODY),
    ("Aging_Detail — every yard-grade position with 30/60/90+ day aging buckets, highlighted by risk", FONT_BODY),
    ("Margin_Analysis — buy/sell spread and 30-day price trend for every grade", FONT_BODY),
    ("At_Risk_Inventory — flagged positions with risk scores and recommended actions (material-to-move list)", FONT_BODY),
    ("Monthly_Trends — 24-month time series of all key metrics", FONT_BODY),
    ("", None),
    ("KEY METRICS", Font(name="Arial", size=12, bold=True, color="1F3864")),
    (f"Total tons on hand: {total_tons:,.0f}", FONT_BODY),
    (f"Book value: ${total_book:,.0f}  |  Market value: ${total_market:,.0f}", FONT_BODY),
    (f"Inventory turnover (annualized): {turnover_annual:.1f}x  |  Days of inventory: {days_of_inv:.0f}", FONT_BODY),
    (f"Aged 90+ days: {total_90p:,.0f} tons ({total_90p/total_tons:.1%} of total)", FONT_BODY),
    (f"At-risk positions: {len(at_risk_df)} yard-grade pairs flagged for action", FONT_BODY),
    ("", None),
    ("TOOLS", Font(name="Arial", size=12, bold=True, color="1F3864")),
    ("Python (pandas, openpyxl) for data processing and workbook generation. All formulas in Excel.", FONT_BODY),
    ("", None),
    ("DATA NOTE", Font(name="Arial", size=12, bold=True, color="1F3864")),
    ("All data is synthetic, generated by a Python script simulating a scrap-metal yard network.", FONT_BODY),
]

for i, (text, font) in enumerate(readme_text):
    cell = ws_readme.cell(row=i + 1, column=1, value=text)
    if font:
        cell.font = font

ws_readme.column_dimensions["A"].width = 100

# ── print preview settings ─────────────────────────────────────────────
for ws_sheet in wb.worksheets:
    ws_sheet.sheet_view.showGridLines = False

# ── save ───────────────────────────────────────────────────────────────
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print(f"\nWorkbook saved to {OUT}")
print(f"Sheets: {wb.sheetnames}")
