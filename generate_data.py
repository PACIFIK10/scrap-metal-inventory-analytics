#!/usr/bin/env python3
"""
Synthetic scrap-metal inventory dataset generator.

Produces a two-year operating history (2024-10-01 to 2026-09-30) for a fictional
regional network of eight scrap-metal yards handling 40 grades across five metal
groups (copper, aluminum, brass, stainless, ferrous).

Design in one paragraph
-----------------------
Market benchmarks are simulated per metal group (or taken from real monthly
series when supplied). Each yard-grade combination receives purchased scrap
daily (volume reacts to price and season), and ships full or partial loads to
customers using a days-of-supply target that tightens or loosens with price
momentum. Inventory is tracked in FIFO lots so aging is exact. A clean ledger
drives the system snapshots (month-end balances, open lots). The raw
transactions extract is then corrupted with a documented set of data-quality
problems, and every planted problem is recorded in a validation file.

All yards, customers and quantities are fictional.
"""
import argparse
import os
from collections import deque

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------
# Calendar
# --------------------------------------------------------------------------
PRICE_START = pd.Timestamp("2024-01-01")   # history for moving averages
WARMUP_START = pd.Timestamp("2024-07-01")  # inventory builds up before the window
START = pd.Timestamp("2024-10-01")         # first reported day
END = pd.Timestamp("2026-09-30")

# --------------------------------------------------------------------------
# Reference data
# --------------------------------------------------------------------------
# id, name, city, state, relative volume, share of grades handled, fill level at the 90th-percentile month
YARDS = [
    ("CAM", "Camden Main Yard", "Camden", "NJ", 1.00, 1.00, 0.88),
    ("PHL", "Philadelphia Yard", "Philadelphia", "PA", 0.85, 0.90, 0.82),
    ("TRE", "Trenton Yard", "Trenton", "NJ", 0.45, 0.75, 0.72),
    ("NWK", "Newark Yard", "Newark", "NJ", 0.75, 0.85, 0.95),
    ("WIL", "Wilmington Yard", "Wilmington", "DE", 0.50, 0.75, 0.78),
    ("ALL", "Allentown Yard", "Allentown", "PA", 0.40, 0.70, 0.70),
    ("BAL", "Baltimore Yard", "Baltimore", "MD", 0.80, 0.85, 0.85),
    ("REA", "Reading Yard", "Reading", "PA", 0.30, 0.65, 0.80),
]

# Group-level settings
#   load: tons per outbound load; target_days: normal days-of-supply held;
#   ticket: average tons per inbound delivery; amp: seasonal amplitude;
#   spread: (min, max) buy discount to sell price; proc: processing cost $/ton
GROUPS = {
    "COPPER":   dict(prefix="CU", load=20, target_days=20, ticket=1.5, amp=0.10, spread=(0.05, 0.09), proc=60),
    "ALUMINUM": dict(prefix="AL", load=20, target_days=25, ticket=3.0, amp=0.15, spread=(0.10, 0.16), proc=35),
    "BRASS":    dict(prefix="BR", load=22, target_days=28, ticket=2.0, amp=0.07, spread=(0.10, 0.15), proc=45),
    "STAINLESS": dict(prefix="SS", load=22, target_days=30, ticket=4.0, amp=0.05, spread=(0.12, 0.20), proc=40),
    "FERROUS":  dict(prefix="FE", load=24, target_days=12, ticket=12.0, amp=0.08, spread=(0.15, 0.24), proc=18),
}

# code, name, group, price factor vs. group benchmark, annual tons for a full-size yard
GRADES = [
    ("CU-01", "Bare Bright Copper", "COPPER", 0.95, 900),
    ("CU-02", "#1 Copper", "COPPER", 0.92, 2200),
    ("CU-03", "#2 Copper", "COPPER", 0.88, 1800),
    ("CU-04", "Insulated Copper Wire", "COPPER", 0.45, 3000),
    ("CU-05", "Copper Turnings", "COPPER", 0.80, 500),
    ("CU-06", "Copper Transformers", "COPPER", 0.35, 700),
    ("CU-07", "Copper Motors", "COPPER", 0.25, 1200),
    ("AL-01", "6063 Extrusions", "ALUMINUM", 0.82, 2500),
    ("AL-02", "6061 Extrusions", "ALUMINUM", 0.80, 2600),
    ("AL-03", "Clean Aluminum Sheet", "ALUMINUM", 0.78, 1800),
    ("AL-04", "Cast Aluminum", "ALUMINUM", 0.68, 3500),
    ("AL-05", "Aluminum Wheels", "ALUMINUM", 0.72, 1500),
    ("AL-06", "Used Beverage Cans", "ALUMINUM", 0.70, 2200),
    ("AL-07", "Clean Aluminum Turnings", "ALUMINUM", 0.60, 1000),
    ("AL-08", "Aluminum Siding", "ALUMINUM", 0.68, 800),
    ("AL-09", "Painted Aluminum Sheet", "ALUMINUM", 0.72, 1200),
    ("BR-01", "Yellow Brass", "BRASS", 1.00, 1200),
    ("BR-02", "Red Brass", "BRASS", 1.22, 800),
    ("BR-03", "Brass Turnings", "BRASS", 0.88, 900),
    ("BR-04", "Brass Shells", "BRASS", 1.05, 300),
    ("BR-05", "Brass Radiators", "BRASS", 0.78, 700),
    ("BR-06", "Heavy Brass", "BRASS", 1.00, 600),
    ("BR-07", "Bronze", "BRASS", 0.95, 400),
    ("SS-01", "304 Stainless Solids", "STAINLESS", 1.00, 2800),
    ("SS-02", "304 Stainless Turnings", "STAINLESS", 0.80, 1200),
    ("SS-03", "316 Stainless Solids", "STAINLESS", 1.35, 900),
    ("SS-04", "316 Stainless Turnings", "STAINLESS", 1.10, 500),
    ("SS-05", "430 Stainless Solids", "STAINLESS", 0.35, 1800),
    ("SS-06", "18-8 Stainless Sheet", "STAINLESS", 0.95, 1600),
    ("SS-07", "Stainless Sinks", "STAINLESS", 0.60, 600),
    ("SS-08", "Mixed Stainless", "STAINLESS", 0.55, 900),
    ("FE-01", "HMS 1", "FERROUS", 1.00, 45000),
    ("FE-02", "HMS 2", "FERROUS", 0.90, 20000),
    ("FE-03", "HMS 1&2 80/20", "FERROUS", 0.97, 38000),
    ("FE-04", "Shredded Steel", "FERROUS", 1.03, 55000),
    ("FE-05", "Busheling", "FERROUS", 1.10, 12000),
    ("FE-06", "Plate & Structural", "FERROUS", 1.02, 15000),
    ("FE-07", "Cast Iron", "FERROUS", 1.25, 6000),
    ("FE-08", "Steel Turnings", "FERROUS", 0.70, 8000),
    ("FE-09", "#1 Bundles", "FERROUS", 1.08, 7000),
]

# Group benchmark starting levels, USD per short ton (illustrative, not quotes)
BENCH = {
    # level, annual vol, mean-reversion speed per day, correlation to market factor
    "COPPER": (10000.0, 0.20, 0.0020, 0.60),
    "ALUMINUM": (2500.0, 0.17, 0.0020, 0.50),
    "STAINLESS": (1300.0, 0.22, 0.0020, 0.40),
    "FERROUS": (380.0, 0.30, 0.0030, 0.35),
}
BRASS_TO_COPPER = 0.62

# id, name, type, groups bought, region, freight $/ton, weight
CUSTOMERS = [
    ("C01", "Copper Rod Mill A", "Mill", ["COPPER"], "PA", 38, 3),
    ("C02", "Copper & Brass Ingot Maker", "Smelter", ["COPPER", "BRASS"], "OH", 62, 2),
    ("C03", "Brass Foundry A", "Foundry", ["BRASS"], "NJ", 22, 3),
    ("C04", "Aluminum Billet Mill A", "Mill", ["ALUMINUM"], "PA", 40, 3),
    ("C05", "Aluminum Smelter B", "Smelter", ["ALUMINUM"], "NY", 55, 2),
    ("C06", "Aluminum Can Sheet Mill", "Mill", ["ALUMINUM"], "OH", 65, 2),
    ("C07", "Stainless Melt Shop A", "Mill", ["STAINLESS"], "PA", 45, 3),
    ("C08", "Export Trader (Non-Ferrous)", "Export Trader",
     ["COPPER", "ALUMINUM", "STAINLESS", "BRASS"], "NJ", 18, 2),
    ("C09", "Steel Mini-Mill A", "Mill", ["FERROUS"], "DE", 14, 4),
    ("C10", "Steel Mini-Mill B", "Mill", ["FERROUS"], "PA", 20, 3),
    ("C11", "Cast Iron Foundry", "Foundry", ["FERROUS"], "MD", 18, 1),
    ("C12", "Export Broker (Bulk Ferrous)", "Export Trader", ["FERROUS"], "NJ", 12, 2),
]

# Scenario grades that become sluggish partway through the window (level 1)
SCENARIO_SLOW_GRADES = {"AL-09": pd.Timestamp("2026-01-01"),
                        "CU-06": pd.Timestamp("2026-02-01"),
                        "SS-05": pd.Timestamp("2026-03-01")}
N_STATIC_SLOW_COMBOS = 14   # yard-grade pairs that are structurally stuck (level 2)


# --------------------------------------------------------------------------
# Market benchmarks
# --------------------------------------------------------------------------
def load_fred_series(path, dates, usd_per_metric_ton=True):
    """Read a FRED CSV (date, value), interpolate to daily, convert to $/short ton."""
    s = pd.read_csv(path)
    s.columns = ["date", "value"]
    s["date"] = pd.to_datetime(s["date"])
    s["value"] = pd.to_numeric(s["value"], errors="coerce")
    s = s.dropna().set_index("date")["value"]
    daily = s.reindex(s.index.union(dates)).interpolate("time").reindex(dates).ffill().bfill()
    if usd_per_metric_ton:
        daily = daily * 0.907185
    return daily.values


def make_benchmarks(dates, rng, fred_dir=None):
    n = len(dates)
    bday = (dates.dayofweek < 5)
    nb = int(bday.sum())
    market = rng.standard_normal(nb)
    out = {}
    for g, (base, vol, kappa, rho) in BENCH.items():
        eps = rng.standard_normal(nb)
        shock = rho * market + np.sqrt(1 - rho ** 2) * eps
        sd = vol / np.sqrt(252)
        x = np.zeros(nb)
        for i in range(1, nb):
            x[i] = x[i - 1] * (1 - kappa) + sd * shock[i]
        series_b = base * np.exp(x)
        full = np.full(n, np.nan)
        full[bday] = series_b
        full = pd.Series(full).ffill().bfill().values  # weekends hold Friday's price
        full = full * base / full[dates >= START].mean()  # keep the reporting window centred on the base level
        out[g] = full
    if fred_dir:
        for g, fname in (("COPPER", "PCOPPUSDM.csv"), ("ALUMINUM", "PALUMUSDM.csv")):
            p = os.path.join(fred_dir, fname)
            if os.path.exists(p):
                real = load_fred_series(p, dates)
                out[g] = real * np.exp(rng.normal(0, 0.003, n))
                print(f"  using real monthly series for {g} from {fname}")
    # Brass follows copper with its own slow-moving basis
    ar = np.zeros(n)
    for i in range(1, n):
        ar[i] = 0.98 * ar[i - 1] + rng.normal(0, 0.002)
    out["BRASS"] = BRASS_TO_COPPER * out["COPPER"] * (1 + ar)
    return out


# --------------------------------------------------------------------------
# Main simulation
# --------------------------------------------------------------------------
def build(out_dir, seed, fred_dir=None):
    rng = np.random.default_rng(seed)
    d_ref = os.path.join(out_dir, "data", "reference")
    d_raw = os.path.join(out_dir, "data", "raw")
    d_sys = os.path.join(out_dir, "data", "system_snapshots")
    d_val = os.path.join(out_dir, "data", "validation")
    for d in (d_ref, d_raw, d_sys, d_val):
        os.makedirs(d, exist_ok=True)

    group_names = list(GROUPS.keys())
    g_idx = {g: i for i, g in enumerate(group_names)}
    n_grades = len(GRADES)
    grade_codes = [g[0] for g in GRADES]
    grade_group = np.array([g_idx[g[2]] for g in GRADES])
    grade_factor = np.array([g[3] for g in GRADES])

    # ---- prices -----------------------------------------------------------
    all_dates = pd.date_range(PRICE_START, END)
    bench = make_benchmarks(all_dates, rng, fred_dir)
    bench_mat = np.column_stack([bench[g] for g in group_names])          # (T, groups)
    ma180 = pd.DataFrame(bench_mat).rolling(180, min_periods=30).mean().bfill().values
    mom30 = np.zeros_like(bench_mat)
    mom30[30:] = bench_mat[30:] / bench_mat[:-30] - 1.0
    price_mult = np.clip((bench_mat / ma180) ** 0.8, 0.75, 1.35)

    spreads = np.array([rng.uniform(*GROUPS[g[2]]["spread"]) for g in GRADES])
    grade_noise = rng.normal(0, 0.0025, (len(all_dates), n_grades))
    sell_px = bench_mat[:, grade_group] * grade_factor[None, :] * (1 + grade_noise)
    buy_px = sell_px * (1 - spreads[None, :] * (1 + rng.normal(0, 0.05, (len(all_dates), n_grades))))
    px_off = (WARMUP_START - PRICE_START).days

    # ---- yard-grade combinations -----------------------------------------
    combos = []
    for yi, y in enumerate(YARDS):
        for gi in range(n_grades):
            if y[0] == "CAM" or rng.random() < y[5]:
                combos.append((yi, gi))
    C = len(combos)
    c_yard = np.array([c[0] for c in combos])
    c_grade = np.array([c[1] for c in combos])
    c_group = grade_group[c_grade]
    ginfo = [GROUPS[group_names[k]] for k in c_group]
    lam_base = np.array([
        GRADES[c_grade[i]][4] * YARDS[c_yard[i]][4] * np.exp(rng.normal(0, 0.2)) / 289.0
        for i in range(C)])
    tdays = np.array([ginfo[i]["target_days"] * rng.uniform(0.8, 1.25) for i in range(C)])
    load_c = np.array([ginfo[i]["load"] for i in range(C)], dtype=float)
    ticket_c = np.array([ginfo[i]["ticket"] for i in range(C)], dtype=float)

    nonfe = [i for i in range(C) if group_names[c_group[i]] != "FERROUS"]
    static_slow = set(rng.choice(nonfe, size=N_STATIC_SLOW_COMBOS, replace=False).tolist())

    # ---- inbound demand matrix -------------------------------------------
    sim_dates = pd.date_range(WARMUP_START, END)
    D = len(sim_dates)
    month = sim_dates.month.values
    dow = sim_dates.dayofweek.values
    dow_f = np.array([1, 1, 1, 1, 1, 0.55, 0.0])[dow]
    amp = np.array([GROUPS[g]["amp"] for g in group_names])
    season = 1 + amp[None, :] * np.sin(2 * np.pi * (month[:, None] - 4) / 12)   # (D, groups)
    lam = (lam_base[None, :] * season[:, c_group] * dow_f[:, None]
           * price_mult[px_off:px_off + D][:, c_group])
    n_tix = rng.poisson(lam / ticket_c[None, :])
    tons_in = np.where(n_tix > 0,
                       ticket_c[None, :] * rng.gamma(np.maximum(2.0 * n_tix, 1e-9), 0.5), 0.0)
    recv_noise = rng.normal(0, 0.012, (D, C))
    ship_noise = rng.normal(0, 0.010, (D, C))

    # ---- customers -------------------------------------------------------
    cust_by_group = {}
    for g in group_names:
        ids = [c for c in CUSTOMERS if g in c[3]]
        w = np.array([c[6] for c in ids], dtype=float)
        cust_by_group[g] = ([c[0] for c in ids], w / w.sum())

    # ---- simulate --------------------------------------------------------
    lots = [deque() for _ in range(C)]   # each lot: [lot_id, date, qty_remaining, price]
    oh = np.zeros(C)
    lot_seq = 0
    txn_rows = []
    opening_rows = []
    snap_rows = []
    day_before_start = START - pd.Timedelta(days=1)

    def consume(ci, qty):
        rem = qty
        dq = lots[ci]
        while rem > 1e-9 and dq:
            lot = dq[0]
            take = min(lot[2], rem)
            lot[2] -= take
            rem -= take
            if lot[2] <= 1e-6:
                dq.popleft()

    def tidy(ci):
        if oh[ci] < 0.005:
            oh[ci] = 0.0
            lots[ci].clear()

    def slow_level(ci, date):
        if ci in static_slow:
            return 2
        code = grade_codes[c_grade[ci]]
        if code in SCENARIO_SLOW_GRADES and date >= SCENARIO_SLOW_GRADES[code]:
            return 1
        return 0

    for di in range(D):
        date = sim_dates[di]
        gi_px = px_off + di
        reporting = date >= START
        is_sunday = dow[di] == 6
        for ci in range(C):
            g = c_grade[ci]
            grp = c_group[ci]
            # inbound
            t = tons_in[di, ci]
            if t > 0.05:
                t = round(float(t), 2)
                price = round(float(buy_px[gi_px, g] * (1 + recv_noise[di, ci])), 2)
                lot_seq += 1
                lid = f"L{lot_seq:07d}"
                lots[ci].append([lid, date, t, price])
                oh[ci] += t
                if reporting:
                    txn_rows.append((date, YARDS[c_yard[ci]][0], grade_codes[g], "RECEIPT", t, price, "", lid))
            # outbound
            if is_sunday or oh[ci] <= 0:
                continue
            lvl = slow_level(ci, date)
            hold = 1 + 4 * float(np.clip(mom30[gi_px, grp], -0.15, 0.15))
            tgt = lam_base[ci] * tdays[ci] * hold * (1, 3, 6)[lvl]
            if rng.random() >= (0.55, 0.20, 0.05)[lvl]:
                continue
            load = load_c[ci]
            ship_list = []
            if oh[ci] >= max(tgt, load):
                n_loads = max(1, int((oh[ci] - 0.6 * tgt) // load))
                n_loads = min(n_loads, int(oh[ci] // load))
                for _ in range(n_loads):
                    ship_list.append(min(oh[ci] - sum(ship_list),
                                         round(load * rng.uniform(0.93, 1.0), 2)))
            elif lvl == 0 and oh[ci] >= 3 and (date - lots[ci][0][1]).days >= 45:
                ship_list.append(min(oh[ci], load))
            for q in ship_list:
                q = round(float(q), 2)
                if q <= 0:
                    continue
                consume(ci, q)
                oh[ci] -= q
                ids, w = cust_by_group[group_names[grp]]
                cust = ids[int(rng.choice(len(ids), p=w))]
                price = round(float(sell_px[gi_px, g] * (1 + ship_noise[di, ci])), 2)
                if reporting:
                    txn_rows.append((date, YARDS[c_yard[ci]][0], grade_codes[g], "SHIPMENT", q, price, cust, ""))
            tidy(ci)

        if date == day_before_start:
            for ci in range(C):
                for lot in lots[ci]:
                    opening_rows.append((lot[0], YARDS[c_yard[ci]][0], grade_codes[c_grade[ci]],
                                         lot[1].date().isoformat(), round(lot[2], 2), lot[3]))

        if reporting and date.is_month_end:
            for ci in range(C):
                g = c_grade[ci]
                # physical count adjustment
                if oh[ci] > 5 and rng.random() < 0.03:
                    if rng.random() < 0.85:
                        adj = -round(float(oh[ci] * rng.uniform(0.004, 0.025)), 2)
                        consume(ci, -adj)
                    else:
                        adj = round(float(oh[ci] * rng.uniform(0.003, 0.015)), 2)
                        lot_seq += 1
                        lots[ci].append([f"L{lot_seq:07d}", date, adj, round(float(buy_px[gi_px, g]), 2)])
                    oh[ci] += adj
                    tidy(ci)
                    txn_rows.append((date, YARDS[c_yard[ci]][0], grade_codes[g], "ADJUSTMENT", adj,
                                     round(float(buy_px[gi_px, g]), 2), "", ""))
                b = [0.0, 0.0, 0.0, 0.0]
                cost = 0.0
                for lot in lots[ci]:
                    age = (date - lot[1]).days
                    k = 0 if age <= 30 else 1 if age <= 60 else 2 if age <= 90 else 3
                    b[k] += lot[2]
                    cost += lot[2] * lot[3]
                tons = sum(b)
                avg_cost = cost / tons if tons > 0 else 0.0
                mkt = float(sell_px[gi_px, g])
                snap_rows.append((date.date().isoformat(), YARDS[c_yard[ci]][0], grade_codes[g],
                                  round(tons, 2), round(avg_cost, 2), round(tons * avg_cost, 2),
                                  round(tons * mkt, 2), round(b[0], 2), round(b[1], 2),
                                  round(b[2], 2), round(b[3], 2)))

    # ---- assemble clean ledger -------------------------------------------
    tx = pd.DataFrame(txn_rows, columns=["date", "yard_id", "grade_id", "txn_type", "tons",
                                         "unit_price", "customer_id", "lot_id"])
    tx = tx.sort_values("date", kind="stable").reset_index(drop=True)
    tx.insert(0, "txn_id", [f"T{i + 1:07d}" for i in range(len(tx))])
    tx_clean = tx.copy()

    snap = pd.DataFrame(snap_rows, columns=[
        "month_end", "yard_id", "grade_id", "tons_on_hand", "avg_cost_per_ton", "book_value",
        "market_value", "tons_0_30", "tons_31_60", "tons_61_90", "tons_over_90"])
    opening = pd.DataFrame(opening_rows, columns=[
        "lot_id", "yard_id", "grade_id", "receipt_date", "tons", "cost_per_ton"])

    # open lots at END
    open_rows = []
    for ci in range(C):
        for lot in lots[ci]:
            open_rows.append((lot[0], YARDS[c_yard[ci]][0], grade_codes[c_grade[ci]],
                              lot[1].date().isoformat(), round(lot[2], 2), lot[3],
                              (END - lot[1]).days))
    open_lots = pd.DataFrame(open_rows, columns=[
        "lot_id", "yard_id", "grade_id", "receipt_date", "tons_remaining", "cost_per_ton", "age_days"])

    # ---- reference tables --------------------------------------------------
    # Capacity is sized so each yard's 90th-percentile month-end load sits near its target
    # fill level, which leaves busy yards running close to (or briefly over) capacity.
    p90 = snap.groupby(["month_end", "yard_id"])["tons_on_hand"].sum().groupby("yard_id").quantile(0.9)
    yards_df = pd.DataFrame([
        (y[0], y[1], y[2], y[3],
         int(round(p90[y[0]] / y[6], -2))) for y in YARDS],
        columns=["yard_id", "yard_name", "city", "state", "capacity_tons"])
    grades_df = pd.DataFrame([
        (g[0], g[1], g[2], GROUPS[g[2]]["load"], GROUPS[g[2]]["proc"]) for g in GRADES],
        columns=["grade_id", "grade_name", "metal_group", "standard_load_tons", "processing_cost_per_ton"])
    cust_df = pd.DataFrame([
        (c[0], c[1], c[2], ";".join(c[3]), c[4], c[5]) for c in CUSTOMERS],
        columns=["customer_id", "customer_name", "customer_type", "metal_groups_bought",
                 "region", "freight_per_ton"])
    yg_df = pd.DataFrame(sorted({(YARDS[c_yard[i]][0], grade_codes[c_grade[i]]) for i in range(C)}),
                         columns=["yard_id", "grade_id"])

    rep_mask = all_dates >= START
    bm = pd.DataFrame(bench_mat[rep_mask], columns=group_names)
    bm.insert(0, "date", all_dates[rep_mask].date)
    bm = bm.melt(id_vars="date", var_name="metal_group", value_name="benchmark_per_ton")
    bm["benchmark_per_ton"] = bm["benchmark_per_ton"].round(2)
    gp = []
    for k, gcode in enumerate(grade_codes):
        gp.append(pd.DataFrame({
            "date": all_dates[rep_mask].date, "grade_id": gcode,
            "buy_price": buy_px[rep_mask, k].round(2), "sell_price": sell_px[rep_mask, k].round(2)}))
    gp = pd.concat(gp, ignore_index=True)

    # ---- corrupt the raw extract ----------------------------------------
    r2 = np.random.default_rng(seed + 1)
    raw = tx.copy()
    issues = []
    used = set()

    def pick(mask, k):
        pool = [i for i in raw.index[mask] if i not in used]
        chosen = r2.choice(pool, size=k, replace=False).tolist()
        used.update(chosen)
        return chosen

    is_recv = (raw["txn_type"] == "RECEIPT").values
    is_ship = (raw["txn_type"] == "SHIPMENT").values
    is_move = is_recv | is_ship
    allrows = np.ones(len(raw), dtype=bool)

    for i in pick(is_move, 40):                         # grade code formatting
        old = raw.at[i, "grade_id"]
        new = r2.choice([old.lower(), old + " ", old.replace("-", "_"), " " + old])
        raw.at[i, "grade_id"] = new
        issues.append((raw.at[i, "txn_id"], "GRADE_CODE_FORMAT", f"'{old}' -> '{new}'"))
    for i in pick(is_move, 10):                         # yard code formatting
        old = raw.at[i, "yard_id"]
        raw.at[i, "yard_id"] = old.lower()
        issues.append((raw.at[i, "txn_id"], "YARD_CODE_FORMAT", f"'{old}' -> '{old.lower()}'"))
    for i in pick(is_move, 25):                         # price decimal errors
        old = raw.at[i, "unit_price"]
        new = round(old * r2.choice([10, 0.1]), 2)
        raw.at[i, "unit_price"] = new
        issues.append((raw.at[i, "txn_id"], "PRICE_OUTLIER", f"{old} -> {new}"))
    for i in pick(is_move & (raw["tons"] > 0).values, 15):   # pounds entered as tons
        old = raw.at[i, "tons"]
        new = round(old * 2000, 2)
        raw.at[i, "tons"] = new
        issues.append((raw.at[i, "txn_id"], "UOM_LBS_AS_TONS", f"{old} -> {new}"))
    for i in pick(is_ship, 30):                         # backdated shipments
        old = raw.at[i, "date"]
        shift = int(r2.integers(20, 61))
        new = max(old - pd.Timedelta(days=shift), START)
        raw.at[i, "date"] = new
        issues.append((raw.at[i, "txn_id"], "BACKDATED_SHIPMENT",
                       f"{old.date()} -> {new.date()} (may create a negative balance)"))
    for i in pick(is_recv, 12):                         # non-positive receipt weights
        old = raw.at[i, "tons"]
        new = 0.0 if r2.random() < 0.34 else -old
        raw.at[i, "tons"] = new
        issues.append((raw.at[i, "txn_id"], "NONPOSITIVE_TONS", f"{old} -> {new}"))
    for i in pick(is_ship, 20):                         # missing customer
        raw.at[i, "customer_id"] = ""
        issues.append((raw.at[i, "txn_id"], "MISSING_CUSTOMER", "customer_id blanked"))
    dup_idx = pick(allrows, 60)                         # duplicate rows (appended later)
    for i in dup_idx:
        issues.append((raw.at[i, "txn_id"], "DUPLICATE_ROW", "row appears twice with the same txn_id"))
    dups = raw.loc[dup_idx].copy()
    del_idx = pick(is_recv, 30)                         # missing receipts (rows removed)
    for i in del_idx:
        issues.append((raw.at[i, "txn_id"], "MISSING_RECEIPT",
                       f"row removed: {raw.at[i, 'yard_id']} {raw.at[i, 'grade_id']} {raw.at[i, 'tons']} t"))
    raw = raw.drop(index=del_idx)
    raw = pd.concat([raw, dups], ignore_index=True).sort_values("date", kind="stable").reset_index(drop=True)
    raw["date"] = raw["date"].dt.date
    tx_clean["date"] = tx_clean["date"].dt.date

    issues_df = pd.DataFrame(issues, columns=["txn_id", "issue_type", "detail"])

    scen_rows = []
    for ci in sorted(static_slow):
        scen_rows.append(("STRUCTURALLY_SLOW_COMBO", YARDS[c_yard[ci]][0], grade_codes[c_grade[ci]],
                          "ships rarely for the entire window; builds aged inventory"))
    for code, d0 in SCENARIO_SLOW_GRADES.items():
        scen_rows.append(("DEMAND_SLOWDOWN_ALL_YARDS", "ALL", code,
                          f"outbound shipments slow materially from {d0.date()}"))
    scen_df = pd.DataFrame(scen_rows, columns=["scenario", "yard_id", "grade_id", "description"])

    # ---- write -------------------------------------------------------------
    yards_df.to_csv(os.path.join(d_ref, "yards.csv"), index=False)
    grades_df.to_csv(os.path.join(d_ref, "grades.csv"), index=False)
    cust_df.to_csv(os.path.join(d_ref, "customers.csv"), index=False)
    yg_df.to_csv(os.path.join(d_ref, "yard_grades.csv"), index=False)
    raw.to_csv(os.path.join(d_raw, "transactions.csv"), index=False)
    bm.to_csv(os.path.join(d_raw, "market_benchmarks_daily.csv"), index=False)
    gp.to_csv(os.path.join(d_raw, "grade_prices_daily.csv"), index=False)
    snap.to_csv(os.path.join(d_sys, "inventory_month_end.csv"), index=False)
    open_lots.to_csv(os.path.join(d_sys, "inventory_open_lots.csv"), index=False)
    opening.to_csv(os.path.join(d_sys, "opening_balance_lots.csv"), index=False)
    issues_df.to_csv(os.path.join(d_val, "planted_issues.csv"), index=False)
    scen_df.to_csv(os.path.join(d_val, "planted_scenarios.csv"), index=False)
    tx_clean.to_csv(os.path.join(d_val, "transactions_clean.csv"), index=False)

    print(f"yard-grade combinations: {C}")
    print(f"transactions (raw extract): {len(raw):,}   clean ledger: {len(tx_clean):,}")
    print(f"planted issues: {len(issues_df)} rows across {issues_df['issue_type'].nunique()} types")
    print(f"output written under: {os.path.join(out_dir, 'data')}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Generate the synthetic scrap-metal inventory dataset.")
    ap.add_argument("--out", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--fred-dir", default=None,
                    help="optional folder with PCOPPUSDM.csv / PALUMUSDM.csv from FRED")
    a = ap.parse_args()
    build(a.out, a.seed, a.fred_dir)
