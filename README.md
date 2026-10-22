# Scrap Metal Inventory & Commercial Planning Analytics

## Project description

A portfolio of inventory analytics workbooks for a fictional regional scrap-metal recycling network: eight yards in the NJ / PA / DE / MD corridor buying scrap from suppliers, sorting it into grades, and selling it to mills, foundries and export traders.

The project answers the questions a commercial planning and inventory analyst faces every week:

- How much material is on hand, by yard and grade, and what is it worth at today's market?
- Which inventory is aging, slow-moving or otherwise at risk?
- Where is inventory sitting relative to yard capacity and target stock levels?
- Which grades should Sales prioritize, and where should Logistics move material first?
- How should grades be classified for differentiated inventory management?
- Are the underlying transaction records trustworthy, and do they agree with the inventory system?

All analysis is delivered as Excel workbooks, with Python build scripts that generate each workbook programmatically from the shared dataset.

## Workbooks

### 1. Inventory Positioning & Movement Plan

`workbook/inventory_positioning.xlsx` | built by `scripts/build_positioning.py`

Calculates target stock levels for every yard-grade position using safety factors derived from ABC-XYZ classification and recent demand. Identifies overstocked, understocked and critically low positions, then generates an inter-yard rebalance plan to move excess material where it is needed.

| Sheet | Purpose |
|---|---|
| README | Dashboard with KPI cards: total positions, critically low count, rebalance tonnage |
| Positioning_Summary | Target vs. on-hand tons, gap analysis, position status for all 254 yard-grade pairs |
| Rebalance_Plan | Proposed inter-yard transfers ranked by priority (grade, from-yard, to-yard, tons) |
| Demand_Analysis | 6-month and 24-month demand trends, growth classification per grade |
| Yard_Capacity | Utilization, headroom, throughput, net flow and aging metrics per yard |
| Grade_Strategy | Combined ABC-XYZ + trend + aging recommendations for each grade |

### 2. ABC/XYZ Classification & Price Volatility

`workbook/abc_xyz_classification.xlsx` | built by `scripts/build_abc_xyz.py`

Classifies 40 grades on two dimensions: ABC by revenue contribution (A = top 80%, B = next 15%, C = bottom 5%) and XYZ by demand variability using coefficient of variation (X &le; 0.5, Y &le; 1.0, Z > 1.0). Combines them into a 9-cell matrix with tailored inventory management strategies. Adds price-volatility analysis with 30/90-day buy and sell price trends.

| Sheet | Purpose |
|---|---|
| README | Overview and methodology |
| ABC_Analysis | Grades ranked by cumulative revenue share with A/B/C classification |
| XYZ_Analysis | Grades ranked by demand variability (coefficient of variation) |
| ABC_XYZ_Matrix | Combined classification with strategy recommendations per cell |
| Price_Volatility | Monthly average buy/sell prices, trends and volatility ranking per grade |
| Strategy_Summary | Actionable strategies mapped to each ABC-XYZ combination |

### 3. Data Quality & Reconciliation

`workbook/inventory_analysis.xlsx`

Reconciles the transaction ledger against month-end system snapshots for all 254 yard-grade pairs across 24 months. Flags data-quality issues (duplicate rows, unit-of-measure errors, price outliers, missing customers) and classifies every variance by root cause.

| Sheet | Purpose |
|---|---|
| DQ_Summary, DQ_Price_Check | Counts of each data-quality flag |
| Recon_Summary | Variance events by cause: count, tons and estimated value |
| Recon_By_Metal | Unresolved variance by metal group |
| Recon_Exceptions | Ranked list of all variance movements |
| Recon_Findings | Written findings and recommendations |
| inventory_month_end | System balances with reconciliation columns added |
| ledger_monthly, opening_balance | Monthly ledger movements and starting balance |

### 4. KPI & Aging Analysis

`workbook/inventory_kpi_aging.xlsx` | built by `scripts/build_kpi_workbook.py`

Tracks operational KPIs across the network: turnover ratios, days-on-hand, carrying costs, throughput velocity and capacity utilization. Breaks down aging exposure by bucket (0-30, 31-60, 61-90, 90+ days) at the yard and grade level.

| Sheet | Purpose |
|---|---|
| README | Executive summary with network-wide KPIs |
| Turnover_Analysis | Inventory turnover ratio and days-on-hand by yard and grade |
| Aging_Summary | Aging distribution by bucket, ranked by exposure |
| Yard_Performance | Throughput, utilization, carrying cost and net flow per yard |
| Grade_Performance | Revenue, margin, velocity and aging risk per grade |
| Monthly_Trends | 24-month time series of key metrics across the network |

## Scope of the business simulated

| Dimension | Detail |
|---|---|
| Yards | 8 (Camden, Philadelphia, Trenton, Newark, Wilmington, Allentown, Baltimore, Reading) |
| Metal groups | Copper, aluminum, brass & bronze, stainless steel, ferrous |
| Grades | 40 (e.g. Bare Bright Copper, 6061 Extrusions, Yellow Brass, 304 Stainless Solids, HMS 1) |
| Customers | 12 mills, smelters, foundries and export traders |
| History | 24 months, October 2024 to September 2026 |
| Volume | About 100,000 purchase receipts and 100,000 outbound loads |
| Unit | Short tons; prices in USD per short ton |

<img width="515" height="728" alt="recon_exceptions" src="https://github.com/user-attachments/assets/fa6dcc7e-bb38-4c7d-b052-046bb4173a6a" />
<img width="726" height="226" alt="recon_by_metal" src="https://github.com/user-attachments/assets/a75e9d2c-2f7c-4237-9c2c-9945313c8f14" />


## Architecture

```
                 +--------------------+
                 |  generate_data.py  |   seeded, reproducible
                 +---------+----------+
                           |
      +--------------------+---------------------+
      |                    |                     |
      v                    v                     v
 data/reference/      data/raw/            data/system_snapshots/
 yards, grades,       transactions,        month-end balances,
 customers,           daily market         open FIFO lots,
 yard-grade map       benchmarks,          opening balances
                      daily grade prices
      |                    |                     |
      +--------------------+---------------------+
                           |
           +---------------+---------------+
           |               |               |
           v               v               v
   scripts/build_*    Power Query      Python analytics
   (openpyxl)         (inventory_      (classification,
                       analysis.xlsx)   positioning, KPIs)
           |               |               |
           v               v               v
      workbook/        workbook/        workbook/
   positioning,     inventory_       abc_xyz, kpi_aging
   abc_xyz, kpi     analysis
```

### Data layers

**Reference.** Static dimensions: yards with capacity, grades with metal group, standard load size and processing cost, customers with freight cost, and the list of grades each yard handles.

**Raw.** Extracts as an operational system would deliver them: the transaction ledger (receipts, shipments, physical-count adjustments), daily market benchmarks per metal group, and daily buy and sell prices per grade. The raw ledger contains the kinds of quality problems found in real extracts: inconsistent code formatting, duplicate rows, unit-of-measure and decimal errors, backdated shipments, missing records, and blank customers.

**System snapshots.** Balances as the inventory system reports them: month-end tons, average cost, book value and market value by yard and grade, split into age buckets (0-30, 31-60, 61-90, 90+ days); lots still open at the end of the period; and the opening balance.

### Data model

| Table | Grain | Key fields |
|---|---|---|
| `transactions` | one row per receipt, shipment or adjustment | txn_id, date, yard_id, grade_id, txn_type, tons, unit_price, customer_id, lot_id |
| `yards` | one row per yard | yard_id, capacity_tons |
| `grades` | one row per grade | grade_id, metal_group, standard_load_tons, processing_cost_per_ton |
| `customers` | one row per customer | customer_id, customer_type, metal_groups_bought, freight_per_ton |
| `yard_grades` | one row per yard-grade handled | yard_id, grade_id |
| `market_benchmarks_daily` | metal group by day | benchmark_per_ton |
| `grade_prices_daily` | grade by day | buy_price, sell_price |
| `inventory_month_end` | yard by grade by month-end | tons_on_hand, avg_cost_per_ton, book_value, market_value, age buckets |
| `inventory_open_lots` | one row per open FIFO lot | lot_id, receipt_date, tons_remaining, cost_per_ton, age_days |

### Simulation logic

- **Prices.** Each metal group follows a mean-reverting price path with its own volatility, correlated to a shared market factor. Brass tracks copper. Grade prices are a fixed percentage of the group benchmark, and each grade carries its own buy-to-sell spread.
- **Inbound material.** Daily purchases per yard and grade respond to season, weekday and recent price level.
- **Outbound shipments.** Full loads ship when a yard's stock exceeds a days-of-supply target. The target loosens or tightens with 30-day price momentum, so yards hold more when prices are rising. Small piles ship as consolidated partial loads once they age.
- **Inventory accounting.** First-in, first-out lots, so aging, average cost and valuation are exact.
- **Built-in business scenarios.** Some yard-grade combinations ship rarely, and a few grades lose demand partway through the period, producing the aged and slow-moving inventory an analyst would be asked to find.
- **Capacity.** Yard capacity is sized so busy yards run close to, and occasionally over, capacity.

## Repository layout

```
.
|-- README.md
|-- generate_data.py
|-- scripts/
|   |-- build_positioning.py
|   |-- build_abc_xyz.py
|   `-- build_kpi_workbook.py
|-- workbook/
|   |-- inventory_positioning.xlsx
|   |-- abc_xyz_classification.xlsx
|   |-- inventory_analysis.xlsx
|   `-- inventory_kpi_aging.xlsx
|-- images/
`-- data/
    |-- reference/
    |-- raw/
    |-- system_snapshots/
    `-- validation/
```

## Data note

All data is synthetic. Yards, customers, quantities and prices are fictional and are not drawn from any company's records. Dollar values are estimates at median purchase price, not accounting values.
