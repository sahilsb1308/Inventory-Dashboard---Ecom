# Swiss Beauty — Inventory Dashboard (Ecom)

Automated sync script that pulls sales, stock, and revenue data from multiple Google Sheets and writes a clean inventory dashboard with priority scoring. Runs daily at 12:00 noon IST via GitHub Actions.

## What it does

| Step | Source | Data |
|------|--------|------|
| Sales (MTD) | `Compiled Sec Data Ecomm` | Units sold, split by B2B / B2C |
| Stock | Inventory Dashboard sheet (tab 0) | B2B: col AB (index 27) · B2C: col AR (index 43) |
| Mother WH Stock | Inventory Dashboard sheet | col T (index 19) — always 0 if absent |
| Revenue | `Secondary Sale (25 to 27)` | MRP Sale, last 90 days; Category/Sub-Category from cols W/X |
| Product names | Easyecomm + ERP sheets | Fallback chain (see below) |

**Output:** `B2B` and `B2C` tabs in the destination sheet with 18 columns:

`IMS Code · EAN Code · Product Name · Variant Title · Brand · Category · Sub-Category · Launch Type · Units Sold (90D) · Current Stock · Mother WH Stock · DRR · Projected Demand (60D) · Days of Cover · Stock Status · Revenue (Last 90D) · Revenue Contribution % · Priority`

## SKU matching

Each inventory sheet row is indexed by **four keys** so any identifier resolves correctly:

1. Col A — ERP Code OLD MRP (e.g. `SB-10-01`)
2. Col B — ERP Code NEW MRP (e.g. `SB-10L-01`)
3. Col F — Old EAN barcode
4. Col G — New EAN barcode

## Product name fallback chain

1. `inv_meta` product name (Inventory Dashboard col D)
2. `ecomm_names` (Inventory Report-Easyecomm)
3. `rev_prod` (revenue sheet)
4. `inv_meta` variant title
5. `meta` product name (sales source)

## DRR and demand

| Metric | Calculation |
|--------|-------------|
| DRR | Units Sold (90D) ÷ 90 |
| Projected Demand (60D) | DRR × 60 |
| Days of Cover | Current Stock ÷ DRR |

90-day units come from the revenue sheet qty column. If no qty column is found, MTD units are extrapolated to 90 days.

## Priority logic

SKUs are ranked by **cumulative MRP revenue** (last 90 days):

| Label | Cumulative revenue band | Action |
|-------|------------------------|--------|
| P0 | Top 50% | Replenish immediately |
| P1 | 50 – 70% | Reorder this week |
| P2 | 70 – 90% | Next planning cycle |
| P3 | 90 – 100% + zero revenue | Monitor |

## Stock Status thresholds

| Status | Days of Cover |
|--------|---------------|
| Critical | < 15 days |
| Low | 15 – 30 days |
| Healthy | 30 – 45 days |
| Overstocked | 45 – 60 days |
| Excess | 60+ days |
| No Data | Stock or DRR missing |

## Daily email (n8n)

An n8n workflow reads the B2B and B2C destination tabs and sends a Gmail alert at the same time as the sync. The email shows **P0 and P1 items only**, grouped into Critical / Low / Excess sections with:

- Online WH Stock (B2B) / B2C WH Stock (B2C)
- Mother WH Stock
- DRR, Projected Demand (60D), DOC, Units to Fill

## GitHub Actions

Cron: `30 6 * * *` (06:30 UTC = 12:00 noon IST). The workflow:
1. Checks out the repo
2. Writes the service account key from the `GCP_SERVICE_ACCOUNT_KEY` secret
3. Runs `python3 sync.py`
4. Deletes the key file

The key file (`*.json`) is git-ignored and **never committed**.

## Setup

1. Create a Google Cloud service account and download the JSON key.
2. Share all source and destination sheets with the service account email (Editor).
3. Place the key file in this directory (git-ignored).
4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
5. Run:
   ```bash
   python3 sync.py
   ```

## Configuration

Edit the constants at the top of `sync.py`:

```python
REVENUE_DAYS = 90                          # rolling revenue window
P0_CUT, P1_CUT, P2_CUT = 50, 70, 90      # priority band %
```

## Sheet IDs

| Variable | Sheet |
|----------|-------|
| `SOURCE_ID` | Sales source (Compiled Sec Data Ecomm) |
| `INV_ID` | Inventory / stock (also contains Easyecomm + ERP tabs) |
| `SALES_ID` | Secondary Sale — revenue + category |
| `DEST_ID` | Destination dashboard |
