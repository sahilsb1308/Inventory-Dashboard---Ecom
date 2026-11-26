# Swiss Beauty — Inventory Dashboard (Ecom)

Automated sync script that pulls sales, stock, and revenue data from multiple Google Sheets and writes a clean inventory dashboard with priority scoring.

## What it does

| Step | Source | Data |
|------|--------|------|
| Sales (MTD) | `Compiled Sec Data Ecomm` | Units sold, split by B2B / B2C |
| Stock | Inventory Dashboard sheet | B2B: col AB · B2C: col AR |
| Revenue | `Secondary Sale (25 to 27)` | MRP Sale, last 90 days |

**Output:** `B2B` and `B2C` tabs in the destination sheet with:

`IMS Code · EAN Code · Product Name · Variant Title · Brand · Category · Sub-Category · Launch Type · Units Sold · Current Stock · DRR · Projected Demand (30D) · Days of Cover · Stock Status · Revenue (Last 90D) · Revenue Contribution % · Priority`

## Priority logic

SKUs are ranked by **cumulative MRP revenue** (last 90 days):

| Label | Cumulative revenue band | Action |
|-------|------------------------|--------|
| P0 | Top 50% | Replenish immediately |
| P1 | 50 – 70% | Reorder this week |
| P2 | 70 – 90% | Next planning cycle |
| P3 | 90 – 100% + zero revenue | Monitor |

Revenue Contribution % and Priority use the same denominator (tab total), so selecting all P0 rows sums to exactly 50%.

## Stock Status thresholds

| Status | Days of Cover |
|--------|---------------|
| Critical | 0 – 15 days |
| Low | 15 – 30 days |
| Healthy | 30 – 45 days |
| Overstocked | 45 – 60 days |
| Excess | 60+ days |
| No Data | Stock or DRR missing |

## Setup

1. Create a Google Cloud service account and download the JSON key.
2. Share all four Google Sheets with the service account email (Editor access).
3. Place the key file in this directory (it is git-ignored).
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
MTD_DAYS     = 26          # days elapsed in current month
REVENUE_DAYS = 90          # rolling revenue window
TODAY        = datetime(2026, 11, 26)
P0_CUT, P1_CUT, P2_CUT = 50, 70, 90   # priority band %
```

## Sheet IDs

| Variable | Sheet |
|----------|-------|
| `SOURCE_ID` | Sales source (Compiled Sec Data Ecomm) |
| `INV_ID` | Inventory / stock |
| `SALES_ID` | Secondary Sale (revenue) |
| `DEST_ID` | Destination dashboard |
