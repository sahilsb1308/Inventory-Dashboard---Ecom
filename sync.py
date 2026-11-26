"""
Swiss Beauty - Inventory Dashboard Sync
Writes B2B and B2C tabs in the destination sheet with:
  Units Sold (MTD), Current Stock, DRR, Projected Demand (30D),
  Days of Cover, Stock Status, Revenue (Last 90D),
  Revenue Contribution %, Priority

Run:  python3 sync.py
Requires: pip install gspread google-auth
"""

import warnings
warnings.filterwarnings("ignore")

import os
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request as GRequest
from collections import defaultdict, Counter
from datetime import datetime, timedelta

# ── Config ────────────────────────────────────────────────────────────────────
SA_FILE   = "inventory-dashboard-ecommerce-a460b764a150.json"   # keep out of git
SOURCE_ID = "1zIhkkTHcca95mj7jiRUO2e4VYWaG7y7l9mSAfwc8zL0"    # sales source
INV_ID    = "1daV5kSvAf19z0LqZ9PKT2Vbae5rhULmi8qcNUEqAL4I"    # inventory / stock
SALES_ID  = "1Lf5WP1WhznatZjs2fLYzFvZmR4KdyUcnJogf_deE_7U"    # secondary sale (revenue)
DEST_ID   = "1ZP1hJXGrRYsm70J6zBwzCLfv2kVpS1VRoCSb4Wh_tYk"   # destination dashboard

REVENUE_DAYS = 90          # rolling revenue window
TODAY        = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
MTD_DAYS     = TODAY.day   # days elapsed so far this month
REVENUE_FROM = TODAY - timedelta(days=REVENUE_DAYS)

# Priority cumulative revenue bands
P0_CUT, P1_CUT, P2_CUT = 50, 70, 90   # %

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ── Colour palette ────────────────────────────────────────────────────────────
def _rgb(r, g, b): return {"red": r, "green": g, "blue": b}

STOCK_COLOURS = {
    "Critical":    {"bg": _rgb(0.83, 0.10, 0.10), "fg": _rgb(1, 1, 1)},
    "Low":         {"bg": _rgb(1.00, 0.95, 0.20), "fg": _rgb(0, 0, 0)},
    "Healthy":     {"bg": _rgb(0.18, 0.56, 0.18), "fg": _rgb(1, 1, 1)},
    "Overstocked": {"bg": _rgb(0.07, 0.07, 0.35), "fg": _rgb(1, 1, 1)},
    "Excess":      {"bg": _rgb(0.50, 0.10, 0.70), "fg": _rgb(1, 1, 1)},
    "No Data":     {"bg": _rgb(0.75, 0.75, 0.75), "fg": _rgb(0, 0, 0)},
}
PRIORITY_COLOURS = {
    "P0": {"bg": _rgb(0.83, 0.10, 0.10), "fg": _rgb(1, 1, 1)},
    "P1": {"bg": _rgb(1.00, 0.60, 0.00), "fg": _rgb(0, 0, 0)},
    "P2": {"bg": _rgb(1.00, 0.95, 0.20), "fg": _rgb(0, 0, 0)},
    "P3": {"bg": _rgb(0.85, 0.92, 0.80), "fg": _rgb(0, 0, 0)},
}
HEADER_BG  = _rgb(0.17, 0.24, 0.31)
HEADER_FG  = _rgb(1, 1, 1)
ROW_ODD    = _rgb(1.00, 1.00, 1.00)
ROW_EVEN   = _rgb(0.94, 0.96, 0.99)

# ── Helpers ───────────────────────────────────────────────────────────────────
def col_letter(idx):
    s, n = "", idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def stock_status(doc):
    if doc is None: return "No Data"
    if doc < 15:    return "Critical"
    if doc < 30:    return "Low"
    if doc < 45:    return "Healthy"
    if doc < 60:    return "Overstocked"
    return "Excess"

def priority_label(pct, total_rev):
    return "P3"   # placeholder; computed after cumulative sort

def parse_month(m):
    try: return datetime.strptime(m, "%b-%y")
    except: return None

def safe_float(val, floor=None):
    try:
        v = float(str(val).replace(",", "").replace(" ", ""))
        return max(floor, v) if floor is not None else v
    except:
        return None

def is_invalid(val):
    v = str(val).strip()
    return not v or v.startswith("#")

# ── Auth ──────────────────────────────────────────────────────────────────────
print("Authenticating…")
creds = Credentials.from_service_account_file(SA_FILE, scopes=SCOPES)
creds.refresh(GRequest())
gc = gspread.authorize(creds)

# ── 1. Read sales (MTD units) ─────────────────────────────────────────────────
print("Reading sales data…")
src_ws  = gc.open_by_key(SOURCE_ID).worksheet("Compiled Sec Data Ecomm")
src_all = src_ws.get_all_values()
src_hdr = [h.strip().lower() for h in src_all[1]]   # row 2 = header

sku_i  = src_hdr.index("ims code")
qty_i  = src_hdr.index("units sold") if "units sold" in src_hdr else src_hdr.index("qty")
vert_i = src_hdr.index("vertical")
ean_i  = src_hdr.index("ean code")      if "ean code"      in src_hdr else None
brd_i  = src_hdr.index("brand")        if "brand"        in src_hdr else None
cat_i  = src_hdr.index("category")     if "category"     in src_hdr else None
sub_i  = src_hdr.index("sub-category") if "sub-category" in src_hdr else (
         src_hdr.index("sub category")  if "sub category"  in src_hdr else None)
var_i  = src_hdr.index("variant title") if "variant title" in src_hdr else (
         src_hdr.index("variant")        if "variant"        in src_hdr else None)
prd_i  = src_hdr.index("product name")  if "product name"  in src_hdr else (
         src_hdr.index("product")        if "product"        in src_hdr else None)

b2b_qty = defaultdict(int)
b2c_qty = defaultdict(int)
meta    = {}   # ims → {ean, brand, cat, sub}

for row in src_all[2:]:
    if len(row) <= max(sku_i, qty_i, vert_i): continue
    ims  = row[sku_i].strip()
    vert = row[vert_i].strip().upper()
    if is_invalid(ims) or is_invalid(vert): continue
    q = safe_float(row[qty_i])
    if q is None: continue
    q = int(q)
    if vert == "B2B": b2b_qty[ims] += q
    elif vert == "B2C": b2c_qty[ims] += q
    if ims not in meta:
        meta[ims] = {
            "ean":     row[ean_i].strip() if ean_i and len(row) > ean_i else "",
            "brand":   row[brd_i].strip() if brd_i and len(row) > brd_i else "",
            "cat":     row[cat_i].strip() if cat_i and len(row) > cat_i else "",
            "sub":     row[sub_i].strip() if sub_i and len(row) > sub_i else "",
            "variant":      row[var_i].strip() if var_i and len(row) > var_i else "",
            "product_name": row[prd_i].strip() if prd_i and len(row) > prd_i else "",
        }

print(f"  B2B SKUs: {len(b2b_qty)}  |  B2C SKUs: {len(b2c_qty)}")

# ── 2. Read stock + product info from inventory sheet ─────────────────────────
print("Reading inventory / stock…")
inv_ws   = gc.open_by_key(INV_ID).get_worksheet(0)   # first sheet
inv_rows = inv_ws.get_all_values()

stock_b2b = {}   # sku → int
stock_b2c = {}
inv_meta  = {}   # sku → {product_name, variant_title, launch_type}

inv_hdr = [h.strip().lower() for h in inv_rows[0]] if inv_rows else []
inv_cat_i = next((inv_hdr.index(c) for c in ("category",) if c in inv_hdr), None)
inv_sub_i = next((inv_hdr.index(c) for c in ("sub-category", "sub category", "subcategory") if c in inv_hdr), None)

for row in inv_rows[1:]:
    if len(row) < 44: continue
    sku = row[1].strip()   # col B = ERP Codes New MRP
    if not sku or sku.startswith("#"): continue
    stock_b2b[sku] = safe_float(row[27], floor=0) or 0   # col AB (index 27)
    stock_b2c[sku] = safe_float(row[43], floor=0) or 0   # col AR (index 43)
    inv_meta[sku]  = {
        "product_name":  row[3].strip()  if len(row) > 3  else "",   # col D
        "variant_title": row[4].strip()  if len(row) > 4  else "",   # col E
        "launch_type":   row[14].strip() if len(row) > 14 else "",   # col O
        "cat": row[inv_cat_i].strip() if inv_cat_i and len(row) > inv_cat_i else "",
        "sub": row[inv_sub_i].strip() if inv_sub_i and len(row) > inv_sub_i else "",
    }

print(f"  Inventory SKUs: {len(stock_b2b)}")

# ── 3. Read revenue (last 90 days, MRP Sale) ──────────────────────────────────
print("Reading revenue data…")
rev_ws   = gc.open_by_key(SALES_ID).worksheet("Secondary Sale (25 to 27)")
rev_rows = rev_ws.get_all_values()
rev_hdr  = [h.strip().lower() for h in rev_rows[0]]

r_sku_i  = rev_hdr.index("common sku")
r_mrp_i  = rev_hdr.index("mrp sale")
r_mon_i  = rev_hdr.index("month")
r_ver_i  = rev_hdr.index("vertical")

valid_months = {
    m for m in {r[r_mon_i].strip() for r in rev_rows[1:] if len(r) > r_mon_i}
    if (d := parse_month(m)) and d >= datetime(REVENUE_FROM.year, REVENUE_FROM.month, 1)
}

b2b_rev = defaultdict(float)
b2c_rev = defaultdict(float)

for row in rev_rows[1:]:
    if len(row) <= max(r_sku_i, r_mrp_i, r_mon_i, r_ver_i): continue
    sku   = row[r_sku_i].strip()
    month = row[r_mon_i].strip()
    vert  = row[r_ver_i].strip().upper()
    if not sku or month not in valid_months: continue
    rev = safe_float(row[r_mrp_i], floor=0)
    if rev is None: continue
    if vert == "B2B":   b2b_rev[sku] += rev
    elif vert == "B2C": b2c_rev[sku] += rev

print(f"  Revenue months included: {sorted(valid_months)}")
print(f"  B2B revenue SKUs: {len(b2b_rev)}  |  B2C revenue SKUs: {len(b2c_rev)}")

# ── 4. Build destination rows ─────────────────────────────────────────────────
FINAL_COLS = [
    "IMS Code", "EAN Code", "Product Name", "Variant Title",
    "Brand", "Category", "Sub-Category", "Launch Type",
    "Units Sold", "Current Stock", "DRR", "Projected Demand (30D)",
    "Days of Cover", "Stock Status",
    "Revenue (Last 90D)", "Revenue Contribution %", "Priority",
]

def build_rows(qty_map, stock_map, rev_map):
    all_skus = sorted(set(qty_map) | set(stock_map))
    total_rev = sum(max(0, rev_map.get(s, 0)) for s in all_skus)

    # Compute priority via cumulative revenue rank
    ranked = sorted(
        [(s, max(0, rev_map.get(s, 0))) for s in all_skus if rev_map.get(s, 0) > 0],
        key=lambda x: -x[1]
    )
    pri_map = {}
    cum = 0.0
    for s, r in ranked:
        cum += r
        pct = cum / total_rev * 100 if total_rev else 0
        if pct <= P0_CUT:   pri_map[s] = "P0"
        elif pct <= P1_CUT: pri_map[s] = "P1"
        elif pct <= P2_CUT: pri_map[s] = "P2"
        else:               pri_map[s] = "P3"
    for s in all_skus:
        if s not in pri_map: pri_map[s] = "P3"

    rows = []
    for sku in all_skus:
        sold  = qty_map.get(sku, 0)
        stock = stock_map.get(sku, 0)
        drr   = int(round(sold / MTD_DAYS)) if sold else 0
        proj  = int(round(drr * 30)) if drr else 0
        doc   = int(round(stock / drr)) if drr else None
        ss    = stock_status(doc)
        rev   = round(max(0, rev_map.get(sku, 0)), 2)
        pct   = round(rev / total_rev * 100, 2) if total_rev else 0.0
        p     = pri_map[sku]
        m     = meta.get(sku, {})
        im    = inv_meta.get(sku, {})

        variant      = im.get("variant_title", "")  or m.get("variant", "")
        product_name = im.get("product_name", "")   or m.get("product_name", "")

        rows.append([
            sku,
            m.get("ean", ""),
            product_name,
            variant,
            m.get("brand", ""),
            m.get("cat", "") or im.get("cat", ""),
            m.get("sub", "") or im.get("sub", ""),
            im.get("launch_type", ""),
            sold,
            int(stock),
            drr,
            int(proj),
            doc if doc is not None else "",
            ss,
            rev,
            pct,
            p,
        ])
    return rows, total_rev

# ── 5. Write to destination ───────────────────────────────────────────────────
dst = gc.open_by_key(DEST_ID)

def _batch(requests, dst, pause=1.5):
    """Send batchUpdate requests in chunks of 500, with a pause between chunks."""
    for i in range(0, len(requests), 500):
        dst.batch_update({"requests": requests[i:i+500]})
        if i + 500 < len(requests):
            time.sleep(pause)

def write_tab(tab_name, qty_map, stock_map, rev_map):
    print(f"\nWriting {tab_name}…")
    data_rows, total_rev = build_rows(qty_map, stock_map, rev_map)
    ws = dst.worksheet(tab_name)
    ws.clear()
    time.sleep(2)

    all_rows = [FINAL_COLS] + data_rows
    for i in range(0, len(all_rows), 500):
        ws.update(f"A{i+1}", all_rows[i:i+500], value_input_option="RAW")
        time.sleep(1.5)

    sheet_id = ws.id
    n_data   = len(data_rows)
    n_cols   = len(FINAL_COLS)
    ss_i     = FINAL_COLS.index("Stock Status")
    pri_i    = FINAL_COLS.index("Priority")
    rev_ci   = FINAL_COLS.index("Revenue (Last 90D)")
    pct_ci   = FINAL_COLS.index("Revenue Contribution %")
    doc_ci   = FINAL_COLS.index("Days of Cover")

    reqs = []

    # Header row formatting
    reqs.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                   "startColumnIndex": 0, "endColumnIndex": n_cols},
        "cell": {"userEnteredFormat": {
            "backgroundColor": HEADER_BG,
            "textFormat": {"foregroundColor": HEADER_FG, "bold": True},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
    }})

    # Remove any existing banding — fetch IDs and delete via raw HTTP to avoid int32 overflow
    sheet_info = dst.client.request(
        "get",
        f"https://sheets.googleapis.com/v4/spreadsheets/{DEST_ID}",
        params={"fields": "sheets(properties.sheetId,bandedRanges)"}
    ).json()
    del_reqs = []
    for s in sheet_info.get("sheets", []):
        if s.get("properties", {}).get("sheetId") == sheet_id:
            for br in s.get("bandedRanges", []):
                br_id = br["bandedRangeId"]
                # Send each deleteBanding separately via raw HTTP to handle IDs > int32 max
                dst.client.request(
                    "post",
                    f"https://sheets.googleapis.com/v4/spreadsheets/{DEST_ID}:batchUpdate",
                    json={"requests": [{"deleteBanding": {"bandedRangeId": br_id}}]}
                )
                time.sleep(0.5)

    # Alternating row banding — single API call instead of one per row
    reqs.append({"addBanding": {
        "bandedRange": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": n_data + 1,
                       "startColumnIndex": 0, "endColumnIndex": n_cols},
            "rowProperties": {
                "headerColor":      ROW_ODD,
                "firstBandColor":   ROW_ODD,
                "secondBandColor":  ROW_EVEN,
            }
        }
    }})

    # Number formats (range-level)
    reqs.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": n_data+1,
                   "startColumnIndex": doc_ci, "endColumnIndex": doc_ci+1},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0"}}},
        "fields": "userEnteredFormat.numberFormat"
    }})
    reqs.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": n_data+1,
                   "startColumnIndex": rev_ci, "endColumnIndex": rev_ci+1},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}},
        "fields": "userEnteredFormat.numberFormat"
    }})
    reqs.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": n_data+1,
                   "startColumnIndex": pct_ci, "endColumnIndex": pct_ci+1},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.00\"%\""}}},
        "fields": "userEnteredFormat.numberFormat"
    }})

    # Freeze header
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount"
    }})

    _batch(reqs, dst)
    time.sleep(2)

    # Stock Status colours — group consecutive rows with same status to cut calls
    ss_reqs = []
    for r_idx, row in enumerate(data_rows):
        ss_val = row[ss_i]
        if ss_val in STOCK_COLOURS:
            c = STOCK_COLOURS[ss_val]
            # Merge with previous request if same status and adjacent row
            if (ss_reqs and
                ss_reqs[-1]["repeatCell"]["cell"]["userEnteredFormat"]["backgroundColor"] == c["bg"] and
                ss_reqs[-1]["repeatCell"]["range"]["endRowIndex"] == r_idx + 1):
                ss_reqs[-1]["repeatCell"]["range"]["endRowIndex"] = r_idx + 2
            else:
                ss_reqs.append({"repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": r_idx+1, "endRowIndex": r_idx+2,
                               "startColumnIndex": ss_i, "endColumnIndex": ss_i+1},
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": c["bg"],
                        "textFormat": {"foregroundColor": c["fg"], "bold": True},
                        "horizontalAlignment": "CENTER",
                    }},
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }})

    _batch(ss_reqs, dst)
    time.sleep(2)

    # Priority colours — same merge trick
    pri_reqs = []
    for r_idx, row in enumerate(data_rows):
        p_val = row[pri_i]
        if p_val in PRIORITY_COLOURS:
            c = PRIORITY_COLOURS[p_val]
            if (pri_reqs and
                pri_reqs[-1]["repeatCell"]["cell"]["userEnteredFormat"]["backgroundColor"] == c["bg"] and
                pri_reqs[-1]["repeatCell"]["range"]["endRowIndex"] == r_idx + 1):
                pri_reqs[-1]["repeatCell"]["range"]["endRowIndex"] = r_idx + 2
            else:
                pri_reqs.append({"repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": r_idx+1, "endRowIndex": r_idx+2,
                               "startColumnIndex": pri_i, "endColumnIndex": pri_i+1},
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": c["bg"],
                        "textFormat": {"foregroundColor": c["fg"], "bold": True},
                        "horizontalAlignment": "CENTER",
                    }},
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }})

    _batch(pri_reqs, dst)

    counts = Counter(row[pri_i] for row in data_rows)
    print(f"  {tab_name}: {n_data} rows written")
    print(f"  Priority: P0={counts['P0']} P1={counts['P1']} P2={counts['P2']} P3={counts['P3']}")

    # Sanity check: read back row 2 and validate key fields
    if data_rows:
        r2 = data_rows[0]
        r2_dict = dict(zip(FINAL_COLS, r2))
        issues = []
        if not r2_dict.get("IMS Code"):           issues.append("IMS Code blank")
        if not r2_dict.get("Product Name"):        issues.append("Product Name blank")
        if not r2_dict.get("Variant Title"):       issues.append("Variant Title blank")
        if r2_dict.get("Current Stock", 0) < 0:   issues.append(f"Negative stock: {r2_dict['Current Stock']}")
        if r2_dict.get("Priority") not in ("P0","P1","P2","P3"): issues.append(f"Bad priority: {r2_dict.get('Priority')}")
        if issues:
            print(f"  ⚠ Row-2 sanity FAIL [{tab_name}]: {', '.join(issues)}")
            print(f"    Row 2 data: {r2_dict}")
        else:
            print(f"  ✓ Row-2 sanity OK  [{tab_name}]: SKU={r2_dict['IMS Code']} | Stock={r2_dict['Current Stock']} | Priority={r2_dict['Priority']} | Variant={r2_dict['Variant Title'] or '(none)'}")

write_tab("B2B", b2b_qty, stock_b2b, b2b_rev)
write_tab("B2C", b2c_qty, stock_b2c, b2c_rev)

print("\n✓ Sync complete.")
