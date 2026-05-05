r"""
=============================================================================
  Honeywell Program Dashboard Generator  -  v2 (Enhanced)
=============================================================================
  Source workbooks:
    Working Honeywell KPI Dashboard.xlsx   ->  'SAW Report Data for Current Day'
    Control_Warehouse_Inventory.xlsx       ->  inventory / ATP data

  HOW TO USE:
    1. One-time setup  ->  pip install pandas openpyxl
    2. Set EXCEL_PATH and INVENTORY_PATH below (or pass as CLI args)
    3. python generate_dashboard_v2.py
    4. Drop the output index.html into your GitLab repo root - done.
=============================================================================
"""

import json
import sys
from datetime import date
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("\nERROR: pandas not installed.\nFix: pip install pandas openpyxl\n")

# =============================================================================
#  CONFIG
# =============================================================================
EXCEL_PATH = Path(
    r"C:\Users\zn424f\OneDrive - The Boeing Company"
    r"\Working KPIs\Honeywell Dashboard (GL)\Working Honeywell KPI Dashboard.xlsx"
)
INVENTORY_PATH = Path(
    r"C:\Users\zn424f\OneDrive - The Boeing Company"
    r"\Working KPIs\Honeywell Dashboard (GL)\Control Warehouse Inventory.xlsx"
)
SHEET_NAME  = "SAW Report Data for Current Day"
OUTPUT_FILE = Path(__file__).resolve().parent / "index.html"

# Plant codes
PLANT_CHANDLER = 8003
PLANT_MIAMI    = 8000

# Month abbreviation -> number
MONTH_MAP = {
    "Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
    "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12
}


# =============================================================================
#  STEP 1 - LOAD DATA
# =============================================================================
def load_data(excel_path, inv_path):
    if not excel_path.exists():
        sys.exit(f"\nERROR: Cannot find workbook:\n  {excel_path}\n")
    if not inv_path.exists():
        sys.exit(f"\nERROR: Cannot find inventory file:\n  {inv_path}\n")

    print(f"Reading order book: {excel_path.name} [{SHEET_NAME}] ...")
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    print(f"  {len(df):,} rows x {len(df.columns)} columns")

    print(f"Reading inventory: {inv_path.name} ...")
    inv = pd.read_excel(inv_path)
    print(f"  {len(inv):,} inventory records")

    # Coerce numeric columns
    for col in ["Extended Price", "Confirmed Qty.", "Open Qty"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Ship Request Date"] = pd.to_datetime(df["Ship Request Date"], errors="coerce")
    inv["ATP Available Quantity"] = pd.to_numeric(inv["ATP Available Quantity"], errors="coerce").fillna(0)
    inv["Inventory On Hand"]      = pd.to_numeric(inv["Inventory On Hand"],      errors="coerce").fillna(0)

    return df, inv


# =============================================================================
#  STEP 2 - BUCKET STATUS
# =============================================================================
def assign_buckets(df):
    today   = pd.Timestamp(date.today())
    cur_yr  = today.year
    cur_mo  = today.month
    cur_q   = (cur_mo - 1) // 3 + 1
    yr_end  = pd.Timestamp(f"{cur_yr}-12-31")

    def classify(status):
        s = str(status).strip()
        if s == "Past Due": return "past_due"
        if s == "Today":    return "today"
        if s == "Future":   return "future"
        # parse "Mon YYYY"
        parts = s.split()
        if len(parts) == 2 and parts[0] in MONTH_MAP:
            mo = MONTH_MAP[parts[0]]
            yr = int(parts[1])
            if yr == cur_yr and mo == cur_mo:
                return "this_month"
            if yr == cur_yr:
                mo_q = (mo - 1) // 3 + 1
                if mo_q == cur_q:
                    return "this_quarter"
                if mo <= 12:
                    return "rest_of_year" if yr == cur_yr and mo > cur_mo else "future"
            return "future"
        return "future"

    df = df.copy()
    df["_bucket"] = df["Status"].apply(classify)

    # Refine: this_quarter includes this_month (it's a superset), but keep them separate for display
    # next_90: use ship date for rows not past_due/today
    next_90 = today + pd.Timedelta(days=90)
    df["_due_90"] = (
        ~df["_bucket"].isin(["past_due", "today"]) &
        df["Ship Request Date"].notna() &
        (df["Ship Request Date"] >= today) &
        (df["Ship Request Date"] <= next_90)
    )

    return df, today


# =============================================================================
#  STEP 3 - INVENTORY ATP ANALYSIS (PAST DUE ONLY)
# =============================================================================
def build_inventory_analysis(df, inv):
    # Aggregate ATP by part + plant
    inv_agg = (
        inv.groupby(["Part Number", "Plant"])
           .agg(atp=("ATP Available Quantity", "sum"),
                onhand=("Inventory On Hand", "sum"))
           .reset_index()
    )
    miami    = inv_agg[inv_agg["Plant"] == PLANT_MIAMI].set_index("Part Number")
    chandler = inv_agg[inv_agg["Plant"] == PLANT_CHANDLER].set_index("Part Number")

    pd_df = df[df["_bucket"] == "past_due"].copy()
    pd_df["miami_atp"]    = pd_df["Part Number"].map(miami["atp"]).fillna(0)
    pd_df["chandler_atp"] = pd_df["Part Number"].map(chandler["atp"]).fillna(0)
    pd_df["total_atp"]    = pd_df["miami_atp"] + pd_df["chandler_atp"]
    pd_df["miami_oh"]     = pd_df["Part Number"].map(miami["onhand"]).fillna(0)
    pd_df["chandler_oh"]  = pd_df["Part Number"].map(chandler["onhand"]).fillna(0)

    def greedy_fill(orders, atp_pool):
        """Fill complete orders biggest-first using available ATP pool."""
        filled, partial, unfilled = [], [], []
        remaining = float(atp_pool)
        for _, row in orders.sort_values("Confirmed Qty.", ascending=False).iterrows():
            qty = float(row["Confirmed Qty."])
            if remaining >= qty and qty > 0:
                filled.append(row)
                remaining -= qty
            elif remaining > 0 and qty > 0:
                partial.append(row)
                # Do NOT consume partial fills per logic request (biggest-order-first full fills only)
            else:
                unfilled.append(row)
        return filled, partial, unfilled, remaining

    # Summarize fill analysis by part/PO-Bin
    fill_rows = []
    for (pn, pob), grp in pd_df.groupby(["Part Number", "PO/Bin"]):
        atp = float(grp["total_atp"].iloc[0])
        miami_atp = float(grp["miami_atp"].iloc[0])
        chand_atp = float(grp["chandler_atp"].iloc[0])
        filled, partial, unfilled, remaining = greedy_fill(grp, atp)
        total_orders = len(grp)
        total_qty    = float(grp["Confirmed Qty."].sum())
        filled_qty   = sum(float(r["Confirmed Qty."]) for r in filled)
        filled_val   = sum(float(r["Extended Price"]) for r in filled)
        unfilled_val = sum(float(r["Extended Price"]) for r in unfilled) + sum(float(r["Extended Price"]) for r in partial)

        fill_rows.append({
            "Part Number":    pn,
            "PO/Bin":         pob,
            "total_orders":   total_orders,
            "total_qty":      round(total_qty, 0),
            "atp_miami":      round(miami_atp, 0),
            "atp_chandler":   round(chand_atp, 0),
            "atp_total":      round(atp, 0),
            "orders_fillable": len(filled),
            "filled_qty":     round(filled_qty, 0),
            "filled_value":   round(filled_val, 2),
            "unfilled_orders": len(unfilled) + len(partial),
            "unfilled_value": round(unfilled_val, 2),
            "fill_rate_orders": round(len(filled) / total_orders * 100, 1) if total_orders else 0,
        })

    fill_df = pd.DataFrame(fill_rows).sort_values("unfilled_value", ascending=False)

    # Summary stats
    total_pd_val    = pd_df["Extended Price"].sum()
    has_atp         = pd_df[pd_df["total_atp"] > 0]
    no_atp          = pd_df[pd_df["total_atp"] == 0]
    fillable_val    = fill_df["filled_value"].sum()
    unfillable_val  = fill_df["unfilled_value"].sum()

    inv_summary = {
        "pd_lines":           len(pd_df),
        "pd_value":           round(total_pd_val, 2),
        "pd_with_atp_lines":  len(has_atp),
        "pd_with_atp_value":  round(has_atp["Extended Price"].sum(), 2),
        "pd_no_atp_lines":    len(no_atp),
        "pd_no_atp_value":    round(no_atp["Extended Price"].sum(), 2),
        "fillable_orders":    int(fill_df["orders_fillable"].sum()),
        "fillable_value":     round(fillable_val, 2),
        "unfillable_value":   round(unfillable_val, 2),
        "miami_atp_parts":    int((pd_df["miami_atp"] > 0).sum()),
        "chandler_atp_parts": int((pd_df["chandler_atp"] > 0).sum()),
    }

    return fill_df.to_dict("records"), inv_summary


# =============================================================================
#  STEP 4 - CALCULATE ALL METRICS
# =============================================================================
def calculate(df, inv):
    df, today = assign_buckets(df)
    total_val = df["Extended Price"].sum()

    def bucket_agg(mask):
        sub = df[mask]
        return {
            "lines": int(len(sub)),
            "value": round(float(sub["Extended Price"].sum()), 2),
            "pct":   round(float(sub["Extended Price"].sum()) / total_val * 100, 2) if total_val else 0,
        }

    is_pd    = df["_bucket"] == "past_due"
    is_today = df["_bucket"] == "today"
    is_mo    = df["_bucket"] == "this_month"
    is_q     = df["_bucket"].isin(["this_quarter", "this_month"])  # this quarter includes this month
    is_90    = df["_due_90"]
    is_yr    = df["_bucket"] == "rest_of_year"
    is_fut   = df["_bucket"] == "future"

    kpis = {
        "total_book_value": round(float(total_val), 2),
        "total_lines":      int(len(df)),
        "as_of_date":       today.strftime("%b %#d, %Y") if sys.platform == "win32" else today.strftime("%b %-d, %Y"),
        "past_due":         bucket_agg(is_pd),
        "today":            bucket_agg(is_today),
        "this_month":       bucket_agg(is_mo),
        "this_quarter":     bucket_agg(is_q),
        "next_90":          bucket_agg(is_90),
        "rest_of_year":     bucket_agg(is_yr),
        "future":           bucket_agg(is_fut),
    }

    # --- Past Due by Site (Customer Name = col AE) ---
    pd_df = df[is_pd].copy()
    pd_site = (
        pd_df.groupby("Customer Name")
             .agg(lines=("SO Number","count"), value=("Extended Price","sum"))
             .reset_index()
             .sort_values("value", ascending=False)
    )
    pd_site["pct"] = (pd_site["value"] / pd_df["Extended Price"].sum() * 100).round(2)
    pd_site["value"] = pd_site["value"].round(2)
    pd_site_data = pd_site.to_dict("records")

    # --- PO vs Bin breakdown (full book + past due) ---
    def pobin_summary(frame):
        rows = []
        for pob, grp in frame.groupby("PO/Bin"):
            rows.append({
                "type":  pob,
                "lines": int(len(grp)),
                "value": round(float(grp["Extended Price"].sum()), 2),
                "pct":   round(float(grp["Extended Price"].sum()) / total_val * 100, 2),
            })
        return sorted(rows, key=lambda x: x["value"], reverse=True)

    pobin_all   = pobin_summary(df)
    pobin_pd    = pobin_summary(pd_df)

    # --- Function breakdown (ActionBy - New = col AU) ---
    func_all = (
        df.groupby("ActionBy - New")
          .agg(lines=("SO Number","count"), value=("Extended Price","sum"))
          .reset_index()
    )
    func_all["pct"]   = (func_all["value"] / total_val * 100).round(2)
    func_all["value"] = func_all["value"].round(2)

    # Function breakdown also sliced by past due
    func_pd = (
        pd_df.groupby("ActionBy - New")
             .agg(pd_lines=("SO Number","count"), pd_value=("Extended Price","sum"))
             .reset_index()
    )
    func_merged = func_all.merge(func_pd, on="ActionBy - New", how="left").fillna(0)
    func_merged["pd_pct"] = (func_merged["pd_value"] / func_merged["value"] * 100).round(1)
    func_data = func_merged.sort_values("value", ascending=False).to_dict("records")

    # --- Site breakdown full book (Whse Site = col CH) ---
    site_all = (
        df.groupby("Whse Site")
          .agg(lines=("SO Number","count"), value=("Extended Price","sum"))
          .reset_index()
    )
    site_all["pct"]   = (site_all["value"] / total_val * 100).round(2)
    site_all["value"] = site_all["value"].round(2)
    site_data = site_all.sort_values("value", ascending=False).to_dict("records")

    # --- Execution Person detail ---
    exec_rows = []
    for name, grp in df.groupby("Execution Person"):
        exec_rows.append({
            "name":       name,
            "lines":      int(len(grp)),
            "value":      round(float(grp["Extended Price"].sum()), 2),
            "pd_value":   round(float(grp.loc[is_pd.loc[grp.index], "Extended Price"].sum()), 2) if is_pd.loc[grp.index].any() else 0,
            "today_val":  round(float(grp.loc[is_today.loc[grp.index], "Extended Price"].sum()), 2) if is_today.loc[grp.index].any() else 0,
            "next90_val": round(float(grp.loc[is_90.loc[grp.index], "Extended Price"].sum()), 2) if is_90.loc[grp.index].any() else 0,
            "pct":        round(float(grp["Extended Price"].sum()) / total_val * 100, 2),
        })
    exec_data = sorted(exec_rows, key=lambda x: x["value"], reverse=True)

    # --- Rep view (P&E / SPS RISK) ---
    risk_df = df[df["ActionBy - New"].isin(["P&E", "SPS RISK"])].copy()
    totals  = df.groupby("Execution Person").agg(
        total_lines=("SO Number","count"), total_value=("Extended Price","sum")
    ).to_dict("index")

    rep_rows = []
    for (exec_name, action), grp in risk_df.groupby(["Execution Person", "ActionBy - New"]):
        tv   = float(totals.get(exec_name, {}).get("total_value", 1) or 1)
        tl   = int(totals.get(exec_name, {}).get("total_lines", 0))
        g_pd = is_pd.loc[grp.index]
        g_td = is_today.loc[grp.index]
        g_90 = is_90.loc[grp.index]
        g_ft = is_fut.loc[grp.index]
        gv   = float(grp["Extended Price"].sum())
        rep_rows.append({
            "exec":          exec_name,
            "action":        action,
            "so_count":      int(grp["SO Number"].nunique()),
            "lines":         int(len(grp)),
            "value":         round(gv, 2),
            "pd_value":      round(float(grp.loc[g_pd, "Extended Price"].sum()), 2),
            "today_value":   round(float(grp.loc[g_td, "Extended Price"].sum()), 2),
            "next90_value":  round(float(grp.loc[g_90, "Extended Price"].sum()), 2),
            "future_value":  round(float(grp.loc[g_ft, "Extended Price"].sum()), 2),
            "pd_lines":      int(g_pd.sum()),
            "total_lines":   tl,
            "total_value":   round(tv, 2),
            "pct_exec":      round(gv / tv * 100, 2),
        })
    rep_data = sorted(rep_rows, key=lambda x: x["value"], reverse=True)

    # --- Inventory / ATP analysis (past due only) ---
    fill_data, inv_summary = build_inventory_analysis(df, inv)

    return {
        "kpis":         kpis,
        "pd_site_data": pd_site_data,
        "pobin_all":    pobin_all,
        "pobin_pd":     pobin_pd,
        "func_data":    func_data,
        "site_data":    site_data,
        "exec_data":    exec_data,
        "rep_data":     rep_data,
        "fill_data":    fill_data[:200],   # top 200 for display
        "inv_summary":  inv_summary,
    }, today


# =============================================================================
#  STEP 5 - HTML TEMPLATE
# =============================================================================
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Honeywell Program Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
:root{
  --bg:#08090d;--s1:#0e1118;--s2:#141720;--s3:#1c2030;--border:#242840;
  --accent:#3d6bff;--accent2:#22d3a0;--red:#ff4757;--amber:#ffb340;--purple:#9b6dff;
  --green:#22d3a0;--text:#dde3f0;--muted:#6b7799;--mono:'JetBrains Mono',monospace;
  --sans:'DM Sans',system-ui,sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;min-height:100vh}
/* NAV */
nav{background:var(--s1);border-bottom:1px solid var(--border);padding:0 28px;display:flex;align-items:center;height:54px;position:sticky;top:0;z-index:200;gap:4px}
.brand{font-weight:700;font-size:14px;letter-spacing:.3px;color:var(--text);margin-right:20px;display:flex;align-items:center;gap:10px;white-space:nowrap}
.brand-dot{width:8px;height:8px;background:var(--accent);border-radius:2px;transform:rotate(45deg)}
.tab{padding:0 14px;height:54px;display:flex;align-items:center;border-bottom:2px solid transparent;cursor:pointer;color:var(--muted);font-size:13px;font-weight:500;transition:all .15s;white-space:nowrap;user-select:none}
.tab:hover{color:var(--text)}
.tab.active{color:var(--text);border-bottom-color:var(--accent)}
.nav-right{margin-left:auto;font-family:var(--mono);font-size:11px;color:var(--muted)}
/* PAGES */
.page{display:none;padding:24px;max-width:1440px;margin:0 auto}
.page.active{display:block}
/* KPI GRID */
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:22px}
.kpi{background:var(--s1);border:1px solid var(--border);border-radius:12px;padding:18px 20px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--kpi-accent,var(--accent))}
.kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);margin-bottom:10px;font-weight:600}
.kpi-val{font-size:24px;font-weight:700;color:var(--kpi-accent,var(--text));font-family:var(--mono)}
.kpi-sub{font-size:11px;color:var(--muted);margin-top:5px}
/* CARDS/SECTIONS */
.card{background:var(--s1);border:1px solid var(--border);border-radius:12px;padding:20px;margin-bottom:18px}
.card-title{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin-bottom:16px;display:flex;align-items:center;gap:10px}
.pill{background:var(--s3);color:var(--muted);font-size:10px;padding:2px 8px;border-radius:20px;font-weight:600;letter-spacing:.3px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;margin-bottom:18px}
@media(max-width:1000px){.row2,.row3{grid-template-columns:1fr}}
.chart-wrap{position:relative;height:260px}
.chart-wrap.tall{height:340px}
/* TABLE */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12.5px}
thead th{text-align:left;font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;padding:8px 12px;border-bottom:1px solid var(--border);background:var(--s2);white-space:nowrap}
tbody tr{border-bottom:1px solid var(--border);transition:background .1s}
tbody tr:hover{background:var(--s2)}
td{padding:8px 12px;white-space:nowrap}
.num{font-family:var(--mono);font-size:12px}
/* BADGES */
.badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;letter-spacing:.3px}
.badge.red{background:#ff475720;color:var(--red)}
.badge.green{background:#22d3a020;color:var(--green)}
.badge.amber{background:#ffb34020;color:var(--amber)}
.badge.blue{background:#3d6bff20;color:var(--accent)}
.badge.purple{background:#9b6dff20;color:var(--purple)}
/* BAR PROGRESS */
.bar-bg{background:var(--s3);border-radius:4px;height:5px;flex:1}
.bar-fill{height:5px;border-radius:4px}
.prog-row{display:flex;align-items:center;gap:8px}
.prog-pct{font-family:var(--mono);font-size:10px;color:var(--muted);width:34px;text-align:right}
/* BUCKET GRID (order book) */
.bucket-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:18px}
.bucket{background:var(--s2);border:1px solid var(--border);border-radius:10px;padding:14px 16px}
.bucket-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:8px}
.bucket-val{font-size:20px;font-weight:700;font-family:var(--mono);color:var(--bkt-col,var(--text))}
.bucket-sub{font-size:11px;color:var(--muted);margin-top:3px}
/* FILL TABLE colors */
.fill-can{color:var(--green)}
.fill-no{color:var(--red)}
/* REP CARDS */
.rep-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:14px}
.rep-card{background:var(--s2);border:1px solid var(--border);border-radius:10px;padding:16px}
.rep-name{font-weight:700;font-size:14px;margin-bottom:3px}
.rep-meta{font-size:11px;color:var(--muted);margin-bottom:12px}
.risk-block{border:1px solid var(--border);border-radius:8px;padding:12px;margin-bottom:8px}
.risk-block:last-child{margin:0}
.risk-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.gauge-row{display:flex;align-items:center;gap:6px;margin-top:5px}
.gauge-lbl{font-size:10px;color:var(--muted);width:72px;flex-shrink:0}
.gauge-bar{flex:1;height:6px;background:var(--s1);border-radius:3px;overflow:hidden}
.gauge-fill{height:100%;border-radius:3px}
.gauge-pct{font-size:10px;font-family:var(--mono);width:34px;text-align:right}
/* FILTER BUTTONS */
.filter-bar{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center}
.filter-btn{padding:5px 14px;border-radius:20px;border:1px solid var(--border);background:var(--s2);color:var(--muted);font-size:11px;cursor:pointer;transition:all .15s;font-weight:600;letter-spacing:.2px}
.filter-btn:hover,.filter-btn.on{background:var(--accent);color:#fff;border-color:var(--accent)}
/* LEGEND */
.legend{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}
.legend-item{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--muted)}
.legend-dot{width:8px;height:8px;border-radius:50%}
/* SECTIONS for inventory */
.inv-stat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:16px}
.inv-stat{background:var(--s3);border-radius:8px;padding:14px 16px}
.inv-stat-val{font-size:22px;font-weight:700;font-family:var(--mono);margin-bottom:3px}
.inv-stat-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-weight:600}
</style>
</head>
<body>
<nav>
  <div class="brand"><div class="brand-dot"></div>Honeywell Program Dashboard</div>
  <div class="tab active" onclick="show('overview',this)">Overview</div>
  <div class="tab" onclick="show('pastdue',this)">Past Due</div>
  <div class="tab" onclick="show('orderbook',this)">Order Book</div>
  <div class="tab" onclick="show('functions',this)">Functions</div>
  <div class="tab" onclick="show('sites',this)">Sites</div>
  <div class="tab" onclick="show('inventory',this)">Inventory / ATP</div>
  <div class="tab" onclick="show('execution',this)">Execution</div>
  <div class="tab" onclick="show('reps',this)">Rep View</div>
  <div class="nav-right">As of __AS_OF__</div>
</nav>

<!-- ============================================================ OVERVIEW -->
<div id="tab-overview" class="page active">
  <div class="kpi-grid">
    <div class="kpi" style="--kpi-accent:var(--accent)">
      <div class="kpi-label">Total Order Book</div>
      <div class="kpi-val" id="kv-book"></div>
      <div class="kpi-sub" id="ks-book"></div>
    </div>
    <div class="kpi" style="--kpi-accent:var(--red)">
      <div class="kpi-label">Past Due</div>
      <div class="kpi-val" id="kv-pd"></div>
      <div class="kpi-sub" id="ks-pd"></div>
    </div>
    <div class="kpi" style="--kpi-accent:var(--amber)">
      <div class="kpi-label">Due Today</div>
      <div class="kpi-val" id="kv-td"></div>
      <div class="kpi-sub" id="ks-td"></div>
    </div>
    <div class="kpi" style="--kpi-accent:var(--purple)">
      <div class="kpi-label">This Month</div>
      <div class="kpi-val" id="kv-mo"></div>
      <div class="kpi-sub" id="ks-mo"></div>
    </div>
    <div class="kpi" style="--kpi-accent:#06b6d4">
      <div class="kpi-label">This Quarter</div>
      <div class="kpi-val" id="kv-q"></div>
      <div class="kpi-sub" id="ks-q"></div>
    </div>
    <div class="kpi" style="--kpi-accent:var(--green)">
      <div class="kpi-label">Rest of Year</div>
      <div class="kpi-val" id="kv-yr"></div>
      <div class="kpi-sub" id="ks-yr"></div>
    </div>
    <div class="kpi" style="--kpi-accent:var(--muted)">
      <div class="kpi-label">Future (2027+)</div>
      <div class="kpi-val" id="kv-fut"></div>
      <div class="kpi-sub" id="ks-fut"></div>
    </div>
  </div>

  <div class="row2">
    <div class="card">
      <div class="card-title">Order Book — Due Date Buckets <span class="pill">$</span></div>
      <div class="chart-wrap"><canvas id="c-bucket"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">PO vs Binstock — Full Book</div>
      <div class="chart-wrap"><canvas id="c-pobin-all"></canvas></div>
    </div>
  </div>
  <div class="row2">
    <div class="card">
      <div class="card-title">Financial Impact by Function <span class="pill">Full Book</span></div>
      <div class="chart-wrap"><canvas id="c-func-pie"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Top Sites by Warehouse Value</div>
      <div class="chart-wrap"><canvas id="c-site-bar"></canvas></div>
    </div>
  </div>
</div>

<!-- ============================================================ PAST DUE -->
<div id="tab-pastdue" class="page">
  <div class="kpi-grid">
    <div class="kpi" style="--kpi-accent:var(--red)">
      <div class="kpi-label">Past Due Value</div>
      <div class="kpi-val" id="kv-pd2"></div>
      <div class="kpi-sub" id="ks-pd2"></div>
    </div>
    <div class="kpi" style="--kpi-accent:var(--amber)">
      <div class="kpi-label">PO Lines Past Due</div>
      <div class="kpi-val" id="kv-pd-po"></div>
      <div class="kpi-sub" id="ks-pd-po"></div>
    </div>
    <div class="kpi" style="--kpi-accent:var(--purple)">
      <div class="kpi-label">Binstock Past Due</div>
      <div class="kpi-val" id="kv-pd-bin"></div>
      <div class="kpi-sub" id="ks-pd-bin"></div>
    </div>
  </div>

  <div class="row2">
    <div class="card">
      <div class="card-title">Past Due by Honeywell Site <span class="pill">Line Items</span></div>
      <div class="chart-wrap tall"><canvas id="c-pdsite-lines"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Past Due by Honeywell Site <span class="pill">$ Value</span></div>
      <div class="chart-wrap tall"><canvas id="c-pdsite-val"></canvas></div>
    </div>
  </div>

  <div class="card">
    <div class="card-title">Past Due — PO vs Binstock Breakout</div>
    <div class="row2" style="margin-bottom:0">
      <div>
        <div class="chart-wrap"><canvas id="c-pobin-pd"></canvas></div>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>Type</th><th>Lines</th><th>Value ($)</th><th>% of PD Book</th><th style="width:140px">Share</th></tr></thead>
          <tbody id="tb-pobin-pd"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-title">Past Due — Site Detail Table</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>#</th><th>Customer Site</th><th>Lines</th><th>Value ($)</th><th>% of PD</th><th style="width:160px">Share</th></tr></thead>
        <tbody id="tb-pdsite"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ============================================================ ORDER BOOK -->
<div id="tab-orderbook" class="page">
  <div class="bucket-grid">
    <div class="bucket" style="--bkt-col:var(--red)"><div class="bucket-label">Past Due</div><div class="bucket-val" id="bv-pd"></div><div class="bucket-sub" id="bs-pd"></div></div>
    <div class="bucket" style="--bkt-col:var(--amber)"><div class="bucket-label">Due Today</div><div class="bucket-val" id="bv-td"></div><div class="bucket-sub" id="bs-td"></div></div>
    <div class="bucket" style="--bkt-col:var(--purple)"><div class="bucket-label">This Month</div><div class="bucket-val" id="bv-mo"></div><div class="bucket-sub" id="bs-mo"></div></div>
    <div class="bucket" style="--bkt-col:#06b6d4"><div class="bucket-label">This Quarter</div><div class="bucket-val" id="bv-q"></div><div class="bucket-sub" id="bs-q"></div></div>
    <div class="bucket" style="--bkt-col:var(--accent)"><div class="bucket-label">Next 90 Days</div><div class="bucket-val" id="bv-90"></div><div class="bucket-sub" id="bs-90"></div></div>
    <div class="bucket" style="--bkt-col:var(--green)"><div class="bucket-label">Rest of Year</div><div class="bucket-val" id="bv-yr"></div><div class="bucket-sub" id="bs-yr"></div></div>
    <div class="bucket" style="--bkt-col:var(--muted)"><div class="bucket-label">Future (2027+)</div><div class="bucket-val" id="bv-fut"></div><div class="bucket-sub" id="bs-fut"></div></div>
  </div>
  <div class="row2">
    <div class="card">
      <div class="card-title">Value by Due Date Bucket</div>
      <div class="chart-wrap"><canvas id="c-ob-bucket"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Lines by Due Date Bucket</div>
      <div class="chart-wrap"><canvas id="c-ob-lines"></canvas></div>
    </div>
  </div>
</div>

<!-- ============================================================ FUNCTIONS -->
<div id="tab-functions" class="page">
  <div class="row2">
    <div class="card">
      <div class="card-title">Financial Impact by Function <span class="pill">Full Book $</span></div>
      <div class="chart-wrap tall"><canvas id="c-func-val"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Line Items by Function <span class="pill">Count</span></div>
      <div class="chart-wrap tall"><canvas id="c-func-lines"></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="card-title">Function Detail — Full Book vs Past Due</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>Function</th><th>Lines</th><th>Value ($)</th><th>% Book</th><th>PD Lines</th><th>PD Value ($)</th><th>PD %</th><th style="width:140px">Share</th></tr></thead>
        <tbody id="tb-func"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ============================================================ SITES -->
<div id="tab-sites" class="page">
  <div class="row2">
    <div class="card">
      <div class="card-title">Top Sites by Warehouse Value <span class="pill">Full Book</span></div>
      <div class="chart-wrap tall"><canvas id="c-site-full"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Site Share (Donut)</div>
      <div style="height:220px;position:relative"><canvas id="c-site-donut"></canvas></div>
      <div class="legend" id="site-legend"></div>
    </div>
  </div>
  <div class="card">
    <div class="card-title">All Sites Detail</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>#</th><th>Warehouse Site</th><th>Lines</th><th>Value ($)</th><th>% of Book</th><th style="width:160px">Share</th></tr></thead>
        <tbody id="tb-sites"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ============================================================ INVENTORY / ATP -->
<div id="tab-inventory" class="page">
  <div class="inv-stat-grid" id="inv-stat-grid"></div>

  <div class="row2">
    <div class="card">
      <div class="card-title">Past Due Fill Coverage <span class="pill">by $ Value</span></div>
      <div class="chart-wrap"><canvas id="c-fill-pie"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Miami (8000) vs Chandler (8003) ATP Coverage</div>
      <div class="chart-wrap"><canvas id="c-plant-atp"></canvas></div>
    </div>
  </div>

  <div class="card">
    <div class="card-title">Inventory Fill Analysis — Past Due Orders
      <span class="pill">Biggest-Order-First · Full Fill Only</span>
      <div class="filter-bar" style="display:inline-flex;margin-left:12px;margin-bottom:0">
        <button class="filter-btn on" onclick="filterFill('all',this)">All</button>
        <button class="filter-btn" onclick="filterFill('PO',this)">PO Only</button>
        <button class="filter-btn" onclick="filterFill('Binstock',this)">Bin Only</button>
        <button class="filter-btn" onclick="filterFill('can',this)">Fillable</button>
        <button class="filter-btn" onclick="filterFill('no',this)">No Stock</button>
      </div>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th>Part Number</th><th>Type</th><th>Open Orders</th><th>Total Qty</th>
            <th>ATP Miami</th><th>ATP Chandler</th><th>ATP Total</th>
            <th>Fillable Orders</th><th>Fillable Value ($)</th>
            <th>Unfillable Orders</th><th>Unfillable Value ($)</th><th>Fill Rate</th>
          </tr>
        </thead>
        <tbody id="tb-fill"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ============================================================ EXECUTION -->
<div id="tab-execution" class="page">
  <div class="card">
    <div class="card-title">Execution Person — Backlog Snapshot <span class="pill">Top 12</span></div>
    <div class="chart-wrap tall"><canvas id="c-exec"></canvas></div>
  </div>
  <div class="card">
    <div class="card-title">Execution Person Detail</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>Person</th><th>Lines</th><th>Total Value ($)</th><th>Past Due ($)</th><th>Due Today ($)</th><th>Next 90d ($)</th><th>% of Book</th><th style="width:130px">Share</th></tr></thead>
        <tbody id="tb-exec"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ============================================================ REP VIEW -->
<div id="tab-reps" class="page">
  <div class="card" style="margin-bottom:14px;padding:14px 20px">
    <div class="card-title" style="margin-bottom:6px">P&amp;E and SPS Risk — Execution Scorecard</div>
    <p style="font-size:12px;color:var(--muted)">Filtered to rows where <strong style="color:#7b9fff">ActionBy = P&amp;E</strong> or <strong style="color:var(--red)">SPS Risk</strong>, grouped by Execution Person. Gauges show time-bucket exposure within each risk category.</p>
  </div>
  <div class="filter-bar">
    <button class="filter-btn on" onclick="filterReps('all',this)">All</button>
    <button class="filter-btn" onclick="filterReps('P&E',this)">P&amp;E Only</button>
    <button class="filter-btn" onclick="filterReps('SPS RISK',this)">SPS Risk Only</button>
  </div>
  <div class="rep-grid" id="rep-grid"></div>
</div>

<script>
const D = __DATA_JSON__;

// ---- helpers ----
const COLS=['#3d6bff','#22d3a0','#ff4757','#ffb340','#9b6dff','#06b6d4','#ec4899','#84cc16','#f97316','#8b5cf6','#fb923c','#14b8a6','#f43f5e','#a78bfa','#34d399'];
const f  = n=>n>=1e6?'$'+(n/1e6).toFixed(2)+'M':n>=1e3?'$'+(n/1e3).toFixed(1)+'K':'$'+n.toFixed(0);
const ff = n=>'$'+n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const fn = n=>n.toLocaleString('en-US');
const fp = n=>n.toFixed(1)+'%';
const pct=(a,b)=>b?((a/b)*100).toFixed(1):'0.0';

function barCell(val,max,col='var(--accent)'){
  const w=Math.max((val/max)*100,.4);
  return `<div class="prog-row"><div class="bar-bg"><div class="bar-fill" style="width:${w}%;background:${col}"></div></div><div class="prog-pct">${(val/max*100).toFixed(1)}%</div></div>`;
}

// ---- KPI fill ----
function fillKPI(){
  const K=D.kpis;
  const sets=[
    ['kv-book','ks-book',K.total_book_value,K.total_lines+' line items'],
    ['kv-pd','ks-pd',K.past_due.value,K.past_due.lines+' lines · '+K.past_due.pct+'% of book'],
    ['kv-td','ks-td',K.today.value,K.today.lines+' lines · '+K.today.pct+'% of book'],
    ['kv-mo','ks-mo',K.this_month.value,K.this_month.lines+' lines · '+K.this_month.pct+'% of book'],
    ['kv-q','ks-q',K.this_quarter.value,K.this_quarter.lines+' lines · '+K.this_quarter.pct+'% of book'],
    ['kv-yr','ks-yr',K.rest_of_year.value,K.rest_of_year.lines+' lines · '+K.rest_of_year.pct+'% of book'],
    ['kv-fut','ks-fut',K.future.value,K.future.lines+' lines · '+K.future.pct+'% of book'],
  ];
  sets.forEach(([vi,si,val,sub])=>{
    const ve=document.getElementById(vi),se=document.getElementById(si);
    if(ve)ve.textContent=f(val);
    if(se)se.textContent=sub;
  });
}

// ---- OVERVIEW charts ----
function buildOverview(){
  const K=D.kpis;
  new Chart(document.getElementById('c-bucket'),{type:'bar',data:{
    labels:['Past Due','Today','This Month','This Quarter','Next 90d','Rest of Year','Future'],
    datasets:[{data:[K.past_due.value,K.today.value,K.this_month.value,K.this_quarter.value,K.next_90.value,K.rest_of_year.value,K.future.value],
      backgroundColor:['#ff4757','#ffb340','#9b6dff','#06b6d4','#3d6bff','#22d3a0','#6b7799'],borderRadius:5,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>f(c.parsed.y)}}},
      scales:{x:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799',callback:v=>f(v)},grid:{color:'#1c2030'}}}}});

  const pb=D.pobin_all;
  new Chart(document.getElementById('c-pobin-all'),{type:'doughnut',data:{
    labels:pb.map(d=>d.type),datasets:[{data:pb.map(d=>d.value),backgroundColor:['#3d6bff','#22d3a0'],borderWidth:2,borderColor:'#0e1118'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:'#6b7799',font:{size:11},boxWidth:10,padding:10}},
      tooltip:{callbacks:{label:c=>c.label+': '+f(c.parsed)+' ('+pb[c.dataIndex].pct+'%)' }}}}});

  new Chart(document.getElementById('c-func-pie'),{type:'doughnut',data:{
    labels:D.func_data.map(d=>d['ActionBy - New']),datasets:[{data:D.func_data.map(d=>d.value),backgroundColor:COLS,borderWidth:2,borderColor:'#0e1118'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{color:'#6b7799',font:{size:11},boxWidth:10,padding:8}},
      tooltip:{callbacks:{label:c=>c.label+': '+f(c.parsed)+' ('+D.func_data[c.dataIndex].pct+'%)'}}}}});

  const top8=D.site_data.slice(0,8);
  new Chart(document.getElementById('c-site-bar'),{type:'bar',data:{
    labels:top8.map(d=>d['Whse Site']),datasets:[{data:top8.map(d=>d.value),backgroundColor:COLS,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>f(c.parsed.y)}}},
      scales:{x:{ticks:{color:'#6b7799',maxRotation:30},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799',callback:v=>f(v)},grid:{color:'#1c2030'}}}}});
}

// ---- PAST DUE ----
function buildPastDue(){
  const K=D.kpis;
  // summary KPIs
  const pdPO=D.pobin_pd.find(x=>x.type==='PO')||{lines:0,value:0,pct:0};
  const pdBin=D.pobin_pd.find(x=>x.type==='Binstock')||{lines:0,value:0,pct:0};
  const setE=(id,v)=>{const e=document.getElementById(id);if(e)e.textContent=v;};
  setE('kv-pd2',f(K.past_due.value));setE('ks-pd2',K.past_due.lines+' lines · '+K.past_due.pct+'% of book');
  setE('kv-pd-po',fn(pdPO.lines));setE('ks-pd-po',f(pdPO.value));
  setE('kv-pd-bin',fn(pdBin.lines));setE('ks-pd-bin',f(pdBin.value));

  // site charts
  const sites=D.pd_site_data.slice(0,15);
  new Chart(document.getElementById('c-pdsite-lines'),{type:'bar',data:{
    labels:sites.map(d=>d['Customer Name']),datasets:[{label:'Lines',data:sites.map(d=>d.lines),backgroundColor:COLS,borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799',font:{size:11}},grid:{color:'#1c2030'}}}}});

  new Chart(document.getElementById('c-pdsite-val'),{type:'bar',data:{
    labels:sites.map(d=>d['Customer Name']),datasets:[{label:'Value',data:sites.map(d=>d.value),backgroundColor:COLS.map(c=>c+'cc'),borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>f(c.parsed.x)}}},
      scales:{x:{ticks:{color:'#6b7799',callback:v=>f(v)},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799',font:{size:11}},grid:{color:'#1c2030'}}}}});

  // pobin chart + table
  const pb=D.pobin_pd; const pdtot=K.past_due.value||1;
  new Chart(document.getElementById('c-pobin-pd'),{type:'doughnut',data:{
    labels:pb.map(d=>d.type),datasets:[{data:pb.map(d=>d.value),backgroundColor:['#3d6bff','#9b6dff'],borderWidth:2,borderColor:'#0e1118'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:'#6b7799',font:{size:11},boxWidth:10,padding:10}},
      tooltip:{callbacks:{label:c=>c.label+': '+f(c.parsed)+' ('+pb[c.dataIndex].pct+'%)'}}}}});

  const tb=document.getElementById('tb-pobin-pd');
  pb.forEach(d=>{
    tb.innerHTML+=`<tr><td><span class="badge blue">${d.type}</span></td><td class="num">${fn(d.lines)}</td><td class="num">${ff(d.value)}</td><td class="num">${fp(d.pct)}</td><td>${barCell(d.value,pdtot,'var(--accent)')}</td></tr>`;
  });

  // site detail table
  const tb2=document.getElementById('tb-pdsite');
  const pdSiteMax=D.pd_site_data[0]?D.pd_site_data[0].value:1;
  D.pd_site_data.forEach((d,i)=>{
    const pctOfPd=pct(d.value,K.past_due.value);
    tb2.innerHTML+=`<tr><td class="num" style="color:var(--muted)">${i+1}</td><td>${d['Customer Name']}</td><td class="num">${fn(d.lines)}</td><td class="num">${ff(d.value)}</td><td class="num">${pctOfPd}%</td><td>${barCell(d.value,pdSiteMax,'var(--red)')}</td></tr>`;
  });
}

// ---- ORDER BOOK ----
function buildOrderBook(){
  const K=D.kpis;
  const bkts=[
    ['bv-pd','bs-pd',K.past_due,'red'],
    ['bv-td','bs-td',K.today,'amber'],
    ['bv-mo','bs-mo',K.this_month,'purple'],
    ['bv-q','bs-q',K.this_quarter,'#06b6d4'],
    ['bv-90','bs-90',K.next_90,'accent'],
    ['bv-yr','bs-yr',K.rest_of_year,'green'],
    ['bv-fut','bs-fut',K.future,'muted'],
  ];
  bkts.forEach(([vi,si,b])=>{
    const ve=document.getElementById(vi),se=document.getElementById(si);
    if(ve)ve.textContent=f(b.value);
    if(se)se.textContent=fn(b.lines)+' lines · '+b.pct+'%';
  });

  const labels=['Past Due','Today','This Month','This Quarter','Next 90d','Rest of Year','Future'];
  const values=bkts.map(([,,b])=>b.value);
  const lvals=bkts.map(([,,b])=>b.lines);
  const bgs=['#ff4757','#ffb340','#9b6dff','#06b6d4','#3d6bff','#22d3a0','#6b7799'];

  new Chart(document.getElementById('c-ob-bucket'),{type:'bar',data:{labels,datasets:[{data:values,backgroundColor:bgs,borderRadius:5,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>f(c.parsed.y)}}},
      scales:{x:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799',callback:v=>f(v)},grid:{color:'#1c2030'}}}}});

  new Chart(document.getElementById('c-ob-lines'),{type:'bar',data:{labels,datasets:[{data:lvals,backgroundColor:bgs.map(c=>c+'99'),borderRadius:5,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fn(c.parsed.y)+' lines'}}},
      scales:{x:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}}}}});
}

// ---- FUNCTIONS ----
function buildFunctions(){
  const fd=[...D.func_data].sort((a,b)=>b.value-a.value);
  new Chart(document.getElementById('c-func-val'),{type:'bar',data:{
    labels:fd.map(d=>d['ActionBy - New']),datasets:[{data:fd.map(d=>d.value),backgroundColor:COLS,borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>f(c.parsed.x)}}},
      scales:{x:{ticks:{color:'#6b7799',callback:v=>f(v)},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}}}}});

  new Chart(document.getElementById('c-func-lines'),{type:'bar',data:{
    labels:fd.map(d=>d['ActionBy - New']),datasets:[{data:fd.map(d=>d.lines),backgroundColor:COLS.map(c=>c+'aa'),borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}}}}});

  const tb=document.getElementById('tb-func');
  const maxV=fd[0]?fd[0].value:1;
  fd.forEach((d,i)=>{
    const fn_=d['ActionBy - New'];
    tb.innerHTML+=`<tr>
      <td><span class="badge" style="background:${COLS[i]}22;color:${COLS[i]}">${fn_}</span></td>
      <td class="num">${fn(d.lines)}</td><td class="num">${ff(d.value)}</td><td class="num">${fp(d.pct)}</td>
      <td class="num">${fn(d.pd_lines||0)}</td><td class="num" style="color:var(--red)">${ff(d.pd_value||0)}</td>
      <td class="num" style="color:var(--amber)">${fp(d.pd_pct||0)}</td>
      <td>${barCell(d.value,maxV,COLS[i])}</td>
    </tr>`;
  });
}

// ---- SITES ----
function buildSites(){
  const sd=[...D.site_data];
  const top8=sd.slice(0,8);
  new Chart(document.getElementById('c-site-full'),{type:'bar',data:{
    labels:top8.map(d=>d['Whse Site']),datasets:[{data:top8.map(d=>d.value),backgroundColor:COLS,borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>f(c.parsed.x)}}},
      scales:{x:{ticks:{color:'#6b7799',callback:v=>f(v)},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}}}}});

  const top5=sd.slice(0,5);
  const otherV=sd.slice(5).reduce((s,d)=>s+d.value,0);
  const dd=[...top5,{'Whse Site':'Other',value:otherV,pct:((otherV/D.kpis.total_book_value)*100).toFixed(2)}];
  new Chart(document.getElementById('c-site-donut'),{type:'doughnut',data:{
    labels:dd.map(d=>d['Whse Site']),datasets:[{data:dd.map(d=>d.value),backgroundColor:COLS,borderWidth:2,borderColor:'#0e1118'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.label+': '+f(c.parsed)}}}}});

  const leg=document.getElementById('site-legend');
  dd.forEach((d,i)=>{leg.innerHTML+=`<div class="legend-item"><div class="legend-dot" style="background:${COLS[i]}"></div>${d['Whse Site']} — ${f(d.value)} (${parseFloat(d.pct).toFixed(1)}%)</div>`;});

  const tb=document.getElementById('tb-sites');
  const maxV=sd[0]?sd[0].value:1;
  sd.forEach((d,i)=>{
    tb.innerHTML+=`<tr><td class="num" style="color:var(--muted)">${i+1}</td>
      <td><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${COLS[i]||'#666'};margin-right:8px"></span>${d['Whse Site']}</td>
      <td class="num">${fn(d.lines)}</td><td class="num">${ff(d.value)}</td><td class="num">${fp(d.pct)}</td>
      <td>${barCell(d.value,maxV,COLS[i]||'var(--accent)')}</td></tr>`;
  });
}

// ---- INVENTORY ----
let _fillFilter='all';
function buildInventory(){
  const IS=D.inv_summary;

  // stat cards
  const stats=[
    ['Past Due Lines',fn(IS.pd_lines),'var(--red)'],
    ['Past Due Value',f(IS.pd_value),'var(--red)'],
    ['Lines w/ ATP Stock',fn(IS.pd_with_atp_lines),'var(--green)'],
    ['ATP-covered Value',f(IS.pd_with_atp_value),'var(--green)'],
    ['No Stock Lines',fn(IS.pd_no_atp_lines),'var(--amber)'],
    ['No Stock Value',f(IS.pd_no_atp_value),'var(--amber)'],
    ['Fillable Orders',fn(IS.fillable_orders),'var(--accent)'],
    ['Fillable Value',f(IS.fillable_value),'var(--accent)'],
  ];
  const sg=document.getElementById('inv-stat-grid');
  stats.forEach(([lbl,val,col])=>{
    sg.innerHTML+=`<div class="inv-stat"><div class="inv-stat-val" style="color:${col}">${val}</div><div class="inv-stat-label">${lbl}</div></div>`;
  });

  // fill pie
  new Chart(document.getElementById('c-fill-pie'),{type:'doughnut',data:{
    labels:['Fillable','Unfillable / No Stock'],
    datasets:[{data:[IS.fillable_value,IS.unfillable_value],backgroundColor:['#22d3a0','#ff4757'],borderWidth:2,borderColor:'#0e1118'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:'#6b7799',font:{size:11},boxWidth:10,padding:10}},
      tooltip:{callbacks:{label:c=>c.label+': '+f(c.parsed)}}}}});

  // plant atp chart
  const pdVal=IS.pd_value||1;
  new Chart(document.getElementById('c-plant-atp'),{type:'bar',data:{
    labels:['Miami (8000)','Chandler (8003)'],
    datasets:[
      {label:'Lines with ATP',data:[IS.miami_atp_parts,IS.chandler_atp_parts],backgroundColor:['#3d6bff','#22d3a0'],borderRadius:4},
    ]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#6b7799',font:{size:11}}}},
      scales:{x:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799'},grid:{color:'#1c2030'}}}}});

  renderFillTable('all');
}

function renderFillTable(filter){
  _fillFilter=filter;
  const tb=document.getElementById('tb-fill');
  tb.innerHTML='';
  let rows=D.fill_data;
  if(filter==='PO') rows=rows.filter(r=>r['PO/Bin']==='PO');
  else if(filter==='Binstock') rows=rows.filter(r=>r['PO/Bin']==='Binstock');
  else if(filter==='can') rows=rows.filter(r=>r.orders_fillable>0);
  else if(filter==='no') rows=rows.filter(r=>r.atp_total===0);
  rows.forEach(r=>{
    const canFill=r.orders_fillable>0;
    const noStock=r.atp_total===0;
    tb.innerHTML+=`<tr>
      <td><span class="badge ${noStock?'red':canFill?'green':'amber'}">${r['Part Number']}</span></td>
      <td><span class="badge blue">${r['PO/Bin']}</span></td>
      <td class="num">${fn(r.total_orders)}</td>
      <td class="num">${fn(r.total_qty)}</td>
      <td class="num">${fn(r.atp_miami)}</td>
      <td class="num">${fn(r.atp_chandler)}</td>
      <td class="num" style="color:${noStock?'var(--red)':canFill?'var(--green)':'var(--amber)'}">${fn(r.atp_total)}</td>
      <td class="num fill-can">${fn(r.orders_fillable)}</td>
      <td class="num fill-can">${ff(r.filled_value)}</td>
      <td class="num fill-no">${fn(r.unfilled_orders)}</td>
      <td class="num fill-no">${ff(r.unfilled_value)}</td>
      <td><span class="badge ${r.fill_rate_orders>=50?'green':r.fill_rate_orders>0?'amber':'red'}">${r.fill_rate_orders}%</span></td>
    </tr>`;
  });
}

function filterFill(filter,btn){
  document.querySelectorAll('.filter-bar .filter-btn').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
  renderFillTable(filter);
}

// ---- EXECUTION ----
function buildExecution(){
  const top=D.exec_data.slice(0,12);
  new Chart(document.getElementById('c-exec'),{type:'bar',data:{
    labels:top.map(d=>d.name.split(' ').pop()),
    datasets:[
      {label:'Total Value',data:top.map(d=>d.value),backgroundColor:'#3d6bff44',borderColor:'#3d6bff',borderWidth:1,borderRadius:4},
      {label:'Past Due',data:top.map(d=>d.pd_value),backgroundColor:'#ff475788',borderColor:'#ff4757',borderWidth:1,borderRadius:4},
      {label:'Due Today',data:top.map(d=>d.today_val),backgroundColor:'#ffb34088',borderColor:'#ffb340',borderWidth:1,borderRadius:4},
      {label:'Next 90d',data:top.map(d=>d.next90_val),backgroundColor:'#22d3a055',borderColor:'#22d3a0',borderWidth:1,borderRadius:4},
    ]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#6b7799',font:{size:11}}},tooltip:{callbacks:{label:c=>c.dataset.label+': '+f(c.parsed.y)}}},
      scales:{x:{ticks:{color:'#6b7799',maxRotation:30},grid:{color:'#1c2030'}},y:{ticks:{color:'#6b7799',callback:v=>f(v)},grid:{color:'#1c2030'}}}}});

  const tb=document.getElementById('tb-exec');
  const maxV=D.exec_data[0]?D.exec_data[0].value:1;
  D.exec_data.forEach(d=>{
    tb.innerHTML+=`<tr>
      <td>${d.name}</td><td class="num">${fn(d.lines)}</td>
      <td class="num">${ff(d.value)}</td>
      <td class="num" style="color:var(--red)">${ff(d.pd_value)}</td>
      <td class="num" style="color:var(--amber)">${ff(d.today_val)}</td>
      <td class="num" style="color:var(--accent)">${ff(d.next90_val)}</td>
      <td class="num">${fp(d.pct)}</td>
      <td>${barCell(d.value,maxV,'var(--accent)')}</td>
    </tr>`;
  });
}

// ---- REP VIEW ----
function buildRepCards(filter){
  const grid=document.getElementById('rep-grid');
  grid.innerHTML='';
  const execs={};
  D.rep_data.forEach(r=>{
    if(!execs[r.exec])execs[r.exec]={name:r.exec,total_lines:r.total_lines,total_value:r.total_value,items:[]};
    if(filter==='all'||r.action===filter)execs[r.exec].items.push(r);
  });
  Object.values(execs).filter(e=>e.items.length)
    .sort((a,b)=>b.items.reduce((s,x)=>s+x.value,0)-a.items.reduce((s,x)=>s+x.value,0))
    .forEach(rep=>{
      const trv=rep.items.reduce((s,x)=>s+x.value,0);
      const blocks=rep.items.map(item=>{
        const cls=item.action==='P&E'?'accent':'red';
        const tv=item.value||1;
        const pd=Math.min((item.pd_value/tv)*100,100);
        const td_=Math.min((item.today_value/tv)*100,100);
        const n90=Math.min((item.next90_value/tv)*100,100);
        const ft=Math.min((item.future_value/tv)*100,100);
        return `<div class="risk-block">
          <div class="risk-lbl" style="color:${item.action==='P&E'?'var(--accent)':'var(--red)'}">${item.action}</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
            <div><div style="font-size:10px;color:var(--muted)">SOs</div><div style="font-size:13px;font-weight:700;font-family:var(--mono)">${fn(item.so_count)}</div></div>
            <div><div style="font-size:10px;color:var(--muted)">Lines</div><div style="font-size:13px;font-weight:700;font-family:var(--mono)">${fn(item.lines)}</div></div>
            <div><div style="font-size:10px;color:var(--muted)">Value</div><div style="font-size:13px;font-weight:700;font-family:var(--mono)">${f(item.value)}</div></div>
            <div><div style="font-size:10px;color:var(--muted)">% of Book</div><div style="font-size:13px;font-weight:700;font-family:var(--mono)">${fp(item.pct_exec)}</div></div>
          </div>
          <div class="gauge-row"><div class="gauge-lbl">Past Due</div><div class="gauge-bar"><div class="gauge-fill" style="width:${pd}%;background:var(--red)"></div></div><div class="gauge-pct" style="color:var(--red)">${pd.toFixed(1)}%</div></div>
          ${td_>0?`<div class="gauge-row"><div class="gauge-lbl">Today</div><div class="gauge-bar"><div class="gauge-fill" style="width:${td_}%;background:var(--amber)"></div></div><div class="gauge-pct" style="color:var(--amber)">${td_.toFixed(1)}%</div></div>`:''}
          <div class="gauge-row"><div class="gauge-lbl">Next 90d</div><div class="gauge-bar"><div class="gauge-fill" style="width:${n90}%;background:var(--accent)"></div></div><div class="gauge-pct" style="color:var(--accent)">${n90.toFixed(1)}%</div></div>
          <div class="gauge-row"><div class="gauge-lbl">Future</div><div class="gauge-bar"><div class="gauge-fill" style="width:${ft}%;background:var(--green)"></div></div><div class="gauge-pct" style="color:var(--green)">${ft.toFixed(1)}%</div></div>
        </div>`;
      }).join('');
      grid.innerHTML+=`<div class="rep-card"><div class="rep-name">${rep.name}</div><div class="rep-meta">Total: ${f(rep.total_value)} · ${fn(rep.total_lines)} lines &nbsp;|&nbsp; Risk Exposure: ${f(trv)} (${pct(trv,rep.total_value)}%)</div>${blocks}</div>`;
    });
}

function filterReps(type,btn){
  document.querySelectorAll('#tab-reps .filter-btn').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');
  buildRepCards(type);
}

// ---- TAB ROUTING ----
const _built={};
function show(name,el){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  if(!_built[name]){
    _built[name]=true;
    if(name==='pastdue')   buildPastDue();
    if(name==='orderbook') buildOrderBook();
    if(name==='functions') buildFunctions();
    if(name==='sites')     buildSites();
    if(name==='inventory') buildInventory();
    if(name==='execution') buildExecution();
    if(name==='reps')      buildRepCards('all');
  }
}

// ---- INIT ----
fillKPI();
buildOverview();
</script>
</body>
</html>"""


# =============================================================================
#  STEP 6 - BUILD HTML
# =============================================================================
def build_html(all_data, today):
    try:
        as_of = today.strftime("%b %#d, %Y")
    except:
        as_of = today.strftime("%b %-d, %Y")

    html = HTML.replace("__DATA_JSON__", json.dumps(all_data, ensure_ascii=False))
    html = html.replace("__AS_OF__", as_of)
    return html


# =============================================================================
#  MAIN
# =============================================================================
def main():
    # Allow CLI overrides: python generate_dashboard_v2.py <excel> <inventory>
    excel_path = Path(sys.argv[1]) if len(sys.argv) > 1 else EXCEL_PATH
    inv_path   = Path(sys.argv[2]) if len(sys.argv) > 2 else INVENTORY_PATH

    print("=" * 62)
    print("  Honeywell Program Dashboard Generator  v2")
    print("=" * 62)

    df, inv       = load_data(excel_path, inv_path)
    print("Calculating metrics ...")
    all_data, today = calculate(df, inv)

    print("Building index.html ...")
    html = build_html(all_data, today)
    OUTPUT_FILE.write_text(html, encoding="utf-8")

    print(f"""
Done!
  Output:  {OUTPUT_FILE}

Dashboard tabs:
  1. Overview       - KPI summary + high-level charts
  2. Past Due       - Site breakdown, PO vs Bin split, detail table
  3. Order Book     - Time bucket analysis (PD / Today / Month / Quarter / 90d / YTD / Future)
  4. Functions      - ActionBy breakdown ($ and lines) vs Past Due exposure
  5. Sites          - Warehouse site breakdown (full book)
  6. Inventory/ATP  - Fill analysis for past due orders (Miami 8000 + Chandler 8003)
  7. Execution      - Execution person backlog + past due exposure
  8. Rep View       - P&E / SPS Risk scorecard by Execution Person
{"=" * 62}""")


if __name__ == "__main__":
    main()
