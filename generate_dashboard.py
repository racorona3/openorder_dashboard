r"""
=============================================================================
  Honeywell Order Book Dashboard Generator
=============================================================================
  Both this script and the output index.html live in:
    C:\Users\zn424f\OneDrive - The Boeing Company\Working KPIs\Order Book Analysis\

  Source workbook:
    C:\Users\zn424f\OneDrive - The Boeing Company\Working KPIs\
        Working Honeywell KPI Dashboard.xlsx
  Sheet:  'SAW Report Data for Current Day'

  HOW TO USE:
    1. One-time setup  ->  open Command Prompt and run:
           pip install pandas openpyxl
    2. Every day       ->  double-click this file  (or: python generate_dashboard.py)
    3. A fresh index.html appears in this same folder.
    4. Drop index.html into your GitHub repo root - done.
=============================================================================
"""

import json
import sys
from datetime import date
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit(
        "\nERROR: pandas is not installed.\n"
        "Fix:   open Command Prompt and run:  pip install pandas openpyxl\n"
    )

# =============================================================================
#  CONFIG  -  only edit if file names or paths ever change
# =============================================================================

EXCEL_PATH = Path(
    r"C:\Users\zn424f\OneDrive - The Boeing Company"
    r"\Working KPIs\Working Honeywell KPI Dashboard.xlsx"
)
SHEET_NAME  = "SAW Report Data for Current Day"

# index.html is written to the same folder as this script
OUTPUT_FILE = Path(__file__).resolve().parent / "index.html"

# Column names as they appear in the sheet
COL_SO          = "SO Number"
COL_EXT_PRICE   = "Extended Price"
COL_SHIP_DATE   = "Ship Request Date"
COL_STATUS      = "Status"
COL_ACTION_BY   = "ActionBy - New"
COL_EXEC_PERSON = "Execution Person"
COL_WHSE_SITE   = "Whse Site"


# =============================================================================
#  STEP 1 - LOAD DATA
# =============================================================================
def load_data():
    if not EXCEL_PATH.exists():
        sys.exit(
            f"\nERROR: Cannot find the workbook at:\n  {EXCEL_PATH}\n\n"
            "Check that OneDrive has finished syncing and the file name matches."
        )

    print(f"Reading: {EXCEL_PATH.name}  [sheet: {SHEET_NAME}] ...")
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    print(f"  {len(df):,} rows x {len(df.columns)} columns loaded.")

    missing = [c for c in [COL_SO, COL_EXT_PRICE, COL_SHIP_DATE, COL_STATUS,
                            COL_ACTION_BY, COL_EXEC_PERSON, COL_WHSE_SITE]
               if c not in df.columns]
    if missing:
        sys.exit(
            "\nERROR: These required columns were not found in the sheet:\n"
            + "\n".join(f"  - {c}" for c in missing)
            + "\n\nAll columns present:\n"
            + ", ".join(df.columns.tolist())
        )

    df[COL_EXT_PRICE] = pd.to_numeric(df[COL_EXT_PRICE], errors="coerce").fillna(0)
    df[COL_SHIP_DATE] = pd.to_datetime(df[COL_SHIP_DATE], errors="coerce")
    return df


# =============================================================================
#  STEP 2 - CALCULATE ALL METRICS
# =============================================================================
def calculate(df):
    today   = pd.Timestamp(date.today())
    next_90 = today + pd.Timedelta(days=90)

    df = df.copy()
    # Use the Status column as the source of truth for Past Due and Today --
    # these are pre-calculated by the source system and match your report exactly.
    # Ship Request Date alone over-counts because some "Future" rows have past dates.
    df["_past_due"] = df[COL_STATUS].astype(str).str.strip() == "Past Due"
    df["_today"]    = df[COL_STATUS].astype(str).str.strip() == "Today"
    # 90-day bucket: ship date within range, but not already flagged Past Due or Today
    df["_due_90"]   = (
        ~df["_past_due"] & ~df["_today"] &
        (df[COL_SHIP_DATE] >= today) &
        (df[COL_SHIP_DATE] <= next_90)
    )
    df["_future"]   = ~df["_past_due"] & ~df["_today"] & ~df["_due_90"]

    total_val = df[COL_EXT_PRICE].sum()

    # --- KPIs ---
    kpis = {
        "total_book_value":     round(total_val, 2),
        "total_lines":          int(len(df)),
        "total_past_due_value": round(df.loc[df["_past_due"], COL_EXT_PRICE].sum(), 2),
        "total_past_due_lines": int(df["_past_due"].sum()),
        "total_today_value":    round(df.loc[df["_today"],    COL_EXT_PRICE].sum(), 2),
        "total_today_lines":    int(df["_today"].sum()),
        "total_due_90_value":   round(df.loc[df["_due_90"],   COL_EXT_PRICE].sum(), 2),
        "total_due_90_lines":   int(df["_due_90"].sum()),
        "total_future_lines":   int(df["_future"].sum()),
    }

    # --- Function breakdown (ActionBy - New) ---
    func = (df.groupby(COL_ACTION_BY)
              .agg(lines=(COL_SO, "count"), value=(COL_EXT_PRICE, "sum"))
              .reset_index()
              .rename(columns={COL_ACTION_BY: "ActionBy - New"}))
    func["pct"]   = (func["value"] / total_val * 100).round(2)
    func["value"] = func["value"].round(2)
    func_data = func.sort_values("value", ascending=False).to_dict("records")

    # --- Site breakdown (Whse Site) ---
    site = (df.groupby(COL_WHSE_SITE)
              .agg(value=(COL_EXT_PRICE, "sum"))
              .reset_index()
              .rename(columns={COL_WHSE_SITE: "Whse Site"}))
    site["pct"]   = (site["value"] / total_val * 100).round(2)
    site["value"] = site["value"].round(2)
    site_data = site.sort_values("value", ascending=False).to_dict("records")

    # --- Execution Person - full backlog ---
    exec_rows = []
    for name, grp in df.groupby(COL_EXEC_PERSON):
        exec_rows.append({
            "Execution Person": name,
            "lines":            len(grp),
            "value":            round(grp[COL_EXT_PRICE].sum(), 2),
            "past_due_value":   round(grp.loc[grp["_past_due"], COL_EXT_PRICE].sum(), 2),
            "due_90_value":     round(grp.loc[grp["_due_90"],   COL_EXT_PRICE].sum(), 2),
            "pct":              round(grp[COL_EXT_PRICE].sum() / total_val * 100, 2),
        })
    exec_data = sorted(exec_rows, key=lambda x: x["value"], reverse=True)

    # --- Rep view: P&E and SPS RISK by Execution Person ---
    risk_df = df[df[COL_ACTION_BY].isin(["P&E", "SPS RISK"])].copy()
    totals  = (df.groupby(COL_EXEC_PERSON)
                 .agg(total_lines=(COL_SO, "count"), total_value=(COL_EXT_PRICE, "sum"))
                 .to_dict("index"))

    rep_rows = []
    for (exec_name, action), grp in risk_df.groupby([COL_EXEC_PERSON, COL_ACTION_BY]):
        tot_val = float(totals.get(exec_name, {}).get("total_value", 1) or 1)
        tot_lines = int(totals.get(exec_name, {}).get("total_lines", 0))
        rep_rows.append({
            "Execution Person":  exec_name,
            "ActionBy - New":    action,
            "so_count":          int(grp[COL_SO].nunique()),
            "lines":             len(grp),
            "value":             round(grp[COL_EXT_PRICE].sum(), 2),
            "past_due_value":    round(grp.loc[grp["_past_due"], COL_EXT_PRICE].sum(), 2),
            "today_value":       round(grp.loc[grp["_today"],    COL_EXT_PRICE].sum(), 2),
            "due_90_value":      round(grp.loc[grp["_due_90"],   COL_EXT_PRICE].sum(), 2),
            "future_value":      round(grp.loc[grp["_future"],   COL_EXT_PRICE].sum(), 2),
            "past_due_lines":    int(grp["_past_due"].sum()),
            "due_90_lines":      int(grp["_due_90"].sum()),
            "total_lines":       tot_lines,
            "total_value":       round(tot_val, 2),
            "pct_of_exec_total": round(grp[COL_EXT_PRICE].sum() / tot_val * 100, 2),
        })
    rep_data = sorted(rep_rows, key=lambda x: x["value"], reverse=True)

    return kpis, func_data, site_data, exec_data, rep_data, today


# =============================================================================
#  STEP 3 - FORMAT HELPERS
# =============================================================================
def fmt_short(n):
    if n >= 1_000_000: return f"${n/1_000_000:.1f}M"
    if n >= 1_000:     return f"${n/1_000:.0f}K"
    return f"${n:.0f}"

def pct_of(val, total):
    return f"{val / total * 100:.1f}" if total else "0.0"

def fmt_date(today):
    try:    return today.strftime("%b %-d, %Y")   # Linux/Mac
    except: return today.strftime("%b %#d, %Y")   # Windows


# =============================================================================
#  STEP 4 - HTML TEMPLATE
# =============================================================================
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Order Book Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  :root{--bg:#0f1117;--surface:#1a1d27;--surface2:#22263a;--border:#2e3350;--accent:#4f6ef7;--green:#22c55e;--red:#ef4444;--amber:#f59e0b;--text:#e2e8f0;--muted:#94a3b8;--font:'Segoe UI',system-ui,sans-serif}
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--text);font-family:var(--font);font-size:14px;min-height:100vh}
  nav{background:var(--surface);border-bottom:1px solid var(--border);padding:0 24px;display:flex;align-items:center;height:52px;position:sticky;top:0;z-index:100}
  .nav-brand{font-weight:700;font-size:15px;color:var(--text);margin-right:32px;display:flex;align-items:center;gap:8px}
  .nav-brand span{color:var(--accent)}
  .nav-tab{padding:0 16px;height:52px;display:flex;align-items:center;border-bottom:2px solid transparent;cursor:pointer;color:var(--muted);font-size:13px;font-weight:500;transition:all .15s;white-space:nowrap}
  .nav-tab:hover{color:var(--text)}
  .nav-tab.active{color:var(--text);border-bottom-color:var(--accent)}
  .nav-right{margin-left:auto;font-size:12px;color:var(--muted)}
  .page{display:none;padding:24px;max-width:1400px;margin:0 auto}
  .page.active{display:block}
  .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px}
  .kpi-card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 20px}
  .kpi-label{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin-bottom:8px}
  .kpi-value{font-size:26px;font-weight:700;color:var(--text)}
  .kpi-sub{font-size:12px;color:var(--muted);margin-top:4px}
  .kpi-card.red .kpi-value{color:var(--red)}
  .kpi-card.amber .kpi-value{color:var(--amber)}
  .kpi-card.green .kpi-value{color:var(--green)}
  .kpi-card.accent .kpi-value{color:var(--accent)}
  .section{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:20px;margin-bottom:20px}
  .section-title{font-size:13px;font-weight:600;color:var(--text);margin-bottom:16px;display:flex;align-items:center;gap:8px}
  .badge{background:var(--surface2);color:var(--muted);font-size:11px;padding:2px 8px;border-radius:20px;font-weight:400}
  .charts-row{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}
  @media(max-width:900px){.charts-row{grid-template-columns:1fr}}
  .chart-wrap{position:relative;height:280px}
  .table-wrap{overflow-x:auto}
  table{width:100%;border-collapse:collapse;font-size:13px}
  thead th{text-align:left;font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;padding:8px 12px;border-bottom:1px solid var(--border);background:var(--surface2);white-space:nowrap}
  tbody tr{border-bottom:1px solid var(--border);transition:background .1s}
  tbody tr:hover{background:var(--surface2)}
  tbody td{padding:9px 12px;white-space:nowrap}
  .tag{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
  .num{font-variant-numeric:tabular-nums}
  .red-text{color:var(--red)}
  .amber-text{color:var(--amber)}
  .green-text{color:var(--green)}
  .bar-bg{background:var(--surface2);border-radius:3px;height:6px}
  .bar-fill{height:6px;border-radius:3px;background:var(--accent)}
  .rep-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:16px}
  .rep-card{background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:18px}
  .rep-name{font-weight:700;font-size:15px;margin-bottom:4px}
  .rep-total{font-size:12px;color:var(--muted);margin-bottom:14px}
  .risk-block{border:1px solid var(--border);border-radius:8px;padding:12px;margin-bottom:10px}
  .risk-block:last-child{margin-bottom:0}
  .risk-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px}
  .risk-title.pe{color:#7b9fff}
  .risk-title.sps{color:#f87171}
  .risk-metrics{display:grid;grid-template-columns:1fr 1fr;gap:8px}
  .risk-metric-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px}
  .risk-metric-val{font-size:14px;font-weight:600}
  .gauge-row{display:flex;align-items:center;gap:8px;margin-top:8px}
  .gauge-label{font-size:11px;color:var(--muted);width:80px;flex-shrink:0}
  .gauge-bar{flex:1;height:8px;background:var(--surface);border-radius:4px;overflow:hidden}
  .gauge-fill{height:100%;border-radius:4px}
  .gauge-pct{font-size:11px;color:var(--muted);width:36px;text-align:right}
  .progress-cell{display:flex;align-items:center;gap:8px}
  .progress-cell .bar-bg{flex:1}
  .pct-lbl{font-size:12px;color:var(--muted);width:36px;text-align:right}
  .filter-bar{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
  .filter-btn{padding:5px 14px;border-radius:20px;border:1px solid var(--border);background:var(--surface2);color:var(--muted);font-size:12px;cursor:pointer;transition:all .15s}
  .filter-btn:hover,.filter-btn.active{background:var(--accent);color:#fff;border-color:var(--accent)}
  .filter-label{font-size:12px;color:var(--muted)}
  .donut-legend{display:flex;flex-direction:column;gap:6px;margin-top:12px}
  .legend-row{display:flex;align-items:center;gap:8px}
  .legend-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
  .legend-text{font-size:12px;color:var(--muted);flex:1}
  .legend-val{font-size:12px;font-weight:600}
  .legend-pct{font-size:11px;color:var(--muted);width:40px;text-align:right}
</style>
</head>
<body>
<nav>
  <div class="nav-brand">&#9632; <span>Order Book</span> Dashboard</div>
  <div class="nav-tab active" onclick="showTab('overview',this)">Overview</div>
  <div class="nav-tab" onclick="showTab('functions',this)">By Function</div>
  <div class="nav-tab" onclick="showTab('sites',this)">By Site</div>
  <div class="nav-tab" onclick="showTab('execution',this)">Execution Team</div>
  <div class="nav-tab" onclick="showTab('reps',this)">Rep View</div>
  <div class="nav-right">As of __AS_OF__</div>
</nav>

<!-- OVERVIEW -->
<div id="tab-overview" class="page active">
  <div class="kpi-grid">
    <div class="kpi-card accent">
      <div class="kpi-label">Total Book Value</div>
      <div class="kpi-value">__KPI_BOOK__</div>
      <div class="kpi-sub">__KPI_LINES__ line items</div>
    </div>
    <div class="kpi-card red">
      <div class="kpi-label">Past Due</div>
      <div class="kpi-value">__KPI_PD_VAL__</div>
      <div class="kpi-sub">__KPI_PD_LINES__ lines &middot; __KPI_PD_PCT__% of book</div>
    </div>
    <div class="kpi-card" style="background:var(--surface);border:1px solid #7c3aed44">
      <div class="kpi-label">Due Today</div>
      <div class="kpi-value" style="color:#a855f7">__KPI_TOD_VAL__</div>
      <div class="kpi-sub">__KPI_TOD_LINES__ lines &middot; __KPI_TOD_PCT__% of book</div>
    </div>
    <div class="kpi-card amber">
      <div class="kpi-label">Due Next 90 Days</div>
      <div class="kpi-value">__KPI_90_VAL__</div>
      <div class="kpi-sub">__KPI_90_LINES__ lines &middot; __KPI_90_PCT__% of book</div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label">Future (90+ Days)</div>
      <div class="kpi-value">__KPI_FUT_VAL__</div>
      <div class="kpi-sub">__KPI_FUT_LINES__ lines &middot; __KPI_FUT_PCT__% of book</div>
    </div>
  </div>
  <div class="charts-row">
    <div class="section">
      <div class="section-title">Book Value &mdash; Due Date Buckets</div>
      <div class="chart-wrap"><canvas id="bucketChart"></canvas></div>
    </div>
    <div class="section">
      <div class="section-title">Financial Impact by Function</div>
      <div class="chart-wrap"><canvas id="funcPieChart"></canvas></div>
    </div>
  </div>
  <div class="section">
    <div class="section-title">Site Financial Breakdown <span class="badge">% of Total Book</span></div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Site</th><th>Value ($)</th><th>% of Book</th><th style="width:200px">Share</th></tr></thead>
        <tbody id="siteTableBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- FUNCTIONS -->
<div id="tab-functions" class="page">
  <div class="section">
    <div class="section-title">Breakdown by Function (ActionBy)</div>
    <div class="charts-row">
      <div><div class="chart-wrap"><canvas id="funcBarChart"></canvas></div></div>
      <div><div class="chart-wrap"><canvas id="funcLinesChart"></canvas></div></div>
    </div>
  </div>
  <div class="section">
    <div class="section-title">Function Detail Table</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Function</th><th>Line Items</th><th>Financial Impact ($)</th><th>% of Total Book</th><th style="width:180px">Share</th></tr></thead>
        <tbody id="funcTableBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- SITES -->
<div id="tab-sites" class="page">
  <div class="charts-row">
    <div class="section">
      <div class="section-title">Financial Impact by Honeywell Site</div>
      <div class="chart-wrap"><canvas id="siteBarChart"></canvas></div>
    </div>
    <div class="section">
      <div class="section-title">Site Share (Donut)</div>
      <div style="height:200px;position:relative"><canvas id="siteDonutChart"></canvas></div>
      <div class="donut-legend" id="siteDonutLegend"></div>
    </div>
  </div>
  <div class="section">
    <div class="section-title">All Sites &mdash; Full Detail</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Site</th><th>Value ($)</th><th>% of Book</th><th style="width:200px">Share</th></tr></thead>
        <tbody id="siteFullTableBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- EXECUTION TEAM -->
<div id="tab-execution" class="page">
  <div class="section">
    <div class="section-title">Execution Person &mdash; Backlog Summary</div>
    <div class="chart-wrap" style="height:320px"><canvas id="execBarChart"></canvas></div>
  </div>
  <div class="section">
    <div class="section-title">Execution Person Detail</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Execution Person</th><th>Lines</th><th>Backlog Value ($)</th><th>Past Due ($)</th><th>Due 90 Days ($)</th><th>% of Book</th><th style="width:140px">Share</th></tr></thead>
        <tbody id="execTableBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- REP VIEW -->
<div id="tab-reps" class="page">
  <div class="section" style="margin-bottom:16px">
    <div class="section-title">P&amp;E and SPS Risk &mdash; Execution Person Scorecard</div>
    <p style="font-size:12px;color:var(--muted)">
      Filtered to orders where ActionBy (col AU) = <strong style="color:#7b9fff">P&amp;E</strong> or
      <strong style="color:#f87171">SPS Risk</strong>, grouped by Execution Person (col AI).
      Gauges show % of that person&rsquo;s P&amp;E / SPS Risk exposure, benchmarked against their total assigned open order book.
    </p>
  </div>
  <div class="filter-bar">
    <span class="filter-label">Filter by type:</span>
    <button class="filter-btn active" onclick="filterReps('all',this)">All</button>
    <button class="filter-btn" onclick="filterReps('P&amp;E',this)">P&amp;E Only</button>
    <button class="filter-btn" onclick="filterReps('SPS RISK',this)">SPS Risk Only</button>
  </div>
  <div class="rep-grid" id="repGrid"></div>
</div>

<script>
const DATA = __DATA_JSON__;
const COLORS = ['#4f6ef7','#7c5cbf','#22c55e','#f59e0b','#ef4444','#06b6d4','#ec4899','#84cc16','#f97316','#8b5cf6','#14b8a6','#f43f5e','#a78bfa','#fb923c','#34d399'];
const fmt     = n => n>=1e6?'$'+(n/1e6).toFixed(2)+'M':n>=1e3?'$'+(n/1e3).toFixed(1)+'K':'$'+n.toFixed(0);
const fmtFull = n => '$'+n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtNum  = n => n.toLocaleString('en-US');
const pct     = n => n.toFixed(2)+'%';

function buildOverviewCharts(){
  const K=DATA.kpis;
  const future=K.total_book_value-K.total_past_due_value-K.total_today_value-K.total_due_90_value;
  new Chart(document.getElementById('bucketChart'),{type:'bar',data:{labels:['Past Due','Due Today','Due \u226490 Days','Future (90+ Days)'],datasets:[{data:[K.total_past_due_value,K.total_today_value,K.total_due_90_value,future],backgroundColor:['#ef4444','#a855f7','#f59e0b','#22c55e'],borderRadius:6,borderSkipped:false}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const l=[K.total_past_due_lines,K.total_today_lines,K.total_due_90_lines,K.total_future_lines][ctx.dataIndex];return[fmt(ctx.parsed.y),fmtNum(l)+' lines'];}}}},scales:{x:{ticks:{color:'#94a3b8'},grid:{color:'#2e3350'}},y:{ticks:{color:'#94a3b8',callback:v=>fmt(v)},grid:{color:'#2e3350'}}}}});
  new Chart(document.getElementById('funcPieChart'),{type:'doughnut',data:{labels:DATA.func_data.map(d=>d['ActionBy - New']),datasets:[{data:DATA.func_data.map(d=>d.value),backgroundColor:COLORS,borderWidth:2,borderColor:'#1a1d27'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{color:'#94a3b8',font:{size:11},boxWidth:12,padding:8}},tooltip:{callbacks:{label:ctx=>ctx.label+': '+fmt(ctx.parsed)+' ('+DATA.func_data[ctx.dataIndex].pct+'%)'}}}}});
  const tb=document.getElementById('siteTableBody');
  DATA.site_data.slice(0,8).forEach(s=>{tb.innerHTML+=`<tr><td>${s['Whse Site']}</td><td class="num">${fmtFull(s.value)}</td><td class="num">${pct(s.pct)}</td><td><div class="bar-bg"><div class="bar-fill" style="width:${Math.max(s.pct,.5)}%"></div></div></td></tr>`;});
}

function buildFunctionCharts(){
  const sorted=[...DATA.func_data].sort((a,b)=>b.value-a.value);
  new Chart(document.getElementById('funcBarChart'),{type:'bar',data:{labels:sorted.map(d=>d['ActionBy - New']),datasets:[{label:'Financial Impact',data:sorted.map(d=>d.value),backgroundColor:COLORS,borderRadius:5}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>fmt(ctx.parsed.x)}}},scales:{x:{ticks:{color:'#94a3b8',callback:v=>fmt(v)},grid:{color:'#2e3350'}},y:{ticks:{color:'#94a3b8'},grid:{color:'#2e3350'}}}}});
  new Chart(document.getElementById('funcLinesChart'),{type:'bar',data:{labels:sorted.map(d=>d['ActionBy - New']),datasets:[{label:'Line Items',data:sorted.map(d=>d.lines),backgroundColor:COLORS.map(c=>c+'bb'),borderRadius:5}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#94a3b8'},grid:{color:'#2e3350'}},y:{ticks:{color:'#94a3b8'},grid:{color:'#2e3350'}}}}});
  const tb=document.getElementById('funcTableBody');
  sorted.forEach((d,i)=>{tb.innerHTML+=`<tr><td><span class="tag" style="background:${COLORS[i]}22;color:${COLORS[i]}">${d['ActionBy - New']}</span></td><td class="num">${fmtNum(d.lines)}</td><td class="num">${fmtFull(d.value)}</td><td class="num">${pct(d.pct)}</td><td><div class="bar-bg"><div class="bar-fill" style="width:${Math.max(d.pct,.5)}%;background:${COLORS[i]}"></div></div></td></tr>`;});
}

function buildSiteCharts(){
  const top=DATA.site_data.slice(0,8);
  new Chart(document.getElementById('siteBarChart'),{type:'bar',data:{labels:top.map(d=>d['Whse Site']),datasets:[{data:top.map(d=>d.value),backgroundColor:COLORS,borderRadius:5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>fmt(ctx.parsed.y)}}},scales:{x:{ticks:{color:'#94a3b8',maxRotation:30},grid:{color:'#2e3350'}},y:{ticks:{color:'#94a3b8',callback:v=>fmt(v)},grid:{color:'#2e3350'}}}}});
  const top5=DATA.site_data.slice(0,5);
  const other=DATA.site_data.slice(5).reduce((s,d)=>s+d.value,0);
  const dd=[...top5,{'Whse Site':'Other',value:other,pct:(other/DATA.kpis.total_book_value*100).toFixed(2)}];
  new Chart(document.getElementById('siteDonutChart'),{type:'doughnut',data:{labels:dd.map(d=>d['Whse Site']),datasets:[{data:dd.map(d=>d.value),backgroundColor:COLORS,borderWidth:2,borderColor:'#1a1d27'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}});
  const leg=document.getElementById('siteDonutLegend');
  dd.forEach((d,i)=>{leg.innerHTML+=`<div class="legend-row"><div class="legend-dot" style="background:${COLORS[i]}"></div><div class="legend-text">${d['Whse Site']}</div><div class="legend-val">${fmt(d.value)}</div><div class="legend-pct">${parseFloat(d.pct).toFixed(1)}%</div></div>`;});
  const tb=document.getElementById('siteFullTableBody');
  DATA.site_data.forEach((s,i)=>{tb.innerHTML+=`<tr><td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${COLORS[i]||'#666'};margin-right:8px"></span>${s['Whse Site']}</td><td class="num">${fmtFull(s.value)}</td><td class="num">${pct(s.pct)}</td><td><div class="bar-bg"><div class="bar-fill" style="width:${Math.max(s.pct,.3)}%"></div></div></td></tr>`;});
}

function buildExecCharts(){
  const top=DATA.exec_data.slice(0,10);
  new Chart(document.getElementById('execBarChart'),{type:'bar',data:{labels:top.map(d=>d['Execution Person'].split(' ').slice(-1)[0]),datasets:[{label:'Total Value',data:top.map(d=>d.value),backgroundColor:'#4f6ef744',borderColor:'#4f6ef7',borderWidth:1,borderRadius:4},{label:'Past Due',data:top.map(d=>d.past_due_value),backgroundColor:'#ef444488',borderColor:'#ef4444',borderWidth:1,borderRadius:4},{label:'Due 90 Days',data:top.map(d=>d.due_90_value),backgroundColor:'#f59e0b88',borderColor:'#f59e0b',borderWidth:1,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8',font:{size:11}}},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+fmt(ctx.parsed.y)}}},scales:{x:{ticks:{color:'#94a3b8',maxRotation:30},grid:{color:'#2e3350'}},y:{ticks:{color:'#94a3b8',callback:v=>fmt(v)},grid:{color:'#2e3350'}}}}});
  const tb=document.getElementById('execTableBody');
  const maxV=DATA.exec_data[0]?DATA.exec_data[0].value:1;
  DATA.exec_data.forEach(d=>{tb.innerHTML+=`<tr><td>${d['Execution Person']}</td><td class="num">${fmtNum(d.lines)}</td><td class="num">${fmtFull(d.value)}</td><td class="num red-text">${fmtFull(d.past_due_value)}</td><td class="num amber-text">${fmtFull(d.due_90_value)}</td><td class="num">${pct(d.pct)}</td><td><div class="progress-cell"><div class="bar-bg" style="flex:1"><div class="bar-fill" style="width:${Math.max(d.value/maxV*100,.5)}%"></div></div><span class="pct-lbl">${d.pct}%</span></div></td></tr>`;});
}

let _repFilter='all';
function buildRepCards(filter){
  _repFilter=filter;
  const grid=document.getElementById('repGrid');
  grid.innerHTML='';
  const execs={};
  DATA.rep_data.forEach(r=>{
    const k=r['Execution Person'];
    if(!execs[k])execs[k]={name:k,total_lines:r.total_lines,total_value:r.total_value,items:[]};
    if(filter==='all'||r['ActionBy - New']===filter)execs[k].items.push(r);
  });
  Object.values(execs).filter(r=>r.items.length>0)
    .sort((a,b)=>b.items.reduce((s,x)=>s+x.value,0)-a.items.reduce((s,x)=>s+x.value,0))
    .forEach(rep=>{
      const totalRiskVal=rep.items.reduce((s,x)=>s+x.value,0);
      const pctBook=(totalRiskVal/rep.total_value*100);
      const risksHTML=rep.items.map(item=>{
        const cls=item['ActionBy - New']==='P&E'?'pe':'sps';
        const pdPct =item.value>0?item.past_due_value/item.value*100:0;
        const todPct=item.value>0?(item.today_value||0)/item.value*100:0;
        const d90Pct=item.value>0?item.due_90_value/item.value*100:0;
        const futPct=item.value>0?item.future_value/item.value*100:0;
        return `<div class="risk-block">
          <div class="risk-title ${cls}">${item['ActionBy - New']}</div>
          <div class="risk-metrics">
            <div><div class="risk-metric-label">SOs</div><div class="risk-metric-val">${fmtNum(item.so_count)}</div></div>
            <div><div class="risk-metric-label">Lines</div><div class="risk-metric-val">${fmtNum(item.lines)}</div></div>
            <div><div class="risk-metric-label">Financial Impact</div><div class="risk-metric-val">${fmt(item.value)}</div></div>
            <div><div class="risk-metric-label">% of Total Book</div><div class="risk-metric-val">${item.pct_of_exec_total}%</div></div>
          </div>
          <div style="margin-top:10px">
            <div class="gauge-row"><div class="gauge-label">Past Due</div><div class="gauge-bar"><div class="gauge-fill" style="width:${Math.min(pdPct,100)}%;background:#ef4444"></div></div><div class="gauge-pct red-text">${pdPct.toFixed(1)}%</div></div>
            ${todPct>0?`<div class="gauge-row"><div class="gauge-label">Due Today</div><div class="gauge-bar"><div class="gauge-fill" style="width:${Math.min(todPct,100)}%;background:#a855f7"></div></div><div class="gauge-pct" style="color:#a855f7">${todPct.toFixed(1)}%</div></div>`:''}
            <div class="gauge-row"><div class="gauge-label">Next 90d</div><div class="gauge-bar"><div class="gauge-fill" style="width:${Math.min(d90Pct,100)}%;background:#f59e0b"></div></div><div class="gauge-pct amber-text">${d90Pct.toFixed(1)}%</div></div>
            <div class="gauge-row"><div class="gauge-label">Future</div><div class="gauge-bar"><div class="gauge-fill" style="width:${Math.min(futPct,100)}%;background:#22c55e"></div></div><div class="gauge-pct green-text">${futPct.toFixed(1)}%</div></div>
          </div>
        </div>`;
      }).join('');
      grid.innerHTML+=`<div class="rep-card"><div class="rep-name">${rep.name}</div><div class="rep-total">Total assigned: ${fmt(rep.total_value)} &middot; ${fmtNum(rep.total_lines)} lines &nbsp;|&nbsp; P&amp;E / SPS Risk: ${fmt(totalRiskVal)} (${pctBook.toFixed(1)}% of book)</div>${risksHTML}</div>`;
    });
}

function filterReps(type,btn){document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');buildRepCards(type);}

const _built={};
function showTab(name,el){
  document.querySelectorAll('.nav-tab').forEach(t=>t.classList.remove('active'));el.classList.add('active');
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));document.getElementById('tab-'+name).classList.add('active');
  if(!_built[name]){_built[name]=true;if(name==='functions')buildFunctionCharts();if(name==='sites')buildSiteCharts();if(name==='execution')buildExecCharts();if(name==='reps')buildRepCards('all');}
}
buildOverviewCharts();
</script>
</body>
</html>"""


# =============================================================================
#  STEP 5 - INJECT DATA AND WRITE FILE
# =============================================================================
def build_html(kpis, func_data, site_data, exec_data, rep_data, today):
    total      = kpis["total_book_value"]
    future_val = total - kpis["total_past_due_value"] \
                       - kpis["total_today_value"] \
                       - kpis["total_due_90_value"]

    data_json = json.dumps({
        "kpis":      kpis,
        "func_data": func_data,
        "site_data": site_data,
        "exec_data": exec_data,
        "rep_data":  rep_data,
    }, ensure_ascii=False)

    html = HTML
    html = html.replace("__AS_OF__",          fmt_date(today))
    html = html.replace("__DATA_JSON__",       data_json)
    html = html.replace("__KPI_BOOK__",        fmt_short(total))
    html = html.replace("__KPI_LINES__",       f"{kpis['total_lines']:,}")
    html = html.replace("__KPI_PD_VAL__",      fmt_short(kpis["total_past_due_value"]))
    html = html.replace("__KPI_PD_LINES__",    f"{kpis['total_past_due_lines']:,}")
    html = html.replace("__KPI_PD_PCT__",      pct_of(kpis["total_past_due_value"], total))
    html = html.replace("__KPI_TOD_VAL__",     fmt_short(kpis["total_today_value"]))
    html = html.replace("__KPI_TOD_LINES__",   f"{kpis['total_today_lines']:,}")
    html = html.replace("__KPI_TOD_PCT__",     pct_of(kpis["total_today_value"], total))
    html = html.replace("__KPI_90_VAL__",      fmt_short(kpis["total_due_90_value"]))
    html = html.replace("__KPI_90_LINES__",    f"{kpis['total_due_90_lines']:,}")
    html = html.replace("__KPI_90_PCT__",      pct_of(kpis["total_due_90_value"], total))
    html = html.replace("__KPI_FUT_VAL__",     fmt_short(future_val))
    html = html.replace("__KPI_FUT_LINES__",   f"{kpis['total_future_lines']:,}")
    html = html.replace("__KPI_FUT_PCT__",     pct_of(future_val, total))
    return html


# =============================================================================
#  MAIN
# =============================================================================
def main():
    print("=" * 62)
    print("  Honeywell Order Book Dashboard Generator")
    print("=" * 62)

    df = load_data()
    print("Calculating metrics ...")
    kpis, func_data, site_data, exec_data, rep_data, today = calculate(df)

    print("Building index.html ...")
    html = build_html(kpis, func_data, site_data, exec_data, rep_data, today)
    OUTPUT_FILE.write_text(html, encoding="utf-8")

    print(f"""
Done!
  Output:  {OUTPUT_FILE}

Next steps:
  1. Open index.html in your browser to verify it looks correct.
  2. Copy / commit index.html to your GitHub repo root.
  3. GitHub Pages will serve the updated dashboard automatically.
{"=" * 62}""")


if __name__ == "__main__":
    main()
