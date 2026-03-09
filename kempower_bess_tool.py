import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import os
import json
from pathlib import Path
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

# ─────────────────────────────────────────────
#  BRANDING & CONFIG
# ─────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent if "__file__" in locals() else Path.cwd()
LOGO_PATH  = SCRIPT_DIR / "kempowerlogo.png"
KEMPOWER_ORANGE = "#FF6400"
KEMPOWER_DARK   = "#1A1A1A"

st.set_page_config(page_title="Kempower | BESS Sizing", layout="wide", page_icon="⚡")

st.markdown(f"""
<style>
  #MainMenu, footer, header {{visibility: hidden;}}
  .block-container {{padding-top: 0.8rem; padding-bottom: 1rem;}}
  [data-testid="stSidebar"] {{background: #F2F2F2; color: #1a1a1a;}}
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] .stMarkdown p,
  [data-testid="stSidebar"] .stMarkdown h3 {{color: #1a1a1a !important;}}
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stNumberInput label,
  [data-testid="stSidebar"] .stSlider label {{color: #333 !important;}}
  div[data-testid="metric-container"] {{
    background: #fff; border-left: 4px solid {KEMPOWER_ORANGE};
    border-radius: 8px; padding: 12px 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
  }}
  div[data-testid="metric-container"] label {{color: #666 !important; font-size:0.78rem !important;}}
  div[data-testid="metric-container"] [data-testid="stMetricValue"] {{
    color: {KEMPOWER_DARK} !important; font-size:1.5rem !important; font-weight:700 !important;
  }}
  .qos-box {{background:#fff; border-radius:8px; padding:14px 18px; box-shadow:0 2px 8px rgba(0,0,0,0.06);}}
  .qos-row {{display:flex; align-items:flex-start; gap:10px; margin-bottom:8px; font-size:0.82rem; color:#333;}}
  .qos-dot {{width:10px; height:10px; border-radius:50%; flex-shrink:0; margin-top:3px;}}
  .fin-card {{background:#fff; border-radius:10px; padding:16px 20px;
              box-shadow: 0 2px 10px rgba(0,0,0,0.07); margin-bottom:12px;}}
  .fin-card h4 {{margin:0 0 4px 0; font-size:0.85rem; color:#666; font-weight:500;}}
  .fin-card .val {{font-size:1.55rem; font-weight:700; color:{KEMPOWER_DARK};}}
  .fin-card .sub {{font-size:0.76rem; color:#888; margin-top:2px;}}
  .positive {{color: #2e7d32 !important;}}
  .negative {{color: #c62828 !important;}}
  .stTabs [data-baseweb="tab-list"] {{gap: 4px;}}
  .stTabs [data-baseweb="tab"] {{
    background: #f0f0f0; border-radius: 8px 8px 0 0;
    padding: 8px 20px; font-weight: 600; color: #555;
  }}
  .stTabs [aria-selected="true"] {{background: {KEMPOWER_ORANGE} !important; color: #fff !important;}}
  .section-header {{
    font-size: 0.68rem; font-weight: 700; letter-spacing: 0.12em;
    text-transform: uppercase; color: {KEMPOWER_ORANGE}; margin: 16px 0 6px 0;
  }}
  [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
    background: #e8e8e8 !important;
    border: 1.5px dashed {KEMPOWER_ORANGE} !important;
    border-radius: 8px !important;
  }}
  [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span,
  [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {{
    color: #555 !important;
  }}
  [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {{
    background: {KEMPOWER_ORANGE} !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
  }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  POLARIUM BESS PRODUCT MAP  (kWh → max kW)
# ─────────────────────────────────────────────
BESS_POWER_MAP = {
    # Polarium: 1 cabinet = 140 kWh + 1 inverter = 75 kW
    140:   75,   # 1 cabinet
    280:  150,   # 2 cabinets
    420:  225,   # 3 cabinets
    560:  300,   # 4 cabinets
    700:  375,   # 5 cabinets
    840:  450,   # 6 cabinets
    1120: 600,   # 8 cabinets
    1400: 750,   # 10 cabinets
}

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=180)
    else:
        st.markdown(f"<h2 style='color:{KEMPOWER_ORANGE}'>⚡ Kempower</h2>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🏢 Site Parameters")
    grid_limit   = st.number_input("Grid Limit (kW)", value=150, min_value=0, step=10)
    num_plugs    = st.number_input("Satellites / Plugs", value=10, min_value=1)
    charger_cap  = st.number_input("Site Charger Capacity (kW)", value=600, min_value=0, step=50,
                                    help="Physical maximum output of all chargers combined")

    st.markdown("### 🔋 BESS (Polarium)")
    cap_options    = sorted(BESS_POWER_MAP.keys())
    nominal_cap    = st.selectbox("BESS Capacity (kWh)", cap_options, index=cap_options.index(560))
    bess_max_power = BESS_POWER_MAP[nominal_cap]
    st.caption(f"↳ Max discharge: **{bess_max_power} kW**")
    usable_factor  = st.slider("Usable SoC Window", 0.70, 0.95, 0.85, 0.01,
                                help="e.g. 0.85 = 85% depth of discharge")

    st.markdown("### 📈 Simulation")
    load_multiplier = st.slider("Load Scaling Factor", 0.25, 2.0, 1.0, 0.05)
    growth_rate     = st.slider("Annual Demand Growth (%)", 0, 20, 5) / 100

    st.markdown("### 💶 Financial Assumptions")
    bess_capex_per_kwh = st.number_input("BESS CAPEX (€/kWh)", value=450, min_value=100, step=10,
                                          help="All-in installed cost per kWh")
    bess_opex_pct      = st.slider("Annual OPEX (% of CAPEX)", 0.5, 3.0, 1.5, 0.1) / 100
    grid_upgrade_cost  = st.number_input("Grid Upgrade Cost (€)", value=80_000, step=5000,
                                          help="One-off cost to raise grid connection capacity")
    demand_fee_per_kw  = st.number_input("Demand Charge (€/kW/month)", value=12.0, step=0.5,
                                          help="Monthly fee charged on peak kW drawn from grid")
    energy_price_eur   = st.number_input("Energy Price (€/kWh)", value=0.18, step=0.01)
    discount_rate      = st.slider("Discount Rate (%)", 0, 15, 7) / 100
    analysis_years     = st.slider("Financial Horizon (years)", 5, 15, 10)

    st.markdown("---")
    st.markdown("### 📂 Load Profile")
    uploaded_file = st.file_uploader(
        "Upload CSV or XLSX (7-day, 5-min intervals)",
        type=["csv", "xlsx"],
        label_visibility="collapsed"
    )


# ─────────────────────────────────────────────
#  SIMULATION ENGINE
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def run_simulation(csv_bytes, grid_limit, charger_cap, nominal_cap, bess_max_power,
                   usable_factor, load_multiplier, growth_rate):
    try:
        df_base = pd.read_csv(io.BytesIO(csv_bytes), sep=None, engine='python')
    except Exception:
        df_base = pd.read_csv(io.BytesIO(csv_bytes), sep=';')

    if df_base.shape[1] < 2:
        df_base = pd.read_csv(io.BytesIO(csv_bytes), sep=';')

    df_base.columns = df_base.columns.str.strip()
    col_map = {df_base.columns[0]: 'timestamp', df_base.columns[1]: 'raw_load'}
    if 'Timestamp'  in df_base.columns: col_map['Timestamp']  = 'timestamp'
    if 'Power [kW]' in df_base.columns: col_map['Power [kW]'] = 'raw_load'
    df_base.rename(columns=col_map, inplace=True)
    df_base['timestamp'] = pd.to_datetime(df_base['timestamp'], dayfirst=True)
    df_base['day_name']  = df_base['timestamp'].dt.day_name()

    step_hrs              = 5 / 60
    year_mult             = 365 / 7
    current_usable_kwh    = nominal_cap * usable_factor
    initial_kwh           = current_usable_kwh
    total_lifetime_cycles = 0
    yearly_results        = []
    plot_data             = {}

    for year in range(1, 11):
        df = df_base.copy()
        df['load_theoretical'] = (df['raw_load'] * load_multiplier
                                   * ((1 + growth_rate) ** (year - 1)))
        df['load_to_serve']    = np.minimum(df['load_theoretical'], charger_cap)

        n          = len(df)
        grid_used  = np.zeros(n)
        bess_disc  = np.zeros(n)
        bess_char  = np.zeros(n)
        missed_energy_yr = 0
        current_soc      = current_usable_kwh * 0.5

        for i in range(n):
            demand = df['load_to_serve'].iloc[i]
            if demand <= grid_limit:
                grid_used[i] = demand
                headroom      = grid_limit - demand
                charge_p      = min(bess_max_power, headroom,
                                    (current_usable_kwh - current_soc) / step_hrs)
                bess_char[i]  = max(0.0, charge_p)
            else:
                grid_used[i]  = grid_limit
                shortfall      = demand - grid_limit
                discharge_p    = min(bess_max_power, shortfall, current_soc / step_hrs)
                bess_disc[i]   = max(0.0, discharge_p)

            current_soc += (bess_char[i] * 0.85 - bess_disc[i] / 0.85) * step_hrs
            current_soc  = float(np.clip(current_soc, 0, current_usable_kwh))

            unmet = df['load_theoretical'].iloc[i] - (grid_used[i] + bess_disc[i])
            if unmet > 0.001:
                missed_energy_yr += unmet * step_hrs

        ev_mwh_yr     = float(df['load_theoretical'].sum()) * step_hrs * year_mult / 1000
        missed_mwh_yr = missed_energy_yr * year_mult / 1000
        bess_mwh_yr   = float(sum(bess_disc)) * step_hrs * year_mult / 1000
        grid_mwh_yr   = float(sum(grid_used)) * step_hrs * year_mult / 1000
        yearly_cycles = bess_mwh_yr / (current_usable_kwh / 1000) if current_usable_kwh > 0 else 0
        total_lifetime_cycles += yearly_cycles
        peak_grid_kw  = float(max(grid_used))

        yearly_results.append({
            "Year":         year,
            "EV MWh":       round(ev_mwh_yr, 1),
            "Grid MWh":     round(grid_mwh_yr, 1),
            "BESS MWh":     round(bess_mwh_yr, 1),
            "Missed MWh":   round(missed_mwh_yr, 1),
            "Unmet %":      round((missed_mwh_yr / ev_mwh_yr * 100) if ev_mwh_yr > 0 else 0, 2),
            "Peak Grid kW": round(peak_grid_kw, 0),
            "SoH %":        round((current_usable_kwh / initial_kwh) * 100, 1),
            "Cycles":       round(yearly_cycles, 0),
        })

        df['grid_used'] = grid_used
        df['bess_disc'] = bess_disc
        plot_data[year] = df.copy()

        # Non-linear degradation (staged by cycle count)
        for _ in range(int(yearly_cycles)):
            rate = (0.0045 if total_lifetime_cycles <= 800
                    else 0.0018 if total_lifetime_cycles <= 2500
                    else 0.0055)
            current_usable_kwh *= (1 - rate / 100)
            total_lifetime_cycles += 1

    return df_base, yearly_results, plot_data, initial_kwh, current_usable_kwh, total_lifetime_cycles


def load_uploaded_file(f):
    """Read CSV or XLSX without pd.read_excel — avoids the openpyxl engine ImportError on Streamlit Cloud."""
    name = f.name.lower()
    if name.endswith(('.xlsx', '.xls')):
        import openpyxl
        raw = f.read()
        wb  = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws  = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            raise ValueError("Empty Excel file")
        headers = [str(c) if c is not None else f"col_{i}" for i, c in enumerate(rows[0])]
        df = pd.DataFrame(rows[1:], columns=headers)
        buf = io.BytesIO()
        df.to_csv(buf, index=False)
        return buf.getvalue()
    return f.read()


# ─────────────────────────────────────────────
#  FINANCIAL MODEL
# ─────────────────────────────────────────────
def compute_financials(yearly_results, nominal_cap, bess_capex_per_kwh, bess_opex_pct,
                        grid_upgrade_cost, demand_fee_per_kw, energy_price_eur,
                        discount_rate, analysis_years, grid_limit):
    bess_capex   = nominal_cap * bess_capex_per_kwh
    annual_opex  = bess_capex * bess_opex_pct
    rows         = []
    cum_npv      = -bess_capex

    for yr_idx in range(1, analysis_years + 1):
        res = yearly_results[min(yr_idx, len(yearly_results)) - 1]

        # Peak shaving: every kW kept below grid_limit saves a monthly demand charge
        peak_kw_saved  = max(0.0, grid_limit - res["Peak Grid kW"])
        demand_saving  = peak_kw_saved * demand_fee_per_kw * 12

        # Energy arbitrage: BESS output displaces peak-tariff energy (10% premium assumed)
        arbitrage_saving = res["BESS MWh"] * energy_price_eur * 1000 * 0.10

        # Avoided grid upgrade amortised linearly
        avoided_upgrade  = grid_upgrade_cost / analysis_years

        total_benefit  = demand_saving + arbitrage_saving + avoided_upgrade
        net_cf         = total_benefit - annual_opex
        discount       = (1 + discount_rate) ** yr_idx
        npv_contrib    = net_cf / discount
        cum_npv       += npv_contrib

        rows.append({
            "Year":                  yr_idx,
            "Demand Saving (€)":     round(demand_saving),
            "Arbitrage Saving (€)":  round(arbitrage_saving),
            "Avoided Upgrade (€)":   round(avoided_upgrade),
            "Total Benefit (€)":     round(total_benefit),
            "OPEX (€)":              round(annual_opex),
            "Net Cash Flow (€)":     round(net_cf),
            "NPV Contrib (€)":       round(npv_contrib),
            "Cumulative NPV (€)":    round(cum_npv),
        })

    fin_df = pd.DataFrame(rows)

    # Simple payback
    cum_simple, payback_yr = -bess_capex, None
    for r in rows:
        cum_simple += r["Net Cash Flow (€)"]
        if cum_simple >= 0 and payback_yr is None:
            payback_yr = r["Year"]

    # IRR via bisection
    cash_flows = [-bess_capex] + [r["Net Cash Flow (€)"] for r in rows]
    def npv_r(r):
        return sum(cf / (1 + r) ** t for t, cf in enumerate(cash_flows))
    lo, hi = -0.5, 5.0
    for _ in range(120):
        mid = (lo + hi) / 2
        if npv_r(mid) > 0:
            lo = mid
        else:
            hi = mid
    irr = lo * 100

    return fin_df, bess_capex, payback_yr, round(irr, 1), round(fin_df["Cumulative NPV (€)"].iloc[-1])


# ─────────────────────────────────────────────
#  PDF EXPORT
# ─────────────────────────────────────────────
def generate_pdf(yearly_results, fin_df, site_params, bess_capex, payback_yr, irr, total_npv):
    buf    = io.BytesIO()
    doc    = SimpleDocTemplate(buf, pagesize=A4,
                                leftMargin=1.8*cm, rightMargin=1.8*cm,
                                topMargin=2*cm,    bottomMargin=2*cm)
    els    = []
    styles = getSampleStyleSheet()
    orange = ParagraphStyle('or', parent=styles['Heading2'],
                             textColor=colors.HexColor(KEMPOWER_ORANGE))

    if LOGO_PATH.exists():
        els += [RLImage(str(LOGO_PATH), width=130, height=43), Spacer(1, 8)]

    els += [Paragraph("BESS Sizing & Financial Report", styles['Title']), Spacer(1, 6)]

    # Params
    els.append(Paragraph("Site & Simulation Parameters", orange))
    pd_data = [
        ["Parameter", "Value", "Parameter", "Value"],
        ["Grid Limit",        f"{site_params['grid_limit']} kW",
         "BESS Capacity",     f"{site_params['nominal_cap']} kWh (Polarium)"],
        ["Charger Capacity",  f"{site_params['charger_cap']} kW",
         "Usable SoC",        f"{int(site_params['usable_factor']*100)}%"],
        ["Satellites/Plugs",  str(site_params['num_plugs']),
         "Load Scaling",      f"{site_params['load_multiplier']}×"],
        ["Annual Growth",     f"{site_params['growth_rate']*100:.1f}%",
         "BESS CAPEX",        f"€{bess_capex:,.0f}"],
    ]
    t = Table(pd_data, colWidths=[4.5*cm, 3.5*cm, 5*cm, 3.5*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,0), colors.HexColor("#333")),
        ('TEXTCOLOR',     (0,0), (-1,0), colors.whitesmoke),
        ('FONTSIZE',      (0,0), (-1,-1), 8),
        ('GRID',          (0,0), (-1,-1), 0.4, colors.lightgrey),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, colors.HexColor("#FFF5EE")]),
    ]))
    els += [t, Spacer(1, 12)]

    # Sim results
    els.append(Paragraph("10-Year Simulation Results", orange))
    cols  = ["Year","EV MWh","Grid MWh","BESS MWh","Missed MWh","Unmet %","Peak Grid kW","SoH %"]
    tdata = [cols] + [[str(r[c]) for c in cols] for r in yearly_results]
    t2    = Table(tdata, colWidths=[1*cm,2*cm,2.3*cm,2.3*cm,2.5*cm,1.7*cm,2.8*cm,1.8*cm])
    t2.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,0), colors.HexColor(KEMPOWER_ORANGE)),
        ('TEXTCOLOR',     (0,0), (-1,0), colors.white),
        ('ALIGN',         (0,0), (-1,-1), 'CENTER'),
        ('FONTSIZE',      (0,0), (-1,-1), 8),
        ('GRID',          (0,0), (-1,-1), 0.4, colors.lightgrey),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, colors.HexColor("#FFF5EE")]),
    ]))
    els += [t2, Spacer(1, 12)]

    # Financial summary
    els.append(Paragraph("Financial Summary", orange))
    fs = [
        ["Metric", "Value"],
        ["BESS CAPEX",                    f"€{bess_capex:,.0f}"],
        ["Simple Payback",                f"{payback_yr} years" if payback_yr else "Outside horizon"],
        ["IRR",                           f"{irr:.1f}%"],
        [f"NPV ({len(fin_df)} years)",    f"€{total_npv:,.0f}"],
    ]
    t3 = Table(fs, colWidths=[8*cm, 8.4*cm])
    t3.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,0), colors.HexColor("#333")),
        ('TEXTCOLOR',     (0,0), (-1,0), colors.whitesmoke),
        ('FONTSIZE',      (0,0), (-1,-1), 9),
        ('GRID',          (0,0), (-1,-1), 0.4, colors.lightgrey),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, colors.HexColor("#FFF5EE")]),
    ]))
    els += [t3, Spacer(1, 12)]

    # Cash flow detail
    els.append(Paragraph("Annual Cash Flow Detail", orange))
    fc   = ["Year","Demand Saving (€)","Avoided Upgrade (€)","OPEX (€)","Net Cash Flow (€)","Cumulative NPV (€)"]
    fdata= [fc] + [[str(r[c]) for c in fc] for _, r in fin_df.iterrows()]
    t4   = Table(fdata, colWidths=[1*cm,3.5*cm,3.5*cm,2*cm,3.5*cm,3*cm])
    t4.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,0), colors.HexColor(KEMPOWER_ORANGE)),
        ('TEXTCOLOR',     (0,0), (-1,0), colors.white),
        ('ALIGN',         (0,0), (-1,-1), 'CENTER'),
        ('FONTSIZE',      (0,0), (-1,-1), 8),
        ('GRID',          (0,0), (-1,-1), 0.4, colors.lightgrey),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [colors.white, colors.HexColor("#FFF5EE")]),
    ]))
    els.append(t4)
    doc.build(els)
    return buf.getvalue()


# ─────────────────────────────────────────────
#  SMART CSV EXPORT
# ─────────────────────────────────────────────
def generate_smart_export(df_base, metadata):
    buffer = io.StringIO()
    buffer.write(f"# METADATA_JSON:{json.dumps(metadata)}\n")
    export_df = df_base[['timestamp', 'raw_load']].copy()
    export_df.columns = ['timestamp', 'load_kw']
    export_df['timestamp'] = export_df['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
    export_df.to_csv(buffer, index=False)
    return buffer.getvalue().encode('utf-8')


# ─────────────────────────────────────────────
#  LANDING PAGE (no file uploaded)
# ─────────────────────────────────────────────
if uploaded_file is None:
    st.markdown(f"""
    <div style="text-align:center; padding:80px 0;">
      <div style="font-size:3rem;">⚡</div>
      <h1 style="color:{KEMPOWER_DARK}; font-weight:800; margin-bottom:8px;">
        Kempower | BESS Sizing Tool
      </h1>
      <p style="color:#666; font-size:1.05rem; max-width:500px; margin:0 auto 24px;">
        Upload a 7-day ChargEye load profile (CSV or XLSX) in the sidebar to begin sizing your
        <strong>Polarium BESS</strong> solution.
      </p>
      <p style="color:#999; font-size:0.85rem;">
        5-minute interval data · 10-year degradation model · Financial ROI analysis
      </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────
#  RUN SIMULATION
# ─────────────────────────────────────────────
file_bytes = load_uploaded_file(uploaded_file)

with st.spinner("Running 10-year simulation…"):
    (df_base, yearly_results, plot_data,
     initial_kwh, final_kwh, total_cycles) = run_simulation(
        file_bytes, grid_limit, charger_cap, nominal_cap, bess_max_power,
        usable_factor, load_multiplier, growth_rate
    )

total_ev_mwh     = sum(r["EV MWh"]   for r in yearly_results)
total_missed_mwh = sum(r["Missed MWh"] for r in yearly_results)
total_bess_mwh   = sum(r["BESS MWh"] for r in yearly_results)
final_soh        = (final_kwh / initial_kwh) * 100
unmet_pct        = (total_missed_mwh / total_ev_mwh * 100) if total_ev_mwh > 0 else 0

fin_df, bess_capex, payback_yr, irr, total_npv = compute_financials(
    yearly_results, nominal_cap, bess_capex_per_kwh, bess_opex_pct,
    grid_upgrade_cost, demand_fee_per_kw, energy_price_eur,
    discount_rate, analysis_years, grid_limit
)

res_df_table = pd.DataFrame(yearly_results)

# ─────────────────────────────────────────────
#  TOP KPI STRIP
# ─────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("BESS (Polarium)",   f"{nominal_cap} kWh",  f"Max: {bess_max_power} kW")
k2.metric("10yr EV Demand",    f"{total_ev_mwh:,.0f} MWh")
k3.metric("10yr BESS Output",  f"{total_bess_mwh:,.0f} MWh")
k4.metric("Unmet Demand",      f"{unmet_pct:.2f}%",
          "✅ Excellent" if unmet_pct < 2 else ("⚠️ Warning" if unmet_pct < 5 else "❌ Critical"))
k5.metric("Final SoH",         f"{final_soh:.1f}%",  f"{total_cycles:,.0f} cycles")

st.markdown("---")

# ─────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "📊  Load Profile & Simulation",
    "📋  Annual Results Table",
    "💶  Financial Analysis",
])

# ══════════════════════════════════════════
#  TAB 1 — Load Profile
# ══════════════════════════════════════════
with tab1:
    chart_col, qos_col = st.columns([4, 1])

    with qos_col:
        st.markdown('<p class="section-header">Service Quality</p>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="qos-box">
          <div class="qos-row"><div class="qos-dot" style="background:#2e7d32"></div>
            <span><b>&lt;2% Excellent</b><br>All EV sessions fully served</span></div>
          <div class="qos-row"><div class="qos-dot" style="background:#f9a825"></div>
            <span><b>2–5% Warning</b><br>Risk of throttled sessions at peak</span></div>
          <div class="qos-row"><div class="qos-dot" style="background:#c62828"></div>
            <span><b>&gt;5% Critical</b><br>Unmet demand exceeds threshold</span></div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<p class="section-header">Year-1 Summary</p>', unsafe_allow_html=True)
        yr1 = yearly_results[0]
        st.metric("EV Demand Y1",   f"{yr1['EV MWh']} MWh")
        st.metric("Grid Supplied",  f"{yr1['Grid MWh']} MWh")
        st.metric("BESS Supplied",  f"{yr1['BESS MWh']} MWh")
        st.metric("Peak Grid Draw", f"{yr1['Peak Grid kW']:.0f} kW")

    with chart_col:
        view_yr = st.selectbox("View year:", list(range(1, 11)))
        p_df    = plot_data[view_yr]

        day_order  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        unique_days, day_indices = [], []
        for day in day_order:
            idx = p_df.index[p_df['day_name'] == day].tolist()
            if idx:
                day_indices.append(idx[len(idx) // 2])
                unique_days.append(day[:3])

        fig = go.Figure()
        fig.add_trace(go.Bar(name='Grid', x=p_df.index, y=p_df['grid_used'],
                              marker_color='#2b579a', opacity=0.85,
                              hovertemplate='Grid: %{y:.0f} kW<extra></extra>'))
        fig.add_trace(go.Bar(name='BESS Supplement', x=p_df.index, y=p_df['bess_disc'],
                              marker_color='#7eb26d',
                              hovertemplate='BESS: %{y:.0f} kW<extra></extra>'))
        fig.add_trace(go.Scatter(name='EV Demand', x=p_df.index, y=p_df['load_theoretical'],
                                  line=dict(color=KEMPOWER_ORANGE, width=2, dash='dot'),
                                  hovertemplate='Demand: %{y:.0f} kW<extra></extra>'))
        fig.add_shape(type="line", x0=p_df.index[0], x1=p_df.index[-1],
                       y0=charger_cap, y1=charger_cap,
                       line=dict(color="red", width=1, dash="dashdot"))
        fig.add_annotation(x=p_df.index[-1], y=charger_cap, text=f"Charger cap {charger_cap} kW",
                            showarrow=False, xanchor="right", font=dict(size=10, color="red"), yshift=8)
        fig.add_shape(type="line", x0=p_df.index[0], x1=p_df.index[-1],
                       y0=grid_limit, y1=grid_limit,
                       line=dict(color="#2b579a", width=1.5, dash="dash"))
        fig.add_annotation(x=p_df.index[-1], y=grid_limit, text=f"Grid limit {grid_limit} kW",
                            showarrow=False, xanchor="right", font=dict(size=10, color="#2b579a"), yshift=8)
        fig.update_layout(
            barmode='stack', height=420,
            margin=dict(l=10, r=10, t=30, b=50),
            legend=dict(orientation='h', y=-0.18),
            xaxis=dict(tickvals=day_indices, ticktext=unique_days,
                        tickfont=dict(size=12), showgrid=False),
            yaxis=dict(title="Power (kW)", gridcolor="#eee"),
            plot_bgcolor='white', paper_bgcolor='white',
            title=dict(text=f"Year {view_yr} — 7-day Load Profile",
                        font=dict(size=13, color=KEMPOWER_DARK))
        )
        st.plotly_chart(fig, use_container_width=True)

    # SoH degradation curve
    st.markdown('<p class="section-header">Battery State-of-Health Degradation</p>',
                unsafe_allow_html=True)
    soh_vals = [r["SoH %"] for r in yearly_results]
    soh_fig  = go.Figure()
    soh_fig.add_trace(go.Scatter(
        x=list(range(1, 11)), y=soh_vals, mode='lines+markers',
        line=dict(color=KEMPOWER_ORANGE, width=3),
        marker=dict(size=8, color=KEMPOWER_ORANGE),
        fill='tozeroy', fillcolor='rgba(255,100,0,0.08)',
        hovertemplate='Year %{x}: %{y:.1f}%<extra></extra>'
    ))
    soh_fig.add_hline(y=80, line_dash="dash", line_color="#c62828",
                       annotation_text="80% SoH threshold",
                       annotation_position="bottom right")
    soh_fig.update_layout(
        height=200, margin=dict(l=10, r=10, t=10, b=30),
        xaxis=dict(title="Year", tickvals=list(range(1, 11))),
        yaxis=dict(title="SoH (%)", range=[60, 102]),
        plot_bgcolor='white', paper_bgcolor='white'
    )
    st.plotly_chart(soh_fig, use_container_width=True)


# ══════════════════════════════════════════
#  TAB 2 — Annual Results Table
# ══════════════════════════════════════════
with tab2:
    st.markdown('<p class="section-header">10-Year Simulation Results</p>', unsafe_allow_html=True)

    def color_unmet(val):
        if val < 2:   return 'color:#2e7d32; font-weight:600'
        elif val < 5: return 'color:#f57f17; font-weight:600'
        return               'color:#c62828; font-weight:600'

    def highlight_bess(val):
        # Light orange tint scaled to value — no matplotlib needed
        try:
            max_val = res_df_table["BESS MWh"].max()
            intensity = int(255 - (val / max_val) * 120) if max_val > 0 else 255
            return f'background-color: rgb(255, {intensity}, {intensity}); color: #1a1a1a'
        except Exception:
            return ''

    styled = (res_df_table.style
              .format({"Unmet %": "{:.2f}%", "SoH %": "{:.1f}%",
                       "EV MWh": "{:.1f}", "Grid MWh": "{:.1f}",
                       "BESS MWh": "{:.1f}", "Missed MWh": "{:.1f}"})
              .applymap(color_unmet, subset=["Unmet %"])
              .applymap(highlight_bess, subset=["BESS MWh"])
              .set_properties(**{'font-size': '0.85rem'}))
    st.dataframe(styled, use_container_width=True, height=400)

    site_params_dict = dict(grid_limit=grid_limit, nominal_cap=nominal_cap,
                             charger_cap=charger_cap, usable_factor=usable_factor,
                             num_plugs=num_plugs, load_multiplier=load_multiplier,
                             growth_rate=growth_rate)
    dl1, dl2 = st.columns(2)
    with dl1:
        pdf_bytes = generate_pdf(yearly_results, fin_df, site_params_dict,
                                  bess_capex, payback_yr, irr, total_npv)
        st.download_button("📥 Download PDF Report", pdf_bytes,
                            "Kempower_BESS_Report.pdf", "application/pdf",
                            use_container_width=True)
    with dl2:
        meta      = dict(grid_limit=grid_limit, bess_capacity=nominal_cap,
                          charger_cap=charger_cap, growth_rate=growth_rate*100,
                          load_multiplier=load_multiplier, total_cycles=int(total_cycles),
                          final_soh=round(final_soh, 1))
        csv_data  = generate_smart_export(df_base, meta)
        st.download_button("📊 Export for Trade-Off Tool", csv_data,
                            "kempower_smart_export.csv", "text/csv",
                            use_container_width=True)


# ══════════════════════════════════════════
#  TAB 3 — Financial Analysis
# ══════════════════════════════════════════
with tab3:
    st.markdown(f"""
    <div style="background:{KEMPOWER_DARK}; border-radius:10px; padding:14px 22px; margin-bottom:18px;">
      <h3 style="color:#fff; margin:0; font-size:1rem;">
        💡 BESS Investment vs. Grid Upgrade — {analysis_years}-Year Business Case
      </h3>
      <p style="color:#aaa; margin:4px 0 0; font-size:0.82rem;">
        Polarium {nominal_cap} kWh ·  CAPEX €{bess_capex:,.0f} ·
        Demand charge €{demand_fee_per_kw}/kW/mo ·
        Grid upgrade avoided: €{grid_upgrade_cost:,.0f}
      </p>
    </div>
    """, unsafe_allow_html=True)

    # KPI cards
    fc1, fc2, fc3, fc4 = st.columns(4)
    npv_cls = "positive" if total_npv >= 0 else "negative"
    irr_cls = "positive" if irr >= discount_rate * 100 else "negative"
    pb_cls  = "positive" if payback_yr and payback_yr <= 7 else "negative"
    pb_txt  = f"{payback_yr} yrs" if payback_yr else "Outside horizon"

    fc1.markdown(f"""<div class="fin-card">
      <h4>BESS Total CAPEX</h4>
      <div class="val">€{bess_capex:,.0f}</div>
      <div class="sub">€{bess_capex_per_kwh}/kWh × {nominal_cap} kWh</div>
    </div>""", unsafe_allow_html=True)
    fc2.markdown(f"""<div class="fin-card">
      <h4>Simple Payback</h4>
      <div class="val {pb_cls}">{pb_txt}</div>
      <div class="sub">vs. grid upgrade alternative</div>
    </div>""", unsafe_allow_html=True)
    fc3.markdown(f"""<div class="fin-card">
      <h4>IRR</h4>
      <div class="val {irr_cls}">{irr:.1f}%</div>
      <div class="sub">Hurdle rate {discount_rate*100:.0f}%</div>
    </div>""", unsafe_allow_html=True)
    fc4.markdown(f"""<div class="fin-card">
      <h4>{analysis_years}-Year NPV</h4>
      <div class="val {npv_cls}">€{total_npv:,.0f}</div>
      <div class="sub">@ {discount_rate*100:.0f}% discount rate</div>
    </div>""", unsafe_allow_html=True)

    # ── Two charts side-by-side ──────────────
    st.markdown('<p class="section-header">Investment Analysis</p>', unsafe_allow_html=True)
    left, right = st.columns(2)
    years_list  = list(range(1, analysis_years + 1))

    with left:
        # NPV Waterfall
        total_demand_sav  = int(fin_df["Demand Saving (€)"].sum())
        total_arb         = int(fin_df["Arbitrage Saving (€)"].sum())
        total_avoided_upg = int(fin_df["Avoided Upgrade (€)"].sum())
        total_opex_cost   = int(fin_df["OPEX (€)"].sum())

        wf_fig = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute","relative","relative","relative","relative","total"],
            x=["CAPEX", "Demand\nSavings", "Arbitrage", "Avoided\nUpgrade", "OPEX\nCost", "Total NPV"],
            y=[-bess_capex, total_demand_sav, total_arb, total_avoided_upg, -total_opex_cost, None],
            connector=dict(line=dict(color="#ddd")),
            decreasing=dict(marker_color="#c62828"),
            increasing=dict(marker_color="#2e7d32"),
            totals=dict(marker_color=KEMPOWER_ORANGE),
            text=[f"€{abs(v):,.0f}" for v in
                  [-bess_capex, total_demand_sav, total_arb, total_avoided_upg, -total_opex_cost]]
                 + [f"€{total_npv:,.0f}"],
            textposition="outside",
            textfont=dict(size=9),
        ))
        wf_fig.update_layout(
            height=360, margin=dict(l=10, r=10, t=30, b=10),
            title=dict(text="NPV Waterfall (undiscounted)", font=dict(size=12)),
            yaxis=dict(tickformat=",.0f"),
            plot_bgcolor='white', paper_bgcolor='white', showlegend=False
        )
        st.plotly_chart(wf_fig, use_container_width=True)

    with right:
        # Cumulative NPV vs grid upgrade
        cum_npv_vals = fin_df["Cumulative NPV (€)"].tolist()
        cf_fig = go.Figure()
        cf_fig.add_trace(go.Scatter(
            x=years_list, y=cum_npv_vals, name='BESS Cumulative NPV',
            mode='lines+markers',
            line=dict(color=KEMPOWER_ORANGE, width=3),
            marker=dict(size=7),
            fill='tozeroy', fillcolor='rgba(255,100,0,0.08)',
            hovertemplate='Year %{x}: €%{y:,.0f}<extra></extra>'
        ))
        cf_fig.add_trace(go.Scatter(
            x=years_list, y=[-grid_upgrade_cost] * analysis_years,
            name='Grid Upgrade (one-off)', mode='lines',
            line=dict(color='#2b579a', width=2, dash='dash'),
            hovertemplate='Grid cost: -€%{customdata:,.0f}<extra></extra>',
            customdata=[grid_upgrade_cost] * analysis_years
        ))
        cf_fig.add_hline(y=0, line_color="#aaa", line_width=1)
        if payback_yr:
            cf_fig.add_vline(x=payback_yr, line_dash="dot", line_color=KEMPOWER_ORANGE,
                              annotation_text=f"Payback Y{payback_yr}",
                              annotation_font=dict(color=KEMPOWER_ORANGE, size=11))
        cf_fig.update_layout(
            height=360, margin=dict(l=10, r=10, t=30, b=10),
            title=dict(text="Cumulative NPV vs Grid Upgrade", font=dict(size=12)),
            xaxis=dict(title="Year"),
            yaxis=dict(tickformat=",.0f"),
            legend=dict(orientation='h', y=-0.22),
            plot_bgcolor='white', paper_bgcolor='white'
        )
        st.plotly_chart(cf_fig, use_container_width=True)

    # ── Annual cash flow bar chart ───────────
    st.markdown('<p class="section-header">Annual Cash Flow Breakdown</p>',
                unsafe_allow_html=True)
    bar_fig = go.Figure()
    bar_fig.add_trace(go.Bar(name='Demand Savings', x=years_list,
                              y=fin_df["Demand Saving (€)"], marker_color='#2e7d32',
                              hovertemplate='Demand: €%{y:,.0f}<extra></extra>'))
    bar_fig.add_trace(go.Bar(name='Avoided Grid Upgrade', x=years_list,
                              y=fin_df["Avoided Upgrade (€)"], marker_color='#1565c0',
                              hovertemplate='Upgrade: €%{y:,.0f}<extra></extra>'))
    bar_fig.add_trace(go.Bar(name='Energy Arbitrage', x=years_list,
                              y=fin_df["Arbitrage Saving (€)"], marker_color='#7eb26d',
                              hovertemplate='Arbitrage: €%{y:,.0f}<extra></extra>'))
    bar_fig.add_trace(go.Bar(name='OPEX', x=years_list,
                              y=-fin_df["OPEX (€)"], marker_color='#c62828',
                              hovertemplate='OPEX: -€%{y:,.0f}<extra></extra>'))
    bar_fig.add_trace(go.Scatter(name='Net Cash Flow', x=years_list,
                                  y=fin_df["Net Cash Flow (€)"],
                                  mode='lines+markers',
                                  line=dict(color=KEMPOWER_ORANGE, width=2.5),
                                  marker=dict(size=7),
                                  hovertemplate='Net: €%{y:,.0f}<extra></extra>'))
    bar_fig.update_layout(
        barmode='relative', height=310,
        margin=dict(l=10, r=10, t=10, b=50),
        legend=dict(orientation='h', y=-0.3),
        xaxis=dict(title="Year", tickvals=years_list),
        yaxis=dict(title="€ / year", tickformat=",.0f"),
        plot_bgcolor='white', paper_bgcolor='white'
    )
    st.plotly_chart(bar_fig, use_container_width=True)

    # ── Sensitivity heat-map ─────────────────
    st.markdown('<p class="section-header">Sensitivity — CAPEX (€/kWh) vs Demand Charge (€/kW/mo)</p>',
                unsafe_allow_html=True)
    st.caption("📍 Payback period in years  ·  🟢 ≤5y  🟡 ≤7y  🟠 ≤10y  🔴 >10y  ⬜ outside horizon")

    capex_range  = [300, 350, 400, 450, 500, 550, 600]
    demand_range = [6, 8, 10, 12, 15, 18, 20]
    sens_matrix  = []
    for cap in capex_range:
        row = []
        for dem in demand_range:
            _, _, pb, _, _ = compute_financials(
                yearly_results, nominal_cap, cap, bess_opex_pct,
                grid_upgrade_cost, dem, energy_price_eur,
                discount_rate, analysis_years, grid_limit
            )
            row.append(pb if pb else 99)
        sens_matrix.append(row)

    sens_df = pd.DataFrame(sens_matrix,
                            index=[f"€{c}/kWh" for c in capex_range],
                            columns=[f"€{d}/kW/mo" for d in demand_range])

    def fmt_pb(val): return "—" if val == 99 else f"{val}y"
    def bg_pb(val):
        if val <= 5:   return 'background-color:#c8e6c9; color:#1b5e20; font-weight:600'
        elif val <= 7: return 'background-color:#fff9c4; color:#f57f17; font-weight:600'
        elif val <= 10:return 'background-color:#ffe0b2; color:#e65100; font-weight:600'
        elif val < 99: return 'background-color:#ffcdd2; color:#b71c1c; font-weight:600'
        return               'background-color:#f5f5f5; color:#aaa'

    st.dataframe(sens_df.style.applymap(bg_pb).format(fmt_pb), use_container_width=True)

    # ── Detail table (collapsible) ───────────
    with st.expander("📄 Full Annual Cash Flow Detail"):
        euro_cols = [c for c in fin_df.columns if "€" in c]
        st.dataframe(
            fin_df.style.format({c: "€{:,.0f}" for c in euro_cols}),
            use_container_width=True
        )
        st.caption(
            "⚠️ Indicative model only. Value drivers: demand charge savings, energy arbitrage, "
            "and amortised avoided grid upgrade. Actual results depend on local tariff structure, "
            "grid operator rules, and site-specific factors. Use as a conversation starter with clients."
        )
