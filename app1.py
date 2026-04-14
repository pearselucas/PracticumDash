"""
Stadium District Real Estate Investment Analyzer  v2.0
Run: streamlit run app.py
Requires: pip install streamlit plotly pandas openpyxl
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np

st.set_page_config(
    page_title="Stadium RE Investment Analyzer",
    page_icon="🏟️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.main-header {
    background: linear-gradient(135deg, #0a2342 0%, #12355b 60%, #1a4a7a 100%);
    color: white; padding: 2rem 2.5rem; border-radius: 14px;
    margin-bottom: 1.5rem; box-shadow: 0 4px 24px rgba(10,35,66,0.18);
}
.main-header h1 { font-size: 2rem; margin: 0; font-weight: 800; letter-spacing:-0.02em; }
.main-header p  { font-size: 0.95rem; margin: 0.4rem 0 0; opacity: 0.8; }
.main-header .subtitle-row { display:flex; gap:1.5rem; margin-top:0.7rem; flex-wrap:wrap; }
.main-header .pill { background:rgba(255,255,255,0.12); border-radius:99px;
                     padding:0.2rem 0.75rem; font-size:0.8rem; font-weight:600; }

.metric-card {
    background:#fff; border:1px solid #e8edf3; border-radius:12px;
    padding:1rem 1.2rem; text-align:center; height:100%;
    box-shadow:0 2px 8px rgba(0,0,0,0.04); transition:box-shadow 0.2s;
}
.metric-card:hover { box-shadow:0 4px 16px rgba(0,0,0,0.09); }
.metric-card .label { font-size:0.72rem; color:#8896a8; text-transform:uppercase;
                      letter-spacing:0.07em; margin-bottom:0.25rem; font-weight:600; }
.metric-card .value { font-size:1.45rem; font-weight:800; color:#0f172a; }
.metric-card .sub   { font-size:0.78rem; color:#a0aec0; margin-top:0.15rem; }

.score-badge-BUY   { background:#dcfce7; color:#15803d; border:1.5px solid #86efac; }
.score-badge-HOLD  { background:#fef9c3; color:#92400e; border:1.5px solid #fde047; }
.score-badge-AVOID { background:#fee2e2; color:#b91c1c; border:1.5px solid #fca5a5; }
.score-badge {
    display:inline-block; padding:0.2rem 0.75rem; border-radius:999px;
    font-weight:700; font-size:0.8rem; letter-spacing:0.04em;
}

.insight-box {
    background:#eff6ff; border-left:4px solid #3b82f6;
    padding:0.75rem 1rem; border-radius:0 8px 8px 0;
    margin:0.35rem 0; font-size:0.88rem; color:#1e3a5f; line-height:1.5;
}
.warning-box {
    background:#fff7ed; border-left:4px solid #f97316;
    padding:0.75rem 1rem; border-radius:0 8px 8px 0;
    margin:0.35rem 0; font-size:0.88rem; color:#7c2d12; line-height:1.5;
}
.section-title {
    font-size:1.05rem; font-weight:700; color:#0f172a;
    border-bottom:2px solid #e8edf3; padding-bottom:0.4rem;
    margin:1.2rem 0 0.8rem;
}
.score-explainer {
    background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px;
    padding:1rem 1.2rem; font-size:0.85rem; color:#475569; margin:0.5rem 0;
}
.score-explainer strong { color:#0f172a; }
div[data-testid="stTabs"] button { font-weight:600; font-size:0.9rem; }
</style>
""", unsafe_allow_html=True)

# ── Stadium Data ─────────────────────────────────────────────────────────────
STADIUMS = {
    "Truist Park": {
        "city": "Atlanta, GA", "team": "Atlanta Braves", "sport": "MLB",
        "league_revenue_rank": 6, "lat": 33.8908, "lon": -84.4677,
        "office":   {"inventory_sf": 9_300_000, "under_construction_sf": 0,
                     "net_absorption_12mo": 534_000, "sale_price_sf": 222,
                     "asking_rent_sf": 34, "cap_rate": 0.084,
                     "vacancy": 0.18, "vacancy_10yr": 0.1977,
                     "rent_growth": 0.026, "rent_growth_10yr": 0.039, "cap_rate_10yr": 0.0752},
        "retail":   {"inventory_sf": 3_800_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -23_300, "sale_price_sf": 249,
                     "asking_rent_sf": 37, "cap_rate": 0.071,
                     "vacancy": 0.0205, "vacancy_10yr": 0.0291,
                     "rent_growth": 0.065, "rent_growth_10yr": 0.048, "cap_rate_10yr": 0.0719},
        "hospitality": {"adr": 154, "revpar": 96, "sale_price_room": 126_000,
                        "cap_rate": 0.086, "occupancy": 0.6245,
                        "occupancy_10yr": 0.6381, "cap_rate_10yr": 0.0828},
        "multifamily": {"inventory_units": 8_804, "under_construction_units": 640,
                        "net_absorption_units": 423, "vacancy": 0.08,
                        "market_rent_unit": 1_762, "sale_price_unit": 256_000,
                        "cap_rate": 0.055, "occupancy": 0.9249, "occupancy_10yr": 0.914,
                        "rent_growth": 0.003, "rent_growth_10yr": 0.032, "cap_rate_10yr": 0.0514},
        "demographics": {"pop_2020": 14_169, "pop_2025": 15_850,
                         "pop_growth_5yr": 0.024, "median_income": 91_802},
    },
    "Mercedes-Benz Stadium": {
        "city": "Atlanta, GA", "team": "Atlanta Falcons / Atlanta United",
        "sport": "NFL/MLS", "league_revenue_rank": 3,
        "lat": 33.7554, "lon": -84.4008,
        "office":   {"inventory_sf": 28_000_000, "under_construction_sf": 0,
                     "net_absorption_12mo": 112_000, "sale_price_sf": 183,
                     "asking_rent_sf": 29, "cap_rate": 0.086,
                     "vacancy": 0.21, "vacancy_10yr": 0.1432,
                     "rent_growth": 0.018, "rent_growth_10yr": 0.041, "cap_rate_10yr": 0.0755},
        "retail":   {"inventory_sf": 2_200_000, "under_construction_sf": 70_000,
                     "net_absorption_12mo": -27_100, "sale_price_sf": 234,
                     "asking_rent_sf": 30, "cap_rate": 0.071,
                     "vacancy": 0.14, "vacancy_10yr": 0.0995,
                     "rent_growth": 0.054, "rent_growth_10yr": 0.049, "cap_rate_10yr": 0.0728},
        "hospitality": {"adr": 206, "revpar": 128, "sale_price_room": 256_000,
                        "cap_rate": 0.079, "occupancy": 0.669,
                        "occupancy_10yr": 0.6195, "cap_rate_10yr": 0.0761},
        "multifamily": {"inventory_units": 8_654, "under_construction_units": 889,
                        "net_absorption_units": 355, "vacancy": 0.18,
                        "market_rent_unit": 1_681, "sale_price_unit": 325_000,
                        "cap_rate": 0.057, "occupancy": 0.8243, "occupancy_10yr": 0.8778,
                        "rent_growth": -0.002, "rent_growth_10yr": 0.02, "cap_rate_10yr": 0.053},
        "demographics": {"pop_2020": 20_811, "pop_2025": 22_444,
                         "pop_growth_5yr": 0.016, "median_income": 59_186},
    },
    "Bank of America Stadium": {
        "city": "Charlotte, NC", "team": "Carolina Panthers",
        "sport": "NFL", "league_revenue_rank": 2,
        "lat": 35.2258, "lon": -80.8528,
        "office":   {"inventory_sf": 30_500_000, "under_construction_sf": 525_000,
                     "net_absorption_12mo": 151_000, "sale_price_sf": 422,
                     "asking_rent_sf": 44.74, "cap_rate": 0.067,
                     "vacancy": 0.2346, "vacancy_10yr": 0.128,
                     "rent_growth": 0.04, "rent_growth_10yr": 0.05, "cap_rate_10yr": 0.0652,
                     "city_avg_vacancy": 0.1782, "city_avg_rent_sf": 37.89, "city_avg_sale_sf": 300},
        "retail":   {"inventory_sf": 1_300_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -29_700, "sale_price_sf": 376,
                     "asking_rent_sf": 36.85, "cap_rate": 0.062,
                     "vacancy": 0.1649, "vacancy_10yr": 0.0714,
                     "rent_growth": 0.051, "rent_growth_10yr": 0.045, "cap_rate_10yr": 0.0644,
                     "city_avg_vacancy": 0.0375, "city_avg_rent_sf": 31.81},
        "hospitality": {"adr": 220.73, "revpar": 144.28, "sale_price_room": 237_000,
                        "cap_rate": 0.088, "occupancy": 0.6536,
                        "occupancy_10yr": 0.6154, "cap_rate_10yr": 0.0813,
                        "city_avg_adr": 140.48, "city_avg_sale_room": 138_000},
        "multifamily": {"inventory_units": 19_281, "under_construction_units": 773,
                        "net_absorption_units": 878, "vacancy": 0.15,
                        "market_rent_unit": 2_111, "sale_price_unit": 355_000,
                        "cap_rate": 0.0552, "occupancy": 0.8504, "occupancy_10yr": 0.8906,
                        "rent_growth": -0.025, "rent_growth_10yr": 0.018, "cap_rate_10yr": 0.0504,
                        "city_avg_vacancy": 0.127, "city_avg_rent_unit": 1_619},
        "demographics": {"pop_2020": 22_659, "pop_2025": 27_835,
                         "pop_growth_5yr": 0.046, "median_income": 107_843},
    },
    "Spectrum Center": {
        "city": "Charlotte, NC", "team": "Charlotte Hornets",
        "sport": "NBA", "league_revenue_rank": 4,
        "lat": 35.2251, "lon": -80.8392,
        "office":   {"inventory_sf": 30_500_000, "under_construction_sf": 525_000,
                     "net_absorption_12mo": 387_000, "sale_price_sf": 415,
                     "asking_rent_sf": 44.08, "cap_rate": 0.067,
                     "vacancy": 0.2248, "vacancy_10yr": 0.1241,
                     "rent_growth": 0.04, "rent_growth_10yr": 0.05, "cap_rate_10yr": 0.0655,
                     "city_avg_vacancy": 0.1782, "city_avg_rent_sf": 37.89},
        "retail":   {"inventory_sf": 1_500_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -28_200, "sale_price_sf": 376,
                     "asking_rent_sf": 36.85, "cap_rate": 0.062,
                     "vacancy": 0.101, "vacancy_10yr": 0.06,
                     "rent_growth": 0.051, "rent_growth_10yr": 0.047, "cap_rate_10yr": 0.0634,
                     "city_avg_vacancy": 0.0375, "city_avg_rent_sf": 31.81},
        "hospitality": {"adr": None, "revpar": None, "sale_price_room": None,
                        "cap_rate": None, "occupancy": None,
                        "occupancy_10yr": None, "cap_rate_10yr": 0.085},
        "multifamily": {"inventory_units": 20_160, "under_construction_units": 1_294,
                        "net_absorption_units": 720, "vacancy": 0.166,
                        "market_rent_unit": 2_028, "sale_price_unit": 348_000,
                        "cap_rate": 0.055, "occupancy": 0.834, "occupancy_10yr": 0.8933,
                        "rent_growth": -0.028, "rent_growth_10yr": 0.017, "cap_rate_10yr": 0.0501,
                        "city_avg_vacancy": 0.127, "city_avg_rent_unit": 1_619},
        "demographics": {"pop_2020": 22_955, "pop_2025": 27_954,
                         "pop_growth_5yr": 0.044, "median_income": 103_429},
    },
    "Allegiant Stadium": {
        "city": "Las Vegas, NV", "team": "Las Vegas Raiders",
        "sport": "NFL", "league_revenue_rank": 2,
        "lat": 36.0909, "lon": -115.1833,
        "office":   {"inventory_sf": 413_000, "under_construction_sf": 0,
                     "net_absorption_12mo": 19_100, "sale_price_sf": 256,
                     "asking_rent_sf": 23.19, "cap_rate": 0.085,
                     "vacancy": 0.0203, "vacancy_10yr": 0.0677,
                     "rent_growth": 0.045, "rent_growth_10yr": 0.042, "cap_rate_10yr": 0.0829},
        "retail":   {"inventory_sf": 894_000, "under_construction_sf": 500_000,
                     "net_absorption_12mo": 0, "sale_price_sf": 557,
                     "asking_rent_sf": 59.37, "cap_rate": 0.06,
                     "vacancy": 0.0036, "vacancy_10yr": 0.0076,
                     "rent_growth": 0.032, "rent_growth_10yr": 0.056, "cap_rate_10yr": 0.0588},
        "hospitality": {"adr": 198, "revpar": 166, "sale_price_room": 115_000,
                        "cap_rate": 0.08, "occupancy": 0.8397,
                        "occupancy_10yr": 0.8149, "cap_rate_10yr": 0.08},
        "multifamily": {"inventory_units": 490, "under_construction_units": 0,
                        "net_absorption_units": -22, "vacancy": 0.231,
                        "market_rent_unit": 1_055, "sale_price_unit": 164_000,
                        "cap_rate": 0.056, "occupancy": 0.904, "occupancy_10yr": 0.7694,
                        "rent_growth": 0.041, "rent_growth_10yr": 0.005, "cap_rate_10yr": 0.0552},
        "demographics": {"pop_2020": None, "pop_2025": None,
                         "pop_growth_5yr": None, "median_income": None},
    },
    "T-Mobile Arena": {
        "city": "Las Vegas, NV", "team": "Vegas Golden Knights",
        "sport": "NHL", "league_revenue_rank": 5,
        "lat": 36.1028, "lon": -115.1784,
        "office":   {"inventory_sf": 414_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -34_300, "sale_price_sf": 228,
                     "asking_rent_sf": 23.69, "cap_rate": 0.085,
                     "vacancy": 0.119, "vacancy_10yr": 0.0845,
                     "rent_growth": 0.059, "rent_growth_10yr": 0.052, "cap_rate_10yr": 0.0825},
        "retail":   {"inventory_sf": 2_500_000, "under_construction_sf": 500_000,
                     "net_absorption_12mo": 5_800, "sale_price_sf": 428,
                     "asking_rent_sf": 70, "cap_rate": 0.063,
                     "vacancy": 0.0997, "vacancy_10yr": 0.0392,
                     "rent_growth": 0.085, "rent_growth_10yr": 0.077, "cap_rate_10yr": 0.0621},
        "hospitality": {"adr": 235, "revpar": 201, "sale_price_room": 115_000,
                        "cap_rate": 0.08, "occupancy": 0.8529,
                        "occupancy_10yr": 0.8414, "cap_rate_10yr": 0.0778},
        "multifamily": {"inventory_units": 4_527, "under_construction_units": 0,
                        "net_absorption_units": 2, "vacancy": 0.228,
                        "market_rent_unit": 1_499, "sale_price_unit": 255_000,
                        "cap_rate": 0.051, "occupancy": 0.7719, "occupancy_10yr": 0.8727,
                        "rent_growth": 0.003, "rent_growth_10yr": 0.028, "cap_rate_10yr": 0.0489},
        "demographics": {"pop_2020": None, "pop_2025": None,
                         "pop_growth_5yr": None, "median_income": None},
    },
    "American Airlines Center": {
        "city": "Dallas, TX", "team": "Dallas Mavericks / Dallas Stars",
        "sport": "NBA/NHL", "league_revenue_rank": 3,
        "lat": 32.7905, "lon": -96.8103,
        "office":   {"inventory_sf": 42_400_000, "under_construction_sf": 1_300_000,
                     "net_absorption_12mo": 421_000, "sale_price_sf": 403,
                     "asking_rent_sf": 42.96, "cap_rate": 0.0634,
                     "vacancy": 0.2602, "vacancy_10yr": 0.2256,
                     "rent_growth": 0.037, "rent_growth_10yr": 0.035, "cap_rate_10yr": 0.0637,
                     "city_avg_vacancy": 0.178, "city_avg_rent_sf": 33},
        "retail":   {"inventory_sf": 1_900_000, "under_construction_sf": 7_700,
                     "net_absorption_12mo": -38_446, "sale_price_sf": 388,
                     "asking_rent_sf": 36.7, "cap_rate": 0.0665,
                     "vacancy": 0.1189, "vacancy_10yr": 0.0676,
                     "rent_growth": 0.033, "rent_growth_10yr": 0.033, "cap_rate_10yr": 0.0651,
                     "city_avg_vacancy": 0.051, "city_avg_rent_sf": 25.32},
        "hospitality": {"adr": 250.69, "revpar": 147.09, "sale_price_room": 253_000,
                        "cap_rate": 0.0793, "occupancy": 0.5867,
                        "occupancy_10yr": 0.5745, "cap_rate_10yr": 0.0767,
                        "city_avg_adr": 128.94, "city_avg_sale_room": 136_000},
        "multifamily": {"inventory_units": 25_430, "under_construction_units": 895,
                        "net_absorption_units": 763, "vacancy": 0.0963,
                        "market_rent_unit": 2_624, "sale_price_unit": 336_000,
                        "cap_rate": 0.0589, "occupancy": 0.9037, "occupancy_10yr": 0.8913,
                        "rent_growth": 0, "rent_growth_10yr": 0.018, "cap_rate_10yr": 0.053,
                        "city_avg_vacancy": 0.121, "city_avg_rent_unit": 1_539},
        "demographics": {"pop_2020": None, "pop_2025": None,
                         "pop_growth_5yr": 0.052, "median_income": 95_700},
    },
    "AT&T Stadium": {
        "city": "Arlington, TX", "team": "Dallas Cowboys",
        "sport": "NFL", "league_revenue_rank": 1,
        "lat": 32.7473, "lon": -97.0945,
        "office":   {"inventory_sf": 1_500_000, "under_construction_sf": 0,
                     "net_absorption_12mo": 27_071, "sale_price_sf": 145,
                     "asking_rent_sf": 24.65, "cap_rate": 0.0934,
                     "vacancy": 0.14, "vacancy_10yr": 0.1307,
                     "rent_growth": 0.016, "rent_growth_10yr": 0.023, "cap_rate_10yr": 0.0829,
                     "city_avg_vacancy": 0.178, "city_avg_rent_sf": 33},
        "retail":   {"inventory_sf": 2_400_000, "under_construction_sf": 0,
                     "net_absorption_12mo": 30_294, "sale_price_sf": 270,
                     "asking_rent_sf": 22.84, "cap_rate": 0.0654,
                     "vacancy": 0.025, "vacancy_10yr": 0.0641,
                     "rent_growth": 0.025, "rent_growth_10yr": 0.041, "cap_rate_10yr": 0.0663,
                     "city_avg_vacancy": 0.051, "city_avg_rent_sf": 25.32},
        "hospitality": {"adr": None, "revpar": None, "sale_price_room": 201_000,
                        "cap_rate": 0.0898, "occupancy": None,
                        "occupancy_10yr": None, "cap_rate_10yr": 0.0827,
                        "city_avg_adr": 128.94, "city_avg_sale_room": 136_000},
        "multifamily": {"inventory_units": 2_940, "under_construction_units": 0,
                        "net_absorption_units": 182, "vacancy": 0.132,
                        "market_rent_unit": 1_336, "sale_price_unit": 132_315,
                        "cap_rate": 0.0575, "occupancy": 0.8684, "occupancy_10yr": 0.9213,
                        "rent_growth": -0.01, "rent_growth_10yr": 0.042, "cap_rate_10yr": 0.0546,
                        "city_avg_vacancy": 0.121, "city_avg_rent_unit": 1_539},
        "demographics": {"pop_2020": None, "pop_2025": None,
                         "pop_growth_5yr": 0.06, "median_income": 52_600},
    },
}

SPORT_COLORS = {
    "NFL": "#013369", "MLB": "#002D62", "NBA": "#C9082A",
    "NHL": "#0e1929", "NFL/MLS": "#013369", "NBA/NHL": "#C9082A",
}

# ── Absolute-Benchmark Scoring ────────────────────────────────────────────────
# Scores are measured against industry-standard RE thresholds, NOT peer-relative.
# This means a Las Vegas tourism market cannot distort scores for a Charlotte CBD.

BENCHMARKS = {
    "office": {
        "vacancy":        {"excellent": 0.08,  "good": 0.12,     "fair": 0.18,    "poor": 0.25},
        "rent_growth":    {"excellent": 0.04,  "good": 0.025,    "fair": 0.01,    "poor": -0.02},
        "net_absorption": {"excellent": 200000,"good": 50000,    "fair": 0,       "poor": -50000},
        "cap_rate":       {"excellent": 0.055, "good": 0.065,    "fair": 0.075,   "poor": 0.10},
        "sale_price_sf":  {"excellent": 350,   "good": 250,      "fair": 150,     "poor": 80},
    },
    "retail": {
        "vacancy":        {"excellent": 0.03,  "good": 0.06,     "fair": 0.10,    "poor": 0.15},
        "rent_growth":    {"excellent": 0.05,  "good": 0.03,     "fair": 0.01,    "poor": -0.02},
        "net_absorption": {"excellent": 20000, "good": 0,        "fair": -15000,  "poor": -40000},
        "cap_rate":       {"excellent": 0.05,  "good": 0.06,     "fair": 0.07,    "poor": 0.09},
        "sale_price_sf":  {"excellent": 400,   "good": 280,      "fair": 200,     "poor": 120},
    },
    "hospitality": {
        "occupancy":       {"excellent": 0.80, "good": 0.70,     "fair": 0.62,    "poor": 0.55},
        "revpar":          {"excellent": 160,  "good": 120,      "fair": 80,      "poor": 50},
        "adr":             {"excellent": 200,  "good": 150,      "fair": 110,     "poor": 75},
        "cap_rate":        {"excellent": 0.065,"good": 0.075,    "fair": 0.085,   "poor": 0.10},
        "sale_price_room": {"excellent": 220000,"good":150000,   "fair": 100000,  "poor": 60000},
    },
    "multifamily": {
        "vacancy":          {"excellent": 0.05, "good": 0.08,   "fair": 0.12,    "poor": 0.18},
        "occupancy":        {"excellent": 0.95, "good": 0.90,   "fair": 0.85,    "poor": 0.80},
        "rent_growth":      {"excellent": 0.04, "good": 0.02,   "fair": 0.00,    "poor": -0.03},
        "rent_growth_10yr": {"excellent": 0.035,"good": 0.025,  "fair": 0.015,   "poor": 0.00},
        "cap_rate":         {"excellent": 0.045,"good": 0.055,  "fair": 0.065,   "poor": 0.08},
        "sale_price_unit":  {"excellent": 300000,"good":200000, "fair": 130000,  "poor": 80000},
        "market_rent_unit": {"excellent": 2500, "good": 1800,   "fair": 1300,    "poor": 900},
    },
    "demographics": {
        "median_income":  {"excellent": 100000,"good": 75000,   "fair": 55000,   "poor": 35000},
        "pop_growth_5yr": {"excellent": 0.05,  "good": 0.03,    "fair": 0.01,    "poor": -0.01},
    },
}

def safe(v):
    if v is None: return None
    if isinstance(v, float) and np.isnan(v): return None
    return v

def abs_score(val, thresholds, higher_is_better=True):
    """Piecewise linear score against absolute benchmarks. Returns 0-100."""
    if safe(val) is None:
        return 50.0
    e, g, f, p = (thresholds["excellent"], thresholds["good"],
                  thresholds["fair"],      thresholds["poor"])
    if higher_is_better:
        if val >= e: return 100.0
        if val >= g: return 75.0 + 25.0 * (val - g) / (e - g)
        if val >= f: return 50.0 + 25.0 * (val - f) / (g - f)
        if val >= p: return 25.0 + 25.0 * (val - p) / (f - p)
        return max(0.0, 25.0 * (val / p)) if p != 0 else 0.0
    else:
        if val <= e: return 100.0
        if val <= g: return 75.0 + 25.0 * (g - val) / (g - e)
        if val <= f: return 50.0 + 25.0 * (f - val) / (f - g)
        if val <= p: return 25.0 + 25.0 * (p - val) / (p - f)
        return max(0.0, 25.0 * (p / val)) if val != 0 else 0.0

def vacancy_trend_penalty(current, avg_10yr):
    """Up to -15 pts if vacancy has materially deteriorated vs 10-yr average."""
    if safe(current) is None or safe(avg_10yr) is None: return 0.0
    return min(15.0, max(0.0, (current - avg_10yr) * 150))

def score_office(d):
    b = BENCHMARKS["office"]
    vs  = abs_score(d.get("vacancy"),             b["vacancy"],        False)
    rgs = abs_score(d.get("rent_growth"),         b["rent_growth"],    True)
    ab  = abs_score(d.get("net_absorption_12mo"), b["net_absorption"], True)
    cs  = abs_score(d.get("cap_rate"),            b["cap_rate"],       False)  # lower cap = more valuable
    ps  = abs_score(d.get("sale_price_sf"),       b["sale_price_sf"],  True)
    pen = vacancy_trend_penalty(d.get("vacancy"), d.get("vacancy_10yr"))
    raw = vs*0.25 + rgs*0.22 + ab*0.18 + cs*0.15 + ps*0.20
    return max(0.0, raw - pen), {
        "Vacancy Rate":       (vs,  f"{d.get('vacancy',0):.1%}",                 "lower = better"),
        "Rent Growth (YoY)":  (rgs, f"{d.get('rent_growth',0):.1%}",             "higher = better"),
        "Net Absorption":     (ab,  f"{d.get('net_absorption_12mo',0):+,.0f} SF","higher = better"),
        "Cap Rate":           (cs,  f"{d.get('cap_rate',0):.2%}",                "lower = more valuable asset"),
        "Sale Price / SF":    (ps,  f"${d.get('sale_price_sf',0):,.0f}",         "higher = better"),
    }, pen

def score_retail(d):
    b = BENCHMARKS["retail"]
    vs  = abs_score(d.get("vacancy"),             b["vacancy"],        False)
    rgs = abs_score(d.get("rent_growth"),         b["rent_growth"],    True)
    ab  = abs_score(d.get("net_absorption_12mo"), b["net_absorption"], True)
    cs  = abs_score(d.get("cap_rate"),            b["cap_rate"],       False)
    ps  = abs_score(d.get("sale_price_sf"),       b["sale_price_sf"],  True)
    pen = vacancy_trend_penalty(d.get("vacancy"), d.get("vacancy_10yr"))
    raw = vs*0.30 + rgs*0.25 + ab*0.15 + cs*0.10 + ps*0.20
    return max(0.0, raw - pen), {
        "Vacancy Rate":      (vs,  f"{d.get('vacancy',0):.1%}",                 "lower = better"),
        "Rent Growth (YoY)": (rgs, f"{d.get('rent_growth',0):.1%}",             "higher = better"),
        "Net Absorption":    (ab,  f"{d.get('net_absorption_12mo',0):+,.0f} SF","higher = better"),
        "Cap Rate":          (cs,  f"{d.get('cap_rate',0):.2%}",                "lower = more valuable"),
        "Sale Price / SF":   (ps,  f"${d.get('sale_price_sf',0):,.0f}",         "higher = better"),
    }, pen

def score_hospitality(d):
    b = BENCHMARKS["hospitality"]
    suppressed = all(safe(d.get(k)) is None for k in ["occupancy", "revpar", "adr"])
    if suppressed:
        cs = abs_score(d.get("cap_rate"),        b["cap_rate"],        False)
        ps = abs_score(d.get("sale_price_room"),  b["sale_price_room"], True)
        score = (cs * 0.5 + ps * 0.5) if (safe(d.get("cap_rate")) or safe(d.get("sale_price_room"))) else 50.0
        return score, {
            "Occupancy":       (50, "Suppressed by CoStar", "single brand >50% of rooms"),
            "RevPAR":          (50, "Suppressed by CoStar", ""),
            "ADR":             (50, "Suppressed by CoStar", ""),
            "Cap Rate":        (cs, f"{d['cap_rate']:.2%}" if safe(d.get("cap_rate")) else "N/A", "lower = more valuable"),
            "Sale Price/Room": (ps, f"${d['sale_price_room']:,}" if safe(d.get("sale_price_room")) else "N/A", "higher = better"),
        }, 0.0
    occ = abs_score(d.get("occupancy"),       b["occupancy"],       True)
    rev = abs_score(d.get("revpar"),          b["revpar"],          True)
    adr = abs_score(d.get("adr"),             b["adr"],             True)
    cs  = abs_score(d.get("cap_rate"),        b["cap_rate"],        False)
    ps  = abs_score(d.get("sale_price_room"), b["sale_price_room"], True)
    raw = occ*0.25 + rev*0.25 + adr*0.20 + cs*0.15 + ps*0.15
    return raw, {
        "Occupancy":       (occ, f"{d.get('occupancy',0):.1%}", "higher = better"),
        "RevPAR":          (rev, f"${d.get('revpar',0):.0f}",   "higher = better"),
        "ADR":             (adr, f"${d.get('adr',0):.0f}",      "higher = better"),
        "Cap Rate":        (cs,  f"{d.get('cap_rate',0):.2%}",  "lower = more valuable"),
        "Sale Price/Room": (ps,  f"${d.get('sale_price_room',0):,.0f}", "higher = better"),
    }, 0.0

def score_multifamily(d):
    b = BENCHMARKS["multifamily"]
    vs   = abs_score(d.get("vacancy"),           b["vacancy"],          False)
    occ  = abs_score(d.get("occupancy"),         b["occupancy"],        True)
    rg_c = abs_score(d.get("rent_growth"),       b["rent_growth"],      True)
    rg_l = abs_score(d.get("rent_growth_10yr"),  b["rent_growth_10yr"], True)
    blended = rg_c * 0.45 + rg_l * 0.55  # structural 10-yr trend anchors short-term dip
    cs   = abs_score(d.get("cap_rate"),          b["cap_rate"],         False)
    ps   = abs_score(d.get("sale_price_unit"),   b["sale_price_unit"],  True)
    mkt  = abs_score(d.get("market_rent_unit"),  b["market_rent_unit"], True)
    pen  = vacancy_trend_penalty(d.get("vacancy"), None)  # MF doesn't always have vacancy_10yr
    raw  = vs*0.20 + occ*0.15 + blended*0.25 + cs*0.15 + ps*0.15 + mkt*0.10
    rg_label = (f"{d.get('rent_growth',0):.1%} now / "
                f"{d.get('rent_growth_10yr',0):.1%} 10-yr")
    return max(0.0, raw - pen), {
        "Vacancy Rate":       (vs,      f"{d.get('vacancy',0):.1%}",           "lower = better"),
        "Occupancy Rate":     (occ,     f"{d.get('occupancy',0):.1%}",         "higher = better"),
        "Rent Growth":        (blended, rg_label,                               "blended current + 10-yr"),
        "Cap Rate":           (cs,      f"{d.get('cap_rate',0):.2%}",          "lower = more valuable"),
        "Sale Price / Unit":  (ps,      f"${d.get('sale_price_unit',0):,.0f}", "higher = better"),
        "Market Rent / Unit": (mkt,     f"${d.get('market_rent_unit',0):,.0f}","higher = better"),
    }, pen

def demo_bonus(d):
    """Demographic quality bonus: up to +8 pts on overall score."""
    b = BENCHMARKS["demographics"]
    inc  = abs_score(d.get("median_income"),  b["median_income"],  True)
    popg = abs_score(d.get("pop_growth_5yr"), b["pop_growth_5yr"], True)
    return (inc * 0.5 + popg * 0.5) * 0.08

def rating(score):
    if score >= 65: return "BUY"
    if score >= 50: return "HOLD"
    return "AVOID"

def color_score(score):
    if score >= 65: return "#16a34a"
    if score >= 50: return "#d97706"
    return "#dc2626"

def badge_html(score):
    r = rating(score)
    return f'<span class="score-badge score-badge-{r}">{r}</span>'

# ── Build Scores ──────────────────────────────────────────────────────────────
@st.cache_data
def build_scores():
    rows = []
    for name, info in STADIUMS.items():
        o_sc, o_comps, o_pen = score_office(info["office"])
        r_sc, r_comps, r_pen = score_retail(info["retail"])
        h_sc, h_comps, h_pen = score_hospitality(info["hospitality"])
        m_sc, m_comps, m_pen = score_multifamily(info["multifamily"])
        db      = demo_bonus(info["demographics"])
        overall = min(100.0, np.mean([o_sc, r_sc, h_sc, m_sc]) + db)
        rows.append({
            "stadium": name, "city": info["city"],
            "team": info["team"], "sport": info["sport"],
            "lat": info["lat"],   "lon": info["lon"],
            "overall": overall,   "office": o_sc,
            "retail": r_sc,       "hospitality": h_sc,
            "multifamily": m_sc,  "demo_bonus": db,
            "office_comps": o_comps, "retail_comps": r_comps,
            "hosp_comps": h_comps,   "mf_comps": m_comps,
            "office_pen": o_pen,     "retail_pen": r_pen, "mf_pen": m_pen,
            "pop_growth":    info["demographics"].get("pop_growth_5yr"),
            "median_income": info["demographics"].get("median_income"),
            "league_revenue_rank": info["league_revenue_rank"],
        })
    return pd.DataFrame(rows)

scores_df = build_scores()

# ── Insight Generator ─────────────────────────────────────────────────────────
def generate_insights(name, info, row):
    ins = {"office": [], "retail": [], "hospitality": [], "multifamily": []}
    wrn = {"office": [], "retail": [], "hospitality": [], "multifamily": []}

    o = info["office"]
    if safe(o.get("vacancy")):
        if o["vacancy"] < 0.10:
            ins["office"].append(f"Vacancy of {o['vacancy']:.1%} is below 10% — landlord-favorable with full pricing power.")
        elif o["vacancy"] > 0.22:
            wrn["office"].append(f"Vacancy of {o['vacancy']:.1%} is above 22%. Industry benchmark for concern is 18%.")
    if safe(o.get("vacancy")) and safe(o.get("vacancy_10yr")) and o["vacancy"] > o["vacancy_10yr"] + 0.05:
        wrn["office"].append(f"Vacancy is {o['vacancy']-o['vacancy_10yr']:.1%} above 10-yr average ({o['vacancy_10yr']:.1%}). Likely post-COVID overhang rather than structural deterioration.")
    if safe(o.get("rent_growth")) and o["rent_growth"] >= 0.04:
        ins["office"].append(f"Rent growth of {o['rent_growth']:.1%} meets the institutional ≥4% benchmark.")
    if safe(o.get("net_absorption_12mo")) and o["net_absorption_12mo"] > 100_000:
        ins["office"].append(f"Positive 12-month net absorption of {o['net_absorption_12mo']:,.0f} SF confirms active tenant demand.")
    elif safe(o.get("net_absorption_12mo")) and o["net_absorption_12mo"] < 0:
        wrn["office"].append(f"Negative net absorption of {o['net_absorption_12mo']:,.0f} SF — more space vacated than leased.")
    if safe(o.get("cap_rate")) and o["cap_rate"] < 0.07:
        ins["office"].append(f"Cap rate of {o['cap_rate']:.2%} reflects institutional-grade demand. Lower cap = higher asset quality.")
    elif safe(o.get("cap_rate")) and o["cap_rate"] > 0.09:
        wrn["office"].append(f"Cap rate of {o['cap_rate']:.2%} is elevated — signals higher perceived risk or secondary-market pricing.")
    if safe(o.get("sale_price_sf")) and o["sale_price_sf"] > 350:
        ins["office"].append(f"Sale price of ${o['sale_price_sf']}/SF exceeds the $350 institutional benchmark — premium submarket.")
    if safe(o.get("city_avg_vacancy")) and safe(o.get("vacancy")):
        diff = o["vacancy"] - o["city_avg_vacancy"]
        if diff > 0.04:
            wrn["office"].append(f"Venue vacancy ({o['vacancy']:.1%}) runs {diff:.1%} above city avg ({o['city_avg_vacancy']:.1%}).")
        elif diff < -0.02:
            ins["office"].append(f"Venue vacancy ({o['vacancy']:.1%}) beats city average ({o['city_avg_vacancy']:.1%}) — outperforming the broader market.")

    r = info["retail"]
    if safe(r.get("vacancy")):
        if r["vacancy"] < 0.04:
            ins["retail"].append(f"Near-zero vacancy ({r['vacancy']:.2%}) — extremely tight supply. Landlord commands full pricing power.")
        elif r["vacancy"] > 0.12:
            wrn["retail"].append(f"Retail vacancy of {r['vacancy']:.1%} exceeds the 10-12% caution threshold.")
    if safe(r.get("rent_growth")) and r["rent_growth"] >= 0.05:
        ins["retail"].append(f"Rent growth of {r['rent_growth']:.1%} meets the ≥5% excellence benchmark.")
    if safe(r.get("net_absorption_12mo")) and r["net_absorption_12mo"] < -20_000:
        wrn["retail"].append(f"Persistent negative absorption ({r['net_absorption_12mo']:,.0f} SF) — tenants are exiting the area.")
    elif safe(r.get("net_absorption_12mo")) and r["net_absorption_12mo"] > 0:
        ins["retail"].append(f"Positive net absorption of {r['net_absorption_12mo']:,.0f} SF — occupier demand is growing.")
    if safe(r.get("sale_price_sf")) and r["sale_price_sf"] > 350:
        ins["retail"].append(f"Sale price of ${r['sale_price_sf']}/SF well above the $280 benchmark — premium location asset.")

    h = info["hospitality"]
    if not safe(h.get("occupancy")):
        wrn["hospitality"].append("CoStar suppressed hotel performance data (OCC/ADR/RevPAR) — one brand exceeds 50% of rooms in radius. Score based on cap rate and sale price only.")
    else:
        if h["occupancy"] > 0.80:
            ins["hospitality"].append(f"Hotel occupancy of {h['occupancy']:.1%} exceeds the 80% excellence threshold — exceptional demand.")
        elif h["occupancy"] < 0.62:
            wrn["hospitality"].append(f"Occupancy of {h['occupancy']:.1%} falls below the 62% fair-market threshold.")
        if safe(h.get("revpar")) and h["revpar"] > 160:
            ins["hospitality"].append(f"RevPAR of ${h['revpar']:.0f} exceeds the $160 excellence benchmark.")
        if safe(h.get("adr")) and h["adr"] > 200:
            ins["hospitality"].append(f"ADR of ${h['adr']:.0f} confirms premium pricing well above the $150 benchmark.")
        if safe(h.get("city_avg_adr")) and safe(h.get("adr")):
            premium = h["adr"] - h["city_avg_adr"]
            if premium > 30:
                ins["hospitality"].append(f"ADR is ${premium:.0f} above city avg (${h['city_avg_adr']:.0f}) — venue drives a clear demand premium.")

    m = info["multifamily"]
    if safe(m.get("vacancy")):
        if m["vacancy"] < 0.06:
            ins["multifamily"].append(f"Very tight vacancy ({m['vacancy']:.1%}) — near-zero availability supports rent increases.")
        elif m["vacancy"] > 0.15:
            wrn["multifamily"].append(f"Vacancy of {m['vacancy']:.1%} exceeds 15% — elevated supply pressure.")
    if safe(m.get("rent_growth")) and m["rent_growth"] < 0:
        if safe(m.get("rent_growth_10yr")) and m["rent_growth_10yr"] > 0.015:
            ins["multifamily"].append(f"Current rent growth is negative ({m['rent_growth']:.1%}) but the 10-yr structural trend ({m['rent_growth_10yr']:.1%}) remains positive — near-term correction, not structural decline.")
        else:
            wrn["multifamily"].append(f"Rent growth of {m['rent_growth']:.1%} is negative with a weak long-term trend — sustained pressure on NOI.")
    if safe(m.get("market_rent_unit")) and m["market_rent_unit"] > 2_000:
        ins["multifamily"].append(f"Market rent of ${m['market_rent_unit']:,}/unit signals an affluent renter base and strong demand quality.")
    if safe(m.get("sale_price_unit")) and m["sale_price_unit"] > 280_000:
        ins["multifamily"].append(f"Sale price of ${m['sale_price_unit']:,}/unit reflects institutional-quality multifamily assets.")
    if safe(m.get("cap_rate")) and m["cap_rate"] < 0.055:
        ins["multifamily"].append(f"Cap rate of {m['cap_rate']:.2%} below 5.5% — institutional capital is compressing yields, signaling quality.")
    if safe(m.get("under_construction_units")) and safe(m.get("inventory_units")) and m["inventory_units"] > 0:
        pp = m["under_construction_units"] / m["inventory_units"]
        if pp > 0.08:
            wrn["multifamily"].append(f"Pipeline of {m['under_construction_units']:,} units ({pp:.1%} of inventory) represents meaningful new supply risk.")

    demo = info["demographics"]
    if safe(demo.get("pop_growth_5yr")) and demo["pop_growth_5yr"] > 0.035:
        for k in ins:
            ins[k].append(f"Strong 5-yr population growth ({demo['pop_growth_5yr']:.1%}) underpins demand across all asset classes.")
    if safe(demo.get("median_income")) and demo["median_income"] > 90_000:
        for k in ins:
            ins[k].append(f"Median income of ${demo['median_income']:,} supports premium rents, ADRs, and consumer spending in the submarket.")
    return ins, wrn

# ── Chart Helpers ─────────────────────────────────────────────────────────────
def quadrant_chart(df, x_col, y_col, x_label, y_label, title):
    fig = go.Figure()
    fig.add_shape(type="rect", x0=65, x1=105, y0=65, y1=105,
                  fillcolor="rgba(22,163,74,0.07)", line_width=0)
    for thresh, color in [(65, "#86efac"), (50, "#fde047")]:
        fig.add_vline(x=thresh, line_dash="dash", line_color=color, line_width=1.5)
        fig.add_hline(y=thresh, line_dash="dash", line_color=color, line_width=1.5)
    for _, row in df.iterrows():
        xv = df.loc[row.name, x_col]
        yv = df.loc[row.name, y_col]
        clr = SPORT_COLORS.get(row["sport"], "#475569")
        label = (row["stadium"].replace(" Stadium","").replace(" Arena","")
                               .replace(" Center","").strip())
        fig.add_trace(go.Scatter(
            x=[xv], y=[yv], mode="markers+text",
            marker=dict(size=20, color=clr, line=dict(width=2.5, color="white")),
            text=[label], textposition="top center",
            textfont=dict(size=10, color="#0f172a", family="Inter"),
            name=row["stadium"], showlegend=False,
            hovertemplate=(f"<b>{row['stadium']}</b><br>{row['team']}<br>"
                           f"{x_label}: %{{x:.1f}}<br>{y_label}: %{{y:.1f}}<extra></extra>"),
        ))
    fig.add_annotation(x=104, y=104, text="▲ Best Zone", showarrow=False,
                       font=dict(color="#16a34a", size=11, family="Inter"), xanchor="right")
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color="#0f172a", family="Inter")),
        xaxis_title=x_label, yaxis_title=y_label,
        plot_bgcolor="#f8fafc", paper_bgcolor="white",
        height=500, margin=dict(l=60, r=30, t=55, b=55),
        xaxis=dict(showgrid=True, gridcolor="#e2e8f0", range=[18, 107]),
        yaxis=dict(showgrid=True, gridcolor="#e2e8f0", range=[18, 107]),
        font=dict(family="Inter"),
    )
    return fig

def component_bar(comps, penalty=0.0):
    labels = list(comps.keys())
    scores = [v[0] for v in comps.values()]
    vals   = [v[1] for v in comps.values()]
    notes  = [v[2] for v in comps.values()]
    fig = go.Figure(go.Bar(
        x=scores, y=labels, orientation="h",
        marker_color=[color_score(s) for s in scores],
        text=[f"{s:.0f}  ({v})" for s, v in zip(scores, vals)],
        textposition="outside", textfont=dict(size=11, family="Inter"),
        hovertemplate="<b>%{y}</b><br>Score: %{x:.1f}<br>%{customdata}<extra></extra>",
        customdata=notes,
    ))
    fig.add_vline(x=65, line_dash="dash", line_color="#16a34a", line_width=1.5,
                  annotation_text="BUY", annotation_position="top right",
                  annotation_font=dict(color="#16a34a", size=10))
    fig.add_vline(x=50, line_dash="dash", line_color="#d97706", line_width=1.5,
                  annotation_text="HOLD", annotation_position="bottom right",
                  annotation_font=dict(color="#d97706", size=10))
    if penalty > 0:
        fig.add_annotation(x=2, y=-0.9, xref="x", yref="y",
                           text=f"⚠ Trend penalty: -{penalty:.1f} pts (vacancy deteriorating vs 10-yr avg)",
                           showarrow=False, font=dict(size=10, color="#92400e"), xanchor="left")
    fig.update_layout(
        xaxis=dict(range=[0, 128], showgrid=True, gridcolor="#e2e8f0"),
        plot_bgcolor="#f8fafc", paper_bgcolor="white",
        height=max(240, len(labels) * 55),
        margin=dict(l=130, r=20, t=10, b=45 if penalty > 0 else 10),
        font=dict(family="Inter"),
    )
    return fig

# ════════════════════════════════════════════════════════════════════════════════
#  LAYOUT
# ════════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
  <h1>🏟️ Stadium District Real Estate Analyzer</h1>
  <p>Institutional-grade investment scoring for Office, Retail, Hospitality &amp; Multifamily assets within 1-mile of major U.S. sports venues</p>
  <div class="subtitle-row">
    <span class="pill">📊 CoStar Analytics · April 2026</span>
    <span class="pill">🏗️ Absolute benchmarks — not peer-relative</span>
    <span class="pill">🔬 8 venues · 4 asset classes · 10-yr trend analysis</span>
  </div>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### 🎛️ Controls")
    selected_sports = st.multiselect(
        "Sport", options=sorted(set(v["sport"] for v in STADIUMS.values())),
        default=sorted(set(v["sport"] for v in STADIUMS.values())),
    )
    asset_focus = st.selectbox("Asset Class Focus",
        ["Overall", "Office", "Retail", "Hospitality", "Multifamily"])
    score_col = {"Overall":"overall","Office":"office","Retail":"retail",
                 "Hospitality":"hospitality","Multifamily":"multifamily"}[asset_focus]
    min_score = st.slider("Minimum Score", 0, 100, 0)

    st.markdown("---")
    st.markdown("### 📖 Rating Guide")
    st.markdown("""
    <span class="score-badge score-badge-BUY">BUY</span> **65–100** Strong vs. benchmarks<br><br>
    <span class="score-badge score-badge-HOLD">HOLD</span> **50–64** Mixed; selective underwriting<br><br>
    <span class="score-badge score-badge-AVOID">AVOID</span> **0–49** Weak fundamentals
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Model Methodology")
    st.markdown("""
    -  **Absolute benchmarks** — Las Vegas tourism market no longer penalises Charlotte CBD
    -  **Cap rate direction fixed** — lower cap = more valuable (was inverted in v1)
    -  **MF rent growth blended** — 45% current / 55% 10-yr structural trend to avoid over-penalising short-term corrections
    -  **Vacancy trend penalty** — up to −15 pts if current vacancy has deteriorated vs 10-yr average
    -  **Demographic quality bonus** — ±8 pts for income + population growth
    -  **Suppressed hospitality data** handled gracefully with neutral fills
    """)

filtered = scores_df[
    scores_df["sport"].apply(lambda s: any(sp in s for sp in selected_sports)) &
    (scores_df[score_col] >= min_score)
].copy()

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Market Overview", "🏟️ Stadium Deep Dive",
    "📈 Quadrant Analysis", "📋 Full Data",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Market Overview
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    sorted_df = filtered.sort_values(score_col, ascending=False).reset_index(drop=True)
    st.markdown(f'<div class="section-title">Investment Scores — {asset_focus}</div>', unsafe_allow_html=True)

    fig_bar = go.Figure(go.Bar(
        x=sorted_df["stadium"], y=sorted_df[score_col].round(1),
        marker_color=[color_score(s) for s in sorted_df[score_col]],
        text=[f"{s:.0f}" for s in sorted_df[score_col]],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Score: %{y:.1f}<extra></extra>",
    ))
    fig_bar.add_hline(y=65, line_dash="dash", line_color="#16a34a", line_width=1.5,
                      annotation_text="BUY ≥ 65", annotation_position="right",
                      annotation_font=dict(color="#16a34a"))
    fig_bar.add_hline(y=50, line_dash="dash", line_color="#d97706", line_width=1.5,
                      annotation_text="HOLD ≥ 50", annotation_position="right",
                      annotation_font=dict(color="#d97706"))
    fig_bar.update_layout(
        yaxis_title="Investment Score (0–100)", plot_bgcolor="#f8fafc",
        paper_bgcolor="white", height=380,
        yaxis=dict(range=[0, 115]),
        margin=dict(l=50, r=100, t=20, b=80),
        xaxis=dict(tickangle=-25), font=dict(family="Inter"),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown('<div class="section-title">Multi-Asset Radar — Top 4 Stadiums</div>', unsafe_allow_html=True)
    top4 = sorted_df.head(4)["stadium"].tolist()
    cats = ["Office", "Retail", "Hospitality", "Multifamily", "Office"]
    palette = ["#0ea5e9", "#f59e0b", "#10b981", "#ef4444"]
    fig_radar = go.Figure()
    for i, sname in enumerate(top4):
        r = scores_df[scores_df["stadium"] == sname].iloc[0]
        vals = [r["office"], r["retail"], r["hospitality"], r["multifamily"], r["office"]]
        fig_radar.add_trace(go.Scatterpolar(
            r=vals, theta=cats, fill="toself", name=sname,
            line=dict(color=palette[i], width=2.5), fillcolor=palette[i], opacity=0.12,
        ))
    fig_radar.add_trace(go.Scatterpolar(
        r=[65, 65, 65, 65, 65], theta=cats, mode="lines", name="BUY threshold",
        line=dict(color="#16a34a", dash="dot", width=1.5),
    ))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100],
                                   tickvals=[25, 50, 65, 100],
                                   ticktext=["25", "50", "BUY", "100"])),
        height=440, showlegend=True, legend=dict(orientation="h", y=-0.15),
        paper_bgcolor="white", font=dict(family="Inter"),
    )
    st.plotly_chart(fig_radar, use_container_width=True)

    st.markdown('<div class="section-title">Full Scorecard</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    for idx, row in sorted_df.iterrows():
        sc = row[score_col]
        with cols[idx % 4]:
            st.markdown(f"""
            <div class="metric-card">
              <div class="label">{row['sport']} · {row['city']}</div>
              <div class="value" style="font-size:0.95rem;margin:0.2rem 0">{row['stadium']}</div>
              <div style="font-size:2.1rem;font-weight:800;color:{color_score(sc)};line-height:1;margin:0.25rem 0">{sc:.0f}</div>
              {badge_html(sc)}
              <div class="sub" style="margin-top:0.4rem">{row['team']}</div>
            </div>""", unsafe_allow_html=True)
            st.markdown("")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – Stadium Deep Dive
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    sel = st.selectbox("Select a Stadium", list(STADIUMS.keys()), index=2,
                       help="Bank of America Stadium pre-selected to illustrate the corrected scoring.")
    info = STADIUMS[sel]
    row  = scores_df[scores_df["stadium"] == sel].iloc[0]
    ins, wrn = generate_insights(sel, info, row)

    st.markdown(f"### 🏟️ {sel}")
    st.caption(f"{info['team']} · {info['city']} · CoStar Analytics, 1-Mile Radius, April 2026")

    kpi_cols = st.columns(6)
    for i, (key, label) in enumerate(zip(
        ["overall","office","retail","hospitality","multifamily","demo_bonus"],
        ["Overall","Office","Retail","Hospitality","Multifamily","Demo Bonus"],
    )):
        sc = row[key]
        with kpi_cols[i]:
            if key == "demo_bonus":
                st.markdown(f"""<div class="metric-card">
                  <div class="label">{label}</div>
                  <div class="value" style="color:#6366f1">+{sc:.1f}</div>
                  <div class="sub">max +8.0</div></div>""", unsafe_allow_html=True)
            else:
                st.markdown(f"""<div class="metric-card">
                  <div class="label">{label}</div>
                  <div class="value" style="color:{color_score(sc)}">{sc:.0f}</div>
                  {badge_html(sc)}</div>""", unsafe_allow_html=True)

    st.markdown("")
    demo = info["demographics"]
    st.markdown('<div class="section-title">📍 1-Mile Demographics</div>', unsafe_allow_html=True)
    d1, d2, d3, d4 = st.columns(4)
    pop_delta = (f"+{demo['pop_2025']-demo['pop_2020']:,}"
                 if demo.get("pop_2020") and demo.get("pop_2025") else None)
    d1.metric("2025 Population", f"{demo['pop_2025']:,}" if demo.get("pop_2025") else "N/A", delta=pop_delta)
    d2.metric("5-Yr Pop Growth", f"{demo['pop_growth_5yr']:.1%}" if demo.get("pop_growth_5yr") else "N/A")
    d3.metric("Median Income", f"${demo['median_income']:,}" if demo.get("median_income") else "N/A")
    d4.metric("2020 Population", f"{demo['pop_2020']:,}" if demo.get("pop_2020") else "N/A")

    st.markdown("")
    asset_tabs = st.tabs(["🏢 Office", "🛍️ Retail", "🏨 Hospitality", "🏠 Multifamily"])

    def render_asset(tab, asset_key, label, comps_key, pen_key):
        d   = info[asset_key]
        sc  = row[asset_key]
        pen = row.get(pen_key, 0) or 0.0
        with tab:
            left, right = st.columns([1, 2])
            with left:
                st.markdown(f"""<div class="metric-card" style="text-align:left;padding:1.2rem">
                  <div class="label">Investment Score</div>
                  <div style="font-size:2.8rem;font-weight:800;color:{color_score(sc)};line-height:1">{sc:.0f}</div>
                  <div style="margin:0.4rem 0">{badge_html(sc)}</div>
                  <div class="sub">out of 100</div></div>""", unsafe_allow_html=True)
            with right:
                r = rating(sc)
                if r == "BUY":
                    st.success(f"**BUY** — {label} near {sel} meets or exceeds institutional benchmarks. Score of {sc:.0f} reflects strong underlying fundamentals.")
                elif r == "HOLD":
                    st.warning(f"**HOLD** — {label} near {sel} shows mixed indicators. Score of {sc:.0f} warrants selective underwriting; value-add strategies may outperform.")
                else:
                    st.error(f"**AVOID** — {label} near {sel} scores below institutional thresholds (score: {sc:.0f}). Elevated risk; redirect capital to higher-scoring markets.")

            st.markdown('<div class="section-title">Score Breakdown</div>', unsafe_allow_html=True)
            st.markdown("""<div class="score-explainer">
            Each metric is scored against <strong>absolute industry benchmarks</strong> — not relative to other
            stadiums in this dataset. Green bars ≥ 65 = BUY territory. The dashed lines mark BUY (65) and HOLD (50).
            </div>""", unsafe_allow_html=True)
            st.plotly_chart(component_bar(row[comps_key], pen), use_container_width=True)

            st.markdown('<div class="section-title">Raw Market Data</div>', unsafe_allow_html=True)
            if asset_key == "office":
                c1,c2,c3 = st.columns(3)
                c1.metric("Vacancy", f"{d.get('vacancy',0):.1%}",
                          delta=f"{d['vacancy']-d['vacancy_10yr']:+.1%} vs 10-yr" if d.get("vacancy_10yr") else None,
                          delta_color="inverse")
                c2.metric("Asking Rent/SF", f"${d.get('asking_rent_sf',0):.2f}")
                c3.metric("Rent Growth (YoY)", f"{d.get('rent_growth',0):.1%}")
                c4,c5,c6 = st.columns(3)
                c4.metric("Sale Price/SF", f"${d.get('sale_price_sf',0):,}")
                c5.metric("Cap Rate", f"{d.get('cap_rate',0):.2%}",
                          help="Lower cap rate = higher asset quality / institutional demand")
                c6.metric("12-Mo Net Absorption", f"{d.get('net_absorption_12mo',0):+,.0f} SF")
                if d.get("city_avg_vacancy"):
                    c7,c8,_ = st.columns(3)
                    c7.metric("10-Yr Avg Vacancy", f"{d.get('vacancy_10yr',0):.1%}")
                    c8.metric("City Avg Vacancy", f"{d.get('city_avg_vacancy',0):.1%}")
            elif asset_key == "retail":
                c1,c2,c3 = st.columns(3)
                c1.metric("Vacancy", f"{d.get('vacancy',0):.1%}",
                          delta=f"{d['vacancy']-d['vacancy_10yr']:+.1%} vs 10-yr" if d.get("vacancy_10yr") else None,
                          delta_color="inverse")
                c2.metric("Asking Rent/SF", f"${d.get('asking_rent_sf',0):.2f}")
                c3.metric("Rent Growth (YoY)", f"{d.get('rent_growth',0):.1%}")
                c4,c5,c6 = st.columns(3)
                c4.metric("Sale Price/SF", f"${d.get('sale_price_sf',0):,}")
                c5.metric("Cap Rate", f"{d.get('cap_rate',0):.2%}", help="Lower cap rate = more valuable")
                c6.metric("12-Mo Net Absorption", f"{d.get('net_absorption_12mo',0):+,.0f} SF")
            elif asset_key == "hospitality":
                c1,c2,c3,c4 = st.columns(4)
                c1.metric("Occupancy", f"{d['occupancy']:.1%}" if safe(d.get("occupancy")) else "Suppressed")
                c2.metric("ADR", f"${d['adr']:.0f}" if safe(d.get("adr")) else "Suppressed")
                c3.metric("RevPAR", f"${d['revpar']:.0f}" if safe(d.get("revpar")) else "Suppressed")
                c4.metric("Cap Rate", f"{d['cap_rate']:.2%}" if safe(d.get("cap_rate")) else "N/A")
                if safe(d.get("sale_price_room")):
                    c5,_ = st.columns([1,2])
                    c5.metric("Sale Price/Room", f"${d['sale_price_room']:,}")
                if safe(d.get("city_avg_adr")) and safe(d.get("adr")):
                    c6,_ = st.columns([1,2])
                    c6.metric("City Avg ADR", f"${d['city_avg_adr']:.0f}",
                              delta=f"+${d['adr']-d['city_avg_adr']:.0f} venue premium")
            else:
                c1,c2,c3 = st.columns(3)
                c1.metric("Vacancy", f"{d.get('vacancy',0):.1%}", delta_color="inverse")
                c2.metric("Market Rent/Unit", f"${d.get('market_rent_unit',0):,}")
                c3.metric("Occupancy", f"{d.get('occupancy',0):.1%}")
                c4,c5,c6 = st.columns(3)
                c4.metric("Rent Growth (Current)", f"{d.get('rent_growth',0):.1%}",
                          delta=f"10-yr avg: {d.get('rent_growth_10yr',0):.1%}")
                c5.metric("Cap Rate", f"{d.get('cap_rate',0):.2%}", help="Lower = more valuable")
                c6.metric("Sale Price/Unit", f"${d.get('sale_price_unit',0):,}")
                if safe(d.get("under_construction_units")) and safe(d.get("inventory_units")):
                    pp = d["under_construction_units"] / d["inventory_units"]
                    c7,_ = st.columns([1,2])
                    c7.metric("Pipeline", f"{d['under_construction_units']:,} units",
                              delta=f"{pp:.1%} of inventory", delta_color="inverse")

            if ins[asset_key] or wrn[asset_key]:
                st.markdown('<div class="section-title">Analyst Notes</div>', unsafe_allow_html=True)
            for note in ins[asset_key]:
                st.markdown(f'<div class="insight-box">✅ {note}</div>', unsafe_allow_html=True)
            for note in wrn[asset_key]:
                st.markdown(f'<div class="warning-box">⚠️ {note}</div>', unsafe_allow_html=True)

    render_asset(asset_tabs[0], "office",      "Office",      "office_comps", "office_pen")
    render_asset(asset_tabs[1], "retail",      "Retail",      "retail_comps", "retail_pen")
    render_asset(asset_tabs[2], "hospitality", "Hospitality", "hosp_comps",   "office_pen")
    render_asset(asset_tabs[3], "multifamily", "Multifamily", "mf_comps",     "mf_pen")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Quadrant Analysis
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### 📈 Quadrant Analysis")
    st.info("Green dashed lines mark the **BUY threshold (65)**. Top-right quadrant = strongest investment opportunity across both axes.")

    q1, q2 = st.columns(2)
    with q1:
        st.plotly_chart(quadrant_chart(filtered, "retail", "multifamily",
            "Retail Score", "Multifamily Score", "Retail vs. Multifamily"), use_container_width=True)
    with q2:
        st.plotly_chart(quadrant_chart(filtered, "office", "hospitality",
            "Office Score", "Hospitality Score", "Office vs. Hospitality"), use_container_width=True)

    st.markdown("### Overall Score vs. Population Growth")
    st.caption("Bubble size = median income. Top-right = high-growth, high-income, high-scoring markets.")
    fig_bub = go.Figure()
    for _, row in filtered.iterrows():
        xv  = row["pop_growth"] if safe(row["pop_growth"]) else 0.02
        yv  = row["overall"]
        sz  = max(12, (row["median_income"] / 4000)) if safe(row["median_income"]) else 18
        clr = SPORT_COLORS.get(row["sport"], "#475569")
        fig_bub.add_trace(go.Scatter(
            x=[xv], y=[yv], mode="markers+text",
            marker=dict(size=sz, color=clr, opacity=0.82, line=dict(width=2, color="white")),
            text=[row["stadium"].split()[0]],
            textposition="top center", textfont=dict(size=9, family="Inter"),
            showlegend=False,
            hovertemplate=(f"<b>{row['stadium']}</b><br>Overall: {yv:.1f}<br>"
                           f"Pop Growth: {xv:.1%}<br>"
                           f"Income: ${row['median_income']:,.0f}<extra></extra>"),
        ))
    fig_bub.update_layout(
        xaxis_title="5-Year Population Growth", yaxis_title="Overall Investment Score",
        plot_bgcolor="#f8fafc", paper_bgcolor="white", height=460,
        xaxis=dict(tickformat=".1%", showgrid=True, gridcolor="#e2e8f0"),
        yaxis=dict(showgrid=True, gridcolor="#e2e8f0"),
        font=dict(family="Inter"),
    )
    st.plotly_chart(fig_bub, use_container_width=True)

    st.markdown("### Score Heatmap — All Venues × All Asset Classes")
    heat = filtered[["stadium","office","retail","hospitality","multifamily"]].set_index("stadium")
    fig_heat = px.imshow(
        heat.round(1),
        color_continuous_scale=[[0,"#fef2f2"],[0.48,"#fefce8"],[0.64,"#f0fdf4"],[1,"#14532d"]],
        zmin=0, zmax=100, text_auto=".0f", aspect="auto", labels=dict(color="Score"),
    )
    fig_heat.update_layout(
        height=360, paper_bgcolor="white",
        coloraxis_colorbar=dict(title="Score", tickvals=[0,50,65,100],
                                ticktext=["Avoid","Hold","Buy","Best"]),
        margin=dict(l=160, r=60, t=20, b=40), font=dict(family="Inter"),
    )
    st.plotly_chart(fig_heat, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Data Table
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown("### 📋 Full Score Table")
    tbl = filtered[["stadium","city","sport","team","overall","office","retail",
                    "hospitality","multifamily","demo_bonus","pop_growth","median_income"]].copy()
    tbl.columns = ["Stadium","City","Sport","Team","Overall","Office","Retail",
                   "Hospitality","Multifamily","Demo Bonus","Pop Growth","Median Income"]
    tbl["Rating"] = tbl["Overall"].apply(rating)
    tbl = tbl.sort_values("Overall", ascending=False)

    def hl(val):
        if isinstance(val, str):
            if val == "BUY":   return "background-color:#dcfce7;color:#15803d;font-weight:700"
            if val == "HOLD":  return "background-color:#fef9c3;color:#92400e;font-weight:700"
            if val == "AVOID": return "background-color:#fee2e2;color:#b91c1c;font-weight:700"
        try:
            fv = float(val)
            if fv >= 65: return "background-color:#dcfce7"
            if fv >= 50: return "background-color:#fef9c3"
            return "background-color:#fee2e2"
        except: return ""

    st.dataframe(
        tbl.style.format({
            "Overall":"{:.1f}","Office":"{:.1f}","Retail":"{:.1f}",
            "Hospitality":"{:.1f}","Multifamily":"{:.1f}","Demo Bonus":"{:.1f}",
            "Pop Growth": lambda x: f"{x:.1%}" if pd.notna(x) else "N/A",
            "Median Income": lambda x: f"${x:,.0f}" if pd.notna(x) else "N/A",
        }).applymap(hl, subset=["Overall","Office","Retail","Hospitality","Multifamily","Rating"]),
        use_container_width=True, height=420,
    )
    st.download_button("⬇️ Download CSV", data=tbl.to_csv(index=False),
                       file_name="stadium_re_scores_v2.csv", mime="text/csv")

    st.markdown("---")
    st.markdown("""<div class="score-explainer">
    <strong>Methodology (v2):</strong> Scores use piecewise-linear interpolation against absolute industry benchmarks
    (Excellent / Good / Fair / Poor), not peer-relative ranking. Cap rates are correctly scored as
    lower = more valuable. Multifamily rent growth blends 45% current + 55% 10-year average to prevent
    short-term corrections from masking structural market strength. A vacancy trend penalty of up to −15 pts
    is applied where current vacancy has deteriorated vs the 10-year average. A demographic quality
    bonus of up to +8 pts rewards high-income, high-growth submarkets. Suppressed CoStar hospitality data
    receives neutral (50) fills on missing metrics.
    <br><br>
    <strong>Source:</strong> CoStar Analytics, 1-Mile Custom Radius, April 2026.
    Analysis does not constitute financial advice.
    </div>""", unsafe_allow_html=True)
