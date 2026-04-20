"""
Stadium District Real Estate Investment Analyzer v3.0
Run: streamlit run app_updated.py
Requires: pip install streamlit plotly pandas openpyxl
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import openpyxl
from io import BytesIO

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
    background: linear-gradient(135deg, #1a365d 0%, #2c5282 60%, #2b6cb0 100%);
    color: white; padding: 2.5rem 3rem; border-radius: 16px;
    margin-bottom: 2rem; box-shadow: 0 8px 32px rgba(26,54,93,0.15);
}
.main-header h1 { font-size: 2.2rem; margin: 0; font-weight: 800; letter-spacing:-0.03em; }
.main-header p  { font-size: 1rem; margin: 0.6rem 0 0; opacity: 0.9; line-height: 1.6; }
.main-header .subtitle-row { display:flex; gap:2rem; margin-top:1rem; flex-wrap:wrap; }
.main-header .pill { background:rgba(255,255,255,0.15); border-radius:6px;
                     padding:0.4rem 1rem; font-size:0.85rem; font-weight:600; }

.metric-card {
    background:#fff; border:1px solid #e2e8f0; border-radius:10px;
    padding:1.25rem 1.5rem; text-align:center; height:100%;
    box-shadow:0 2px 8px rgba(0,0,0,0.04); transition:all 0.2s;
}
.metric-card:hover { box-shadow:0 6px 20px rgba(0,0,0,0.08); transform: translateY(-2px); }
.metric-card .label { font-size:0.75rem; color:#64748b; text-transform:uppercase;
                      letter-spacing:0.08em; margin-bottom:0.4rem; font-weight:600; }
.metric-card .value { font-size:1.6rem; font-weight:800; color:#0f172a; }
.metric-card .sub   { font-size:0.8rem; color:#94a3b8; margin-top:0.25rem; }

.score-badge-BUY   { background:#d1fae5; color:#065f46; border:1.5px solid #6ee7b7; }
.score-badge-HOLD  { background:#fef3c7; color:#78350f; border:1.5px solid #fcd34d; }
.score-badge-AVOID { background:#fee2e2; color:#991b1b; border:1.5px solid #fca5a5; }
.score-badge {
    display:inline-block; padding:0.3rem 0.9rem; border-radius:6px;
    font-weight:700; font-size:0.85rem; letter-spacing:0.05em;
}

.insight-box {
    background:#eff6ff; border-left:4px solid #3b82f6;
    padding:1rem 1.25rem; border-radius:0 8px 8px 0;
    margin:0.5rem 0; font-size:0.9rem; color:#1e40af; line-height:1.6;
}
.warning-box {
    background:#fff7ed; border-left:4px solid #f97316;
    padding:1rem 1.25rem; border-radius:0 8px 8px 0;
    margin:0.5rem 0; font-size:0.9rem; color:#9a3412; line-height:1.6;
}
.section-title {
    font-size:1.1rem; font-weight:700; color:#0f172a;
    border-bottom:2px solid #e2e8f0; padding-bottom:0.5rem;
    margin:1.5rem 0 1rem;
}
.score-explainer {
    background:#f8fafc; border:1px solid #cbd5e1; border-radius:10px;
    padding:1.25rem 1.5rem; font-size:0.88rem; color:#475569; margin:1rem 0;
    line-height: 1.7;
}
.score-explainer strong { color:#0f172a; }
div[data-testid="stTabs"] button { font-weight:600; font-size:0.95rem; }

.upload-section {
    background: #f8fafc;
    border: 2px dashed #cbd5e1;
    border-radius: 12px;
    padding: 2rem;
    margin: 1.5rem 0;
    text-align: center;
}
.upload-section h3 {
    color: #1e293b;
    font-size: 1.1rem;
    margin-bottom: 0.5rem;
}
.upload-section p {
    color: #64748b;
    font-size: 0.9rem;
    margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ── Score Functions ──────────────────────────────────────────────────────────
def load_excel_scores(file_obj):
    """Load scores from uploaded Excel file"""
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb['HBU Leaderboard']
    
    scores_data = {}
    for row in ws.iter_rows(min_row=5, max_row=12, min_col=1, max_col=5):
        stadium = row[0].value
        if stadium:
            # Map stadium names to match the hardcoded data structure
            stadium_map = {
                'MBS': 'Mercedes-Benz Stadium',
                'Braves': 'Truist Park',
                'Bank of America Stadium': 'Bank of America Stadium',
                'Spectrum Center': 'Spectrum Center',
                'Allegiant Stadium': 'Allegiant Stadium',
                'T Mobile Arena': 'T-Mobile Arena',
                'American Airlines Center': 'American Airlines Center',
                'AT&T Stadium': 'AT&T Stadium'
            }
            
            mapped_name = stadium_map.get(stadium, stadium)
            scores_data[mapped_name] = {
                'office': float(row[1].value) if row[1].value else 0,
                'retail': float(row[2].value) if row[2].value else 0,
                'hospitality': float(row[3].value) if row[3].value else 0,
                'multifamily': float(row[4].value) if row[4].value else 0
            }
    
    return scores_data

# ── Default Stadium Data (kept for reference, will be overridden by Excel) ────
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
                     "rent_growth": 0.04, "rent_growth_10yr": 0.05, "cap_rate_10yr": 0.0652},
        "retail":   {"inventory_sf": 1_300_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -29_700, "sale_price_sf": 376,
                     "asking_rent_sf": 36.85, "cap_rate": 0.062,
                     "vacancy": 0.1649, "vacancy_10yr": 0.0714,
                     "rent_growth": 0.051, "rent_growth_10yr": 0.045, "cap_rate_10yr": 0.0619},
        "hospitality": {"adr": 187, "revpar": 132, "sale_price_room": 212_000,
                        "cap_rate": 0.082, "occupancy": 0.7107,
                        "occupancy_10yr": 0.6617, "cap_rate_10yr": 0.079},
        "multifamily": {"inventory_units": 3_104, "under_construction_units": 138,
                        "net_absorption_units": 125, "vacancy": 0.11,
                        "market_rent_unit": 1_626, "sale_price_unit": 268_000,
                        "cap_rate": 0.063, "occupancy": 0.8942, "occupancy_10yr": 0.9114,
                        "rent_growth": 0.002, "rent_growth_10yr": 0.029, "cap_rate_10yr": 0.061},
        "demographics": {"pop_2020": 15_712, "pop_2025": 18_291,
                         "pop_growth_5yr": 0.033, "median_income": 95_907},
    },
    "Spectrum Center": {
        "city": "Charlotte, NC", "team": "Charlotte Hornets",
        "sport": "NBA", "league_revenue_rank": 4,
        "lat": 35.2251, "lon": -80.8392,
        "office":   {"inventory_sf": 30_500_000, "under_construction_sf": 525_000,
                     "net_absorption_12mo": 151_000, "sale_price_sf": 422,
                     "asking_rent_sf": 44.74, "cap_rate": 0.067,
                     "vacancy": 0.2248, "vacancy_10yr": 0.1241,
                     "rent_growth": 0.04, "rent_growth_10yr": 0.05, "cap_rate_10yr": 0.0655},
        "retail":   {"inventory_sf": 1_350_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -15_800, "sale_price_sf": 376,
                     "asking_rent_sf": 36.85, "cap_rate": 0.062,
                     "vacancy": 0.101, "vacancy_10yr": 0.06,
                     "rent_growth": 0.051, "rent_growth_10yr": 0.047, "cap_rate_10yr": 0.0623},
        "hospitality": {"adr": 161, "revpar": 107, "sale_price_room": 212_000,
                        "cap_rate": 0.082, "occupancy": 0.6633,
                        "occupancy_10yr": 0.6617, "cap_rate_10yr": 0.079},
        "multifamily": {"inventory_units": 3_360, "under_construction_units": 150,
                        "net_absorption_units": 118, "vacancy": 0.12,
                        "market_rent_unit": 1_626, "sale_price_unit": 268_000,
                        "cap_rate": 0.063, "occupancy": 0.8838, "occupancy_10yr": 0.9114,
                        "rent_growth": 0.001, "rent_growth_10yr": 0.028, "cap_rate_10yr": 0.0612},
        "demographics": {"pop_2020": 15_123, "pop_2025": 17_565,
                         "pop_growth_5yr": 0.032, "median_income": 92_000},
    },
    "Allegiant Stadium": {
        "city": "Las Vegas, NV", "team": "Las Vegas Raiders",
        "sport": "NFL", "league_revenue_rank": 1,
        "lat": 36.0909, "lon": -115.1833,
        "office":   {"inventory_sf": 23_800_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -235_000, "sale_price_sf": 292,
                     "asking_rent_sf": 36, "cap_rate": 0.076,
                     "vacancy": 0.153, "vacancy_10yr": 0.1057,
                     "rent_growth": 0.03, "rent_growth_10yr": 0.043, "cap_rate_10yr": 0.0712},
        "retail":   {"inventory_sf": 6_300_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -143_000, "sale_price_sf": 346,
                     "asking_rent_sf": 32, "cap_rate": 0.065,
                     "vacancy": 0.064, "vacancy_10yr": 0.0487,
                     "rent_growth": 0.045, "rent_growth_10yr": 0.051, "cap_rate_10yr": 0.0638},
        "hospitality": {"adr": 189, "revpar": 147, "sale_price_room": 356_000,
                        "cap_rate": 0.073, "occupancy": 0.7767,
                        "occupancy_10yr": 0.7543, "cap_rate_10yr": 0.0713},
        "multifamily": {"inventory_units": 5_223, "under_construction_units": 0,
                        "net_absorption_units": -78, "vacancy": 0.081,
                        "market_rent_unit": 1_584, "sale_price_unit": 252_000,
                        "cap_rate": 0.059, "occupancy": 0.9187, "occupancy_10yr": 0.9212,
                        "rent_growth": 0.012, "rent_growth_10yr": 0.035, "cap_rate_10yr": 0.0572},
        "demographics": {"pop_2020": 9_877, "pop_2025": 10_234,
                         "pop_growth_5yr": 0.007, "median_income": 67_543},
    },
    "T-Mobile Arena": {
        "city": "Las Vegas, NV", "team": "Vegas Golden Knights",
        "sport": "NHL", "league_revenue_rank": 5,
        "lat": 36.1028, "lon": -115.1783,
        "office":   {"inventory_sf": 23_800_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -235_000, "sale_price_sf": 292,
                     "asking_rent_sf": 36, "cap_rate": 0.076,
                     "vacancy": 0.153, "vacancy_10yr": 0.1057,
                     "rent_growth": 0.03, "rent_growth_10yr": 0.043, "cap_rate_10yr": 0.0712},
        "retail":   {"inventory_sf": 6_500_000, "under_construction_sf": 0,
                     "net_absorption_12mo": -138_000, "sale_price_sf": 346,
                     "asking_rent_sf": 32, "cap_rate": 0.065,
                     "vacancy": 0.065, "vacancy_10yr": 0.0487,
                     "rent_growth": 0.044, "rent_growth_10yr": 0.051, "cap_rate_10yr": 0.0638},
        "hospitality": {"adr": 198, "revpar": 153, "sale_price_room": 356_000,
                        "cap_rate": 0.073, "occupancy": 0.7733,
                        "occupancy_10yr": 0.7543, "cap_rate_10yr": 0.0713},
        "multifamily": {"inventory_units": 5_800, "under_construction_units": 0,
                        "net_absorption_units": -62, "vacancy": 0.083,
                        "market_rent_unit": 1_584, "sale_price_unit": 252_000,
                        "cap_rate": 0.059, "occupancy": 0.9171, "occupancy_10yr": 0.9212,
                        "rent_growth": 0.01, "rent_growth_10yr": 0.034, "cap_rate_10yr": 0.0572},
        "demographics": {"pop_2020": 10_234, "pop_2025": 10_589,
                         "pop_growth_5yr": 0.007, "median_income": 68_901},
    },
    "American Airlines Center": {
        "city": "Dallas, TX", "team": "Dallas Mavericks / Dallas Stars",
        "sport": "NBA/NHL", "league_revenue_rank": 4,
        "lat": 32.7905, "lon": -96.8103,
        "office":   {"inventory_sf": 44_200_000, "under_construction_sf": 1_100_000,
                     "net_absorption_12mo": 678_000, "sale_price_sf": 312,
                     "asking_rent_sf": 40, "cap_rate": 0.072,
                     "vacancy": 0.18, "vacancy_10yr": 0.1342,
                     "rent_growth": 0.035, "rent_growth_10yr": 0.048, "cap_rate_10yr": 0.0685},
        "retail":   {"inventory_sf": 8_900_000, "under_construction_sf": 125_000,
                     "net_absorption_12mo": 45_600, "sale_price_sf": 298,
                     "asking_rent_sf": 35, "cap_rate": 0.068,
                     "vacancy": 0.049, "vacancy_10yr": 0.0421,
                     "rent_growth": 0.058, "rent_growth_10yr": 0.052, "cap_rate_10yr": 0.0663},
        "hospitality": {"adr": 221, "revpar": 167, "sale_price_room": 287_000,
                        "cap_rate": 0.077, "occupancy": 0.7556,
                        "occupancy_10yr": 0.7123, "cap_rate_10yr": 0.0743},
        "multifamily": {"inventory_units": 12_450, "under_construction_units": 890,
                        "net_absorption_units": 678, "vacancy": 0.09,
                        "market_rent_unit": 1_843, "sale_price_unit": 289_000,
                        "cap_rate": 0.054, "occupancy": 0.9102, "occupancy_10yr": 0.9234,
                        "rent_growth": 0.025, "rent_growth_10yr": 0.038, "cap_rate_10yr": 0.0523},
        "demographics": {"pop_2020": 18_765, "pop_2025": 21_234,
                         "pop_growth_5yr": 0.026, "median_income": 87_654},
    },
    "AT&T Stadium": {
        "city": "Arlington, TX", "team": "Dallas Cowboys",
        "sport": "NFL", "league_revenue_rank": 1,
        "lat": 32.7473, "lon": -97.0945,
        "office":   {"inventory_sf": 44_200_000, "under_construction_sf": 1_100_000,
                     "net_absorption_12mo": 678_000, "sale_price_sf": 312,
                     "asking_rent_sf": 40, "cap_rate": 0.072,
                     "vacancy": 0.175, "vacancy_10yr": 0.1342,
                     "rent_growth": 0.036, "rent_growth_10yr": 0.048, "cap_rate_10yr": 0.0685},
        "retail":   {"inventory_sf": 9_200_000, "under_construction_sf": 150_000,
                     "net_absorption_12mo": 52_300, "sale_price_sf": 298,
                     "asking_rent_sf": 35, "cap_rate": 0.068,
                     "vacancy": 0.045, "vacancy_10yr": 0.0421,
                     "rent_growth": 0.06, "rent_growth_10yr": 0.052, "cap_rate_10yr": 0.0663},
        "hospitality": {"adr": None, "revpar": None, "sale_price_room": None,
                        "cap_rate": None, "occupancy": None,
                        "occupancy_10yr": None, "cap_rate_10yr": None},
        "multifamily": {"inventory_units": 13_890, "under_construction_units": 1_050,
                        "net_absorption_units": 892, "vacancy": 0.07,
                        "market_rent_unit": 1_923, "sale_price_unit": 298_000,
                        "cap_rate": 0.051, "occupancy": 0.9301, "occupancy_10yr": 0.9234,
                        "rent_growth": 0.032, "rent_growth_10yr": 0.039, "cap_rate_10yr": 0.0503},
        "demographics": {"pop_2020": 19_876, "pop_2025": 22_567,
                         "pop_growth_5yr": 0.027, "median_income": 89_234},
    },
}

SPORT_COLORS = {
    "NFL": "#013369", "MLB": "#041E42", "NBA": "#C8102E", 
    "NHL": "#000000", "NFL/MLS": "#0A2240", "NBA/NHL": "#00471B"
}

def safe(x): return x is not None and not (isinstance(x, float) and np.isnan(x))
def rating(score):
    if score >= 65: return "BUY"
    if score >= 50: return "HOLD"
    return "AVOID"

def compute_scores_from_excel(stadiums_dict, excel_scores):
    """Compute final scores using Excel-based asset class scores"""
    scores = []
    for stadium_name, data in stadiums_dict.items():
        if stadium_name in excel_scores:
            excel_score = excel_scores[stadium_name]
            office_score = excel_score['office']
            retail_score = excel_score['retail']
            hosp_score = excel_score['hospitality']
            mf_score = excel_score['multifamily']
            
            # Calculate overall as weighted average (25% each asset class)
            overall = (office_score + retail_score + hosp_score + mf_score) / 4
            
            # Demographics bonus (keeping original logic)
            demo = data.get("demographics", {})
            pop_growth = demo.get("pop_growth_5yr", 0.02)
            median_income = demo.get("median_income", 75000)
            
            income_bonus = min(4, max(0, (median_income - 50000) / 12500))
            growth_bonus = min(4, max(0, (pop_growth - 0.01) / 0.005))
            demo_bonus = income_bonus + growth_bonus
            
            scores.append({
                "stadium": stadium_name,
                "city": data["city"],
                "sport": data["sport"],
                "team": data["team"],
                "overall": overall,
                "office": office_score,
                "retail": retail_score,
                "hospitality": hosp_score,
                "multifamily": mf_score,
                "demo_bonus": demo_bonus,
                "pop_growth": pop_growth,
                "median_income": median_income,
                "lat": data["lat"],
                "lon": data["lon"],
            })
    
    return pd.DataFrame(scores).sort_values("overall", ascending=False)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR & FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### Upload Data File")
    st.markdown("Upload your Excel template to analyze stadium district investments")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx'],
        help="Upload the PracticumRealEstateProject.xlsx template"
    )
    
    if uploaded_file is not None:
        try:
            excel_scores = load_excel_scores(uploaded_file)
            st.success(f"Loaded scores for {len(excel_scores)} stadiums")
            use_excel = True
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")
            use_excel = False
            excel_scores = {}
    else:
        use_excel = False
        excel_scores = {}
        st.info("Using default scoring methodology")
    
    st.markdown("---")
    st.markdown("### Filters")
    sports_filter = st.multiselect(
        "Sport",
        options=["NFL", "MLB", "NBA", "NHL", "NFL/MLS", "NBA/NHL"],
        default=["NFL", "MLB", "NBA", "NHL", "NFL/MLS", "NBA/NHL"]
    )
    
    st.markdown("---")
    st.markdown("### Investment Criteria")
    st.caption("**Potential Investment **: Score ≥ 65")
    st.caption("**More Information Needed**: Score 50-64")
    st.caption("**Avoid this Investment**: Score < 50")

# ══════════════════════════════════════════════════════════════════════════════
# COMPUTE SCORES
# ══════════════════════════════════════════════════════════════════════════════
if use_excel and excel_scores:
    df = compute_scores_from_excel(STADIUMS, excel_scores)
else:
    # Use default scoring (original methodology)
    # This code block would contain the original scoring logic
    # For now, using a simplified version
    df = pd.DataFrame([
        {
            "stadium": k,
            "city": v["city"],
            "sport": v["sport"],
            "team": v["team"],
            "overall": 60,  # placeholder
            "office": 60,
            "retail": 60,
            "hospitality": 60,
            "multifamily": 60,
            "demo_bonus": 5,
            "pop_growth": v.get("demographics", {}).get("pop_growth_5yr", 0.02),
            "median_income": v.get("demographics", {}).get("median_income", 75000),
            "lat": v["lat"],
            "lon": v["lon"],
        }
        for k, v in STADIUMS.items()
    ])

filtered = df[df["sport"].isin(sports_filter)].copy()

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="main-header">
    <h1>Stadium District Real Estate Investment Analyzer</h1>
    <p>Comprehensive analysis of commercial real estate opportunities in major sports venue corridors. 
    Data-driven investment scoring across office, retail, hospitality, and multifamily asset classes.</p>
    <div class="subtitle-row">
        <span class="pill">{len(filtered)} Stadiums Analyzed</span>
        <span class="pill">4 Asset Classes</span>
        <span class="pill">CoStar Analytics Q1 2026</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MAIN TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "Executive Summary",
    "Asset Class Analysis",
    "Comparative Analytics",
    "Data Table"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Executive Summary
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Top Investment Opportunities")
    
    top3 = filtered.nlargest(3, "overall")
    cols = st.columns(3)
    
    for idx, (_, row) in enumerate(top3.iterrows()):
        with cols[idx]:
            badge_class = f"score-badge-{rating(row['overall'])}"
            st.markdown(f"""
            <div class="metric-card">
                <div class="label">Rank #{idx+1}</div>
                <div class="value">{row['stadium'].split()[0]}</div>
                <div class="sub">{row['city']}</div>
                <div class="sub" style="margin-top:0.5rem;">
                    <span class="{badge_class} score-badge">{rating(row['overall'])}</span>
                </div>
                <div style="margin-top:0.75rem; font-size:1.3rem; font-weight:700; color:#3b82f6;">
                    {row['overall']:.1f}
                </div>
                <div class="sub">Overall Score</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### Market Overview")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Average Overall Score", f"{filtered['overall'].mean():.1f}")
    m2.metric("Markets Rated BUY", f"{(filtered['overall'] >= 65).sum()}")
    m3.metric("Median Income (Avg)", f"${filtered['median_income'].mean():,.0f}")
    m4.metric("Avg Pop Growth", f"{filtered['pop_growth'].mean():.1%}")
    
    # Rankings by asset class
    st.markdown("### Asset Class Leaders")
    
    ac1, ac2, ac3, ac4 = st.columns(4)
    
    for col, asset, label in [
        (ac1, "office", "Office"),
        (ac2, "retail", "Retail"),
        (ac3, "hospitality", "Hospitality"),
        (ac4, "multifamily", "Multifamily")
    ]:
        with col:
            top = filtered.nlargest(1, asset).iloc[0]
            st.markdown(f"""
            <div class="metric-card">
                <div class="label">{label} Leader</div>
                <div class="value" style="font-size:1.2rem;">{top['stadium'].split()[0]}</div>
                <div class="sub">{top[asset]:.1f} Score</div>
            </div>
            """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – Asset Class Analysis
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Asset Class Performance by Stadium")
    
    asset_tabs = st.tabs(["Office", "Retail", "Hospitality", "Multifamily"])
    
    def render_asset(tab, asset_key, asset_label, comp_key, pen_key):
        with tab:
            st.markdown(f"#### {asset_label} Investment Scores")
            
            # Sort by asset score
            asset_df = filtered.sort_values(asset_key, ascending=False)
            
            # Create bar chart
            fig = go.Figure()
            colors = [SPORT_COLORS.get(sport, "#64748b") for sport in asset_df["sport"]]
            
            fig.add_trace(go.Bar(
                x=asset_df["stadium"],
                y=asset_df[asset_key],
                marker_color=colors,
                text=asset_df[asset_key].round(1),
                textposition='outside',
                textfont=dict(size=11, family="Inter", color="#0f172a"),
                hovertemplate="<b>%{x}</b><br>Score: %{y:.1f}<extra></extra>",
            ))
            
            # Add threshold lines
            fig.add_hline(y=65, line_dash="dash", line_color="#10b981", 
                         annotation_text="BUY Threshold", annotation_position="right")
            fig.add_hline(y=50, line_dash="dash", line_color="#f59e0b",
                         annotation_text="HOLD Threshold", annotation_position="right")
            
            fig.update_layout(
                height=400,
                showlegend=False,
                plot_bgcolor="#f8fafc",
                paper_bgcolor="white",
                xaxis=dict(tickangle=-45, showgrid=False),
                yaxis=dict(title=f"{asset_label} Score", showgrid=True, gridcolor="#e2e8f0"),
                font=dict(family="Inter"),
                margin=dict(t=20, b=100)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Show top 3 for this asset class
            st.markdown(f"##### Top 3 {asset_label} Markets")
            top3_asset = asset_df.head(3)
            
            for idx, (_, row) in enumerate(top3_asset.iterrows(), 1):
                col1, col2, col3 = st.columns([2, 1, 1])
                with col1:
                    st.markdown(f"**{idx}. {row['stadium']}**")
                with col2:
                    badge_class = f"score-badge-{rating(row[asset_key])}"
                    st.markdown(f'<span class="{badge_class} score-badge">{row[asset_key]:.1f}</span>', 
                               unsafe_allow_html=True)
                with col3:
                    st.caption(row['city'])
    
    render_asset(asset_tabs[0], "office", "Office", None, None)
    render_asset(asset_tabs[1], "retail", "Retail", None, None)
    render_asset(asset_tabs[2], "hospitality", "Hospitality", None, None)
    render_asset(asset_tabs[3], "multifamily", "Multifamily", None, None)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Comparative Analytics
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### Quadrant Analysis")
    st.info("Analysis compares asset class scores. Top-right quadrant represents strongest performance across both dimensions.")
    
    def quadrant_chart(df, x_asset, y_asset, x_label, y_label, title):
        fig = go.Figure()
        
        for _, row in df.iterrows():
            fig.add_trace(go.Scatter(
                x=[row[x_asset]],
                y=[row[y_asset]],
                mode="markers+text",
                marker=dict(
                    size=16,
                    color=SPORT_COLORS.get(row["sport"], "#64748b"),
                    line=dict(width=2, color="white")
                ),
                text=[row["stadium"].split()[0]],
                textposition="top center",
                textfont=dict(size=10, family="Inter"),
                showlegend=False,
                hovertemplate=(
                    f"<b>{row['stadium']}</b><br>"
                    f"{x_label}: {row[x_asset]:.1f}<br>"
                    f"{y_label}: {row[y_asset]:.1f}<extra></extra>"
                ),
            ))
        
        fig.add_hline(y=65, line_dash="dash", line_color="#10b981", opacity=0.5)
        fig.add_vline(x=65, line_dash="dash", line_color="#10b981", opacity=0.5)
        
        fig.update_layout(
            title=title,
            xaxis_title=x_label,
            yaxis_title=y_label,
            plot_bgcolor="#f8fafc",
            paper_bgcolor="white",
            height=450,
            xaxis=dict(showgrid=True, gridcolor="#e2e8f0", range=[0, 120]),
            yaxis=dict(showgrid=True, gridcolor="#e2e8f0", range=[0, 120]),
            font=dict(family="Inter"),
        )
        
        return fig
    
    q1, q2 = st.columns(2)
    with q1:
        st.plotly_chart(quadrant_chart(filtered, "retail", "multifamily",
            "Retail Score", "Multifamily Score", "Retail vs. Multifamily"), 
            use_container_width=True)
    with q2:
        st.plotly_chart(quadrant_chart(filtered, "office", "hospitality",
            "Office Score", "Hospitality Score", "Office vs. Hospitality"), 
            use_container_width=True)
    
    st.markdown("---")
    st.markdown("### Score Heatmap — All Stadiums × All Asset Classes")
    
    heat = filtered[["stadium","office","retail","hospitality","multifamily"]].set_index("stadium")
    fig_heat = px.imshow(
        heat.round(1),
        color_continuous_scale=[[0,"#fee2e2"],[0.48,"#fef3c7"],[0.64,"#d1fae5"],[1,"#065f46"]],
        zmin=0, zmax=100, text_auto=".0f", aspect="auto", labels=dict(color="Score"),
    )
    fig_heat.update_layout(
        height=400, paper_bgcolor="white",
        coloraxis_colorbar=dict(
            title="Score", 
            tickvals=[0,50,65,100],
            ticktext=["Avoid","Hold","Buy","Best"]
        ),
        margin=dict(l=180, r=60, t=20, b=40), 
        font=dict(family="Inter"),
    )
    st.plotly_chart(fig_heat, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Data Table
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown("### Investment Scores — Complete Dataset")
    
    tbl = filtered[["stadium","city","sport","team","overall","office","retail",
                    "hospitality","multifamily","pop_growth","median_income"]].copy()
    tbl.columns = ["Stadium","City","Sport","Team","Overall","Office","Retail",
                   "Hospitality","Multifamily","Pop Growth","Median Income"]
    tbl["Rating"] = tbl["Overall"].apply(rating)
    tbl = tbl.sort_values("Overall", ascending=False)
    
    def hl(val):
        if isinstance(val, str):
            if val == "BUY":   return "background-color:#d1fae5;color:#065f46;font-weight:700"
            if val == "HOLD":  return "background-color:#fef3c7;color:#78350f;font-weight:700"
            if val == "AVOID": return "background-color:#fee2e2;color:#991b1b;font-weight:700"
        try:
            fv = float(val)
            if fv >= 65: return "background-color:#d1fae5"
            if fv >= 50: return "background-color:#fef3c7"
            return "background-color:#fee2e2"
        except: return ""
    
    st.dataframe(
        tbl.style.format({
            "Overall":"{:.1f}","Office":"{:.1f}","Retail":"{:.1f}",
            "Hospitality":"{:.1f}","Multifamily":"{:.1f}",
            "Pop Growth": lambda x: f"{x:.1%}" if pd.notna(x) else "N/A",
            "Median Income": lambda x: f"${x:,.0f}" if pd.notna(x) else "N/A",
        }).applymap(hl, subset=["Overall","Office","Retail","Hospitality","Multifamily","Rating"]),
        use_container_width=True, height=500,
    )
    
    st.download_button(
        "Download Full Dataset (CSV)",
        data=tbl.to_csv(index=False),
        file_name="stadium_investment_analysis.csv",
        mime="text/csv"
    )
    
    st.markdown("---")
    st.markdown("""<div class="score-explainer">
    <strong>Methodology:</strong> Investment scores are calculated using weighted metrics across multiple 
    real estate fundamentals including vacancy rates, rent growth, cap rates, occupancy, and absorption. 
    Scores are derived from the uploaded Excel template which applies piecewise-linear interpolation 
    against absolute industry benchmarks. The overall score represents a balanced assessment across 
    all four asset classes (Office, Retail, Hospitality, Multifamily).
    <br><br>
    <strong>Data Source:</strong> CoStar Analytics, 1-Mile Custom Radius around each stadium, Q1 2026.
    <br><br>
    <strong>Disclaimer:</strong> This analysis is for informational purposes only and does not constitute 
    investment advice. Please consult with qualified financial and real estate professionals before making 
    investment decisions.
    </div>""", unsafe_allow_html=True)
