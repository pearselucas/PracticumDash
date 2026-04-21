"""
Stadium District Real Estate Investment Analyzer - Final Version
Run: streamlit run app_final.py
Requires: pip install streamlit plotly pandas openpyxl
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np

# Standardize uploaded Excel/raw names to dashboard display names
DISPLAY_NAME_REMAP = {
    "Bank of America Properties": "Bank of America Stadium",
    "Spectrum Center Properties": "Spectrum Center",
    "Allegiant Stadium Properties": "Allegiant Stadium",
    "Bank of America Properties (Charlotte)": "Bank of America Stadium (Charlotte)",
    "Spectrum Center Properties (Charlotte)": "Spectrum Center (Charlotte)",
    "Allegiant Stadium Properties (Las Vegas)": "Allegiant Stadium (Las Vegas)",
}

def standardize_stadium_names(df, column='Stadium/Market'):
    """Rename workbook/raw labels to dashboard display labels without changing scores."""
    if df is None or column not in df.columns:
        return df
    out = df.copy()
    out[column] = out[column].replace(DISPLAY_NAME_REMAP)
    return out

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
    position: relative;
}

.bec-logo {
    position: absolute;
    top: 2rem;
    right: 3rem;
    height: 50px;
    width: auto;
    background: white;
    padding: 0.5rem 1rem;
    border-radius: 6px;
    font-weight: 700;
    font-size: 0.85rem;
    color: #1a365d;
}

.main-header h1 { 
    font-size: 2.2rem; margin: 0; font-weight: 800; 
    letter-spacing:-0.03em; 
    color: white;
}
.main-header p  { 
    font-size: 1rem; margin: 0.6rem 0 0; opacity: 0.95; 
    line-height: 1.6; max-width: 75%;
}
.main-header .subtitle-row { 
    display:flex; gap:2rem; margin-top:1rem; flex-wrap:wrap; 
}
.main-header .pill { 
    background: rgba(255,255,255,0.15); 
    border-radius:6px;
    padding:0.4rem 1rem; font-size:0.85rem; font-weight:600; 
}

.metric-card {
    background:#fff; 
    border:1px solid #e2e8f0; 
    border-radius:10px;
    padding:1.25rem 1.5rem; 
    text-align:center; 
    height:100%;
    box-shadow:0 2px 8px rgba(0,0,0,0.04); 
    transition:all 0.2s;
}
.metric-card:hover { 
    box-shadow:0 6px 20px rgba(0,0,0,0.08); 
    transform: translateY(-2px); 
}
.metric-card .label { 
    font-size:0.75rem; 
    color:#64748b; 
    text-transform:uppercase;
    letter-spacing:0.08em; 
    margin-bottom:0.4rem; 
    font-weight:600; 
}
.metric-card .value { 
    font-size:1.6rem; 
    font-weight:800; 
    color:#0f172a; 
}
.metric-card .sub   { 
    font-size:0.8rem; 
    color:#94a3b8; 
    margin-top:0.25rem; 
}

.score-badge-POTENTIAL   { 
    background:#d1fae5; 
    color:#065f46; 
    border:1.5px solid #6ee7b7; 
}
.score-badge-INFO  { 
    background:#fef3c7; 
    color:#78350f; 
    border:1.5px solid #fcd34d; 
}
.score-badge-AVOID { 
    background:#fee2e2; 
    color:#991b1b; 
    border:1.5px solid #fca5a5; 
}
.score-badge {
    display:inline-block; 
    padding:0.3rem 0.9rem; 
    border-radius:6px;
    font-weight:700; 
    font-size:0.85rem; 
    letter-spacing:0.05em;
}

.section-title {
    font-size:1.1rem; 
    font-weight:700; 
    color:#0f172a;
    border-bottom:2px solid #e2e8f0; 
    padding-bottom:0.5rem;
    margin:1.5rem 0 1rem;
}

.score-explainer {
    background:#f8fafc; 
    border:1px solid #cbd5e1; 
    border-radius:10px;
    padding:1.25rem 1.5rem; 
    font-size:0.88rem; 
    color:#475569; 
    margin:1rem 0;
    line-height: 1.7;
}
.score-explainer strong { 
    color:#0f172a; 
}

div[data-testid="stTabs"] button { 
    font-weight:600; 
    font-size:0.95rem; 
}

/* Team color indicators */
.team-indicator {
    display: inline-block;
    width: 12px;
    height: 12px;
    border-radius: 50%;
    margin-right: 0.5rem;
    vertical-align: middle;
}
</style>
""", unsafe_allow_html=True)

# ── Data Loading Functions ────────────────────────────────────────────────────
def load_matrix_from_excel(file_obj):
    """
    Load data directly from the Matrix tab in the Excel file.
    Reads from specific range E9:I17 where the data is located.
    NO CALCULATIONS - just reads the scores as-is.
    """
    try:
        # Get all sheet names
        xl = pd.ExcelFile(file_obj)
        
        # Find the Matrix sheet - prioritize exact match
        matrix_sheet = None
        
        # First look for exact match "Matrix"
        for sheet_name in xl.sheet_names:
            if sheet_name == 'Matrix':
                matrix_sheet = sheet_name
                break
        
        # If not found, look for any sheet containing "matrix" but not "weighting"
        if matrix_sheet is None:
            for sheet_name in xl.sheet_names:
                if 'matrix' in sheet_name.lower() and 'weighting' not in sheet_name.lower():
                    matrix_sheet = sheet_name
                    break
        
        # If still not found, use the last sheet
        if matrix_sheet is None:
            matrix_sheet = xl.sheet_names[-1]
        
        # Read the Matrix sheet starting from column E (index 4), row 9 (index 8)
        # This reads the range E9:I17
        df = pd.read_excel(
            file_obj, 
            sheet_name=matrix_sheet, 
            usecols="E:I",  # Columns E through I
            header=8        # Row 9 (0-indexed as 8) is the header
        )
        
        # The columns should now be: Stadium/Market, Office Score, Retail Score, Hospitality Score, Multifamily Score
        # Remove any rows with NaN in Stadium/Market
        df_clean = df[df.iloc[:, 0].notna()].copy()
        
        # Ensure we have exactly 5 columns
        if len(df_clean.columns) >= 5:
            df_clean = df_clean.iloc[:, :5].copy()
            
            # Set proper column names
            df_clean.columns = ['Stadium/Market', 'Office Score', 'Retail Score', 
                               'Hospitality Score', 'Multifamily Score']
            
            # Round scores to whole numbers for display
            df_clean['Office Score'] = pd.to_numeric(df_clean['Office Score'], errors='coerce').round(0)
            df_clean['Retail Score'] = pd.to_numeric(df_clean['Retail Score'], errors='coerce').round(0)
            df_clean['Hospitality Score'] = pd.to_numeric(df_clean['Hospitality Score'], errors='coerce').round(0)
            df_clean['Multifamily Score'] = pd.to_numeric(df_clean['Multifamily Score'], errors='coerce').round(0)
            
            # Remove any rows where all scores are NaN
            df_clean = df_clean.dropna(
                subset=['Office Score', 'Retail Score', 'Hospitality Score', 'Multifamily Score'], 
                how='all'
            )
            
            return standardize_stadium_names(df_clean)
        else:
            st.error(f"Expected 5 columns but found {len(df_clean.columns)}")
            return None
        
    except Exception as e:
        st.error(f"Error loading Matrix tab from Excel: {str(e)}")
        import traceback
        st.error(f"Details: {traceback.format_exc()}")
        return None

def load_weights_and_index_data(file_obj):
    """
    Load the weights and index calculation data for interactive weight adjustment.
    Returns weights_df and index_df.
    """
    try:
        # Read weights
        weights_df = pd.read_excel(file_obj, sheet_name='Asset Class Weighting Matrix', header=3)
        # Remove the total row (last row with NaN metric)
        weights_df = weights_df[weights_df['Metric'].notna()].copy()
        
        # Read index calculation data
        index_df = pd.read_excel(file_obj, sheet_name='Index Calculation', header=3)
        
        # Filter for stadium data only
        stadiums = ['MBS', 'Braves', 'American Airlines Center', 'AT&T Stadium', 
                   'Bank of America Stadium', 'Spectrum Center', 
                   'T Mobile Arena', 'Allegiant Stadium']
        index_df = index_df[index_df['Stadium/Market'].isin(stadiums)].copy()
        
        return weights_df, standardize_stadium_names(index_df)
        
    except Exception as e:
        st.error(f"Error loading weights and index data: {str(e)}")
        return None, None



def prepare_weights_for_editor(weights_df):
    """
    Convert workbook weights to an editable percentage table for Streamlit.
    Returns columns: Metric, Office, Retail, Hospitality, Multifamily
    where weight values are percentages from 0 to 100.
    """
    editor_df = weights_df.copy()
    editor_df = editor_df[editor_df['Metric'].notna()].copy()
    keep_cols = ['Metric', 'Office', 'Retail', 'Hospitality', 'Multifamily']
    editor_df = editor_df[keep_cols].copy()
    for col in ['Office', 'Retail', 'Hospitality', 'Multifamily']:
        editor_df[col] = pd.to_numeric(editor_df[col], errors='coerce').fillna(0.0) * 100.0
    return editor_df.reset_index(drop=True)


def validate_weight_totals(editor_df, tolerance=0.01):
    """
    Validate that each asset class column totals 100%.
    Returns (is_valid, totals_dict).
    """
    totals = {}
    valid = True
    for col in ['Office', 'Retail', 'Hospitality', 'Multifamily']:
        totals[col] = float(pd.to_numeric(editor_df[col], errors='coerce').fillna(0.0).sum())
        if abs(totals[col] - 100.0) > tolerance:
            valid = False
    return valid, totals


def convert_editor_to_custom_weights(editor_df):
    """
    Convert percentage editor table back to decimal-weight dict:
    {metric: {Office: 0.20, ...}}
    """
    custom_weights = {}
    for _, row in editor_df.iterrows():
        metric = row['Metric']
        custom_weights[metric] = {}
        for col in ['Office', 'Retail', 'Hospitality', 'Multifamily']:
            val = pd.to_numeric(pd.Series([row[col]]), errors='coerce').fillna(0.0).iloc[0]
            custom_weights[metric][col] = float(val) / 100.0
    return custom_weights

def calculate_scores_with_custom_weights(index_df, custom_weights_source):
    """
    Calculate scores using custom weights.
    
    Parameters:
    - index_df: DataFrame with indexed metric values for each stadium/product type
    - custom_weights: Dict with structure {metric_name: {asset_class: weight}}
    
    Returns:
    - DataFrame with Stadium/Market and scores for each asset class
    """
    if isinstance(custom_weights_source, pd.DataFrame):
        custom_weights = convert_editor_to_custom_weights(custom_weights_source)
    else:
        custom_weights = custom_weights_source

    results = []
    
    # Stadium name mapping
    name_mapping = {
        'MBS': 'Mercedez Benz (Atlanta)',
        'Braves': 'Truist Park (Atlanta)',
        'Bank of America Stadium': 'Bank of America Stadium (Charlotte)',
        'Spectrum Center': 'Spectrum Center (Charlotte)',
        'Allegiant Stadium': 'Allegiant Stadium (Las Vegas)',
        'T Mobile Arena': 'T Mobile Arena (Las Vegas)',
        'American Airlines Center': 'American Airlines Center (DFW)',
        'AT&T Stadium': 'AT&T Stadium (DFW)'
    }
    
    for stadium in index_df['Stadium/Market'].unique():
        stadium_data = index_df[index_df['Stadium/Market'] == stadium]
        
        scores = {'Stadium/Market': name_mapping.get(stadium, stadium)}
        
        for product_type in ['Office', 'Retail', 'Hospitality', 'Multifamily']:
            product_data = stadium_data[stadium_data['Product Type'] == product_type]
            
            if len(product_data) > 0:
                row_data = product_data.iloc[0]
                
                # Calculate weighted score using custom weights
                total_score = 0
                
                for metric, weights in custom_weights.items():
                    weight = weights.get(product_type, 0)
                    if weight > 0 and metric in row_data.index:
                        value = row_data[metric]
                        if pd.notna(value):
                            total_score += value * weight
                
                scores[f'{product_type} Score'] = round(total_score, 0)
            else:
                scores[f'{product_type} Score'] = 0
        
        results.append(scores)
    
    return pd.DataFrame(results)

def create_default_data():
    """Create default data with exact Matrix scores"""
    return pd.DataFrame([
        {'Stadium/Market': 'Mercedez Benz (Atlanta)', 'Office Score': 68, 'Retail Score': 74, 'Hospitality Score': 131, 'Multifamily Score': 49},
        {'Stadium/Market': 'Truist Park (Atlanta)', 'Office Score': 82, 'Retail Score': 128, 'Hospitality Score': 108, 'Multifamily Score': 46},
        {'Stadium/Market': 'American Airlines Center (DFW)', 'Office Score': 96, 'Retail Score': 98, 'Hospitality Score': 137, 'Multifamily Score': 79},
        {'Stadium/Market': 'AT&T Stadium (DFW)', 'Office Score': 67, 'Retail Score': 110, 'Hospitality Score': 27, 'Multifamily Score': 80},
        {'Stadium/Market': 'Bank of America Stadium (Charlotte)', 'Office Score': 84, 'Retail Score': 82, 'Hospitality Score': 127, 'Multifamily Score': 73},
        {'Stadium/Market': 'Spectrum Center (Charlotte)', 'Office Score': 84, 'Retail Score': 51, 'Hospitality Score': 87, 'Multifamily Score': 71},
        {'Stadium/Market': 'T Mobile Arena (Las Vegas)', 'Office Score': 55, 'Retail Score': 84, 'Hospitality Score': 79, 'Multifamily Score': 161},
        {'Stadium/Market': 'Allegiant Stadium (Las Vegas)', 'Office Score': 65, 'Retail Score': 30, 'Hospitality Score': 62, 'Multifamily Score': 11},
    ])

# Stadium metadata for visualization with team colors
STADIUM_META = {
    "Mercedez Benz (Atlanta)": {
        "city": "Atlanta, GA", 
        "team": "Atlanta Falcons / Atlanta United", 
        "sport": "NFL/MLS", 
        "short": "Mercedes-Benz",
        "color": "#A71930"  # Falcons Red
    },
    "Truist Park (Atlanta)": {
        "city": "Atlanta, GA", 
        "team": "Atlanta Braves", 
        "sport": "MLB", 
        "short": "Truist Park",
        "color": "#CE1141"  # Braves Red
    },
    "Bank of America Stadium (Charlotte)": {
        "city": "Charlotte, NC", 
        "team": "Carolina Panthers", 
        "sport": "NFL", 
        "short": "Bank of America",
        "color": "#0085CA"  # Panthers Blue
    },
    "Spectrum Center (Charlotte)": {
        "city": "Charlotte, NC", 
        "team": "Charlotte Hornets", 
        "sport": "NBA", 
        "short": "Spectrum Center",
        "color": "#1D1160"  # Hornets Purple
    },
    "Allegiant Stadium (Las Vegas)": {
        "city": "Las Vegas, NV", 
        "team": "Las Vegas Raiders", 
        "sport": "NFL", 
        "short": "Allegiant",
        "color": "#000000"  # Raiders Black/Silver
    },
    "T Mobile Arena (Las Vegas)": {
        "city": "Las Vegas, NV", 
        "team": "Vegas Golden Knights", 
        "sport": "NHL", 
        "short": "T-Mobile",
        "color": "#B4975A"  # Golden Knights Gold
    },
    "American Airlines Center (DFW)": {
        "city": "Dallas, TX", 
        "team": "Dallas Mavericks / Dallas Stars", 
        "sport": "NBA/NHL", 
        "short": "American Airlines",
        "color": "#00538C"  # Mavericks Blue
    },
    "AT&T Stadium (DFW)": {
        "city": "Arlington, TX", 
        "team": "Dallas Cowboys", 
        "sport": "NFL", 
        "short": "AT&T Stadium",
        "color": "#041E42"  # Cowboys Navy Blue
    },
}

# Use team colors for visualizations
TEAM_COLORS = {stadium: meta["color"] for stadium, meta in STADIUM_META.items()}

def rating(score):
    """Convert score to rating category"""
    if score > 100: return "POTENTIAL"
    if score >= 60: return "INFO"
    return "AVOID"

def rating_display(score):
    """Convert score to display text"""
    if score > 100: return "Potential Investment"
    if score >= 60: return "More Information Needed"
    return "Avoid This Asset Class"

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR & FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### Upload Data File")
    st.markdown("Upload your Excel template with Matrix tab containing final scores")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx'],
        help="Upload Excel file with Matrix tab (last sheet)"
    )
    
    # Initialize session state for workbook-driven custom weights
    if 'weights_df' not in st.session_state:
        st.session_state.weights_df = None
    if 'weights_editor_df' not in st.session_state:
        st.session_state.weights_editor_df = None
    if 'index_df' not in st.session_state:
        st.session_state.index_df = None
    if 'df_original' not in st.session_state:
        st.session_state.df_original = None
    if 'current_df' not in st.session_state:
        st.session_state.current_df = None
    if 'last_weight_validation' not in st.session_state:
        st.session_state.last_weight_validation = None

    if uploaded_file is not None:
        with st.spinner("Loading Matrix data from Excel..."):
            excel_bytes = uploaded_file.getvalue()
            df_original = load_matrix_from_excel(BytesIO(excel_bytes))
            weights_df, index_df = load_weights_and_index_data(BytesIO(excel_bytes))

            if df_original is not None and len(df_original) > 0:
                st.session_state.df_original = df_original.copy()
                if uploaded_file is not None and st.session_state.last_weight_validation != "success":
                    st.session_state.current_df = df_original.copy()
                elif st.session_state.current_df is None:
                    st.session_state.current_df = df_original.copy()

            if weights_df is not None and index_df is not None:
                st.session_state.weights_df = weights_df.copy()
                st.session_state.index_df = index_df.copy()
                if st.session_state.weights_editor_df is None:
                    st.session_state.weights_editor_df = prepare_weights_for_editor(weights_df)

        if st.session_state.df_original is None or len(st.session_state.df_original) == 0:
            st.error("Could not load Matrix data. Using default scores.")
            if st.session_state.current_df is None:
                st.session_state.current_df = create_default_data()
            df = st.session_state.current_df.copy()
            data_loaded = True
        else:
            st.success(f"Loaded {len(st.session_state.df_original)} stadiums from Matrix tab")

            if st.session_state.current_df is None:
                st.session_state.current_df = st.session_state.df_original.copy()

            df = st.session_state.current_df.copy()
            data_loaded = True
    else:
        st.info("Using default Matrix scores")
        if st.session_state.current_df is None:
            st.session_state.current_df = create_default_data()
        df = st.session_state.current_df.copy()
        data_loaded = True
    # Add metadata to dataframe
    if data_loaded:
        df['City'] = df['Stadium/Market'].map(lambda x: STADIUM_META.get(x, {}).get('city', ''))
        df['Team'] = df['Stadium/Market'].map(lambda x: STADIUM_META.get(x, {}).get('team', ''))
        df['Sport'] = df['Stadium/Market'].map(lambda x: STADIUM_META.get(x, {}).get('sport', ''))
        df['Short Name'] = df['Stadium/Market'].map(lambda x: STADIUM_META.get(x, {}).get('short', x.split()[0]))
        df['Team Color'] = df['Stadium/Market'].map(lambda x: STADIUM_META.get(x, {}).get('color', '#64748b'))
        
        # Calculate best use and top score
        df['Best Score'] = df[['Office Score', 'Retail Score', 'Hospitality Score', 'Multifamily Score']].max(axis=1)
        
        def get_best_use(row):
            scores = {
                'Office': row['Office Score'],
                'Retail': row['Retail Score'],
                'Hospitality': row['Hospitality Score'],
                'Multifamily': row['Multifamily Score']
            }
            return max(scores, key=scores.get)
        
        df['Highest & Best Use'] = df.apply(get_best_use, axis=1)
        
        st.markdown("---")
        st.markdown("### Filters")
        sports_filter = st.multiselect(
            "Sport",
            options=sorted(df['Sport'].unique().tolist()),
            default=sorted(df['Sport'].unique().tolist())
        )
        
        st.markdown("---")
        st.markdown("### Investment Criteria")
        st.caption("**Potential Investment**: Score > 100")
        st.caption("**More Information Needed**: Score 60-100")
        st.caption("**Avoid This Asset Class**: Score < 60")

# Apply filters
if data_loaded and len(df) > 0:
    filtered = df[df['Sport'].isin(sports_filter)].copy()
else:
    filtered = df

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="main-header">
    <div class="bec-logo">BLUE EAGLE CAPITAL</div>
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
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Executive Summary",
    "Asset Class Analysis",
    "Spider Charts",
    "Comparative Analytics",
    "Data Table",
    "Custom Weights"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Executive Summary
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Top Investment Opportunities by Best Use Score")
    
    top3 = filtered.nlargest(3, "Best Score")
    cols = st.columns(3)
    
    for idx, (_, row) in enumerate(top3.iterrows()):
        with cols[idx]:
            badge_class = f"score-badge-{rating(row['Best Score'])}"
            team_color = row['Team Color']
            
            st.markdown(f"""
            <div class="metric-card" style="border-top: 4px solid {team_color};">
                <div class="label">Rank #{idx+1}</div>
                <div class="value">{row['Stadium/Market']}</div>
                <div class="sub">{row['Team']} • {row['City']}</div>
                <div class="sub" style="margin-top:0.5rem;">
                    <span class="{badge_class} score-badge">{rating_display(row['Best Score'])}</span>
                </div>
                <div style="margin-top:0.75rem; font-size:1.3rem; font-weight:700; color:#3b82f6;">
                    {row['Best Score']:.0f}
                </div>
                <div class="sub">Best Use Score ({row['Highest & Best Use']})</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### Market Overview")
    
    m1, m2, m3, m4 = st.columns(4)
    
    # Calculate averages
    avg_office = filtered['Office Score'].mean()
    avg_retail = filtered['Retail Score'].mean()
    avg_hosp = filtered['Hospitality Score'].mean()
    avg_mf = filtered['Multifamily Score'].mean()
    overall_avg = (avg_office + avg_retail + avg_hosp + avg_mf) / 4
    
    m1.metric("Average Score (All Assets)", f"{overall_avg:.0f}")
    
    potential_count = (
        (filtered['Office Score'] > 100).sum() +
        (filtered['Retail Score'] > 100).sum() +
        (filtered['Hospitality Score'] > 100).sum() +
        (filtered['Multifamily Score'] > 100).sum()
    )
    m2.metric("Potential Investment Opportunities", f"{potential_count}")
    m3.metric("Total Markets Analyzed", f"{len(filtered)}")
    m4.metric("Asset Classes", "4")
    
    # Asset class leaders
    st.markdown("### Asset Class Leaders")
    
    ac1, ac2, ac3, ac4 = st.columns(4)
    
    for col, asset, label in [
        (ac1, "Office Score", "Office"),
        (ac2, "Retail Score", "Retail"),
        (ac3, "Hospitality Score", "Hospitality"),
        (ac4, "Multifamily Score", "Multifamily")
    ]:
        with col:
            top = filtered.nlargest(1, asset).iloc[0]
            team_color = top['Team Color']
            
            st.markdown(f"""
            <div class="metric-card" style="border-top: 4px solid {team_color};">
                <div class="label">{label} Leader</div>
                <div class="value" style="font-size:1rem;">{top['Stadium/Market']}</div>
                <div class="sub">{top[asset]:.0f} Score</div>
            </div>
            """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – Asset Class Analysis
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Asset Class Performance by Stadium")
    
    asset_tabs = st.tabs(["Office", "Retail", "Hospitality", "Multifamily"])
    
    def render_asset(tab, asset_col, asset_label):
        with tab:
            st.markdown(f"#### {asset_label} Investment Scores")
            
            asset_df = filtered.sort_values(asset_col, ascending=False)
            
            fig = go.Figure()
            # Use team colors for each stadium
            colors = [asset_df.loc[asset_df['Stadium/Market'] == stadium, 'Team Color'].values[0] 
                     for stadium in asset_df['Stadium/Market']]
            
            fig.add_trace(go.Bar(
                x=asset_df["Stadium/Market"],
                y=asset_df[asset_col],
                marker_color=colors,
                text=asset_df[asset_col].round(0),
                textposition='outside',
                textfont=dict(size=11, family="Inter", color="#0f172a"),
                hovertemplate="<b>%{x}</b><br>Score: %{y:.0f}<extra></extra>",
            ))
            
            fig.add_hline(y=100, line_dash="dash", line_color="#10b981", 
                         annotation_text="Potential Investment (>100)", annotation_position="right")
            fig.add_hline(y=60, line_dash="dash", line_color="#f59e0b",
                         annotation_text="More Info Needed (60-100)", annotation_position="right")
            
            fig.update_layout(
                height=400,
                showlegend=False,
                plot_bgcolor="#f8fafc",
                paper_bgcolor="white",
                xaxis=dict(tickangle=-45, showgrid=False),
                yaxis=dict(
                    title=f"{asset_label} Score", 
                    showgrid=True, 
                    gridcolor="#e2e8f0",
                    range=[0, asset_df[asset_col].max() * 1.15]  # Add 15% padding at top
                ),
                font=dict(family="Inter"),
                margin=dict(t=20, b=100)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown(f"##### Top 3 {asset_label} Markets")
            top3_asset = asset_df.head(3)
            
            for idx, (_, row) in enumerate(top3_asset.iterrows(), 1):
                team_color = row['Team Color']
                
                col1, col2, col3 = st.columns([3, 2, 2])
                with col1:
                    st.markdown(f'<span class="team-indicator" style="background-color:{team_color};"></span>**{idx}. {row["Stadium/Market"]}**', unsafe_allow_html=True)
                with col2:
                    badge_class = f"score-badge-{rating(row[asset_col])}"
                    st.markdown(f'<span class="{badge_class} score-badge">{row[asset_col]:.0f}</span>', 
                               unsafe_allow_html=True)
                with col3:
                    st.caption(row['City'])
    
    render_asset(asset_tabs[0], "Office Score", "Office")
    render_asset(asset_tabs[1], "Retail Score", "Retail")
    render_asset(asset_tabs[2], "Hospitality Score", "Hospitality")
    render_asset(asset_tabs[3], "Multifamily Score", "Multifamily")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Spider Charts
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### Stadium Comparison - Spider/Radar Charts")
    st.caption("Compare investment scores across all four asset classes for each stadium")
    
    # Allow user to select stadiums to compare
    selected_stadiums = st.multiselect(
        "Select stadiums to compare (up to 4 for clarity)",
        options=filtered['Stadium/Market'].tolist(),
        default=filtered.nlargest(3, 'Best Score')['Stadium/Market'].tolist()[:3]
    )
    
    if selected_stadiums:
        # Create spider chart
        fig = go.Figure()
        
        categories = ['Office', 'Retail', 'Hospitality', 'Multifamily']
        
        for stadium in selected_stadiums[:4]:  # Limit to 4 for readability
            stadium_data = filtered[filtered['Stadium/Market'] == stadium].iloc[0]
            values = [
                stadium_data['Office Score'],
                stadium_data['Retail Score'],
                stadium_data['Hospitality Score'],
                stadium_data['Multifamily Score']
            ]
            # Close the polygon
            values.append(values[0])
            
            team_color = stadium_data['Team Color']
            
            fig.add_trace(go.Scatterpolar(
                r=values,
                theta=categories + [categories[0]],
                fill='toself',
                name=stadium_data['Short Name'],
                line=dict(color=team_color, width=2),
                marker=dict(size=8)
            ))
        
        fig.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 180],
                    tickmode='linear',
                    tick0=0,
                    dtick=30,
                    showline=True,
                    showgrid=True,
                    gridcolor='#e2e8f0'
                ),
                angularaxis=dict(
                    showline=True,
                    showgrid=True,
                    gridcolor='#e2e8f0'
                )
            ),
            showlegend=True,
            height=600,
            paper_bgcolor="white",
            font=dict(family="Inter"),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5
            )
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Show individual stadium spider charts in a grid
        st.markdown("---")
        st.markdown("### Individual Stadium Profiles")
        
        cols = st.columns(2)
        for idx, stadium in enumerate(selected_stadiums):
            with cols[idx % 2]:
                stadium_data = filtered[filtered['Stadium/Market'] == stadium].iloc[0]
                
                fig_single = go.Figure()
                
                values = [
                    stadium_data['Office Score'],
                    stadium_data['Retail Score'],
                    stadium_data['Hospitality Score'],
                    stadium_data['Multifamily Score']
                ]
                values.append(values[0])
                
                team_color = stadium_data['Team Color']
                
                fig_single.add_trace(go.Scatterpolar(
                    r=values,
                    theta=categories + [categories[0]],
                    fill='toself',
                    line=dict(color=team_color, width=3),
                    marker=dict(size=10)
                ))
                
                fig_single.update_layout(
                    polar=dict(
                        radialaxis=dict(
                            visible=True,
                            range=[0, 180],
                            tickmode='linear',
                            tick0=0,
                            dtick=30
                        )
                    ),
                    showlegend=False,
                    height=400,
                    title=dict(text=stadium_data['Short Name'], font=dict(size=16)),
                    paper_bgcolor="white",
                    font=dict(family="Inter")
                )
                
                st.plotly_chart(fig_single, use_container_width=True)
                
                # Show scores below chart
                st.markdown(f"""
                **Best Use:** {stadium_data['Highest & Best Use']} ({stadium_data['Best Score']:.0f})  
                Office: {stadium_data['Office Score']:.0f} | Retail: {stadium_data['Retail Score']:.0f} | 
                Hospitality: {stadium_data['Hospitality Score']:.0f} | Multifamily: {stadium_data['Multifamily Score']:.0f}
                """)
    else:
        st.info("Please select at least one stadium to display spider charts")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Comparative Analytics
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown("### Quadrant Analysis")
    st.info("Analysis compares asset class scores. Top-right quadrant represents strongest performance across both dimensions.")
    
    def quadrant_chart(df, x_col, y_col, x_label, y_label, title):
        fig = go.Figure()
        
        for _, row in df.iterrows():
            fig.add_trace(go.Scatter(
                x=[row[x_col]],
                y=[row[y_col]],
                mode="markers+text",
                marker=dict(
                    size=16,
                    color=row['Team Color'],
                    line=dict(width=2, color="white")
                ),
                text=[row["Short Name"]],
                textposition="top center",
                textfont=dict(size=10, family="Inter"),
                showlegend=False,
                hovertemplate=(
                    f"<b>{row['Stadium/Market']}</b><br>"
                    f"{x_label}: {row[x_col]:.0f}<br>"
                    f"{y_label}: {row[y_col]:.0f}<extra></extra>"
                ),
            ))
        
        fig.add_hline(y=100, line_dash="dash", line_color="#10b981", opacity=0.5)
        fig.add_vline(x=100, line_dash="dash", line_color="#10b981", opacity=0.5)
        fig.add_hline(y=60, line_dash="dash", line_color="#f59e0b", opacity=0.5)
        fig.add_vline(x=60, line_dash="dash", line_color="#f59e0b", opacity=0.5)
        
        fig.update_layout(
            title=title,
            xaxis_title=x_label,
            yaxis_title=y_label,
            plot_bgcolor="#f8fafc",
            paper_bgcolor="white",
            height=450,
            xaxis=dict(showgrid=True, gridcolor="#e2e8f0", range=[0, max(df[x_col].max(), 180)]),
            yaxis=dict(showgrid=True, gridcolor="#e2e8f0", range=[0, max(df[y_col].max(), 180)]),
            font=dict(family="Inter"),
        )
        
        return fig
    
    q1, q2 = st.columns(2)
    with q1:
        st.plotly_chart(quadrant_chart(filtered, "Retail Score", "Multifamily Score",
            "Retail Score", "Multifamily Score", "Retail vs. Multifamily"), 
            use_container_width=True)
    with q2:
        st.plotly_chart(quadrant_chart(filtered, "Office Score", "Hospitality Score",
            "Office Score", "Hospitality Score", "Office vs. Hospitality"), 
            use_container_width=True)
    
    st.markdown("---")
    st.markdown("### Score Heatmap — All Stadiums × All Asset Classes")
    
    heat = filtered[["Stadium/Market","Office Score","Retail Score","Hospitality Score","Multifamily Score"]].set_index("Stadium/Market")
    fig_heat = px.imshow(
        heat.round(0),
        color_continuous_scale=[[0,"#fee2e2"],[0.33,"#fef3c7"],[0.56,"#d1fae5"],[1,"#065f46"]],
        zmin=0, zmax=180, text_auto=".0f", aspect="auto", labels=dict(color="Score"),
    )
    fig_heat.update_layout(
        height=400, paper_bgcolor="white",
        coloraxis_colorbar=dict(
            title="Score", 
            tickvals=[0,60,100,150,180],
            ticktext=["0","60","100","150","180"]
        ),
        margin=dict(l=250, r=60, t=20, b=40), 
        font=dict(family="Inter"),
    )
    st.plotly_chart(fig_heat, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – Data Table
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown("### Investment Scores — Complete Dataset")
    st.caption("Data sourced directly from Matrix tab - exact scores without modification")
    
    # Prepare table
    tbl = filtered[["Stadium/Market", "City", "Sport", "Team", "Office Score", "Retail Score", 
                    "Hospitality Score", "Multifamily Score", "Highest & Best Use", "Best Score"]].copy()
    
    # Add rating columns
    tbl["Office Rating"] = tbl["Office Score"].apply(rating_display)
    tbl["Retail Rating"] = tbl["Retail Score"].apply(rating_display)
    tbl["Hospitality Rating"] = tbl["Hospitality Score"].apply(rating_display)
    tbl["Multifamily Rating"] = tbl["Multifamily Score"].apply(rating_display)
    
    tbl = tbl.sort_values("Best Score", ascending=False)
    
    def hl(val):
        if isinstance(val, str):
            if val == "Potential Investment":   
                return "background-color:#d1fae5;color:#065f46;font-weight:700"
            if val == "More Information Needed":  
                return "background-color:#fef3c7;color:#78350f;font-weight:700"
            if val == "Avoid This Asset Class": 
                return "background-color:#fee2e2;color:#991b1b;font-weight:700"
        try:
            fv = float(val)
            if fv > 100: return "background-color:#d1fae5"
            if fv >= 60: return "background-color:#fef3c7"
            return "background-color:#fee2e2"
        except: return ""
    
    st.dataframe(
        tbl.style.format({
            "Office Score": "{:.0f}",
            "Retail Score": "{:.0f}",
            "Hospitality Score": "{:.0f}",
            "Multifamily Score": "{:.0f}",
            "Best Score": "{:.0f}",
        }).applymap(hl, subset=["Office Score", "Retail Score", "Hospitality Score", 
                                 "Multifamily Score", "Best Score", "Office Rating", 
                                 "Retail Rating", "Hospitality Rating", "Multifamily Rating"]),
        use_container_width=True, 
        height=500,
    )
    
    st.download_button(
        "Download Full Dataset (CSV)",
        data=tbl.to_csv(index=False),
        file_name="stadium_investment_scores.csv",
        mime="text/csv"
    )
    
    st.markdown("---")
    st.markdown("""<div class="score-explainer">
    <strong>About This Data:</strong> All scores displayed in this dashboard are sourced directly from 
    the Matrix tab (last sheet) of the uploaded Excel file without any modification or recalculation. 
    These scores represent the final weighted investment analysis for each stadium and asset class combination.
    <br><br>
    <strong>Investment Criteria:</strong>
    <ul>
        <li><strong>Potential Investment</strong> (Score > 100): Asset class shows exceptional fundamentals and strong investment potential</li>
        <li><strong>More Information Needed</strong> (Score 60-100): Asset class shows moderate performance and requires additional analysis before investment decision</li>
        <li><strong>Avoid This Asset Class</strong> (Score < 60): Asset class shows weak fundamentals, not recommended for investment</li>
    </ul>
    <strong>Data Source:</strong> CoStar Analytics, 1-Mile Custom Radius around each stadium, Q1 2026.
    <br><br>
    <strong>Disclaimer:</strong> This analysis is for informational purposes only and does not constitute 
    investment advice. Please consult with qualified financial and real estate professionals before making 
    investment decisions.
    </div>""", unsafe_allow_html=True)
    
    # Blue Eagle Capital Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; padding:2rem 0; color:#64748b; font-size:0.85rem;">
        <div style="font-weight:700; color:#003366; font-size:1.1rem; margin-bottom:0.5rem;">
            BLUE EAGLE CAPITAL
        </div>
        <div>Stadium District Real Estate Research & Analytics</div>
        <div style="margin-top:0.5rem; font-size:0.75rem;">
            Powered by CoStar Analytics | © 2026 Blue Eagle Capital. All rights reserved.
        </div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 - Custom Weights
# ══════════════════════════════════════════════════════════════════════════════
with tab6:
    st.markdown("### Custom Weighting Matrix")
    st.caption("Edit the 26 metric weights directly. Each asset class column must sum to 100 before the dashboard updates.")

    if uploaded_file is None or st.session_state.weights_df is None or st.session_state.index_df is None:
        st.info("Upload the Excel workbook to enable custom weights based on the Asset Class Weighting Matrix and Index Calculation sheets.")
    else:
        if st.session_state.weights_editor_df is None:
            st.session_state.weights_editor_df = prepare_weights_for_editor(st.session_state.weights_df)

        edited_weights_df = st.data_editor(
            st.session_state.weights_editor_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "Metric": st.column_config.TextColumn("Metric", disabled=True, width="medium"),
                "Office": st.column_config.NumberColumn("Office Weight %", min_value=0.0, max_value=100.0, step=0.1, format="%.2f"),
                "Retail": st.column_config.NumberColumn("Retail Weight %", min_value=0.0, max_value=100.0, step=0.1, format="%.2f"),
                "Hospitality": st.column_config.NumberColumn("Hospitality Weight %", min_value=0.0, max_value=100.0, step=0.1, format="%.2f"),
                "Multifamily": st.column_config.NumberColumn("Multifamily Weight %", min_value=0.0, max_value=100.0, step=0.1, format="%.2f"),
            },
            key="weights_editor_table"
        )

        totals_valid, totals = validate_weight_totals(edited_weights_df)

        c1, c2, c3, c4 = st.columns(4)
        for col_obj, asset_col in zip([c1, c2, c3, c4], ["Office", "Retail", "Hospitality", "Multifamily"]):
            total_val = totals[asset_col]
            with col_obj:
                if abs(total_val - 100.0) <= 0.01:
                    st.success(f"{asset_col}: {total_val:.2f}%")
                else:
                    st.error(f"{asset_col}: {total_val:.2f}%")

        b1, b2, b3 = st.columns([1, 1, 2])

        with b1:
            if st.button("Apply Weights", use_container_width=True):
                st.session_state.weights_editor_df = edited_weights_df.copy()
                totals_valid, totals = validate_weight_totals(st.session_state.weights_editor_df)

                if totals_valid:
                    recalculated_df = calculate_scores_with_custom_weights(
                        st.session_state.index_df,
                        st.session_state.weights_editor_df
                    )
                    st.session_state.current_df = standardize_stadium_names(recalculated_df.copy())
                    st.session_state.last_weight_validation = "success"
                    st.success("Custom weights applied. Dashboard scores and recommendations have been refreshed.")
                    st.rerun()
                else:
                    st.session_state.last_weight_validation = "invalid"
                    st.warning("Weights were not applied. Each asset class column must sum to 100. The dashboard is still using the last valid scores.")

        with b2:
            if st.button("Reset to Excel Defaults", use_container_width=True):
                st.session_state.weights_editor_df = prepare_weights_for_editor(st.session_state.weights_df)
                st.session_state.current_df = st.session_state.df_original.copy()
                st.session_state.last_weight_validation = "reset"
                st.rerun()

        with b3:
            st.caption("The rest of the dashboard uses the original Matrix scores until you apply a valid 100% weighting set for all four asset classes.")

        st.markdown("---")
        st.markdown("#### How this updates the dashboard")
        st.markdown(
            """
            - The editable grid mirrors the workbook's Asset Class Weighting Matrix.
            - When you click **Apply Weights**, the app recalculates the weighted scores from the Index Calculation sheet.
            - Those recalculated scores replace the dashboard's current score table, which updates the charts, rankings, matrix, and recommendations.
            - If any asset class does not total 100%, the dashboard keeps the last valid result so nothing breaks.
            """
        )
