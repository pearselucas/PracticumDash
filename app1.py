"""
Stadium District Real Estate Investment Analyzer v4.0
Run: streamlit run app_updated.py
Requires: pip install streamlit plotly pandas openpyxl
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import openpyxl

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

.score-badge-POTENTIAL   { background:#d1fae5; color:#065f46; border:1.5px solid #6ee7b7; }
.score-badge-INFO  { background:#fef3c7; color:#78350f; border:1.5px solid #fcd34d; }
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
</style>
""", unsafe_allow_html=True)

# ── Data Loading Functions ────────────────────────────────────────────────────
def load_excel_data(file_obj):
    """Load all data from the HBU Leaderboard sheet"""
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb['HBU Leaderboard']
    
    data = []
    # Stadium name mapping for display
    stadium_display_names = {
        'MBS': 'Mercedes-Benz Stadium',
        'Braves': 'Truist Park',
        'Bank of America Properties': 'Bank of America Stadium',
        'Spectrum Center Properties': 'Spectrum Center',
        'Allegiant Stadium Properties': 'Allegiant Stadium',
        'T Mobile Arena': 'T-Mobile Arena',
        'American Airlines Center': 'American Airlines Center',
        'AT&T Stadium': 'AT&T Stadium'
    }
    
    # Read rows 5-12 (data rows)
    for row_idx in range(5, 13):
        row = ws[row_idx]
        stadium_key = row[0].value
        if stadium_key:
            display_name = stadium_display_names.get(stadium_key, stadium_key)
            data.append({
                'Stadium': display_name,
                'Office Score': float(row[1].value) if row[1].value else 0,
                'Retail Score': float(row[2].value) if row[2].value else 0,
                'Hospitality Score': float(row[3].value) if row[3].value else 0,
                'Multifamily Score': float(row[4].value) if row[4].value else 0,
                'Highest & Best Use': row[9].value if row[9].value else '',
                'Top Score': float(row[10].value) if row[10].value else 0,
            })
    
    return pd.DataFrame(data)

# Stadium metadata for visualization
STADIUM_META = {
    "Mercedes-Benz Stadium": {"city": "Atlanta, GA", "team": "Atlanta Falcons / Atlanta United", "sport": "NFL/MLS"},
    "Truist Park": {"city": "Atlanta, GA", "team": "Atlanta Braves", "sport": "MLB"},
    "Bank of America Stadium": {"city": "Charlotte, NC", "team": "Carolina Panthers", "sport": "NFL"},
    "Spectrum Center": {"city": "Charlotte, NC", "team": "Charlotte Hornets", "sport": "NBA"},
    "Allegiant Stadium": {"city": "Las Vegas, NV", "team": "Las Vegas Raiders", "sport": "NFL"},
    "T-Mobile Arena": {"city": "Las Vegas, NV", "team": "Vegas Golden Knights", "sport": "NHL"},
    "American Airlines Center": {"city": "Dallas, TX", "team": "Dallas Mavericks / Dallas Stars", "sport": "NBA/NHL"},
    "AT&T Stadium": {"city": "Arlington, TX", "team": "Dallas Cowboys", "sport": "NFL"},
}

SPORT_COLORS = {
    "NFL": "#013369", "MLB": "#041E42", "NBA": "#C8102E", 
    "NHL": "#000000", "NFL/MLS": "#0A2240", "NBA/NHL": "#00471B"
}

def rating(score):
    """Convert score to rating category"""
    if score >= 65: return "POTENTIAL"
    if score >= 50: return "INFO"
    return "AVOID"

def rating_display(score):
    """Convert score to display text"""
    if score >= 65: return "Potential Investment"
    if score >= 50: return "More Information Needed"
    return "Avoid This Asset Class"

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR & FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### Upload Data File")
    st.markdown("Upload your Excel template to visualize stadium district investment analysis")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx'],
        help="Upload the PracticumRealEstateProject.xlsx template"
    )
    
    if uploaded_file is not None:
        try:
            df = load_excel_data(uploaded_file)
            # Add metadata
            df['City'] = df['Stadium'].map(lambda x: STADIUM_META.get(x, {}).get('city', ''))
            df['Team'] = df['Stadium'].map(lambda x: STADIUM_META.get(x, {}).get('team', ''))
            df['Sport'] = df['Stadium'].map(lambda x: STADIUM_META.get(x, {}).get('sport', ''))
            
            st.success(f"Loaded data for {len(df)} stadiums")
            data_loaded = True
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")
            data_loaded = False
            df = pd.DataFrame()
    else:
        data_loaded = False
        df = pd.DataFrame()
        st.warning("Please upload an Excel file to begin analysis")
    
    if data_loaded:
        st.markdown("---")
        st.markdown("### Filters")
        sports_filter = st.multiselect(
            "Sport",
            options=df['Sport'].unique().tolist(),
            default=df['Sport'].unique().tolist()
        )
        
        st.markdown("---")
        st.markdown("### Investment Criteria")
        st.caption("**Potential Investment**: Score ≥ 65")
        st.caption("**More Information Needed**: Score 50-64")
        st.caption("**Avoid This Asset Class**: Score < 50")

# Apply filters if data is loaded
if data_loaded and len(df) > 0:
    filtered = df[df['Sport'].isin(sports_filter)].copy()
else:
    filtered = df

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="main-header">
    <h1>Stadium District Real Estate Investment Analyzer</h1>
    <p>Comprehensive analysis of commercial real estate opportunities in major sports venue corridors. 
    Data-driven investment scoring across office, retail, hospitality, and multifamily asset classes.</p>
    <div class="subtitle-row">
        <span class="pill">{len(filtered) if data_loaded else 0} Stadiums Analyzed</span>
        <span class="pill">4 Asset Classes</span>
        <span class="pill">CoStar Analytics Q1 2026</span>
    </div>
</div>
""", unsafe_allow_html=True)

if not data_loaded:
    st.info("👆 Please upload your Excel file using the sidebar to begin the analysis")
    st.stop()

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
    st.markdown("### Top Investment Opportunities by Best Use Score")
    
    top3 = filtered.nlargest(3, "Top Score")
    cols = st.columns(3)
    
    for idx, (_, row) in enumerate(top3.iterrows()):
        with cols[idx]:
            badge_class = f"score-badge-{rating(row['Top Score'])}"
            st.markdown(f"""
            <div class="metric-card">
                <div class="label">Rank #{idx+1}</div>
                <div class="value">{row['Stadium']}</div>
                <div class="sub">{row['City']}</div>
                <div class="sub" style="margin-top:0.5rem;">
                    <span class="{badge_class} score-badge">{rating_display(row['Top Score'])}</span>
                </div>
                <div style="margin-top:0.75rem; font-size:1.3rem; font-weight:700; color:#3b82f6;">
                    {row['Top Score']:.1f}
                </div>
                <div class="sub">Best Use Score ({row['Highest & Best Use']})</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### Market Overview")
    
    m1, m2, m3, m4 = st.columns(4)
    
    # Calculate averages across all asset classes
    avg_office = filtered['Office Score'].mean()
    avg_retail = filtered['Retail Score'].mean()
    avg_hosp = filtered['Hospitality Score'].mean()
    avg_mf = filtered['Multifamily Score'].mean()
    overall_avg = (avg_office + avg_retail + avg_hosp + avg_mf) / 4
    
    m1.metric("Average Score (All Assets)", f"{overall_avg:.1f}")
    
    # Count potential investments across all asset classes
    potential_count = (
        (filtered['Office Score'] >= 65).sum() +
        (filtered['Retail Score'] >= 65).sum() +
        (filtered['Hospitality Score'] >= 65).sum() +
        (filtered['Multifamily Score'] >= 65).sum()
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
            st.markdown(f"""
            <div class="metric-card">
                <div class="label">{label} Leader</div>
                <div class="value" style="font-size:1rem;">{top['Stadium']}</div>
                <div class="sub">{top[asset]:.1f} Score</div>
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
            
            # Sort by asset score
            asset_df = filtered.sort_values(asset_col, ascending=False)
            
            # Create bar chart
            fig = go.Figure()
            colors = [SPORT_COLORS.get(sport, "#64748b") for sport in asset_df["Sport"]]
            
            fig.add_trace(go.Bar(
                x=asset_df["Stadium"],
                y=asset_df[asset_col],
                marker_color=colors,
                text=asset_df[asset_col].round(1),
                textposition='outside',
                textfont=dict(size=11, family="Inter", color="#0f172a"),
                hovertemplate="<b>%{x}</b><br>Score: %{y:.1f}<extra></extra>",
            ))
            
            # Add threshold lines
            fig.add_hline(y=65, line_dash="dash", line_color="#10b981", 
                         annotation_text="Potential Investment", annotation_position="right")
            fig.add_hline(y=50, line_dash="dash", line_color="#f59e0b",
                         annotation_text="More Information Needed", annotation_position="right")
            
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
                col1, col2, col3 = st.columns([3, 2, 2])
                with col1:
                    st.markdown(f"**{idx}. {row['Stadium']}**")
                with col2:
                    badge_class = f"score-badge-{rating(row[asset_col])}"
                    st.markdown(f'<span class="{badge_class} score-badge">{row[asset_col]:.1f}</span>', 
                               unsafe_allow_html=True)
                with col3:
                    st.caption(row['City'])
    
    render_asset(asset_tabs[0], "Office Score", "Office")
    render_asset(asset_tabs[1], "Retail Score", "Retail")
    render_asset(asset_tabs[2], "Hospitality Score", "Hospitality")
    render_asset(asset_tabs[3], "Multifamily Score", "Multifamily")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Comparative Analytics
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
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
                    color=SPORT_COLORS.get(row["Sport"], "#64748b"),
                    line=dict(width=2, color="white")
                ),
                text=[row["Stadium"].split()[0]],
                textposition="top center",
                textfont=dict(size=10, family="Inter"),
                showlegend=False,
                hovertemplate=(
                    f"<b>{row['Stadium']}</b><br>"
                    f"{x_label}: {row[x_col]:.1f}<br>"
                    f"{y_label}: {row[y_col]:.1f}<extra></extra>"
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
        st.plotly_chart(quadrant_chart(filtered, "Retail Score", "Multifamily Score",
            "Retail Score", "Multifamily Score", "Retail vs. Multifamily"), 
            use_container_width=True)
    with q2:
        st.plotly_chart(quadrant_chart(filtered, "Office Score", "Hospitality Score",
            "Office Score", "Hospitality Score", "Office vs. Hospitality"), 
            use_container_width=True)
    
    st.markdown("---")
    st.markdown("### Score Heatmap — All Stadiums × All Asset Classes")
    
    heat = filtered[["Stadium","Office Score","Retail Score","Hospitality Score","Multifamily Score"]].set_index("Stadium")
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
            ticktext=["Avoid","More Info","Potential","Best"]
        ),
        margin=dict(l=200, r=60, t=20, b=40), 
        font=dict(family="Inter"),
    )
    st.plotly_chart(fig_heat, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Data Table
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown("### Investment Scores — Complete Dataset")
    st.caption("Data sourced directly from HBU Leaderboard analysis")
    
    # Prepare table - use exact data from Excel
    tbl = filtered[["Stadium", "City", "Sport", "Team", "Office Score", "Retail Score", 
                    "Hospitality Score", "Multifamily Score", "Highest & Best Use", "Top Score"]].copy()
    
    # Add rating columns for each asset class
    tbl["Office Rating"] = tbl["Office Score"].apply(rating_display)
    tbl["Retail Rating"] = tbl["Retail Score"].apply(rating_display)
    tbl["Hospitality Rating"] = tbl["Hospitality Score"].apply(rating_display)
    tbl["Multifamily Rating"] = tbl["Multifamily Score"].apply(rating_display)
    
    # Sort by top score
    tbl = tbl.sort_values("Top Score", ascending=False)
    
    def hl(val):
        """Highlight cells based on score"""
        if isinstance(val, str):
            if val == "Potential Investment":   
                return "background-color:#d1fae5;color:#065f46;font-weight:700"
            if val == "More Information Needed":  
                return "background-color:#fef3c7;color:#78350f;font-weight:700"
            if val == "Avoid This Asset Class": 
                return "background-color:#fee2e2;color:#991b1b;font-weight:700"
        try:
            fv = float(val)
            if fv >= 65: return "background-color:#d1fae5"
            if fv >= 50: return "background-color:#fef3c7"
            return "background-color:#fee2e2"
        except: return ""
    
    st.dataframe(
        tbl.style.format({
            "Office Score": "{:.1f}",
            "Retail Score": "{:.1f}",
            "Hospitality Score": "{:.1f}",
            "Multifamily Score": "{:.1f}",
            "Top Score": "{:.1f}",
        }).applymap(hl, subset=["Office Score", "Retail Score", "Hospitality Score", 
                                 "Multifamily Score", "Top Score", "Office Rating", 
                                 "Retail Rating", "Hospitality Rating", "Multifamily Rating"]),
        use_container_width=True, 
        height=500,
    )
    
    st.download_button(
        "Download Full Dataset (CSV)",
        data=tbl.to_csv(index=False),
        file_name="stadium_investment_analysis.csv",
        mime="text/csv"
    )
    
    st.markdown("---")
    st.markdown("""<div class="score-explainer">
    <strong>About This Data:</strong> All scores displayed in this dashboard are sourced directly from 
    the uploaded Excel template without modification. The analysis reflects weighted investment scores 
    calculated using comprehensive real estate metrics including vacancy rates, rent growth, cap rates, 
    occupancy, and absorption across a 1-mile radius of each stadium.
    <br><br>
    <strong>Investment Criteria:</strong>
    <ul>
        <li><strong>Potential Investment</strong> (Score ≥ 65): Asset class shows strong fundamentals and investment potential</li>
        <li><strong>More Information Needed</strong> (Score 50-64): Asset class requires additional analysis before investment decision</li>
        <li><strong>Avoid This Asset Class</strong> (Score < 50): Asset class shows weak fundamentals, not recommended for investment</li>
    </ul>
    <strong>Data Source:</strong> CoStar Analytics, 1-Mile Custom Radius around each stadium, Q1 2026.
    <br><br>
    <strong>Disclaimer:</strong> This analysis is for informational purposes only and does not constitute 
    investment advice. Please consult with qualified financial and real estate professionals before making 
    investment decisions.
    </div>""", unsafe_allow_html=True)
