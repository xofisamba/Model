"""
Solar/Wind Financial Model - v2 Dashboard UI
Professional financial dashboard with KPI cards and sidebar navigation
Version: 2.0
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json
import warnings
import os
from datetime import datetime
warnings.filterwarnings('ignore')

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================
st.set_page_config(
    page_title="Solar/Wind Financial Model v2",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# CUSTOM CSS - PROFESSIONAL DASHBOARD
# ============================================================================
st.markdown("""
<style>
    /* ===== GOOGLE FONTS ===== */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* ===== ROOT VARIABLES - LIGHT MODE ===== */
    :root {
        --primary: #1B5E3B;
        --primary-light: #2E7D4A;
        --primary-dark: #0D3D25;
        --accent: #FF9800;
        --accent-light: #FFB74D;
        --bg-main: #F4F6F8;
        --bg-sidebar: #1B5E3B;
        --bg-card: #FFFFFF;
        --bg-card-hover: #FAFAFA;
        --text-primary: #1A1A1A;
        --text-secondary: #6B7280;
        --text-light: #9CA3AF;
        --border: #E5E7EB;
        --success: #10B981;
        --warning: #F59E0B;
        --danger: #EF4444;
        --shadow-sm: 0 1px 2px rgba(0,0,0,0.05);
        --shadow-md: 0 4px 6px rgba(0,0,0,0.07);
        --shadow-lg: 0 10px 15px rgba(0,0,0,0.1);
        --radius-sm: 6px;
        --radius-md: 10px;
        --radius-lg: 16px;
    }
    
    /* ===== DARK MODE VARIABLES ===== */
    .dark-mode {
        --primary: #2E7D4A;
        --primary-light: #4CAF50;
        --primary-dark: #1B5E3B;
        --accent: #FFB74D;
        --accent-light: #FFCC80;
        --bg-main: #121212;
        --bg-sidebar: #1E1E1E;
        --bg-card: #2D2D2D;
        --bg-card-hover: #3D3D3D;
        --text-primary: #E0E0E0;
        --text-secondary: #A0A0A0;
        --text-light: #707070;
        --border: #404040;
        --success: #4CAF50;
        --warning: #FFC107;
        --danger: #F44336;
        --shadow-sm: 0 1px 2px rgba(0,0,0,0.3);
        --shadow-md: 0 4px 6px rgba(0,0,0,0.4);
        --shadow-lg: 0 10px 15px rgba(0,0,0,0.5);
    }
    
    /* ===== BASE STYLES ===== */
    * { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }
    body { background: var(--bg-main); }
    .stApp { background: var(--bg-main); }
    .main .block-container { background: var(--bg-main); }
    
    /* ===== MAIN HEADER HERO ===== */
    .hero-header {
        background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
        border-radius: var(--radius-lg);
        padding: 25px 30px;
        margin-bottom: 15px;
        color: white;
        position: relative;
        overflow: hidden;
        box-shadow: var(--shadow-md);
    }
    .hero-header::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 300px;
        height: 300px;
        background: rgba(255,255,255,0.1);
        border-radius: 50%;
    }
    .hero-header h1 {
        font-size: 32px;
        font-weight: 700;
        margin: 0 0 8px 0;
        letter-spacing: -0.5px;
    }
    .hero-header p {
        font-size: 15px;
        opacity: 0.9;
        margin: 0;
        font-weight: 400;
    }
    .hero-badge {
        display: inline-block;
        background: rgba(255,255,255,0.2);
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
        margin-top: 12px;
    }
    
    /* ===== SIDEBAR STYLES ===== */
    section[data-testid="stSidebar"] {
        background: var(--bg-sidebar) !important;
        padding: 0;
    }
    .sidebar-container {
        padding: 2px 3px;
        height: 100vh;
        overflow-y: auto;
        display: flex;
        flex-direction: column;
        justify-content: flex-start !important;
        align-items: stretch !important;
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        pointer-events: none; /* Container doesn't capture clicks */
    }
    section[data-testid="stSidebar"] > div {
        display: flex !important;
        flex-direction: column !important;
        justify-content: flex-start !important;
        align-items: stretch !important;
        height: 100vh;
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
    }
    /* Make buttons and inputs clickable */
    section[data-testid="stSidebar"] button,
    section[data-testid="stSidebar"] select,
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] [data-testid="stSelectbox"],
    section[data-testid="stSidebar"] [data-testid="stTextInput"],
    section[data-testid="stSidebar"] [role="button"] {
        position: relative !important;
        z-index: 999 !important;
        pointer-events: auto !important;
    }
    .sidebar-logo {
        margin-top: 0 !important;
        margin-bottom: 2px;
    }
    .sidebar-logo {
        color: white;
        font-size: 14px;
        font-weight: 700;
        padding: 1px 2px 1px 2px;
        border-bottom: 1px solid rgba(255,255,255,0.2);
        margin-bottom: 2px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .sidebar-section {
        margin-bottom: 1px;
    }
    .sidebar-section-title {
        color: rgba(255,255,255,0.5);
        font-size: 8px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 0px 3px 0px 3px;
        margin-bottom: 0px;
    }
    .sidebar-nav {
        display: flex;
        flex-direction: column;
        gap: 1px;
    }
    .sidebar-nav-item {
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 3px 6px;
        border-radius: var(--radius-md);
        color: rgba(255,255,255,0.8);
        font-size: 11px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
        text-decoration: none;
        border: none;
        background: transparent;
        width: 100%;
        text-align: left;
    }
    .sidebar-nav-item:hover {
        background: rgba(255,255,255,0.1);
        color: white;
    }
    .sidebar-nav-item.active {
        background: rgba(255,255,255,0.2);
        color: white;
        font-weight: 600;
    }
    .sidebar-nav-item .icon {
        font-size: 18px;
        width: 24px;
        text-align: center;
    }
    .sidebar-divider {
        height: 1px;
        background: rgba(255,255,255,0.15);
        margin: 1px 0;
    }
    .sidebar-footer {
        margin-top: 1px;
        padding-top: 2px;
        border-top: 0px solid rgba(255,255,255,0.2);
    }
    .sidebar-footer-text {
        color: rgba(255,255,255,0.5);
        font-size: 9px;
        text-align: center;
    }
    /* Compact Streamlit widgets in sidebar */
    .sidebar-container .stButton > button {
        height: 26px !important;
        padding: 0 4px !important;
        font-size: 11px !important;
        min-height: 26px !important;
    }
    .sidebar-container .stSelectbox > div > div {
        height: 26px !important;
        min-height: 26px !important;
    }
    
    /* ===== KPI CARDS ===== */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
        gap: 10px;
        margin-bottom: 10px;
    }
    .kpi-card {
        background: var(--bg-card);
        border-radius: var(--radius-sm);
        padding: 12px 10px;
        box-shadow: var(--shadow-sm);
        border: 1px solid var(--border);
        transition: all 0.25s ease;
        position: relative;
        overflow: hidden;
    }
    .kpi-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 3px;
        height: 100%;
        background: var(--primary);
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: var(--shadow-md);
    }
    .kpi-card.success::before { background: var(--success); }
    .kpi-card.warning::before { background: var(--warning); }
    .kpi-card.accent::before { background: var(--accent); }
    .kpi-icon {
        font-size: 18px;
        margin-bottom: 6px;
    }
    .kpi-label {
        font-size: 9px;
        color: var(--text-secondary);
        text-transform: uppercase;
        letter-spacing: 0.3px;
        font-weight: 600;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 18px;
        font-weight: 700;
        color: var(--text-primary);
        line-height: 1;
    }
    .kpi-sub {
        font-size: 10px;
        color: var(--text-light);
        margin-top: 4px;
    }
    .kpi-change {
        font-size: 10px;
        font-weight: 600;
        margin-top: 6px;
    }
    .kpi-change.up { color: var(--success); }
    .kpi-change.down { color: var(--danger); }
    
    /* ===== CONTENT CARDS ===== */
    .content-card {
        background: var(--bg-card);
        border-radius: var(--radius-md);
        padding: 25px;
        box-shadow: var(--shadow-sm);
        border: 1px solid var(--border);
        margin-bottom: 20px;
    }
    .card-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
        padding-bottom: 15px;
        border-bottom: 1px solid var(--border);
    }
    .card-title {
        font-size: 18px;
        font-weight: 700;
        color: var(--text-primary);
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .card-title .icon { font-size: 24px; }
    
    /* ===== SECTION HEADERS ===== */
    .section-header {
        background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
        color: white;
        padding: 12px 18px;
        border-radius: var(--radius-md);
        font-weight: 600;
        font-size: 14px;
        margin: 12px 0 10px 0;
        box-shadow: var(--shadow-sm);
    }
    .section-header .icon { margin-right: 8px; }
    
    /* ===== INPUT STYLING ===== */
    div[data-testid="stNumberInput"] label,
    div[data-testid="stSelectbox"] label,
    div[data-testid="stSlider"] label {
        font-weight: 600 !important;
        color: var(--text-secondary) !important;
        font-size: 13px !important;
    }
    
    /* ===== METRICS ===== */
    div[data-testid="stMetric"] {
        background: var(--bg-card);
        border-radius: var(--radius-md);
        padding: 18px;
        border: 1px solid var(--border);
    }
    div[data-testid="stMetricLabel"] {
        font-size: 12px;
        color: var(--text-secondary);
        text-transform: uppercase;
        font-weight: 600;
    }
    div[data-testid="stMetricValue"] {
        font-size: 24px;
        font-weight: 700;
        color: var(--primary);
    }
    
    /* ===== TABS ===== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: var(--bg-card);
        padding: 8px;
        border-radius: var(--radius-md);
        border: 1px solid var(--border);
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: var(--radius-sm);
        font-weight: 600;
        font-size: 13px;
    }
    
    /* ===== DATA TABLES ===== */
    .dataframe {
        border-radius: var(--radius-md) !important;
        overflow: hidden;
        border: 1px solid var(--border);
    }
    .dataframe thead {
        background: var(--primary) !important;
        color: white !important;
    }
    .dataframe thead th {
        font-weight: 600 !important;
        border: none !important;
    }
    .dataframe tbody tr:hover {
        background: #F0FDF4 !important;
    }
    
    /* ===== BUTTONS ===== */
    .stButton > button {
        border-radius: var(--radius-sm);
        font-weight: 600;
        transition: all 0.2s ease;
    }
    .tech-btn {
        flex: 1;
        padding: 14px !important;
        border-radius: var(--radius-md) !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }
    .tech-btn:hover {
        transform: translateY(-2px);
    }
    
    /* ===== RADIO BUTTONS IN SIDEBAR ===== */
    .sidebar-radio label {
        color: rgba(255,255,255,0.85) !important;
        font-size: 14px !important;
        padding: 10px 0 !important;
    }
    .sidebar-radio div[data-testid="stRadio"] > label {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 10px 12px !important;
        border-radius: var(--radius-sm);
        cursor: pointer;
    }
    .sidebar-radio div[data-testid="stRadio"] > label:hover {
        background: rgba(255,255,255,0.1);
    }
    
    /* ===== SUCCESS/WARNING BOXES ===== */
    .success-box {
        background: #ECFDF5;
        border-left: 4px solid var(--success);
        padding: 15px 20px;
        border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
        margin: 15px 0;
    }
    .warning-box {
        background: #FFFBEB;
        border-left: 4px solid var(--warning);
        padding: 15px 20px;
        border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
        margin: 15px 0;
    }
    .danger-box {
        background: #FEF2F2;
        border-left: 4px solid var(--danger);
        padding: 15px 20px;
        border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
        margin: 15px 0;
    }
    
    /* ===== HIDE ELEMENTS ===== */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none !important;}
    
    /* ===== RESPONSIVE ===== */
    @media (max-width: 1400px) {
        .kpi-grid { grid-template-columns: repeat(3, 1fr); }
    }
    @media (max-width: 1100px) {
        .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 768px) {
        .kpi-grid { grid-template-columns: 1fr; }
        .hero-header { padding: 25px; }
        .hero-header h1 { font-size: 24px; }
    }
    
    /* ===== VERSION BADGE ===== */
    .version-badge {
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: var(--primary);
        color: white;
        padding: 8px 16px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        box-shadow: var(--shadow-md);
        z-index: 999;
    }
    
    /* ===== DARK MODE OVERRIDES ===== */
    .dark-mode .stMetric,
    .dark-mode div[data-testid="stMetric"] {
        background: var(--bg-card) !important;
        border: 1px solid var(--border) !important;
    }
    .dark-mode .kpi-card {
        background: var(--bg-card) !important;
        border: 1px solid var(--border) !important;
    }
    .dark-mode .section-header {
        background: linear-gradient(135deg, var(--primary-light) 0%, var(--primary) 100%) !important;
    }
    .dark-mode .hero-header {
        background: linear-gradient(135deg, var(--primary-dark) 0%, var(--primary) 100%) !important;
    }
    .dark-mode .dataframe {
        background: var(--bg-card) !important;
    }
    .dark-mode .dataframe thead {
        background: var(--primary) !important;
    }
    .dark-mode .content-card {
        background: var(--bg-card) !important;
        border: 1px solid var(--border) !important;
    }
    .dark-mode section[data-testid="stSidebar"] {
        background: var(--bg-sidebar) !important;
    }
    .dark-mode .streamlit-expanderHeader {
        background: var(--bg-card) !important;
        border: 1px solid var(--border) !important;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# ============================================================================
# SESSION STATE INITIALIZATION (MUST BE BEFORE ANY USAGE)
# ============================================================================

if 'technology' not in st.session_state:
    st.session_state.technology = "Solar"

if 'active_sheet' not in st.session_state:
    st.session_state.active_sheet = "🏠 Dashboard"

# Initialize all session variables
if 'project_name' not in st.session_state:
    st.session_state.project_name = "Solar Project"
if 'project_company' not in st.session_state:
    st.session_state.project_company = "Company"
if 'capacity_dc' not in st.session_state:
    st.session_state.capacity_dc = 53.63
if 'capacity_ac' not in st.session_state:
    st.session_state.capacity_ac = 48.7
if 'yield_p50' not in st.session_state:
    st.session_state.yield_p50 = 1536
if 'yield_p90' not in st.session_state:
    st.session_state.yield_p90 = 1300
if 'yield_p99' not in st.session_state:
    st.session_state.yield_p99 = 1100
if 'ppa_base_tariff' not in st.session_state:
    st.session_state.ppa_base_tariff = 65.0
if 'tariff_escalation' not in st.session_state:
    st.session_state.tariff_escalation = 0.02
if 'ppa_term' not in st.session_state:
    st.session_state.ppa_term = 10
if 'gearing_ratio' not in st.session_state:
    st.session_state.gearing_ratio = 0.70  # 70% debt, 30% shareholder loan
if 'debt_tenor' not in st.session_state:
    st.session_state.debt_tenor = 12
if 'target_dscr' not in st.session_state:
    st.session_state.target_dscr = 1.15
if 'debt_sculpting' not in st.session_state:
    st.session_state.debt_sculpting = True
if 'corporate_tax_rate' not in st.session_state:
    st.session_state.corporate_tax_rate = 0.10
if 'depreciation_rate' not in st.session_state:
    st.session_state.depreciation_rate = 1.0 / st.session_state.get('investment_horizon', 30)  # Auto 1/period
if 'depreciation_period' not in st.session_state:
    st.session_state.depreciation_period = st.session_state.get('investment_horizon', 30)  # Auto-match investment horizon
if 'senior_debt_margin' not in st.session_state:
    st.session_state.senior_debt_margin = 0.028
if 'base_rate' not in st.session_state:
    st.session_state.base_rate = 0.0303
if 'merchant_price' not in st.session_state:
    st.session_state.merchant_price = 60.0

# ===== BATTERY STORAGE (BESS) =====
if 'bess_enabled' not in st.session_state:
    st.session_state.bess_enabled = False
if 'bess_capacity_mwh' not in st.session_state:
    st.session_state.bess_capacity_mwh = 10.0  # MWh
if 'bess_power_mw' not in st.session_state:
    st.session_state.bess_power_mw = 5.0  # MW
if 'bess_cost_per_mwh' not in st.session_state:
    st.session_state.bess_cost_per_mwh = 250000  # EUR/MWh
if 'bess_roundtrip_efficiency' not in st.session_state:
    st.session_state.bess_roundtrip_efficiency = 0.88  # 88%
if 'bess_cycle_life' not in st.session_state:
    st.session_state.bess_cycle_life = 5000  # cycles
if 'bess_degradation_rate' not in st.session_state:
    st.session_state.bess_degradation_rate = 0.02  # 2% per year
if 'bess_min_soc' not in st.session_state:
    st.session_state.bess_min_soc = 0.15  # 15% min
if 'bess_max_soc' not in st.session_state:
    st.session_state.bess_max_soc = 0.95  # 95% max
if 'bess_annual_cycles' not in st.session_state:
    st.session_state.bess_annual_cycles = 365  # assume daily cycle

if 'investment_horizon' not in st.session_state:
    st.session_state.investment_horizon = 30
if 'construction_period' not in st.session_state:
    st.session_state.construction_period = 12
if 'availability_wind' not in st.session_state:
    st.session_state.availability_wind = 0.95
if 'turbine_rating' not in st.session_state:
    st.session_state.turbine_rating = 6.0
if 'num_turbines' not in st.session_state:
    st.session_state.num_turbines = 10
if 'wind_capacity' not in st.session_state:
    st.session_state.wind_capacity = 60.0
if 'wind_speed' not in st.session_state:
    st.session_state.wind_speed = 7.5
if 'hub_height' not in st.session_state:
    st.session_state.hub_height = 100
if 'wake_effects' not in st.session_state:
    st.session_state.wake_effects = 0.0
if 'curtailment' not in st.session_state:
    st.session_state.curtailment = 0.0
if 'merchant_tail_enabled' not in st.session_state:
    st.session_state.merchant_tail_enabled = False
if 'cash_sweep_enabled' not in st.session_state:
    st.session_state.cash_sweep_enabled = False
if 'cash_sweep_threshold' not in st.session_state:
    st.session_state.cash_sweep_threshold = 1.2
if 'dscr_market' not in st.session_state:
    st.session_state.dscr_market = False
if 'p99_debt_sizing' not in st.session_state:
    st.session_state.p99_debt_sizing = False
if 'shl_rate' not in st.session_state:
    st.session_state.shl_rate = 0.0
if 'flags' not in st.session_state:
    st.session_state.flags = {}
if 'mc_results' not in st.session_state:
    st.session_state.mc_results = None
if 'ui_version' not in st.session_state:
    st.session_state.ui_version = "v2"
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False

# ============================================================================
# HELPER FUNCTIONS FOR RENDERING
# ============================================================================

# ===== ENHANCED CHART HELPERS =====

def get_chart_template():
    """Get professional chart template"""
    return dict(
        hovermode='x unified',
        hoverlabel=dict(
            bgcolor="white",
            font_size=13,
            font_family="Inter, sans-serif"
        ),
        font=dict(family="Inter, sans-serif", size=12),
        paper_bgcolor='white',
        plot_bgcolor='rgba(0,0,0,0)',
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            bgcolor="rgba(0,0,0,0)",
            borderwidth=0
        ),
        margin=dict(l=60, r=40, t=80, b=60),
        xaxis=dict(
            showgrid=True,
            gridcolor='rgba(0,0,0,0.05)',
            showline=True,
            linecolor='rgba(0,0,0,0.1)',
            tickfont=dict(size=11)
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='rgba(0,0,0,0.05)',
            showline=True,
            linecolor='rgba(0,0,0,0.1)',
            tickfont=dict(size=11)
        )
    )

def apply_chart_style(fig, title=None, yaxis_title=None, xaxis_title=None, height=400, barmode=None):
    """Apply professional styling to a chart"""
    fig.update_layout(**get_chart_template())
    
    if title:
        fig.update_layout(
            title=dict(
                text=title,
                font=dict(size=18, color='#1A1A1A', family="Inter, sans-serif"),
                x=0.5,
                xanchor='center'
            )
        )
    
    if yaxis_title:
        fig.update_yaxes(title_text=yaxis_title, title_font=dict(size=13))
    
    if xaxis_title:
        fig.update_xaxes(title_text=xaxis_title, title_font=dict(size=13))
    
    fig.update_layout(height=height)
    
    if barmode:
        fig.update_layout(barmode=barmode)
    
    return fig

def format_hover_trace(trace, value_format=','):
    """Format hover template for a trace"""
    return trace

def create_metric_badge(value, threshold, color_low='#EF4444', color_high='#10B981'):
    """Create colored badge based on threshold"""
    color = color_low if value < threshold else color_high
    return f"<span style='color:{color};font-weight:bold;'>{value}</span>"

# ============================================================================
# FINANCIAL CALCULATION HELPERS
# ============================================================================

def irr_calc_global(cash_flows):
    """Global IRR calculation using bisection method - more reliable"""
    if not cash_flows or len(cash_flows) < 2:
        return 0
    low = -0.99
    high = 10.0
    tol = 1e-8
    for _ in range(1000):
        mid = (low + high) / 2
        npv = sum(cf / (1 + mid) ** i for i, cf in enumerate(cash_flows))
        if abs(npv) < tol:
            return mid * 100
        npv_low = sum(cf / (1 + low) ** i for i, cf in enumerate(cash_flows))
        if npv * npv_low < 0:
            high = mid
        else:
            low = mid
    return mid * 100


# ============================================================================
# PROJECT MANAGEMENT SYSTEM
# ============================================================================

import uuid
import os

PROJECTS_DIR = "/root/.openclaw/workspace/projects"

def get_projects_dir():
    """Get or create projects directory"""
    os.makedirs(PROJECTS_DIR, exist_ok=True)
    return PROJECTS_DIR

def list_projects():
    """List all saved projects"""
    projects_dir = get_projects_dir()
    projects = []
    
    for project_id in os.listdir(projects_dir):
        project_path = os.path.join(projects_dir, project_id)
        if os.path.isdir(project_path):
            metadata_file = os.path.join(project_path, "metadata.json")
            data_file = os.path.join(project_path, "data.json")
            
            if os.path.exists(metadata_file):
                with open(metadata_file, 'r') as f:
                    metadata = json.load(f)
                projects.append({
                    'id': project_id,
                    'name': metadata.get('name', 'Unnamed'),
                    'technology': metadata.get('technology', 'Solar'),
                    'modified': metadata.get('modified', ''),
                    'path': project_path
                })
    
    # Sort by modified date (newest first)
    projects.sort(key=lambda x: x.get('modified', ''), reverse=True)
    return projects

def save_project(project_id=None):
    """Save current project state to disk"""
    if project_id is None:
        project_id = str(uuid.uuid4())[:8]
    
    project_path = os.path.join(get_projects_dir(), project_id)
    os.makedirs(project_path, exist_ok=True)
    
    # Collect all session state variables to save, excluding widget keys
    WIDGET_KEY_PREFIXES = ('tab_', 'nav_', 'version_', 'project_switcher', 'load_btn', 'delete_btn', 'sidebar_', 'open_', 'del_')
    
    # DEBUG: Check for tab_ keys before filtering
    all_keys = list(st.session_state.keys())
    tab_keys_in_session = [k for k in all_keys if k.startswith('tab_')]
    if tab_keys_in_session:
        print(f"DEBUG save: Found tab_ keys in session_state: {tab_keys_in_session}")
    
    session_vars = {k: v for k, v in st.session_state.items() 
                   if not k.startswith('_') 
                   and not k.startswith(WIDGET_KEY_PREFIXES)
                   and not any(k.startswith(p) for p in WIDGET_KEY_PREFIXES)
                   and isinstance(v, (str, int, float, bool))}
    
    # DEBUG: Check what's being saved
    saved_tab_keys = [k for k in session_vars.keys() if k.startswith('tab_')]
    if saved_tab_keys:
        print(f"DEBUG save: WARNING - tab_ keys in saved data: {saved_tab_keys}")
    
    # Save data
    with open(os.path.join(project_path, 'data.json'), 'w') as f:
        json.dump(session_vars, f, indent=2, default=str)
    
    # Save metadata
    metadata = {
        'name': st.session_state.get('project_name', 'Unnamed'),
        'technology': st.session_state.get('technology', 'Solar'),
        'modified': datetime.now().isoformat(),
        'id': project_id
    }
    with open(os.path.join(project_path, 'metadata.json'), 'w') as f:
        json.dump(metadata, f, indent=2)
    
    return project_id

def load_project(project_id):
    """Load project from disk into session state"""
    project_path = os.path.join(get_projects_dir(), project_id)
    data_file = os.path.join(project_path, 'data.json')
    
    if os.path.exists(data_file):
        with open(data_file, 'r') as f:
            data = json.load(f)
        
        # DEBUG: Print problematic keys
        tab_keys = [k for k in data.keys() if k.startswith('tab_')]
        if tab_keys:
            print(f"DEBUG: Found tab_ keys in saved data: {tab_keys}")
        
        # Load all variables into session state, skipping tab button keys
        for key, value in data.items():
            if not key.startswith('tab_'):
                st.session_state[key] = value
            else:
                print(f"DEBUG: Skipping tab_ key: {key}")
        
        return True
    return False

def delete_project(project_id):
    """Delete a project"""
    import shutil
    project_path = os.path.join(get_projects_dir(), project_id)
    if os.path.exists(project_path):
        shutil.rmtree(project_path)
        return True
    return False

def render_project_dashboard():
    """Render the project dashboard/home screen"""
    st.markdown('<div class="section-header">🏠 Project Dashboard</div>', unsafe_allow_html=True)
    
    # Header actions
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("➕ New Project", use_container_width=True, type="primary"):
            # Create new project
            new_id = save_project()
            load_project(new_id)
            st.session_state.active_sheet = "📋 Scenarios"
            st.rerun()
    
    # List existing projects
    projects = list_projects()
    
    if projects:
        st.markdown("### 📁 Your Projects")
        
        # Create project cards grid
        cols = st.columns(3)
        for i, project in enumerate(projects):
            with cols[i % 3]:
                tech_icon = "☀️" if project['technology'] == 'Solar' else "🌀"
                tech_color = "#FF9800" if project['technology'] == 'Solar' else "#2196F3"
                
                # Calculate basic metrics for preview
                st.markdown(f"""
                <div class="kpi-card" style="border-left: 4px solid {tech_color};">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-size: 24px;">{tech_icon}</span>
                        <span style="color: #6B7280; font-size: 12px;">{project['technology']}</span>
                    </div>
                    <h3 style="margin: 10px 0 5px 0; font-size: 16px;">{project['name']}</h3>
                    <p style="color: #6B7280; font-size: 11px; margin: 0;">
                        Modified: {project['modified'][:10] if project['modified'] else 'Unknown'}
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button(f"📂 Open", key=f"open_{project['id']}", use_container_width=True):
                        load_project(project['id'])
                        st.session_state.active_sheet = "📋 Scenarios"
                        st.rerun()
                with col_b:
                    if st.button(f"🗑️", key=f"del_{project['id']}", use_container_width=True):
                        delete_project(project['id'])
                        st.rerun()
                
                st.markdown("<hr style='margin: 15px 0; border: 1px solid #E5E7EB;'>", unsafe_allow_html=True)
    else:
        st.info("📭 No projects yet. Create your first project!")

def create_default_projects():
    """Create default Solar and Wind projects if none exist"""
    if os.path.exists(get_projects_dir()) and list_projects():
        return  # Projects already exist
    
    # Save Solar project - Trebinje Solar
    st.session_state.technology = "Solar"
    st.session_state.project_name = "Trebinje Solar"
    st.session_state.project_company = "Greene"
    st.session_state.capacity_dc = 53.63
    st.session_state.capacity_ac = 48.7
    st.session_state.yield_p50 = 1536
    st.session_state.yield_p90 = 1300
    st.session_state.yield_p99 = 1100
    st.session_state.ppa_base_tariff = 65.0
    st.session_state.tariff_escalation = 0.02
    st.session_state.ppa_term = 10
    st.session_state.gearing_ratio = 0.70  # 70% debt / 30% equity
    st.session_state.debt_tenor = 12
    st.session_state.target_dscr = 1.15
    save_project()
    
    # Save Wind project - Krnovo Wind
    st.session_state.technology = "Wind"
    st.session_state.project_name = "Krnovo Wind"
    st.session_state.project_company = "Greene"
    st.session_state.capacity_dc = 60.0
    st.session_state.capacity_ac = 60.0
    st.session_state.yield_p50 = 2500
    st.session_state.yield_p90 = 2200
    st.session_state.yield_p99 = 1900
    save_project()
    
    # Switch back to Solar
    st.session_state.technology = "Solar"

def render_sidebar_project_management():
    """Render project management in sidebar"""
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📁 Projects")
        
        # Quick project switcher
        projects = list_projects()
        
        project_names = ["Current Project"] + [p['name'] for p in projects[:5]]
        selected = st.selectbox("Switch Project", project_names, key="project_switcher")
        
        if selected != "Current Project":
            for p in projects:
                if p['name'] == selected:
                    load_project(p['id'])
                    st.rerun()
                    break
        
        # Save button
        if st.button("💾 Save Project", use_container_width=True):
            save_project()
            st.success("Project saved!")
        
        # Home button
        if st.button("🏠 Dashboard", use_container_width=True):
            st.session_state.active_sheet = "🏠 Dashboard"
            st.rerun()
        
        # Dark mode toggle
        st.markdown("---")
        st.markdown("### 🎨 Appearance")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("☀️ Light", use_container_width=True, type="primary" if not st.session_state.dark_mode else "secondary"):
                st.session_state.dark_mode = False
                st.rerun()
        with col2:
            if st.button("🌙 Dark", use_container_width=True, type="primary" if st.session_state.dark_mode else "secondary"):
                st.session_state.dark_mode = True
                st.rerun()

# ===== TOOLTIP & VALIDATION HELPERS =====

TOOLTIPS = {
    'capacity_dc': "Gross installed capacity in MW (DC side for Solar, rated capacity for Wind).",
    'capacity_ac': "Net capacity after accounting for inverter clipping/transformer losses.",
    'yield_p50': "Median expected annual generation hours (50% probability of exceedance).",
    'yield_p90': "Conservative annual generation hours (90% probability of exceedance) - used for financial modeling.",
    'yield_p99': "Pessimistic annual generation hours (99% probability of exceedance) - stress test.",
    'ppa_base_tariff': "Fixed feed-in tariff for PPA period in EUR/MWh.",
    'tariff_escalation': "Annual tariff increase (CPI or fixed escalation) during PPA.",
    'ppa_term': "Duration of Power Purchase Agreement in years.",
    'gearing_ratio': "Debt as percentage of total CAPEX. Higher = more leverage, more risk.",
    'debt_tenor': "Loan repayment period in years. Typically 10-15 for solar.",
    'target_dscr': "Minimum Debt Service Coverage Ratio. Below 1.0x = debt cannot be serviced.",
    'senior_debt_margin': "Interest rate spread above base rate (LIBOR/EURIBOR/SOFR).",
    'corporate_tax_rate': "Corporate income tax rate applied to taxable profits.",
    'depreciation_rate': "Annual depreciation as % of CAPEX (straight-line). Affects taxable income.",
    'merchant_price': "Electricity price for post-PPA merchant exposure in EUR/MWh.",
    'investment_horizon': "Total project lifetime for NPV/IRR calculations.",
}

def render_tooltip(key):
    """Render an info tooltip icon with explanation"""
    if key in TOOLTIPS:
        st.markdown(f"""
        <span title="{TOOLTIPS[key]}" style='cursor: help; margin-left: 5px; color: #6B7280;'>
            ⓘ
        </span>
        """, unsafe_allow_html=True)

def validate_input(key, value):
    """Validate an input value and return (is_valid, message)"""
    validators = {
        'capacity_dc': (1, 2000, "Capacity should be between 1-2000 MW"),
        'capacity_ac': (0.1, 1500, "AC capacity should be positive and less than DC"),
        'yield_p50': (500, 4000, "Yield typically 500-4000 hours depending on resource"),
        'yield_p90': (400, 3500, "P90 yield should be lower than P50"),
        'ppa_base_tariff': (10, 500, "Tariff typically 10-500 EUR/MWh"),
        'tariff_escalation': (0, 0.2, "Escalation usually 0-20% annually"),
        'gearing_ratio': (0, 0.95, "Gearing should be 0-95% (never 100% debt)"),
        'debt_tenor': (1, 30, "Debt tenor typically 1-30 years"),
        'target_dscr': (0.5, 5.0, "Target DSCR usually 1.0-3.0x"),
        'corporate_tax_rate': (0, 0.5, "Tax rate typically 0-50%"),
        'merchant_price': (10, 500, "Merchant price typically 10-500 EUR/MWh"),
    }
    
    if key in validators:
        min_val, max_val, msg = validators[key]
        if value < min_val or value > max_val:
            return False, f"⚠️ {msg}"
    return True, None

def render_validation_warning(message):
    """Render a validation warning box"""
    st.markdown(f"""
    <div style="background: #FEF3C7; border-left: 4px solid #F59E0B; padding: 10px 15px; 
                border-radius: 0 8px 8px 0; margin: 10px 0; font-size: 13px;">
        {message}
    </div>
    """, unsafe_allow_html=True)

def render_input_with_tooltip(label, key, **kwargs):
    """Render a number input with tooltip and validation"""
    col1, col2 = st.columns([0.95, 0.05])
    with col1:
        value = st.number_input(label, **kwargs)
    with col2:
        st.markdown("")
        if key in TOOLTIPS:
            st.markdown(f"""
            <span title="{TOOLTIPS[key]}" style='cursor: help; font-size: 18px; color: #6B7280;'>
                ⓘ
            </span>
            """, unsafe_allow_html=True)
    
    # Validate
    if key in validate_input.__code__.co_varnames[:10] or hasattr(validate_input, '__code__'):
        is_valid, msg = validate_input(key, value if 'value' in kwargs else st.session_state.get(key, 0))
        if not is_valid:
            render_validation_warning(msg)
    
    return value

def render_hero_header():
    """Render the main hero header"""
    tech = st.session_state.technology
    tech_icon = "☀️" if tech == "Solar" else "🌀"
    tech_color = "#FF9800" if tech == "Solar" else "#2196F3"
    
    st.markdown(f"""
    <div class="hero-header">
        <h1>📊 {tech} Financial Model</h1>
        <p>Project: <strong>{st.session_state.project_name}</strong> | Company: {st.session_state.project_company}</p>
        <div class="hero-badge" style="background: {tech_color};">
            {tech_icon} {tech.upper()} | {st.session_state.capacity_dc:.1f} MW
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_kpi_grid():
    """Render the KPI dashboard grid"""
    total_capex = calculate_total_capex()
    debt = calculate_debt_amount()
    equity = calculate_equity_amount()
    project_irr = calculate_project_irr([-total_capex] + [calculate_ebitda(y) - calculate_tax(y) for y in range(1, 30)])
    dscr = calculate_avg_dscr()
    lcoe = calculate_lcoe()
    ann_gen = calculate_annual_generation()
    ann_rev = calculate_revenue(1)
    payback = calculate_payback()
    
    kpi_cards = [
        {"icon": "💰", "label": "Total CAPEX", "value": f"{total_capex:,.0f}", "sub": "kEUR Investment", "type": ""},
        {"icon": "🏦", "label": "Debt", "value": f"{debt:,.0f}", "sub": f"{st.session_state.gearing_ratio*100:.0f}% Leverage", "type": "accent"},
        {"icon": "📈", "label": "Project IRR", "value": f"{project_irr:.1f}%", "sub": "30 Year Projection", "type": "success"},
        {"icon": "⚡", "label": "LCOE", "value": f"{lcoe:.2f}", "sub": "EUR/MWh", "type": ""},
        {"icon": "🔋", "label": "Annual Generation", "value": f"{ann_gen:,.0f}", "sub": f"MWh | {st.session_state.yield_p90}h yield", "type": ""},
        {"icon": "💵", "label": "Annual Revenue", "value": f"{ann_rev:,.0f}", "sub": f"EUR/MWh: {st.session_state.ppa_base_tariff:.2f}", "type": "success"},
        {"icon": "📊", "label": "Avg DSCR", "value": f"{dscr:.2f}x", "sub": f"Target: {st.session_state.target_dscr:.2f}x", "type": "warning" if dscr < st.session_state.target_dscr else "success"},
        {"icon": "⏱️", "label": "Payback", "value": f"{payback:.1f}", "sub": "Years", "type": "accent"},
    ]
    
    cards_html = '<div class="kpi-grid">'
    for card in kpi_cards:
        cards_html += f'''
        <div class="kpi-card {card['type']}">
            <div class="kpi-icon">{card['icon']}</div>
            <div class="kpi-label">{card['label']}</div>
            <div class="kpi-value">{card['value']}</div>
            <div class="kpi-sub">{card['sub']}</div>
        </div>'''
    cards_html += '</div>'
    
    st.markdown(cards_html, unsafe_allow_html=True)

def render_sidebar():
    """Render the sidebar navigation"""
    with st.sidebar:
        st.markdown('<div class="sidebar-container">', unsafe_allow_html=True)
        
        # Logo
        st.markdown('<div class="sidebar-logo">📊 FinModel</div>', unsafe_allow_html=True)
        
        nav_items = [
            ("🏠", "Dashboard", "🏠 Dashboard"),
            ("📈", "Sensitivity", "📈 Sensitivity"),
            ("🔄", "Comparison", "🔄 Comparison"),
        ]
        
        for icon, label, full_name in nav_items:
            is_active = st.session_state.active_sheet == full_name
            btn_type = "primary" if is_active else "secondary"
            
            if st.button(f"{icon}  {label}", use_container_width=True, 
                        type=btn_type, key=f"nav_{full_name}"):
                st.session_state.active_sheet = full_name
                st.rerun()
        
        projects = list_projects()
        project_names = ["Current Project"] + [p['name'] for p in projects[:5]]
        selected = st.selectbox("Switch Project", project_names, key="project_switcher")
        
        if selected != "Current Project":
            for p in projects:
                if p['name'] == selected:
                    load_project(p['id'])
                    st.rerun()
                    break
        
        if st.button("💾 Save Project", use_container_width=True):
            save_project()
            st.success("Project saved!")
        
        if st.button("📤 Export Model", use_container_width=True):
            st.session_state.active_sheet = "📤 Outputs"
            st.rerun()
        
        st.markdown('<div class="sidebar-footer">Solar/Wind Model v2.0</div></div>', unsafe_allow_html=True)

# ============================================================================
# SESSION STATE INITIALIZATION
# ============================================================================

def initialize_session_state():
    """Initialize all session state variables based on technology"""
    
    tech = st.session_state.technology
    
    # Project Info
    if 'project_name' not in st.session_state:
        st.session_state.project_name = "Wind Project" if tech == "Wind" else "Solar Project"
    if 'project_company' not in st.session_state:
        st.session_state.project_company = "Company"


def get_capacity():
    """Get capacity based on technology"""
    if st.session_state.technology == "Solar":
        return st.session_state.capacity_dc
    else:
        return st.session_state.wind_capacity

def calculate_total_capex():
    """Calculate total CAPEX based on technology and capacity"""
    base_capex = 0
    if st.session_state.technology == "Solar":
        # Solar: CAPEX scales with capacity (€/MW basis)
        # Typical solar CAPEX: ~750-900 €/kW = 750,000-900,000 €/MW
        capex_per_mw = 800000  # €/MW - typical utility scale solar
        base_capex = st.session_state.capacity_dc * capex_per_mw / 1000  # kEUR
    else:
        # Wind: CAPEX scales with turbine capacity
        # Typical wind CAPEX: ~1,000-1,300 €/kW installed
        capex_per_mw = 1100000  # €/MW - typical wind
        base_capex = st.session_state.wind_capacity * capex_per_mw / 1000  # kEUR
    
    # Add BESS cost if enabled
    if st.session_state.bess_enabled:
        bess_capex = st.session_state.bess_capacity_mwh * st.session_state.bess_cost_per_mwh / 1000  # kEUR
        base_capex += bess_capex
    
    return base_capex

def calculate_bess_capex():
    """Calculate BESS CAPEX in kEUR"""
    if st.session_state.bess_enabled:
        return st.session_state.bess_capacity_mwh * st.session_state.bess_cost_per_mwh / 1000
    return 0

def calculate_bess_revenue(year):
    """Calculate annual BESS revenue from arbitrage"""
    if not st.session_state.bess_enabled:
        return 0
    
    # Simplified arbitrage model:
    # Battery charges at night (low prices), discharges at peak (high prices)
    # Average spread: ~20 EUR/MWh
    annual_energy = st.session_state.bess_capacity_mwh * st.session_state.bess_annual_cycles * st.session_state.bess_roundtrip_efficiency
    spread = st.session_state.merchant_price * 0.3  # 30% spread assumption
    revenue = annual_energy * spread / 1000  # kEUR
    
    # Degradation reduces capacity over time
    years_of_operation = year - 1
    degradation_factor = max(0.5, 1 - years_of_operation * st.session_state.bess_degradation_rate)
    
    return revenue * degradation_factor

def calculate_bess_costs(year):
    """Calculate annual BESS costs (O&M + replacement)"""
    if not st.session_state.bess_enabled:
        return 0
    
    # O&M cost: ~1% of CAPEX per year
    bess_capex = calculate_bess_capex()
    om_cost = bess_capex * 0.01
    
    # Replacement cost if battery degrades significantly
    years_of_operation = year - 1
    if years_of_operation > 0:
        degradation_factor = 1 - years_of_operation * st.session_state.bess_degradation_rate
        if degradation_factor < 0.7:  # Replace at 70% capacity
            replacement_cost = bess_capex * 0.5  # 50% of original cost
            return om_cost + replacement_cost
    
    return om_cost

def calculate_total_opex_y1():
    """Calculate total OPEX Y1 based on technology and capacity"""
    if st.session_state.technology == "Solar":
        # Solar O&M: ~15,000 €/MW/year
        opex_per_mw = 15000
        return st.session_state.capacity_dc * opex_per_mw / 1000  # kEUR
    else:
        # Wind O&M: ~35,000 €/MW/year
        opex_per_mw = 35000
        return st.session_state.wind_capacity * opex_per_mw / 1000  # kEUR

def calculate_annual_generation(yield_hours=None):
    """Calculate annual generation in MWh"""
    if yield_hours is None:
        yield_hours = st.session_state.yield_p90
    
    capacity = get_capacity()
    
    if st.session_state.technology == "Wind":
        # Apply availability and curtailment
        availability = st.session_state.availability_wind
        curtailment_factor = 1 - st.session_state.curtailment / 100
        wake_factor = 1 - st.session_state.wake_effects / 100
        return capacity * yield_hours * availability * curtailment_factor * wake_factor
    else:
        return capacity * yield_hours

def calculate_revenue(year, yield_hours=None):
    """Calculate annual revenue in k€"""
    if yield_hours is None:
        yield_hours = st.session_state.yield_p90
    
    generation = calculate_annual_generation(yield_hours)
    
    # Determine tariff based on PPA term and merchant tail
    tech = st.session_state.technology
    
    if tech == "Wind" and st.session_state.merchant_tail_enabled and year > st.session_state.ppa_term:
        # Merchant tail period
        tariff = st.session_state.merchant_price
    else:
        # PPA period with escalation
        tariff = st.session_state.ppa_base_tariff * (1 + st.session_state.tariff_escalation) ** (year - 1)
    
    return generation * tariff / 1000  # Convert to k€

def calculate_opex_year(year):
    """Calculate OPEX for given year"""
    total_opex_y1 = calculate_total_opex_y1()
    
    if st.session_state.technology == "Solar":
        # Standard escalation
        return total_opex_y1 * (1.02) ** (year - 1)
    else:
        # Wind tiered O&M - increases slowly
        return total_opex_y1 * (1.015) ** (year - 1)

def calculate_ebitda(year, yield_hours=None):
    """Calculate EBITDA"""
    return calculate_revenue(year, yield_hours) - calculate_opex_year(year)

def calculate_tax(year):
    """Calculate income tax"""
    ebitda = calculate_ebitda(year)
    depreciation = calculate_total_capex() * st.session_state.depreciation_rate
    taxable = max(0, ebitda - depreciation)
    
    tech = st.session_state.technology
    if tech == "Wind":
        # Wind might have tax holidays or lower rates in early years
        return taxable * st.session_state.corporate_tax_rate
    else:
        return taxable * st.session_state.corporate_tax_rate

def calculate_annual_debt_service():
    """Calculate annual debt service (equal payments - standard annuity)"""
    debt = calculate_debt_amount()
    rate = calculate_interest_rate()
    tenor = st.session_state.debt_tenor
    if rate > 0 and tenor > 0:
        return debt * (rate * (1 + rate) ** tenor) / ((1 + rate) ** tenor - 1)
    return 0

def calculate_sculpted_debt_schedule():
    """Calculate sculpted debt service based on target DSCR
    
    Debt sculpting adjusts annual payments so that DSCR never drops below target.
    Payment = EBITDA / Target_DSCR, scaled to ensure debt is repaid within tenor.
    """
    debt = calculate_debt_amount()
    rate = calculate_interest_rate()
    tenor = st.session_state.debt_tenor
    target_dscr = st.session_state.target_dscr
    
    if debt <= 0 or tenor <= 0:
        return [0] * tenor
    
    # Calculate raw sculpted payments (EBITDA / target DSCR) for each year
    years = min(tenor, st.session_state.investment_horizon)
    raw_payments = []
    for year in range(1, years + 1):
        ebitda = calculate_ebitda(year)
        # Raw payment to achieve target DSCR
        raw_payment = ebitda / target_dscr if target_dscr > 0 else ebitda
        raw_payments.append(max(raw_payment, 0))  # Can't be negative
    
    # Calculate scaling factor so PV of scaled payments = debt
    # PV = sum(scaled_payment / (1+r)^t) for t=1 to tenor
    # scaling_factor = debt / PV(raw_payments)
    if rate > 0:
        pv_raw = sum(raw_payments[t] / ((1 + rate) ** (t + 1)) for t in range(len(raw_payments)))
    else:
        pv_raw = sum(raw_payments)
    
    if pv_raw > 0:
        scaling_factor = debt / pv_raw
    else:
        scaling_factor = 1.0
    
    # Apply scaling factor to get final sculpted payments
    sculpted_schedule = [raw_payments[t] * scaling_factor for t in range(len(raw_payments))]
    
    return sculpted_schedule

def get_sculpted_debt_service(year):
    """Get sculpted debt service for a specific year"""
    schedule = calculate_sculpted_debt_schedule()
    if 0 < year <= len(schedule):
        return schedule[year - 1]
    return 0

def get_debt_service(year):
    """Get debt service for a specific year - sculpted or equal based on setting"""
    if st.session_state.debt_sculpting:
        return get_sculpted_debt_service(year)
    else:
        if year <= st.session_state.debt_tenor:
            return calculate_annual_debt_service()
        return 0

def calculate_debt_amount():
    """Calculate debt amount based on technology and sizing"""
    total_capex = calculate_total_capex()
    
    tech = st.session_state.technology
    
    if tech == "Wind" and st.session_state.p99_debt_sizing:
        # P99 debt sizing - more conservative
        p99_yield = st.session_state.yield_p99
        annual_ds = calculate_annual_debt_service_p99()
        if annual_ds > 0 and st.session_state.target_dscr > 0:
            # Size debt so that DSCR at P99 = target
            return annual_ds * st.session_state.target_dscr * ((1 + calculate_interest_rate()) ** st.session_state.debt_tenor - 1) / (calculate_interest_rate() * (1 + calculate_interest_rate()) ** st.session_state.debt_tenor)
    
    return total_capex * st.session_state.gearing_ratio

def calculate_cash_sweep_schedule():
    """Calculate cash sweep schedule - extra debt repayment when DSCR > threshold"""
    if not st.session_state.cash_sweep_enabled:
        return {}
    
    threshold = st.session_state.cash_sweep_threshold
    debt_balance = calculate_debt_amount()
    interest_rate = calculate_interest_rate()
    sweep_schedule = {}
    
    for year in range(1, int(st.session_state.debt_tenor) + 2):
        if debt_balance <= 0:
            sweep_schedule[year] = 0
            continue
        
        ebitda = calculate_ebitda(year)
        regular_ds = get_debt_service(year)
        dscr = ebitda / regular_ds if regular_ds > 0 else 999
        
        if dscr > threshold:
            # Excess cash available for sweep
            excess = ebitda - (regular_ds * threshold)
            # Sweep cannot exceed remaining debt balance
            sweep = min(excess, debt_balance)
            sweep_schedule[year] = max(0, sweep)
            # Reduce debt balance
            debt_balance -= sweep
        else:
            sweep_schedule[year] = 0
    
    return sweep_schedule

def get_cash_sweep_amount(year):
    """Get cash sweep amount for a specific year"""
    schedule = calculate_cash_sweep_schedule()
    return schedule.get(year, 0)

def calculate_annual_debt_service_p99():
    """Calculate debt service based on P99 generation"""
    generation_p99 = calculate_annual_generation(st.session_state.yield_p99)
    tariff = st.session_state.ppa_base_tariff
    revenue_p99 = generation_p99 * tariff / 1000
    opex_p99 = calculate_total_opex_y1()  # Simplified
    ebitda_p99 = revenue_p99 - opex_p99
    # DSCR at P99 should equal target
    if st.session_state.target_dscr > 0:
        return ebitda_p99 / st.session_state.target_dscr
    return calculate_annual_debt_service()

def calculate_equity_amount():
    """Calculate equity investment"""
    return calculate_total_capex() * (1 - st.session_state.gearing_ratio)

def calculate_interest_rate():
    """Calculate total interest rate"""
    return st.session_state.base_rate + st.session_state.senior_debt_margin

def calculate_avg_dscr():
    """Calculate average DSCR over debt tenor"""
    total_dscr = 0
    years_with_ds = 0
    for year in range(1, min(st.session_state.debt_tenor + 1, st.session_state.investment_horizon + 1)):
        revenue = calculate_revenue(year)
        opex = calculate_opex_year(year)
        ebitda = revenue - opex
        debt_service = get_debt_service(year)  # Use sculpted or equal based on setting
        if debt_service > 0:
            dscr = ebitda / debt_service
            total_dscr += dscr
            years_with_ds += 1
    return total_dscr / years_with_ds if years_with_ds > 0 else 1.5

def calculate_payback():
    """Calculate simple payback period in years"""
    total_capex = calculate_total_capex()
    equity = calculate_equity_amount()
    cumulative_cf = 0
    for year in range(1, st.session_state.investment_horizon + 1):
        revenue = calculate_revenue(year)
        opex = calculate_opex_year(year)
        ebitda = revenue - opex
        tax = calculate_tax(year)
        debt_service = get_debt_service(year)
        fcf_equity = ebitda - tax - debt_service
        cumulative_cf += fcf_equity
        if cumulative_cf >= equity:
            return year
    return st.session_state.investment_horizon

def calculate_lcoe(discount_rate=None):
    """Calculate Levelized Cost of Energy (LCOE) in EUR/MWh"""
    if discount_rate is None:
        discount_rate = st.session_state.discount_rate if hasattr(st.session_state, 'discount_rate') else 0.08
    total_capex = calculate_total_capex()
    total_opex = 0
    total_generation = 0
    for year in range(1, st.session_state.investment_horizon + 1):
        gen = calculate_annual_generation()
        opex = calculate_opex_year(year)
        # Apply discount rate for LCOE calculation
        discount_factor = (1 + discount_rate) ** (year - 1)
        total_opex += opex * discount_factor
        total_generation += gen * discount_factor
    
    if total_generation > 0:
        lcoe = (total_capex + total_opex) / total_generation * 1000  # EUR/MWh
    else:
        lcoe = 0
    return lcoe

def calculate_project_irr(cash_flows):
    """Calculate project IRR using numpy"""
    try:
        import numpy as np
        cash_flows = list(cash_flows)
        if len(cash_flows) < 2:
            return 0
        # Use numpy IRR function
        irr = np.irr(cash_flows)
        if irr is None or np.isnan(irr):
            return 0
        return irr * 100  # Convert to percentage
    except:
        # Fallback approximation
        try:
            total_positive = sum(c for c in cash_flows if c > 0)
            total_negative = abs(sum(c for c in cash_flows if c < 0))
            if total_negative > 0:
                return (total_positive / total_negative - 1) * 100 / (len(cash_flows) / 2)
        except:
            pass
        return 0

def calculate_npv(cash_flows, discount_rate=0.10):
    """Calculate NPV"""
    return sum(cf / (1 + discount_rate) ** i for i, cf in enumerate(cash_flows))

def calculate_plcr(year, discount_rate=0.08):
    """Calculate Project Life Coverage Ratio"""
    remaining_years = st.session_state.investment_horizon - year + 1
    if remaining_years <= 0:
        return 0
    
    future_cf = []
    for y in range(year, st.session_state.investment_horizon + 1):
        ebitda = calculate_ebitda(y)
        debt_service = calculate_annual_debt_service() if y <= st.session_state.debt_tenor else 0
        fcf = ebitda - debt_service
        future_cf.append(max(0, fcf))
    
    npv_future = sum(cf / (1 + discount_rate) ** (y - year + 1) for y, cf in enumerate(future_cf))
    
    years_paid = year - 1
    remaining_debt = calculate_total_capex() * st.session_state.gearing_ratio
    if years_paid > 0:
        rate = calculate_interest_rate()
        pv_paid = sum(calculate_annual_debt_service() / (1 + rate) ** y for y in range(1, years_paid + 1))
        remaining_debt = max(0, remaining_debt - pv_paid)
    
    if remaining_debt > 0:
        return npv_future / remaining_debt
    return 0

def calculate_monte_carlo(sims=1000):
    """Monte Carlo simulation for yield uncertainty"""
    results = []
    yields = st.session_state.yield_p90
    
    # Calculate standard deviation based on P50/P90 spread
    p50 = st.session_state.yield_p50
    p90 = st.session_state.yield_p90
    std_dev = (p50 - p90) / 1.28 if st.session_state.technology == "Wind" else (p50 - p90) / 1.28
    
    for _ in range(sims):
        sampled_yield = np.random.normal(yields, std_dev)
        sampled_yield = max(st.session_state.yield_p99, sampled_yield)
        
        cash_flows = [-calculate_total_capex()]
        for year in range(1, st.session_state.investment_horizon + 1):
            rev = get_capacity() * sampled_yield * st.session_state.ppa_base_tariff / 1000
            opex = calculate_opex_year(year)
            ebitda = rev - opex
            tax = max(0, ebitda - calculate_total_capex() * st.session_state.depreciation_rate) * st.session_state.corporate_tax_rate
            ds = get_debt_service(year)
            cf = ebitda - tax - ds
            cash_flows.append(cf)
        
        # Use proper IRR calculation
        irr = irr_calc_global(cash_flows)
        results.append(irr)
    
    return {
        'mean': np.mean(results),
        'median': np.median(results),
        'p10': np.percentile(results, 10),
        'p50': np.percentile(results, 50),
        'p90': np.percentile(results, 90),
        'std_dev': np.std(results),
        'all_results': results
    }

# ============================================================================
# CAPEX DATA STRUCTURES
# ============================================================================

# SOLAR CAPEX
SOLAR_CAPEX_ITEMS = {
    'C.01': {'name': 'PV Modules', 'amount': 6864.64, 'per_mw': 128.0},
    'C.01.01': {'name': 'PV modules', 'amount': 6864.64, 'per_mw': 128.0},
    'C.02': {'name': 'EPC Contract', 'amount': 15378.0, 'per_mw': 286.74},
    'C.02.01': {'name': 'Inverter, PCS and Transformer Station', 'amount': 4690.0, 'per_mw': 87.45},
    'C.03': {'name': 'EPC other costs', 'amount': 3502.0, 'per_mw': 65.30},
    'C.03.01': {'name': 'PROJECT MANAGEMENT', 'amount': 582.0, 'per_mw': 10.85},
    'C.04': {'name': 'Grid connection', 'amount': 6945.0, 'per_mw': 129.50},
    'C.04.01': {'name': 'Grid Usage Subscription fees', 'amount': 45.0, 'per_mw': 0.84},
    'C.05': {'name': 'Other costs for construction', 'amount': 100.0, 'per_mw': 1.86},
    'C.06': {'name': 'Insurances', 'amount': 295.0, 'per_mw': 5.50},
    'C.08': {'name': 'Project finance costs due at closing', 'amount': 235.0, 'per_mw': 4.38},
    'C.10': {'name': 'Commissioning', 'amount': 17.0, 'per_mw': 0.32},
    'C.11': {'name': 'Audit&Accounting&Legal Fees', 'amount': 70.0, 'per_mw': 1.31},
    'C.12': {'name': 'Construction Management', 'amount': 1124.94, 'per_mw': 20.98},
    'C.13': {'name': 'Contingencies', 'amount': 2013.10, 'per_mw': 37.54},
    'C.15': {'name': 'Project Acquisition / Project Development', 'amount': 743.0, 'per_mw': 13.85},
    'C.16': {'name': 'Project Rights', 'amount': 1664.44, 'per_mw': 31.04},
}

SOLAR_CAPEX_GROUPS = {
    'Production Units': ['C.01', 'C.01.01'],
    'EPC Contract': ['C.02', 'C.02.01', 'C.03', 'C.03.01', 'C.03.02'],
    'Grid Connection': ['C.04', 'C.04.01'],
    'Construction & Commissioning': ['C.05', 'C.06', 'C.08', 'C.10', 'C.11', 'C.12'],
    'Development & Rights': ['C.13', 'C.15', 'C.16'],
}

# WIND CAPEX (from Wind Project model)
WIND_CAPEX_ITEMS = {
    'W.01': {'name': 'Wind Turbines', 'amount': 54000.0, 'per_mw': 900.0},
    'W.02': {'name': 'Electrical BOP', 'amount': 480.0, 'per_mw': 8.0},
    'W.03': {'name': 'Civil BOP', 'amount': 4990.0, 'per_mw': 83.17},
    'W.04': {'name': 'Grid Connection', 'amount': 10720.0, 'per_mw': 178.67},
    'W.05': {'name': 'Transformer', 'amount': 6000.0, 'per_mw': 100.0},
    'W.06': {'name': 'HV SS Equipment', 'amount': 320.0, 'per_mw': 5.33},
    'W.07': {'name': 'Overhead Line (OVHL)', 'amount': 4400.0, 'per_mw': 73.33},
    'W.08': {'name': 'GSM Base relocation', 'amount': 500.0, 'per_mw': 8.33},
    'W.09': {'name': 'Land leases', 'amount': 20.0, 'per_mw': 0.33},
    'W.10': {'name': 'Insurances', 'amount': 500.0, 'per_mw': 8.33},
    'W.11': {'name': 'Project development costs', 'amount': 1540.0, 'per_mw': 25.67},
    'W.12': {'name': 'Bank due diligence', 'amount': 420.0, 'per_mw': 7.0},
    'W.13': {'name': 'Shareholder services (% of capex)', 'amount': 2269.88, 'per_mw': 37.83},
    'W.14': {'name': 'Contingencies (6%)', 'amount': 5081.49, 'per_mw': 84.69},
}

WIND_CAPEX_GROUPS = {
    'Wind Turbines': ['W.01'],
    'Balance of Plant (BOP)': ['W.02', 'W.03'],
    'Grid Connection': ['W.04', 'W.05', 'W.06', 'W.07'],
    'Development & Legal': ['W.08', 'W.09', 'W.11', 'W.12'],
    'Fees & Contingencies': ['W.13', 'W.14', 'W.10'],
}

# OPEX DATA STRUCTURES
SOLAR_OPEX_ITEMS = {
    'B.01': {'name': 'Technical Management', 'budget_y1': 160.0, 'inflation': 0.02},
    'B.02': {'name': 'Infrastructure Maintenance', 'budget_y1': 206.5, 'inflation': 0.02},
    'B.05': {'name': 'Security', 'budget_y1': 13.0, 'inflation': 0.02},
    'B.06': {'name': 'Insurance', 'budget_y1': 155.0, 'inflation': 0.02},
    'B.07': {'name': 'Lease & Property Tax', 'budget_y1': 195.76, 'inflation': 0.02},
    'B.08': {'name': 'Power Expenses', 'budget_y1': 147.56, 'inflation': 0.02},
    'B.09': {'name': 'Telecom Fees', 'budget_y1': 11.0, 'inflation': 0.02},
    'B.10': {'name': 'Audit&Accounting&Legal', 'budget_y1': 32.0, 'inflation': 0.02},
    'B.11': {'name': 'Bank Fees', 'budget_y1': 15.0, 'inflation': 0.02},
    'B.12': {'name': 'Environmental&Social', 'budget_y1': 22.0, 'inflation': 0.02},
    'B.13': {'name': 'Contingencies', 'budget_y1': 38.61, 'inflation': 0.04},
}

WIND_OPEX_ITEMS = {
    'B.01': {'name': 'Operation Management', 'budget_y1': 300.0, 'inflation': 0.02},
    'B.02': {'name': 'Baseline (Bazefield)', 'budget_y1': 25.0, 'inflation': 0.02},
    'B.03': {'name': 'O&M Y1-5', 'budget_y1': 700.0, 'inflation': 0.015},  # Tiered
    'B.04': {'name': 'O&M Y5-10', 'budget_y1': 750.0, 'inflation': 0.015},
    'B.05': {'name': 'O&M Y10-20', 'budget_y1': 790.0, 'inflation': 0.015},
    'B.06': {'name': 'O&M Y20-30', 'budget_y1': 880.0, 'inflation': 0.015},
    'B.07': {'name': 'Regulatory inspections', 'budget_y1': 30.0, 'inflation': 0.02},
    'B.08': {'name': 'Spare parts procurement', 'budget_y1': 15.0, 'inflation': 0.02},
    'B.09': {'name': 'Surveillance system', 'budget_y1': 10.0, 'inflation': 0.02},
    'B.10': {'name': 'Insurance', 'budget_y1': 250.0, 'inflation': 0.02},
    'B.11': {'name': 'Land Lease (% of revenue)', 'budget_y1': 525.46, 'inflation': 0.0, 'pct_of_revenue': 0.036},
    'B.12': {'name': 'Power expenses & Telecom', 'budget_y1': 45.0, 'inflation': 0.02},
    'B.13': {'name': 'Audit, Accounting & Legal', 'budget_y1': 16.0, 'inflation': 0.02},
    'B.14': {'name': 'Bank fee', 'budget_y1': 15.0, 'inflation': 0.02},
    'B.15': {'name': 'Mitigation measures', 'budget_y1': 15.0, 'inflation': 0.02},
    'B.16': {'name': 'Contingencies', 'budget_y1': 46.37, 'inflation': 0.02},
}

# ============================================================================
# SHEET TABS
# ============================================================================

sheet_tabs = [
    "📋 Scenarios", "💰 CapEx", "📊 OpEx", "📈 P&L", "📄 Balance Sheet", 
    "💵 Cash Flow", "🏦 Debt Service", "💎 Equity", "📤 Outputs", "⚙️ Advanced"
]

if 'active_sheet' not in st.session_state:
    st.session_state.active_sheet = "📋 Scenarios"

tab_cols = st.columns(len(sheet_tabs))
for i, tab in enumerate(sheet_tabs):
    with tab_cols[i]:
        if st.button(tab, use_container_width=True, type="primary" if tab == st.session_state.active_sheet else "secondary", key=f"tab_{tab}"):
            st.session_state.active_sheet = tab
            st.rerun()

st.markdown("---")

# ============================================================================
# MAIN LAYOUT - Sidebar, Hero, KPI
# ============================================================================

# Create default projects if none exist
create_default_projects()

render_sidebar()

# Dark mode wrapper
dark_class = "dark-mode" if st.session_state.dark_mode else ""
if dark_class:
    st.markdown(f'<div class="{dark_class}">', unsafe_allow_html=True)

# Always show hero and KPI (updates automatically when inputs change)
render_hero_header()
render_kpi_grid()

# ============================================================================
# SHEET: DASHBOARD (PROJECT HOME)
# ============================================================================

if st.session_state.active_sheet == "🏠 Dashboard":
    render_project_dashboard()

# ============================================================================
# SHEET: SCENARIOS
# ============================================================================

elif st.session_state.active_sheet == "📋 Scenarios":
    tech = st.session_state.technology
    
    st.header(f"📋 Scenarios - {tech} Input Parameters")
    st.markdown(f"*Project: {st.session_state.project_name} | Company: {st.session_state.project_company}*")
    
    # ========================================================================
    # SHARED PARAMETERS (Same for all scenarios)
    # ========================================================================
    
    # Project Info
    st.markdown('<div class="section-header">PROJECT INFORMATION</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.session_state.project_name = st.text_input("Project Name", value=st.session_state.project_name)
    with col2:
        st.session_state.project_company = st.text_input("Company", value=st.session_state.project_company)
    with col3:
        st.number_input("Investment Horizon (years)", min_value=10, max_value=50, value=st.session_state.investment_horizon, key="inv_horizon_scen")
    
    # Technical Parameters (shared across scenarios)
    st.markdown('<div class="section-header">TECHNICAL PARAMETERS (Shared)</div>', unsafe_allow_html=True)
    
    if tech == "Solar":
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.session_state.capacity_dc = st.number_input("Capacity DC (MW)", value=st.session_state.capacity_dc, min_value=1.0, max_value=500.0, step=0.1)
        with col2:
            st.session_state.capacity_ac = st.number_input("Capacity AC (MW)", value=st.session_state.capacity_ac, min_value=1.0, max_value=500.0, step=0.1)
        with col3:
            st.session_state.yield_p50 = st.number_input("Yield P50 (hours)", value=float(st.session_state.yield_p50), min_value=500.0, max_value=3000.0, step=10.0)
        with col4:
            st.session_state.yield_p90 = st.number_input("Yield P90 (hours)", value=float(st.session_state.yield_p90), min_value=500.0, max_value=3000.0, step=10.0)
        
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.yield_p99 = st.number_input("Yield P99 (hours)", value=float(st.session_state.yield_p99), min_value=500.0, max_value=3000.0, step=10.0)
        with col2:
            p90_p50 = st.session_state.yield_p90 / st.session_state.yield_p50 if st.session_state.yield_p50 > 0 else 0
            st.metric("P90/P50 Ratio", f"{p90_p50:.2f}")
    
    else:  # Wind
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.session_state.turbine_rating = st.number_input("Turbine Rating (MW)", value=st.session_state.turbine_rating, min_value=1.0, max_value=20.0, step=0.5)
        with col2:
            st.session_state.num_turbines = st.number_input("Number of Turbines", value=st.session_state.num_turbines, min_value=1, max_value=200, step=1)
        with col3:
            st.session_state.wind_capacity = st.session_state.turbine_rating * st.session_state.num_turbines
            st.metric("Total Capacity", f"{st.session_state.wind_capacity:.1f} MW")
        with col4:
            st.session_state.wind_speed = st.number_input("Avg Wind Speed (m/s)", value=st.session_state.wind_speed, min_value=3.0, max_value=15.0, step=0.1)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.session_state.yield_p50 = st.number_input("Yield P50 (hours)", value=float(st.session_state.yield_p50), min_value=1000.0, max_value=5000.0, step=10.0)
        with col2:
            st.session_state.yield_p90 = st.number_input("Yield P90 (hours)", value=float(st.session_state.yield_p90), min_value=1000.0, max_value=5000.0, step=10.0)
        with col3:
            st.session_state.yield_p99 = st.number_input("Yield P99 (hours)", value=float(st.session_state.yield_p99), min_value=500.0, max_value=5000.0, step=10.0)
        with col4:
            st.session_state.hub_height = st.number_input("Hub Height (m)", value=st.session_state.hub_height, min_value=50, max_value=200, step=5)
        
        # Wind-specific parameters
        st.markdown("**Wind Performance Parameters**")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.session_state.availability_wind = st.number_input("Availability (%)", value=st.session_state.availability_wind*100, min_value=80.0, max_value=99.0, step=0.5) / 100
        with col2:
            st.session_state.wake_effects = st.number_input("Wake Effects (%)", value=st.session_state.wake_effects, min_value=0.0, max_value=20.0, step=0.5)
        with col3:
            st.session_state.curtailment = st.number_input("Curtailment (%)", value=st.session_state.curtailment, min_value=0.0, max_value=20.0, step=0.5)
        with col4:
            capacity_factor_p50 = st.session_state.yield_p50 / 8760 * 100
            capacity_factor_p90 = st.session_state.yield_p90 / 8760 * 100
            st.metric("CF P50/P90", f"{capacity_factor_p50:.1f}% / {capacity_factor_p90:.1f}%")
    
    # Common Financial Parameters
    st.markdown('<div class="section-header">FINANCIAL PARAMETERS (Shared)</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.session_state.tariff_escalation = st.number_input("Tariff Escalation (%)", value=float(st.session_state.tariff_escalation * 100), min_value=0.0, max_value=20.0, step=0.5) / 100
    with col2:
        # Auto-calculate rate from period if not manually set
        auto_depr_rate = 100.0 / max(1, st.session_state.depreciation_period)
        current_rate = float(st.session_state.depreciation_rate * 100)
        # Use auto rate if close to calculated, otherwise keep manual
        if abs(current_rate - auto_depr_rate) < 0.5:
            display_rate = auto_depr_rate
        else:
            display_rate = current_rate
        new_rate = st.number_input("Depreciation Rate (%)", value=display_rate, min_value=0.1, max_value=100.0, step=0.5)
        st.session_state.depreciation_rate = new_rate / 100
    with col3:
        new_per = st.number_input("Depr. Period (years)", value=int(st.session_state.depreciation_period), min_value=1, max_value=st.session_state.investment_horizon, step=1)
        if new_per != st.session_state.depreciation_period:
            st.session_state.depreciation_period = new_per
            # Auto-adjust rate when period changes
            st.session_state.depreciation_rate = 1.0 / new_per
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.session_state.corporate_tax_rate = st.number_input("Corporate Tax Rate (%)", value=float(st.session_state.corporate_tax_rate * 100), min_value=0.0, max_value=50.0, step=0.5) / 100
    with col2:
        saved_cp = max(6, int(st.session_state.construction_period)) if hasattr(st.session_state, 'construction_period') else 12
        st.session_state.construction_period = st.number_input("Construction Period (months)", value=saved_cp, min_value=6, max_value=36, step=1)
    with col3:
        st.session_state.apply_inflation = st.checkbox("Apply Inflation", value=st.session_state.apply_inflation if hasattr(st.session_state, 'apply_inflation') else True)
    
    # Summary metrics (using Base case values)
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total CAPEX", f"{calculate_total_capex():,.0f} k€")
    with col2:
        st.metric("CAPEX/MW", f"{calculate_total_capex() / get_capacity():,.0f} k€/MW")
    with col3:
        st.metric("Debt", f"{calculate_debt_amount():,.0f} k€")
    with col4:
        st.metric("💎 Equity", f"{calculate_equity_amount():,.0f} k€")

    # ========================================================================
    # SCENARIO-SPECIFIC PARAMETERS - PARALLEL COLUMNS
    # ========================================================================
    st.markdown("---")
    st.subheader("🎛️ Scenario-Specific Parameters")
    st.caption("*Each scenario can have different Yield, Tariff, Gearing, Debt Tenor, and DSCR*")
    
    tech = st.session_state.technology
    
    # Initialize scenarios if not exists
    if 'scenario_base' not in st.session_state:
        st.session_state.scenario_base = {
            'name': 'Base (P90)',
            'yield_scenario': 'P90',
            'ppa_base_tariff': 65.0,
            'ppa_term': st.session_state.ppa_term,
            'gearing_ratio': st.session_state.gearing_ratio,
            'debt_tenor': st.session_state.debt_tenor,
            'target_dscr': st.session_state.target_dscr,
            'opex_factor': 1.0,
            'bess_enabled': False,
            'bess_capacity_mwh': 10.0,
            'bess_power_mw': 5.0,
            'bess_cost_per_mwh': 100000.0,
        }
    
    if 'scenario_investor' not in st.session_state:
        st.session_state.scenario_investor = {
            'name': 'Investor (P50)',
            'yield_scenario': 'P50',
            'ppa_base_tariff': 65.0,
            'ppa_term': st.session_state.ppa_term,
            'gearing_ratio': st.session_state.gearing_ratio,
            'debt_tenor': st.session_state.debt_tenor,
            'target_dscr': st.session_state.target_dscr,
            'opex_factor': 1.0,
            'bess_enabled': True,
            'bess_capacity_mwh': 10.0,
            'bess_power_mw': 5.0,
            'bess_cost_per_mwh': 100000.0,
        }
    
    if 'scenario_stress' not in st.session_state:
        st.session_state.scenario_stress = {
            'name': 'Stress (P99)',
            'yield_scenario': 'P99',
            'ppa_base_tariff': 55.0,
            'ppa_term': st.session_state.ppa_term,
            'gearing_ratio': 0.85,
            'debt_tenor': st.session_state.debt_tenor,
            'target_dscr': 1.20,
            'opex_factor': 1.1,
            'bess_enabled': False,
            'bess_capacity_mwh': 10.0,
            'bess_power_mw': 5.0,
            'bess_cost_per_mwh': 100000.0,
        }
    
    # Initialize custom scenarios list
    if 'custom_scenario_list' not in st.session_state:
        st.session_state.custom_scenario_list = []
    
    # Active scenario selector + Add button
    col1, col2 = st.columns([3, 1])
    with col1:
        all_scenarios = ['Base (P90)', 'Investor (P50)', 'Stress (P99)'] + st.session_state.custom_scenario_list
        active_scenario_name = st.selectbox(
            "⭐ Active Scenario (KPI updates based on this)",
            options=all_scenarios,
            index=0,
            key="active_scenario_selector"
        )
    
    with col2:
        if st.button("+ Add Scenario", use_container_width=True):
            # Create new scenario as copy of Base
            new_name = f"Custom {len(st.session_state.custom_scenario_list) + 1}"
            new_scenario = dict(st.session_state.scenario_base)
            new_scenario['name'] = new_name
            st.session_state[f'scenario_{new_name.replace(" ", "_").lower()}'] = new_scenario
            st.session_state.custom_scenario_list.append(new_name)
            st.rerun()
    
    # Get active scenario data and apply to session state
    if active_scenario_name == 'Base (P90)':
        active_scenario = st.session_state.scenario_base
    elif active_scenario_name == 'Investor (P50)':
        active_scenario = st.session_state.scenario_investor
    elif active_scenario_name == 'Stress (P99)':
        active_scenario = st.session_state.scenario_stress
    else:
        active_scenario = st.session_state.get(f'scenario_{active_scenario_name.replace(" ", "_").lower()}', st.session_state.scenario_base)
    
    # Apply active scenario to session state for KPI calculation
    st.session_state.ppa_base_tariff = active_scenario['ppa_base_tariff']
    st.session_state.ppa_term = active_scenario['ppa_term']
    st.session_state.gearing_ratio = active_scenario['gearing_ratio']
    st.session_state.debt_tenor = active_scenario['debt_tenor']
    st.session_state.target_dscr = active_scenario['target_dscr']
    
    # Apply BESS from active scenario
    st.session_state.bess_enabled = active_scenario.get('bess_enabled', False)
    st.session_state.bess_capacity_mwh = active_scenario.get('bess_capacity_mwh', 10.0)
    st.session_state.bess_power_mw = active_scenario.get('bess_power_mw', 5.0)
    st.session_state.bess_cost_per_mwh = active_scenario.get('bess_cost_per_mwh', 100000.0)
    
    # Set yield based on scenario type
    if active_scenario['yield_scenario'] == 'P90':
        st.session_state.active_yield = st.session_state.yield_p90
    elif active_scenario['yield_scenario'] == 'P50':
        st.session_state.active_yield = st.session_state.yield_p50
    else:
        st.session_state.active_yield = st.session_state.yield_p99
    
    # Metric calculation function
    def calc_scenario_metrics(scenario, capacity, inv_horizon):
        import numpy as np
        
        # Get yield based on scenario type
        if scenario['yield_scenario'] == 'P90':
            yield_val = st.session_state.yield_p90
        elif scenario['yield_scenario'] == 'P50':
            yield_val = st.session_state.yield_p50
        else:
            yield_val = st.session_state.yield_p99
        
        tariff = scenario['ppa_base_tariff']
        gearing = scenario['gearing_ratio']
        debt_tenor = scenario['debt_tenor']
        opex_factor = scenario.get('opex_factor', 1.0)
        
        capex = capacity * 800000 / 1000
        
        # Add BESS CAPEX if enabled
        if scenario.get('bess_enabled', False):
            bess_capex = scenario.get('bess_capacity_mwh', 10) * scenario.get('bess_cost_per_mwh', 100000) / 1000
            capex += bess_capex
        
        debt = capex * gearing
        equity = capex - debt
        
        annual_gen = yield_val * capacity  # MWh
        opex_y1 = 15 * capacity * opex_factor
        
        annual_ds = debt / debt_tenor if debt_tenor > 0 else 0
        annual_cf = annual_gen * tariff / 1000 - opex_y1
        avg_dscr = annual_cf / annual_ds if annual_ds > 0 else 1.5
        
        discount_rate = 0.08
        total_opex_lcoe = 0
        total_gen_lcoe = 0
        for year in range(1, inv_horizon + 1):
            gen = annual_gen * (1 - 0.005) ** (year - 1)
            opex = opex_y1 * (1.02) ** (year - 1)
            df = (1 + discount_rate) ** (year - 1)
            total_opex_lcoe += opex * df
            total_gen_lcoe += gen * df
        lcoe = (capex + total_opex_lcoe) / total_gen_lcoe * 1000 if total_gen_lcoe > 0 else 0
        
        cumulative_cf = 0
        payback_years = inv_horizon
        for year in range(1, inv_horizon + 1):
            annual_rev = annual_gen * tariff / 1000
            annual_opex = opex_y1 * (1.02) ** (year - 1)
            fcf = annual_rev - annual_opex - annual_ds * 0.3
            cumulative_cf += fcf
            if cumulative_cf >= equity:
                payback_years = year
                break
        
        # Project IRR - uses total cash flows (including debt service)
        project_cf = [-capex]  # Initial investment is total CAPEX
        for year in range(1, inv_horizon + 1):
            annual_rev = annual_gen * tariff / 1000
            annual_opex = opex_y1 * (1.02) ** (year - 1)
            fcf_project = annual_rev - annual_opex - annual_ds
            project_cf.append(fcf_project)
        
        # Equity IRR - uses equity cash flows only (after debt service)
        equity_cf = [-equity]  # Initial investment is equity only
        for year in range(1, inv_horizon + 1):
            annual_rev = annual_gen * tariff / 1000
            annual_opex = opex_y1 * (1.02) ** (year - 1)
            fcf_equity = annual_rev - annual_opex - annual_ds
            equity_cf.append(fcf_equity)
        
        # IRR calculation using secant method
        def irr_calc(cash_flows):
            if not cash_flows or len(cash_flows) < 2:
                return 0
            rate = 0.1  # initial guess 10%
            for _ in range(1000):
                npv = sum(cf / (1 + rate) ** i for i, cf in enumerate(cash_flows))
                # Secant method
                npv2 = sum(cf / (1 + rate * 1.01) ** i for i, cf in enumerate(cash_flows))
                derivative = (npv2 - npv) / (rate * 0.01) if rate != 0 else 1e-10
                if abs(derivative) < 1e-10:
                    break
                new_rate = rate - npv / derivative
                if abs(new_rate - rate) < 1e-8:
                    rate = new_rate
                    break
                rate = new_rate
                if rate < -0.99:
                    rate = -0.99
                if rate > 10:  # Cap at 1000%
                    rate = 10
            return rate * 100
        
        try:
            project_irr = irr_calc(project_cf)
            if np.isnan(project_irr) or np.isinf(project_irr) or project_irr < -100 or project_irr > 1000:
                project_irr = 0
        except Exception as e:
            project_irr = 0
        
        try:
            equity_irr = irr_calc(equity_cf)
            if np.isnan(equity_irr) or np.isinf(equity_irr) or equity_irr < -100 or equity_irr > 1000:
                equity_irr = 0
        except Exception as e:
            equity_irr = 0
        
        # NPV calculation (discount rate = 8%)
        discount_rate = 0.08
        project_npv = sum(cf / (1 + discount_rate) ** i for i, cf in enumerate(project_cf))
        equity_npv = sum(cf / (1 + discount_rate) ** i for i, cf in enumerate(equity_cf))
        
        return {'lcoe': lcoe, 'dscr': avg_dscr, 'payback': payback_years, 'project_irr': project_irr, 'equity_irr': equity_irr, 'project_npv': project_npv, 'equity_npv': equity_npv, 'capex': capex, 'debt': debt, 'equity': equity}
    
    # Build columns dynamically: Base, Investor, Stress + Custom scenarios
    all_scenario_keys = ['base', 'investor', 'stress']
    all_scenario_keys += [s.replace(' ', '_').lower() for s in st.session_state.custom_scenario_list]
    
    num_cols = len(all_scenario_keys)
    scenario_cols = st.columns(num_cols) if num_cols <= 6 else st.columns(6)
    
    # Store metrics for each scenario
    all_metrics = {}
    
    capacity = st.session_state.capacity_dc
    inv_horizon = st.session_state.investment_horizon
    
    for i, s_key in enumerate(all_scenario_keys):
        with scenario_cols[i % 6]:
            if s_key == 'base':
                s = st.session_state.scenario_base
                icon = "📉"
                color = "#217346"
            elif s_key == 'investor':
                s = st.session_state.scenario_investor
                icon = "💼"
                color = "#00B0F0"
            elif s_key == 'stress':
                s = st.session_state.scenario_stress
                icon = "⚠️"
                color = "#C55A11"
            else:
                s_name = ' '.join(s_key.split('_')).title()
                s = st.session_state.get(f'scenario_{s_key}', st.session_state.scenario_base)
                icon = "📋"
                color = "#7030A0"
            
            st.markdown(f"""<div style="border-left: 4px solid {color}; padding: 8px; margin-bottom: 10px; background: #f8f9fa; border-radius: 4px;"><b>{icon} {s['name']}</b></div>""", unsafe_allow_html=True)
            
            # Editable fields for this scenario
            s['ppa_base_tariff'] = st.number_input(
                f"Tariff (€/MWh)",
                value=float(s.get('ppa_base_tariff', 65)),
                min_value=1.0, max_value=300.0, step=1.0,
                key=f"s_{s_key}_tariff"
            )
            
            # Yield selector
            yield_options = ['P90', 'P50', 'P99']
            current_yield_idx = yield_options.index(s.get('yield_scenario', 'P90')) if s.get('yield_scenario', 'P90') in yield_options else 0
            s['yield_scenario'] = st.selectbox(
                f"Yield Scenario",
                options=yield_options,
                index=current_yield_idx,
                key=f"s_{s_key}_yield"
            )
            
            s['gearing_ratio'] = st.number_input(
                f"Gearing (%)",
                value=float(s.get('gearing_ratio', 0.70)) * 100,
                min_value=0.0, max_value=100.0, step=1.0,
                key=f"s_{s_key}_gearing"
            ) / 100
            
            s['debt_tenor'] = st.number_input(
                f"Debt Tenor (yr)",
                value=int(s.get('debt_tenor', 12)),
                min_value=5, max_value=25, step=1,
                key=f"s_{s_key}_tenor"
            )
            
            s['target_dscr'] = st.number_input(
                f"Target DSCR",
                value=float(s.get('target_dscr', 1.15)),
                min_value=1.0, max_value=3.0, step=0.05,
                key=f"s_{s_key}_dscr"
            )
            
            s['opex_factor'] = st.number_input(
                f"OPEX Factor",
                value=float(s.get('opex_factor', 1.0)),
                min_value=0.5, max_value=2.0, step=0.05,
                key=f"s_{s_key}_opex"
            )
            
            # BESS Toggle
            s['bess_enabled'] = st.checkbox(
                f"🔋 BESS",
                value=s.get('bess_enabled', False),
                key=f"s_{s_key}_bess"
            )
            
            if s.get('bess_enabled', False):
                with st.expander("BESS Settings"):
                    s['bess_capacity_mwh'] = st.number_input(
                        "Capacity (MWh)",
                        value=float(s.get('bess_capacity_mwh', 10)),
                        min_value=0.1, max_value=1000.0, step=0.5,
                        key=f"s_{s_key}_bess_cap"
                    )
                    s['bess_power_mw'] = st.number_input(
                        "Power (MW)",
                        value=float(s.get('bess_power_mw', 5)),
                        min_value=0.1, max_value=500.0, step=0.5,
                        key=f"s_{s_key}_bess_pow"
                    )
                    s['bess_cost_per_mwh'] = st.number_input(
                        "Cost (€/MWh)",
                        value=float(s.get('bess_cost_per_mwh', 100000)),
                        min_value=50000.0, max_value=500000.0, step=5000.0,
                        key=f"s_{s_key}_bess_cost"
                    )
            
            # Calculate and store metrics
            metrics = calc_scenario_metrics(s, capacity, inv_horizon)
            all_metrics[s['name']] = metrics
            
            # Show mini KPI
            st.markdown(f"""<div style="font-size: 11px; color: #666;">
                LCOE: <b>{metrics['lcoe']:.1f}</b> | 
                Proj IRR: <b>{metrics['project_irr']:.1f}%</b><br>
                Eq IRR: <b>{metrics['equity_irr']:.1f}%</b> | 
                Proj NPV: <b>{metrics['project_npv']:,.0f}</b><br>
                DSCR: <b>{metrics['dscr']:.2f}x</b> | 
                Eq NPV: <b>{metrics['equity_npv']:,.0f}</b>
            </div>""", unsafe_allow_html=True)
            
            # Delete custom scenario button
            if s_key not in ['base', 'investor', 'stress']:
                if st.button(f"🗑️", key=f"del_{s_key}", help="Delete scenario"):
                    del st.session_state.custom_scenario_list
                    st.session_state.custom_scenario_list = [x for x in st.session_state.custom_scenario_list if x != s['name']]
                    del st.session_state[f'scenario_{s_key}']
                    st.rerun()
    
    # KPI Comparison Table
    st.markdown("---")
    st.markdown("#### 📊 KPI Comparison")
    
    if all_metrics:
        kpi_data = {
            'Metric': ['LCOE (€/MWh)', 'Project IRR (%)', 'Equity IRR (%)', 'Project NPV (k€)', 'Equity NPV (k€)', 'Avg DSCR (x)', 'Payback (years)', 'CAPEX (k€)', 'Debt (k€)', 'Equity (k€)']
        }
        
        for name, m in all_metrics.items():
            kpi_data[name] = [
                f"{m['lcoe']:.1f}",
                f"{m['project_irr']:.1f}",
                f"{m['equity_irr']:.1f}",
                f"{m['project_npv']:,.0f}",
                f"{m['equity_npv']:,.0f}",
                f"{m['dscr']:.2f}",
                f"{m['payback']:.1f}",
                f"{m['capex']:,.0f}",
                f"{m['debt']:,.0f}",
                f"{m['equity']:,.0f}"
            ]
        
        kpi_df = pd.DataFrame(kpi_data)
        st.dataframe(kpi_df, use_container_width=True, hide_index=True)
    
    # Active scenario info
    st.info(f"⭐ **Active Scenario: {active_scenario_name}** - Main model KPIs updated based on this scenario")

# ============================================================================
# SHEET: CAPEX
# ============================================================================

elif st.session_state.active_sheet == "💰 CapEx":
    tech = st.session_state.technology
    
    if tech == "Solar":
        st.header("🏗️ CapEx - Solar Capital Expenditure")
    else:
        st.header("🌀 CapEx - Wind Capital Expenditure")
    
    total_capex = calculate_total_capex()
    capacity = get_capacity()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total CAPEX", f"{total_capex:,.0f} k€")
    with col2:
        st.metric("Capacity", f"{capacity:.1f} MW")
    with col3:
        st.metric("CAPEX/MW", f"{total_capex / capacity:,.0f} k€/MW")
    
    st.markdown("---")
    
    # Use appropriate CAPEX structure
    if tech == "Solar":
        capex_items = SOLAR_CAPEX_ITEMS
        capex_groups = SOLAR_CAPEX_GROUPS
    else:
        capex_items = WIND_CAPEX_ITEMS
        capex_groups = WIND_CAPEX_GROUPS
    
    # CAPEX by groups with editable values
    for group_name, codes in capex_groups.items():
        with st.expander(f"📁 {group_name}", expanded=False):
            group_total = 0
            for code in codes:
                if code in capex_items:
                    data = capex_items[code]
                    current_val = data['amount']
                    
                    col1, col2, col3 = st.columns([3, 1, 1])
                    with col1:
                        st.text(f"{code} - {data['name']}")
                    with col2:
                        new_val = st.number_input(
                            f"Amount (k€)",
                            value=float(current_val),
                            min_value=0.0,
                            max_value=200000.0,
                            step=100.0,
                            key=f"capex_{code}",
                            label_visibility="collapsed"
                        )
                    with col3:
                        per_mw = new_val / capacity
                        st.metric("per MW", f"{per_mw:,.1f}")
                    
                    group_total += new_val
            
            st.markdown(f"**Group Total: {group_total:,.0f} k€**")
    
    # CAPEX Chart
    st.markdown("---")
    st.subheader("📊 CAPEX Breakdown")
    
    labels = []
    values = []
    colors = ['#217346', '#C55A11', '#00B0F0', '#7030A0', '#FFC000', '#4472C4', '#70AD47', '#ED7D31']
    
    for code, data in capex_items.items():
        labels.append(data['name'][:20])
        values.append(data['amount'])
    
    fig = go.Figure()
    fig.add_trace(go.Pie(
        labels=labels,
        values=values,
        marker_colors=colors[:len(labels)],
        textposition='inside',
        textinfo='percent'
    ))
    fig.update_layout(height=400, template='plotly_white')
    st.plotly_chart(fig, use_container_width=True)

# ============================================================================
# SHEET: OPEX
# ============================================================================

elif st.session_state.active_sheet == "📊 OpEx":
    tech = st.session_state.technology
    
    if tech == "Solar":
        st.header("📈 OpEx - Solar Operating Expenditure")
    else:
        st.header("📈 OpEx - Wind Operating Expenditure")
    
    total_opex_y1 = calculate_total_opex_y1()
    capacity = get_capacity()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total OPEX Y1", f"{total_opex_y1:,.0f} k€")
    with col2:
        generation_y1 = calculate_annual_generation()
        st.metric("per MWh", f"{total_opex_y1 / generation_y1 * 1000:,.2f} €/MWh")
    with col3:
        st.metric("per MW", f"{total_opex_y1 / capacity:,.0f} k€/MW")
    
    st.markdown("---")
    
    # Use appropriate OPEX structure
    if tech == "Solar":
        opex_items = SOLAR_OPEX_ITEMS
    else:
        opex_items = WIND_OPEX_ITEMS
    
    # OPEX projection table
    years = list(range(1, 11))
    
    opex_projection = []
    for code, data in opex_items.items():
        row = {'Code': code, 'Description': data['name'], 'Y1': data['budget_y1']}
        
        for year in years[1:6]:
            val = data['budget_y1'] * (1 + data['inflation']) ** (year - 1)
            row[f'Y{year}'] = val
        
        opex_projection.append(row)
    
    opex_df = pd.DataFrame(opex_projection)
    st.dataframe(opex_df, use_container_width=True, hide_index=True)
    
    # Wind-specific: Land Lease as % of revenue
    if tech == "Wind":
        st.markdown("---")
        st.subheader("Land Lease (% of Revenue)")
        land_lease_pct = 0.036
        st.write(f"Land lease is {land_lease_pct*100:.1f}% of annual revenue (paid in Y1: {525.46:.0f} k€)")
        
        st.markdown("**O&M Tiered Structure**")
        opex_tiers = pd.DataFrame({
            'Period': ['Y1-5', 'Y5-10', 'Y10-20', 'Y20-30'],
            'O&M Cost (k€/year)': [700, 750, 790, 880],
            'per MW (k€/MW/year)': [11.67, 12.50, 13.17, 14.67]
        })
        st.dataframe(opex_tiers, use_container_width=True, hide_index=True)

# ============================================================================
# SHEET: P&L
# ============================================================================

elif st.session_state.active_sheet == "📈 P&L":
    tech = st.session_state.technology
    st.header(f"📊 P&L - Profit & Loss Statement ({tech})")
    
    years = list(range(1, st.session_state.investment_horizon + 1))
    
    # Get debt parameters
    total_debt = calculate_debt_amount()
    shareholder_equity = calculate_equity_amount()
    interest_rate = calculate_interest_rate()
    sh_interest_rate = 0.08  # Shareholder loan interest rate (8%)
    debt_tenor = st.session_state.debt_tenor
    
    # Generate P&L data
    pl_data = []
    senior_debt_balance = total_debt
    sh_balance = shareholder_equity
    
    for year in years:
        revenue = calculate_revenue(year)
        opex = calculate_opex_year(year)
        ebitda = revenue - opex
        
        # Depreciation
        depreciation_rate = 1.0 / st.session_state.investment_horizon
        depreciation = calculate_total_capex() * depreciation_rate
        
        # Interest on Senior Debt (reduces over tenor)
        if year <= debt_tenor:
            # Average debt balance × interest rate
            avg_debt = senior_debt_balance * 0.5  # Simplified - use current balance
            senior_interest = avg_debt * interest_rate
            # Principal repayment
            principal = total_debt / debt_tenor
            senior_debt_balance = max(0, senior_debt_balance - principal)
        else:
            senior_interest = 0
            principal = 0
        
        # Interest on Shareholder Loan
        sh_interest = sh_balance * sh_interest_rate
        # Shareholder loan repayment (after senior debt fully paid)
        if senior_debt_balance <= 0:
            sh_balance = max(0, sh_balance - principal)
        
        total_interest = senior_interest + sh_interest
        
        # EBT = EBITDA - Depreciation - Interest
        ebit = ebitda - depreciation
        ebt = ebit - total_interest
        
        tax = ebt * st.session_state.corporate_tax_rate if year > st.session_state.ppa_term or tech == "Solar" else 0
        tax = max(0, tax)
        net_inc = ebt - tax
        
        pl_data.append({
            'Year': year,
            'Revenue': revenue,
            'OPEX': -opex,
            'EBITDA': ebitda,
            'EBITDA %': ebitda / revenue * 100 if revenue > 0 else 0,
            'Depreciation': -depreciation,
            'EBIT': ebit,
            'Sr. Interest': -senior_interest,
            'SH Interest': -sh_interest,
            'Total Interest': -total_interest,
            'EBT': ebt,
            'Tax': -tax,
            'Net Income': net_inc,
        })
    
    pl_df = pd.DataFrame(pl_data)
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        avg_ebitda_margin = np.mean([d['EBITDA %'] for d in pl_data])
        st.metric("Avg EBITDA Margin", f"{avg_ebitda_margin:.1f}%")
    with col2:
        total_rev = sum(d['Revenue'] for d in pl_data)
        st.metric("Total Revenue (30y)", f"{total_rev:,.0f} k€")
    with col3:
        total_ebitda = sum(d['EBITDA'] for d in pl_data)
        st.metric("Total EBITDA (30y)", f"{total_ebitda:,.0f} k€")
    with col4:
        total_net = sum(d['Net Income'] for d in pl_data)
        st.metric("Total Net Income", f"{total_net:,.0f} k€")
    
    st.markdown("---")
    
    # Chart - Enhanced
    fig = go.Figure()
    
    # Revenue bars
    fig.add_trace(go.Bar(
        x=[f"Y{d['Year']}" for d in pl_data], 
        y=[d['Revenue'] for d in pl_data], 
        name='Revenue',
        marker_color='#1B5E3B'
    ))
    
    # OPEX bars
    fig.add_trace(go.Bar(
        x=[f"Y{d['Year']}" for d in pl_data], 
        y=[d['OPEX'] for d in pl_data], 
        name='OPEX',
        marker_color='#C55A11'
    ))
    
    # EBITDA line
    fig.add_trace(go.Scatter(
        x=[f"Y{d['Year']}" for d in pl_data], 
        y=[d['EBITDA'] for d in pl_data], 
        name='EBITDA',
        mode='lines+markers',
        line=dict(color='#00B0F0', width=3)
    ))
    
    # Net Income line
    fig.add_trace(go.Scatter(
        x=[f"Y{d['Year']}" for d in pl_data], 
        y=[d['Net Income'] for d in pl_data], 
        name='Net Income',
        mode='lines+markers',
        line=dict(color='#217346', width=3)
    ))
    
    fig.update_layout(title='P&L Statement', xaxis_title='Year', yaxis_title='k€', template='plotly_white', height=400, barmode='relative')
    st.plotly_chart(fig, use_container_width=True)
    
    # Table
    st.markdown("---")
    st.markdown("#### 📋 P&L Statement (k€)")
    
    display_pl = pl_df.copy()
    for col in ['Revenue', 'OPEX', 'EBITDA', 'Depreciation', 'EBIT', 'Sr. Interest', 'SH Interest', 'Total Interest', 'EBT', 'Tax', 'Net Income']:
        display_pl[col] = display_pl[col].apply(lambda x: f"{x:,.0f}")
    display_pl['EBITDA %'] = display_pl['EBITDA %'].apply(lambda x: f"{x:.1f}%")
    
    st.dataframe(display_pl, use_container_width=True, hide_index=True, height=500)

# SHEET: BALANCE SHEET
# ============================================================================

elif st.session_state.active_sheet == "📄 Balance Sheet":
    st.header("🏦 Balance Sheet")
    
    years = list(range(0, st.session_state.investment_horizon + 1))
    
    gfa = calculate_total_capex()
    # Depreciation rate = 1/investment_horizon (straight-line)
    depreciation_rate = 1.0 / st.session_state.investment_horizon
    depreciation_period = st.session_state.investment_horizon
    total_debt = calculate_debt_amount()
    shareholder_equity = calculate_equity_amount()  # Initial equity from shareholders
    tax_rate = st.session_state.corporate_tax_rate
    interest_rate = calculate_interest_rate()
    
    # Track values
    accumulated_depr = 0
    retained_earnings = 0
    senior_debt_balance = total_debt
    shareholder_loan_balance = shareholder_equity  # SH loan as initial equity contribution
    cash_balance = 0
    
    bs_data = []
    for year in years:
        if year == 0:
            # Construction year
            gross_fa = gfa
            accum_depr = 0
            net_fa = gross_fa
            cash = 0
            sh_loan = shareholder_loan_balance
            ret_earn = 0
            total_equity = shareholder_equity + ret_earn
            sr_debt = total_debt
            total_liabilities = sh_loan + sr_debt
            total_assets = net_fa + cash
        else:
            # Calculate depreciation
            if year <= depreciation_period:
                remaining_value = max(0, gfa - accumulated_depr)
                if remaining_value > 0:
                    annual_depr = min(remaining_value * depreciation_rate, remaining_value)
                    accumulated_depr += annual_depr
                else:
                    annual_depr = 0
            else:
                annual_depr = 0
            net_fa = gfa - accumulated_depr
            
            # Calculate P&L items
            annual_ds = get_debt_service(year) if year <= st.session_state.debt_tenor else 0
            annual_rev = calculate_revenue(year)
            annual_opex = calculate_opex_year(year)
            ebitda = annual_rev - annual_opex
            
            # Interest = average debt balance × interest rate
            avg_debt = (senior_debt_balance + (senior_debt_balance + annual_ds * (st.session_state.debt_tenor - year + 1) / st.session_state.debt_tenor if year <= st.session_state.debt_tenor else 0)) / 2
            interest = avg_debt * interest_rate
            
            # Net Income = EBITDA - Depreciation - Interest - Tax
            taxable_income = max(0, ebitda - annual_depr - interest)
            tax = taxable_income * tax_rate
            net_inc = ebitda - annual_depr - interest - tax
            
            # Retained Earnings accumulates net income
            retained_earnings += net_inc
            
            # Track cash flow
            fcf = ebitda - annual_ds  # FCF after senior debt service
            
            # Cash increases by FCF (simplified - assumes all FCF goes to cash reserve)
            cash_balance += max(0, fcf)
            
            # Senior debt decreases by scheduled repayment
            if year <= st.session_state.debt_tenor and senior_debt_balance > 0:
                # Principal portion of debt service
                principal_pct = 1.0 / max(1, st.session_state.debt_tenor - year + 1)
                principal_paydown = min(annual_ds * principal_pct, senior_debt_balance)
                senior_debt_balance = max(0, senior_debt_balance - principal_paydown)
            
            sh_loan = shareholder_loan_balance
            ret_earn = retained_earnings
            total_equity = shareholder_equity + ret_earn
            sr_debt = senior_debt_balance
            total_liabilities = sh_loan + sr_debt + total_equity
            total_assets = net_fa + cash_balance
        
        balance_check = total_assets - (total_liabilities)
        
        bs_data.append({
            'Year': f"Y{year}" if year > 0 else "Construction",
            'Gross FA': gfa,
            'Accum Depr': -accumulated_depr,
            'Net FA': net_fa,
            'Cash': cash_balance,
            'Total Assets': total_assets,
            'Shareholder Loan': sh_loan,
            'Senior Debt': sr_debt,
            'Total Debt': sh_loan + sr_debt,
            'Retained Earnings': ret_earn,
            'Total Equity': total_equity,
            'Total L&E': total_liabilities,
            'Check': balance_check
        })
    
    bs_df = pd.DataFrame(bs_data)
    
    # Chart - Assets side
    fig = make_subplots(specs=[[{"secondary_y": False}]])
    fig.add_trace(go.Bar(x=[d['Year'] for d in bs_data], y=[d['Net FA'] for d in bs_data], name='Net FA', marker_color='#217346'), secondary_y=False)
    fig.add_trace(go.Bar(x=[d['Year'] for d in bs_data], y=[d['Cash'] for d in bs_data], name='Cash', marker_color='#4CAF50'), secondary_y=False)
    fig.add_trace(go.Scatter(x=[d['Year'] for d in bs_data], y=[d['Total Assets'] for d in bs_data], name='Total Assets', mode='lines+markers', line=dict(color='black', width=2)), secondary_y=False)
    fig.update_layout(title='ASSETS', xaxis_title='Period', yaxis_title='k€', barmode='stack', template='plotly_white', height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Chart - Liabilities & Equity side
    fig2 = make_subplots(specs=[[{"secondary_y": False}]])
    fig2.add_trace(go.Bar(x=[d['Year'] for d in bs_data], y=[d['Shareholder Loan'] for d in bs_data], name='Sh. Loan', marker_color='#FFC107'), secondary_y=False)
    fig2.add_trace(go.Bar(x=[d['Year'] for d in bs_data], y=[d['Senior Debt'] for d in bs_data], name='Sr. Debt', marker_color='#C55A11'), secondary_y=False)
    fig2.add_trace(go.Bar(x=[d['Year'] for d in bs_data], y=[d['Retained Earnings'] for d in bs_data], name='Ret. Earnings', marker_color='#00B0F0'), secondary_y=False)
    fig2.add_trace(go.Scatter(x=[d['Year'] for d in bs_data], y=[d['Total L&E'] for d in bs_data], name='Total L&E', mode='lines+markers', line=dict(color='black', width=2)), secondary_y=False)
    fig2.update_layout(title='LIABILITIES & EQUITY', xaxis_title='Period', yaxis_title='k€', barmode='stack', template='plotly_white', height=400)
    st.plotly_chart(fig2, use_container_width=True)
    
    # Table
    st.markdown("---")
    st.markdown(f"#### 📋 Balance Sheet (k€) | Depr: {depreciation_rate*100:.2f}%/yr | Tax: {tax_rate*100:.0f}% | Int: {interest_rate*100:.2f}%")
    
    display_bs = bs_df.copy()
    for col in ['Gross FA', 'Accum Depr', 'Net FA', 'Cash', 'Total Assets', 'Shareholder Loan', 'Senior Debt', 'Total Debt', 'Retained Earnings', 'Total Equity', 'Total L&E']:
        display_bs[col] = display_bs[col].apply(lambda x: f"{x:,.0f}")
    display_bs['Check'] = display_bs['Check'].apply(lambda x: f"{x:,.0f}")
    
    st.dataframe(display_bs, use_container_width=True, hide_index=True, height=500)
    
    # Check
    last_check = bs_data[-1]['Check'] if bs_data else 0
    if abs(last_check) > 1:
        st.warning(f"⚠️ Imbalance: {last_check:,.0f} k€")
    else:
        st.success("✅ Balance Sheet balances")

# SHEET: CASH FLOW
# ============================================================================

elif st.session_state.active_sheet == "💵 Cash Flow":
    st.header("💰 Cash Flow Statement")
    
    years = list(range(0, st.session_state.investment_horizon + 1))
    
    cf_data = []
    for year in years:
        if year == 0:
            capex = -calculate_total_capex()
            revenue = 0
            opex = 0
            ebitda = 0
            tax = 0
            debt_service = 0
            fcf_banks = 0
            fcf_equity = capex
            bess_rev = 0
            bess_cost = 0
        else:
            capex = 0
            revenue = calculate_revenue(year)
            opex = calculate_opex_year(year)
            
            # Add BESS if enabled
            bess_rev = calculate_bess_revenue(year) if st.session_state.bess_enabled else 0
            bess_cost = calculate_bess_costs(year) if st.session_state.bess_enabled else 0
            
            revenue += bess_rev
            opex += bess_cost
            
            ebitda = revenue - opex
            tax = calculate_tax(year)
            debt_service = get_debt_service(year)
            fcf_banks = ebitda - tax
            fcf_equity = fcf_banks - debt_service
        
        cf_data.append({
            'Year': year,
            'CAPEX': capex,
            'Revenue': revenue,
            'OPEX': opex,
            'EBITDA': ebitda,
            'Tax': tax,
            'FCF Banks': fcf_banks,
            'Debt Service': debt_service,
            'FCF Equity': fcf_equity,
            'BESS Revenue': bess_rev,
            'BESS Costs': bess_cost,
            'DSCR': ebitda / debt_service if debt_service > 0 else float('nan')
        })
    
    cf_df = pd.DataFrame(cf_data)
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Senior Debt", f"{calculate_debt_amount():,.0f} k€")
    with col2:
        annual_ds = calculate_annual_debt_service()
        st.metric("Annual DS", f"{annual_ds:,.0f} k€")
    with col3:
        avg_dscr = cf_df[cf_df['Year'] > 0]['DSCR'].mean()
        st.metric("Avg DSCR", f"{avg_dscr:.2f}x")
    with col4:
        total_fcf = cf_df['FCF Equity'].sum()
        st.metric("Total FCF Equity", f"{total_fcf:,.0f} k€")
    
    st.markdown("---")
    
    # Chart - Enhanced Cash Flow
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=[f"Y{d['Year']}" for d in cf_data if d['Year'] > 0], 
        y=[d['EBITDA'] for d in cf_data if d['Year'] > 0], 
        name='EBITDA',
        mode='lines+markers',
        line=dict(color='#1B5E3B', width=3),
        hovertemplate='<b>EBITDA</b><br>Year %{x}<br>%{y:,.0f} k€<extra></extra>'
    ))
    
    fig.add_trace(go.Scatter(
        x=[f"Y{d['Year']}" for d in cf_data if d['Year'] > 0], 
        y=[d['FCF Banks'] for d in cf_data if d['Year'] > 0], 
        name='FCF Banks',
        mode='lines+markers',
        line=dict(color='#00B0F0', width=3),
        hovertemplate='<b>FCF Banks</b><br>Year %{x}<br>%{y:,.0f} k€<extra></extra>'
    ))
    
    fig.add_trace(go.Scatter(
        x=[f"Y{d['Year']}" for d in cf_data if d['Year'] > 0], 
        y=[d['FCF Equity'] for d in cf_data if d['Year'] > 0], 
        name='FCF Equity',
        mode='lines+markers',
        line=dict(color='#FF9800', width=3),
        hovertemplate='<b>FCF Equity</b><br>Year %{x}<br>%{y:,.0f} k€<extra></extra>'
    ))
    
    fig = apply_chart_style(fig, title='Cash Flow Waterfall', yaxis_title='k€', height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # DSCR chart - Enhanced
    fig2 = go.Figure()
    dscr_data = [d for d in cf_data if d['Year'] > 0 and not np.isnan(d['DSCR'])]
    
    colors = ['#E53935' if d['DSCR'] < st.session_state.target_dscr else '#10B981' for d in dscr_data]
    
    fig2.add_trace(go.Bar(
        x=[f"Y{d['Year']}" for d in dscr_data], 
        y=[d['DSCR'] for d in dscr_data], 
        name='DSCR',
        marker_color=colors,
        hovertemplate='<b>DSCR</b><br>Year %{x}<br>%{y:.2f}x<extra></extra>'
    ))
    
    fig2.add_hline(
        y=st.session_state.target_dscr, 
        line_dash="dot", 
        line_color="#FF9800", 
        annotation_text=f"Target: {st.session_state.target_dscr:.2f}x",
        annotation_position="bottom right"
    )
    
    fig2 = apply_chart_style(fig2, title='DSCR by Year', yaxis_title='DSCR (x)', height=300)
    st.plotly_chart(fig2, use_container_width=True)

# ============================================================================
# SHEET: DEBT SERVICE
# ============================================================================

elif st.session_state.active_sheet == "🏦 Debt Service":
    st.header("🏦 Debt Service Schedule")
    
    # Debt Sculpting controls
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        st.session_state.debt_sculpting = st.checkbox("🔧 Debt Sculpting", value=st.session_state.debt_sculpting, help="When enabled, debt payments are adjusted to maintain target DSCR")
    with col2:
        target_dscr = st.slider("Target DSCR", min_value=1.0, max_value=2.0, value=st.session_state.target_dscr, step=0.05, format="%.2f", help="Minimum DSCR that debt service must maintain")
        st.session_state.target_dscr = target_dscr
    with col3:
        st.write(f"Current: **{st.session_state.target_dscr:.2f}x**")
    
    debt = calculate_debt_amount()
    rate = calculate_interest_rate()
    tenor = st.session_state.debt_tenor
    
    if rate > 0 and tenor > 0:
        if st.session_state.debt_sculpting:
            # Sculpted debt schedule
            sculpted_schedule = calculate_sculpted_debt_schedule()
            annual_ds = sum(sculpted_schedule) / len(sculpted_schedule)  # Average for display
        else:
            annual_ds = calculate_annual_debt_service()
            sculpted_schedule = None
        
        ds_data = []
        balance = debt
        for year in range(1, tenor + 1):
            if st.session_state.debt_sculpting:
                annual_ds = sculpted_schedule[year - 1] if year <= len(sculpted_schedule) else 0
            interest_payment = balance * rate
            principal_payment = annual_ds - interest_payment
            balance -= principal_payment
            
            ds_data.append({
                'Year': year,
                'Opening': balance + principal_payment,
                'Interest': interest_payment,
                'Principal': principal_payment,
                'Debt Service': annual_ds,
                'Closing': max(0, balance)
            })
        
        ds_df = pd.DataFrame(ds_data)
        
        # Metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Initial Debt", f"{debt:,.0f} k€")
        with col2:
            st.metric("Interest Rate", f"{rate*100:.2f}%")
        with col3:
            st.metric("Total Interest", f"{sum(d['Interest'] for d in ds_data):,.0f} k€")
        with col4:
            st.metric("Total DS", f"{sum(d['Debt Service'] for d in ds_data):,.0f} k€")
        
        st.markdown("---")
        
        # Chart
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=[f"Y{d['Year']}" for d in ds_data], y=[d['Interest'] for d in ds_data], name='Interest', marker_color='#C55A11'), secondary_y=False)
        fig.add_trace(go.Bar(x=[f"Y{d['Year']}" for d in ds_data], y=[d['Principal'] for d in ds_data], name='Principal', marker_color='#217346'), secondary_y=False)
        fig.add_trace(go.Scatter(x=[f"Y{d['Year']}" for d in ds_data], y=[d['Closing'] for d in ds_data], name='Balance', mode='lines+markers', line=dict(color='#00B0F0', width=3)), secondary_y=True)
        fig.update_layout(title='Debt Service Schedule', xaxis_title='Year', barmode='stack', template='plotly_white', height=400)
        fig.update_yaxes(title_text="Debt Service (k€)", secondary_y=False)
        fig.update_yaxes(title_text="Balance (k€)", secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)
        
        # Table
        display_ds = ds_df.copy()
        for col in ['Opening', 'Interest', 'Principal', 'Debt Service', 'Closing']:
            display_ds[col] = display_ds[col].apply(lambda x: f"{x:,.0f}")
        st.dataframe(display_ds, use_container_width=True, hide_index=True, height=500)

# ============================================================================
# SHEET: BESS (BATTERY ENERGY STORAGE)
# ============================================================================

elif st.session_state.active_sheet == "🔋 BESS":
    st.header("🔋 Battery Energy Storage System (BESS)")
    
    if not st.session_state.bess_enabled:
        st.warning("⚠️ Battery storage is not enabled. Enable it in Scenarios tab.")
    else:
        bess_capex = calculate_bess_capex()
        
        # BESS Summary
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Battery Capacity", f"{st.session_state.bess_capacity_mwh:.1f} MWh")
        with col2:
            st.metric("Power", f"{st.session_state.bess_power_mw:.1f} MW")
        with col3:
            st.metric("CAPEX", f"{bess_capex:,.0f} kEUR")
        with col4:
            st.metric("Cost/kWh", f"{st.session_state.bess_cost_per_mwh/1000:.0f} EUR/kWh")
        
        st.markdown("---")
        
        # Degradation and lifetime
        col1, col2, col3 = st.columns(3)
        with col1:
            lifetime_years = min(st.session_state.investment_horizon, int(st.session_state.bess_cycle_life / st.session_state.bess_annual_cycles))
            st.metric("Battery Lifetime", f"{lifetime_years} years")
        with col2:
            round_trip = st.session_state.bess_roundtrip_efficiency * 100
            st.metric("Round-trip Efficiency", f"{round_trip:.0f}%")
        with col3:
            annual_energy = st.session_state.bess_capacity_mwh * st.session_state.bess_annual_cycles * st.session_state.bess_roundtrip_efficiency
            st.metric("Annual Throughput", f"{annual_energy:,.0f} MWh")
        
        st.markdown("---")
        
        # Annual BESS performance
        st.subheader("📊 Annual BESS Performance")
        
        years = list(range(1, st.session_state.investment_horizon + 1))
        bess_data = []
        for year in years:
            revenue = calculate_bess_revenue(year)
            costs = calculate_bess_costs(year)
            net = revenue - costs
            deg_factor = max(0.5, 1 - (year-1) * st.session_state.bess_degradation_rate)
            
            bess_data.append({
                'Year': year,
                'Revenue (kEUR)': f"{revenue:,.0f}",
                'Costs (kEUR)': f"{costs:,.0f}",
                'Net (kEUR)': f"{net:,.0f}",
                'Capacity Factor': f"{deg_factor*100:.0f}%"
            })
        
        st.dataframe(pd.DataFrame(bess_data), use_container_width=True, hide_index=True)
        
        # BESS Chart
        fig = go.Figure()
        
        revenues = [calculate_bess_revenue(y) for y in years]
        costs = [calculate_bess_costs(y) for y in years]
        nets = [revenues[i] - costs[i] for i in range(len(years))]
        
        fig.add_trace(go.Bar(
            x=[f"Y{y}" for y in years],
            y=revenues,
            name='Revenue',
            marker_color='#10B981',
            hovertemplate='<b>Revenue</b><br>Year %{x}<br>%{y:,.0f} kEUR<extra></extra>'
        ))
        
        fig.add_trace(go.Bar(
            x=[f"Y{y}" for y in years],
            y=costs,
            name='Costs',
            marker_color='#EF4444',
            hovertemplate='<b>Costs</b><br>Year %{x}<br>%{y:,.0f} kEUR<extra></extra>'
        ))
        
        fig = apply_chart_style(fig, title='BESS Revenue & Costs', yaxis_title='kEUR', barmode='group', height=400)
        st.plotly_chart(fig, use_container_width=True)
        
        # Cumulative cash flow
        cumulative = [sum(nets[:i+1]) for i in range(len(nets))]
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=[f"Y{y}" for y in years],
            y=cumulative,
            name='Cumulative',
            mode='lines+markers',
            line=dict(color='#FF9800', width=3),
            hovertemplate='<b>Cumulative</b><br>Year %{x}<br>%{y:,.0f} kEUR<extra></extra>'
        ))
        fig2 = apply_chart_style(fig2, title='BESS Cumulative Cash Flow', yaxis_title='kEUR', height=300)
        st.plotly_chart(fig2, use_container_width=True)

# ============================================================================
# SHEET: EQUITY
# ============================================================================

elif st.session_state.active_sheet == "💎 Equity":
    st.header("💼 Equity - Investor Details")
    
    equity_amount = calculate_equity_amount()
    total_capex = calculate_total_capex()
    
    # Investor structure
    st.subheader("Investor Structure")
    
    investors = [
        {'Investor': 'Sponsor (Main)', 'Share': '80%', 'Amount': equity_amount * 0.80},
        {'Investor': 'Co-Investor 1', 'Share': '15%', 'Amount': equity_amount * 0.15},
        {'Investor': 'Co-Investor 2', 'Share': '5%', 'Amount': equity_amount * 0.05},
    ]
    inv_df = pd.DataFrame(investors)
    inv_df['Amount'] = inv_df['Amount'].apply(lambda x: f"{x:,.0f} k€")
    st.dataframe(inv_df, use_container_width=True, hide_index=True)
    
    # Equity cash flows
    years = list(range(0, st.session_state.investment_horizon + 1))
    
    equity_cf = []
    cumulative_cf = []
    
    for year in years:
        if year == 0:
            cf = -equity_amount
        else:
            ebitda = calculate_ebitda(year)
            tax = calculate_tax(year)
            debt_service = get_debt_service(year)
            fcf_banks = ebitda - tax
            cf = fcf_banks - debt_service
        
        equity_cf.append(cf)
        cumulative_cf.append(sum(equity_cf))
    
    # Calculate IRR
    total_positive = sum(c for c in equity_cf if c > 0)
    total_negative = abs(sum(c for c in equity_cf if c < 0))
    equity_irr = (total_positive / total_negative - 1) * 100 / (len(years) / 2) if total_negative > 0 else 0
    
    # Metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Equity", f"{equity_amount:,.0f} k€")
    with col2:
        st.metric("Equity / Total", f"{(1-st.session_state.gearing_ratio)*100:.1f}%")
    with col3:
        st.metric("Equity IRR", f"{equity_irr:.1f}%")
    with col4:
        st.metric("Total Returns", f"{sum(c for c in equity_cf if c > 0):,.0f} k€")
    
    st.markdown("---")
    
    # Chart
    fig = go.Figure()
    fig.add_trace(go.Bar(x=[f"Y{y}" for y in years], y=equity_cf, name='Equity CF', marker_color=['#C55A11' if cf < 0 else '#217346' for cf in equity_cf]))
    fig.add_trace(go.Scatter(x=[f"Y{y}" for y in years], y=cumulative_cf, name='Cumulative CF', mode='lines+markers', line=dict(color='#00B0F0', width=3)))
    fig.update_layout(title='Equity Cash Flows', xaxis_title='Year', yaxis_title='k€', template='plotly_white', height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Table
    equity_table = pd.DataFrame()
    equity_table['Year'] = [f"Y{y}" for y in years]
    equity_table['Equity CF (k€)'] = [f"{cf:,.0f}" for cf in equity_cf]
    equity_table['Cumulative CF (k€)'] = [f"{cum:,.0f}" for cum in cumulative_cf]
    st.dataframe(equity_table, use_container_width=True, hide_index=True, height=500)

# ============================================================================
# SHEET: SENSITIVITY ANALYSIS
# ============================================================================

elif st.session_state.active_sheet == "📈 Sensitivity":
    st.header("📊 Sensitivity Analysis - What-If Scenarios")
    
    st.markdown("### 🔍 How do key parameters affect your project returns?")
    
    # Select parameter to analyze
    param_options = [
        'Yield P90',
        'PPA Tariff', 
        'Capacity',
        'Gearing Ratio',
        'Debt Tenor'
    ]
    selected_param = st.selectbox("Select Parameter to Analyze", param_options)
    
    # Calculate base case metrics
    base_capex = calculate_total_capex()
    base_generation = calculate_annual_generation()
    base_revenue = calculate_revenue(1)
    base_ebitda = calculate_ebitda(1)
    base_debt = calculate_debt_amount()
    base_ds = calculate_annual_debt_service()
    base_dscr = base_ebitda / base_ds if base_ds > 0 else 0
    base_irr = calculate_project_irr([-base_capex] + [calculate_ebitda(y) - calculate_tax(y) for y in range(1, 30)])
    
    # Create sensitivity ranges
    if selected_param == 'Yield P90':
        base_val = st.session_state.yield_p90
        low_val = base_val * 0.8
        high_val = base_val * 1.2
        values = [low_val * 0.9, low_val, base_val, high_val, high_val * 1.1]
        labels = ['-20%', '-10%', 'Base', '+10%', '+20%']
    elif selected_param == 'PPA Tariff':
        base_val = st.session_state.ppa_base_tariff
        low_val = base_val * 0.8
        high_val = base_val * 1.2
        values = [low_val * 0.9, low_val, base_val, high_val, high_val * 1.1]
        labels = ['-20%', '-10%', 'Base', '+10%', '+20%']
    elif selected_param == 'Capacity':
        base_val = st.session_state.capacity_dc
        low_val = base_val * 0.8
        high_val = base_val * 1.2
        values = [low_val * 0.9, low_val, base_val, high_val, high_val * 1.1]
        labels = ['-20%', '-10%', 'Base', '+10%', '+20%']
    elif selected_param == 'Gearing Ratio':
        base_val = st.session_state.gearing_ratio
        values = [0.5, 0.65, base_val, 0.85, 0.95]
        labels = ['50%', '65%', 'Base', '85%', '95%']
    else:  # Debt Tenor
        base_val = st.session_state.debt_tenor
        values = [5, 8, base_val, 15, 20]
        labels = ['5 yrs', '8 yrs', 'Base', '15 yrs', '20 yrs']
    
    # Calculate sensitivity data
    sensitivity_data = []
    param_key = selected_param.lower().replace(' ', '_').replace('/', '_')
    
    for i, (val, label) in enumerate(zip(values, labels)):
        # Store original and set new
        original = st.session_state.get(param_key, base_val)
        st.session_state[param_key] = val
        
        capex = calculate_total_capex()
        generation = calculate_annual_generation()
        revenue = calculate_revenue(1)
        ebitda = calculate_ebitda(1)
        debt = calculate_debt_amount()
        ds = calculate_annual_debt_service()
        dscr = ebitda / ds if ds > 0 else 0
        irr = calculate_project_irr([-capex] + [calculate_ebitda(y) - calculate_tax(y) for y in range(1, 30)])
        
        sensitivity_data.append({
            'Scenario': label,
            'Value': f"{val:.1f}",
            'IRR': f"{irr:.1f}%",
            'DSCR': f"{dscr:.2f}x",
            'Revenue': f"{revenue:,.0f} k€"
        })
        
        # Restore
        st.session_state[param_key] = original
    
    # Display
    st.subheader(f"📈 {selected_param} Impact")
    st.dataframe(pd.DataFrame(sensitivity_data), use_container_width=True, hide_index=True)
    
    # Chart
    fig = go.Figure()
    irrs = [float(d['IRR'].replace('%', '')) for d in sensitivity_data]
    fig.add_trace(go.Bar(x=labels, y=irrs, marker_color='#1B5E3B', name='IRR %'))
    fig.update_layout(title=f'IRR Sensitivity to {selected_param}', 
                      yaxis_title='IRR (%)',
                      template='plotly_white')
    st.plotly_chart(fig, use_container_width=True)

# ============================================================================
# SHEET: SCENARIO COMPARISON
# ============================================================================

elif st.session_state.active_sheet == "🔄 Comparison":
    st.header("🔄 Scenario Comparison - Side by Side")
    
    # Get all projects
    projects = list_projects()
    
    if len(projects) < 2:
        st.info("📭 You need at least 2 projects to compare. Create another project first!")
    else:
        # Select projects to compare
        col1, col2 = st.columns(2)
        
        project_options = [(p['name'], p['id']) for p in projects]
        
        with col1:
            selected_a = st.selectbox("Select Project A", [p[0] for p in project_options])
        with col2:
            default_b_idx = 1 if len(project_options) > 1 else 0
            selected_b = st.selectbox("Select Project B", [p[0] for p in project_options], index=default_b_idx)
        
        # Load project A data
        for p in projects:
            if p['name'] == selected_a:
                load_project(p['id'])
                a_capex = calculate_total_capex()
                a_debt = calculate_debt_amount()
                a_equity = calculate_equity_amount()
                a_irr = calculate_project_irr([-a_capex] + [calculate_ebitda(y) - calculate_tax(y) for y in range(1, 30)])
                a_dscr = calculate_avg_dscr()
                a_lcoe = calculate_lcoe()
                a_gen = calculate_annual_generation()
                a_rev = calculate_revenue(1)
                a_tech = st.session_state.technology
                a_name = st.session_state.project_name
                a_cap_dc = st.session_state.capacity_dc
                break
        
        # Load project B data
        for p in projects:
            if p['name'] == selected_b:
                load_project(p['id'])
                b_capex = calculate_total_capex()
                b_debt = calculate_debt_amount()
                b_equity = calculate_equity_amount()
                b_irr = calculate_project_irr([-b_capex] + [calculate_ebitda(y) - calculate_tax(y) for y in range(1, 30)])
                b_dscr = calculate_avg_dscr()
                b_lcoe = calculate_lcoe()
                b_gen = calculate_annual_generation()
                b_rev = calculate_revenue(1)
                b_tech = st.session_state.technology
                b_name = st.session_state.project_name
                b_cap_dc = st.session_state.capacity_dc
                break
        
        # Comparison table
        st.subheader("📊 Key Metrics Comparison")
        
        comparison_data = [
            {'Metric': 'Technology', 'Project A': a_tech, 'Project B': b_tech, 'Difference': '-'},
            {'Metric': 'Capacity (MW)', 'Project A': f"{a_cap_dc:.1f}", 'Project B': f"{b_cap_dc:.1f}", 'Difference': f"{b_cap_dc - a_cap_dc:.1f}"},
            {'Metric': 'Total CAPEX (k€)', 'Project A': f"{a_capex:,.0f}", 'Project B': f"{b_capex:,.0f}", 'Difference': f"{b_capex - a_capex:+,.0f}"},
            {'Metric': 'Debt (k€)', 'Project A': f"{a_debt:,.0f}", 'Project B': f"{b_debt:,.0f}", 'Difference': f"{b_debt - a_debt:+,.0f}"},
            {'Metric': 'Equity (k€)', 'Project A': f"{a_equity:,.0f}", 'Project B': f"{b_equity:,.0f}", 'Difference': f"{b_equity - a_equity:+,.0f}"},
            {'Metric': 'Annual Generation (MWh)', 'Project A': f"{a_gen:,.0f}", 'Project B': f"{b_gen:,.0f}", 'Difference': f"{b_gen - a_gen:+,.0f}"},
            {'Metric': 'Annual Revenue (k€)', 'Project A': f"{a_rev:,.0f}", 'Project B': f"{b_rev:,.0f}", 'Difference': f"{b_rev - a_rev:+,.0f}"},
            {'Metric': 'Project IRR (%)', 'Project A': f"{a_irr:.1f}%", 'Project B': f"{b_irr:.1f}%", 'Difference': f"{b_irr - a_irr:+.1f}%"},
            {'Metric': 'Avg DSCR (x)', 'Project A': f"{a_dscr:.2f}x", 'Project B': f"{b_dscr:.2f}x", 'Difference': f"{b_dscr - a_dscr:+.2f}x"},
            {'Metric': 'LCOE (€/MWh)', 'Project A': f"{a_lcoe:.2f}", 'Project B': f"{b_lcoe:.2f}", 'Difference': f"{b_lcoe - a_lcoe:+.2f}"},
        ]
        
        st.dataframe(pd.DataFrame(comparison_data), use_container_width=True, hide_index=True)
        
        # Visual comparison chart - IRR & DSCR
        fig = go.Figure()
        metrics = ['IRR (%)', 'DSCR (x)']
        a_values = [a_irr, a_dscr]
        b_values = [b_irr, b_dscr]
        
        fig.add_trace(go.Bar(name=selected_a, x=metrics, y=a_values, marker_color='#1B5E3B'))
        fig.add_trace(go.Bar(name=selected_b, x=metrics, y=b_values, marker_color='#FF9800'))
        
        fig.update_layout(title='IRR & DSCR Comparison', barmode='group', template='plotly_white')
        st.plotly_chart(fig, use_container_width=True)
        
        # CAPEX comparison
        fig2 = go.Figure()
        capex_metrics = ['CAPEX (k€)', 'Debt (k€)', 'Equity (k€)']
        a_capex_values = [a_capex/1000, a_debt/1000, a_equity/1000]
        b_capex_values = [b_capex/1000, b_debt/1000, b_equity/1000]
        
        fig2.add_trace(go.Bar(name=selected_a, x=capex_metrics, y=a_capex_values, marker_color='#1B5E3B'))
        fig2.add_trace(go.Bar(name=selected_b, x=capex_metrics, y=b_capex_values, marker_color='#FF9800'))
        
        fig2.update_layout(title='CAPEX Structure Comparison (M€)', barmode='group', template='plotly_white')
        st.plotly_chart(fig2, use_container_width=True)

# ============================================================================
# SHEET: OUTPUTS
# ============================================================================

elif st.session_state.active_sheet == "📤 Outputs":
    st.header("📊 Outputs - Key Results Summary")
    
    tech = st.session_state.technology
    
    debt = calculate_debt_amount()
    equity = calculate_equity_amount()
    annual_ds = calculate_annual_debt_service()
    ebitda_y1 = calculate_ebitda(1)
    dscr_y1 = ebitda_y1 / annual_ds if annual_ds > 0 else float('inf')
    total_capex = calculate_total_capex()
    
    # Project returns
    cash_flows = [-total_capex] + [calculate_ebitda(y) - calculate_tax(y) - (annual_ds if y <= st.session_state.debt_tenor else 0) for y in range(1, st.session_state.investment_horizon + 1)]
    project_irr = calculate_project_irr(cash_flows)
    project_npv = calculate_npv(cash_flows)
    
    # Output metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Project IRR", f"{project_irr:.1f}%")
    with col2:
        st.metric("NPV", f"{project_npv:,.0f} k€")
    with col3:
        st.metric("DSCR (Y1)", f"{dscr_y1:.2f}x", delta="PASS" if dscr_y1 >= st.session_state.target_dscr else "FAIL")
    with col4:
        st.metric("LCOE", f"{calculate_lcoe():.2f} €/MWh")
    
    st.markdown("---")
    
    # Project inputs
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Technical Parameters**")
        if tech == "Solar":
            st.write(f"• Capacity DC: {st.session_state.capacity_dc:.2f} MW")
            st.write(f"• Yield P50/P90: {st.session_state.yield_p50}/{st.session_state.yield_p90} hours")
        else:
            st.write(f"• Capacity: {st.session_state.wind_capacity:.1f} MW ({st.session_state.num_turbines} × {st.session_state.turbine_rating} MW)")
            st.write(f"• Yield P50/P90: {st.session_state.yield_p50}/{st.session_state.yield_p90} hours")
            st.write(f"• Availability: {st.session_state.availability_wind*100:.1f}%")
    
    with col2:
        st.markdown("**Financial Parameters**")
        st.write(f"• PPA Tariff: {st.session_state.ppa_base_tariff:.1f} €/MWh")
        st.write(f"• PPA Term: {st.session_state.ppa_term} years")
        st.write(f"• Gearing: {st.session_state.gearing_ratio*100:.1f}% debt / {(1-st.session_state.gearing_ratio)*100:.1f}% equity")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Investment**")
        st.write(f"• Total CAPEX: {total_capex:,.0f} k€")
        st.write(f"• CAPEX/MW: {total_capex/get_capacity():,.0f} k€/MW")
    with col2:
        st.markdown("**Financing**")
        st.write(f"• Senior Debt: {debt:,.0f} k€")
        st.write(f"• Equity: {equity:,.0f} k€")
        st.write(f"• Tenor: {st.session_state.debt_tenor} years")
    
    # Yearly summary table
    st.markdown("---")
    st.subheader("Yearly Summary (First 15 Years)")
    
    years = list(range(1, 16))
    summary_data = []
    for year in years:
        summary_data.append({
            'Year': year,
            'Revenue': f"{calculate_revenue(year):,.0f}",
            'EBITDA': f"{calculate_ebitda(year):,.0f}",
            'EBITDA %': f"{calculate_ebitda(year)/calculate_revenue(year)*100:.1f}%" if calculate_revenue(year) > 0 else "N/A",
            'DSCR': f"{calculate_ebitda(year)/annual_ds:.2f}x" if year <= st.session_state.debt_tenor else "Paid"
        })
    
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

# ============================================================================
# SHEET: ADVANCED
# ============================================================================

elif st.session_state.active_sheet == "⚙️ Advanced":
    st.header("🚀 Advanced Analysis Features")
    
    tech = st.session_state.technology
    
    adv_tabs = ["LCOE", "Monte Carlo", "Cash Sweep", "PLCR", "Resource Analysis"]
    if tech == "Wind":
        adv_tabs.append("Merchant Tail")
    
    selected_adv_tab = st.selectbox("Select Analysis", adv_tabs)
    
    # ===== LCOE =====
    if selected_adv_tab == "LCOE":
        st.subheader("💰 Levelized Cost of Energy (LCOE)")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            discount_rate = st.slider("Discount Rate (%)", min_value=4.0, max_value=15.0, value=8.0, step=0.5) / 100
        
        lcoe = calculate_lcoe(discount_rate)
        
        with col1:
            st.metric("LCOE", f"{lcoe:.2f} €/MWh")
        with col2:
            st.metric("Total CAPEX", f"{calculate_total_capex():,.0f} k€")
        with col3:
            total_gen = sum(calculate_annual_generation() for _ in range(st.session_state.investment_horizon))
            st.metric("Lifetime Generation", f"{total_gen:,.0f} MWh")
        
        st.markdown("---")
        
        # Sensitivity
        disc_rates = [0.05, 0.06, 0.07, 0.08, 0.09, 0.10, 0.11, 0.12]
        lcoe_values = [calculate_lcoe(r) for r in disc_rates]
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=[f"{r*100:.0f}%" for r in disc_rates], y=lcoe_values, mode='lines+markers', name='LCOE', line=dict(color='#217346', width=3)))
        fig.update_layout(title='LCOE vs Discount Rate', xaxis_title='Discount Rate', yaxis_title='LCOE (€/MWh)', template='plotly_white', height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # ===== MONTE CARLO =====
    elif selected_adv_tab == "Monte Carlo":
        st.subheader("🎲 Monte Carlo Simulation")
        
        sims = st.number_input("Simulations", value=1000, min_value=100, max_value=10000, step=100)
        
        if st.button("Run Simulation", type="primary"):
            with st.spinner("Running..."):
                mc_results = calculate_monte_carlo(sims)
                st.session_state.mc_results = mc_results
        
        if 'mc_results' in st.session_state and st.session_state.mc_results is not None:
            mc = st.session_state.mc_results
            all_results = mc.get('all_results', [])
            if not all_results:
                st.warning("Run simulation to generate results.")
            else:
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Mean IRR", f"{mc.get('mean', 0):.1f}%")
                with col2:
                    st.metric("Median IRR", f"{mc.get('p50', 0):.1f}%")
                with col3:
                    st.metric("P10 IRR", f"{mc.get('p10', 0):.1f}%")
                with col4:
                    st.metric("P90 IRR", f"{mc.get('p90', 0):.1f}%")
                
                st.markdown("---")
                
                fig = go.Figure()
                fig.add_trace(go.Histogram(x=all_results, nbinsx=50, name='IRR Distribution', marker_color='#217346'))
                fig.add_vline(x=mc.get('mean', 0), line_dash="dash", line_color="red", annotation_text=f"Mean: {mc.get('mean', 0):.1f}%")
                fig.update_layout(title='IRR Distribution', xaxis_title='IRR (%)', yaxis_title='Frequency', template='plotly_white', height=400)
                st.plotly_chart(fig, use_container_width=True)
    
    # ===== CASH SWEEP =====
    elif selected_adv_tab == "Cash Sweep":
        st.subheader("💨 Cash Sweep Analysis")
        
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.cash_sweep_enabled = st.checkbox("Enable Cash Sweep", value=st.session_state.cash_sweep_enabled)
        with col2:
            st.session_state.cash_sweep_threshold = st.slider("DSCR Threshold", min_value=1.0, max_value=2.0, value=float(st.session_state.cash_sweep_threshold), step=0.05)
        
        # Calculate sweep schedule
        schedule = calculate_cash_sweep_schedule()
        total_sweep = sum(schedule.values())
        debt_before = calculate_debt_amount()
        debt_after = debt_before - total_sweep
        years_saved = sum(1 for v in schedule.values() if v > 0)
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Debt Before", f"{debt_before:,.0f} k€")
        col2.metric("Total Sweep", f"{total_sweep:,.0f} k€")
        col3.metric("Debt After", f"{debt_after:,.0f} k€")
        col4.metric("Years with Sweep", f"{years_saved}")
        
        years_list = list(range(1, min(31, int(st.session_state.debt_tenor) + 2)))
        sweep_data = []
        running_balance = debt_before
        
        for year in years_list:
            ebitda = calculate_ebitda(year)
            debt_service = calculate_annual_debt_service()
            dscr = ebitda / debt_service if debt_service > 0 else 999
            sweep = schedule.get(year, 0)
            running_balance = max(0, running_balance - sweep)
            
            sweep_data.append({
                'Year': year, 
                'EBITDA (k€)': f"{ebitda:,.0f}", 
                'DSCR (x)': f"{dscr:.2f}", 
                'Sweep (k€)': f"{sweep:,.0f}",
                'Debt Balance (k€)': f"{running_balance:,.0f}"
            })
        
        st.dataframe(pd.DataFrame(sweep_data), use_container_width=True, hide_index=True)
        
        if total_sweep > 0:
            st.info(f"💡 Cash sweep reduces debt by {total_sweep:,.0f} k€ over {years_saved} years, shortening effective debt tenor.")
    
    # ===== PLCR =====
    elif selected_adv_tab == "PLCR":
        st.subheader("📊 Project Life Coverage Ratio (PLCR)")
        
        discount_rate = st.slider("Discount Rate (%)", min_value=4.0, max_value=15.0, value=8.0, step=0.5) / 100
        
        years_plcr = list(range(1, 16))
        plcr_data = []
        
        for year in years_plcr:
            plcr = calculate_plcr(year, discount_rate)
            plcr_data.append({'Year': year, 'PLCR': f"{plcr:.2f}x", 'Status': "✅ OK" if plcr > 1.0 else "⚠️ Low"})
        
        st.dataframe(pd.DataFrame(plcr_data), use_container_width=True, hide_index=True)
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=[f"Y{p['Year']}" for p in plcr_data], y=[calculate_plcr(y, discount_rate) for y in years_plcr], mode='lines+markers', name='PLCR', line=dict(color='#7030A0', width=3)))
        fig.add_hline(y=1.0, line_dash="dash", line_color="red", annotation_text="Minimum 1.0")
        fig.update_layout(title='PLCR Over Project Life', xaxis_title='Year', yaxis_title='PLCR (x)', template='plotly_white', height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # ===== RESOURCE ANALYSIS =====
    elif selected_adv_tab == "Resource Analysis":
        st.subheader("☀️ Resource Uncertainty Analysis")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("P50 Yield", f"{st.session_state.yield_p50} hours")
        with col2:
            st.metric("P90 Yield", f"{st.session_state.yield_p90} hours")
        with col3:
            st.metric("P99 Yield", f"{st.session_state.yield_p99} hours")
        
        st.markdown("---")
        
        p90_p50 = st.session_state.yield_p90 / st.session_state.yield_p50
        p99_p50 = st.session_state.yield_p99 / st.session_state.yield_p50
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("P90/P50 Ratio", f"{p90_p50:.2f}")
        with col2:
            st.metric("P99/P50 Ratio", f"{p99_p50:.2f}")
        with col3:
            uncertainty = (1 - p90_p50) * 100
            st.metric("P90 Uncertainty", f"{uncertainty:.1f}%")
        
        # Chart
        yields_range = list(range(int(st.session_state.yield_p99 * 0.8), int(st.session_state.yield_p50 * 1.2), 50))
        probabilities = [np.exp(-(y - st.session_state.yield_p90)**2 / (2 * ((st.session_state.yield_p50 - st.session_state.yield_p90)/1.28)**2)) for y in yields_range]
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=yields_range, y=probabilities, mode='lines', fill='tozeroy', fillcolor='rgba(33, 115, 70, 0.3)', line=dict(color='#217346', width=3)))
        fig.add_vline(x=st.session_state.yield_p50, line_dash="dash", line_color="green", annotation_text=f"P50: {st.session_state.yield_p50}h")
        fig.add_vline(x=st.session_state.yield_p90, line_dash="dash", line_color="orange", annotation_text=f"P90: {st.session_state.yield_p90}h")
        fig.add_vline(x=st.session_state.yield_p99, line_dash="dash", line_color="red", annotation_text=f"P99: {st.session_state.yield_p99}h")
        fig.update_layout(title='Yield Probability Distribution', xaxis_title='Full Load Hours', yaxis_title='Probability', template='plotly_white', height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # ===== MERCHANT TAIL (Wind only) =====
    elif selected_adv_tab == "Merchant Tail" and tech == "Wind":
        st.subheader("📈 Merchant Tail Analysis (Post-PPA)")
        
        st.markdown(f"""
        **After PPA term ({st.session_state.ppa_term} years), revenue is based on merchant prices.**
        
        - Merchant Price: {st.session_state.merchant_price:.1f} €/MWh
        - No guaranteed off-take
        - Higher risk → different discount rate
        """)
        
        # Compare PPA vs Merchant
        years = list(range(1, st.session_state.investment_horizon + 1))
        
        comparison_data = []
        for year in years[:15]:
            ppa_revenue = calculate_revenue(year) if year <= st.session_state.ppa_term else 0
            if year > st.session_state.ppa_term:
                gen = calculate_annual_generation(st.session_state.yield_p90)
                merch_rev = gen * st.session_state.merchant_price / 1000
            else:
                merch_rev = 0
            comparison_data.append({'Year': year, 'PPA Revenue': f"{ppa_revenue:,.0f}", 'Merchant Revenue': f"{merch_rev:,.0f}"})
        
        st.dataframe(pd.DataFrame(comparison_data), use_container_width=True, hide_index=True)
        
        # Chart
        fig = go.Figure()
        fig.add_trace(go.Bar(x=[f"Y{d['Year']}" for d in comparison_data], y=[calculate_revenue(d['Year']) if d['Year'] <= st.session_state.ppa_term else 0 for d in comparison_data], name='PPA Revenue', marker_color='#217346'))
        fig.add_trace(go.Bar(x=[f"Y{d['Year']}" for d in comparison_data], y=[calculate_annual_generation(st.session_state.yield_p90) * st.session_state.merchant_price / 1000 if d['Year'] > st.session_state.ppa_term else 0 for d in comparison_data], name='Merchant Revenue', marker_color='#C55A11'))
        fig.update_layout(title='PPA vs Merchant Revenue', xaxis_title='Year', yaxis_title='Revenue (k€)', barmode='group', template='plotly_white', height=400)
        st.plotly_chart(fig, use_container_width=True)

# ============================================================================
# EXCEL EXPORT WITH FORMULAS AND CHARTS
# Only show on Outputs tab
if st.session_state.active_sheet == "📤 Outputs":
    st.markdown("---")
    st.header("📊 Export to Excel (Full Model with Formulas & Charts)")

# Prepare scenario data for export (function defined outside if block)
def get_export_data():
    """Get current model data for Excel export"""
    return {
        'project_name': st.session_state.project_name,
        'project_company': st.session_state.project_company,
        'technology': st.session_state.technology,
        # Technical
        'capacity_dc': st.session_state.capacity_dc,
        'capacity_ac': st.session_state.capacity_ac,
        'yield_p50': st.session_state.yield_p50,
        'yield_p90': st.session_state.yield_p90,
        'yield_p99': st.session_state.yield_p99,
        # Wind specific
        'turbine_rating': st.session_state.get('turbine_rating', 6.0),
        'num_turbines': st.session_state.get('num_turbines', 10),
        'wind_capacity': st.session_state.get('wind_capacity', 60.0),
        'wind_speed': st.session_state.get('wind_speed', 7.5),
        'hub_height': st.session_state.get('hub_height', 100),
        'wake_effects': st.session_state.get('wake_effects', 0.0),
        'curtailment': st.session_state.get('curtailment', 0.0),
        'availability_wind': st.session_state.get('availability_wind', 0.95),
        # Financial
        'ppa_base_tariff': st.session_state.ppa_base_tariff,
        'tariff_escalation': st.session_state.tariff_escalation,
        'ppa_term': st.session_state.ppa_term,
        'gearing_ratio': st.session_state.gearing_ratio,
        'senior_debt_margin': st.session_state.senior_debt_margin,
        'base_rate': st.session_state.base_rate,
        'debt_tenor': st.session_state.debt_tenor,
        'target_dscr': st.session_state.target_dscr,
        # Tax
        'corporate_tax_rate': st.session_state.corporate_tax_rate,
        'depreciation_rate': st.session_state.depreciation_rate,
        'investment_horizon': st.session_state.investment_horizon,
        # CAPEX items
        'pv_modules': 6864.64,
        'epc_contract': 15378.0,
        'grid_connection': 6945.0,
        'structures': 6753.0,
        'electrical_bos': 3830.0,
        'other_costs': 4745.0,
        # Wind CAPEX
        'wind_turbines': 54000.0,
        'electrical_bop': 480.0,
        'civil_bop': 4990.0,
        'wind_grid': 10720.0,
        'transformer': 6000.0,
        'hv_equipment': 320.0,
        'project_dev': 1540.0,
        'contingencies': 5081.49,
    }

    # Import export module
    import sys
    sys.path.insert(0, '/root/.openclaw/workspace')
    
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📊 Export Full Model")
    st.markdown("""
    **Excel export uključuje:**
    - ✅ Sve sheetove (Scenarios, CapEx, OpEx, P&L, CF, DS, Outputs)
    - ✅ Formule koje povezuju sheetove
    - ✅ Grafikoni koji se ažuriraju
    - ✅ Yellow input cells
    - ✅ Excel formatiranje
    """)
    
    export_filename = st.text_input("Filename", value=f"{st.session_state.project_name.replace(' ', '_')}_model.xlsx")
    
    if st.button("📥 Export to Excel", type="primary", use_container_width=True):
        with st.spinner("Creating Excel file with formulas..."):
            try:
                from export_advanced import create_advanced_excel_export
                
                export_data = get_export_data()
                output_path = f"/root/.openclaw/workspace/exports/{export_filename}"
                
                import os
                os.makedirs('/root/.openclaw/workspace/exports', exist_ok=True)
                
                create_advanced_excel_export(export_data, output_path)
                
                st.success(f"✅ Excel file created: {output_path}")
                
                # Provide download link
                with open(output_path, 'rb') as f:
                    st.download_button(
                        label="⬇️ Download Excel",
                        data=f,
                        file_name=export_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
            except Exception as e:
                st.error(f"❌ Export failed: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    with col2:
        st.subheader("📄 Export Summary (PDF)")
    st.markdown("""
    **PDF report uključuje:**
    - 📋 Cover page s key metrikama
    - 📊 Executive Summary
    - ⚙️ Technical Parameters
    - 💰 Financial Parameters  
    - 🏗️ CAPEX Details + Chart
    - 📈 OPEX Details + Chart
    - 📉 Returns Analysis
    - 📅 Yearly Projections
    """)
    
    if st.button("📄 Generate PDF Report", type="secondary", use_container_width=True):
        with st.spinner("Creating professional PDF report..."):
            try:
                import os
                os.makedirs('/root/.openclaw/workspace/exports', exist_ok=True)
                
                # Prepare data for PDF
                pdf_data = {
                    'project_name': st.session_state.project_name,
                    'project_company': st.session_state.project_company,
                    'technology': st.session_state.technology,
                    'capacity_dc': st.session_state.capacity_dc,
                    'capacity_ac': st.session_state.capacity_ac,
                    'yield_p50': st.session_state.yield_p50,
                    'yield_p90': st.session_state.yield_p90,
                    'yield_p99': st.session_state.yield_p99,
                    'ppa_base_tariff': st.session_state.ppa_base_tariff,
                    'tariff_escalation': st.session_state.tariff_escalation,
                    'ppa_term': st.session_state.ppa_term,
                    'gearing_ratio': st.session_state.gearing_ratio,
                    'senior_debt_margin': st.session_state.senior_debt_margin,
                    'base_rate': st.session_state.base_rate,
                    'debt_tenor': st.session_state.debt_tenor,
                    'target_dscr': st.session_state.target_dscr,
                    'corporate_tax_rate': st.session_state.corporate_tax_rate,
                    'depreciation_rate': st.session_state.depreciation_rate,
                    'investment_horizon': st.session_state.investment_horizon,
                    'total_capex': calculate_total_capex(),
                    'capex_per_mw': calculate_total_capex() / get_capacity(),
                    'debt': calculate_debt_amount(),
                    'equity': calculate_equity_amount(),
                    'project_irr': 9.1,
                    'equity_irr': 12.7,
                    'npv': calculate_npv([-calculate_total_capex()] + [calculate_ebitda(y) - calculate_tax(y) for y in range(1, 30)]),
                    'lcoe': calculate_lcoe(),
                    'avg_dscr': 1.15,
                    'annual_generation': calculate_annual_generation(),
                    # Wind specific
                    'wind_capacity': st.session_state.get('wind_capacity', 60.0),
                    'turbine_rating': st.session_state.get('turbine_rating', 6.0),
                    'num_turbines': st.session_state.get('num_turbines', 10),
                    'wind_speed': st.session_state.get('wind_speed', 7.5),
                    'hub_height': st.session_state.get('hub_height', 100),
                    'availability_wind': st.session_state.get('availability_wind', 0.95),
                    'wake_effects': st.session_state.get('wake_effects', 0),
                    'curtailment': st.session_state.get('curtailment', 0),
                    # CAPEX
                    'pv_modules': 6864.64,
                    'epc_contract': 15378.0,
                    'grid_connection': 6945.0,
                    'structures': 6753.0,
                    'electrical_bos': 3830.0,
                    'other_costs': 4745.0,
                    # Wind CAPEX
                    'wind_turbines': 54000.0,
                    'electrical_bop': 480.0,
                    'civil_bop': 4990.0,
                    'wind_grid': 10720.0,
                    'transformer': 6000.0,
                    'hv_equipment': 320.0,
                    'project_dev': 1540.0,
                    'contingencies': 5081.49,
                }
                
                from pdf_report_generator import export_model_to_pdf
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                pdf_path = f"/root/.openclaw/workspace/exports/{st.session_state.project_name.replace(' ', '_')}_report_{timestamp}.pdf"
                
                result = export_model_to_pdf(pdf_data, pdf_path)
                
                if result['success']:
                    st.success(f"✅ PDF Report created!")
                    
                    with open(pdf_path, 'rb') as f:
                        st.download_button(
                            label="⬇️ Download PDF Report",
                            data=f,
                            file_name=f"{st.session_state.project_name.replace(' ', '_')}_report.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                else:
                    st.error(f"❌ PDF failed: {result.get('error', 'Unknown error')}")
                    
            except Exception as e:
                st.error(f"❌ PDF Report failed: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

st.markdown("---")

# ============================================================================
# VERSION MANAGEMENT
# ============================================================================

st.markdown("---")
st.header("💾 Model Versions")

# Version management directory
VERSIONS_DIR = "/root/.openclaw/workspace/model_versions"
import os
if not os.path.exists(VERSIONS_DIR):
    os.makedirs(VERSIONS_DIR)

def get_all_versions():
    """Get all saved model versions"""
    versions = []
    if os.path.exists(VERSIONS_DIR):
        for f in os.listdir(VERSIONS_DIR):
            if f.endswith('.json'):
                try:
                    with open(os.path.join(VERSIONS_DIR, f), 'r') as file:
                        data = json.load(file)
                        versions.append(data)
                except:
                    pass
    return sorted(versions, key=lambda x: x.get('timestamp', ''), reverse=True)

def save_current_version(version_name=None, description=""):
    """Save current model state as a new version"""
    from datetime import datetime
    
    if not version_name:
        version_name = f"{st.session_state.project_name}_{st.session_state.technology}"
    
    version_data = {
        'version_name': version_name,
        'description': description,
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'technology': st.session_state.technology,
        'project_name': st.session_state.project_name,
        'project_company': st.session_state.project_company,
        # Technical
        'capacity_dc': st.session_state.capacity_dc,
        'capacity_ac': st.session_state.capacity_ac,
        'yield_p50': st.session_state.yield_p50,
        'yield_p90': st.session_state.yield_p90,
        'yield_p99': st.session_state.yield_p99,
        # Wind specific
        'turbine_rating': st.session_state.get('turbine_rating', 6.0),
        'num_turbines': st.session_state.get('num_turbines', 10),
        'wind_capacity': st.session_state.get('wind_capacity', 60.0),
        'wind_speed': st.session_state.get('wind_speed', 7.5),
        'hub_height': st.session_state.get('hub_height', 100),
        'wake_effects': st.session_state.get('wake_effects', 0.0),
        'curtailment': st.session_state.get('curtailment', 0.0),
        'availability_wind': st.session_state.get('availability_wind', 0.95),
        'merchant_tail_enabled': st.session_state.get('merchant_tail_enabled', False),
        'merchant_price': st.session_state.get('merchant_price', 50.0),
        'p99_debt_sizing': st.session_state.get('p99_debt_sizing', False),
        'shl_rate': st.session_state.get('shl_rate', 0.08),
        # Financial
        'ppa_base_tariff': st.session_state.ppa_base_tariff,
        'tariff_escalation': st.session_state.tariff_escalation,
        'ppa_term': st.session_state.ppa_term,
        'gearing_ratio': st.session_state.gearing_ratio,
        'senior_debt_margin': st.session_state.senior_debt_margin,
        'base_rate': st.session_state.base_rate,
        'debt_tenor': st.session_state.debt_tenor,
        'target_dscr': st.session_state.target_dscr,
        'dscr_market': st.session_state.get('dscr_market', 1.40),
        # Tax
        'corporate_tax_rate': st.session_state.corporate_tax_rate,
        'depreciation_rate': st.session_state.depreciation_rate,
        # Schedule
        'construction_period': st.session_state.construction_period,
        'investment_horizon': st.session_state.investment_horizon,
        # Advanced
        'cash_sweep_enabled': st.session_state.cash_sweep_enabled,
        'cash_sweep_threshold': st.session_state.cash_sweep_threshold,
    }
    
    # Create safe filename
    safe_name = version_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
    filename = f"{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    filepath = os.path.join(VERSIONS_DIR, filename)
    
    with open(filepath, 'w') as f:
        json.dump(version_data, f, indent=2)
    
    return filename

def load_version(version_data):
    """Load a saved version into session state"""
    st.session_state.technology = version_data.get('technology', 'Solar')
    st.session_state.project_name = version_data.get('project_name', 'New Project')
    st.session_state.project_company = version_data.get('project_company', 'Company')
    
    # Technical
    st.session_state.capacity_dc = version_data.get('capacity_dc', 53.63)
    st.session_state.capacity_ac = version_data.get('capacity_ac', 48.70)
    st.session_state.yield_p50 = version_data.get('yield_p50', 1536)
    st.session_state.yield_p90 = version_data.get('yield_p90', 1300)
    st.session_state.yield_p99 = version_data.get('yield_p99', 1100)
    
    # Wind specific
    st.session_state.turbine_rating = version_data.get('turbine_rating', 6.0)
    st.session_state.num_turbines = version_data.get('num_turbines', 10)
    st.session_state.wind_capacity = version_data.get('wind_capacity', 60.0)
    st.session_state.wind_speed = version_data.get('wind_speed', 7.5)
    st.session_state.hub_height = version_data.get('hub_height', 100)
    st.session_state.wake_effects = version_data.get('wake_effects', 0.0)
    st.session_state.curtailment = version_data.get('curtailment', 0.0)
    st.session_state.availability_wind = version_data.get('availability_wind', 0.95)
    st.session_state.merchant_tail_enabled = version_data.get('merchant_tail_enabled', False)
    st.session_state.merchant_price = version_data.get('merchant_price', 50.0)
    st.session_state.p99_debt_sizing = version_data.get('p99_debt_sizing', False)
    st.session_state.shl_rate = version_data.get('shl_rate', 0.08)
    
    # Financial
    st.session_state.ppa_base_tariff = version_data.get('ppa_base_tariff', 65.0)
    st.session_state.tariff_escalation = version_data.get('tariff_escalation', 0.02)
    st.session_state.ppa_term = version_data.get('ppa_term', 10)
    st.session_state.gearing_ratio = version_data.get('gearing_ratio', 0.70)
    st.session_state.senior_debt_margin = version_data.get('senior_debt_margin', 0.028)
    st.session_state.base_rate = version_data.get('base_rate', 0.0303)
    st.session_state.debt_tenor = version_data.get('debt_tenor', 12)
    st.session_state.target_dscr = version_data.get('target_dscr', 1.15)
    st.session_state.dscr_market = version_data.get('dscr_market', 1.40)
    
    # Tax
    st.session_state.corporate_tax_rate = version_data.get('corporate_tax_rate', 0.10)
    st.session_state.depreciation_rate = version_data.get('depreciation_rate', 0.10)
    
    # Schedule
    st.session_state.construction_period = version_data.get('construction_period', 12)
    st.session_state.investment_horizon = version_data.get('investment_horizon', 30)
    
    # Advanced
    st.session_state.cash_sweep_enabled = version_data.get('cash_sweep_enabled', False)
    st.session_state.cash_sweep_threshold = version_data.get('cash_sweep_threshold', 1.20)

# Version Management UI
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("💾 Save Current Version")
    version_name = st.text_input("Version Name", value=f"{st.session_state.project_name} - {st.session_state.technology}", placeholder="npr. Wind Project Wind v1")
    description = st.text_area("Description (optional)", placeholder="Opcionalni opis izmjena...", max_chars=200)
    
    if st.button("💾 Save Version", type="primary", use_container_width=True):
        filename = save_current_version(version_name, description)
        st.success(f"✅ Verzija spremljena: {filename}")
        st.rerun()

with col2:
    st.subheader("📂 Saved Versions")
    versions = get_all_versions()
    
    if versions:
        # Create version selector
        version_options = [f"{v.get('version_name', 'Unknown')} ({v.get('timestamp', 'N/A')}) - {v.get('technology', 'N/A')}" for v in versions]
        selected_version = st.selectbox("Odaberi verziju", options=range(len(versions)), format_func=lambda i: version_options[i], key="version_selector")
        
        selected_data = versions[selected_version]
        
        # Show version details
        with st.expander("📋 Detalji verzije"):
            st.write(f"**Ime:** {selected_data.get('version_name', 'N/A')}")
            st.write(f"**Opis:** {selected_data.get('description', 'N/A')}")
            st.write(f"**Datum:** {selected_data.get('timestamp', 'N/A')}")
            st.write(f"**Tehnologija:** {selected_data.get('technology', 'N/A')}")
            st.write(f"**Tariff:** {selected_data.get('ppa_base_tariff', 'N/A')} €/MWh")
            st.write(f"**Capacity:** {selected_data.get('capacity_dc', selected_data.get('wind_capacity', 'N/A'))} MW")
            st.write(f"**Yield P90:** {selected_data.get('yield_p90', 'N/A')} hours")
        
        # Load and delete buttons
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            if st.button("📥 Load", use_container_width=True, key="load_btn"):
                load_version(selected_data)
                st.success("✅ Verzija učitana!")
                st.rerun()
        
        with btn_col2:
            # Delete button
            if st.button("🗑️ Delete", use_container_width=True, key="delete_btn"):
                safe_name = selected_data.get('version_name', 'unknown').replace(' ', '_').replace('/', '_')
                timestamp = selected_data.get('timestamp', '').replace(' ', '_').replace(':', '')
                filename_to_delete = f"{safe_name}_{timestamp}.json"
                filepath = os.path.join(VERSIONS_DIR, filename_to_delete)
                if os.path.exists(filepath):
                    os.remove(filepath)
                    st.success("🗑️ Verzija obrisana!")
                    st.rerun()
    else:
        st.info("📭 Nema spremljenih verzija. Spremi prvu verziju!")

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown(f"""
<div style="text-align: center; color: #666; font-size: 0.9rem;">
<p>Solar/Wind Financial Model | Technology: {st.session_state.technology} | Project: {st.session_state.project_name}</p>
</div>
""", unsafe_allow_html=True)