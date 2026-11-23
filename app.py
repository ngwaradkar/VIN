# -*- coding: utf-8 -*-
import streamlit as st
import base64
import os
import re
import datetime
import io
import openpyxl
from engine import (
    initialize_tcf1_workbook,
    initialize_tcf2_workbook,
    update_today_vin_generation,
    update_paint_float_data
)

# Helper to load image to base64
def get_base64_image(img_path):
    if os.path.exists(img_path):
        with open(img_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return ""

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="TML VIN Generation & Production Plan",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# THEME STATE & CONTROLLER
# ─────────────────────────────────────────────
if "theme" not in st.session_state:
    st.session_state["theme"] = "light"  # Default to Enterprise Light

# Top Navigation Bar with Enterprise Branding & Theme Switcher
nav_col_left, nav_col_right = st.columns([7, 2])
with nav_col_left:
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:10px; padding: 4px 0;">
            <span style="font-weight:800; font-size:1.15rem; color:#002B49; letter-spacing:-0.4px;">
                TATA MOTORS
            </span>
            <span style="height:16px; width:1px; background:#CBD5E1; display:inline-block;"></span>
            <span style="font-size:0.85rem; font-weight:600; color:#64748B; text-transform:uppercase; letter-spacing:0.06em;">
                Passenger Vehicles Planning Division
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )
with nav_col_right:
    theme_choice = st.segmented_control(
        "Dashboard Theme",
        options=["☀️ Enterprise Light", "🌙 Dark Mode"],
        default="☀️ Enterprise Light" if st.session_state["theme"] == "light" else "🌙 Dark Mode",
        label_visibility="collapsed",
        key="app_theme_switcher"
    )
    if theme_choice is not None:
        target_theme = "light" if "Light" in theme_choice else "dark"
        if target_theme != st.session_state["theme"]:
            st.session_state["theme"] = target_theme
            st.rerun()

is_light = (st.session_state.get("theme", "light") == "light")

# ─────────────────────────────────────────────
# CUSTOM CSS — ENTERPRISE TATA MOTORS SYSTEM
# ─────────────────────────────────────────────
if is_light:
    # ── LIGHT MODE STYLING ──
    css_theme = """
    :root {
        --bg-main: #F5F7FA;
        --card-bg: #FFFFFF;
        --card-border: #E2E8F0;
        --primary-navy: #002B49;
        --accent-cyan: #0284C7;
        --text-primary: #0F172A;
        --text-secondary: #475569;
        --text-muted: #64748B;
        --badge-bg: #F1F5F9;
        --badge-border: #CBD5E1;
        --shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.05), 0 1px 2px rgba(0, 0, 0, 0.04);
        --shadow-md: 0 4px 14px rgba(15, 23, 42, 0.06), 0 1px 3px rgba(15, 23, 42, 0.04);
        --shadow-lg: 0 10px 25px rgba(15, 23, 42, 0.08), 0 4px 10px rgba(15, 23, 42, 0.04);
    }
    .stApp {
        background-color: var(--bg-main) !important;
        background: radial-gradient(circle at 10% 10%, rgba(2, 132, 199, 0.035), transparent 45%),
                    radial-gradient(circle at 90% 90%, rgba(16, 185, 129, 0.03), transparent 45%),
                    #F5F7FA !important;
        color: var(--text-primary) !important;
    }
    .tml-card {
        background: var(--card-bg);
        border: 1px solid var(--card-border);
        border-radius: 14px;
        box-shadow: var(--shadow-sm);
        padding: 1.5rem;
        margin-bottom: 1.25rem;
        transition: box-shadow 0.2s ease, border-color 0.2s ease;
    }
    .tml-card:hover {
        border-color: #CBD5E1;
        box-shadow: var(--shadow-md);
    }
    .hero-banner {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-top: 4px solid #002B49;
        border-radius: 16px;
        padding: 1.75rem 2rem;
        margin-bottom: 1.5rem;
        box-shadow: var(--shadow-md);
        position: relative;
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #002B49;
        margin: 0.35rem 0 0.5rem 0;
        letter-spacing: -0.6px;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-radius: 9999px;
        padding: 4px 12px;
        font-size: 0.76rem;
        font-weight: 700;
        color: #002B49;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .hero-description {
        font-size: 0.95rem;
        color: #475569;
        line-height: 1.55;
        max-width: 860px;
        margin: 0;
    }
    .hero-feature-item {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 5px 12px;
        font-size: 0.8rem;
        color: #334155;
        font-weight: 600;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #002B49;
        margin: 0;
    }
    .section-subtitle {
        font-size: 0.82rem;
        color: #64748B;
        margin-top: 2px;
    }
    .capacity-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 1.4rem;
        box-shadow: var(--shadow-sm);
        height: 100%;
    }
    .capacity-day-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1rem 1.1rem;
        margin-bottom: 0.85rem;
        transition: border-color 0.2s ease;
    }
    .capacity-day-card:hover {
        border-color: #CBD5E1;
        background: #FFFFFF;
    }
    div[data-testid="stNumberInput"] input {
        background: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }
    div[data-testid="stNumberInput"] input:focus {
        border-color: #0284C7 !important;
        box-shadow: 0 0 0 2px rgba(2, 132, 199, 0.15) !important;
    }
    div[data-testid="stFileUploader"] {
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 14px !important;
        padding: 1.2rem !important;
        box-shadow: var(--shadow-sm) !important;
    }
    div[data-testid="stFileUploaderDropzone"] {
        background: #F8FAFC !important;
        border: 1.5px dashed #CBD5E1 !important;
        border-radius: 10px !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stFileUploaderDropzone"]:hover {
        border-color: #0284C7 !important;
        background: #F0F9FF !important;
    }
    div[data-testid="stFileUploaderDropzone"] button {
        background: #002B49 !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
        border: none !important;
        font-weight: 600 !important;
    }
    div[data-testid="stExpander"] {
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        box-shadow: var(--shadow-sm) !important;
    }
    .styled-divider {
        height: 1px;
        background: #E2E8F0;
        margin: 1.75rem 0;
        border: none;
    }
    """
else:
    # ── DARK MODE STYLING ──
    css_theme = """
    :root {
        --bg-main: #0B132B;
        --card-bg: #151E33;
        --card-border: rgba(255, 255, 255, 0.12);
        --primary-navy: #38BDF8;
        --accent-cyan: #0284C7;
        --text-primary: #F8FAFC;
        --text-secondary: #CBD5E1;
        --text-muted: #94A3B8;
        --badge-bg: rgba(56, 189, 248, 0.15);
        --badge-border: rgba(56, 189, 248, 0.35);
        --shadow-sm: 0 2px 8px rgba(0, 0, 0, 0.25);
        --shadow-md: 0 8px 24px rgba(0, 0, 0, 0.35);
        --shadow-lg: 0 16px 40px rgba(0, 0, 0, 0.45);
    }
    .stApp {
        background-color: var(--bg-main) !important;
        background: radial-gradient(circle at 10% 10%, rgba(56, 189, 248, 0.08), transparent 45%),
                    radial-gradient(circle at 90% 90%, rgba(16, 185, 129, 0.06), transparent 45%),
                    #0B132B !important;
        color: var(--text-primary) !important;
    }
    .tml-card {
        background: var(--card-bg);
        border: 1px solid var(--card-border);
        border-radius: 14px;
        box-shadow: var(--shadow-sm);
        padding: 1.5rem;
        margin-bottom: 1.25rem;
    }
    .hero-banner {
        background: #151E33;
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-top: 4px solid #38BDF8;
        border-radius: 16px;
        padding: 1.75rem 2rem;
        margin-bottom: 1.5rem;
        box-shadow: var(--shadow-md);
        position: relative;
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #38BDF8;
        margin: 0.35rem 0 0.5rem 0;
        letter-spacing: -0.6px;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(56, 189, 248, 0.15);
        border: 1px solid rgba(56, 189, 248, 0.35);
        border-radius: 9999px;
        padding: 4px 12px;
        font-size: 0.76rem;
        font-weight: 700;
        color: #38BDF8;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .hero-description {
        font-size: 0.95rem;
        color: #CBD5E1;
        line-height: 1.55;
        max-width: 860px;
        margin: 0;
    }
    .hero-feature-item {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #1E293B;
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 8px;
        padding: 5px 12px;
        font-size: 0.8rem;
        color: #E2E8F0;
        font-weight: 600;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #F8FAFC;
        margin: 0;
    }
    .section-subtitle {
        font-size: 0.82rem;
        color: #94A3B8;
        margin-top: 2px;
    }
    .capacity-panel {
        background: #151E33;
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 14px;
        padding: 1.4rem;
        box-shadow: var(--shadow-sm);
        height: 100%;
    }
    .capacity-day-card {
        background: #1E293B;
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 10px;
        padding: 1rem 1.1rem;
        margin-bottom: 0.85rem;
    }
    div[data-testid="stNumberInput"] input {
        background: #1E293B !important;
        color: #F8FAFC !important;
        border: 1px solid rgba(255, 255, 255, 0.15) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }
    div[data-testid="stFileUploader"] {
        background: #151E33 !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        border-radius: 14px !important;
        padding: 1.2rem !important;
    }
    div[data-testid="stFileUploaderDropzone"] {
        background: rgba(0, 0, 0, 0.25) !important;
        border: 1.5px dashed rgba(255, 255, 255, 0.2) !important;
        border-radius: 10px !important;
    }
    div[data-testid="stFileUploaderDropzone"] button {
        background: #0284C7 !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
        border: none !important;
        font-weight: 600 !important;
    }
    div[data-testid="stExpander"] {
        background: #151E33 !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        border-radius: 12px !important;
    }
    .styled-divider {
        height: 1px;
        background: rgba(255, 255, 255, 0.12);
        margin: 1.75rem 0;
        border: none;
    }
    """

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    html, body, [class*="css"] {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }}
    {css_theme}

    /* ── Global Typography & Layout Polishes ── */
    h1, h2, h3, h4, h5, h6 {{
        font-family: 'Inter', sans-serif;
    }}
    
    /* ── Primary CTA Button ── */
    div[data-testid="stButton"] > button {{
        background: linear-gradient(135deg, #002B49 0%, #003F6E 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid #002B49 !important;
        border-radius: 10px !important;
        padding: 0.8rem 2rem !important;
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.3px !important;
        box-shadow: 0 4px 12px rgba(0, 43, 73, 0.25) !important;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}
    div[data-testid="stButton"] > button:hover {{
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        border-color: #0284C7 !important;
        box-shadow: 0 6px 18px rgba(2, 132, 199, 0.35) !important;
        transform: translateY(-1px) !important;
    }}

    /* ── Secondary / Download Buttons ── */
    div[data-testid="stDownloadButton"] > button {{
        background: #FFFFFF !important;
        color: #002B49 !important;
        border: 1.5px solid #002B49 !important;
        border-radius: 10px !important;
        padding: 0.65rem 1.4rem !important;
        font-weight: 700 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05) !important;
        transition: all 0.2s ease !important;
    }}
    div[data-testid="stDownloadButton"] > button:hover {{
        background: #002B49 !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 12px rgba(0, 43, 73, 0.2) !important;
    }}

    /* ── Status Dots ── */
    .status-dot {{
        width: 8px;
        height: 8px;
        border-radius: 50%;
        display: inline-block;
        margin-right: 6px;
    }}
    .status-dot.uploaded {{
        background: #10B981;
        box-shadow: 0 0 6px rgba(16, 185, 129, 0.4);
    }}
    .status-dot.pending {{
        background: #F59E0B;
        box-shadow: 0 0 6px rgba(245, 158, 11, 0.4);
    }}

    footer {{ visibility: hidden; }}
    #MainMenu {{ visibility: hidden; }}
    </style>
    """,
    unsafe_allow_html=True
)

# ─────────────────────────────────────────────
# FILE REGISTRY
# ─────────────────────────────────────────────
FILE_LABELS = {
    "vin_plan_tcf1_yesterday": "Yesterday's VIN Plan – TCF 1",
    "vin_plan_tcf2_yesterday": "Yesterday's VIN Plan – TCF 2",
    "vin_list_tcf1_today":     "Today's VIN List – TCF 1",
    "vin_list_tcf2_today":     "Today's VIN List – TCF 2",
    "paint_float_report":      "Paint Float Report",
    "wip_q5":                  "WIP – Q5",
    "nova_wip":                "Nova – WIP",
    "x1_biw_wip":              "X1 BIW WIP",
    "pending_plan_biw":        "Pending Plan for BIW",
    "next_3days_biw_plan":     "Next 3-Days BIW Plan",
}
FILE_KEYS = list(FILE_LABELS.keys())
TOTAL_FILES = len(FILE_KEYS)

# ─────────────────────────────────────────────
# FILE DETECTION & RESOLUTION HELPERS
# ─────────────────────────────────────────────
def detect_file_type(filename: str):
    """
    Classifies a filename into one of the 10 target categories regardless
    of date, month, year, or timestamp variations in the filename.
    """
    if not filename:
        return None
    fn = filename.upper().strip()
    fn_norm = re.sub(r'[^A-Z0-9]+', ' ', fn)
    
    # 1. Paint Float Report (PPC Float Report...)
    if ('PPC' in fn_norm and ('FLOAT' in fn_norm or 'REPORT' in fn_norm)) or ('PAINT' in fn_norm and 'FLOAT' in fn_norm):
        return 'paint_float_report'
        
    # 2. Q5 WIP
    if 'Q5' in fn_norm and ('WIP' in fn_norm or 'FLOAT' in fn_norm):
        return 'wip_q5'
        
    # 3. Nova WIP
    if 'NOVA' in fn_norm and ('WIP' in fn_norm or 'FLOAT' in fn_norm):
        return 'nova_wip'
        
    # 4. X1 / X451 WIP
    if ('X451' in fn_norm or 'X1' in fn_norm) and ('WIP' in fn_norm or 'FLOAT' in fn_norm or 'BIW' in fn_norm):
        return 'x1_biw_wip'
        
    # 5. Pending Plan for BIW
    if 'PENDING' in fn_norm and ('BIW' in fn_norm or 'PLAN' in fn_norm):
        return 'pending_plan_biw'
        
    # 6. Next 3-Days BIW Plan
    has_3day_kw = any(k in fn_norm for k in ['3 DAYS', '3DAYS', '3 DAY', '3DAY', 'NEXT 3', 'NEXT 2', 'DECLARATION'])
    has_biw_kw = ('BIW' in fn_norm or 'PLAN' in fn_norm or 'DECLARATION' in fn_norm)
    if has_3day_kw and has_biw_kw and 'PENDING' not in fn_norm:
        return 'next_3days_biw_plan'
        
    # 7 & 8: Today's VIN Generation Lists
    is_list = ('VEHICLE GENERATION LIST' in fn_norm or 'GENERATION LIST' in fn_norm or 'VIN LIST' in fn_norm or ('VEHICLE' in fn_norm and 'LIST' in fn_norm) or ('VEHICLE' in fn_norm and 'GENERATION' in fn_norm))
    if is_list:
        if 'TCF2' in fn_norm or 'TCF 2' in fn_norm:
            return 'vin_list_tcf2_today'
        else:
            return 'vin_list_tcf1_today'
            
    # 9 & 10: Yesterday's VIN Generation Plans
    is_plan = ('VIN GENERATION PLAN' in fn_norm or 'VIN PLAN' in fn_norm or 'GENERATION PLAN' in fn_norm or ('VIN' in fn_norm and 'PLAN' in fn_norm))
    if is_plan:
        if 'TCF2' in fn_norm or 'TCF 2' in fn_norm:
            return 'vin_plan_tcf2_yesterday'
        elif 'TCF1' in fn_norm or 'TCF 1' in fn_norm:
            return 'vin_plan_tcf1_yesterday'
            
    return None

def get_active_file(key: str):
    """
    Returns the active uploaded file for a given category key:
    Prioritizes manual individual upload slot, falling back to bulk auto-detected file.
    """
    manual = st.session_state.get(key)
    if manual is not None and getattr(manual, 'name', None):
        return manual
    detected = st.session_state.get("bulk_detected_files", {})
    if key in detected and detected[key] is not None:
        return detected[key]
    return None

def upload_card(label: str, key: str, icon: str = "📄") -> object:
    """Renders a cohesive styled upload widget with auto-detection awareness."""
    active_f = get_active_file(key)
    if active_f is not None:
        is_manual = (active_f == st.session_state.get(key))
        badge_text = "Manual Override" if is_manual else "Auto-Detected (Bulk)"
        badge_bg = "#EFF6FF" if is_manual else "#ECFDF5"
        badge_color = "#1D4ED8" if is_manual else "#047857"
        badge_border = "#BFDBFE" if is_manual else "#A7F3D0"
        st.markdown(
            f"""
            <div style="background:{'#F8FAFC' if is_light else '#1E293B'}; border:1px solid {'#E2E8F0' if is_light else 'rgba(255,255,255,0.1)'}; border-radius:8px; padding:6px 12px; margin-bottom:8px; display:flex; align-items:center; justify-content:space-between; gap:10px;">
                <span style="font-size:0.8rem; color:{'#0F172A' if is_light else '#F8FAFC'}; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:600;">
                    📁 {active_f.name}
                </span>
                <span style="font-size:0.7rem; color:{badge_color}; background:{badge_bg}; border:1px solid {badge_border}; padding:2px 8px; border-radius:9999px; font-weight:700; white-space:nowrap;">
                    {badge_text}
                </span>
            </div>
            """,
            unsafe_allow_html=True
        )
    return st.file_uploader(
        label=f"{icon}  {label}",
        type=["csv", "xlsx", "xls", "xlsb", "xlsm"],
        key=key,
    )

# ─────────────────────────────────────────────
# 3. HEADER AND HERO SECTION
# ─────────────────────────────────────────────
st.markdown(
    """
    <div class="hero-banner">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
                <div class="hero-badge">
                    <span style="display:inline-block; width:8px; height:8px; background:#10B981; border-radius:50%;"></span>
                    TATA MOTORS LIMITED – PRODUCTION PLANNING
                </div>
                <h1 class="hero-title">TML VIN Generation &amp; Production Plan</h1>
                <p class="hero-description">
                    Enterprise manufacturing sequence optimizer for <strong>TCF 1</strong> (Punch EV / Petrol) and 
                    <strong>TCF 2</strong> (Harrier / Safari) production lines. Dynamic stage-wise float matching, 
                    WIP prioritization, and automated 3-day rolling capacity balancing.
                </p>
            </div>
        </div>
        <div style="display:flex; gap:12px; flex-wrap:wrap; margin-top:1.25rem; border-top:1px solid #F1F5F9; padding-top:1rem;">
            <div class="hero-feature-item">
                <span style="color:#0284C7; font-size:1rem;">🔄</span> 3-Day Rolling Plan
            </div>
            <div class="hero-feature-item">
                <span style="color:#10B981; font-size:1rem;">⚡</span> Stage-wise Float Matching
            </div>
            <div class="hero-feature-item">
                <span style="color:#DC2626; font-size:1rem;">🛡️</span> Pune Plant Exclusions Active
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# 4. BULK FILE UPLOAD SECTION
# ─────────────────────────────────────────────
if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0

bulk_key = f"bulk_file_uploader_{st.session_state['uploader_version']}"

st.markdown(
    """
    <div class="tml-card">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 0.75rem; flex-wrap:wrap; gap:10px;">
            <div>
                <div style="display:flex; align-items:center; gap:8px;">
                    <span style="font-size:1.25rem;">📂</span>
                    <h3 class="section-title">Bulk File Upload &amp; Automatic Name Detection</h3>
                </div>
                <p class="section-subtitle">
                    Drag and drop all 10 planning, float, and WIP files at once. Names are automatically identified regardless of date or month.
                </p>
            </div>
            <div style="font-size:0.75rem; color:#64748B; background:#F8FAFC; border:1px solid #E2E8F0; padding:4px 10px; border-radius:6px; font-weight:600;">
                Supported: .xlsx, .xls, .xlsm, .csv, .xlsb (Max 200MB)
            </div>
        </div>
    """,
    unsafe_allow_html=True,
)

bulk_files = st.file_uploader(
    "Drop all 10 files together here (or click to browse)",
    type=["csv", "xlsx", "xls", "xlsb", "xlsm"],
    accept_multiple_files=True,
    key=bulk_key,
    help="Upload all required planning, float, and WIP files simultaneously. The engine will automatically classify each file."
)

if "bulk_detected_files" not in st.session_state:
    st.session_state["bulk_detected_files"] = {}

if bulk_files:
    detected_mapping = {}
    unrecognized = []
    
    for f in bulk_files:
        override_val = st.session_state.get(f"override_{f.name}")
        if override_val and override_val in FILE_LABELS:
            detected_mapping[override_val] = f
        else:
            cat = detect_file_type(f.name)
            if cat:
                detected_mapping[cat] = f
            else:
                unrecognized.append(f)
                
    st.session_state["bulk_detected_files"] = detected_mapping

    # Render mapping review table
    st.markdown("<h4 style='color:#002B49; font-size:0.95rem; font-weight:700; margin-top:1rem; margin-bottom:0.5rem;'>🎯 File Auto-Detection &amp; Mapping Matrix</h4>", unsafe_allow_html=True)
    
    head_col1, head_col2, head_col3 = st.columns([3, 3, 2])
    with head_col1:
        st.markdown("<span style='font-size:0.72rem; color:#64748B; font-weight:700; text-transform:uppercase;'>Uploaded File</span>", unsafe_allow_html=True)
    with head_col2:
        st.markdown("<span style='font-size:0.72rem; color:#64748B; font-weight:700; text-transform:uppercase;'>Auto-Detected Category</span>", unsafe_allow_html=True)
    with head_col3:
        st.markdown("<span style='font-size:0.72rem; color:#64748B; font-weight:700; text-transform:uppercase;'>Override Category</span>", unsafe_allow_html=True)
        
    for f in bulk_files:
        matched_cat = None
        for cat_k, file_obj in st.session_state["bulk_detected_files"].items():
            if file_obj.name == f.name:
                matched_cat = cat_k
                break
                
        r_c1, r_c2, r_c3 = st.columns([3, 3, 2])
        with r_c1:
            st.markdown(f"<div style='font-size:0.82rem; color:#0F172A; font-weight:500; padding:6px 0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{f.name}'>📄 {f.name}</div>", unsafe_allow_html=True)
        with r_c2:
            if matched_cat:
                st.markdown(f"<div style='padding:4px 0;'><span style='font-size:0.75rem; color:#065F46; background:#ECFDF5; border:1px solid #A7F3D0; padding:3px 10px; border-radius:9999px; font-weight:700;'>✓ {FILE_LABELS[matched_cat]}</span></div>", unsafe_allow_html=True)
            else:
                st.markdown("<div style='padding:4px 0;'><span style='font-size:0.75rem; color:#991B1B; background:#FEF2F2; border:1px solid #FECACA; padding:3px 10px; border-radius:9999px; font-weight:700;'>⚠️ Unrecognized</span></div>", unsafe_allow_html=True)
        with r_c3:
            options = ["(Auto-detect)"] + list(FILE_LABELS.keys())
            curr_index = 0
            if st.session_state.get(f"override_{f.name}"):
                curr_index = options.index(st.session_state[f"override_{f.name}"])
            elif matched_cat:
                curr_index = options.index(matched_cat)
                
            chosen = st.selectbox(
                "Override Category",
                options=options,
                index=curr_index,
                format_func=lambda x: "(Auto-detect)" if x == "(Auto-detect)" else FILE_LABELS[x],
                key=f"select_override_{f.name}",
                label_visibility="collapsed"
            )
            if chosen != "(Auto-detect)" and chosen != matched_cat:
                st.session_state[f"override_{f.name}"] = chosen
                st.session_state["bulk_detected_files"][chosen] = f
                st.rerun()
            elif chosen == "(Auto-detect)" and st.session_state.get(f"override_{f.name}"):
                del st.session_state[f"override_{f.name}"]
                st.rerun()

    # Clear button
    c_col1, c_col2 = st.columns([5, 1])
    with c_col2:
        if st.button("🗑️ Clear Bulk", key="btn_clear_bulk", use_container_width=True):
            st.session_state["bulk_detected_files"] = {}
            st.session_state["uploader_version"] += 1
            for k in list(st.session_state.keys()):
                if k.startswith("override_"):
                    del st.session_state[k]
            st.rerun()
else:
    st.session_state["bulk_detected_files"] = {}

st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# LIVE UPLOAD PROGRESS INDICATOR
# ─────────────────────────────────────────────
uploaded_count = sum(
    1 for k in FILE_KEYS if get_active_file(k) is not None
)
pct = int((uploaded_count / TOTAL_FILES) * 100)

col_prog_1, col_prog_2 = st.columns([1, 5])

with col_prog_1:
    st.markdown(
        f"""
        <div class="tml-card" style="text-align:center; padding: 1rem 0.5rem; margin-bottom:0;">
            <div style="font-size:2.2rem; font-weight:800; color:#002B49; line-height:1;">{uploaded_count}/{TOTAL_FILES}</div>
            <div style="font-size:0.75rem; color:#64748B; margin-top:4px; font-weight:700; text-transform:uppercase;">Files Verified</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_prog_2:
    st.progress(pct / 100, text=f"Pipeline Readiness — {pct}% files prepared")
    # Mini status checklist
    status_html = ""
    for k, label in FILE_LABELS.items():
        dot_cls = "uploaded" if get_active_file(k) is not None else "pending"
        status_html += (
            f'<span style="margin-right:12px; font-size:0.75rem; color:#475569; font-weight:500;">'
            f'<span class="status-dot {dot_cls}"></span>{label}</span>'
        )
    st.markdown(
        f'<div style="margin-top:6px; line-height:2.0;">{status_html}</div>',
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────
# 5. FILE UPLOAD WORKFLOW GROUPS (COLLAPSIBLE / OVERRIDES)
# ─────────────────────────────────────────────
st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)

with st.expander("📂 Individual File Uploaders (Manual Upload & Single File Overrides)", expanded=(uploaded_count == 0)):
    # GROUP 1 – Yesterday's VIN Generation Plans
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:8px; margin: 1rem 0 0.8rem 0;">
            <span style="background:#EFF6FF; color:#1D4ED8; padding:4px 8px; border-radius:6px; font-size:0.85rem; font-weight:700;">Group 1</span>
            <div>
                <p style="margin:0; font-size:0.95rem; font-weight:700; color:#0F172A;">Yesterday's VIN Generation Plans</p>
                <p style="margin:0; font-size:0.75rem; color:#64748B;">Baseline plans from the previous production day for TCF 1 &amp; TCF 2</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    g1_col1, g1_col2 = st.columns(2, gap="large")
    with g1_col1:
        upload_card(
            "Yesterday's VIN Generation Plan — TCF 1",
            key="vin_plan_tcf1_yesterday",
            icon="📄",
        )
    with g1_col2:
        upload_card(
            "Yesterday's VIN Generation Plan — TCF 2",
            key="vin_plan_tcf2_yesterday",
            icon="📄",
        )

    # GROUP 2 – Today's VIN Generation Lists
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:8px; margin: 1.5rem 0 0.8rem 0;">
            <span style="background:#ECFDF5; color:#047857; padding:4px 8px; border-radius:6px; font-size:0.85rem; font-weight:700;">Group 2</span>
            <div>
                <p style="margin:0; font-size:0.95rem; font-weight:700; color:#0F172A;">Today's VIN Generation Lists</p>
                <p style="margin:0; font-size:0.75rem; color:#64748B;">Current-day VIN lists for TCF 1 &amp; TCF 2 production tracks</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    g2_col1, g2_col2 = st.columns(2, gap="large")
    with g2_col1:
        upload_card(
            "Today's TCF 1 VIN Generation List",
            key="vin_list_tcf1_today",
            icon="📝",
        )
    with g2_col2:
        upload_card(
            "Today's TCF 2 VIN Generation List",
            key="vin_list_tcf2_today",
            icon="📝",
        )

    # GROUP 3 – Reports & Work-In-Progress Data
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:8px; margin: 1.5rem 0 0.8rem 0;">
            <span style="background:#FFFBEB; color:#B45309; padding:4px 8px; border-radius:6px; font-size:0.85rem; font-weight:700;">Group 3</span>
            <div>
                <p style="margin:0; font-size:0.95rem; font-weight:700; color:#0F172A;">Reports &amp; Work-In-Progress Data</p>
                <p style="margin:0; font-size:0.75rem; color:#64748B;">Paint float, WIP trackers, BIW plan data and stage-wise reports</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    g3_row1_col1, g3_row1_col2 = st.columns(2, gap="large")
    with g3_row1_col1:
        upload_card("Paint Float Report", key="paint_float_report", icon="🎨")
    with g3_row1_col2:
        upload_card("WIP — Q5", key="wip_q5", icon="📈")

    g3_row2_col1, g3_row2_col2 = st.columns(2, gap="large")
    with g3_row2_col1:
        upload_card("Nova — WIP", key="nova_wip", icon="🔧")
    with g3_row2_col2:
        upload_card("Pending Plan for BIW", key="pending_plan_biw", icon="🏗️")

    g3_row3_col1, g3_row3_col2 = st.columns(2, gap="large")
    with g3_row3_col1:
        upload_card("X1 BIW WIP", key="x1_biw_wip", icon="🏭")
    with g3_row3_col2:
        upload_card("Next 3-Days BIW Plan", key="next_3days_biw_plan", icon="📋")


# ─────────────────────────────────────────────
# 6. PHASE 3 – SHIFT PRODUCTION TARGETS (MANUAL ENTRY)
# ─────────────────────────────────────────────
st.markdown(
    """
    <div class="tml-card" style="border-left: 4px solid #0284C7;">
        <div style="display:flex; align-items:center; gap:8px; margin-bottom:0.4rem;">
            <span style="font-size:1.2rem;">⏱️</span>
            <h4 style="margin:0; font-size:1.05rem; font-weight:700; color:#002B49;">
                Phase 3 — Shift Production Targets (Manual Entry)
            </h4>
        </div>
        <p style="margin:0 0 1rem 0; font-size:0.85rem; color:#64748B; line-height:1.4;">
            Specify the <strong>Expected VIN B &amp; C shift tentative production quantities</strong> for TCF 1 and TCF 2 tracks before initializing the 3-day sequence.
        </p>
    """,
    unsafe_allow_html=True,
)

qty_col1, qty_col2 = st.columns(2, gap="large")

with qty_col1:
    st.number_input(
        "Expected TCF 1 B & C Shift Tentative Qty",
        min_value=0,
        max_value=1000,
        value=500,
        step=1,
        key="expected_qty_tcf1",
        help="Enter or select the tentative production quantity for TCF 1 B & C shift."
    )

with qty_col2:
    st.number_input(
        "Expected TCF 2 B & C Shift Tentative Qty",
        min_value=0,
        max_value=1000,
        value=100,
        step=1,
        key="expected_qty_tcf2",
        help="Enter or select the tentative production quantity for TCF 2 B & C shift."
    )

st.markdown("</div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 7. PRODUCTION LINE CAPACITIES
# ─────────────────────────────────────────────
st.markdown(
    """
    <div style="margin: 1.5rem 0 1rem 0;">
        <div style="display:flex; align-items:center; gap:8px;">
            <span style="font-size:1.25rem;">📅</span>
            <h3 class="section-title">Production Line Capacities</h3>
        </div>
        <p class="section-subtitle">Configure 3-day planning capacity limits and powertrain mix constraints for TCF-1 and TCF-2</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Dynamically calculate planning dates
baseline_date = None
yest_tcf1 = get_active_file("vin_plan_tcf1_yesterday")
if yest_tcf1 is not None:
    try:
        wb_temp = openpyxl.load_workbook(io.BytesIO(yest_tcf1.getvalue()), read_only=True, data_only=True)
        if 'Plan Summary' in wb_temp.sheetnames:
            ws_temp = wb_temp['Plan Summary']
            val_h2 = ws_temp.cell(row=2, column=8).value
            if isinstance(val_h2, (datetime.datetime, datetime.date)):
                baseline_date = val_h2
    except Exception:
        pass

if baseline_date is None:
    baseline_date = datetime.date.today()
elif isinstance(baseline_date, datetime.datetime):
    baseline_date = baseline_date.date()

def get_next_planning_days(start_date, count=3):
    days = []
    curr = start_date
    while len(days) < count:
        curr += datetime.timedelta(days=1)
        if curr.weekday() == 6:  # Sunday exclusion
            continue
        days.append(curr)
    return days

planning_days = get_next_planning_days(baseline_date, 3)
d1, d2, d3 = planning_days

d1_str = d1.strftime('%Y-%m-%d (%A)')
d2_str = d2.strftime('%Y-%m-%d (%A)')
d3_str = d3.strftime('%Y-%m-%d (%A)')
day_dates = [d1_str, d2_str, d3_str]

cap_col1, cap_col2 = st.columns(2, gap="large")

# ── TCF-1 PANEL (Punch / Altroz) ──
with cap_col1:
    st.markdown(
        """
        <div class="capacity-panel" style="border-top: 4px solid #0284C7;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:1.1rem;">
                <div style="display:flex; align-items:center; gap:8px;">
                    <span style="background:#E0F2FE; color:#0369A1; font-weight:800; font-size:0.75rem; padding:3px 10px; border-radius:6px; border:1px solid #BAE6FD;">
                        TCF-1 TRACK
                    </span>
                    <h4 style="margin:0; font-size:1.1rem; font-weight:700; color:#002B49;">
                        Punch / Altroz Daily Targets
                    </h4>
                </div>
            </div>
        """,
        unsafe_allow_html=True
    )
    
    tcf1_caps = []
    tcf1_ev = []
    tcf1_cng = []
    
    for idx, (day_num, def_total, def_ev, def_cng) in enumerate([
        ("1", 900, 160, 350),
        ("2", 900, 160, 350),
        ("3", 900, 160, 350)
    ]):
        day_date_label = day_dates[idx]
        st.markdown(
            f"""
            <div class="capacity-day-card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.6rem;">
                    <span style="font-weight:700; font-size:0.85rem; color:#002B49;">
                        🗓️ Day {day_num}
                    </span>
                    <span style="font-size:0.75rem; color:#64748B; font-weight:500;">
                        {day_date_label}
                    </span>
                </div>
            """,
            unsafe_allow_html=True
        )
        sc1, sc2, sc3, sc4 = st.columns([1.1, 1.1, 1.1, 1.3])
        with sc1:
            total_cap = st.number_input("Total Target", min_value=0, max_value=2000, value=def_total, step=50, key=f"tcf1_cap_d{day_num}")
        with sc2:
            ev_cap = st.number_input("Punch EV Limit", min_value=0, max_value=2000, value=def_ev, step=10, key=f"tcf1_ev_d{day_num}")
        with sc3:
            cng_cap = st.number_input("Punch CNG Limit", min_value=0, max_value=2000, value=def_cng, step=10, key=f"tcf1_cng_d{day_num}")
        with sc4:
            rem_petrol = total_cap - ev_cap - cng_cap
            st.markdown("<div style='font-size:0.72rem; color:#64748B; font-weight:700; text-transform:uppercase; margin-bottom:4px;'>Petrol Balance</div>", unsafe_allow_html=True)
            if rem_petrol < 0:
                st.markdown(
                    f"<div style='background:#FEF2F2; border:1px solid #FECACA; color:#B91C1C; padding:7px 10px; border-radius:8px; font-weight:700; font-size:0.85rem; text-align:center;'>⚠️ Deficit: {rem_petrol}</div>",
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    f"<div style='background:#F1F5F9; border:1px solid #CBD5E1; color:#0F172A; padding:7px 10px; border-radius:8px; font-weight:700; font-size:0.85rem; text-align:center;'>⛽ Petrol: {rem_petrol}</div>",
                    unsafe_allow_html=True
                )
        st.markdown("</div>", unsafe_allow_html=True)
        tcf1_caps.append(total_cap)
        tcf1_ev.append(ev_cap)
        tcf1_cng.append(cng_cap)
        
    st.markdown("</div>", unsafe_allow_html=True)
    st.session_state["tcf1_capacities"] = tcf1_caps
    st.session_state["tcf1_ev"] = tcf1_ev
    st.session_state["tcf1_cng"] = tcf1_cng

# ── TCF-2 PANEL (Harrier / Safari) ──
with cap_col2:
    st.markdown(
        """
        <div class="capacity-panel" style="border-top: 4px solid #1E3A8A;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:1.1rem;">
                <div style="display:flex; align-items:center; gap:8px;">
                    <span style="background:#EDE9FE; color:#5B21B6; font-weight:800; font-size:0.75rem; padding:3px 10px; border-radius:6px; border:1px solid #DDD6FE;">
                        TCF-2 TRACK
                    </span>
                    <h4 style="margin:0; font-size:1.1rem; font-weight:700; color:#002B49;">
                        Harrier / Safari Daily Targets
                    </h4>
                </div>
            </div>
        """,
        unsafe_allow_html=True
    )
    
    tcf2_caps = []
    tcf2_eturna = []
    tcf2_tgdi = []
    
    for idx, (day_num, def_total, def_ev, def_tgdi) in enumerate([
        ("1", 250, 120, 40),
        ("2", 250, 120, 40),
        ("3", 250, 120, 40)
    ]):
        day_date_label = day_dates[idx]
        st.markdown(
            f"""
            <div class="capacity-day-card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.6rem;">
                    <span style="font-weight:700; font-size:0.85rem; color:#002B49;">
                        🗓️ Day {day_num}
                    </span>
                    <span style="font-size:0.75rem; color:#64748B; font-weight:500;">
                        {day_date_label}
                    </span>
                </div>
            """,
            unsafe_allow_html=True
        )
        sc1, sc2, sc3, sc4 = st.columns([1.1, 1.1, 1.1, 1.3])
        with sc1:
            total_cap = st.number_input("Total Target", min_value=0, max_value=2000, value=def_total, step=50, key=f"tcf2_cap_d{day_num}")
        with sc2:
            et_cap = st.number_input("Eturna EV Limit", min_value=0, max_value=2000, value=def_ev, step=10, key=f"tcf2_et_d{day_num}")
        with sc3:
            tg_cap = st.number_input("TGDI Petrol Limit", min_value=0, max_value=2000, value=def_tgdi, step=10, key=f"tcf2_tg_d{day_num}")
        with sc4:
            rem_diesel = total_cap - et_cap - tg_cap
            st.markdown("<div style='font-size:0.72rem; color:#64748B; font-weight:700; text-transform:uppercase; margin-bottom:4px;'>Diesel Balance</div>", unsafe_allow_html=True)
            if rem_diesel < 0:
                st.markdown(
                    f"<div style='background:#FEF2F2; border:1px solid #FECACA; color:#B91C1C; padding:7px 10px; border-radius:8px; font-weight:700; font-size:0.85rem; text-align:center;'>⚠️ Deficit: {rem_diesel}</div>",
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    f"<div style='background:#F1F5F9; border:1px solid #CBD5E1; color:#0F172A; padding:7px 10px; border-radius:8px; font-weight:700; font-size:0.85rem; text-align:center;'>⛽ Diesel: {rem_diesel}</div>",
                    unsafe_allow_html=True
                )
        st.markdown("</div>", unsafe_allow_html=True)
        tcf2_caps.append(total_cap)
        tcf2_eturna.append(et_cap)
        tcf2_tgdi.append(tg_cap)
        
    st.markdown("</div>", unsafe_allow_html=True)
    st.session_state["tcf2_capacities"] = tcf2_caps
    st.session_state["tcf2_eturna"] = tcf2_eturna
    st.session_state["tcf2_tgdi"] = tcf2_tgdi

st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 8. GENERATE 3-DAY PLAN BUTTON (PRIMARY CTA)
# ─────────────────────────────────────────────
btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])
with btn_col2:
    generate_clicked = st.button(
        "⚙️  Generate 3-Day Production Plan",
        key="generate_btn",
        use_container_width=True,
    )

if generate_clicked:
    required_keys = [k for k in FILE_KEYS if k not in ("pending_plan_biw", "next_3days_biw_plan")]
    missing = [
        FILE_LABELS[k]
        for k in required_keys
        if get_active_file(k) is None
    ]
    if missing:
        missing_list = "\n".join(f"- {m}" for m in missing)
        st.warning(
            f"**⚠️  Missing {len(missing)} required file(s). Please upload:**\n\n{missing_list}",
            icon="⚠️",
        )
    else:
        with st.spinner("Processing stage-wise floats and sequencing 3-day production plans..."):
            try:
                yest_tcf1 = get_active_file("vin_plan_tcf1_yesterday")
                yest_tcf2 = get_active_file("vin_plan_tcf2_yesterday")
                today_list_tcf1 = get_active_file("vin_list_tcf1_today")
                today_list_tcf2 = get_active_file("vin_list_tcf2_today")
                paint_float = get_active_file("paint_float_report")
                wip_q5 = get_active_file("wip_q5")
                nova_wip = get_active_file("nova_wip")
                x1_biw_wip = get_active_file("x1_biw_wip")
                
                next_3days_biw_plan = get_active_file("next_3days_biw_plan")
                pending_plan_biw = get_active_file("pending_plan_biw")
                
                expected_qty_tcf1 = st.session_state.get("expected_qty_tcf1", 500)
                expected_qty_tcf2 = st.session_state.get("expected_qty_tcf2", 100)
                
                # 1. Process TCF-1 3-Day Plan
                tcf1_wb = initialize_tcf1_workbook(yest_tcf1)
                tcf1_wb = update_today_vin_generation(tcf1_wb, today_list_tcf1, "TCF1")
                
                tcf1_wip_files = []
                if nova_wip is not None:
                    tcf1_wip_files.append(nova_wip)
                if x1_biw_wip is not None:
                    tcf1_wip_files.append(x1_biw_wip)
                    
                tcf1_capacities = st.session_state.get("tcf1_capacities", [900, 900, 900])
                tcf1_ev = st.session_state.get("tcf1_ev", [160, 160, 160])
                tcf1_cng = st.session_state.get("tcf1_cng", [350, 350, 350])
                tcf1_wb = update_paint_float_data(
                    tcf1_wb, paint_float, "TCF1", 
                    expected_qty=expected_qty_tcf1,
                    wip_files=tcf1_wip_files,
                    yest_plan_file_or_stream=yest_tcf1,
                    pending_plan_biw=pending_plan_biw,
                    next_3days_biw_plan=next_3days_biw_plan,
                    daily_capacities=tcf1_capacities,
                    sub_limits_1=tcf1_ev,
                    sub_limits_2=tcf1_cng
                )
                
                tcf1_out = io.BytesIO()
                tcf1_wb.save(tcf1_out)
                tcf1_out.seek(0)
                
                # 2. Process TCF-2 3-Day Plan
                tcf2_wb = initialize_tcf2_workbook(yest_tcf2)
                tcf2_wb = update_today_vin_generation(tcf2_wb, today_list_tcf2, "TCF2")
                
                tcf2_wip_files = []
                if wip_q5 is not None:
                    tcf2_wip_files.append(wip_q5)
                    
                tcf2_capacities = st.session_state.get("tcf2_capacities", [250, 250, 250])
                tcf2_eturna = st.session_state.get("tcf2_eturna", [120, 120, 120])
                tcf2_tgdi = st.session_state.get("tcf2_tgdi", [40, 40, 40])
                tcf2_wb = update_paint_float_data(
                    tcf2_wb, paint_float, "TCF2", 
                    expected_qty=expected_qty_tcf2,
                    wip_files=tcf2_wip_files,
                    yest_plan_file_or_stream=yest_tcf2,
                    pending_plan_biw=pending_plan_biw,
                    next_3days_biw_plan=next_3days_biw_plan,
                    daily_capacities=tcf2_capacities,
                    sub_limits_1=tcf2_eturna,
                    sub_limits_2=tcf2_tgdi
                )
                
                tcf2_out = io.BytesIO()
                tcf2_wb.save(tcf2_out)
                tcf2_out.seek(0)
                
                # Persist output file bytes and summary details in st.session_state
                st.session_state["plans_generated"] = True
                st.session_state["tcf1_out_bytes"] = tcf1_out.getvalue()
                st.session_state["tcf2_out_bytes"] = tcf2_out.getvalue()
                st.session_state["tcf1_summary"] = tcf1_wb.summary_counts
                st.session_state["tcf2_summary"] = tcf2_wb.summary_counts
                
                st.toast("3-Day production plans generated successfully!", icon="✅")
            except Exception as e:
                st.session_state["plans_generated"] = False
                st.error(f"Error generating plans: {e}")

# ─────────────────────────────────────────────
# OUTPUT DASHBOARDS (AFTER GENERATION)
# ─────────────────────────────────────────────
if st.session_state.get("plans_generated", False):
    tcf1_caps = st.session_state.get("tcf1_capacities", [900, 900, 900])
    tcf2_caps = st.session_state.get("tcf2_capacities", [250, 250, 250])
    tcf1_cap_strs = [f"{cap} (Holiday)" if cap == 0 else str(cap) for cap in tcf1_caps]
    tcf2_cap_strs = [f"{cap} (Holiday)" if cap == 0 else str(cap) for cap in tcf2_caps]
    
    st.markdown(
        f"""
        <div class="tml-card" style="border-left: 5px solid #10B981; background:{'#F0FDF4' if is_light else '#142922'};">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
                <span style="font-size:1.3rem;">✅</span>
                <h4 style="margin:0; font-size:1.15rem; font-weight:800; color:{'#065F46' if is_light else '#34D399'};">
                    Production Plans Generated Successfully
                </h4>
            </div>
            <p style="margin:0; font-size:0.88rem; color:{'#1E293B' if is_light else '#E2E8F0'}; line-height:1.5;">
                • <strong>TCF-1 (Punch/Altroz):</strong> Day 1: <code>{tcf1_cap_strs[0]}</code> | Day 2: <code>{tcf1_cap_strs[1]}</code> | Day 3: <code>{tcf1_cap_strs[2]}</code><br/>
                • <strong>TCF-2 (Harrier/Safari):</strong> Day 1: <code>{tcf2_cap_strs[0]}</code> | Day 2: <code>{tcf2_cap_strs[1]}</code> | Day 3: <code>{tcf2_cap_strs[2]}</code>
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Download buttons in enterprise cards
    dl_col1, dl_col2 = st.columns(2, gap="large")
    with dl_col1:
        st.download_button(
            label="📥 Download TCF-1 (Punch/Altroz) 3-Day Plan",
            data=st.session_state["tcf1_out_bytes"],
            file_name="TCF-1_3-Day_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_tcf1_3day",
            use_container_width=True
        )
    with dl_col2:
        st.download_button(
            label="📥 Download TCF-2 (Harrier/Safari) 3-Day Plan",
            data=st.session_state["tcf2_out_bytes"],
            file_name="TCF-2_3-Day_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_tcf2_3day",
            use_container_width=True
        )
        
    st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
    
    # Color-coded source breakdown
    t1 = st.session_state["tcf1_summary"]
    t2 = st.session_state["tcf2_summary"]
    
    st.markdown(
        """
        <div class="tml-card">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:1.2rem;">
                <span style="font-size:1.2rem;">🎨</span>
                <h4 class="section-title">Planned Vehicle Count by Source Type (Color-Coded)</h4>
            </div>
        """,
        unsafe_allow_html=True
    )
    
    sc_c1, sc_c2 = st.columns(2, gap="large")
    with sc_c1:
        st.markdown(
            f"""
            <div style="background:{'#F8FAFC' if is_light else '#1E293B'}; border:1px solid {'#E2E8F0' if is_light else 'rgba(255,255,255,0.1)'}; border-radius:10px; padding:1.2rem;">
                <h5 style="margin:0 0 0.8rem 0; color:#0284C7; font-size:0.95rem; font-weight:700;">TCF-1 (Punch/Altroz)</h5>
                <div style="display:flex; flex-direction:column; gap:10px; font-size:0.86rem;">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#E2EFDA; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Paint Floor Float (Light Green):</span>
                        <strong>{t1['colors']['light_green']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#DDEBF7; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>WIP in BIW (Light Blue):</span>
                        <strong>{t1['colors']['light_blue']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FFF2CC; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>BIW Plan Loaded (Production Pending) (Light Yellow):</span>
                        <strong>{t1['colors']['light_yellow']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FCE4D6; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Next 3-Days BIW Plan (Light Pink):</span>
                        <strong>{t1['colors']['light_pink']}</strong>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with sc_c2:
        st.markdown(
            f"""
            <div style="background:{'#F8FAFC' if is_light else '#1E293B'}; border:1px solid {'#E2E8F0' if is_light else 'rgba(255,255,255,0.1)'}; border-radius:10px; padding:1.2rem;">
                <h5 style="margin:0 0 0.8rem 0; color:#059669; font-size:0.95rem; font-weight:700;">TCF-2 (Harrier/Safari)</h5>
                <div style="display:flex; flex-direction:column; gap:10px; font-size:0.86rem;">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#E2EFDA; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Paint Floor Float (Light Green):</span>
                        <strong>{t2['colors']['light_green']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#DDEBF7; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>WIP in BIW (Light Blue):</span>
                        <strong>{t2['colors']['light_blue']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FFF2CC; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>BIW Plan Loaded (Production Pending) (Light Yellow):</span>
                        <strong>{t2['colors']['light_yellow']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FCE4D6; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Next 3-Days BIW Plan (Light Pink):</span>
                        <strong>{t2['colors']['light_pink']}</strong>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    st.markdown("</div>", unsafe_allow_html=True)
    
    # 4 Model Summary Cards
    st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:8px; margin-bottom:1rem;">
            <span style="font-size:1.2rem;">📊</span>
            <h4 class="section-title">Constrained Powertrain &amp; Model Summaries</h4>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    tcf1_ev = st.session_state.get("tcf1_ev", [160, 160, 160])
    tcf1_cng = st.session_state.get("tcf1_cng", [350, 350, 350])
    tcf2_eturna = st.session_state.get("tcf2_eturna", [120, 120, 120])
    tcf2_tgdi = st.session_state.get("tcf2_tgdi", [40, 40, 40])

    sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4, gap="large")
    
    with sum_col1:
        st.markdown(
            f"""
            <div class="tml-card" style="border-top: 4px solid #0284C7; text-align:center;">
                <div style="font-size:2rem; margin-bottom:0.2rem;">⚡</div>
                <h4 style="color:#0284C7; margin:0 0 0.1rem 0; font-size:1.15rem; font-weight:700;">NOVA (Punch EV)</h4>
                <p style="color:#64748B; font-size:0.75rem; margin:0 0 1rem 0; text-transform:uppercase; letter-spacing:0.06em; font-weight:600;">Daily Targets</p>
                <div style="background:{'#F8FAFC' if is_light else '#1E293B'}; border-radius:8px; padding:0.8rem; font-size:0.84rem; text-align:left;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem; border-bottom:1px solid {'#E2E8F0' if is_light else 'rgba(255,255,255,0.1)'}; padding-bottom:0.4rem;">
                        <span>B&C Drops Today:</span> <strong style="color:#0284C7;">{t1['picked']['nova']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Today (Day 1):</span> <span><strong>{t1['day1']['nova']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf1_ev[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t1['day2']['nova']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf1_ev[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t1['day3']['nova']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf1_ev[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
    with sum_col2:
        st.markdown(
            f"""
            <div class="tml-card" style="border-top: 4px solid #EA580C; text-align:center;">
                <div style="font-size:2rem; margin-bottom:0.2rem;">🔥</div>
                <h4 style="color:#EA580C; margin:0 0 0.1rem 0; font-size:1.15rem; font-weight:700;">Punch CNG</h4>
                <p style="color:#64748B; font-size:0.75rem; margin:0 0 1rem 0; text-transform:uppercase; letter-spacing:0.06em; font-weight:600;">Daily Targets</p>
                <div style="background:{'#F8FAFC' if is_light else '#1E293B'}; border-radius:8px; padding:0.8rem; font-size:0.84rem; text-align:left;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem; border-bottom:1px solid {'#E2E8F0' if is_light else 'rgba(255,255,255,0.1)'}; padding-bottom:0.4rem;">
                        <span>B&C Drops Today:</span> <strong style="color:#EA580C;">{t1['picked']['cng']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Today (Day 1):</span> <span><strong>{t1['day1']['cng']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf1_cng[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t1['day2']['cng']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf1_cng[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t1['day3']['cng']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf1_cng[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
    with sum_col3:
        st.markdown(
            f"""
            <div class="tml-card" style="border-top: 4px solid #059669; text-align:center;">
                <div style="font-size:2rem; margin-bottom:0.2rem;">🌱</div>
                <h4 style="color:#059669; margin:0 0 0.1rem 0; font-size:1.15rem; font-weight:700;">Eturna (Harrier EV)</h4>
                <p style="color:#64748B; font-size:0.75rem; margin:0 0 1rem 0; text-transform:uppercase; letter-spacing:0.06em; font-weight:600;">Daily Targets</p>
                <div style="background:{'#F8FAFC' if is_light else '#1E293B'}; border-radius:8px; padding:0.8rem; font-size:0.84rem; text-align:left;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem; border-bottom:1px solid {'#E2E8F0' if is_light else 'rgba(255,255,255,0.1)'}; padding-bottom:0.4rem;">
                        <span>B&C Drops Today:</span> <strong style="color:#059669;">{t2['picked']['eturna']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Today (Day 1):</span> <span><strong>{t2['day1']['eturna']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf2_eturna[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t2['day2']['eturna']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf2_eturna[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t2['day3']['eturna']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf2_eturna[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with sum_col4:
        st.markdown(
            f"""
            <div class="tml-card" style="border-top: 4px solid #7C3AED; text-align:center;">
                <div style="font-size:2rem; margin-bottom:0.2rem;">🏎️</div>
                <h4 style="color:#7C3AED; margin:0 0 0.1rem 0; font-size:1.15rem; font-weight:700;">TGDI (Petrol)</h4>
                <p style="color:#64748B; font-size:0.75rem; margin:0 0 1rem 0; text-transform:uppercase; letter-spacing:0.06em; font-weight:600;">Daily Targets</p>
                <div style="background:{'#F8FAFC' if is_light else '#1E293B'}; border-radius:8px; padding:0.8rem; font-size:0.84rem; text-align:left;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem; border-bottom:1px solid {'#E2E8F0' if is_light else 'rgba(255,255,255,0.1)'}; padding-bottom:0.4rem;">
                        <span>B&C Drops Today:</span> <strong style="color:#7C3AED;">{t2['picked']['tgdi']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Today (Day 1):</span> <span><strong>{t2['day1']['tgdi']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf2_tgdi[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.3rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t2['day2']['tgdi']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf2_tgdi[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t2['day3']['tgdi']}</strong> <span style="font-size:0.75rem; color:#64748B;">/ {tcf2_tgdi[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # Shortage & Production Remarks
    remarks = []
    t1_targets = t1.get('targets', {})
    for day_key, day_name in [('day1', 'Day 1 (Today)'), ('day2', 'Day 2'), ('day3', 'Day 3')]:
        t_val = t1_targets.get(day_key, {})
        p_val = t1.get(day_key, {})
        day_cap = st.session_state.get("tcf1_capacities", [900, 900, 900])[int(day_key[-1]) - 1]
        if day_cap > 0:
            for variant, label in [('nova', 'Punch EV'), ('cng', 'Punch CNG')]:
                target = t_val.get(variant, 0)
                planned = p_val.get(variant, 0)
                if planned < target:
                    shortage = target - planned
                    remarks.append(f"TCF-1 {day_name} - **{label}**: Target was `{target}`, but only `{planned}` vehicles were available (Shortage of `{shortage}`).")
                    
    t2_targets = t2.get('targets', {})
    for day_key, day_name in [('day1', 'Day 1 (Today)'), ('day2', 'Day 2'), ('day3', 'Day 3')]:
        t_val = t2_targets.get(day_key, {})
        p_val = t2.get(day_key, {})
        day_cap = st.session_state.get("tcf2_capacities", [250, 250, 250])[int(day_key[-1]) - 1]
        if day_cap > 0:
            for variant, label in [('eturna', 'Eturna EV'), ('tgdi', 'TGDI Petrol')]:
                target = t_val.get(variant, 0)
                planned = p_val.get(variant, 0)
                if planned < target:
                    shortage = target - planned
                    remarks.append(f"TCF-2 {day_name} - **{label}**: Target was `{target}`, but only `{planned}` vehicles were available (Shortage of `{shortage}`).")
                    
    if remarks:
        st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:0.8rem;">
                <span style="font-size:1.2rem;">⚠️</span>
                <h4 class="section-title">Target Shortage &amp; Production Remarks</h4>
            </div>
            """,
            unsafe_allow_html=True
        )
        for r_text in remarks:
            st.warning(r_text)
    else:
        st.success("All daily target quantities were fully met. No shortages identified in input files!")

    # Scheduled VCs by Category Summary List
    st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
    with st.expander("🔍 View Scheduled Vehicle Codes (VCs) by Category", expanded=False):
        tab1, tab2 = st.tabs(["TCF-1 (Punch/Altroz)", "TCF-2 (Harrier/Safari)"])
        
        with tab1:
            t1_lists = t1.get("vc_lists", {})
            for color_key, title, bg_color, text_color in [
                ("light_green", "Paint Floor Float", "#E2EFDA", "#385723"),
                ("light_blue", "WIP in BIW", "#DDEBF7", "#1F4E78"),
                ("light_yellow", "BIW Plan Loaded (Production Pending)", "#FFF2CC", "#7F6000"),
                ("light_pink", "Next 3-Days BIW Plan", "#FCE4D6", "#C65911")
            ]:
                vcs = t1_lists.get(color_key, [])
                if vcs:
                    st.markdown(
                        f"""
                        <div style="background:{bg_color}; padding:8px 12px; border-radius:6px; margin-bottom:8px; border:1px solid rgba(0,0,0,0.1); margin-top:12px;">
                            <strong style="color:{text_color};">{title} ({len(vcs)} units)</strong>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    st.text_area(f"VCs for {title} (TCF-1)", value=", ".join(vcs), height=100, key=f"tcf1_vc_text_{color_key}", disabled=True)
                else:
                    st.info(f"No vehicles scheduled from {title} for TCF-1.")
                    
        with tab2:
            t2_lists = t2.get("vc_lists", {})
            for color_key, title, bg_color, text_color in [
                ("light_green", "Paint Floor Float", "#E2EFDA", "#385723"),
                ("light_blue", "WIP in BIW", "#DDEBF7", "#1F4E78"),
                ("light_yellow", "BIW Plan Loaded (Production Pending)", "#FFF2CC", "#7F6000"),
                ("light_pink", "Next 3-Days BIW Plan", "#FCE4D6", "#C65911")
            ]:
                vcs = t2_lists.get(color_key, [])
                if vcs:
                    st.markdown(
                        f"""
                        <div style="background:{bg_color}; padding:8px 12px; border-radius:6px; margin-bottom:8px; border:1px solid rgba(0,0,0,0.1); margin-top:12px;">
                            <strong style="color:{text_color};">{title} ({len(vcs)} units)</strong>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    st.text_area(f"VCs for {title} (TCF-2)", value=", ".join(vcs), height=100, key=f"tcf2_vc_text_{color_key}", disabled=True)
                else:
                    st.info(f"No vehicles scheduled from {title} for TCF-2.")


# ─────────────────────────────────────────────
# 9. VEHICLE SELECTION / INFORMATION CARDS
# ─────────────────────────────────────────────
harrier_b64 = get_base64_image("assets/images/tata_harrier_3d.png")
safari_b64 = get_base64_image("assets/images/tata_safari_3d.png")
punch_b64 = get_base64_image("assets/images/tata_punch_3d.png")

with st.expander("🚗 TML Fleet Overview (Interactive Automotive Cards)", expanded=True):
    body_color = "#0F172A" if is_light else "#F8FAFC"
    card_bg = "#FFFFFF" if is_light else "#151E33"
    card_border = "#E2E8F0" if is_light else "rgba(255, 255, 255, 0.12)"
    card_shadow = "0 4px 14px rgba(15, 23, 42, 0.06)" if is_light else "0 8px 24px rgba(0, 0, 0, 0.4)"
    card_title_color = "#002B49" if is_light else "#38BDF8"
    card_desc_color = "#475569" if is_light else "#CBD5E1"

    cards_html = f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/vanilla-tilt/1.8.1/vanilla-tilt.min.js"></script>
    <style>
        body {{
            margin: 0;
            padding: 0;
            background: transparent;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            color: {body_color};
            display: flex;
            justify-content: center;
            align-items: center;
            overflow: hidden;
        }}
        .container {{
            display: flex;
            justify-content: space-around;
            align-items: center;
            flex-wrap: wrap;
            width: 100%;
            max-width: 1200px;
            padding: 16px;
            gap: 20px;
        }}
        .card {{
            position: relative;
            width: 330px;
            height: 380px;
            background: {card_bg};
            border: 1px solid {card_border};
            border-radius: 16px;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-end;
            padding: 24px 20px;
            box-sizing: border-box;
            box-shadow: {card_shadow};
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            transform-style: preserve-3d;
        }}
        .card:hover {{
            border-color: #0284C7;
            box-shadow: 0 12px 30px rgba(2, 132, 199, 0.15);
        }}
        .card .img-box {{
            position: absolute;
            top: 20px;
            width: 290px;
            height: 180px;
            display: flex;
            justify-content: center;
            align-items: center;
            transform: translateZ(50px);
            transition: 0.4s;
        }}
        .card .img-box img {{
            max-width: 92%;
            max-height: 92%;
            object-fit: contain;
            filter: drop-shadow(0 10px 15px rgba(0, 0, 0, 0.2));
        }}
        .card .content-box {{
            text-align: center;
            transform: translateZ(25px);
            z-index: 10;
        }}
        .card .content-box h2 {{
            font-size: 1.25rem;
            font-weight: 800;
            margin: 0;
            color: {card_title_color};
            letter-spacing: 0.3px;
        }}
        .card .content-box p {{
            font-size: 0.8rem;
            color: {card_desc_color};
            margin: 8px 0 0 0;
            line-height: 1.45;
        }}
        .card .badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 0.72rem;
            font-weight: 700;
            margin-top: 14px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
        }}
        .tcf2-badge {{
            background: #EDE9FE;
            color: #5B21B6;
            border: 1px solid #DDD6FE;
        }}
        .tcf1-badge {{
            background: #E0F2FE;
            color: #0369A1;
            border: 1px solid #BAE6FD;
        }}
    </style>
    <div class="container">
        <!-- Harrier Card -->
        <div class="card" data-tilt data-tilt-max="12" data-tilt-speed="400" data-tilt-perspective="1000">
            <div class="img-box">
                <img src="data:image/png;base64,{harrier_b64}" alt="Tata Harrier">
            </div>
            <div class="content-box">
                <h2>TATA HARRIER</h2>
                <p>Premium 5-Seater SUV built on the OMEGA Arc platform. Renowned for dynamic driving performance and bold stance.</p>
                <span class="badge tcf2-badge">TCF 2 Track</span>
            </div>
        </div>
        <!-- Safari Card -->
        <div class="card" data-tilt data-tilt-max="12" data-tilt-speed="400" data-tilt-perspective="1000">
            <div class="img-box">
                <img src="data:image/png;base64,{safari_b64}" alt="Tata Safari">
            </div>
            <div class="content-box">
                <h2>TATA SAFARI</h2>
                <p>Flagship 7-Seater SUV offering executive comfort, advanced tech, and commanding road presence.</p>
                <span class="badge tcf2-badge">TCF 2 Track</span>
            </div>
        </div>
        <!-- Punch Card -->
        <div class="card" data-tilt data-tilt-max="12" data-tilt-speed="400" data-tilt-perspective="1000">
            <div class="img-box">
                <img src="data:image/png;base64,{punch_b64}" alt="Tata Punch">
            </div>
            <div class="content-box">
                <h2>TATA PUNCH</h2>
                <p>Sub-compact SUV combining agile dimensions, 5-star safety rating, and tough SUV character.</p>
                <span class="badge tcf1-badge">TCF 1 Track</span>
            </div>
        </div>
    </div>
    """
    st.components.v1.html(cards_html, height=410)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown(
    """
    <div style="margin-top: 3rem; padding: 1.5rem 0; border-top: 1px solid #E2E8F0; text-align: center; color: #64748B; font-size: 0.78rem;">
        <strong>Tata Motors Limited</strong> &nbsp;·&nbsp; VIN Generation &amp; Production Planning Dashboard &nbsp;·&nbsp;
        For Internal Manufacturing Operations Use Only
    </div>
    """,
    unsafe_allow_html=True,
)
