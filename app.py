# -*- coding: utf-8 -*-
import streamlit as st
import base64
import os
from engine import initialize_tcf1_workbook, initialize_tcf2_workbook, update_today_vin_generation, update_paint_float_data

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
    page_title="TML VIN Generation Plan",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# CUSTOM CSS — Glasswork Liquid Glass Dark Theme
# ─────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* ── Google Font ── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* ── Ambient Dark Mesh Wallpaper Background ── */
    .stApp {
        background: 
            radial-gradient(ellipse 80% 50% at 50% -20%, rgba(56, 189, 248, 0.16), transparent 70%),
            radial-gradient(ellipse 60% 40% at 85% 60%, rgba(16, 185, 129, 0.12), transparent 60%),
            radial-gradient(ellipse 60% 40% at 15% 85%, rgba(99, 102, 241, 0.14), transparent 60%),
            linear-gradient(160deg, #090c14 0%, #0e1526 45%, #07090e 100%);
        background-attachment: fixed;
        min-height: 100vh;
    }

    /* ── Glasswork Liquid Glass Core Tokens & Utilities ── */
    .liquid-glass {
        background: color-mix(in srgb, rgba(18, 24, 40, 0.58) 80%, transparent);
        -webkit-backdrop-filter: blur(24px) saturate(190%) brightness(1.08);
        backdrop-filter: blur(24px) saturate(190%) brightness(1.08);
        border: 1px solid rgba(255, 255, 255, 0.12);
        box-shadow:
            inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.42),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.16),
            inset 0 0 0 0.5px rgba(255, 255, 255, 0.08),
            0 12px 36px rgba(0, 0, 0, 0.42),
            inset 0 6px 10px -6px rgba(255, 255, 255, 0.16),
            inset 4px 0 8px -6px rgba(255, 255, 255, 0.08),
            inset -4px 0 8px -6px rgba(255, 255, 255, 0.05),
            inset 0 -6px 8px -6px rgba(0, 0, 0, 0.28);
        border-radius: 20px;
        transition: all 0.35s cubic-bezier(0.16, 1, 0.3, 1);
    }

    .liquid-glass:hover {
        border-color: rgba(99, 179, 237, 0.4);
        box-shadow:
            inset 1.5px 1.5px 1.5px -0.5px rgba(255, 255, 255, 0.55),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.22),
            inset 0 0 0 0.5px rgba(255, 255, 255, 0.14),
            0 18px 48px rgba(0, 0, 0, 0.5),
            0 0 24px rgba(56, 189, 248, 0.16);
    }

    /* ── Liquid Glass Pill / Badge ── */
    .glass-pill {
        background: color-mix(in srgb, rgba(30, 41, 59, 0.65) 85%, transparent);
        -webkit-backdrop-filter: blur(14px) saturate(180%);
        backdrop-filter: blur(14px) saturate(180%);
        border: 1px solid rgba(255, 255, 255, 0.16);
        box-shadow:
            inset 1px 1px 0.5px rgba(255, 255, 255, 0.35),
            inset -1px -1px 0.5px rgba(255, 255, 255, 0.12),
            0 4px 12px rgba(0, 0, 0, 0.25);
        border-radius: 9999px;
        padding: 5px 16px;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }

    /* ── Hero banner Liquid Glass Deck ── */
    .hero-banner {
        background: 
            radial-gradient(circle at 10% 20%, rgba(56, 189, 248, 0.18), transparent 45%), 
            radial-gradient(circle at 90% 80%, rgba(16, 185, 129, 0.12), transparent 45%),
            color-mix(in srgb, rgba(15, 23, 42, 0.55) 80%, transparent);
        -webkit-backdrop-filter: blur(28px) saturate(200%) brightness(1.08);
        backdrop-filter: blur(28px) saturate(200%) brightness(1.08);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 24px;
        padding: 3rem;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
        box-shadow:
            inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.45),
            inset -1.5px -1.5px 1px -0.5px rgba(255, 255, 255, 0.18),
            inset 0 0 0 0.5px rgba(255, 255, 255, 0.1),
            0 20px 50px rgba(0, 0, 0, 0.45),
            inset 0 8px 12px -8px rgba(255, 255, 255, 0.2),
            inset 0 -8px 12px -8px rgba(0, 0, 0, 0.35);
    }
    .hero-title {
        font-size: 2.9rem;
        font-weight: 800;
        background: linear-gradient(135deg, #38BDF8 0%, #60A5FA 40%, #34D399 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin: 0 0 1rem 0;
        letter-spacing: -1px;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: color-mix(in srgb, rgba(56, 189, 248, 0.15) 85%, transparent);
        -webkit-backdrop-filter: blur(12px) saturate(180%);
        backdrop-filter: blur(12px) saturate(180%);
        border: 1px solid rgba(56, 189, 248, 0.35);
        box-shadow:
            inset 1px 1px 0.5px rgba(255, 255, 255, 0.4),
            inset -1px -1px 0.5px rgba(255, 255, 255, 0.15),
            0 4px 15px rgba(56, 189, 248, 0.2);
        border-radius: 9999px;
        padding: 6px 18px;
        font-size: 0.8rem;
        font-weight: 600;
        color: #38BDF8;
        margin-bottom: 1.2rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .hero-description {
        font-size: 1.05rem;
        color: #e2e8f0;
        line-height: 1.7;
        max-width: 820px;
        margin-bottom: 1.5rem;
    }
    .hero-features {
        display: flex;
        gap: 2.5rem;
        flex-wrap: wrap;
        margin-top: 1.8rem;
    }
    .hero-feature-item {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 0.88rem;
        color: #cbd5e1;
        font-weight: 500;
    }

    /* ── Section labels ── */
    .section-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin: 2rem 0 1.2rem 0;
    }
    .section-icon {
        width: 38px;
        height: 38px;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.1rem;
        flex-shrink: 0;
        background: color-mix(in srgb, rgba(56, 189, 248, 0.15) 80%, transparent);
        -webkit-backdrop-filter: blur(10px);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(56, 189, 248, 0.25);
        box-shadow: inset 1px 1px 0.5px rgba(255, 255, 255, 0.3);
    }
    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #f1f5f9;
        margin: 0;
        letter-spacing: -0.2px;
    }
    .section-subtitle {
        font-size: 0.82rem;
        color: #94a3b8;
        margin: 0;
    }

    /* ── Liquid Glass Upload Card Wrapper ── */
    .upload-card {
        background: color-mix(in srgb, rgba(20, 26, 42, 0.52) 80%, transparent);
        -webkit-backdrop-filter: blur(16px) saturate(180%) brightness(1.05);
        backdrop-filter: blur(16px) saturate(180%) brightness(1.05);
        border: 1px solid rgba(255, 255, 255, 0.12);
        box-shadow:
            inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.35),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.12),
            inset 0 0 0 0.5px rgba(255, 255, 255, 0.06),
            0 8px 24px rgba(0, 0, 0, 0.32);
        border-radius: 16px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 0.8rem;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .upload-card:hover {
        border-color: rgba(99, 179, 237, 0.4);
        box-shadow:
            inset 1.5px 1.5px 1.5px -0.5px rgba(255, 255, 255, 0.48),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.18),
            0 12px 32px rgba(0, 0, 0, 0.4),
            0 0 18px rgba(56, 189, 248, 0.12);
        transform: translateY(-1px);
    }
    .upload-label {
        font-size: 0.84rem;
        font-weight: 600;
        color: #e2e8f0;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .required-dot {
        width: 6px;
        height: 6px;
        background: #f87171;
        border-radius: 50%;
        display: inline-block;
        flex-shrink: 0;
        box-shadow: 0 0 8px #f87171;
    }

    /* ── Divider ── */
    .styled-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(99, 179, 237, 0.35), transparent);
        margin: 2.2rem 0;
        border: none;
    }

    /* ── Liquid Glass Generate Button ── */
    div[data-testid="stButton"] > button {
        width: 100%;
        background: linear-gradient(135deg, rgba(37, 99, 235, 0.85) 0%, rgba(30, 64, 175, 0.90) 100%) !important;
        -webkit-backdrop-filter: blur(12px) saturate(180%) !important;
        backdrop-filter: blur(12px) saturate(180%) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.25) !important;
        border-radius: 14px !important;
        padding: 0.85rem 2rem !important;
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.3px;
        box-shadow:
            inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.5),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.2),
            0 8px 25px rgba(37, 99, 235, 0.45) !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    div[data-testid="stButton"] > button:hover {
        background: linear-gradient(135deg, rgba(59, 130, 246, 0.95) 0%, rgba(37, 99, 235, 0.95) 100%) !important;
        border-color: rgba(255, 255, 255, 0.4) !important;
        box-shadow:
            inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.65),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.25),
            0 12px 35px rgba(37, 99, 235, 0.6),
            0 0 20px rgba(56, 189, 248, 0.35) !important;
        transform: translateY(-2px);
    }

    /* ── Cohesive Liquid Glass File Uploader ── */
    div[data-testid="stFileUploader"] {
        background: color-mix(in srgb, rgba(20, 26, 42, 0.5) 80%, transparent) !important;
        -webkit-backdrop-filter: blur(16px) saturate(180%) brightness(1.05) !important;
        backdrop-filter: blur(16px) saturate(180%) brightness(1.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        border-radius: 16px !important;
        padding: 1.2rem !important;
        margin-bottom: 1.2rem !important;
        box-shadow:
            inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.35),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.12),
            0 8px 24px rgba(0, 0, 0, 0.3) !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(99, 179, 237, 0.4) !important;
        box-shadow:
            inset 1.5px 1.5px 1.5px -0.5px rgba(255, 255, 255, 0.5),
            inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.18),
            0 12px 32px rgba(0, 0, 0, 0.4),
            0 0 18px rgba(56, 189, 248, 0.15) !important;
    }
    /* Label styling */
    div[data-testid="stFileUploader"] label {
        color: #f8fafc !important;
        font-size: 1.2rem !important;
        font-weight: 700 !important;
        margin-bottom: 0.8rem !important;
    }
    /* Required marker */
    div[data-testid="stFileUploader"] label [data-testid="stWidgetLabel"] p::after {
        content: "  *";
        color: #f87171;
        font-weight: 900;
        font-size: 1.3rem !important;
    }
    /* Dropzone area styling (Glass Well) */
    div[data-testid="stFileUploaderDropzone"] {
        background: rgba(0, 0, 0, 0.32) !important;
        border: 1.5px dashed rgba(255, 255, 255, 0.2) !important;
        border-radius: 12px !important;
        padding: 1rem !important;
        box-shadow: inset 0 2px 8px rgba(0, 0, 0, 0.4) !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stFileUploaderDropzone"]:hover {
        border-color: rgba(99, 179, 237, 0.6) !important;
        background: rgba(56, 189, 248, 0.05) !important;
    }
    /* Text inside dropzone */
    div[data-testid="stFileUploaderDropzone"] p {
        color: #cbd5e1 !important;
        font-size: 0.82rem !important;
    }
    /* Browse files button (Liquid Pill) */
    div[data-testid="stFileUploaderDropzone"] button {
        background: color-mix(in srgb, rgba(56, 189, 248, 0.18) 85%, transparent) !important;
        -webkit-backdrop-filter: blur(10px) !important;
        backdrop-filter: blur(10px) !important;
        color: #38bdf8 !important;
        border: 1px solid rgba(56, 189, 248, 0.35) !important;
        border-radius: 9999px !important;
        padding: 0.35rem 1rem !important;
        font-size: 0.78rem !important;
        font-weight: 600 !important;
        box-shadow:
            inset 1px 1px 0.5px rgba(255, 255, 255, 0.3),
            0 2px 8px rgba(0, 0, 0, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stFileUploaderDropzone"] button:hover {
        background: rgba(56, 189, 248, 0.3) !important;
        border-color: #38bdf8 !important;
        box-shadow: 0 0 12px rgba(56, 189, 248, 0.4) !important;
        color: #ffffff !important;
    }

    /* ── Status indicator ── */
    .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        display: inline-block;
        margin-right: 6px;
    }
    .status-dot.uploaded { 
        background: #34d399; 
        box-shadow: 0 0 8px #34d399; 
    }
    .status-dot.pending  { 
        background: #f87171; 
        box-shadow: 0 0 8px #f87171; 
    }

    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────
# FILE REGISTRY  (module-level — accessible everywhere)
# ─────────────────────────────────────────────
# Each entry: session-state key → human-readable label
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
import re

def detect_file_type(filename: str):
    """
    Classifies a filename into one of the 10 target categories regardless
    of date, month, year, or timestamp variations in the filename.
    """
    if not filename:
        return None
    fn = filename.upper().strip()
    # Normalize separators (underscores, hyphens, periods, extra spaces) to single spaces
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
        
    # 6. Next 3-Days BIW Plan (e.g. BIW Plan 3 days declaration..., Next 3 days BIW...)
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
        badge_bg = "color-mix(in srgb, rgba(59, 130, 246, 0.25) 85%, transparent)" if is_manual else "color-mix(in srgb, rgba(16, 185, 129, 0.25) 85%, transparent)"
        badge_color = "#93c5fd" if is_manual else "#6ee7b7"
        badge_border = "rgba(59, 130, 246, 0.4)" if is_manual else "rgba(16, 185, 129, 0.4)"
        st.markdown(
            f"""
            <div style="background:color-mix(in srgb, rgba(20,26,42,0.65) 85%, transparent); -webkit-backdrop-filter:blur(14px) saturate(180%); backdrop-filter:blur(14px) saturate(180%); border:1px solid rgba(255,255,255,0.14); box-shadow: inset 1px 1px 0.5px rgba(255,255,255,0.3), inset -1px -1px 0.5px rgba(255,255,255,0.1), 0 4px 14px rgba(0,0,0,0.3); border-radius:12px; padding:7px 14px; margin-bottom:10px; display:flex; align-items:center; justify-content:space-between; gap:10px;">
                <span style="font-size:0.82rem; color:#f1f5f9; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:500;">
                    📁 <strong>{active_f.name}</strong>
                </span>
                <span style="font-size:0.7rem; color:{badge_color}; background:{badge_bg}; border:1px solid {badge_border}; padding:3px 10px; border-radius:9999px; font-weight:600; white-space:nowrap; box-shadow:inset 1px 1px 0.5px rgba(255,255,255,0.2);">
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
# HERO BANNER
# ─────────────────────────────────────────────
st.markdown(
    """
    <div class="hero-banner">
        <div class="hero-badge">
            <span style="display:inline-block; width:8px; height:8px; background:#10b981; border-radius:50%; box-shadow:0 0 8px #10b981; margin-right:6px;"></span>
            🏭 &nbsp; Tata Motors Limited — Production Planning
        </div>
        <p class="hero-title">TML VIN Generation & Production Plan</p>
        <p class="hero-description">
            Automated multi-day rolling sequence optimizer for <strong>TCF 1</strong> (Punch EV / Petrol) and 
            <strong>TCF 2</strong> (Harrier / Safari) production lines. Dynamic stage-wise float matching and 
            WIP prioritization packed to daily capacity constraints.
        </p>
        <div class="hero-features">
            <div class="hero-feature-item">
                <span style="font-size: 1.1rem; color: #38bdf8;">🔄</span> 3-Day Rolling Plan
            </div>
            <div class="hero-feature-item">
                <span style="font-size: 1.1rem; color: #10b981;">⚡</span> Stage-wise Float Matching
            </div>
            <div class="hero-feature-item">
                <span style="font-size: 1.1rem; color: #f43f5e;">🛡️</span> Pune Plant Exclusions Active
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────
# BULK FILE UPLOAD & AUTO-DETECTION
# ─────────────────────────────────────────────
if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0

bulk_key = f"bulk_file_uploader_{st.session_state['uploader_version']}"

st.markdown(
    """
    <div class="liquid-glass" style="padding: 1.8rem; margin-bottom: 1.8rem; border-color: rgba(56, 189, 248, 0.35); box-shadow: inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.45), 0 16px 40px rgba(0,0,0,0.4), 0 0 25px rgba(56, 189, 248, 0.12);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 0.8rem; flex-wrap:wrap; gap:10px;">
            <div style="display:flex; align-items:center; gap:12px;">
                <span class="glass-pill" style="font-size: 1.3rem; padding: 6px 12px; color:#38bdf8;">⚡</span>
                <div>
                    <h3 style="margin:0; font-size:1.3rem; font-weight:700; color:#f8fafc; letter-spacing:-0.3px;">
                        Bulk File Upload &amp; Automatic Name Detection
                    </h3>
                    <p style="margin:0; font-size:0.84rem; color:#94a3b8; margin-top:2px;">
                        Drag and drop all 10 planning and float files at once — names are automatically identified regardless of date or month.
                    </p>
                </div>
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
    st.markdown("<h4 style='color:#38bdf8; font-size:1rem; margin-top:1rem; margin-bottom:0.5rem;'>🎯 Auto-Detection &amp; Mapping Results</h4>", unsafe_allow_html=True)
    
    head_col1, head_col2, head_col3 = st.columns([3, 3, 2])
    with head_col1:
        st.markdown("<span style='font-size:0.75rem; color:#94a3b8; font-weight:700; text-transform:uppercase;'>Uploaded File</span>", unsafe_allow_html=True)
    with head_col2:
        st.markdown("<span style='font-size:0.75rem; color:#94a3b8; font-weight:700; text-transform:uppercase;'>Auto-Detected Category</span>", unsafe_allow_html=True)
    with head_col3:
        st.markdown("<span style='font-size:0.75rem; color:#94a3b8; font-weight:700; text-transform:uppercase;'>Override Category</span>", unsafe_allow_html=True)
        
    for f in bulk_files:
        matched_cat = None
        for cat_k, file_obj in st.session_state["bulk_detected_files"].items():
            if file_obj.name == f.name:
                matched_cat = cat_k
                break
                
        r_c1, r_c2, r_c3 = st.columns([3, 3, 2])
        with r_c1:
            st.markdown(f"<div style='font-size:0.85rem; color:#f1f5f9; padding:6px 0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{f.name}'>📄 {f.name}</div>", unsafe_allow_html=True)
        with r_c2:
            if matched_cat:
                st.markdown(f"<div style='padding:4px 0;'><span class='glass-pill' style='font-size:0.78rem; color:#6ee7b7; background:color-mix(in srgb, rgba(16,185,129,0.2) 85%, transparent); border:1px solid rgba(16,185,129,0.4); font-weight:600;'>✓ {FILE_LABELS[matched_cat]}</span></div>", unsafe_allow_html=True)
            else:
                st.markdown("<div style='padding:4px 0;'><span class='glass-pill' style='font-size:0.78rem; color:#f87171; background:color-mix(in srgb, rgba(239,68,68,0.2) 85%, transparent); border:1px solid rgba(239,68,68,0.4); font-weight:600;'>⚠️ Unrecognized</span></div>", unsafe_allow_html=True)
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
# LIVE UPLOAD PROGRESS
# ─────────────────────────────────────────────
uploaded_count = sum(
    1 for k in FILE_KEYS if get_active_file(k) is not None
)
pct = int((uploaded_count / TOTAL_FILES) * 100)

col_prog_1, col_prog_2 = st.columns([1, 5])

with col_prog_1:
    st.markdown(
        f"""
        <div class="liquid-glass" style="text-align:center; border-radius:18px; padding: 1.1rem 0.6rem;">
            <div style="font-size:2.3rem;font-weight:800;color:#38bdf8; text-shadow:0 0 20px rgba(56,189,248,0.35);">{uploaded_count}/{TOTAL_FILES}</div>
            <div style="font-size:0.75rem;color:#94a3b8;margin-top:4px; font-weight:600;">Files Ready</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_prog_2:
    st.progress(pct / 100, text=f"Upload Progress — {pct}% complete")
    # Mini status checklist (2 rows for readability with 10 files)
    status_html = ""
    for k, label in FILE_LABELS.items():
        dot_cls = "uploaded" if get_active_file(k) is not None else "pending"
        status_html += (
            f'<span style="margin-right:14px;font-size:0.75rem;color:#a0aec0;">'
            f'<span class="status-dot {dot_cls}"></span>{label}</span>'
        )
    st.markdown(
        f'<div style="margin-top:8px;line-height:2.2;">{status_html}</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# INDIVIDUAL FILE UPLOADERS & OVERRIDES (COLLAPSIBLE)
# ─────────────────────────────────────────────
st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)

with st.expander("📂 Individual File Uploaders (Manual Upload & Single File Overrides)", expanded=(uploaded_count == 0)):
    # GROUP 1
    st.markdown(
        """
        <div class="section-header">
            <div class="section-icon" style="background:rgba(66,153,225,0.15);">📋</div>
            <div>
                <p class="section-title">Group 1 — Yesterday's VIN Generation Plans</p>
                <p class="section-subtitle">Baseline plans from the previous production day for TCF 1 &amp; TCF 2</p>
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

    # GROUP 2
    st.markdown(
        """
        <div class="section-header">
            <div class="section-icon" style="background:rgba(72,187,120,0.15);">📝</div>
            <div>
                <p class="section-title">Group 2 — Today's VIN Generation Lists</p>
                <p class="section-subtitle">Current-day VIN lists for TCF 1 &amp; TCF 2 production tracks</p>
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

    # GROUP 3
    st.markdown(
        """
        <div class="section-header">
            <div class="section-icon" style="background:rgba(237,137,54,0.15);">📊</div>
            <div>
                <p class="section-title">Group 3 — Reports &amp; Work-In-Progress Data</p>
                <p class="section-subtitle">Paint float, WIP trackers, BIW plan data and stage-wise reports</p>
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
# SHIFT PRODUCTION QUANTITIES (MANUAL INPUT)
# ─────────────────────────────────────────────
st.markdown(
    """
    <div style="background:rgba(237,137,54,0.05); border:1px solid rgba(237,137,54,0.2); border-radius:12px; padding:1.2rem; margin-bottom:1.5rem;">
        <h4 style="margin:0 0 0.5rem 0; color:#ed8936; font-size:1.1rem; display:flex; align-items:center; gap:8px;">
            📊 Phase 3: Shift Production Targets (Manual Entry)
        </h4>
        <p style="margin:0; font-size:0.85rem; color:#a0aec0; line-height:1.4;">
            Specify the <strong>Expected VIN B &amp; C shift tentative production quantities</strong> for TCF 1 and TCF 2 tracks before initializing plans.
        </p>
    </div>
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

st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# PRODUCTION LINE CAPACITIES (MANUAL ADJUSTMENT PER DAY)
# ─────────────────────────────────────────────
st.markdown(
    """
    <div class="section-header">
        <div class="section-icon" style="background:rgba(156,163,175,0.15);">📅</div>
        <div>
            <p class="section-title">Production Line Capacities</p>
            <p class="section-subtitle">Set planning capacities for each specific day on both tracks</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Dynamically calculate planning dates
baseline_date = None
yest_tcf1 = get_active_file("vin_plan_tcf1_yesterday")
if yest_tcf1 is not None:
    try:
        import io
        import openpyxl
        import datetime
        wb_temp = openpyxl.load_workbook(io.BytesIO(yest_tcf1.getvalue()), read_only=True, data_only=True)
        if 'Plan Summary' in wb_temp.sheetnames:
            ws_temp = wb_temp['Plan Summary']
            val_h2 = ws_temp.cell(row=2, column=8).value # Yesterday's Day 1
            if isinstance(val_h2, (datetime.datetime, datetime.date)):
                baseline_date = val_h2
    except Exception:
        pass

import datetime
if baseline_date is None:
    baseline_date = datetime.date.today()
elif isinstance(baseline_date, datetime.datetime):
    baseline_date = baseline_date.date()

def get_next_planning_days(start_date, count=3):
    days = []
    curr = start_date
    while len(days) < count:
        curr += datetime.timedelta(days=1)
        if curr.weekday() == 6: # Sunday
            continue
        days.append(curr)
    return days

planning_days = get_next_planning_days(baseline_date, 3)
d1, d2, d3 = planning_days

d1_str = d1.strftime('%Y-%m-%d (%A)')
d2_str = d2.strftime('%Y-%m-%d (%A)')
d3_str = d3.strftime('%Y-%m-%d (%A)')

cap_col1, cap_col2 = st.columns(2, gap="large")

with cap_col1:
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:10px; margin-bottom:1.2rem;">
            <span class="glass-pill" style="color:#38bdf8; font-weight:700; border-color:rgba(56,189,248,0.4);">TCF-1</span>
            <h3 style="margin:0; font-size:1.15rem; color:#f8fafc; font-weight:700;">Punch / Altroz Daily Targets</h3>
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
        st.markdown(f"<div style='margin-top:0.8rem; margin-bottom:0.3rem;'><span class='glass-pill' style='padding:2px 12px; font-size:0.75rem; color:#e2e8f0; font-weight:600;'>Day {day_num}</span></div>", unsafe_allow_html=True)
        sc1, sc2, sc3, sc4 = st.columns(4)
        with sc1:
            total_cap = st.number_input("Total Target", min_value=0, max_value=2000, value=def_total, step=50, key=f"tcf1_cap_d{day_num}")
        with sc2:
            ev_cap = st.number_input("Punch EV Limit", min_value=0, max_value=2000, value=def_ev, step=10, key=f"tcf1_ev_d{day_num}")
        with sc3:
            cng_cap = st.number_input("Punch CNG Limit", min_value=0, max_value=2000, value=def_cng, step=10, key=f"tcf1_cng_d{day_num}")
        with sc4:
            rem_petrol = total_cap - ev_cap - cng_cap
            if rem_petrol < 0:
                st.markdown(f"<span style='color:#fc8181;font-size:0.85rem;font-weight:600;'>Petrol: {rem_petrol} ⚠️</span>", unsafe_allow_html=True)
            else:
                st.markdown(f"<span style='color:#a0aec0;font-size:0.85rem;'>Petrol: <strong>{rem_petrol}</strong></span>", unsafe_allow_html=True)
        tcf1_caps.append(total_cap)
        tcf1_ev.append(ev_cap)
        tcf1_cng.append(cng_cap)
        
    st.session_state["tcf1_capacities"] = tcf1_caps
    st.session_state["tcf1_ev"] = tcf1_ev
    st.session_state["tcf1_cng"] = tcf1_cng

with cap_col2:
    st.markdown(
        """
        <div style="display:flex; align-items:center; gap:10px; margin-bottom:1.2rem;">
            <span class="glass-pill" style="color:#34d399; font-weight:700; border-color:rgba(52,211,153,0.4);">TCF-2</span>
            <h3 style="margin:0; font-size:1.15rem; color:#f8fafc; font-weight:700;">Harrier / Safari Daily Targets</h3>
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
        st.markdown(f"<div style='margin-top:0.8rem; margin-bottom:0.3rem;'><span class='glass-pill' style='padding:2px 12px; font-size:0.75rem; color:#e2e8f0; font-weight:600;'>Day {day_num}</span></div>", unsafe_allow_html=True)
        sc1, sc2, sc3, sc4 = st.columns(4)
        with sc1:
            total_cap = st.number_input("Total Target", min_value=0, max_value=2000, value=def_total, step=50, key=f"tcf2_cap_d{day_num}")
        with sc2:
            et_cap = st.number_input("Eturna EV Limit", min_value=0, max_value=2000, value=def_ev, step=10, key=f"tcf2_et_d{day_num}")
        with sc3:
            tg_cap = st.number_input("TGDI Petrol Limit", min_value=0, max_value=2000, value=def_tgdi, step=10, key=f"tcf2_tg_d{day_num}")
        with sc4:
            rem_diesel = total_cap - et_cap - tg_cap
            if rem_diesel < 0:
                st.markdown(f"<span style='color:#fc8181;font-size:0.85rem;font-weight:600;'>Diesel: {rem_diesel} ⚠️</span>", unsafe_allow_html=True)
            else:
                st.markdown(f"<span style='color:#a0aec0;font-size:0.85rem;'>Diesel: <strong>{rem_diesel}</strong></span>", unsafe_allow_html=True)
        tcf2_caps.append(total_cap)
        tcf2_eturna.append(et_cap)
        tcf2_tgdi.append(tg_cap)
        
    st.session_state["tcf2_capacities"] = tcf2_caps
    st.session_state["tcf2_eturna"] = tcf2_eturna
    st.session_state["tcf2_tgdi"] = tcf2_tgdi

st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)

btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])
with btn_col2:
    generate_clicked = st.button(
        "⚙️  Generate 3-Day Plan",
        key="generate_btn",
        use_container_width=True,
    )

if generate_clicked:
    # pending_plan_biw and next_3days_biw_plan are optional
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
        st.info("⚙️  **Processing and generating 3-Day Production Plans...**")
        try:
            import io
            
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

# Render generated output dashboards if state is set
if st.session_state.get("plans_generated", False):
    
    tcf1_caps = st.session_state.get("tcf1_capacities", [900, 900, 900])
    tcf2_caps = st.session_state.get("tcf2_capacities", [250, 250, 250])
    tcf1_cap_strs = [f"{cap} (Holiday)" if cap == 0 else str(cap) for cap in tcf1_caps]
    tcf2_cap_strs = [f"{cap} (Holiday)" if cap == 0 else str(cap) for cap in tcf2_caps]
    
    st.success(
        f"✅  **Production Plans generated successfully!**\n\n"
        f"- TCF-1 (Punch/Altroz): Day 1 Capacity: `{tcf1_cap_strs[0]}`, Day 2: `{tcf1_cap_strs[1]}`, Day 3: `{tcf1_cap_strs[2]}`.\n"
        f"- TCF-2 (Harrier/Safari): Day 1 Capacity: `{tcf2_cap_strs[0]}`, Day 2: `{tcf2_cap_strs[1]}`, Day 3: `{tcf2_cap_strs[2]}`.",
        icon="✅"
    )
    
    dl_col1, dl_col2 = st.columns(2, gap="large")
    with dl_col1:
        st.download_button(
            label="📥 Download TCF-1 3-Day Plan",
            data=st.session_state["tcf1_out_bytes"],
            file_name="TCF-1_3-Day_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_tcf1_3day",
            use_container_width=True
        )
    with dl_col2:
        st.download_button(
            label="📥 Download TCF-2 3-Day Plan",
            data=st.session_state["tcf2_out_bytes"],
            file_name="TCF-2_3-Day_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_tcf2_3day",
            use_container_width=True
        )
        
    # Render Planned Vehicle Source Color Counts
    st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
    t1 = st.session_state["tcf1_summary"]
    t2 = st.session_state["tcf2_summary"]
    st.markdown(
        f"""
        <div class="liquid-glass" style="padding:1.6rem; margin-bottom:1.8rem; border-color:rgba(56,189,248,0.35);">
            <h4 style="margin:0 0 1.2rem 0; color:#38BDF8; font-size:1.15rem; display:flex; align-items:center; gap:10px;">
                <span class="glass-pill" style="font-size:1.1rem; padding:4px 10px;">🎨</span> Planned Vehicle Count by Source Type (Color-Coded)
            </h4>
            <div style="display:flex; justify-content:space-between; gap:20px; flex-wrap:wrap;">
                <!-- TCF-1 Column -->
                <div style="flex:1; min-width:280px; background:color-mix(in srgb, rgba(15,23,42,0.65) 85%, transparent); -webkit-backdrop-filter:blur(14px); backdrop-filter:blur(14px); border-radius:14px; padding:1.2rem; border:1px solid rgba(255,255,255,0.1); box-shadow:inset 1px 1px 0.5px rgba(255,255,255,0.25);">
                    <h5 style="margin:0 0 0.8rem 0; color:#90cdf4; font-size:0.95rem; font-weight:700;">TCF-1 (Punch/Altroz)</h5>
                    <div style="display:flex; flex-direction:column; gap:10px; font-size:0.86rem;">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#E2EFDA; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #E2EFDA;"></span>Paint Floor Float (Light Green):</span>
                            <strong style="color:#f8fafc;">{t1['colors']['light_green']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#DDEBF7; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #DDEBF7;"></span>WIP in BIW (Light Blue):</span>
                            <strong style="color:#f8fafc;">{t1['colors']['light_blue']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FFF2CC; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #FFF2CC;"></span>BIW Plan Loaded (Production Pending) (Light Yellow):</span>
                            <strong style="color:#f8fafc;">{t1['colors']['light_yellow']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FCE4D6; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #FCE4D6;"></span>Next 3-Days BIW Plan (Light Pink):</span>
                            <strong style="color:#f8fafc;">{t1['colors']['light_pink']}</strong>
                        </div>
                    </div>
                </div>
                <!-- TCF-2 Column -->
                <div style="flex:1; min-width:280px; background:color-mix(in srgb, rgba(15,23,42,0.65) 85%, transparent); -webkit-backdrop-filter:blur(14px); backdrop-filter:blur(14px); border-radius:14px; padding:1.2rem; border:1px solid rgba(255,255,255,0.1); box-shadow:inset 1px 1px 0.5px rgba(255,255,255,0.25);">
                    <h5 style="margin:0 0 0.8rem 0; color:#98db9c; font-size:0.95rem; font-weight:700;">TCF-2 (Harrier/Safari)</h5>
                    <div style="display:flex; flex-direction:column; gap:10px; font-size:0.86rem;">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#E2EFDA; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #E2EFDA;"></span>Paint Floor Float (Light Green):</span>
                            <strong style="color:#f8fafc;">{t2['colors']['light_green']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#DDEBF7; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #DDEBF7;"></span>WIP in BIW (Light Blue):</span>
                            <strong style="color:#f8fafc;">{t2['colors']['light_blue']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FFF2CC; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #FFF2CC;"></span>BIW Plan Loaded (Production Pending) (Light Yellow):</span>
                            <strong style="color:#f8fafc;">{t2['colors']['light_yellow']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FCE4D6; margin-right:8px; border:1px solid rgba(0,0,0,0.15); box-shadow:0 0 6px #FCE4D6;"></span>Next 3-Days BIW Plan (Light Pink):</span>
                            <strong style="color:#f8fafc;">{t2['colors']['light_pink']}</strong>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    # Render Planned Models Summary Dashboard
    st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="liquid-glass" style="padding:1.4rem; margin-bottom:1.8rem; border-color:rgba(16,185,129,0.35); box-shadow: inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.4), 0 16px 40px rgba(0,0,0,0.4), 0 0 25px rgba(16,185,129,0.12);">
            <h4 style="margin:0 0 0.4rem 0; color:#34D399; font-size:1.15rem; display:flex; align-items:center; gap:10px;">
                <span class="glass-pill" style="font-size:1.1rem; padding:4px 10px;">📊</span> Model Plan Summary (NOVA, CNG, Eturna, TGDI)
            </h4>
            <p style="margin:0; font-size:0.85rem; color:#cbd5e1; line-height:1.4;">
                Highlighting the number of constrained model units planned across yesterday's float drops, today, and the next 3 days.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    tcf1_ev = st.session_state.get("tcf1_ev", [160, 90, 160])
    tcf1_cng = st.session_state.get("tcf1_cng", [350, 350, 350])
    tcf2_eturna = st.session_state.get("tcf2_eturna", [120, 120, 120])
    tcf2_tgdi = st.session_state.get("tcf2_tgdi", [40, 40, 40])

    sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4, gap="large")
    
    with sum_col1:
        st.markdown(
            f"""
            <div class="liquid-glass" style="border-radius:20px; padding:1.6rem; text-align:center; border-color:rgba(56,189,248,0.35); box-shadow: inset 1.5px 1.5px 1px -0.5px rgba(255,255,255,0.45), inset -1px -1px 1px -0.5px rgba(56,189,248,0.2), 0 16px 40px rgba(0,0,0,0.4), 0 0 24px rgba(56,189,248,0.12);">
                <div style="font-size:2.5rem; margin-bottom:0.4rem; filter: drop-shadow(0 0 12px rgba(56,189,248,0.5));">⚡</div>
                <h4 style="color:#38bdf8; margin:0 0 0.2rem 0; font-size:1.25rem; font-weight:700;">NOVA (Punch EV)</h4>
                <p style="color:#94a3b8; font-size:0.75rem; margin:0 0 1.1rem 0; text-transform:uppercase; letter-spacing:0.06em;">Daily Target Limits</p>
                <div style="text-align:left; font-size:0.86rem; color:#f1f5f9; background:rgba(0,0,0,0.32); -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); border:1px solid rgba(255,255,255,0.08); border-radius:12px; padding:0.9rem; box-shadow:inset 0 2px 6px rgba(0,0,0,0.3);">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; border-bottom:1px solid rgba(255,255,255,0.08); padding-bottom:0.4rem;">
                        <span style="color:#cbd5e1;">B&C Drops Today:</span> <strong style="color:#38bdf8;">{t1['picked']['nova']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Today (Day 1):</span> <span><strong>{t1['day1']['nova']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf1_ev[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t1['day2']['nova']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf1_ev[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t1['day3']['nova']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf1_ev[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
    with sum_col2:
        st.markdown(
            f"""
            <div class="liquid-glass" style="border-radius:20px; padding:1.6rem; text-align:center; border-color:rgba(251,146,60,0.35); box-shadow: inset 1.5px 1.5px 1px -0.5px rgba(255,255,255,0.45), inset -1px -1px 1px -0.5px rgba(251,146,60,0.2), 0 16px 40px rgba(0,0,0,0.4), 0 0 24px rgba(251,146,60,0.12);">
                <div style="font-size:2.5rem; margin-bottom:0.4rem; filter: drop-shadow(0 0 12px rgba(251,146,60,0.5));">🔥</div>
                <h4 style="color:#fb923c; margin:0 0 0.2rem 0; font-size:1.25rem; font-weight:700;">Punch CNG</h4>
                <p style="color:#94a3b8; font-size:0.75rem; margin:0 0 1.1rem 0; text-transform:uppercase; letter-spacing:0.06em;">Daily Target Limits</p>
                <div style="text-align:left; font-size:0.86rem; color:#f1f5f9; background:rgba(0,0,0,0.32); -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); border:1px solid rgba(255,255,255,0.08); border-radius:12px; padding:0.9rem; box-shadow:inset 0 2px 6px rgba(0,0,0,0.3);">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; border-bottom:1px solid rgba(255,255,255,0.08); padding-bottom:0.4rem;">
                        <span style="color:#cbd5e1;">B&C Drops Today:</span> <strong style="color:#fb923c;">{t1['picked']['cng']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Today (Day 1):</span> <span><strong>{t1['day1']['cng']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf1_cng[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t1['day2']['cng']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf1_cng[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t1['day3']['cng']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf1_cng[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
    with sum_col3:
        st.markdown(
            f"""
            <div class="liquid-glass" style="border-radius:20px; padding:1.6rem; text-align:center; border-color:rgba(52,211,153,0.35); box-shadow: inset 1.5px 1.5px 1px -0.5px rgba(255,255,255,0.45), inset -1px -1px 1px -0.5px rgba(52,211,153,0.2), 0 16px 40px rgba(0,0,0,0.4), 0 0 24px rgba(52,211,153,0.12);">
                <div style="font-size:2.5rem; margin-bottom:0.4rem; filter: drop-shadow(0 0 12px rgba(52,211,153,0.5));">🌱</div>
                <h4 style="color:#34d399; margin:0 0 0.2rem 0; font-size:1.25rem; font-weight:700;">Eturna (Harrier EV)</h4>
                <p style="color:#94a3b8; font-size:0.75rem; margin:0 0 1.1rem 0; text-transform:uppercase; letter-spacing:0.06em;">Daily Target Limits</p>
                <div style="text-align:left; font-size:0.86rem; color:#f1f5f9; background:rgba(0,0,0,0.32); -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); border:1px solid rgba(255,255,255,0.08); border-radius:12px; padding:0.9rem; box-shadow:inset 0 2px 6px rgba(0,0,0,0.3);">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; border-bottom:1px solid rgba(255,255,255,0.08); padding-bottom:0.4rem;">
                        <span style="color:#cbd5e1;">B&C Drops Today:</span> <strong style="color:#34d399;">{t2['picked']['eturna']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Today (Day 1):</span> <span><strong>{t2['day1']['eturna']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf2_eturna[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t2['day2']['eturna']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf2_eturna[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t2['day3']['eturna']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf2_eturna[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with sum_col4:
        st.markdown(
            f"""
            <div class="liquid-glass" style="border-radius:20px; padding:1.6rem; text-align:center; border-color:rgba(192,132,252,0.35); box-shadow: inset 1.5px 1.5px 1px -0.5px rgba(255,255,255,0.45), inset -1px -1px 1px -0.5px rgba(192,132,252,0.2), 0 16px 40px rgba(0,0,0,0.4), 0 0 24px rgba(192,132,252,0.12);">
                <div style="font-size:2.5rem; margin-bottom:0.4rem; filter: drop-shadow(0 0 12px rgba(192,132,252,0.5));">🏎️</div>
                <h4 style="color:#c084fc; margin:0 0 0.2rem 0; font-size:1.25rem; font-weight:700;">TGDI (Petrol)</h4>
                <p style="color:#94a3b8; font-size:0.75rem; margin:0 0 1.1rem 0; text-transform:uppercase; letter-spacing:0.06em;">Daily Target Limits</p>
                <div style="text-align:left; font-size:0.86rem; color:#f1f5f9; background:rgba(0,0,0,0.32); -webkit-backdrop-filter:blur(10px); backdrop-filter:blur(10px); border:1px solid rgba(255,255,255,0.08); border-radius:12px; padding:0.9rem; box-shadow:inset 0 2px 6px rgba(0,0,0,0.3);">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; border-bottom:1px solid rgba(255,255,255,0.08); padding-bottom:0.4rem;">
                        <span style="color:#cbd5e1;">B&C Drops Today:</span> <strong style="color:#c084fc;">{t2['picked']['tgdi']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Today (Day 1):</span> <span><strong>{t2['day1']['tgdi']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf2_tgdi[0]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t2['day2']['tgdi']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf2_tgdi[1]}</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 3 Plan:</span> <span><strong>{t2['day3']['tgdi']}</strong> <span style="color:#94a3b8;font-size:0.75rem;">/ {tcf2_tgdi[2]}</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    # ─────────────────────────────────────────────
    # TARGET SHORTAGE & PRODUCTION REMARKS
    # ─────────────────────────────────────────────
    st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="liquid-glass" style="padding:1.4rem; margin-bottom:1.8rem; border-color:rgba(239,68,68,0.35); box-shadow: inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.4), 0 16px 40px rgba(0,0,0,0.4), 0 0 25px rgba(239,68,68,0.12);">
            <h4 style="margin:0 0 0.4rem 0; color:#EF4444; font-size:1.15rem; display:flex; align-items:center; gap:10px;">
                <span class="glass-pill" style="font-size:1.1rem; padding:4px 10px;">⚠️</span> Target Shortage &amp; Production Remarks
            </h4>
            <p style="margin:0; font-size:0.85rem; color:#cbd5e1; line-height:1.4;">
                Generated warnings if available vehicles in float or BIW next 3 days plan fell short of target quantities.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    remarks = []
    
    # Check TCF-1
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
                    
    # Check TCF-2
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
        for r_text in remarks:
            st.warning(r_text)
    else:
        st.success("All daily target quantities were fully met. No shortages identified in input files!")

    # Scheduled VCs by Category Summary List
    st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
    with st.expander("🔍 View Scheduled Vehicle Codes (VCs) by Category", expanded=True):
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
# VEHICLE 3D CARDS SHOWCASE
# ─────────────────────────────────────────────
harrier_b64 = get_base64_image("assets/images/tata_harrier_3d.png")
safari_b64 = get_base64_image("assets/images/tata_safari_3d.png")
punch_b64 = get_base64_image("assets/images/tata_punch_3d.png")

with st.expander("🚗 TML Fleet Overview (Interactive 3D Cards)", expanded=True):
    # Constructing HTML for the 3D cards
    cards_html = f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/vanilla-tilt/1.8.1/vanilla-tilt.min.js"></script>
    <style>
        body {{
            margin: 0;
            padding: 0;
            background: transparent;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            color: #e2e8f0;
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
            padding: 20px;
            gap: 20px;
        }}
        .card {{
            position: relative;
            width: 320px;
            height: 380px;
            background: radial-gradient(circle at 20% 20%, rgba(56, 189, 248, 0.12), transparent 60%),
                        rgba(18, 24, 40, 0.65);
            -webkit-backdrop-filter: blur(24px) saturate(190%) brightness(1.08);
            backdrop-filter: blur(24px) saturate(190%) brightness(1.08);
            border: 1px solid rgba(255, 255, 255, 0.14);
            border-radius: 24px;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-end;
            padding: 25px 20px;
            box-sizing: border-box;
            box-shadow:
                inset 1.5px 1.5px 1px -0.5px rgba(255, 255, 255, 0.45),
                inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.18),
                inset 0 0 0 0.5px rgba(255, 255, 255, 0.1),
                0 16px 40px rgba(0, 0, 0, 0.45),
                inset 0 8px 12px -8px rgba(255, 255, 255, 0.2),
                inset 0 -8px 12px -8px rgba(0, 0, 0, 0.35);
            transition: 0.4s cubic-bezier(0.16, 1, 0.3, 1);
            transform-style: preserve-3d;
        }}
        .card:hover {{
            border-color: rgba(99, 179, 237, 0.5);
            box-shadow:
                inset 1.5px 1.5px 1.5px -0.5px rgba(255, 255, 255, 0.65),
                inset -1px -1px 1px -0.5px rgba(255, 255, 255, 0.25),
                0 25px 60px rgba(0, 0, 0, 0.6),
                0 0 30px rgba(56, 189, 248, 0.22);
        }}
        .card .img-box {{
            position: absolute;
            top: 20px;
            width: 280px;
            height: 180px;
            display: flex;
            justify-content: center;
            align-items: center;
            transform: translateZ(60px);
            transition: 0.5s;
        }}
        .card .img-box img {{
            max-width: 90%;
            max-height: 90%;
            object-fit: contain;
            filter: drop-shadow(0 15px 15px rgba(0, 0, 0, 0.6));
        }}
        .card .content-box {{
            text-align: center;
            transform: translateZ(30px);
            z-index: 10;
        }}
        .card .content-box h2 {{
            font-size: 1.3rem;
            font-weight: 700;
            margin: 0;
            color: #fff;
            letter-spacing: 0.5px;
        }}
        .card .content-box p {{
            font-size: 0.8rem;
            color: #cbd5e1;
            margin: 10px 0 0 0;
            line-height: 1.4;
        }}
        .card .badge {{
            display: inline-block;
            padding: 5px 14px;
            border-radius: 9999px;
            font-size: 0.72rem;
            font-weight: 700;
            margin-top: 15px;
            text-transform: uppercase;
            letter-spacing: 1px;
            backdrop-filter: blur(10px);
            box-shadow: inset 1px 1px 0.5px rgba(255, 255, 255, 0.3);
        }}
        .harrier-badge {{
            background: rgba(99, 179, 237, 0.12);
            color: #63b3ed;
            border: 1px solid rgba(99, 179, 237, 0.25);
        }}
        .safari-badge {{
            background: rgba(237, 137, 54, 0.12);
            color: #ed8936;
            border: 1px solid rgba(237, 137, 54, 0.25);
        }}
        .punch-badge {{
            background: rgba(72, 187, 120, 0.12);
            color: #48bb78;
            border: 1px solid rgba(72, 187, 120, 0.25);
        }}
    </style>
    <div class="container">
        <!-- Harrier Card -->
        <div class="card" data-tilt data-tilt-max="15" data-tilt-speed="400" data-tilt-perspective="1000">
            <div class="img-box">
                <img src="data:image/png;base64,{harrier_b64}" alt="Tata Harrier">
            </div>
            <div class="content-box">
                <h2>TATA HARRIER</h2>
                <p>Premium 5-Seater SUV Built on the OMEGA Arc platform. Known for its dynamic driving performance and bold stance.</p>
                <span class="badge harrier-badge">TCF 2 Track</span>
            </div>
        </div>
        <!-- Safari Card -->
        <div class="card" data-tilt data-tilt-max="15" data-tilt-speed="400" data-tilt-perspective="1000">
            <div class="img-box">
                <img src="data:image/png;base64,{safari_b64}" alt="Tata Safari">
            </div>
            <div class="content-box">
                <h2>TATA SAFARI</h2>
                <p>Flagship 7-Seater SUV offering premium comfort, advanced features, and a commanding road presence.</p>
                <span class="badge safari-badge">TCF 2 Track</span>
            </div>
        </div>
        <!-- Punch Card -->
        <div class="card" data-tilt data-tilt-max="15" data-tilt-speed="400" data-tilt-perspective="1000">
            <div class="img-box">
                <img src="data:image/png;base64,{punch_b64}" alt="Tata Punch">
            </div>
            <div class="content-box">
                <h2>TATA PUNCH</h2>
                <p>Sub-compact SUV combining compact dimensions with tough SUV character, high safety ratings, and agility.</p>
                <span class="badge punch-badge">TCF 1 Track</span>
            </div>
        </div>
    </div>
    """
    st.components.v1.html(cards_html, height=420)


# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown(
    """
    <div style="margin-top: 3rem; text-align: center; color: #4a5568; font-size: 0.75rem;">
        TML VIN Generation Plan &nbsp;·&nbsp; Phase 1: Data Input &nbsp;·&nbsp;
        For internal production planning use only
    </div>
    """,
    unsafe_allow_html=True,
)
