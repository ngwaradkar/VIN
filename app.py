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
# CUSTOM CSS — clean professional dark theme
# ─────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* ── Google Font ── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* ── Root background ── */
    .stApp {
        background: linear-gradient(135deg, #0f1117 0%, #1a1d2e 50%, #0f1117 100%);
        min-height: 100vh;
    }

    /* ── Hero banner ── */
    .hero-banner {
        background: radial-gradient(circle at top left, rgba(59, 130, 246, 0.12), transparent 50%), 
                    radial-gradient(circle at bottom right, rgba(16, 185, 129, 0.06), transparent 50%),
                    rgba(15, 23, 42, 0.45);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 20px;
        padding: 3rem;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.25);
    }
    .hero-title {
        font-size: 2.8rem;
        font-weight: 800;
        background: linear-gradient(135deg, #38BDF8 0%, #3B82F6 40%, #10B981 100%);
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
        background: linear-gradient(90deg, rgba(56, 189, 248, 0.1), rgba(16, 185, 129, 0.1));
        border: 1px solid rgba(56, 189, 248, 0.25);
        border-radius: 30px;
        padding: 6px 16px;
        font-size: 0.8rem;
        font-weight: 600;
        color: #38BDF8;
        margin-bottom: 1.2rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .hero-description {
        font-size: 1.05rem;
        color: #cbd5e1;
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
        color: #94a3b8;
        font-weight: 500;
    }

    /* ── Section labels ── */
    .section-header {
        display: flex;
        align-items: center;
        gap: 10px;
        margin: 1.8rem 0 1rem 0;
    }
    .section-icon {
        width: 36px;
        height: 36px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1rem;
        flex-shrink: 0;
    }
    .section-title {
        font-size: 1rem;
        font-weight: 600;
        color: #e2e8f0;
        margin: 0;
    }
    .section-subtitle {
        font-size: 0.8rem;
        color: #718096;
        margin: 0;
    }

    /* ── Upload card wrapper ── */
    .upload-card {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 12px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 0.8rem;
        transition: border-color 0.2s ease;
    }
    .upload-card:hover {
        border-color: rgba(99,179,237,0.3);
    }
    .upload-label {
        font-size: 0.82rem;
        font-weight: 600;
        color: #cbd5e0;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .required-dot {
        width: 6px;
        height: 6px;
        background: #fc8181;
        border-radius: 50%;
        display: inline-block;
        flex-shrink: 0;
    }

    /* ── Divider ── */
    .styled-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(99,179,237,0.25), transparent);
        margin: 2rem 0;
        border: none;
    }

    /* ── Generate button override ── */
    div[data-testid="stButton"] > button {
        width: 100%;
        background: linear-gradient(135deg, #2b6cb0 0%, #2c5282 100%);
        color: white !important;
        border: 1px solid rgba(99,179,237,0.4) !important;
        border-radius: 10px !important;
        padding: 0.75rem 2rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.3px;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 15px rgba(43,108,176,0.3);
    }
    div[data-testid="stButton"] > button:hover {
        background: linear-gradient(135deg, #3182ce 0%, #2b6cb0 100%) !important;
        border-color: rgba(99,179,237,0.7) !important;
        box-shadow: 0 6px 20px rgba(43,108,176,0.4);
        transform: translateY(-1px);
    }

    /* ── Cohesive File Uploader Style ── */
    div[data-testid="stFileUploader"] {
        background: rgba(255, 255, 255, 0.02) !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-radius: 12px !important;
        padding: 1.2rem !important;
        margin-bottom: 1.2rem !important;
        transition: all 0.25s ease !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(99, 179, 237, 0.4) !important;
        background: rgba(99, 179, 237, 0.02) !important;
        box-shadow: 0 4px 15px rgba(99, 179, 237, 0.05) !important;
    }
    /* Label styling */
    div[data-testid="stFileUploader"] label {
        color: #f7fafc !important;
        font-size: 1.25rem !important;
        font-weight: 700 !important;
        margin-bottom: 0.8rem !important;
    }
    /* Required marker */
    div[data-testid="stFileUploader"] label [data-testid="stWidgetLabel"] p::after {
        content: "  *";
        color: #fc8181;
        font-weight: 900;
        font-size: 1.4rem !important;
    }
    /* Dropzone area styling */
    div[data-testid="stFileUploaderDropzone"] {
        background: rgba(0, 0, 0, 0.2) !important;
        border: 1.5px dashed rgba(255, 255, 255, 0.15) !important;
        border-radius: 8px !important;
        padding: 0.8rem !important;
    }
    div[data-testid="stFileUploaderDropzone"]:hover {
        border-color: rgba(99, 179, 237, 0.6) !important;
    }
    /* Text inside dropzone */
    div[data-testid="stFileUploaderDropzone"] p {
        color: #a0aec0 !important;
        font-size: 0.8rem !important;
    }
    /* Browse files button */
    div[data-testid="stFileUploaderDropzone"] button {
        background: rgba(99, 179, 237, 0.15) !important;
        color: #90cdf4 !important;
        border: 1px solid rgba(99, 179, 237, 0.3) !important;
        border-radius: 6px !important;
        padding: 0.3rem 0.8rem !important;
        font-size: 0.75rem !important;
        font-weight: 500 !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stFileUploaderDropzone"] button:hover {
        background: rgba(99, 179, 237, 0.25) !important;
        border-color: #63b3ed !important;
        color: #fff !important;
    }

    /* ── Status indicator ── */
    .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        display: inline-block;
        margin-right: 6px;
    }
    .status-dot.uploaded { background: #68d391; }
    .status-dot.pending  { background: #fc8181; }

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
# HELPER: upload widget wrapped in styled card
# ─────────────────────────────────────────────
def upload_card(label: str, key: str, icon: str = "📄") -> object:
    """Renders a cohesive styled upload widget."""
    st.file_uploader(
        label=f"{icon}  {label}",
        type=["csv", "xlsx", "xls", "xlsb"],
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
# LIVE UPLOAD PROGRESS
# ─────────────────────────────────────────────
uploaded_count = sum(
    1 for k in FILE_KEYS if st.session_state.get(k) not in (None, [])
)
pct = int((uploaded_count / TOTAL_FILES) * 100)

col_prog_1, col_prog_2 = st.columns([1, 5])

with col_prog_1:
    st.markdown(
        f"""
        <div style="text-align:center; background:rgba(255,255,255,0.04);
                    border:1px solid rgba(255,255,255,0.08); border-radius:12px;
                    padding: 1rem 0.5rem;">
            <div style="font-size:2.2rem;font-weight:800;color:#63b3ed;">{uploaded_count}/{TOTAL_FILES}</div>
            <div style="font-size:0.75rem;color:#718096;margin-top:4px;">Files Uploaded</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_prog_2:
    st.progress(pct / 100, text=f"Upload Progress — {pct}% complete")
    # Mini status checklist (2 rows for readability with 9 files)
    status_html = ""
    for k, label in FILE_LABELS.items():
        dot_cls = "uploaded" if st.session_state.get(k) not in (None, []) else "pending"
        status_html += (
            f'<span style="margin-right:14px;font-size:0.75rem;color:#a0aec0;">'
            f'<span class="status-dot {dot_cls}"></span>{label}</span>'
        )
    st.markdown(
        f'<div style="margin-top:8px;line-height:2.2;">{status_html}</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# DIVIDER
# ─────────────────────────────────────────────
st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# GROUP 1 — YESTERDAY'S PLANS
# ─────────────────────────────────────────────
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
# GROUP 2 — TODAY'S LISTS
# ─────────────────────────────────────────────
st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
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


# ─────────────────────────────────────────────
# GROUP 3 — REPORTS & WIP
# ─────────────────────────────────────────────
st.markdown('<div class="styled-divider"></div>', unsafe_allow_html=True)
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
    upload_card(
        "Paint Float Report",
        key="paint_float_report",
        icon="🎨",
    )
with g3_row1_col2:
    upload_card(
        "WIP — Q5",
        key="wip_q5",
        icon="📈",
    )

g3_row2_col1, g3_row2_col2 = st.columns(2, gap="large")
with g3_row2_col1:
    upload_card(
        "Nova — WIP",
        key="nova_wip",
        icon="🔧",
    )
with g3_row2_col2:
    upload_card(
        "Pending Plan for BIW",
        key="pending_plan_biw",
        icon="🏗️",
    )

g3_row3_col1, g3_row3_col2 = st.columns(2, gap="large")
with g3_row3_col1:
    upload_card(
        "X1 BIW WIP",
        key="x1_biw_wip",
        icon="🏭",
    )
with g3_row3_col2:
    upload_card(
        "Next 3-Days BIW Plan",
        key="next_3days_biw_plan",
        icon="📋",
    )




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
        if st.session_state.get(k) in (None, [])
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
            
            yest_tcf1 = st.session_state.get("vin_plan_tcf1_yesterday")
            yest_tcf2 = st.session_state.get("vin_plan_tcf2_yesterday")
            today_list_tcf1 = st.session_state.get("vin_list_tcf1_today")
            today_list_tcf2 = st.session_state.get("vin_list_tcf2_today")
            paint_float = st.session_state.get("paint_float_report")
            wip_q5 = st.session_state.get("wip_q5")
            nova_wip = st.session_state.get("nova_wip")
            x1_biw_wip = st.session_state.get("x1_biw_wip")
            
            next_3days_biw_plan = st.session_state.get("next_3days_biw_plan")
            
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
                
            tcf1_wb = update_paint_float_data(
                tcf1_wb, paint_float, "TCF1", 
                expected_qty=expected_qty_tcf1,
                wip_files=tcf1_wip_files,
                yest_plan_file_or_stream=yest_tcf1,
                next_3days_biw_plan=next_3days_biw_plan
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
                
            tcf2_wb = update_paint_float_data(
                tcf2_wb, paint_float, "TCF2", 
                expected_qty=expected_qty_tcf2,
                wip_files=tcf2_wip_files,
                yest_plan_file_or_stream=yest_tcf2,
                next_3days_biw_plan=next_3days_biw_plan
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
    
    st.success(
        f"✅  **3-Day Production Plans generated successfully!**\n\n"
        f"- TCF-1 (Punch/Altroz): Daily target capacity `900`. Day 1 Sequence completed, Days 2 & 3 generated.\n"
        f"- TCF-2 (Harrier/Safari): Daily target capacity `250`. Day 1 Sequence completed, Days 2 & 3 generated.",
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
        <div style="background:rgba(255,255,255,0.02); border:1px solid rgba(255,255,255,0.08); border-radius:12px; padding:1.2rem; margin-bottom:1.5rem;">
            <h4 style="margin:0 0 1rem 0; color:#38BDF8; font-size:1.1rem; display:flex; align-items:center; gap:8px;">
                🎨 Planned Vehicle Count by Source Type (Color-Coded)
            </h4>
            <div style="display:flex; justify-content:space-between; gap:20px; flex-wrap:wrap;">
                <!-- TCF-1 Column -->
                <div style="flex:1; min-width:280px; background:rgba(255,255,255,0.01); border-radius:8px; padding:1rem; border:1px solid rgba(255,255,255,0.03);">
                    <h5 style="margin:0 0 0.8rem 0; color:#90cdf4; font-size:0.95rem;">TCF-1 (Punch/Altroz)</h5>
                    <div style="display:flex; flex-direction:column; gap:8px; font-size:0.85rem;">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#E2EFDA; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Paint Floor Float (Light Green):</span>
                            <strong>{t1['colors']['light_green']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#DDEBF7; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Critical WIPs (Light Blue):</span>
                            <strong>{t1['colors']['light_blue']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FFF2CC; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Pending BIW Backlog (Light Yellow):</span>
                            <strong>{t1['colors']['light_yellow']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FCE4D6; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Next 3-Days BIW Plan (Light Pink):</span>
                            <strong>{t1['colors']['light_pink']}</strong>
                        </div>
                    </div>
                </div>
                <!-- TCF-2 Column -->
                <div style="flex:1; min-width:280px; background:rgba(255,255,255,0.01); border-radius:8px; padding:1rem; border:1px solid rgba(255,255,255,0.03);">
                    <h5 style="margin:0 0 0.8rem 0; color:#98db9c; font-size:0.95rem;">TCF-2 (Harrier/Safari)</h5>
                    <div style="display:flex; flex-direction:column; gap:8px; font-size:0.85rem;">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#E2EFDA; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Paint Floor Float (Light Green):</span>
                            <strong>{t2['colors']['light_green']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#DDEBF7; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Critical WIPs (Light Blue):</span>
                            <strong>{t2['colors']['light_blue']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FFF2CC; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Pending BIW Backlog (Light Yellow):</span>
                            <strong>{t2['colors']['light_yellow']}</strong>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span><span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#FCE4D6; margin-right:8px; border:1px solid rgba(0,0,0,0.15);"></span>Next 3-Days BIW Plan (Light Pink):</span>
                            <strong>{t2['colors']['light_pink']}</strong>
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
        <div style="background:rgba(16,185,129,0.05); border:1px solid rgba(16,185,129,0.2); border-radius:12px; padding:1.2rem; margin-bottom:1.5rem;">
            <h4 style="margin:0 0 0.5rem 0; color:#10B981; font-size:1.1rem; display:flex; align-items:center; gap:8px;">
                📊 Model Plan Summary (NOVA, CNG, Eturna)
            </h4>
            <p style="margin:0; font-size:0.85rem; color:#a0aec0; line-height:1.4;">
                Highlighting the number of constrained model units planned across yesterday's float drops, today, and the next 3 days.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    sum_col1, sum_col2, sum_col3 = st.columns(3, gap="large")
    
    with sum_col1:
        st.markdown(
            f"""
            <div style="background:rgba(99,179,237,0.03); border:1px solid rgba(99,179,237,0.1); border-radius:15px; padding:1.5rem; text-align:center; box-shadow: 0 10px 30px rgba(0,0,0,0.2);">
                <div style="font-size:2.5rem; margin-bottom:0.5rem;">⚡</div>
                <h4 style="color:#63b3ed; margin:0 0 0.2rem 0; font-size:1.2rem;">NOVA (Punch EV)</h4>
                <p style="color:#a0aec0; font-size:0.75rem; margin:0 0 1rem 0;">Max Limit: 150 units/day</p>
                <div style="text-align:left; font-size:0.85rem; color:#e2e8f0; background:rgba(255,255,255,0.02); border-radius:10px; padding:0.8rem;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; border-bottom:1px solid rgba(255,255,255,0.05); padding-bottom:0.3rem;">
                        <span>B&C Drops Today:</span> <strong>{t1['picked']['nova']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Today (Day 1):</span> <span><strong>{t1['day1']['nova']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 150</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t1['day2']['nova']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 150</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 3 Plan:</span> <span><strong>{t1['day3']['nova']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 150</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 4 Plan:</span> <span><strong>{t1['day4']['nova']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 150</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
    with sum_col2:
        st.markdown(
            f"""
            <div style="background:rgba(237,137,54,0.03); border:1px solid rgba(237,137,54,0.1); border-radius:15px; padding:1.5rem; text-align:center; box-shadow: 0 10px 30px rgba(0,0,0,0.2);">
                <div style="font-size:2.5rem; margin-bottom:0.5rem;">🔥</div>
                <h4 style="color:#ed8936; margin:0 0 0.2rem 0; font-size:1.2rem;">Punch CNG</h4>
                <p style="color:#a0aec0; font-size:0.75rem; margin:0 0 1rem 0;">Max Limit: 350 units/day</p>
                <div style="text-align:left; font-size:0.85rem; color:#e2e8f0; background:rgba(255,255,255,0.02); border-radius:10px; padding:0.8rem;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; border-bottom:1px solid rgba(255,255,255,0.05); padding-bottom:0.3rem;">
                        <span>B&C Drops Today:</span> <strong>{t1['picked']['cng']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Today (Day 1):</span> <span><strong>{t1['day1']['cng']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 350</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t1['day2']['cng']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 350</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 3 Plan:</span> <span><strong>{t1['day3']['cng']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 350</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between;">
                        <span>Day 4 Plan:</span> <span><strong>{t1['day4']['cng']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 350</span></span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        
    with sum_col3:
        st.markdown(
            f"""
            <div style="background:rgba(72,187,120,0.03); border:1px solid rgba(72,187,120,0.1); border-radius:15px; padding:1.5rem; text-align:center; box-shadow: 0 10px 30px rgba(0,0,0,0.2);">
                <div style="font-size:2.5rem; margin-bottom:0.5rem;">🌱</div>
                <h4 style="color:#48bb78; margin:0 0 0.2rem 0; font-size:1.2rem;">Eturna (Harrier EV)</h4>
                <p style="color:#a0aec0; font-size:0.75rem; margin:0 0 1rem 0;">Max Limit: 160 units/day</p>
                <div style="text-align:left; font-size:0.85rem; color:#e2e8f0; background:rgba(255,255,255,0.02); border-radius:10px; padding:0.8rem;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; border-bottom:1px solid rgba(255,255,255,0.05); padding-bottom:0.3rem;">
                        <span>B&C Drops Today:</span> <strong>{t2['picked']['eturna']}</strong>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Today (Day 1):</span> <span><strong>{t2['day1']['eturna']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 160</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 2 Plan:</span> <span><strong>{t2['day2']['eturna']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 160</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.4rem;">
                        <span>Day 3 Plan:</span> <span><strong>{t2['day3']['eturna']}</strong> <span style="color:#a0aec0;font-size:0.75rem;">/ 160</span></span>
                    </div>
                    <div style="display:flex; justify-content:space-between; color:#718096;">
                        <span>Day 4 Plan:</span> <span>Not Applicable</span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        



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
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 20px;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-end;
            padding: 25px 20px;
            box-sizing: border-box;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.4);
            transition: 0.5s;
            transform-style: preserve-3d;
        }}
        .card:hover {{
            border-color: rgba(99, 179, 237, 0.4);
            box-shadow: 0 25px 50px rgba(99, 179, 237, 0.15);
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
            color: #a0aec0;
            margin: 10px 0 0 0;
            line-height: 1.4;
        }}
        .card .badge {{
            display: inline-block;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.7rem;
            font-weight: 600;
            margin-top: 15px;
            text-transform: uppercase;
            letter-spacing: 1px;
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
