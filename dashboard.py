import streamlit as st
import plotly.express as px
from datetime import datetime
import requests
from utils.excel_reader import load_data, get_weather
from utils.voice import listen
import streamlit.components.v1 as components
from rapidfuzz import process, fuzz
from io import BytesIO
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image
from utils.pdf_report import generate_executive_pdf
import os
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table as PDFTable,
    TableStyle,
    PageBreak,
    Image as PDFImage
)
from reportlab.lib.units import inch





def speak(text):

    text = str(text).replace('"', '\\"')

    components.html(
        f"""
        <script>

        window.speechSynthesis.cancel();

        var msg = new SpeechSynthesisUtterance("{text}");

        msg.lang = "en-US";
        msg.rate = 0.95;
        msg.pitch = 1;
        msg.volume = 1;

        window.speechSynthesis.speak(msg);

        </script>
        """,
        height=0,
    )

# =====================================================
# Normalize Voice
# =====================================================

def normalize_voice(text):

    if not text:
        return ""

    text = str(text).lower().strip()

    aliases = {

        # ==========================
        # Staff
        # ==========================

        "sada asif": "saad asif",
        "sad asif": "saad asif",
        "saada asif": "saad asif",
        "saad asifa": "saad asif",
        "saad asif": "saad asif",

        "show all projects": "all",
        "show all": "all",
        "all projects": "all",

        "afzal": "afzal",

        "adnan": "adnan",
        "adnan younis": "adnan younis",

        "mobashir": "mubashir",
        "mubashar": "mubashir",
        "mubashir": "mubashir",

        "bhawan": "bhawan",
        "bawan": "bhawan",
        "bhavan": "bhawan",

        "ayesha": "ayesha",

        # ==========================
        # Status
        # ==========================

        "life": "live",
        "live": "live",

        "uat": "uat",
        "sit": "sit",

        "scoping": "under scoping",
        "under scoping": "under scoping",

        "development": "under development",
        "under development": "under development",

        "process": "in process",
        "in process": "in process",

        # ==========================
        # Projects
        # ==========================

        "swap": "swap",
        "swaps": "swap",

        "ata": "cnic screening -ata",

        "siem": "siem intergration",
        "cm": "siem intergration",

        "soap": "discontinuation of soap to rest",

        "token": "tokinization on smartpay",
        "tokenization": "tokinization on smartpay",

        "estamping": "e-stamping",
        "e stamping": "e-stamping",

        "outward": "outward clearing",

        "gbm": "gbm",

        "optimization": "optimization",

        "rf": "rf account on smartpay",
        "show live projects": "live",
        "live projects": "live",
        "only live": "live",

        "show uat projects": "uat",
        "uat projects": "uat",

        "show sit projects": "sit",
        "sit projects": "sit",

        "show development projects": "under development",
        "under development": "under development",

        "show all projects": "all",
        "all projects": "all",
         # =====================================
        # Voice Commands
        # =====================================

        "show all projects": "all",
        "show all": "all",
        "all projects": "all",
        "all": "all",

        "show live projects": "live",
        "live projects": "live",
        "only live": "live",

        "show uat projects": "uat",
        "uat projects": "uat",

        "show sit projects": "sit",
        "sit projects": "sit",

        "show development projects": "under development",
        "development projects": "under development",

        # ==========================
        # Commands
        # ==========================

        "show all projects": "all",
        "show all": "all",
        "all projects": "all",
        "all": "all",
    }

    if text in aliases:
        return aliases[text]
        # Partial Match
    for key, value in aliases.items():
        if key in text:
            return value

    return text

# =====================================================
# Page Settings
# =====================================================
st.set_page_config(
    page_title="SmartPay Project Dashboard",
    page_icon="💳",
    layout="wide"
)


st.markdown("""
<style>

/* ===========================
   Main App
=========================== */

.stApp{
    background:#F8F9FA;
}
/* ==========================================
   COMPACT TOP SPACE
========================================== */

.block-container {
    padding-top: 2.5rem !important;
}

/* ===========================
   Headings
=========================== */

h1,h2,h3,h4,h5,h6,
[data-testid="stHeading"],
[data-testid="stHeading"] *{
    color:#006747 !important;
    font-weight:700 !important;
}

/* ===========================
   Normal Text
=========================== */

p,
span{
    color:inherit !important;
}

/* ===========================
   Field Labels
=========================== */

label,
label p,
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] p{
    color:#006747 !important;
    font-weight:bold !important;
    font-size:16px !important;
}

/* ===========================
   Sidebar
=========================== */

/* ==========================================
   PREMIUM SIDEBAR
========================================== */

section[data-testid="stSidebar"]{
    background:linear-gradient(180deg,#004B34,#006747,#008A5A) !important;
    border-right:3px solid #D4AF37;
}

section[data-testid="stSidebar"] *{
    color:white !important;
    font-family:"Segoe UI",sans-serif !important;
}

/* Navigation Title */

section[data-testid="stSidebar"] label{
    color:#FFD700 !important;
    font-size:18px !important;
    font-weight:700 !important;
}

/* Radio Buttons */

div[role="radiogroup"] label{
    background:rgba(255,255,255,.08);
    margin-bottom:8px;
    padding:10px;
    border-radius:10px;
    transition:.3s;
}

div[role="radiogroup"] label:hover{
    background:rgba(255,255,255,.20);
}

/* Selected Page */

div[role="radiogroup"] label[data-selected="true"]{
    background:white !important;
    color:#006747 !important;
    font-weight:bold !important;
    border-left:5px solid #FFD700;
}

/* ===========================
   KPI Cards
=========================== */

div[data-testid="metric-container"]{
    background:white !important;
    border-left:6px solid #006747 !important;
    border-radius:12px !important;
    padding:18px !important;
    box-shadow:0 2px 8px rgba(0,0,0,.15);
}

[data-testid="stMetricLabel"]{
    color:#006747 !important;
    font-weight:bold !important;
}

[data-testid="stMetricValue"]{
    color:#222222 !important;
    font-size:34px !important;
    font-weight:700 !important;
}

[data-testid="stMetricDelta"]{
    color:#006747 !important;
}

/* ===========================
   Buttons
=========================== */

.stButton>button{
    background:#006747 !important;
    color:white !important;
    border:none !important;
    border-radius:8px !important;
    font-weight:bold !important;
}

.stButton>button:hover{
    background:#008A5A !important;
}

/* ===========================
   Text Input
=========================== */

.stTextInput input{
    background:white !important;
    color:black !important;
    border:2px solid #006747 !important;
    border-radius:8px !important;
}

/* ===========================
   Select Box
=========================== */

div[data-baseweb="select"]>div{
    background:white !important;
    color:black !important;
    border:2px solid #006747 !important;
    border-radius:8px !important;
}

/* ===========================
   DataFrame
=========================== */

[data-testid="stDataFrame"]{
    border:2px solid #006747 !important;
    border-radius:10px !important;
}

/* ===========================
   Alert Boxes
=========================== */

[data-testid="stAlert"]{
    border-radius:10px !important;
}

[data-testid="stAlert"] *{
    color:#222222 !important;
}

</style>
""", unsafe_allow_html=True)
# =====================================================
# Load Data
# =====================================================
df = load_data()

# Clean Status
df["Status"] = (
    df["Status"]
    .astype(str)
    .str.strip()
    .replace({
        "Under development": "Under Development",
        "InProcess": "In Process",
        "LIVE ": "LIVE"
    })
)

# =====================================================
# SIDEBAR
# =====================================================

st.markdown("""
<style>

/* =========================
   Sidebar
========================= */

section[data-testid="stSidebar"]{
    background:#006747;
}

/* Navigation title */
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p{
    color:white;
}

/* Radio Labels */
div[role="radiogroup"] label{
    background:transparent !important;
    border:none !important;
    border-radius:10px;
    padding:10px 12px;
    margin-bottom:6px;
    color:white !important;
    font-weight:600;
    transition:.2s;
}

/* Hover */
div[role="radiogroup"] label:hover{
    background:#0B8758 !important;
}

/* Selected */
div[role="radiogroup"] label:has(input:checked){
    background:#0FA968 !important;
    color:white !important;
}

/* Hide Radio Circle */
div[role="radiogroup"] input{
    display:none !important;
}

/* Hide Empty Header Space */
section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"]{
    padding-top:.4rem;
}

</style>
""", unsafe_allow_html=True)

# =====================================================
# LOGO
# =====================================================

st.sidebar.image(
    "smartpay_logo/smartpay_logo.png",
    use_container_width=True
)

# =====================================================
# TITLE
# =====================================================

st.sidebar.markdown("""
<h2 style="
text-align:center;
color:white;
margin-bottom:0;">
🏦 SmartPay
</h2>

<p style="
text-align:center;
color:#D1FAE5;
margin-top:2px;
margin-bottom:8px;">
Project Dashboard
</p>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")

# =====================================================
# NAVIGATION
# =====================================================

# Dashboard card se navigation request aaye to
if "navigate_to" in st.session_state:

    st.session_state["page_navigation"] = st.session_state["navigate_to"]

    del st.session_state["navigate_to"]


page = st.sidebar.radio(
    "Menu",
    [
        "Dashboard",
        "Projects",
        "Analytics",
        "Project Timeline",
        "BAU Monitoring",
        "CRPL",
        "PAYSYS",
        "Export"
    ],
    label_visibility="collapsed",
    key="page_navigation"
)

st.sidebar.markdown("---")
# ==========================================
# RESET SEARCH + SCROLL WHEN PAGE CHANGES
# ==========================================

if "previous_page" not in st.session_state:
    st.session_state.previous_page = page

elif st.session_state.previous_page != page:

    # Search / filter reset
    for key in [
        "project_search",
        "project",
        "search",
        "search_project",
        "selected_project",
        "selected_member"
    ]:
        if key in st.session_state:
            del st.session_state[key]

    # New page remember
    st.session_state.previous_page = page

    # Scroll to top
    # Scroll position reset
    st.markdown("""
    <script>
    window.parent.scrollTo(0, 0);
    window.parent.document.documentElement.scrollTop = 0;
    window.parent.document.body.scrollTop = 0;
    </script>
    """, unsafe_allow_html=True)

    st.rerun()

# =====================================================
# INFORMATION
# =====================================================

st.sidebar.markdown("""
<div style="
background:rgba(255,255,255,.12);
padding:12px;
border-radius:12px;
color:white;
font-size:14px;
line-height:1.5;">

<b>Department</b><br>
Digital Banking Group

<hr style="margin:8px 0;border:.5px solid rgba(255,255,255,.25);">

<b>Organization</b><br>
National Bank of Pakistan

<hr style="margin:8px 0;border:.5px solid rgba(255,255,255,.25);">

<b>Version</b><br>
2.0

<hr style="margin:8px 0;border:.5px solid rgba(255,255,255,.25);">

<b>Developer</b><br>
Muhammad Saad Asif

</div>
""", unsafe_allow_html=True)

st.sidebar.markdown("<br>", unsafe_allow_html=True)

st.sidebar.caption(
    "© 2026 SmartPay Dashboard"
)
# =====================================================
# DASHBOARD
# =====================================================

if page == "Dashboard":

    # ==========================================
    # TITLE - COMPACT HEADER
    # ==========================================

    st.markdown("""
    <div style="
    background:linear-gradient(135deg,#ffffff,#f8fbff);
    border-radius:18px;
    padding:14px 22px;
    border:1px solid #E5E7EB;
    box-shadow:0 8px 22px rgba(0,0,0,.08);
    box-sizing:border-box;
    margin-bottom:10px;
    ">

    <div style="
    color:#006747;
    font-size:12px;
    font-weight:700;
    letter-spacing:1.8px;
    margin-bottom:5px;">
    NATIONAL BANK OF PAKISTAN
    </div>

    <div style="
    font-size:28px;
    color:#006747;
    font-weight:800;
    line-height:1.15;
    margin:0;">
    SmartPay Project Dashboard
    </div>

    <div style="
    margin-top:5px;
    font-size:14px;
    color:#374151;">
    Digital Banking Group
    </div>

    <div style="
    margin-top:8px;
    display:inline-block;
    background:#ECFDF5;
    color:#006747;
    padding:4px 11px;
    border-radius:20px;
    font-size:11px;
    font-weight:700;">
    ● LIVE Dashboard
    </div>

    </div>
    """, unsafe_allow_html=True)


    # =====================================================
    # KPI - VIP ENTERPRISE CLICKABLE CARDS
    # =====================================================

    status = df["Status"].astype(str).str.upper().str.strip()

    total_projects = len(df)

    scoping_projects = len(
        df[status.isin(["SCOPING", "UNDER SCOPING"])]
    )

    uat_projects = len(
        df[status == "UAT"]
    )

    review_projects = len(
        df[status == "IS REVIEW"]
    )

    cmc_projects = len(
        df[status == "CMC"]
    )

    live_projects = len(
        df[status == "LIVE"]
    )

    bau_projects = len(
        df[status == "BAU"]
    )


    # =====================================================
    # VIP CARD CSS
    # =====================================================

    st.markdown("""
    <style>

    /* =========================================
       KPI CARD CONTAINER
    ========================================= */

    .st-key-dashboard_kpis div[data-testid="stButton"] {
        width:100% !important;
    }


    /* =========================================
       BASE KPI CARD
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stButton"] button {

        width:100% !important;

        min-height:96px !important;

        height:96px !important;

        background:#FFFFFF !important;

        border:1px solid #E5E7EB !important;

        border-radius:16px !important;

        padding:10px 6px !important;

        box-shadow:
            0 6px 16px rgba(0,0,0,.08) !important;

        color:#111827 !important;

        font-size:13px !important;

        font-weight:600 !important;

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        line-height:1.2 !important;

        white-space:pre-line !important;

        text-align:center !important;

        display:flex !important;

        align-items:center !important;

        justify-content:center !important;

        transition:all .2s ease !important;
    }


    /* =========================================
       BUTTON TEXT
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stButton"] button p {

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        font-size:13px !important;

        font-weight:650 !important;

        line-height:1.3 !important;

        white-space:pre-line !important;

        text-align:center !important;

        display:block !important;

        width:100% !important;

        margin:0 !important;

        padding:0 !important;

        color:#111827 !important;
    }


    /* =========================================
       HOVER
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stButton"] button:hover {

        background:#F8FAFC !important;

        transform:translateY(-2px) !important;

        box-shadow:
            0 10px 22px rgba(0,0,0,.11) !important;
    }


    /* =========================================
       FOCUS / SELECTED
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stButton"] button:focus {

        outline:none !important;

        min-height:96px !important;

        height:96px !important;

        background:#EAF5F0 !important;

        border:2px solid #006747 !important;

        box-shadow:
            0 0 0 3px rgba(0,103,71,0.12),
            0 10px 22px rgba(0,103,71,0.16) !important;

        color:#006747 !important;

        transform:translateY(-2px) !important;
    }


    .st-key-dashboard_kpis
    div[data-testid="stButton"] button:focus p {

        color:#006747 !important;

        font-weight:750 !important;
    }


    /* =========================================
       CARD 1 - TOTAL
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(1)
    div[data-testid="stButton"] button {

        border-top:6px solid #006747 !important;
    }


    /* =========================================
       CARD 2 - SCOPING
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(2)
    div[data-testid="stButton"] button {

        border-top:6px solid #8E24AA !important;
    }


    /* =========================================
       CARD 3 - UAT
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(3)
    div[data-testid="stButton"] button {

        border-top:6px solid #F9A825 !important;
    }


    /* =========================================
       CARD 4 - IS REVIEW
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(4)
    div[data-testid="stButton"] button {

        border-top:6px solid #00ACC1 !important;
    }


    /* =========================================
       CARD 5 - CMC
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(5)
    div[data-testid="stButton"] button {

        border-top:6px solid #3949AB !important;
    }


    /* =========================================
       CARD 6 - LIVE
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(6)
    div[data-testid="stButton"] button {

        border-top:6px solid #00C853 !important;
    }


    /* =========================================
       CARD 7 - BAU
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(7)
    div[data-testid="stButton"] button {

        border-top:6px solid #607D8B !important;
    }


    /* =========================================
       REMOVE EXTRA GAP
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stVerticalBlock"] {

        gap:0 !important;
    }


    /* =========================================
       COLUMN SPACING
    ========================================= */

    .st-key-dashboard_kpis
    div[data-testid="stHorizontalBlock"] {

        gap:10px !important;

        align-items:stretch !important;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # CLICKABLE KPI CARDS
    # =====================================================

    with st.container(key="dashboard_kpis"):

        c1, c2, c3, c4, c5, c6, c7 = st.columns(
            [1.2, 1.1, 1.0, 1.1, 1.0, 1.0, 1.0]
        )


        # -----------------------------------------
        # TOTAL PROJECTS
        # -----------------------------------------

        with c1:

            if st.button(
                f"TOTAL PROJECTS\n{total_projects}",
                key="dashboard_total",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "All"

                st.session_state["navigate_to"] = "Projects"

                st.rerun()


        # -----------------------------------------
        # SCOPING
        # -----------------------------------------

        with c2:

            if st.button(
                f"SCOPING\n{scoping_projects}",
                key="dashboard_scoping",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "SCOPING"

                st.session_state["navigate_to"] = "Projects"

                st.rerun()


        # -----------------------------------------
        # UAT
        # -----------------------------------------

        with c3:

            if st.button(
                f"UAT\n{uat_projects}",
                key="dashboard_uat",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "UAT"

                st.session_state["navigate_to"] = "Projects"

                st.rerun()


        # -----------------------------------------
        # IS REVIEW
        # -----------------------------------------

        with c4:

            if st.button(
                f"IS REVIEW\n{review_projects}",
                key="dashboard_review",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "IS REVIEW"

                st.session_state["navigate_to"] = "Projects"

                st.rerun()


        # -----------------------------------------
        # CMC
        # -----------------------------------------

        with c5:

            if st.button(
                f"CMC\n{cmc_projects}",
                key="dashboard_cmc",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "CMC"

                st.session_state["navigate_to"] = "Projects"

                st.rerun()


        # -----------------------------------------
        # LIVE
        # -----------------------------------------

        with c6:

            if st.button(
                f"LIVE\n{live_projects}",
                key="dashboard_live",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "LIVE"

                st.session_state["navigate_to"] = "Projects"

                st.rerun()


        # -----------------------------------------
        # BAU
        # -----------------------------------------

        with c7:

            if st.button(
                f"BAU\n{bau_projects}",
                key="dashboard_bau",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "BAU"

                st.session_state["navigate_to"] = "Projects"

                st.rerun()


    st.markdown("<br>", unsafe_allow_html=True)
    # =====================================================
    # TEAM OVERVIEW
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:26px;
    font-weight:700;
    margin-bottom:10px;">
    👥 Team Overview
    </h2>
    """, unsafe_allow_html=True)


    allocation = (
        df["Allocation"]
        .value_counts()
        .reset_index()
    )

    allocation.columns = ["Allocation", "Projects"]


    cols = st.columns(4)

    for i, row in allocation.iterrows():

        with cols[i % 4]:

            st.markdown(f"""
    <div style="
    background:linear-gradient(180deg,#ffffff,#f7f9fc);
    border-radius:14px;
    padding:9px 8px;
    text-align:center;
    border:1px solid #E5E7EB;
    box-shadow:0 4px 12px rgba(0,0,0,.05);
    min-height:100px;
    box-sizing:border-box;">

    <div style="
    width:34px;
    height:34px;
    border-radius:50%;
    background:#E8F5E9;
    margin:auto;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:17px;">
    👤
    </div>

    <div style="
    margin-top:5px;
    font-size:14px;
    font-weight:700;
    line-height:1.2;
    color:#006747;
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;">
    {row["Allocation"]}
    </div>

    <div style="
    margin-top:2px;
    font-size:26px;
    font-weight:800;
    line-height:1;
    color:#111827;">
    {row["Projects"]}
    </div>

    <div style="
    margin-top:2px;
    font-size:10px;
    line-height:1.1;
    color:#6B7280;">
    Projects Assigned
    </div>

    </div>
    """, unsafe_allow_html=True)


    st.markdown("<br>", unsafe_allow_html=True)


    # =====================================================
    # TEAM WORKLOAD - COMPACT
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:24px;
    font-weight:700;
    margin-top:8px;
    margin-bottom:10px;">
    📊 Team Workload
    </h2>
    """, unsafe_allow_html=True)


    team_df = (
        df.groupby("Allocation")
        .size()
        .reset_index(name="Projects")
        .sort_values("Projects", ascending=False)
    )


    fig = px.bar(
        team_df,
        x="Allocation",
        y="Projects",
        text="Projects",
        color="Projects",
        color_continuous_scale="Greens"
    )


    fig.update_traces(
        textposition="outside",
        marker_line_width=0,
        textfont=dict(size=12)
    )


    fig.update_layout(
        height=380,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(
            l=15,
            r=15,
            t=8,
            b=15
        ),
        coloraxis_showscale=False,
        xaxis_title="",
        yaxis_title="Projects",
        font=dict(size=12),
        xaxis=dict(
            showgrid=False,
            tickfont=dict(size=12)
        ),
        yaxis=dict(
            gridcolor="#ECECEC",
            tickfont=dict(size=11)
        )
    )


    st.plotly_chart(
        fig,
        width="stretch"
    )

    # =====================================================
    # TEAM SUMMARY - VIP
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:34px;
    font-weight:700;
    margin-top:30px;
    margin-bottom:20px;">
    📋 Team Summary
    </h2>
    """, unsafe_allow_html=True)


    summary = (
        df.groupby("Allocation")
        .agg(
            Total=("Mandate", "count"),
            Live=("Status", lambda x: (x.str.upper() == "LIVE").sum()),
            UAT=("Status", lambda x: (x.str.upper() == "UAT").sum())
        )
        .reset_index()
    )


    # -----------------------------
    # TEAM SUMMARY TABLE CSS
    # -----------------------------

    st.markdown("""
    <style>

    .vip-summary-wrapper {
        background: #FFFFFF;
        border: 1px solid #DDE5E1;
        border-radius: 18px;
        padding: 6px;
        box-shadow: 0 8px 25px rgba(0,103,71,0.08);
        overflow: hidden;
    }

    .vip-summary-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 14px;
    }

    .vip-summary-table thead th {
        background: #006747;
        color: #FFFFFF;
        padding: 15px 16px;
        font-weight: 700;
        text-align: left;
    }

    .vip-summary-table thead th:first-child {
        border-top-left-radius: 12px;
    }

    .vip-summary-table thead th:last-child {
        border-top-right-radius: 12px;
    }

    .vip-summary-table tbody td {
        padding: 15px 16px;
        color: #1F2937;
        border-bottom: 1px solid #E5E7EB;
        background: #FFFFFF;
    }

    .vip-summary-table tbody tr:nth-child(even) td {
        background: #F8FAFC;
    }

    .vip-summary-table tbody tr:hover td {
        background: #ECFDF5;
    }

    .team-name {
        color: #006747 !important;
        font-weight: 800;
    }

    .total-count {
        background: #F0FDF4;
        color: #006747 !important;
        font-weight: 800;
        padding: 5px 10px;
        border-radius: 20px;
    }

    .live-count {
        background: #DCFCE7;
        color: #166534 !important;
        font-weight: 800;
        padding: 5px 10px;
        border-radius: 20px;
    }

    .uat-count {
        background: #FEF3C7;
        color: #92400E !important;
        font-weight: 800;
        padding: 5px 10px;
        border-radius: 20px;
    }

    </style>
    """, unsafe_allow_html=True)
  


    # =====================================================
    # CREATE SUMMARY
    # =====================================================

    summary = (
        df.groupby("Allocation")
        .agg(

            # Total Projects
            Total=("Mandate", "count"),

            # SCOPING
            Scoping=("Status", lambda x:
                x.astype(str)
                .str.strip()
                .str.upper()
                .isin([
                    "SCOPING",
                    "UNDER SCOPING"
                ])
                .sum()
            ),

            # DEVELOPMENT
            Development=("Status", lambda x:
                x.astype(str)
                .str.strip()
                .str.upper()
                .isin([
                    "DEVELOPMENT",
                    "UNDER DEVELOPMENT",
                    "SIT"
                ])
                .sum()
            ),

            # UAT
            UAT=("Status", lambda x:
                (
                    x.astype(str)
                    .str.strip()
                    .str.upper()
                    == "UAT"
                ).sum()
            ),

            # IS REVIEW
            IS_Review=("Status", lambda x:
                (
                    x.astype(str)
                    .str.strip()
                    .str.upper()
                    == "IS REVIEW"
                ).sum()
            ),

            # CMC
            CMC=("Status", lambda x:
                (
                    x.astype(str)
                    .str.strip()
                    .str.upper()
                    == "CMC"
                ).sum()
            ),

            # LIVE
            Live=("Status", lambda x:
                (
                    x.astype(str)
                    .str.strip()
                    .str.upper()
                    == "LIVE"
                ).sum()
            )
        )
        .reset_index()
    )


    # =====================================================
    # RENAME COLUMN
    # =====================================================

    summary.rename(
        columns={
            "IS_Review": "IS Review"
        },
        inplace=True
    )


    # =====================================================
    # EXACT COLUMN ORDER
    # =====================================================

    summary = summary[
        [
            "Allocation",
            "Total",
            "Scoping",
            "UAT",
            "IS Review",
            "CMC",
            "Live"
        ]
    ]


    # =====================================================
    # FORMAT SUMMARY
    # =====================================================

    summary_display = summary.copy()


    # Allocation
    summary_display["Allocation"] = summary_display[
        "Allocation"
    ].apply(
        lambda x:
        f'<span class="team-name">👤 {x}</span>'
    )


    # Total
    summary_display["Total"] = summary_display[
        "Total"
    ].apply(
        lambda x:
        f'<span class="total-count">{x}</span>'
    )


    # Scoping
    summary_display["Scoping"] = summary_display[
        "Scoping"
    ].apply(
        lambda x:
        f'<span class="scoping-count">🟠 {x}</span>'
    )





    # UAT
    summary_display["UAT"] = summary_display[
        "UAT"
    ].apply(
        lambda x:
        f'<span class="uat-count">🟡 {x}</span>'
    )


    # IS Review
    summary_display["IS Review"] = summary_display[
        "IS Review"
    ].apply(
        lambda x:
        f'<span class="review-count">🔷 {x}</span>'
    )


    # CMC
    summary_display["CMC"] = summary_display[
        "CMC"
    ].apply(
        lambda x:
        f'<span class="cmc-count">🟣 {x}</span>'
    )


    # Live
    summary_display["Live"] = summary_display[
        "Live"
    ].apply(
        lambda x:
        f'<span class="live-count">🟢 {x}</span>'
    )


    # =====================================================
    # CREATE HTML TABLE
    # =====================================================

    summary_html = summary_display.to_html(
        index=False,
        escape=False,
        classes="vip-summary-table"
    )


    # =====================================================
    # DISPLAY VIP TABLE
    # =====================================================

    st.markdown(
        f"""
        <div class="vip-summary-wrapper">
            {summary_html}
        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown("<br>", unsafe_allow_html=True)
    # =====================================================
    # SMART SEARCH - VIP PROJECT TABLE
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:34px;
    font-weight:700;
    margin-top:35px;
    margin-bottom:20px;">
    🔍 Smart Search
    </h2>
    """, unsafe_allow_html=True)


    project = st.text_input(
        "Search Project",
        placeholder="🔍 Search Project...",
        label_visibility="collapsed",
        key="project_search"
    )


    if project:

        result = df[
            df["Mandate"].astype(str).str.contains(
                project,
                case=False,
                na=False
            )
        ]


        if result.empty:

            st.error("❌ Project not found.")


        else:

            st.markdown(
                f"""
                <div style="
                background:#ECFDF5;
                border:1px solid #A7F3D0;
                border-left:6px solid #006747;
                padding:14px 18px;
                border-radius:12px;
                margin-bottom:15px;
                color:#006747;
                font-weight:700;
                font-size:15px;">
                ✅ {len(result)} Project(s) Found
                </div>
                """,
                unsafe_allow_html=True
            )


            # =================================================
            # VIP SEARCH RESULT TABLE
            # =================================================

            st.markdown("""
            <div style="
            background:#FFFFFF;
            border-radius:18px;
            padding:6px;
            border:1px solid #DDE5E1;
            box-shadow:0 8px 25px rgba(0,103,71,0.08);
            margin-bottom:20px;">
            """, unsafe_allow_html=True)


            st.dataframe(
                result,
                width="stretch",
                hide_index=True,
                height=400
            )


            st.markdown(
                "</div>",
                unsafe_allow_html=True
            )


            # =================================================
            # SEARCH SUMMARY
            # =================================================

            total_found = len(result)

            live_found = len(
                result[
                    result["Status"]
                    .astype(str)
                    .str.upper()
                    == "LIVE"
                ]
            )

            uat_found = len(
                result[
                    result["Status"]
                    .astype(str)
                    .str.upper()
                    == "UAT"
                ]
            )

            c1, c2, c3 = st.columns(3)


            with c1:
                st.metric(
                    "📋 Projects Found",
                    total_found
                )


            with c2:
                st.metric(
                    "🟢 Live",
                    live_found
                )


            with c3:
                st.metric(
                    "🟡 UAT",
                    uat_found
                )


    st.markdown("<br>", unsafe_allow_html=True)
# =====================================================
# PROJECTS
# =====================================================

elif page == "Projects":

    project_file = os.path.join(
        os.path.dirname(__file__),
        "data",
        "projects.xlsx"
    )

    # =====================================================
    # COMPACT HEADER
    # =====================================================

    st.markdown("""
    <div style="
    background:linear-gradient(135deg,#ffffff,#f8fbff);
    border-radius:18px;
    padding:14px 22px;
    border:1px solid #E5E7EB;
    box-shadow:0 8px 22px rgba(0,0,0,.08);
    box-sizing:border-box;
    margin-bottom:4px;">

    <div style="
    color:#006747;
    font-size:12px;
    font-weight:700;
    letter-spacing:1.8px;
    margin-bottom:4px;">
    SMARTPAY PROJECT MANAGEMENT
    </div>

    <div style="
    color:#006747;
    font-size:28px;
    font-weight:800;
    line-height:1.15;
    margin:0;">
    📁 Project Portfolio
    </div>

    <div style="
    margin-top:4px;
    color:#6B7280;
    font-size:14px;">
    SmartPay Project Management System
    </div>

    </div>
    """, unsafe_allow_html=True)


    # =====================================================
    # FILTERS - COMPACT
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:24px;
    font-weight:700;
    margin-top:2px;
    margin-bottom:6px;
    padding-top:0;">
    🎯 Filters
    </h2>
    """, unsafe_allow_html=True)


    c1, c2, c3, c4 = st.columns(4)


    with c1:

        search = st.text_input(
            "Search Project",
            placeholder="🔍 Search Project...",
            key="project_filter_search"
        )


    with c2:

        allocation = st.selectbox(
            "Team Member",
            ["All"] + sorted(
                df["Allocation"]
                .dropna()
                .unique()
            ),
            key="project_filter_allocation"
        )


    with c3:

        status = st.selectbox(
            "Status",
            ["All"] + sorted(
                df["Status"]
                .dropna()
                .unique()
            ),
            key="project_filter_status"
        )


    with c4:

        category = st.selectbox(
            "Category",
            ["All"] + sorted(
                df["Category"]
                .dropna()
                .unique()
            ),
            key="project_filter_category"
        )


    # =====================================================
    # FILTERING
    # =====================================================

    filtered_df = df.copy()


    if search:

        filtered_df = filtered_df[
            filtered_df["Mandate"]
            .astype(str)
            .str.contains(
                search,
                case=False,
                na=False
            )
        ]


    if allocation != "All":

        filtered_df = filtered_df[
            filtered_df["Allocation"] == allocation
        ]


    if status != "All":

        filtered_df = filtered_df[
            filtered_df["Status"] == status
        ]


    if category != "All":

        filtered_df = filtered_df[
            filtered_df["Category"] == category
        ]


    # =====================================================
    # COMPACT SUMMARY
    # =====================================================

    left, right = st.columns([3, 1])


    with left:

        st.markdown(f"""
        <div style="
        background:#E8F5E9;
        padding:9px 14px;
        border-radius:11px;
        color:#006747;
        font-size:15px;
        font-weight:700;
        margin-top:6px;">
        📌 Showing <b>{len(filtered_df)}</b> Project(s)
        </div>
        """, unsafe_allow_html=True)


    with right:

        csv = filtered_df.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "⬇ Download CSV",
            csv,
            "Projects.csv",
            "text/csv",
            width="stretch"
        )
    # =====================================================
    # VIP CLICKABLE KPI CARDS - COMPACT
    # =====================================================

    # -----------------------------------------
    # STATUS CLEAN
    # -----------------------------------------

    status_clean = (
        filtered_df["Status"]
        .astype(str)
        .str.strip()
        .str.upper()
    )


    # -----------------------------------------
    # COUNTS
    # -----------------------------------------

    all_count = len(filtered_df)

    scoping_count = status_clean.isin(
        ["SCOPING", "UNDER SCOPING"]
    ).sum()

    uat_count = (
        status_clean == "UAT"
    ).sum()

    review_count = (
        status_clean == "IS REVIEW"
    ).sum()

    cmc_count = (
        status_clean == "CMC"
    ).sum()

    live_count = (
        status_clean == "LIVE"
    ).sum()

    bau_count = (
        status_clean == "BAU"
    ).sum()


    # =====================================================
    # COMPACT KPI CARD CSS
    # =====================================================

    st.markdown("""
    <style>

    .st-key-status_kpis div[data-testid="stButton"] {
        width:100% !important;
    }

    .st-key-status_kpis
    div[data-testid="stButton"] button {

        width:100% !important;

        min-height:96px !important;
        height:96px !important;

        background:#FFFFFF !important;

        border:1px solid #E5E7EB !important;
        border-radius:16px !important;

        padding:10px 6px !important;

        box-shadow:
            0 6px 16px rgba(0,0,0,.07) !important;

        color:#111827 !important;

        font-size:13px !important;
        font-weight:650 !important;

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        line-height:1.3 !important;

        white-space:pre-line !important;

        text-align:center !important;

        display:flex !important;
        align-items:center !important;
        justify-content:center !important;

        transition:all .2s ease !important;
    }


    .st-key-status_kpis
    div[data-testid="stButton"] button p {

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        font-size:13px !important;

        font-weight:650 !important;

        line-height:1.3 !important;

        white-space:pre-line !important;

        text-align:center !important;

        display:block !important;

        width:100% !important;

        margin:0 !important;
        padding:0 !important;

        color:#111827 !important;
    }


    .st-key-status_kpis
    div[data-testid="stButton"] button:hover {

        background:#F8FAFC !important;

        transform:translateY(-2px) !important;

        box-shadow:
            0 10px 22px rgba(0,0,0,.11) !important;
    }


    .st-key-status_kpis
    div[data-testid="stButton"] button:focus {

        outline:none !important;

        min-height:96px !important;
        height:96px !important;

        background:#EAF5F0 !important;

        border:2px solid #006747 !important;

        box-shadow:
            0 0 0 3px rgba(0,103,71,0.12),
            0 10px 22px rgba(0,103,71,0.16) !important;

        color:#006747 !important;

        transform:translateY(-2px) !important;
    }


    /* ALL */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(1)
    div[data-testid="stButton"] button {
        border-top:6px solid #006747 !important;
    }


    /* SCOPING */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(2)
    div[data-testid="stButton"] button {
        border-top:6px solid #8E24AA !important;
    }


    /* UAT */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(3)
    div[data-testid="stButton"] button {
        border-top:6px solid #F9A825 !important;
    }


    /* IS REVIEW */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(4)
    div[data-testid="stButton"] button {
        border-top:6px solid #00ACC1 !important;
    }


    /* CMC */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(5)
    div[data-testid="stButton"] button {
        border-top:6px solid #3949AB !important;
    }


    /* LIVE */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(6)
    div[data-testid="stButton"] button {
        border-top:6px solid #00C853 !important;
    }


    /* BAU */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(7)
    div[data-testid="stButton"] button {
        border-top:6px solid #607D8B !important;
    }


    /* REMOVE EXTRA GAPS */
    .st-key-status_kpis
    div[data-testid="stVerticalBlock"] {
        gap:0 !important;
    }


    /* COLUMN SPACING */
    .st-key-status_kpis
    div[data-testid="stHorizontalBlock"] {
        gap:8px !important;
        align-items:stretch !important;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # KPI CARD CONTAINER
    # =====================================================

    with st.container(key="status_kpis"):

        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)


        # -----------------------------------------
        # ALL
        # -----------------------------------------

        with c1:

            if st.button(
                f"ALL\n{all_count}",
                key="kpi_all",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "ALL"
                st.rerun()


        # -----------------------------------------
        # SCOPING
        # -----------------------------------------

        with c2:

            if st.button(
                f"SCOPING\n{scoping_count}",
                key="kpi_scoping",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "SCOPING"
                st.rerun()


        # -----------------------------------------
        # UAT
        # -----------------------------------------

        with c3:

            if st.button(
                f"UAT\n{uat_count}",
                key="kpi_uat",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "UAT"
                st.rerun()


        # -----------------------------------------
        # IS REVIEW
        # -----------------------------------------

        with c4:

            if st.button(
                f"IS REVIEW\n{review_count}",
                key="kpi_review",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "IS REVIEW"
                st.rerun()


        # -----------------------------------------
        # CMC
        # -----------------------------------------

        with c5:

            if st.button(
                f"CMC\n{cmc_count}",
                key="kpi_cmc",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "CMC"
                st.rerun()


        # -----------------------------------------
        # LIVE
        # -----------------------------------------

        with c6:

            if st.button(
                f"LIVE\n{live_count}",
                key="kpi_live",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "LIVE"
                st.rerun()


        # -----------------------------------------
        # BAU
        # -----------------------------------------

        with c7:

            if st.button(
                f"BAU\n{bau_count}",
                key="kpi_bau",
                use_container_width=True
            ):

                st.session_state["project_status_filter"] = "BAU"
                st.rerun()


    st.markdown("<br>", unsafe_allow_html=True)
    # =====================================================
    # APPLY KPI STATUS FILTER TO EXISTING TABLE
    # =====================================================

    selected_status = st.session_state.get(
        "project_status_filter",
        "ALL"
    )

    status_upper = (
        filtered_df["Status"]
        .astype(str)
        .str.strip()
        .str.upper()
    )


    if selected_status == "SCOPING":

        filtered_df = filtered_df[
            status_upper.isin([
                "SCOPING",
                "UNDER SCOPING"
            ])
        ].copy()


    elif selected_status == "DEVELOPMENT":

        filtered_df = filtered_df[
            status_upper.isin([
                "DEVELOPMENT",
                "UNDER DEVELOPMENT",
                "SIT"
            ])
        ].copy()


    elif selected_status == "UAT":

        filtered_df = filtered_df[
            status_upper == "UAT"
        ].copy()


    elif selected_status == "IS REVIEW":

        filtered_df = filtered_df[
            status_upper == "IS REVIEW"
        ].copy()


    elif selected_status == "CMC":

        filtered_df = filtered_df[
            status_upper == "CMC"
        ].copy()


    elif selected_status == "LIVE":

        filtered_df = filtered_df[
            status_upper == "LIVE"
        ].copy()


    elif selected_status == "BAU":

        filtered_df = filtered_df[
            status_upper == "BAU"
        ].copy()
    # =====================================================
    # PROJECT TABLE - VIP STYLE
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:30px;
    font-weight:700;
    margin-bottom:18px;">
    📋 Project Details
    </h2>
    """, unsafe_allow_html=True)


    # -----------------------------
    # VIP TABLE CSS
    # -----------------------------

    st.markdown("""
    <style>

    .project-table-wrapper {
        background: #FFFFFF;
        border: 1px solid #DDE5E1;
        border-radius: 16px;
        padding: 6px;
        box-shadow: 0 6px 20px rgba(0,103,71,0.08);
        overflow: hidden;
    }

    .project-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 14px;
    }

    .project-table thead th {
        background: #006747;
        color: white;
        font-weight: 700;
        padding: 14px 12px;
        text-align: left;
        border: none;
    }

    .project-table thead th:first-child {
        border-top-left-radius: 11px;
    }

    .project-table thead th:last-child {
        border-top-right-radius: 11px;
    }

    .project-table tbody td {
        padding: 13px 12px;
        color: #1F2937;
        border-bottom: 1px solid #E5E7EB;
        background: #FFFFFF;
    }

    .project-table tbody tr:nth-child(even) td {
        background: #F8FAFC;
    }

    .project-table tbody tr:hover td {
        background: #ECFDF5;
    }

    .project-name {
        font-weight: 700;
        color: #006747 !important;
    }

    .status-badge {
        display: inline-block;
        padding: 5px 11px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 800;
        white-space: nowrap;
    }

    .status-live {
        background: #DCFCE7;
        color: #166534;
    }

    .status-uat {
        background: #FEF3C7;
        color: #92400E;
    }

    .status-development {
        background: #DBEAFE;
        color: #1E40AF;
    }

    .status-review {
        background: #CCFBF1;
        color: #115E59;
    }

    .status-cmc {
        background: #EDE9FE;
        color: #5B21B6;
    }

    .status-scoping {
        background: #F3E8FF;
        color: #7E22CE;
    }

    .status-default {
        background: #F3F4F6;
        color: #374151;
    }

    </style>
    """, unsafe_allow_html=True)


    # -----------------------------
    # CREATE VIP TABLE
    # -----------------------------

    display_df = filtered_df.copy()


    # ==========================================
    # DATE FORMAT
    # ==========================================

    if "Date" in display_df.columns:

        display_df["Date"] = pd.to_datetime(
            display_df["Date"],
            errors="coerce"
        ).dt.strftime("%Y-%m-%d")


    # ==========================================
    # LIVE DATE FORMAT
    # ==========================================

    if "Live Date" in display_df.columns:

        def format_live_date(value):

            if pd.isna(value):
                return "TBD"

            value = str(value).strip()

            if value.upper() == "BAU":
                return "BAU"

            parsed = pd.to_datetime(
                value,
                errors="coerce"
            )

            if pd.notna(parsed):
                return parsed.strftime("%Y-%m-%d")

            return value


        display_df["Live Date"] = (
            display_df["Live Date"]
            .apply(format_live_date)
        )

    def status_badge(status):

        status = str(status).strip()
        status_upper = status.upper()

        if status_upper == "LIVE":
            css = "status-live"

        elif status_upper == "UAT":
            css = "status-uat"

        elif status_upper in [
            "DEVELOPMENT",
            "UNDER DEVELOPMENT",
            "SIT"
        ]:
            css = "status-development"

        elif status_upper == "IS REVIEW":
            css = "status-review"

        elif status_upper == "CMC":
            css = "status-cmc"

        elif status_upper in [
            "SCOPING",
            "UNDER SCOPING"
        ]:
            css = "status-scoping"

        else:
            css = "status-default"

        return f'<span class="status-badge {css}">{status}</span>'


    # Status ko badge mein convert karo
    if "Status" in display_df.columns:
        display_df["Status"] = display_df["Status"].apply(status_badge)


    # Project/Mandate name ko highlight karo
    if "Mandate" in display_df.columns:
        display_df["Mandate"] = display_df["Mandate"].apply(
            lambda x: f'<span class="project-name">📁 {x}</span>'
        )


    # -----------------------------
    # HTML TABLE
    # -----------------------------

    table_html = display_df.to_html(
        index=False,
        escape=False,
        classes="project-table"
    )


    st.markdown(
        f"""
        <div class="project-table-wrapper">
            {table_html}
        </div>
        """,
        unsafe_allow_html=True
    )
    # =====================================================
    # EDIT PROJECTS
    # =====================================================

    st.markdown("<br>", unsafe_allow_html=True)

    if "project_edit_mode" not in st.session_state:
        st.session_state["project_edit_mode"] = False


    # =====================================================
    # EDIT BUTTON
    # =====================================================

    if not st.session_state["project_edit_mode"]:

        if st.button(
            "✏️ Edit Projects",
            key="project_edit_button",
            use_container_width=True
        ):

            st.session_state["project_edit_mode"] = True
            st.rerun()


    # =====================================================
    # EDITOR
    # =====================================================

    if st.session_state["project_edit_mode"]:

        st.markdown("""
        <h2 style="
        color:#006747;
        font-size:28px;
        font-weight:700;
        margin-top:20px;
        margin-bottom:15px;">
        ✏️ Edit Projects
        </h2>
        """, unsafe_allow_html=True)

        st.info(
            "✏️ Edit project data or add new projects."
        )


        # =================================================
        # EDITABLE PROJECT TABLE
        # =================================================

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=550,
            key="project_editor"
        )


        # =================================================
        # SAVE / CLOSE
        # =================================================

        save_col, close_col = st.columns(2)


        # =================================================
        # SAVE
        # =================================================

        with save_col:

            if st.button(
                "💾 Save Changes",
                type="primary",
                key="project_save",
                use_container_width=True
            ):

                try:

                    edited_df = edited_df.fillna("")

                    edited_df.to_excel(
                        project_file,
                        index=False
                    )

                    st.success(
                        "✅ Project data saved successfully!"
                    )

                    st.session_state[
                        "project_edit_mode"
                    ] = False

                    if "project_editor" in st.session_state:
                        del st.session_state["project_editor"]

                    st.rerun()

                except Exception as e:

                    st.error(
                        f"Save error: {e}"
                    )


        # =================================================
        # CLOSE
        # =================================================

        with close_col:

            if st.button(
                "❌ Close Editor",
                key="project_close_editor",
                use_container_width=True
            ):

                st.session_state[
                    "project_edit_mode"
                ] = False

                if "project_editor" in st.session_state:
                    del st.session_state["project_editor"]

                st.rerun()
# =====================================================
# ANALYTICS
# =====================================================

elif page == "Analytics":

    # ==========================================
    # ANALYTICS DATA
    # ==========================================

    analytics_source_df = df[
        df["Status"]
        .astype(str)
        .str.strip()
        .str.upper()
        != "BAU"
    ].copy()

    bau_df = df[
        df["Status"]
        .astype(str)
        .str.strip()
        .str.upper()
        == "BAU"
    ].copy()


    # ==========================================
    # COMPACT HEADER
    # ==========================================

    st.markdown("""
    <div style="
    background:linear-gradient(135deg,#ffffff,#f8fbff);
    border-radius:18px;
    padding:14px 22px;
    border:1px solid #E5E7EB;
    box-shadow:0 8px 22px rgba(0,0,0,.08);
    box-sizing:border-box;
    margin-bottom:6px;">

    <div style="
    color:#006747;
    font-size:12px;
    font-weight:700;
    letter-spacing:1.8px;
    margin-bottom:4px;">
    SMARTPAY ANALYTICS
    </div>

    <div style="
    color:#006747;
    font-size:28px;
    font-weight:800;
    line-height:1.15;
    margin:0;">
    📊 Analytics Dashboard
    </div>

    <div style="
    margin-top:4px;
    color:#6B7280;
    font-size:14px;">
    SmartPay Project Insights & Team Performance
    </div>

    </div>
    """, unsafe_allow_html=True)


    # ==========================================
    # KPI CARDS
    # ==========================================

    status = (
        df["Status"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    total_projects = len(df)

    live_projects = len(
        df[status == "LIVE"]
    )

    uat_projects = len(
        df[status == "UAT"]
    )

    team_members = (
        df["Allocation"]
        .nunique()
    )

    live_percent = (
        round(
            (live_projects / total_projects) * 100,
            1
        )
        if total_projects
        else 0
    )

    uat_percent = (
        round(
            (uat_projects / total_projects) * 100,
            1
        )
        if total_projects
        else 0
    )


    # ==========================================
    # COMPACT ANALYTICS KPI CARD
    # ==========================================

    def analytics_card(title, value, color):

        st.markdown(f"""
        <div style="
        background:#FFFFFF;
        border-radius:14px;
        padding:10px 8px;
        text-align:center;
        border-top:5px solid {color};
        border-left:1px solid #E5E7EB;
        border-right:1px solid #E5E7EB;
        border-bottom:1px solid #E5E7EB;
        box-shadow:0 5px 14px rgba(0,0,0,.06);
        min-height:90px;
        box-sizing:border-box;">

        <div style="
        color:#6B7280;
        font-size:12px;
        font-weight:600;
        line-height:1.2;">
        {title}
        </div>

        <div style="
        color:{color};
        font-size:30px;
        font-weight:800;
        line-height:1;
        margin-top:7px;">
        {value}
        </div>

        </div>
        """, unsafe_allow_html=True)


    k1, k2, k3, k4 = st.columns(4)


    with k1:

        analytics_card(
            "Total Projects",
            total_projects,
            "#006747"
        )


    with k2:

        analytics_card(
            "Live %",
            f"{live_percent}%",
            "#00C853"
        )


    with k3:

        analytics_card(
            "UAT %",
            f"{uat_percent}%",
            "#F9A825"
        )


    with k4:

        analytics_card(
            "Team Members",
            team_members,
            "#3949AB"
        )


    # ==========================================
    # SMALL GAP BEFORE FILTERS
    # ==========================================

    st.markdown(
        "<div style='height:6px;'></div>",
        unsafe_allow_html=True
    )


    # ==========================================
    # FILTERS
    # ==========================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:24px;
    font-weight:700;
    margin-top:2px;
    margin-bottom:7px;">
    🎯 Analytics Filters
    </h2>
    """, unsafe_allow_html=True)


    f1, f2 = st.columns([2, 2])


    with f1:

        selected_allocation = st.selectbox(
            "👤 Team Member",
            ["All"] + sorted(
                analytics_source_df[
                    "Allocation"
                ]
                .dropna()
                .unique()
            )
        )


    with f2:

        selected_status = st.selectbox(
            "📌 Status",
            ["All"] + sorted(
                analytics_source_df[
                    "Status"
                ]
                .dropna()
                .unique()
            )
        )


    # ==========================================
    # ANALYTICS FILTER DATA
    # ==========================================

    analytics_df = analytics_source_df.copy()


    if selected_allocation != "All":

        analytics_df = analytics_df[
            analytics_df["Allocation"]
            == selected_allocation
        ]


    if selected_status != "All":

        analytics_df = analytics_df[
            analytics_df["Status"]
            == selected_status
        ]


    analytics_df["Status"] = (
        analytics_df["Status"]
        .astype(str)
        .str.upper()
        .str.strip()
    )


    # ==========================================
    # STATUS ORDER
    # ==========================================

    status_order = [
        "LIVE",
        "UAT",
        "IS REVIEW",
        "CMC",
        "UNDER SCOPING"
    ]


    color_map = {
        "LIVE": "#00C853",
        "UAT": "#F9A825",
        "IS REVIEW": "#00ACC1",
        "CMC": "#3949AB",
        "UNDER SCOPING": "#8E24AA"
    }

    # ==========================================
    # STATUS CHART - COMPACT
    # ==========================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:24px;
    font-weight:700;
    margin-top:6px;
    margin-bottom:8px;">
    📈 Project Status
    </h2>
    """, unsafe_allow_html=True)

    status_count = (
        analytics_df["Status"]
        .value_counts()
        .reindex(status_order, fill_value=0)
        .reset_index()
    )

    status_count.columns = [
        "Status",
        "Projects"
    ]

    fig1 = px.bar(
        status_count,
        x="Status",
        y="Projects",
        text="Projects",
        color="Status",
        color_discrete_map=color_map
    )

    fig1.update_layout(
        height=380,
        template="plotly_white",
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_text="",
        coloraxis_showscale=False,

        font=dict(
            color="#111827",
            size=12
        ),

        xaxis=dict(
            title="Status",
            title_font=dict(
                color="#374151",
                size=12
            ),
            tickfont=dict(
                color="#374151",
                size=11
            ),
            showgrid=False,
            zeroline=False
        ),

        yaxis=dict(
            title="Projects",
            title_font=dict(
                color="#374151",
                size=12
            ),
            tickfont=dict(
                color="#374151",
                size=11
            ),
            gridcolor="#E5E7EB",
            zeroline=False
        ),

        legend=dict(
            font=dict(
                color="#374151",
                size=11
            )
        ),

        margin=dict(
            l=35,
            r=20,
            t=10,
            b=45
        )
    )

    fig1.update_traces(
        textposition="inside",
        textfont=dict(size=11),
        marker_line_width=0
    )

    st.plotly_chart(
        fig1,
        width="stretch"
    )


    # ==========================================
    # TEAM PERFORMANCE - COMPACT
    # ==========================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:24px;
    font-weight:700;
    margin-top:6px;
    margin-bottom:8px;">
    👥 Team Performance
    </h2>
    """, unsafe_allow_html=True)

    allocation_count = (
        analytics_df["Allocation"]
        .value_counts()
        .reset_index()
    )

    allocation_count.columns = [
        "Allocation",
        "Projects"
    ]

    fig2 = px.bar(
        allocation_count,
        x="Allocation",
        y="Projects",
        text="Projects",
        color="Projects",
        color_continuous_scale="Greens"
    )

    fig2.update_layout(
        height=380,
        template="plotly_white",
        plot_bgcolor="white",
        paper_bgcolor="white",
        coloraxis_showscale=False,

        font=dict(
            color="#111827",
            size=12
        ),

        xaxis=dict(
            title="",
            tickfont=dict(
                color="#374151",
                size=11
            ),
            showgrid=False,
            zeroline=False
        ),

        yaxis=dict(
            title="Projects",
            title_font=dict(
                color="#374151",
                size=12
            ),
            tickfont=dict(
                color="#374151",
                size=11
            ),
            gridcolor="#E5E7EB",
            zeroline=False
        ),

        margin=dict(
            l=35,
            r=20,
            t=10,
            b=55
        )
    )

    fig2.update_traces(
        textposition="outside",
        textfont=dict(
            color="#111827",
            size=11
        )
    )

    st.plotly_chart(
        fig2,
        width="stretch"
    )


    # ==========================================
    # PERSON ANALYTICS - COMPACT
    # ==========================================

    if selected_allocation != "All":

        # ==========================================
        # PERSON DATA
        # ==========================================

        person_df = analytics_df[
            analytics_df["Allocation"] == selected_allocation
        ].copy()


        # ==========================================
        # PIE CHART
        # ==========================================

        st.markdown(
            f"""
            <h2 style="
            color:#006747;
            font-size:24px;
            font-weight:700;
            margin-top:6px;
            margin-bottom:5px;">
            🥧 {selected_allocation}
            </h2>
            """,
            unsafe_allow_html=True
        )

        person_status = (
            person_df["Status"]
            .value_counts()
            .reindex(status_order, fill_value=0)
            .reset_index()
        )

        person_status.columns = [
            "Status",
            "Projects"
        ]

        fig3 = px.pie(
            person_status,
            names="Status",
            values="Projects",
            hole=.55,
            color="Status",
            color_discrete_map=color_map
        )

        fig3.update_layout(
            height=360,
            template="plotly_white",
            plot_bgcolor="white",
            paper_bgcolor="white",

            font=dict(
                color="#111827",
                size=12
            ),

            legend=dict(
                font=dict(
                    color="#374151",
                    size=11
                )
            ),

            margin=dict(
                l=10,
                r=10,
                t=5,
                b=5
            )
        )

        fig3.update_traces(
            textfont=dict(
                color="white",
                size=11
            )
        )

        st.plotly_chart(
            fig3,
            width="stretch"
        )

        # ==========================================
        # PROJECT STATUS PROGRESS - BELOW PIE
        # ==========================================

        st.markdown(
            """
            <h2 style="
            color:#006747;
            font-size:24px;
            font-weight:700;
            margin-top:8px;
            margin-bottom:12px;">
            📊 Project Status Progress
            </h2>
            """,
            unsafe_allow_html=True
        )


        status_stages = [
            "SCOPING",
            "UAT",
            "IS REVIEW",
            "CMC",
            "LIVE"
        ]


        status_progress = {
            "SCOPING": 1,
            "UNDER SCOPING": 1,
            "UAT": 2,
            "IS REVIEW": 3,
            "CMC": 4,
            "LIVE": 5
        }


        # ==========================================
        # REMOVE BAU PROJECTS
        # ==========================================

        project_df = person_df[
            person_df["Status"]
            .astype(str)
            .str.strip()
            .str.upper()
            != "BAU"
        ].copy()


        # ==========================================
        # INDIVIDUAL PROJECT CARDS
        # ==========================================

        for _, row in project_df.iterrows():

            project_name = str(
                row["Mandate"]
            ).strip()

            current_status = str(
                row["Status"]
            ).strip().upper()

            current_stage = status_progress.get(
                current_status,
                1
            )

            progress_percent = int(
                (current_stage / 5) * 100
            )


            # ======================================
            # STATUS COLORS
            # ======================================

            if current_status == "LIVE":

                status_bg = "#DCFCE7"
                status_color = "#166534"

            elif current_status == "UAT":

                status_bg = "#FEF3C7"
                status_color = "#92400E"

            elif current_status == "CMC":

                status_bg = "#EDE9FE"
                status_color = "#5B21B6"

            elif current_status == "IS REVIEW":

                status_bg = "#CCFBF1"
                status_color = "#115E59"

            elif current_status in [
                "SCOPING",
                "UNDER SCOPING"
            ]:

                status_bg = "#F3E8FF"
                status_color = "#7E22CE"

            else:

                status_bg = "#F3F4F6"
                status_color = "#374151"


            # ======================================
            # PROJECT CARD
            # ======================================

            card_html = f"""
            <div style="
                background:#FFFFFF;
                border:1px solid #DDE5E1;
                border-radius:14px;
                padding:12px 14px;
                margin-bottom:10px;
                box-shadow:0 4px 12px rgba(0,103,71,.06);
                box-sizing:border-box;
            ">

                <!-- PROJECT HEADER -->
                <div style="
                    display:flex;
                    justify-content:space-between;
                    align-items:center;
                    gap:12px;
                    margin-bottom:9px;
                ">

                    <div style="
                        color:#006747;
                        font-size:14px;
                        font-weight:750;
                        line-height:1.3;
                        flex:1;
                        word-break:break-word;
                    ">
                        📁 {project_name}
                    </div>

                    <div style="
                        background:{status_bg};
                        color:{status_color};
                        padding:4px 9px;
                        border-radius:14px;
                        font-size:9px;
                        font-weight:800;
                        white-space:nowrap;
                        flex-shrink:0;
                    ">
                        {current_status}
                    </div>

                </div>


                <!-- PROGRESS LINE -->

                <div style="
                    width:100%;
                    height:6px;
                    background:#E5E7EB;
                    border-radius:10px;
                    overflow:hidden;
                ">

                    <div style="
                        width:{progress_percent}%;
                        height:100%;
                        background:#006747;
                        border-radius:10px;
                    ">
                    </div>

                </div>


                <!-- STATUS STEPS -->

                <div style="
                    display:flex;
                    justify-content:space-between;
                    align-items:flex-start;
                    margin-top:9px;
                ">
            """


            for i, stage in enumerate(
                status_stages,
                start=1
            ):

                if i <= current_stage:

                    dot_color = "#006747"
                    text_color = "#006747"

                else:

                    dot_color = "#D1D5DB"
                    text_color = "#9CA3AF"


                card_html += f"""
                    <div style="
                        flex:1;
                        text-align:center;
                    ">

                        <div style="
                            width:8px;
                            height:8px;
                            background:{dot_color};
                            border-radius:50%;
                            margin:auto;
                        ">
                        </div>

                        <div style="
                            margin-top:4px;
                            color:{text_color};
                            font-size:8px;
                            font-weight:700;
                            white-space:nowrap;
                        ">
                            {stage}
                        </div>

                    </div>
                """


            card_html += """
                </div>

            </div>
            """


            st.html(card_html)


        # ==========================================
        # PROJECT DETAILS
        # ==========================================

        display_person_df = person_df.copy()


        # ==========================================
        # DATE FORMAT
        # ==========================================

        for col in ["Date", "Live Date"]:

            if col in display_person_df.columns:

                original = display_person_df[col].copy()

                parsed = pd.to_datetime(
                    original,
                    errors="coerce"
                )

                formatted = parsed.dt.strftime(
                    "%Y-%m-%d"
                )

                # Keep text values like BAU
                display_person_df[col] = formatted.where(
                    parsed.notna(),
                    original.astype(str)
                )

                # Blank values
                display_person_df[col] = (
                    display_person_df[col]
                    .replace(
                        ["nan", "NaT", "", "None"],
                        "TBD"
                    )
                )


        st.dataframe(
            display_person_df,
            width="stretch",
            hide_index=True
        )


# =====================================================
        # =====================================================
        # BAU MONITORING ANALYTICS - COMPACT
        # =====================================================

        st.markdown(
            "<div style='height:6px;'></div>",
            unsafe_allow_html=True
        )


        st.markdown("""
        <h2 style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:8px;
        margin-bottom:5px;">
        🏦 BAU Monitoring
        </h2>

        <p style="
        color:#6B7280;
        font-size:13px;
        margin-top:0;
        margin-bottom:12px;">
        Business as Usual projects are monitored separately from delivery projects.
        </p>
        """, unsafe_allow_html=True)


        # =====================================================
        # BAU DATA
        # =====================================================

        bau_df = df[
            df["Status"]
            .astype(str)
            .str.strip()
            .str.upper()
            == "BAU"
        ].copy()


        # =====================================================
        # BAU KPI
        # =====================================================

        bau_total = len(bau_df)

        bau_owners = (
            bau_df["Allocation"]
            .dropna()
            .astype(str)
            .nunique()
        )

        bau_updates = (
            bau_df["Update"]
            .notna()
            .sum()
            if "Update" in bau_df.columns
            else 0
        )


        b1, b2, b3 = st.columns(3)


        with b1:

            analytics_card(
                "BAU Projects",
                bau_total,
                "#607D8B"
            )


        with b2:

            analytics_card(
                "BAU Owners",
                bau_owners,
                "#3949AB"
            )


        with b3:

            analytics_card(
                "BAU Updates",
                bau_updates,
                "#00897B"
            )


        # =====================================================
        # BAU PROJECTS BY OWNER
        # =====================================================

        if not bau_df.empty:

            st.markdown("""
            <h2 style="
            color:#006747;
            font-size:22px;
            font-weight:700;
            margin-top:8px;
            margin-bottom:8px;">
            👥 BAU Projects by Owner
            </h2>
            """, unsafe_allow_html=True)


            bau_owner_count = (
                bau_df["Allocation"]
                .astype(str)
                .value_counts()
                .reset_index()
            )

            bau_owner_count.columns = [
                "Allocation",
                "Projects"
            ]


            fig_bau = px.bar(
                bau_owner_count,
                x="Allocation",
                y="Projects",
                text="Projects",
                color="Projects",
                color_continuous_scale="Greens"
            )


            fig_bau.update_traces(
                textposition="outside",
                marker_line_width=0,
                textfont=dict(size=11)
            )


            fig_bau.update_layout(
                height=370,
                template="plotly_white",
                plot_bgcolor="white",
                paper_bgcolor="white",
                coloraxis_showscale=False,

                font=dict(
                    color="#111827",
                    size=11
                ),

                xaxis=dict(
                    title="",
                    showgrid=False,
                    tickfont=dict(
                        color="#374151",
                        size=11
                    )
                ),

                yaxis=dict(
                    title="BAU Projects",
                    gridcolor="#E5E7EB",
                    tickfont=dict(
                        color="#374151",
                        size=10
                    )
                ),

                margin=dict(
                    l=35,
                    r=20,
                    t=8,
                    b=45
                )
            )


            st.plotly_chart(
                fig_bau,
                width="stretch"
            )
## =====================================================
# PROJECT TIMELINE
# =====================================================

elif page == "Project Timeline":

    # ==========================================
    # TIMELINE DATA
    # ==========================================

    timeline_df = df[
        df["Status"]
        .astype(str)
        .str.strip()
        .str.upper()
        != "BAU"
    ].copy()

    # ==========================================
    # HEADER
    # ==========================================

    st.html("""
    <div style="
        background:linear-gradient(135deg,#ffffff,#f8fbff);
        border-radius:18px;
        padding:14px 22px;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 22px rgba(0,0,0,.08);
        margin-bottom:8px;
        font-family:Segoe UI,Arial,sans-serif;
    ">
        <div style="
            color:#006747;
            font-size:12px;
            font-weight:700;
            letter-spacing:1.8px;
            margin-bottom:4px;
        ">
            SMARTPAY PROJECT MANAGEMENT
        </div>

        <div style="
            color:#006747;
            font-size:28px;
            font-weight:800;
            line-height:1.15;
        ">
            📅 Project Timeline
        </div>

        <div style="
            color:#6B7280;
            font-size:14px;
            margin-top:4px;
        ">
            Track the current progress of SmartPay projects across every delivery stage.
        </div>
    </div>
    """)

    # ==========================================
    # LEGEND
    # ==========================================

    st.html("""
    <div style="
        background:#FFFFFF;
        border-radius:14px;
        padding:9px 14px;
        border:1px solid #E5E7EB;
        box-shadow:0 4px 12px rgba(0,0,0,.05);
        margin-bottom:10px;
        font-family:Segoe UI,Arial,sans-serif;
    ">
        <div style="
            color:#006747;
            font-size:13px;
            font-weight:700;
            margin-bottom:7px;
        ">
            Timeline Status
        </div>

        <div style="
            display:flex;
            align-items:center;
            gap:22px;
            flex-wrap:wrap;
            font-size:11px;
            font-weight:600;
        ">
            <div>
                <span style="
                    display:inline-block;
                    width:9px;
                    height:9px;
                    background:#16A34A;
                    border-radius:50%;
                    margin-right:5px;
                "></span>
                <span style="color:#374151;">Completed</span>
            </div>

            <div>
                <span style="
                    display:inline-block;
                    width:9px;
                    height:9px;
                    background:#F59E0B;
                    border-radius:50%;
                    margin-right:5px;
                "></span>
                <span style="color:#374151;">Current Stage</span>
            </div>

            <div>
                <span style="
                    display:inline-block;
                    width:9px;
                    height:9px;
                    background:#D1D5DB;
                    border-radius:50%;
                    margin-right:5px;
                "></span>
                <span style="color:#374151;">Pending</span>
            </div>
        </div>
    </div>
    """)

    # ==========================================
    # SEARCH
    # ==========================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:2px;
        margin-bottom:4px;
        font-family:Segoe UI,Arial,sans-serif;">
        🔎 Find Project Timeline
    </div>
    """)

    search_type = st.radio(
        "Search By",
        ["Project", "Team Member"],
        horizontal=True,
        label_visibility="visible"
    )

    st.markdown(
        "<div style='height:2px;'></div>",
        unsafe_allow_html=True
    )

    if search_type == "Project":

        selected = st.selectbox(
            "Select Project",
            ["All Projects"] + sorted(
                timeline_df["Mandate"]
                .dropna()
                .astype(str)
                .unique()
            ),
            key="timeline_project_select"
        )

    else:

        selected = st.selectbox(
            "Select Team Member",
            ["All Members"] + sorted(
                timeline_df["Allocation"]
                .dropna()
                .astype(str)
                .unique()
            ),
            key="timeline_member_select"
        )
    # ==========================================
    # STAGES
    # ==========================================

    stages = [
        "SCOPING",
        "DEVELOPMENT",
        "UAT",
        "IS REVIEW",
        "CMC",
        "LIVE"
    ]
    # ==========================================
    # TIMELINE FUNCTION - COMPACT
    # ==========================================

    def show_timeline(project):

        current = str(
            project["Status"]
        ).upper().strip()


        # ==========================================
        # STATUS MAPPING
        # ==========================================

        if current in [
            "SIT",
            "UNDER DEVELOPMENT",
            "DEVELOPMENT"
        ]:

            current = "DEVELOPMENT"

        elif current in [
            "UNDER SCOPING",
            "SCOPING"
        ]:

            current = "SCOPING"

        elif current not in stages:

            current = "SCOPING"


        current_index = stages.index(current)


        # ==========================================
        # PROJECT TITLE - COMPACT
        # ==========================================

        st.html(f"""
        <div style="
            background:linear-gradient(180deg,#ffffff,#f9fbfd);
            border-radius:15px;
            padding:11px 15px;
            margin-bottom:7px;
            border:1px solid #E5E7EB;
            box-shadow:0 4px 12px rgba(0,0,0,.06);
            font-family:Segoe UI,Arial,sans-serif;">

            <div style="
                color:#6B7280;
                font-size:10px;
                font-weight:600;
                letter-spacing:1px;">
                PROJECT TIMELINE
            </div>

            <div style="
                color:#006747;
                font-size:18px;
                font-weight:750;
                line-height:1.3;
                margin-top:3px;
                word-break:break-word;">
                📌 {project["Mandate"]}
            </div>

        </div>
        """)


        # ==========================================
        # TIMELINE - COMPACT
        # ==========================================

        timeline_html = """
        <div style="
            background:white;
            border-radius:15px;
            padding:15px 12px;
            border:1px solid #E5E7EB;
            box-shadow:0 5px 14px rgba(0,0,0,.05);
            font-family:Segoe UI,Arial,sans-serif;
            overflow-x:auto;
            margin-bottom:8px;">

            <div style="
                display:flex;
                align-items:flex-start;
                min-width:620px;">
        """


        for i, stage in enumerate(stages):

            if i < current_index:

                color = "#16A34A"
                symbol = "✓"

            elif i == current_index:

                color = "#F59E0B"
                symbol = "●"

            else:

                color = "#D1D5DB"
                symbol = "○"


            # ======================================
            # CONNECTOR
            # ======================================

            connector = ""

            if i < len(stages) - 1:

                if i < current_index:

                    line_color = "#16A34A"

                else:

                    line_color = "#D1D5DB"


                connector = f"""
                <div style="
                    flex:1;
                    height:3px;
                    background:{line_color};
                    margin-top:14px;">
                </div>
                """


            timeline_html += f"""

            <div style="
                width:78px;
                text-align:center;
                flex-shrink:0;">

                <div style="
                    width:28px;
                    height:28px;
                    border-radius:50%;
                    background:{color};
                    color:white;
                    margin:auto;
                    line-height:28px;
                    font-size:13px;
                    font-weight:700;
                    box-shadow:0 3px 8px rgba(0,0,0,.12);">
                    {symbol}
                </div>

                <div style="
                    margin-top:6px;
                    color:#374151;
                    font-size:9px;
                    font-weight:700;
                    white-space:nowrap;">
                    {stage}
                </div>

            </div>

            {connector}
            """


        timeline_html += """
            </div>
        </div>
        """


        st.html(
            timeline_html
        )


        # ==========================================
        # PROJECT INFORMATION - COMPACT
        # ==========================================

        c1, c2, c3 = st.columns(3)


        with c1:

            st.html(f"""
            <div style="
                background:white;
                border-radius:13px;
                padding:10px 12px;
                border:1px solid #E5E7EB;
                box-shadow:0 3px 10px rgba(0,0,0,.04);
                font-family:Segoe UI,Arial,sans-serif;
                min-height:66px;">

                <div style="
                    color:#6B7280;
                    font-size:10px;
                    font-weight:600;">
                    PROJECT
                </div>

                <div style="
                    color:#111827;
                    font-size:13px;
                    font-weight:700;
                    line-height:1.25;
                    margin-top:3px;">
                    {project["Mandate"]}
                </div>

            </div>
            """)


        with c2:

            st.html(f"""
            <div style="
                background:white;
                border-radius:13px;
                padding:10px 12px;
                border:1px solid #E5E7EB;
                box-shadow:0 3px 10px rgba(0,0,0,.04);
                min-height:66px;
                font-family:Segoe UI,Arial,sans-serif;">

                <div style="
                    color:#6B7280;
                    font-size:10px;
                    font-weight:600;">
                    OWNER
                </div>

                <div style="
                    color:#006747;
                    font-size:13px;
                    font-weight:700;
                    margin-top:3px;">
                    {project["Allocation"]}
                </div>

            </div>
            """)


        with c3:

            st.html(f"""
            <div style="
                background:white;
                border-radius:13px;
                padding:10px 12px;
                border:1px solid #E5E7EB;
                box-shadow:0 3px 10px rgba(0,0,0,.04);
                min-height:66px;
                font-family:Segoe UI,Arial,sans-serif;">

                <div style="
                    color:#92400E;
                    font-size:10px;
                    font-weight:600;">
                    CURRENT STAGE
                </div>

                <div style="
                    color:#F59E0B;
                    font-size:13px;
                    font-weight:700;
                    margin-top:3px;">
                    {current}
                </div>

            </div>
            """)


        st.markdown(
            "<div style='height:8px;'></div>",
            unsafe_allow_html=True
        )


    # ==========================================
    # DISPLAY PROJECTS
    # ==========================================

    if search_type == "Project":

        if selected == "All Projects":

            for _, project in timeline_df.iterrows():

                show_timeline(project)

        else:

            selected_project = timeline_df[
                timeline_df["Mandate"].astype(str)
                == selected
            ]

            if not selected_project.empty:

                show_timeline(
                    selected_project.iloc[0]
                )


    else:

        if selected == "All Members":

            for _, project in timeline_df.iterrows():

                show_timeline(project)

        else:

            member_df = timeline_df[
                timeline_df["Allocation"].astype(str)
                == selected
            ]

            st.success(
                f"{selected} is handling "
                f"{len(member_df)} project(s)."
            )

            for _, project in member_df.iterrows():

                show_timeline(project)
# # =====================================================
# # TEAM PERFORMANCE
# # =====================================================

# elif page == "Team Performance":

#     # ==========================================
#     # HEADER
#     # ==========================================

#     st.html("""
#     <div style="
#         background:linear-gradient(180deg,#ffffff,#f8fbff);
#         border-radius:24px;
#         padding:28px;
#         border:1px solid #E5E7EB;
#         box-shadow:0 14px 35px rgba(0,0,0,.10);
#         margin-bottom:20px;
#         font-family:Segoe UI,Arial,sans-serif;">

#         <div style="
#             color:#006747;
#             font-size:14px;
#             font-weight:700;
#             letter-spacing:2px;
#             margin-bottom:8px;">
#             SMARTPAY PROJECT MANAGEMENT
#         </div>

#         <div style="
#             color:#006747;
#             font-size:40px;
#             font-weight:700;">
#             Team Performance
#         </div>

#         <div style="
#             color:#6B7280;
#             font-size:17px;
#             margin-top:10px;">
#             Monitor team workload, project distribution and delivery progress.
#         </div>

#     </div>
#     """)

#     st.markdown("<br>", unsafe_allow_html=True)


#     # ==========================================
#     # TEAM MEMBER FILTER
#     # ==========================================

#     st.markdown("""
#     <h2 style="
#         color:#006747;
#         font-size:28px;
#         margin-bottom:12px;">
#         👤 Team Member
#     </h2>
#     """, unsafe_allow_html=True)

#     member = st.selectbox(
#         "Select Team Member",
#         ["All"] + sorted(
#             df["Allocation"]
#             .dropna()
#             .astype(str)
#             .unique()
#         )
#     )

#     team_df = df.copy()

#     if member != "All":
#         team_df = team_df[
#             team_df["Allocation"].astype(str) == member
#         ]


#     # ==========================================
#     # CLEAN STATUS
#     # ==========================================

#     team_df["Status_Clean"] = (
#         team_df["Status"]
#         .astype(str)
#         .str.upper()
#         .str.strip()
#         .replace({
#             "SIT": "UNDER DEVELOPMENT",
#             "DEVELOPMENT": "UNDER DEVELOPMENT",
#             "SCOPING": "UNDER SCOPING",
#             "IS  REVIEW": "IS REVIEW"
#         })
#     )

#     status = team_df["Status_Clean"]


#     # ==========================================
#     # KPI VALUES
#     # ==========================================

#     total = len(team_df)
#     live = len(team_df[status == "LIVE"])
#     uat = len(team_df[status == "UAT"])
#     development = len(team_df[status == "UNDER DEVELOPMENT"])
#     review = len(team_df[status == "IS REVIEW"])
#     cmc = len(team_df[status == "CMC"])
#     scoping = len(team_df[status == "UNDER SCOPING"])


#     # ==========================
#     # KPI CARDS
#     # ==========================

#     k1, k2, k3, k4, k5, k6, k7 = st.columns(7)

#     with k1:
#         st.metric("Total", total)

#     with k2:
#         st.metric("Live", live)

#     with k3:
#         st.metric("UAT", uat)

#     with k4:
#         st.metric("Development", development)

#     with k5:
#         st.metric("IS Review", review)

#     with k6:
#         st.metric("CMC", cmc)

#     with k7:
#         st.metric("Scoping", scoping)

#     st.markdown("<br>", unsafe_allow_html=True)

#     # ==========================================
#     # PROJECT STATUS DISTRIBUTION
#     # ==========================================

#     st.markdown("""
#     <h2 style="
#         color:#006747;
#         font-size:30px;">
#         📊 Project Status Distribution
#     </h2>
#     """, unsafe_allow_html=True)


#     status_order = [
#         "LIVE",
#         "UAT",
#         "UNDER DEVELOPMENT",
#         "IS REVIEW",
#         "CMC",
#         "UNDER SCOPING"
#     ]


#     color_map = {
#         "LIVE": "#00C853",
#         "UAT": "#F9A825",
#         "UNDER DEVELOPMENT": "#FF9800",
#         "IS REVIEW": "#00ACC1",
#         "CMC": "#3949AB",
#         "UNDER SCOPING": "#8E24AA"
#     }


#     status_df = (
#         team_df["Status_Clean"]
#         .value_counts()
#         .reindex(status_order, fill_value=0)
#         .reset_index()
#     )

#     status_df.columns = [
#         "Status",
#         "Projects"
#     ]


#     fig1 = px.bar(
#         status_df,
#         x="Status",
#         y="Projects",
#         text="Projects",
#         color="Status",
#         color_discrete_map=color_map
#     )


#     # ==========================================
#     # CHART DESIGN + DARK MODE FIX
#     # ==========================================

#     fig1.update_layout(
#         height=500,

#         plot_bgcolor="white",
#         paper_bgcolor="white",

#         showlegend=False,

#         font=dict(
#             color="#111827",
#             family="Segoe UI, Arial"
#         ),

#         xaxis=dict(
#             title="",
#             tickangle=0,
#             tickfont=dict(
#                 size=11,
#                 color="#111827"
#             ),
#             automargin=True
#         ),

#         yaxis=dict(
#             title="Projects",
#             title_font=dict(
#                 color="#111827"
#             ),
#             tickfont=dict(
#                 color="#111827"
#             ),
#             gridcolor="#E5E7EB"
#         ),

#         margin=dict(
#             l=30,
#             r=30,
#             t=30,
#             b=100
#         )
#     )


#     fig1.update_traces(
#         textposition="outside",
#         textfont=dict(
#             color="#111827",
#             size=13
#         ),
#         marker_line_width=0
#     )


#     st.plotly_chart(
#         fig1,
#         width="stretch"
#     )


#     st.markdown("<br>", unsafe_allow_html=True)


#     # ==========================================
#     # TEAM WORKLOAD
#     # ==========================================

#     if member == "All":

#         st.markdown("""
#         <h2 style="
#             color:#006747;
#             font-size:30px;">
#             👥 Team Workload
#         </h2>
#         """, unsafe_allow_html=True)


#         allocation_df = (
#             df["Allocation"]
#             .value_counts()
#             .reset_index()
#         )

#         allocation_df.columns = [
#             "Allocation",
#             "Projects"
#         ]


#         fig2 = px.bar(
#             allocation_df,
#             x="Allocation",
#             y="Projects",
#             text="Projects",
#             color="Projects",
#             color_continuous_scale="Greens"
#         )


#         fig2.update_layout(
#             height=500,
#             plot_bgcolor="white",
#             paper_bgcolor="white",

#             font=dict(
#                 color="#111827",
#                 family="Segoe UI, Arial"
#             ),

#             coloraxis_showscale=False,

#             xaxis=dict(
#                 title="",
#                 tickfont=dict(
#                     color="#111827"
#                 )
#             ),

#             yaxis=dict(
#                 title="Projects",
#                 title_font=dict(
#                     color="#111827"
#                 ),
#                 tickfont=dict(
#                     color="#111827"
#                 ),
#                 gridcolor="#E5E7EB"
#             ),

#             margin=dict(
#                 l=20,
#                 r=20,
#                 t=25,
#                 b=30
#             )
#         )


#         fig2.update_traces(
#             textposition="outside",
#             textfont=dict(
#                 color="#111827",
#                 size=13
#             ),
#             marker_line_width=0
#         )


#         st.plotly_chart(
#             fig2,
#             width="stretch"
#         )


#     else:

#         # ==========================================
#         # SELECTED MEMBER PROJECTS - VIP
#         # ==========================================

#         st.markdown(f"""
#         <h2 style="
#         color:#006747;
#         font-size:30px;
#         font-weight:700;
#         margin-top:25px;
#         margin-bottom:15px;">
#         📋 {member} — Project Portfolio
#         </h2>
#         """, unsafe_allow_html=True)


#         # ==========================================
#         # PREPARE MEMBER DATA
#         # ==========================================

#         member_display_df = team_df[
#             [
#                 "Mandate",
#                 "Status",
#                 "Allocation"
#             ]
#         ].copy()


#         # ==========================================
#         # STATUS BADGE
#         # ==========================================

#         def member_status_badge(status):

#             status = str(status).strip()
#             status_upper = status.upper()

#             if status_upper == "LIVE":
#                 css = "status-live"

#             elif status_upper == "UAT":
#                 css = "status-uat"

#             elif status_upper in [
#                 "DEVELOPMENT",
#                 "UNDER DEVELOPMENT",
#                 "SIT"
#             ]:
#                 css = "status-development"

#             elif status_upper == "IS REVIEW":
#                 css = "status-review"

#             elif status_upper == "CMC":
#                 css = "status-cmc"

#             elif status_upper in [
#                 "SCOPING",
#                 "UNDER SCOPING"
#             ]:
#                 css = "status-scoping"

#             else:
#                 css = "status-default"

#             return f'<span class="status-badge {css}">{status}</span>'


#         # ==========================================
#         # APPLY VIP FORMATTING
#         # ==========================================

#         member_display_df["Status"] = (
#             member_display_df["Status"]
#             .apply(member_status_badge)
#         )


#         member_display_df["Mandate"] = (
#             member_display_df["Mandate"]
#             .apply(
#                 lambda x:
#                 f'<span class="project-name">📁 {x}</span>'
#             )
#         )


#         # ==========================================
#         # CREATE VIP HTML TABLE
#         # ==========================================

#         member_table_html = member_display_df.to_html(
#             index=False,
#             escape=False,
#             classes="project-table"
#         )


#         st.markdown(
#             f"""
#             <div class="project-table-wrapper">
#                 {member_table_html}
#             </div>
#             """,
#             unsafe_allow_html=True
#         )


#         st.markdown("<br>", unsafe_allow_html=True)


#     # ==========================================
#     # TOP WORKLOAD
#     # ==========================================

#     st.markdown("""
#     <h2 style="
#         color:#006747;
#         font-size:30px;">
#         🏆 Team Workload Leader
#     </h2>
#     """, unsafe_allow_html=True)


#     top = (
#         df["Allocation"]
#         .value_counts()
#         .reset_index()
#     )

#     top.columns = [
#         "Member",
#         "Projects"
#     ]


#     if not top.empty:

#         winner = top.iloc[0]


#         st.html(f"""
#         <div style="
#             background:linear-gradient(
#                 135deg,
#                 #ffffff,
#                 #f1f8f5
#             );
#             border-radius:22px;
#             padding:24px;
#             border:1px solid #D1FAE5;
#             border-left:7px solid #006747;
#             box-shadow:0 10px 28px rgba(0,0,0,.08);
#             font-family:Segoe UI,Arial,sans-serif;">

#             <div style="
#                 color:#6B7280;
#                 font-size:13px;
#                 font-weight:600;
#                 text-transform:uppercase;
#                 letter-spacing:1px;">
#                 Highest Project Workload
#             </div>

#             <div style="
#                 color:#006747;
#                 font-size:28px;
#                 font-weight:700;
#                 margin-top:8px;">
#                 🏆 {winner["Member"]}
#             </div>

#             <div style="
#                 color:#111827;
#                 font-size:17px;
#                 margin-top:5px;">
#                 Currently handling
#                 <b>{winner["Projects"]}</b>
#                 project(s)
#             </div>

#         </div>
#         """)
# # =====================================================
# # VOICE SEARCH
# # =====================================================

# elif page == "Voice Search":

#     # ==========================================
#     # VIP HEADER
#     # ==========================================

#     st.html("""
#     <div style="
#         background:linear-gradient(135deg,#ffffff,#f4fbf8);
#         border-radius:24px;
#         padding:30px;
#         border:1px solid #DDEBE5;
#         box-shadow:0 14px 35px rgba(0,0,0,.10);
#         margin-bottom:22px;
#         font-family:Segoe UI,Arial,sans-serif;">

#         <div style="
#             color:#006747;
#             font-size:14px;
#             font-weight:700;
#             letter-spacing:2px;
#             margin-bottom:8px;">
#             SMARTPAY INTELLIGENT SEARCH
#         </div>

#         <div style="
#             color:#006747;
#             font-size:40px;
#             font-weight:700;">
#             🎤 Voice Search
#         </div>

#         <div style="
#             color:#6B7280;
#             font-size:17px;
#             margin-top:10px;">
#             Find SmartPay projects instantly using your voice.
#         </div>

#     </div>
#     """)

#     # ==========================================
#     # HOW TO SEARCH
#     # ==========================================

#     st.html("""
#     <div style="
#         background:white;
#         border-radius:20px;
#         padding:22px;
#         border:1px solid #E5E7EB;
#         box-shadow:0 8px 22px rgba(0,0,0,.06);
#         margin-bottom:22px;">

#         <div style="
#             color:#006747;
#             font-size:18px;
#             font-weight:700;
#             margin-bottom:8px;">
#             🎙️ How to Search
#         </div>

#         <div style="
#             color:#4B5563;
#             font-size:15px;
#             line-height:1.7;">
#             Speak a <b>Project Name</b>, <b>Team Member</b>,
#             <b>Status</b>, <b>Category</b> or say
#             <b>"Show All Projects"</b>.
#         </div>

#     </div>
#     """)

#     # ==========================================
#     # VOICE COMMAND GUIDE
#     # ==========================================

#     st.markdown("""
#     <h2 style="
#         color:#006747;
#         font-size:28px;
#         margin-bottom:15px;">
#         💡 Voice Commands
#     </h2>
#     """, unsafe_allow_html=True)

#     vc1, vc2, vc3, vc4 = st.columns(4)

#     with vc1:

#         st.html("""
#         <div style="
#             background:#F0FDF4;
#             border:1px solid #BBF7D0;
#             border-radius:18px;
#             padding:20px;
#             min-height:125px;">

#             <div style="font-size:25px;">
#                 🟢
#             </div>

#             <div style="
#                 color:#166534;
#                 font-weight:700;
#                 margin-top:8px;">
#                 Live Projects
#             </div>

#             <div style="
#                 color:#6B7280;
#                 font-size:13px;
#                 margin-top:5px;">
#                 Say: <b>Live</b>
#             </div>

#         </div>
#         """)

#     with vc2:

#         st.html("""
#         <div style="
#             background:#FFFBEB;
#             border:1px solid #FDE68A;
#             border-radius:18px;
#             padding:20px;
#             min-height:125px;">

#             <div style="font-size:25px;">
#                 🟡
#             </div>

#             <div style="
#                 color:#92400E;
#                 font-weight:700;
#                 margin-top:8px;">
#                 UAT Projects
#             </div>

#             <div style="
#                 color:#6B7280;
#                 font-size:13px;
#                 margin-top:5px;">
#                 Say: <b>UAT</b>
#             </div>

#         </div>
#         """)

#     with vc3:

#         st.html("""
#         <div style="
#             background:#EFF6FF;
#             border:1px solid #BFDBFE;
#             border-radius:18px;
#             padding:20px;
#             min-height:125px;">

#             <div style="font-size:25px;">
#                 🔎
#             </div>

#             <div style="
#                 color:#1D4ED8;
#                 font-weight:700;
#                 margin-top:8px;">
#                 Project Search
#             </div>

#             <div style="
#                 color:#6B7280;
#                 font-size:13px;
#                 margin-top:5px;">
#                 Say the <b>project name</b>
#             </div>

#         </div>
#         """)

#     with vc4:

#         st.html("""
#         <div style="
#             background:#F5F3FF;
#             border:1px solid #DDD6FE;
#             border-radius:18px;
#             padding:20px;
#             min-height:125px;">

#             <div style="font-size:25px;">
#                 👥
#             </div>

#             <div style="
#                 color:#5B21B6;
#                 font-weight:700;
#                 margin-top:8px;">
#                 Team Search
#             </div>

#             <div style="
#                 color:#6B7280;
#                 font-size:13px;
#                 margin-top:5px;">
#                 Say the <b>team member</b>
#             </div>

#         </div>
#         """)

#     st.markdown("<br>", unsafe_allow_html=True)

#     # ==========================================
#     # MICROPHONE AREA
#     # ==========================================

#     st.html("""
#     <div style="
#         background:linear-gradient(135deg,#006747,#00875A);
#         border-radius:22px;
#         padding:25px;
#         text-align:center;
#         color:white;
#         box-shadow:0 12px 30px rgba(0,103,71,.20);
#         margin-bottom:20px;">

#         <div style="
#             font-size:42px;">
#             🎤
#         </div>

#         <div style="
#             font-size:22px;
#             font-weight:700;
#             margin-top:8px;">
#             Speak Your Command
#         </div>

#         <div style="
#             font-size:14px;
#             opacity:.9;
#             margin-top:6px;">
#             Use your microphone to search SmartPay projects
#         </div>

#     </div>
#     """)

#     # =====================================================
#     # ORIGINAL VOICE CODE — DO NOT CHANGE
#     # =====================================================

#     voice_text = listen()

#     st.write("Raw Voice :", repr(voice_text))

#     voice_text = normalize_voice(str(voice_text))

#     st.write("Normalized :", voice_text)

#     if voice_text:

#         voice_text = voice_text.lower().strip()

#         st.success(
#             f"🎤 You said: {voice_text}"
#         )

#         speak(
#             f"You said {voice_text}"
#         )

#         # =========================================
#         # SMART COMMANDS
#         # =========================================

#         if voice_text == "all":

#             result = df.copy()

#         elif voice_text == "live":

#             result = df[
#                 df["Status"]
#                 .astype(str)
#                 .str.lower() == "live"
#             ]

#         elif voice_text == "uat":

#             result = df[
#                 df["Status"]
#                 .astype(str)
#                 .str.lower() == "uat"
#             ]

#         elif voice_text == "sit":

#             result = df[
#                 df["Status"]
#                 .astype(str)
#                 .str.lower() == "sit"
#             ]

#         elif voice_text == "under development":

#             result = df[
#                 df["Status"]
#                 .astype(str)
#                 .str.lower() == "under development"
#             ]

#         else:

#             search_cols = [
#                 "Mandate",
#                 "Allocation",
#                 "Status",
#                 "Category",
#                 "Update"
#             ]

#             result = df[
#                 df.apply(
#                     lambda row: any(
#                         voice_text in str(
#                             row[col]
#                         ).lower()
#                         for col in search_cols
#                     ),
#                     axis=1
#                 )
#             ]

#             # =========================================
#             # FUZZY SEARCH
#             # =========================================

#             if result.empty:

#                 search_values = []

#                 for col in search_cols:

#                     search_values.extend(
#                         df[col]
#                         .dropna()
#                         .astype(str)
#                         .tolist()
#                     )

#                 match = process.extractOne(
#                     voice_text,
#                     search_values,
#                     scorer=fuzz.token_sort_ratio
#                 )

#                 if match and match[1] >= 70:

#                     matched = match[0].lower()

#                     result = df[
#                         df.apply(
#                             lambda row: any(
#                                 matched in str(
#                                     row[col]
#                                 ).lower()
#                                 for col in search_cols
#                             ),
#                             axis=1
#                         )
#                     ]

#         # =========================================
#         # RESULT
#         # =========================================

#         if result.empty:

#             st.error(
#                 "❌ No Project Found"
#             )

#             speak(
#                 "Sorry. No matching project found."
#             )

#         else:

#             st.success(
#                 f"✅ {len(result)} Project(s) Found"
#             )

#             st.metric(
#                 "Total Results",
#                 len(result)
#             )

#             # =========================================
#             # SINGLE RESULT
#             # =========================================

#             if len(result) == 1:

#                 first = result.iloc[0]

#                 status = str(
#                     first["Status"]
#                 ).upper()

#                 if status == "LIVE":

#                     badge = "#16A34A"

#                 elif status == "UAT":

#                     badge = "#F59E0B"

#                 elif status == "SIT":

#                     badge = "#2563EB"

#                 else:

#                     badge = "#6B7280"

#                 st.markdown(
#                     f"""
#                     <div style="
#                     background:white;
#                     border-radius:18px;
#                     padding:25px;
#                     border-left:8px solid {badge};
#                     box-shadow:0 8px 18px rgba(0,0,0,.08);
#                     margin-bottom:20px;">

#                     <h2 style="
#                     color:#006747;
#                     margin-top:0;">
#                     {first['Mandate']}
#                     </h2>

#                     <table style="
#                     width:100%;
#                     font-size:16px;">

#                     <tr>
#                     <td><b>Status</b></td>
#                     <td>{first['Status']}</td>
#                     </tr>

#                     <tr>
#                     <td><b>Owner</b></td>
#                     <td>{first['Allocation']}</td>
#                     </tr>

#                     <tr>
#                     <td><b>Category</b></td>
#                     <td>{first['Category']}</td>
#                     </tr>

#                     <tr>
#                     <td><b>Latest Update</b></td>
#                     <td>{first['Update']}</td>
#                     </tr>

#                     </table>

#                     </div>
#                     """,
#                     unsafe_allow_html=True
#                 )

#                 response = (
#                     f"{first['Mandate']} is currently "
#                     f"{first['Status']} and allocated to "
#                     f"{first['Allocation']}. "
#                     f"Latest update is "
#                     f"{first['Update']}."
#                 )

#             # =========================================
#             # MULTIPLE RESULTS
#             # =========================================

#             else:

#                 st.dataframe(
#                     result,
#                     use_container_width=True,
#                     hide_index=True
#                 )

#                 st.markdown(
#                     "### 📋 Projects Found"
#                 )

#                 c1, c2, c3 = st.columns(3)

#                 with c1:

#                     st.metric(
#                         "Projects",
#                         len(result)
#                     )

#                 with c2:

#                     st.metric(
#                         "Owners",
#                         result["Allocation"].nunique()
#                     )

#                 with c3:

#                     st.metric(
#                         "Live",
#                         len(
#                             result[
#                                 result["Status"]
#                                 .astype(str)
#                                 .str.upper() == "LIVE"
#                             ]
#                         )
#                     )

#                 for _, row in result.iterrows():

#                     st.markdown(
#                         f"""
#                         <div style="
#                         background:white;
#                         padding:18px;
#                         border-radius:15px;
#                         margin-bottom:12px;
#                         border:1px solid #E5E7EB;
#                         box-shadow:0 4px 10px rgba(0,0,0,.06);">

#                         <h4 style="
#                         color:#006747;
#                         margin:0;">
#                         {row['Mandate']}
#                         </h4>

#                         <p style="
#                         margin-top:8px;">

#                         <b>Owner:</b>
#                         {row['Allocation']}<br>

#                         <b>Status:</b>
#                         {row['Status']}<br>

#                         <b>Category:</b>
#                         {row['Category']}

#                         </p>

#                         </div>
#                         """,
#                         unsafe_allow_html=True
#                     )

#                 names = ", ".join(
#                     result["Mandate"]
#                     .astype(str)
#                     .tolist()
#                 )

#                 response = (
#                     f"{len(result)} projects found. "
#                     f"The projects are {names}."
#                 )

#             # =========================================
#             # VOICE RESPONSE
#             # =========================================

#             st.info(response)

#             speak(response)

# =====================================================
# BAU MONITORING
# =====================================================

elif page == "BAU Monitoring":

    # =================================================
    # BAU DATA
    # =================================================

    bau_df = df[
        df["Status"]
        .astype(str)
        .str.strip()
        .str.upper()
        == "BAU"
    ].copy()


    # =================================================
    # COMPACT HEADER
    # =================================================

    st.html("""
    <div style="
        background:linear-gradient(135deg,#ffffff,#f7fbf9);
        border-radius:18px;
        padding:14px 22px;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 22px rgba(0,0,0,.08);
        margin-bottom:8px;
        font-family:Segoe UI,Arial,sans-serif;
    ">

        <div style="
            color:#006747;
            font-size:12px;
            font-weight:700;
            letter-spacing:1.8px;
            margin-bottom:4px;">
            SMARTPAY BUSINESS AS USUAL
        </div>

        <div style="
            color:#006747;
            font-size:28px;
            font-weight:800;
            line-height:1.15;">
            🏦 BAU Monitoring
        </div>

        <div style="
            color:#6B7280;
            font-size:14px;
            margin-top:4px;">
            Live Business Operations Monitoring, ownership and ongoing updates.
        </div>

    </div>
    """)


    # =================================================
    # KPI VALUES
    # =================================================

    total_bau = len(bau_df)

    bau_members = (
        bau_df["Allocation"]
        .dropna()
        .astype(str)
        .nunique()
    )

    bau_updates = (
        bau_df["Update"]
        .notna()
        .sum()
        if "Update" in bau_df.columns
        else 0
    )

    bau_categories = (
        bau_df["Category"]
        .dropna()
        .astype(str)
        .nunique()
        if "Category" in bau_df.columns
        else 0
    )


    # =================================================
    # COMPACT KPI CARDS
    # =================================================

    k1, k2, k3, k4 = st.columns(4)


    def bau_card(title, value, color):

        st.html(
            f"""
            <div style="
                background:#FFFFFF;
                border-radius:14px;
                padding:10px 8px;
                height:88px;
                border-top:5px solid {color};
                box-shadow:0 4px 12px rgba(0,0,0,.06);
                display:flex;
                flex-direction:column;
                justify-content:center;
                align-items:center;
                text-align:center;
                font-family:Segoe UI,Arial,sans-serif;
            ">

                <div style="
                    color:#6B7280;
                    font-size:11px;
                    font-weight:700;
                    margin-bottom:5px;">
                    {title}
                </div>

                <div style="
                    color:{color};
                    font-size:28px;
                    font-weight:800;
                    line-height:1;">
                    {value}
                </div>

            </div>
            """
        )


    with k1:
        bau_card(
            "Total BAU Projects",
            total_bau,
            "#006747"
        )

    with k2:
        bau_card(
            "Team Members",
            bau_members,
            "#3949AB"
        )

    with k3:
        bau_card(
            "Updated Projects",
            bau_updates,
            "#F9A825"
        )

    with k4:
        bau_card(
            "Categories",
            bau_categories,
            "#00ACC1"
        )


    st.markdown(
        "<div style='height:8px;'></div>",
        unsafe_allow_html=True
    )


    # =================================================
    # FILTERS
    # =================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:2px;
        margin-bottom:5px;
        font-family:Segoe UI,Arial,sans-serif;">
        🎯 BAU Filters
    </div>
    """)


    f1, f2 = st.columns(2)


    with f1:

        if "Allocation" in bau_df.columns:

            selected_bau_member = st.selectbox(
                "👤 Team Member",
                ["All"] +
                sorted(
                    bau_df["Allocation"]
                    .dropna()
                    .astype(str)
                    .unique()
                ),
                key="bau_member_filter"
            )

        else:

            selected_bau_member = "All"


    with f2:

        if "Category" in bau_df.columns:

            selected_bau_category = st.selectbox(
                "📂 Category",
                ["All"] +
                sorted(
                    bau_df["Category"]
                    .dropna()
                    .astype(str)
                    .unique()
                ),
                key="bau_category_filter"
            )

        else:

            selected_bau_category = "All"


    # =================================================
    # APPLY FILTERS
    # =================================================

    filtered_bau_df = bau_df.copy()


    if selected_bau_member != "All":

        filtered_bau_df = filtered_bau_df[
            filtered_bau_df["Allocation"].astype(str)
            == selected_bau_member
        ]


    if selected_bau_category != "All":

        filtered_bau_df = filtered_bau_df[
            filtered_bau_df["Category"].astype(str)
            == selected_bau_category
        ]


    # =================================================
    # SEARCH
    # =================================================

    bau_search = st.text_input(
        "",
        placeholder="🔍 Search BAU Project...",
        label_visibility="collapsed",
        key="bau_project_search"
    )


    if bau_search:

        filtered_bau_df = filtered_bau_df[
            filtered_bau_df["Mandate"]
            .astype(str)
            .str.contains(
                bau_search,
                case=False,
                na=False
            )
        ]


    st.markdown(
        "<div style='height:8px;'></div>",
        unsafe_allow_html=True
    )


    # =================================================
    # OWNER SUMMARY
    # =================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:2px;
        margin-bottom:7px;
        font-family:Segoe UI,Arial,sans-serif;">
        👥 BAU Owner Summary
    </div>
    """)


    owner_summary = (
        filtered_bau_df
        .groupby("Allocation")
        .size()
        .reset_index(name="Projects")
        .sort_values(
            "Projects",
            ascending=False
        )
    )


    if owner_summary.empty:

        st.info("No BAU projects found.")

    else:

        owner_cols = st.columns(
            min(4, len(owner_summary))
        )


        for i, (_, owner) in enumerate(
            owner_summary.iterrows()
        ):

            with owner_cols[
                i % len(owner_cols)
            ]:

                st.html(
                    f"""
                    <div style="
                        background:linear-gradient(
                            180deg,
                            #ffffff,
                            #f7fbf9
                        );
                        border:1px solid #E5E7EB;
                        border-radius:14px;
                        padding:10px 8px;
                        text-align:center;
                        box-shadow:0 4px 12px rgba(0,0,0,.05);
                        margin-bottom:8px;
                        font-family:Segoe UI,Arial,sans-serif;
                    ">

                        <div style="
                            font-size:22px;
                            line-height:1;">
                            👤
                        </div>

                        <div style="
                            color:#006747;
                            font-size:14px;
                            font-weight:700;
                            margin-top:5px;">
                            {owner["Allocation"]}
                        </div>

                        <div style="
                            color:#111827;
                            font-size:26px;
                            font-weight:800;
                            margin-top:4px;
                            line-height:1;">
                            {owner["Projects"]}
                        </div>

                        <div style="
                            color:#6B7280;
                            font-size:10px;
                            margin-top:3px;">
                            BAU Projects
                        </div>

                    </div>
                    """
                )

    # =================================================
    # BAU PROJECT TABLE
    # =================================================

    st.markdown("""
    <h2 style="
        color:#006747;
        font-size:30px;
        font-weight:700;
        margin-top:20px;
        margin-bottom:18px;">
        📋 BAU Project Portfolio
    </h2>
    """, unsafe_allow_html=True)


    display_bau = filtered_bau_df.copy()


    

    # =================================================
    # VIP TABLE CSS
    # =================================================

    st.markdown("""
    <style>

    .bau-table-wrapper {
        background:#FFFFFF;
        border:1px solid #DDE5E1;
        border-radius:18px;
        padding:6px;
        box-shadow:0 8px 25px rgba(0,103,71,.08);
        overflow-x:auto;
    }

    .bau-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:14px;
    }

    .bau-table thead th {
        background:#006747;
        color:white;
        padding:14px 12px;
        font-weight:700;
        text-align:left;
    }

    .bau-table tbody td {
        padding:13px 12px;
        color:#1F2937;
        border-bottom:1px solid #E5E7EB;
        background:#FFFFFF;
    }

    .bau-table tbody tr:nth-child(even) td {
        background:#F8FAFC;
    }

    .bau-table tbody tr:hover td {
        background:#ECFDF5;
    }

    .bau-project-name {
        color:#006747 !important;
        font-weight:700;
    }

    .bau-status {
        display:inline-block;
        padding:5px 12px;
        border-radius:20px;
        background:#E8F5E9;
        color:#166534;
        font-size:11px;
        font-weight:800;
    }

    </style>
    """, unsafe_allow_html=True)


    # =================================================
    # TABLE FORMATTING
    # =================================================

    if "Mandate" in display_bau.columns:

        display_bau["Mandate"] = (
            display_bau["Mandate"]
            .apply(
                lambda x:
                f'<span class="bau-project-name">📁 {x}</span>'
            )
        )


    if "Status" in display_bau.columns:

        display_bau["Status"] = (
            display_bau["Status"]
            .apply(
                lambda x:
                '<span class="bau-status">🟢 BAU</span>'
            )
        )


    # =================================================
    # SELECT USEFUL COLUMNS
    # =================================================

    preferred_columns = [
        "Mandate",
        "Allocation",
        "Status",
        "Category",
        "Update"
    ]


    table_columns = [
        col
        for col in preferred_columns
        if col in display_bau.columns
    ]


    if table_columns:

        display_bau = display_bau[
            table_columns
        ]


    # =================================================
    # CREATE TABLE
    # =================================================

    bau_table_html = display_bau.to_html(
        index=False,
        escape=False,
        classes="bau-table"
    )


    st.markdown(
        f"""
        <div class="bau-table-wrapper">
            {bau_table_html}
        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown("<br>", unsafe_allow_html=True)


    # =================================================
    # RECENT / ONGOING UPDATES
    # =================================================

    if "Update" in filtered_bau_df.columns:

        st.markdown("""
        <h2 style="
            color:#006747;
            font-size:28px;
            font-weight:700;
            margin-top:20px;
            margin-bottom:18px;">
            🔄 BAU Monitoring Updates
        </h2>
        """, unsafe_allow_html=True)


        updates_df = filtered_bau_df[
            [
                col
                for col in [
                    "Mandate",
                    "Allocation",
                    "Update"
                ]
                if col in filtered_bau_df.columns
            ]
        ].copy()


        if not updates_df.empty:

            for _, row in updates_df.iterrows():

                project_name = str(
                    row.get("Mandate", "")
                )

                owner_name = str(
                    row.get("Allocation", "")
                )

                update_text = str(
                    row.get(
                        "Update",
                        "Business as usual."
                    )
                )


                st.html(
                    f"""
                    <div style="
                        background:white;
                        border:1px solid #E5E7EB;
                        border-left:5px solid #006747;
                        border-radius:14px;
                        padding:15px 18px;
                        margin-bottom:10px;
                        box-shadow:0 4px 12px rgba(0,0,0,.05);
                    ">

                        <div style="
                            display:flex;
                            justify-content:space-between;
                            align-items:center;
                            gap:15px;
                        ">

                            <div style="
                                color:#006747;
                                font-size:16px;
                                font-weight:700;">
                                📁 {project_name}
                            </div>

                            <div style="
                                color:#6B7280;
                                font-size:12px;
                                font-weight:600;">
                                👤 {owner_name}
                            </div>

                        </div>

                        <div style="
                            color:#374151;
                            font-size:14px;
                            margin-top:9px;
                            line-height:1.5;">
                            {update_text}
                        </div>

                    </div>
                    """
                )

    else:

        st.info(
            "No BAU update field is available in the Excel data."
        )

# =====================================================
# CRPL
# =====================================================

elif page == "CRPL":

    import os
    import pandas as pd

    # =====================================================
    # LOAD CRPL FILE
    # =====================================================

    file = os.path.join(
        os.path.dirname(__file__),
        "data",
        "CRPL_All_Data_Exactly_20.xlsx"
    )

    try:

        raw = pd.read_excel(
            file,
            sheet_name="CRPL",
            header=None
        )

        header_row = None

        for i in range(min(20, len(raw))):

            row = " ".join(
                str(x).lower()
                for x in raw.iloc[i]
                if pd.notna(x)
            )

            if "crf no" in row and "crf name" in row:
                header_row = i
                break

        if header_row is None:

            st.error("CRPL header row not found.")
            st.dataframe(raw.head(20))
            st.stop()

        df = pd.read_excel(
            file,
            sheet_name="CRPL",
            header=header_row
        )

        df.columns = [
            str(x).strip()
            for x in df.columns
        ]

        df = df.dropna(
            how="all"
        ).reset_index(drop=True)

    except Exception as e:

        st.error(f"CRPL file error: {e}")
        st.stop()

    # =====================================================
    # REQUIRED COLUMNS
    # =====================================================

    columns = [
        "CRF No",
        "CRF Name",
        "Stage",
        "CRPL Remarks",
        "NBP Remarks"
    ]

    for col in columns:

        if col not in df.columns:
            df[col] = ""

    df = df[columns].fillna("")

    # =====================================================
    # HEADER
    # =====================================================

    st.html("""
    <div style="
        background:linear-gradient(135deg,#ffffff,#f8fbff);
        border-radius:18px;
        padding:14px 22px;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 22px rgba(0,0,0,.08);
        box-sizing:border-box;
        margin-bottom:2px;
        font-family:Segoe UI,Arial,sans-serif;">

        <div style="
            color:#006747;
            font-size:12px;
            font-weight:700;
            letter-spacing:1.8px;
            margin-bottom:4px;">
            CRPL MONITORING
        </div>

        <div style="
            color:#006747;
            font-size:28px;
            font-weight:800;
            line-height:1.15;">
            🏦 CRPL
        </div>

        <div style="
            color:#6B7280;
            font-size:14px;
            margin-top:4px;">
            CRPL Issues &amp; Progress Tracking
        </div>

    </div>
    """)


    # =====================================================
    # FILTERS
    # =====================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:0;
        margin-bottom:3px;
        font-family:Segoe UI,Arial,sans-serif;">
        🎯 Filters
    </div>
    """)

    c1, c2 = st.columns(2)

    with c1:

        search = st.text_input(
            "Search CRF",
            placeholder="🔍 Search CRF No or CRF Name..."
        )

    with c2:

        stage_filter = st.selectbox(
            "Stage",
            [
                "All",
                "HOLD",
                "WIP",
                "LIVE",
                "UAT"
            ]
        )

    # =====================================================
    # FILTER DATA
    # =====================================================

    filtered_df = df.copy()

    if search:

        filtered_df = filtered_df[
            filtered_df["CRF No"]
            .astype(str)
            .str.contains(
                search,
                case=False,
                na=False
            )
            |
            filtered_df["CRF Name"]
            .astype(str)
            .str.contains(
                search,
                case=False,
                na=False
            )
        ]

    if stage_filter != "All":

        if stage_filter == "LIVE":

            filtered_df = filtered_df[
                filtered_df["Stage"]
                .astype(str)
                .str.upper()
                .isin([
                    "LIVE",
                    "PRODUCTION",
                    "LIVE / PRODUCTION",
                    "LIVE/PRODUCTION"
                ])
            ]

        else:

            filtered_df = filtered_df[
                filtered_df["Stage"]
                .astype(str)
                .str.upper()
                == stage_filter
            ]

    # =====================================================
    # KPI COUNTS
    # =====================================================

    stage = (
        df["Stage"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    hold_count = (
        stage == "HOLD"
    ).sum()

    wip_count = (
        stage == "WIP"
    ).sum()

    uat_count = (
        stage == "UAT"
    ).sum()

    live_count = stage.isin([
        "LIVE",
        "PRODUCTION",
        "LIVE / PRODUCTION",
        "LIVE/PRODUCTION"
    ]).sum()

    # =====================================================
    # VIP CLICKABLE KPI CARDS
    # =====================================================

    st.markdown("""
    <style>

    /* =====================================================
    KPI CARD CONTAINER
    ===================================================== */

    .st-key-crpl_kpis div[data-testid="stButton"] {
        width:100%;
    }


    /* =====================================================
    BASE CARD
    ===================================================== */

    .st-key-crpl_kpis div[data-testid="stButton"] button {

        width:100% !important;
        min-height:88px !important;

        background:#FFFFFF !important;

        border:1px solid #E5E7EB !important;
        border-radius:14px !important;

        padding:10px 6px !important;

        box-shadow:
            0 4px 12px rgba(0,0,0,.06) !important;

        color:#111827 !important;

        font-size:13px !important;
        font-weight:700 !important;

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        line-height:1.25 !important;

        white-space:pre-line !important;

        text-align:center !important;

        transition:all .2s ease !important;
    }


    /* =====================================================
    HOVER
    ===================================================== */

    .st-key-crpl_kpis div[data-testid="stButton"] button:hover {

        background:#F8FAFC !important;

        color:#111827 !important;

        transform:translateY(-2px);

        box-shadow:
            0 8px 18px rgba(0,0,0,.10) !important;
    }

    .st-key-crpl_kpis
    div[data-testid="stButton"] button:hover p {

        color:#111827 !important;
    }


    /* =====================================================
    SELECTED CARD
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stButton"] button:focus {

        outline:none !important;

        background:#EAF5F0 !important;

        border:2px solid #006747 !important;

        box-shadow:
            0 0 0 3px rgba(0,103,71,0.12),
            0 8px 18px rgba(0,103,71,0.15) !important;

        color:#006747 !important;
    }

    .st-key-crpl_kpis
    div[data-testid="stButton"] button:focus p {

        color:#006747 !important;

        font-weight:700 !important;
    }


    /* =====================================================
    ALL
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(1)
    div[data-testid="stButton"] button {

        border-top:5px solid #006747 !important;
    }


    /* =====================================================
    HOLD
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(2)
    div[data-testid="stButton"] button {

        border-top:5px solid #607D8B !important;
    }


    /* =====================================================
    WIP
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(3)
    div[data-testid="stButton"] button {

        border-top:5px solid #FF9800 !important;
    }


    /* =====================================================
    LIVE
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(4)
    div[data-testid="stButton"] button {

        border-top:5px solid #00C853 !important;
    }


    /* =====================================================
    UAT
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(5)
    div[data-testid="stButton"] button {

        border-top:5px solid #F9A825 !important;
    }


    /* =====================================================
    BUTTON TEXT
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stButton"] button p {

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        font-size:13px !important;

        font-weight:700 !important;

        line-height:1.3 !important;

        color:#111827 !important;
    }


    /* =====================================================
    REMOVE EXTRA GAPS
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stVerticalBlock"] {

        gap:0 !important;
    }


    /* =====================================================
    COLUMN SPACING
    ===================================================== */

    .st-key-crpl_kpis
    div[data-testid="stHorizontalBlock"] {

        gap:8px !important;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # KPI CARD CONTAINER
    # =====================================================

    with st.container(key="crpl_kpis"):

        c1, c2, c3, c4, c5 = st.columns(5)


        # ALL
        with c1:

            if st.button(
                f"ALL\n{len(df)}",
                key="crpl_all",
                use_container_width=True
            ):
                st.session_state["crpl_stage"] = "All"
                st.rerun()


        # HOLD
        with c2:

            if st.button(
                f"HOLD\n{hold_count}",
                key="crpl_hold",
                use_container_width=True
            ):
                st.session_state["crpl_stage"] = "HOLD"
                st.rerun()


        # WIP
        with c3:

            if st.button(
                f"WIP\n{wip_count}",
                key="crpl_wip",
                use_container_width=True
            ):
                st.session_state["crpl_stage"] = "WIP"
                st.rerun()


        # LIVE
        with c4:

            if st.button(
                f"LIVE\n{live_count}",
                key="crpl_live",
                use_container_width=True
            ):
                st.session_state["crpl_stage"] = "LIVE"
                st.rerun()


        # UAT
        with c5:

            if st.button(
                f"UAT\n{uat_count}",
                key="crpl_uat",
                use_container_width=True
            ):
                st.session_state["crpl_stage"] = "UAT"
                st.rerun()


    st.markdown(
        "<div style='height:8px;'></div>",
        unsafe_allow_html=True
    )
    # =====================================================
    # CARD FILTER
    # =====================================================

    selected_card = st.session_state.get(
        "crpl_stage",
        "All"
    )

    if selected_card != "All":

        if selected_card == "LIVE":

            filtered_df = filtered_df[
                filtered_df["Stage"]
                .astype(str)
                .str.upper()
                .isin([
                    "LIVE",
                    "PRODUCTION",
                    "LIVE / PRODUCTION",
                    "LIVE/PRODUCTION"
                ])
            ]

        else:

            filtered_df = filtered_df[
                filtered_df["Stage"]
                .astype(str)
                .str.upper()
                == selected_card
            ]


    # =====================================================
    # TABLE TITLE
    # =====================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:2px;
        margin-bottom:3px;
        font-family:Segoe UI,Arial,sans-serif;">
        📋 CRPL Details
    </div>
    """)


    # =====================================================
    # VIP TABLE CSS - GREEN THEME
    # =====================================================

    st.markdown("""
    <style>

    .crpl-table-wrapper {
        background:#FFFFFF;
        border:1px solid #DDE5E1;
        border-radius:16px;
        padding:6px;
        box-shadow:0 6px 20px rgba(0,103,71,0.08);
        overflow:hidden;
    }

    .crpl-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:14px;
        table-layout:fixed;
    }

    .crpl-table thead th {
        background:#006747;
        color:white;
        font-weight:700;
        padding:14px 12px;
        text-align:left;
    }

    .crpl-table thead th:first-child {
        border-top-left-radius:11px;
    }

    .crpl-table thead th:last-child {
        border-top-right-radius:11px;
    }

    .crpl-table tbody td {
        padding:13px 12px;
        color:#1F2937;
        border-bottom:1px solid #E5E7EB;
        background:#FFFFFF;

        white-space:normal !important;
        word-wrap:break-word !important;
        overflow-wrap:anywhere !important;
        vertical-align:top;
        line-height:1.6;
    }

    .crpl-table tbody tr:nth-child(even) td {
        background:#F8FAFC;
    }

    .crpl-table tbody tr:hover td {
        background:#ECFDF5;
    }

    .crf-name {
        font-weight:700;
        color:#006747 !important;
    }

    .stage-badge {
        display:inline-block;
        padding:5px 11px;
        border-radius:20px;
        font-size:11px;
        font-weight:800;
        white-space:nowrap;
    }

    .stage-hold {
        background:#FEE2E2;
        color:#991B1B;
    }

    .stage-wip {
        background:#DBEAFE;
        color:#1E40AF;
    }

    .stage-uat {
        background:#FEF3C7;
        color:#92400E;
    }

    .stage-live {
        background:#DCFCE7;
        color:#166534;
    }

    .stage-default {
        background:#F3F4F6;
        color:#374151;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # TABLE DATA
    # =====================================================

    display_df = filtered_df.copy()


    # =====================================================
    # CRF NAME
    # =====================================================

    display_df["CRF Name"] = display_df[
        "CRF Name"
    ].apply(
        lambda x:
        f'<span class="crf-name">{x}</span>'
    )


    # =====================================================
    # STAGE BADGE
    # =====================================================

    def stage_badge(value):

        value = str(value).strip()
        upper = value.upper()

        if upper == "HOLD":

            css = "stage-hold"

        elif upper == "WIP":

            css = "stage-wip"

        elif upper == "UAT":

            css = "stage-uat"

        elif upper in [
            "LIVE",
            "PRODUCTION",
            "LIVE / PRODUCTION",
            "LIVE/PRODUCTION"
        ]:

            css = "stage-live"

        else:

            css = "stage-default"

        return (
            f'<span class="stage-badge {css}">'
            f'{value}'
            f'</span>'
        )


    display_df["Stage"] = display_df[
        "Stage"
    ].apply(stage_badge)


    # =====================================================
    # HTML TABLE
    # =====================================================

    table_html = display_df.to_html(
        index=False,
        escape=False,
        classes="crpl-table"
    )


    st.markdown(
        f"""
        <div class="crpl-table-wrapper">
            {table_html}
        </div>
        """,
        unsafe_allow_html=True
    )

    # =====================================================
    # EDIT CRPL BUTTON
    # =====================================================

    st.markdown(
        "<br>",
        unsafe_allow_html=True
    )


    if "crpl_edit_mode" not in st.session_state:

        st.session_state["crpl_edit_mode"] = False


    if "crpl_new_field" not in st.session_state:

        st.session_state["crpl_new_field"] = ""


    # =====================================================
    # EDIT BUTTON
    # =====================================================

    if not st.session_state["crpl_edit_mode"]:

        if st.button(
            "✏️ Edit CRPL",
            key="crpl_edit_button",
            type="secondary",
            use_container_width=True
        ):

            st.session_state["crpl_edit_mode"] = True

            st.rerun()


    # =====================================================
    # EDIT MODE
    # =====================================================

    if st.session_state["crpl_edit_mode"]:

        st.markdown("""
        <h2 style="
        color:#006747;
        font-size:28px;
        font-weight:700;
        margin-top:20px;
        margin-bottom:15px;">
        ✏️ Edit CRPL
        </h2>
        """, unsafe_allow_html=True)


        st.info(
            "✏️ Edit existing records, add new records, "
            "or create a new field/column."
        )


        # =================================================
        # ADD NEW FIELD
        # =================================================

        st.markdown("""
        <h3 style="
        color:#006747;
        font-size:20px;
        font-weight:700;
        margin-top:10px;">
        ➕ Add New Field
        </h3>
        """, unsafe_allow_html=True)


        field_col1, field_col2 = st.columns([3, 1])


        with field_col1:

            new_field = st.text_input(
                "New Field Name",
                placeholder="e.g. Vendor Remarks",
                key="crpl_new_field_input"
            )


        with field_col2:

            st.markdown("<br>", unsafe_allow_html=True)

            if st.button(
                "➕ Add Field",
                key="crpl_add_field",
                use_container_width=True
            ):

                if new_field.strip():

                    new_field = new_field.strip()


                    if new_field not in df.columns:

                        df[new_field] = ""


                        # Save new field in session
                        st.session_state[
                            "crpl_added_fields"
                        ] = df.columns.tolist()


                        st.success(
                            f"✅ '{new_field}' field added!"
                        )

                        st.rerun()

                    else:

                        st.warning(
                            "⚠️ This field already exists."
                        )

                else:

                    st.warning(
                        "⚠️ Please enter a field name."
                    )


        # =================================================
        # GET CURRENT DATA
        # =================================================

        edit_df = df.copy()


        # =================================================
        # APPLY SEARCH FILTER
        # =================================================

        if search:

            search_mask = pd.Series(
                False,
                index=edit_df.index
            )


            for col in edit_df.columns:

                search_mask = (
                    search_mask
                    |
                    edit_df[col]
                    .astype(str)
                    .str.contains(
                        search,
                        case=False,
                        na=False
                    )
                )


            edit_df = edit_df[
                search_mask
            ]


        # =================================================
        # APPLY STAGE FILTER
        # =================================================

        if stage_filter != "All":

            if stage_filter == "LIVE":

                edit_df = edit_df[
                    edit_df["Stage"]
                    .astype(str)
                    .str.upper()
                    .isin([
                        "LIVE",
                        "PRODUCTION",
                        "LIVE / PRODUCTION",
                        "LIVE/PRODUCTION"
                    ])
                ]

            else:

                edit_df = edit_df[
                    edit_df["Stage"]
                    .astype(str)
                    .str.upper()
                    == stage_filter
                ]


        # =================================================
        # APPLY KPI CARD FILTER
        # =================================================

        selected_card = st.session_state.get(
            "crpl_stage",
            "All"
        )


        if selected_card != "All":

            if selected_card == "LIVE":

                edit_df = edit_df[
                    edit_df["Stage"]
                    .astype(str)
                    .str.upper()
                    .isin([
                        "LIVE",
                        "PRODUCTION",
                        "LIVE / PRODUCTION",
                        "LIVE/PRODUCTION"
                    ])
                ]

            else:

                edit_df = edit_df[
                    edit_df["Stage"]
                    .astype(str)
                    .str.upper()
                    == selected_card
                ]


        # =================================================
        # EDITABLE TABLE
        # =================================================

        edited_df = st.data_editor(
            edit_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=500,
            key="crpl_editor"
        )


        # =================================================
        # SAVE / CLOSE
        # =================================================

        save_col, close_col = st.columns(2)


        # =================================================
        # SAVE CHANGES
        # =================================================

        with save_col:

            if st.button(
                "💾 Save Changes",
                type="primary",
                key="crpl_save",
                use_container_width=True
            ):

                try:

                    # -----------------------------------------
                    # READ ORIGINAL EXCEL
                    # -----------------------------------------

                    original = pd.read_excel(
                        file,
                        sheet_name="CRPL",
                        header=header_row
                    )


                    original.columns = [
                        str(x).strip()
                        for x in original.columns
                    ]


                    original = original.fillna("")


                    # -----------------------------------------
                    # EXISTING COLUMNS
                    # -----------------------------------------

                    for col in edited_df.columns:

                        if col not in original.columns:

                            original[col] = ""


                    # -----------------------------------------
                    # UPDATE EXISTING / FILTERED ROWS
                    # -----------------------------------------

                    for idx in edited_df.index:

                        if idx < len(original):

                            for col in edited_df.columns:

                                original.loc[
                                    idx,
                                    col
                                ] = edited_df.loc[
                                    idx,
                                    col
                                ]


                    # -----------------------------------------
                    # HANDLE NEW ROWS
                    # -----------------------------------------

                    original_indexes = set(
                        original.index
                    )


                    for idx in edited_df.index:

                        if idx not in original_indexes:

                            new_row = {}

                            for col in edited_df.columns:

                                new_row[col] = edited_df.loc[
                                    idx,
                                    col
                                ]


                            original = pd.concat(
                                [
                                    original,
                                    pd.DataFrame([new_row])
                                ],
                                ignore_index=True
                            )


                    # -----------------------------------------
                    # SAVE TO EXCEL
                    # -----------------------------------------

                    with pd.ExcelWriter(
                        file,
                        engine="openpyxl",
                        mode="a",
                        if_sheet_exists="replace"
                    ) as writer:

                        original.to_excel(
                            writer,
                            sheet_name="CRPL",
                            index=False
                        )


                    # -----------------------------------------
                    # SUCCESS
                    # -----------------------------------------

                    st.success(
                        "✅ CRPL changes, new rows and new fields "
                        "saved successfully!"
                    )


                    st.session_state[
                        "crpl_edit_mode"
                    ] = False


                    if "crpl_editor" in st.session_state:

                        del st.session_state[
                            "crpl_editor"
                        ]


                    st.rerun()


                except Exception as e:

                    st.error(
                        f"Save error: {e}"
                    )


        # =================================================
        # CLOSE EDITOR
        # =================================================

        with close_col:

            if st.button(
                "❌ Close Editor",
                key="crpl_close_editor",
                use_container_width=True
            ):

                st.session_state[
                    "crpl_edit_mode"
                ] = False


                if "crpl_editor" in st.session_state:

                    del st.session_state[
                        "crpl_editor"
                    ]


                st.rerun()
# =====================================================
# PAYSYS
# =====================================================

elif page == "PAYSYS":

    import os
    import pandas as pd

    # =====================================================
    # LOAD PAYSYS FILE
    # =====================================================

    file = os.path.join(
        os.path.dirname(__file__),
        "data",
        "PAYSYS_Weekly_Project_Update_04-Sep-2026.xlsx"
    )

    try:

        excel_file = pd.ExcelFile(file)

        sheet_name = excel_file.sheet_names[0]

        raw = pd.read_excel(
            file,
            sheet_name=sheet_name,
            header=None
        )

        header_row = None

        for i in range(min(20, len(raw))):

            row = " ".join(
                str(x).lower()
                for x in raw.iloc[i]
                if pd.notna(x)
            )

            if (
                "uat/live" in row
                and "paysys response" in row
            ):
                header_row = i
                break

        if header_row is None:

            st.error("PAYSYS header row not found.")
            st.dataframe(raw.head(20))
            st.stop()

        df = pd.read_excel(
            file,
            sheet_name=sheet_name,
            header=header_row
        )

        df.columns = [
            str(x).strip()
            for x in df.columns
        ]

        df = df.dropna(
            how="all"
        ).reset_index(drop=True)

    except Exception as e:

        st.error(
            f"PAYSYS file error: {e}"
        )

        st.stop()


    # =====================================================
    # REQUIRED COLUMNS
    # =====================================================

    columns = [
        "UAT/Live",
        "Current NBP Remarks",
        "PAYSYS Response",
        "NBP Remarks / Action Required",
        "Last Update Date"
    ]

    for col in columns:

        if col not in df.columns:
            df[col] = ""

    df = df[columns].fillna("")


    # =====================================================
    # HEADER
    # =====================================================

    st.html("""
    <div style="
        background:linear-gradient(135deg,#ffffff,#f8fbff);
        border-radius:18px;
        padding:14px 22px;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 22px rgba(0,0,0,.08);
        box-sizing:border-box;
        margin-bottom:2px;
        font-family:Segoe UI,Arial,sans-serif;">

        <div style="
            color:#006747;
            font-size:12px;
            font-weight:700;
            letter-spacing:1.8px;
            margin-bottom:4px;">
            PAYSYS MONITORING
        </div>

        <div style="
            color:#006747;
            font-size:28px;
            font-weight:800;
            line-height:1.15;">
            💳 PAYSYS
        </div>

        <div style="
            color:#6B7280;
            font-size:14px;
            margin-top:4px;">
            PAYSYS Issues &amp; Progress Tracking
        </div>

    </div>
    """)


    # =====================================================
    # FILTERS
    # =====================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:0;
        margin-bottom:3px;
        font-family:Segoe UI,Arial,sans-serif;">
        🎯 Filters
    </div>
    """)

    c1, c2 = st.columns(2)


    with c1:

        search = st.text_input(
            "Search PAYSYS",
            placeholder="🔍 Search PAYSYS remarks, response or status..."
        )


    with c2:

        stage_filter = st.selectbox(
            "Stage",
            [
                "All",
                "UAT",
                "LIVE"
            ]
        )


    # =====================================================
    # FILTER DATA
    # =====================================================

    filtered_df = df.copy()


    if search:

        search_columns = [
            "UAT/Live",
            "Current NBP Remarks",
            "PAYSYS Response",
            "NBP Remarks / Action Required",
            "Last Update Date"
        ]

        search_mask = pd.Series(
            False,
            index=filtered_df.index
        )

        for col in search_columns:

            search_mask = (
                search_mask
                |
                filtered_df[col]
                .astype(str)
                .str.contains(
                    search,
                    case=False,
                    na=False
                )
            )

        filtered_df = filtered_df[
            search_mask
        ]


    if stage_filter != "All":

        filtered_df = filtered_df[
            filtered_df["UAT/Live"]
            .astype(str)
            .str.strip()
            .str.upper()
            == stage_filter
        ]


   # =====================================================
    # =====================================================
    # KPI COUNTS
    # =====================================================

    stage = (
        df["UAT/Live"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    uat_count = (
        stage == "UAT"
    ).sum()

    live_count = (
        stage.isin([
            "LIVE",
            "PRODUCTION"
        ])
    ).sum()


    # =====================================================
    # VIP CLICKABLE KPI CARDS
    # =====================================================

    st.markdown("""
    <style>

    /* =====================================================
    KPI CARD CONTAINER
    ===================================================== */

    .st-key-paysys_kpis div[data-testid="stButton"] {
        width:100%;
    }


    /* =====================================================
    BASE CARD
    ===================================================== */

    .st-key-paysys_kpis div[data-testid="stButton"] button {

        width:100% !important;
        min-height:88px !important;

        background:#FFFFFF !important;

        border:1px solid #E5E7EB !important;
        border-radius:14px !important;

        padding:10px 6px !important;

        box-shadow:
            0 4px 12px rgba(0,0,0,.06) !important;

        color:#111827 !important;

        font-size:13px !important;
        font-weight:700 !important;

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        line-height:1.25 !important;

        white-space:pre-line !important;

        text-align:center !important;

        transition:all .2s ease !important;
    }


    /* =====================================================
    HOVER
    ===================================================== */

    .st-key-paysys_kpis div[data-testid="stButton"] button:hover {

        background:#F8FAFC !important;

        color:#111827 !important;

        transform:translateY(-2px);

        box-shadow:
            0 8px 18px rgba(0,0,0,.10) !important;
    }


    .st-key-paysys_kpis
    div[data-testid="stButton"] button:hover p {

        color:#111827 !important;
    }


    /* =====================================================
    SELECTED / FOCUSED CARD
    ===================================================== */

    .st-key-paysys_kpis
    div[data-testid="stButton"] button:focus {

        outline:none !important;

        background:#EAF5F0 !important;

        border:2px solid #006747 !important;

        box-shadow:
            0 0 0 3px rgba(0,103,71,0.12),
            0 8px 18px rgba(0,103,71,0.15) !important;

        color:#006747 !important;
    }


    .st-key-paysys_kpis
    div[data-testid="stButton"] button:focus p {

        color:#006747 !important;

        font-weight:700 !important;
    }


    /* =====================================================
    ALL
    ===================================================== */

    .st-key-paysys_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(1)
    div[data-testid="stButton"] button {

        border-top:5px solid #006747 !important;
    }


    /* =====================================================
    UAT
    ===================================================== */

    .st-key-paysys_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(2)
    div[data-testid="stButton"] button {

        border-top:5px solid #F9A825 !important;
    }


    /* =====================================================
    LIVE
    ===================================================== */

    .st-key-paysys_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(3)
    div[data-testid="stButton"] button {

        border-top:5px solid #00C853 !important;
    }


    /* =====================================================
    BUTTON TEXT
    ===================================================== */

    .st-key-paysys_kpis
    div[data-testid="stButton"] button p {

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        font-size:13px !important;

        font-weight:700 !important;

        line-height:1.3 !important;

        color:#111827 !important;
    }


    /* =====================================================
    COLUMN SPACING
    ===================================================== */

    .st-key-paysys_kpis
    div[data-testid="stHorizontalBlock"] {

        gap:8px !important;
    }


    /* =====================================================
    REMOVE EXTRA GAPS
    ===================================================== */

    .st-key-paysys_kpis
    div[data-testid="stVerticalBlock"] {

        gap:0 !important;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # KPI CARD CONTAINER
    # =====================================================

    with st.container(key="paysys_kpis"):

        c1, c2, c3 = st.columns(3)


        # =================================================
        # ALL
        # =================================================

        with c1:

            if st.button(
                f"ALL\n{len(df)}",
                key="paysys_all",
                use_container_width=True
            ):

                st.session_state["paysys_stage"] = "All"

                st.rerun()


        # =================================================
        # UAT
        # =================================================

        with c2:

            if st.button(
                f"UAT\n{uat_count}",
                key="paysys_uat",
                use_container_width=True
            ):

                st.session_state["paysys_stage"] = "UAT"

                st.rerun()


        # =================================================
        # LIVE
        # =================================================

        with c3:

            if st.button(
                f"LIVE\n{live_count}",
                key="paysys_live",
                use_container_width=True
            ):

                st.session_state["paysys_stage"] = "LIVE"

                st.rerun()


    st.markdown(
        "<div style='height:2px;'></div>",
        unsafe_allow_html=True
    )


    # =====================================================
    # CARD FILTER
    # =====================================================

    selected_card = st.session_state.get(
        "paysys_stage",
        "All"
    )


    if selected_card != "All":

        if selected_card == "LIVE":

            filtered_df = filtered_df[
                filtered_df["UAT/Live"]
                .astype(str)
                .str.strip()
                .str.upper()
                .isin([
                    "LIVE",
                    "PRODUCTION"
                ])
            ]

        else:

            filtered_df = filtered_df[
                filtered_df["UAT/Live"]
                .astype(str)
                .str.strip()
                .str.upper()
                == selected_card
            ]


    # =====================================================
    # TABLE TITLE
    # =====================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:2px;
        margin-bottom:3px;
        font-family:Segoe UI,Arial,sans-serif;">
        📋 PAYSYS Details
    </div>
    """)


    # =====================================================
    # VIP TABLE CSS - GREEN THEME
    # =====================================================

    st.markdown("""
    <style>

    .paysys-table-wrapper {
        background:#FFFFFF;
        border:1px solid #DDE5E1;
        border-radius:16px;
        padding:6px;
        box-shadow:0 6px 20px rgba(0,103,71,0.08);
        overflow:hidden;
    }

    .paysys-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:14px;
        table-layout:fixed;
    }

    .paysys-table thead th {
        background:#006747;
        color:white;
        font-weight:700;
        padding:14px 12px;
        text-align:left;
    }

    .paysys-table thead th:first-child {
        border-top-left-radius:11px;
    }

    .paysys-table thead th:last-child {
        border-top-right-radius:11px;
    }

    .paysys-table tbody td {
        padding:13px 12px;
        color:#1F2937;
        border-bottom:1px solid #E5E7EB;
        background:#FFFFFF;

        white-space:normal !important;
        word-wrap:break-word !important;
        overflow-wrap:anywhere !important;
        vertical-align:top;
        line-height:1.6;
    }

    .paysys-table tbody tr:nth-child(even) td {
        background:#F8FAFC;
    }

    .paysys-table tbody tr:hover td {
        background:#ECFDF5;
    }

    .paysys-name {
        font-weight:700;
        color:#006747 !important;
    }

    .stage-badge {
        display:inline-block;
        padding:5px 11px;
        border-radius:20px;
        font-size:11px;
        font-weight:800;
        white-space:nowrap;
    }

    .stage-uat {
        background:#FEF3C7;
        color:#92400E;
    }

    .stage-live {
        background:#DCFCE7;
        color:#166534;
    }

    .stage-default {
        background:#F3F4F6;
        color:#374151;
    }


    /* =====================================================
    PAYSYS KPI BUTTONS
    ===================================================== */

    div.stButton > button {
        background-color:#006747 !important;
        color:white !important;
        border:1px solid #006747 !important;
        border-radius:10px !important;
        font-weight:700 !important;
    }

    div.stButton > button:hover {
        background-color:#00553A !important;
        color:white !important;
        border-color:#00553A !important;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # TABLE DATA
    # =====================================================

    display_df = filtered_df.copy()


    # =====================================================
    # STAGE BADGE
    # =====================================================

    def stage_badge(value):

        value = str(value).strip()

        upper = value.upper()


        if upper == "UAT":

            css = "stage-uat"


        elif upper in [
            "LIVE",
            "PRODUCTION"
        ]:

            css = "stage-live"


        else:

            css = "stage-default"


        return (
            f'<span class="stage-badge {css}">'
            f'{value}'
            f'</span>'
        )


    display_df["UAT/Live"] = (
        display_df["UAT/Live"]
        .apply(stage_badge)
    )


    # =====================================================
    # HTML TABLE
    # =====================================================

    table_html = display_df.to_html(
        index=False,
        escape=False,
        classes="paysys-table"
    )


    st.markdown(
        f"""
        <div class="paysys-table-wrapper">
            {table_html}
        </div>
        """,
        unsafe_allow_html=True
    )

    # =====================================================
    # EDIT PAYSYS BUTTON
    # =====================================================

    st.markdown(
        "<br>",
        unsafe_allow_html=True
    )


    if "paysys_edit_mode" not in st.session_state:

        st.session_state[
            "paysys_edit_mode"
        ] = False


    # =====================================================
    # EDIT BUTTON
    # =====================================================

    if not st.session_state["paysys_edit_mode"]:

        if st.button(
            "✏️ Edit PAYSYS",
            key="paysys_edit_button",
            type="secondary",
            use_container_width=True
        ):

            st.session_state[
                "paysys_edit_mode"
            ] = True

            st.rerun()


    # =====================================================
    # EDITOR
    # =====================================================

    if st.session_state["paysys_edit_mode"]:

        st.markdown("""
        <h2 style="
        color:#006747;
        font-size:28px;
        font-weight:700;
        margin-top:20px;
        margin-bottom:15px;">
        ✏️ Edit PAYSYS
        </h2>
        """, unsafe_allow_html=True)


        st.info(
            "✏️ Edit existing records, add new records, "
            "or create a new field/column."
        )


        # =================================================
        # ADD NEW FIELD
        # =================================================

        st.markdown("""
        <h3 style="
        color:#006747;
        font-size:20px;
        font-weight:700;
        margin-top:10px;">
        ➕ Add New Field
        </h3>
        """, unsafe_allow_html=True)


        field_col1, field_col2 = st.columns([3, 1])


        with field_col1:

            new_field = st.text_input(
                "New Field Name",
                placeholder="e.g. Vendor Remarks",
                key="paysys_new_field_input"
            )


        with field_col2:

            st.markdown(
                "<br>",
                unsafe_allow_html=True
            )

            if st.button(
                "➕ Add Field",
                key="paysys_add_field",
                use_container_width=True
            ):

                if new_field.strip():

                    new_field = new_field.strip()


                    if new_field not in df.columns:

                        df[new_field] = ""


                        st.session_state[
                            "paysys_added_fields"
                        ] = df.columns.tolist()


                        st.success(
                            f"✅ '{new_field}' field added!"
                        )

                        st.rerun()

                    else:

                        st.warning(
                            "⚠️ This field already exists."
                        )

                else:

                    st.warning(
                        "⚠️ Please enter a field name."
                    )


        # =================================================
        # GET CURRENT DATA
        # =================================================

        edit_df = df.copy()


        # =================================================
        # APPLY SEARCH FILTER
        # =================================================

        if search:

            search_mask = pd.Series(
                False,
                index=edit_df.index
            )


            for col in edit_df.columns:

                search_mask = (
                    search_mask
                    |
                    edit_df[col]
                    .astype(str)
                    .str.contains(
                        search,
                        case=False,
                        na=False
                    )
                )


            edit_df = edit_df[
                search_mask
            ]


        # =================================================
        # APPLY STAGE FILTER
        # =================================================

        if stage_filter != "All":

            if stage_filter == "LIVE":

                edit_df = edit_df[
                    edit_df["UAT/Live"]
                    .astype(str)
                    .str.upper()
                    .isin([
                        "LIVE",
                        "PRODUCTION"
                    ])
                ]

            else:

                edit_df = edit_df[
                    edit_df["UAT/Live"]
                    .astype(str)
                    .str.upper()
                    == stage_filter
                ]


        # =================================================
        # APPLY KPI CARD FILTER
        # =================================================

        selected_card = st.session_state.get(
            "paysys_stage",
            "All"
        )


        if selected_card != "All":

            if selected_card == "LIVE":

                edit_df = edit_df[
                    edit_df["UAT/Live"]
                    .astype(str)
                    .str.upper()
                    .isin([
                        "LIVE",
                        "PRODUCTION"
                    ])
                ]

            else:

                edit_df = edit_df[
                    edit_df["UAT/Live"]
                    .astype(str)
                    .str.upper()
                    == selected_card
                ]


        # =================================================
        # EDITABLE TABLE
        # =================================================

        edited_df = st.data_editor(
            edit_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=500,
            disabled=[],
            key="paysys_editor"
        )


        # =================================================
        # SAVE / CLOSE BUTTONS
        # =================================================

        save_col, close_col = st.columns(2)


        # =================================================
        # SAVE CHANGES
        # =================================================

        with save_col:

            if st.button(
                "💾 Save Changes",
                type="primary",
                key="paysys_save",
                use_container_width=True
            ):

                try:

                    # -----------------------------------------
                    # READ ORIGINAL EXCEL
                    # -----------------------------------------

                    original = pd.read_excel(
                        file,
                        sheet_name=sheet_name,
                        header=header_row
                    )


                    original.columns = [
                        str(x).strip()
                        for x in original.columns
                    ]


                    original = original.fillna("")


                    # -----------------------------------------
                    # ADD NEW COLUMNS
                    # -----------------------------------------

                    for col in edited_df.columns:

                        if col not in original.columns:

                            original[col] = ""


                    # -----------------------------------------
                    # UPDATE EXISTING ROWS
                    # -----------------------------------------

                    for idx in edited_df.index:

                        if idx < len(original):

                            for col in edited_df.columns:

                                original.loc[
                                    idx,
                                    col
                                ] = edited_df.loc[
                                    idx,
                                    col
                                ]


                    # -----------------------------------------
                    # SAVE EXCEL
                    # -----------------------------------------

                    with pd.ExcelWriter(
                        file,
                        engine="openpyxl",
                        mode="a",
                        if_sheet_exists="replace"
                    ) as writer:

                        original.to_excel(
                            writer,
                            sheet_name=sheet_name,
                            index=False
                        )


                    # -----------------------------------------
                    # SUCCESS
                    # -----------------------------------------

                    st.success(
                        "✅ PAYSYS changes, new rows and "
                        "new fields saved successfully!"
                    )


                    st.session_state[
                        "paysys_edit_mode"
                    ] = False


                    if "paysys_editor" in st.session_state:

                        del st.session_state[
                            "paysys_editor"
                        ]


                    st.rerun()


                except Exception as e:

                    st.error(
                        f"Save error: {e}"
                    )


        # =================================================
        # CLOSE EDITOR
        # =================================================

        with close_col:

            if st.button(
                "❌ Close Editor",
                key="paysys_close_editor",
                use_container_width=True
            ):

                st.session_state[
                    "paysys_edit_mode"
                ] = False


                if "paysys_editor" in st.session_state:

                    del st.session_state[
                        "paysys_editor"
                    ]


                st.rerun()
# =====================================================
# EXPORT
# =====================================================

elif page == "Export":

    # ==========================================
    # HEADER
    # ==========================================

    st.html("""
    <div style="
        background:linear-gradient(135deg,#ffffff,#f4fbf8);
        border-radius:18px;
        padding:14px 22px;
        border:1px solid #DDEBE5;
        box-shadow:0 8px 22px rgba(0,0,0,.08);
        margin-bottom:8px;
        font-family:Segoe UI,Arial,sans-serif;">

        <div style="
            color:#006747;
            font-size:12px;
            font-weight:700;
            letter-spacing:1.8px;
            margin-bottom:4px;">
            SMARTPAY PROJECT MANAGEMENT
        </div>

        <div style="
            color:#006747;
            font-size:28px;
            font-weight:800;
            line-height:1.15;">
            Executive Report Center
        </div>

        <div style="
            color:#6B7280;
            font-size:14px;
            margin-top:4px;">
            Generate, preview and download SmartPay project reports.
        </div>

    </div>
    """)


    # ==========================================
    # REPORT INFORMATION
    # ==========================================

    st.html("""
    <div style="
        background:#FFFFFF;
        border-radius:14px;
        padding:12px 16px;
        border:1px solid #E5E7EB;
        box-shadow:0 4px 12px rgba(0,0,0,.06);
        margin-bottom:8px;
        font-family:Segoe UI,Arial,sans-serif;">

        <div style="
            color:#006747;
            font-size:18px;
            font-weight:700;
            margin-bottom:8px;">
            📋 Report Includes
        </div>

        <div style="
            color:#374151;
            font-size:13px;
            line-height:1.7;">

            ✅ Dashboard Summary<br>
            ✅ KPI Overview<br>
            ✅ Project Details<br>
            ✅ Team Performance<br>
            ✅ Analytics Summary<br>
            ✅ Project Timeline

        </div>

    </div>
    """)


    # ==========================================
    # GENERATED DATE
    # ==========================================

    generated_time = datetime.now().strftime(
        "%d %B %Y  |  %I:%M %p"
    )

    st.html(f"""
    <div style="
        background:#F0FDF4;
        border:1px solid #BBF7D0;
        border-radius:12px;
        padding:9px 14px;
        margin-bottom:8px;
        font-family:Segoe UI,Arial,sans-serif;">

        <div style="
            color:#166534;
            font-size:10px;
            font-weight:700;
            text-transform:uppercase;
            letter-spacing:1px;">
            Generated On
        </div>

        <div style="
            color:#166534;
            font-size:15px;
            font-weight:700;
            margin-top:3px;">
            {generated_time}
        </div>

    </div>
    """)


    # ==========================================
    # EXPORT OPTIONS
    # ==========================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:2px;
        margin-bottom:5px;
        font-family:Segoe UI,Arial,sans-serif;">
        📤 Export Reports
    </div>
    """)


    # ==========================================
    # LOAD ALL DATA
    # ==========================================

    # ------------------------------------------
    # SMARTPAY PROJECT DATA
    # ------------------------------------------

    projects_export_df = df.copy()


    # ------------------------------------------
    # CRPL DATA
    # ------------------------------------------

    crpl_file = os.path.join(
        os.path.dirname(__file__),
        "data",
        "CRPL_All_Data_Exactly_20.xlsx"
    )

    try:

        crpl_export_df = pd.read_excel(
            crpl_file,
            sheet_name="CRPL"
        )

    except Exception:

        crpl_export_df = pd.DataFrame()

        # ------------------------------------------
    # PAYSYS DATA
    # ------------------------------------------

    paysys_file = os.path.join(
        os.path.dirname(__file__),
        "data",
        "PAYSYS_Weekly_Project_Update_04-Sep-2026.xlsx"
    )

    try:

        paysys_export_df = pd.read_excel(
            paysys_file
        )

    except Exception:

        paysys_export_df = pd.DataFrame()


    # ==========================================
    # DATA SUMMARY
    # ==========================================

    s1, s2, s3 = st.columns(3)


    with s1:

        st.metric(
            "SmartPay Projects",
            len(projects_export_df)
        )


    with s2:

        st.metric(
            "CRPL Records",
            len(crpl_export_df)
        )


    with s3:

        st.metric(
            "PAYSYS Records",
            len(paysys_export_df)
        )


    st.markdown(
        "<div style='height:4px;'></div>",
        unsafe_allow_html=True
    )


    # ==========================================
    # CREATE CSV FILES
    # ==========================================

    smartpay_csv = projects_export_df.to_csv(
        index=False
    ).encode("utf-8-sig")


    crpl_csv = crpl_export_df.to_csv(
        index=False
    ).encode("utf-8-sig")


    paysys_csv = paysys_export_df.to_csv(
        index=False
    ).encode("utf-8-sig")


    # ==========================================
    # CREATE EXCEL FILE
    # ==========================================

    excel_buffer = BytesIO()


    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:


        # --------------------------------------
        # SMARTPAY SHEET
        # --------------------------------------

        projects_export_df.to_excel(
            writer,
            index=False,
            sheet_name="SmartPay Projects"
        )


        # --------------------------------------
        # CRPL SHEET
        # --------------------------------------

        if not crpl_export_df.empty:

            crpl_export_df.to_excel(
                writer,
                index=False,
                sheet_name="CRPL"
            )


        # --------------------------------------
        # PAYSYS SHEET
        # --------------------------------------

        if not paysys_export_df.empty:

            paysys_export_df.to_excel(
                writer,
                index=False,
                sheet_name="PAYSYS"
            )


        # ======================================
        # FORMAT EVERY SHEET
        # ======================================

        for sheet_name, worksheet in writer.sheets.items():


            # ----------------------------------
            # FREEZE HEADER
            # ----------------------------------

            worksheet.freeze_panes = "A2"


            # ----------------------------------
            # AUTO COLUMN WIDTH
            # ----------------------------------

            for column in worksheet.columns:

                max_length = 0

                column_letter = (
                    column[0].column_letter
                )

                for cell in column:

                    try:

                        if cell.value is not None:

                            max_length = max(
                                max_length,
                                len(str(cell.value))
                            )

                    except Exception:

                        pass


                worksheet.column_dimensions[
                    column_letter
                ].width = min(
                    max_length + 3,
                    45
                )

            # ----------------------------------
    # EXCEL TABLE
    # ----------------------------------

    last_row = worksheet.max_row
    last_col = worksheet.max_column


    if last_row > 1 and last_col > 0:

        last_col_letter = get_column_letter(
            last_col
        )


        table_ref = (
            f"A1:{last_col_letter}{last_row}"
        )


        safe_name = (
            sheet_name
            .replace(" ", "")
            .replace("-", "")
            .replace("/", "")
        )


        table = Table(
            displayName=f"{safe_name}Table",
            ref=table_ref
        )


        style = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )


        table.tableStyleInfo = style

        worksheet.add_table(table)


    excel_data = excel_buffer.getvalue()


    # ==========================================
    # EXCLUSIVE EXECUTIVE PDF
    # ==========================================

    pdf_buffer = BytesIO()

    pdf_doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=20,
        bottomMargin=20
    )


    pdf_styles = getSampleStyleSheet()


    # ==========================================
    # PDF TITLE STYLE
    # ==========================================

    pdf_title_style = ParagraphStyle(
        "PDFTitle",
        parent=pdf_styles["Title"],
        fontSize=22,
        leading=26,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#006747"),
        spaceAfter=8
    )


    # ==========================================
    # PDF HEADING STYLE
    # ==========================================

    pdf_heading_style = ParagraphStyle(
        "PDFHeading",
        parent=pdf_styles["Heading2"],
        fontSize=16,
        leading=19,
        textColor=colors.HexColor("#006747"),
        spaceBefore=8,
        spaceAfter=8
    )


    # ==========================================
    # PDF CELL STYLE
    # ==========================================

    pdf_cell_style = ParagraphStyle(
        "PDFCell",
        parent=pdf_styles["Normal"],
        fontSize=7,
        leading=8
    )


    # ==========================================
    # PDF HEADER STYLE
    # ==========================================

    pdf_header_style = ParagraphStyle(
        "PDFHeader",
        parent=pdf_styles["Normal"],
        fontSize=7,
        leading=8,
        textColor=colors.white,
        fontName="Helvetica-Bold"
    )


    pdf_story = []


    # ==========================================
    # PDF COVER
    # ==========================================

    pdf_story.append(
        Spacer(1, 0.25 * inch)
    )


    pdf_story.append(
        Paragraph(
            "SMARTPAY PROJECT MANAGEMENT",
            pdf_title_style
        )
    )


    pdf_story.append(
        Paragraph(
            "Executive Report",
            ParagraphStyle(
                "PDFSubTitle",
                parent=pdf_styles["Heading2"],
                alignment=TA_CENTER,
                fontSize=15,
                leading=18,
                textColor=colors.HexColor("#6B7280"),
                spaceAfter=5
            )
        )
    )


    pdf_story.append(
        Spacer(1, 0.12 * inch)
    )


    pdf_story.append(
        Paragraph(
            f"Generated On: "
            f"{datetime.now().strftime('%d %B %Y | %I:%M %p')}",
            ParagraphStyle(
                "PDFDate",
                parent=pdf_styles["Normal"],
                alignment=TA_CENTER,
                fontSize=9,
                textColor=colors.HexColor("#555555")
            )
        )
    )


    pdf_story.append(
        Spacer(1, 0.25 * inch)
    )


    # ==========================================
    # EXECUTIVE SUMMARY
    # ==========================================

    summary_data = [
        ["REPORT", "TOTAL RECORDS"],
        [
            "SmartPay Projects",
            str(len(projects_export_df))
        ],
        [
            "CRPL",
            str(len(crpl_export_df))
        ],
        [
            "PAYSYS",
            str(len(paysys_export_df))
        ]
    ]


    summary_table = PDFTable(
        summary_data,
        colWidths=[
            5 * inch,
            2 * inch
        ]
    )


    summary_table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#006747")
            ),
            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),
            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),
            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                8
            ),
            (
                "ALIGN",
                (1, 0),
                (1, -1),
                "CENTER"
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor("#DDEBE5")
            ),
            (
                "BACKGROUND",
                (0, 1),
                (-1, -1),
                colors.HexColor("#F4FBF8")
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                6
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                6
            )
        ])
    )


    pdf_story.append(
        summary_table
    )


    pdf_story.append(
        PageBreak()
    )

    # ==========================================
    # DATAFRAME → PDF FUNCTION
    # ==========================================

    def dataframe_to_pdf(dataframe, title):

        if dataframe.empty:
            return

        pdf_story.append(
            Paragraph(
                title,
                pdf_heading_style
            )
        )

        # Keep PDF readable
        export_pdf_df = dataframe.copy()

        # Maximum 100 records per section
        export_pdf_df = export_pdf_df.head(100)

        columns = list(export_pdf_df.columns)

        table_data = []

        # Header
        table_data.append([
            Paragraph(
                str(column),
                pdf_header_style
            )
            for column in columns
        ])

        # Rows
        for _, row in export_pdf_df.iterrows():

            row_data = []

            for value in row:

                if pd.isna(value):
                    value = ""

                value = str(value)

                # Avoid huge cells
                if len(value) > 180:
                    value = value[:180] + "..."

                row_data.append(
                    Paragraph(
                        value,
                        pdf_cell_style
                    )
                )

            table_data.append(row_data)

        # Dynamic column width
        available_width = 10.5 * inch

        column_count = max(
            len(columns),
            1
        )

        column_width = (
            available_width /
            column_count
        )

        pdf_table = PDFTable(
            table_data,
            repeatRows=1,
            colWidths=[
                column_width
            ] * column_count
        )

        pdf_table.setStyle(
            TableStyle([
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#006747")
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.3,
                    colors.HexColor("#D1D5DB")
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP"
                ),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [
                        colors.white,
                        colors.HexColor("#F7FBF9")
                    ]
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                )
            ])
        )

        pdf_story.append(pdf_table)

        pdf_story.append(
            Spacer(1, 0.15 * inch)
        )

        pdf_story.append(PageBreak())


    # ==========================================
    # ADD DATA TO PDF
    # ==========================================

    dataframe_to_pdf(
        projects_export_df,
        "SmartPay Projects"
    )

    dataframe_to_pdf(
        crpl_export_df,
        "CRPL"
    )

    dataframe_to_pdf(
        paysys_export_df,
        "PAYSYS"
    )


    # ==========================================
    # BUILD PDF
    # ==========================================

    pdf_doc.build(pdf_story)

    pdf_buffer.seek(0)

    executive_pdf = pdf_buffer.getvalue()


    # ==========================================
    # EXPORT CARDS
    # ==========================================

    e1, e2, e3 = st.columns(3)


    # ==========================================
    # SMARTPAY EXPORT
    # ==========================================

    with e1:

        st.html("""
        <div style="
            background:#FFFFFF;
            border-radius:14px;
            padding:12px 14px;
            border:1px solid #DDEBE5;
            border-top:4px solid #006747;
            box-shadow:0 4px 12px rgba(0,0,0,.06);
            font-family:Segoe UI,Arial,sans-serif;">

            <div style="
                color:#006747;
                font-size:16px;
                font-weight:700;">
                SmartPay Projects
            </div>

            <div style="
                color:#6B7280;
                font-size:12px;
                margin-top:4px;">
                Download SmartPay project data
            </div>

        </div>
        """)


        st.download_button(
            "⬇ Download SmartPay CSV",
            smartpay_csv,
            "SmartPay_Projects.csv",
            "text/csv",
            use_container_width=True
        )
        # ==========================================
    # EXECUTIVE PDF EXPORT
    # ==========================================

    st.markdown(
        "<div style='height:2px;'></div>",
        unsafe_allow_html=True
    )

    st.html("""
    <div style="
        background:linear-gradient(135deg,#F0FDF4,#FFFFFF);
        border-radius:14px;
        padding:12px 16px;
        border:1px solid #BBF7D0;
        border-left:4px solid #006747;
        box-shadow:0 4px 12px rgba(0,0,0,.06);
        font-family:Segoe UI,Arial,sans-serif;">

        <div style="
            color:#006747;
            font-size:18px;
            font-weight:700;">
            📕 Exclusive Executive PDF Report
        </div>

        <div style="
            color:#6B7280;
            font-size:12px;
            margin-top:4px;">
            Complete executive PDF containing SmartPay,
            CRPL and PAYSYS project information.
        </div>

    </div>
    """)

    st.download_button(
        "⬇ Download Exclusive Executive PDF",
        executive_pdf,
        "SmartPay_Exclusive_Executive_Report.pdf",
        "application/pdf",
        use_container_width=True
    )


    # ==========================================
    # CRPL EXPORT
    # ==========================================

    with e2:

        st.html("""
        <div style="
            background:#FFFFFF;
            border-radius:14px;
            padding:12px 14px;
            border:1px solid #DDEBE5;
            border-top:4px solid #FF9800;
            box-shadow:0 4px 12px rgba(0,0,0,.06);
            font-family:Segoe UI,Arial,sans-serif;">

            <div style="
                color:#E65100;
                font-size:16px;
                font-weight:700;">
                CRPL
            </div>

            <div style="
                color:#6B7280;
                font-size:12px;
                margin-top:4px;">
                Download CRPL vendor data
            </div>

        </div>
        """)

        st.download_button(
            "⬇ Download CRPL CSV",
            crpl_csv,
            "CRPL_Data.csv",
            "text/csv",
            use_container_width=True
        )


    # ==========================================
    # PAYSYS EXPORT
    # ==========================================

    with e3:

        st.html("""
        <div style="
            background:#FFFFFF;
            border-radius:14px;
            padding:12px 14px;
            border:1px solid #DDEBE5;
            border-top:4px solid #3949AB;
            box-shadow:0 4px 12px rgba(0,0,0,.06);
            font-family:Segoe UI,Arial,sans-serif;">

            <div style="
                color:#3949AB;
                font-size:16px;
                font-weight:700;">
                PAYSYS
            </div>

            <div style="
                color:#6B7280;
                font-size:12px;
                margin-top:4px;">
                Download PAYSYS vendor data
            </div>

        </div>
        """)

        st.download_button(
            "⬇ Download PAYSYS CSV",
            paysys_csv,
            "PAYSYS_Data.csv",
            "text/csv",
            use_container_width=True
        )


    # ==========================================
    # COMPLETE EXCEL EXPORT
    # ==========================================

    st.markdown(
        "<div style='height:4px;'></div>",
        unsafe_allow_html=True
    )

    st.html("""
    <div style="
        background:linear-gradient(135deg,#F0FDF4,#FFFFFF);
        border-radius:14px;
        padding:12px 16px;
        border:1px solid #BBF7D0;
        border-left:4px solid #006747;
        box-shadow:0 4px 12px rgba(0,0,0,.06);
        font-family:Segoe UI,Arial,sans-serif;">

        <div style="
            color:#006747;
            font-size:18px;
            font-weight:700;">
            📊 Complete Executive Excel Report
        </div>

        <div style="
            color:#6B7280;
            font-size:12px;
            margin-top:4px;">
            One Excel file containing SmartPay Projects, CRPL and PAYSYS
            in separate worksheets.
        </div>

    </div>
    """)

    st.download_button(
        "⬇ Download Complete Excel Report",
        excel_data,
        "SmartPay_Executive_Data.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


    st.markdown(
        "<div style='height:4px;'></div>",
        unsafe_allow_html=True
    )

    st.markdown("---")


    # ==========================================
    # REPORT PREVIEW
    # ==========================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:2px;
        margin-bottom:5px;
        font-family:Segoe UI,Arial,sans-serif;">
        👁️ Report Preview
    </div>
    """)


    # ==========================================
    # PREVIEW TABS
    # ==========================================

    preview_projects, preview_crpl, preview_paysys = st.tabs(
        [
            "SmartPay Projects",
            "CRPL",
            "PAYSYS"
        ]
    )


    # ==========================================
    # SMARTPAY PREVIEW
    # ==========================================

    with preview_projects:

        st.html("""
        <div style="
            color:#006747;
            font-size:18px;
            font-weight:700;
            margin-bottom:5px;
            font-family:Segoe UI,Arial,sans-serif;">
            SmartPay Project Data
        </div>
        """)

        st.dataframe(
            projects_export_df,
            use_container_width=True,
            hide_index=True,
            height=500
        )


    # ==========================================
    # CRPL PREVIEW
    # ==========================================

    with preview_crpl:

        st.html("""
        <div style="
            color:#E65100;
            font-size:18px;
            font-weight:700;
            margin-bottom:5px;
            font-family:Segoe UI,Arial,sans-serif;">
            CRPL Data
        </div>
        """)

        if not crpl_export_df.empty:

            st.dataframe(
                crpl_export_df,
                use_container_width=True,
                hide_index=True,
                height=500
            )

        else:

            st.warning(
                "CRPL data could not be loaded."
            )


    # ==========================================
    # PAYSYS PREVIEW
    # ==========================================

    with preview_paysys:

        st.html("""
        <div style="
            color:#3949AB;
            font-size:18px;
            font-weight:700;
            margin-bottom:5px;
            font-family:Segoe UI,Arial,sans-serif;">
            PAYSYS Data
        </div>
        """)

        if not paysys_export_df.empty:

            st.dataframe(
                paysys_export_df,
                use_container_width=True,
                hide_index=True,
                height=500
            )

        else:

            st.warning(
                "PAYSYS data could not be loaded."
            )


    # ==========================================
    # FINAL STATUS
    # ==========================================

    st.success(
        f"Report Ready — "
        f"{len(projects_export_df)} SmartPay Projects | "
        f"{len(crpl_export_df)} CRPL Records | "
        f"{len(paysys_export_df)} PAYSYS Records"
    )