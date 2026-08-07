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

st.sidebar.markdown(
    "<h4 style='color:white;'>Navigation</h4>",
    unsafe_allow_html=True
)

page = st.sidebar.radio(
    "Menu",
    [
        "Dashboard",
        "Projects",
        "Analytics",
        "Project Timeline",
        "Team Performance",
        "Voice Search",
        "Export"
    ],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")

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
    # HEADER - VIP ENTERPRISE
    # ==========================================

    city, temp, weather = get_weather()

    left, center, right = st.columns([5.5, 2.2, 2])

    # ==========================================
    # LEFT - TITLE
    # ==========================================

    with left:

        st.markdown("""
        <div style="
        background:linear-gradient(135deg,#ffffff,#f8fbff);
        border-radius:24px;
        padding:35px;
        border:1px solid #E5E7EB;
        box-shadow:0 14px 35px rgba(0,0,0,.10);
        min-height:230px;">

        <div style="
        color:#006747;
        font-size:18px;
        font-weight:700;
        letter-spacing:2px;">
        NATIONAL BANK OF PAKISTAN
        </div>

        <div style="
        font-size:52px;
        color:#006747;
        font-weight:800;
        line-height:1.15;
        margin-top:18px;">
        SmartPay Project Dashboard
        </div>

        <div style="
        margin-top:20px;
        font-size:20px;
        color:#374151;">
        Digital Banking Group
        </div>

        <div style="
        margin-top:28px;
        display:inline-block;
        background:#ECFDF5;
        color:#006747;
        padding:8px 18px;
        border-radius:30px;
        font-size:14px;
        font-weight:600;">
        ● LIVE Dashboard
        </div>

        </div>
        """, unsafe_allow_html=True)

    # ==========================================
    # CENTER - WEATHER
    # ==========================================

    with center:

        st.markdown(f"""
        <div style="
        background:linear-gradient(180deg,#ffffff,#f8fbff);
        border-radius:24px;
        padding:25px;
        min-height:230px;
        text-align:center;
        border:1px solid #E5E7EB;
        box-shadow:0 14px 35px rgba(0,0,0,.10);">

        <div style="
        font-size:32px;">
        🌤
        </div>

        <div style="
        font-size:20px;
        color:#006747;
        font-weight:700;
        margin-top:5px;">
        {city}
        </div>

        <div style="
        font-size:58px;
        color:#006747;
        font-weight:800;
        margin:18px 0 5px;">
        {temp}°
        </div>

        <div style="
        color:#6B7280;
        font-size:18px;">
        {weather}
        </div>

        </div>
        """, unsafe_allow_html=True)

    # ==========================================
    # RIGHT - DATE & TIME
    # ==========================================

    with right:

        components.html("""
        <div style="
        background:linear-gradient(180deg,#ffffff,#f8fbff);
        border-radius:24px;
        padding:25px;
        min-height:230px;
        text-align:center;
        border:1px solid #E5E7EB;
        box-shadow:0 14px 35px rgba(0,0,0,.10);
        font-family:Segoe UI;">

        <div style="
        color:#006747;
        font-size:18px;
        font-weight:700;">
        TODAY
        </div>

        <div id="date"
        style="
        margin-top:18px;
        font-size:34px;
        font-weight:700;
        color:#111827;">
        </div>

        <hr style="margin:22px 0;">

        <div id="clock"
        style="
        font-size:34px;
        font-weight:800;
        color:#006747;">
        </div>

        <div style="
        margin-top:15px;
        color:#6B7280;
        font-size:14px;">
        Pakistan Standard Time
        </div>

        </div>

        <script>

        function updateClock(){

            const now = new Date();

            document.getElementById("date").innerHTML =
            now.toLocaleDateString("en-GB",{
                day:"2-digit",
                month:"short",
                year:"numeric"
            });

            document.getElementById("clock").innerHTML =
            now.toLocaleTimeString("en-US");

        }

        updateClock();

        setInterval(updateClock,1000);

        </script>
        """, height=440, scrolling=False)
    # ==========================================
    # KPI - VIP ENTERPRISE CARDS
    # ==========================================

    status = df["Status"].astype(str).str.upper().str.strip()

    total_projects = len(df)
    scoping_projects = len(df[status == "UNDER SCOPING"])
    development_projects = len(df[status == "UNDER DEVELOPMENT"])
    uat_projects = len(df[status == "UAT"])
    review_projects = len(df[status == "IS REVIEW"])
    cmc_projects = len(df[status == "CMC"])
    live_projects = len(df[status == "LIVE"])

    # Better spacing for cards
    c1, c2, c3, c4, c5, c6, c7 = st.columns(
        [1.25, 1.15, 1.25, 1.0, 1.15, 1.0, 1.0]
    )

    ## ------------------------------------------
    # KPI Card Function
    # ------------------------------------------

    def vip_card(title, value, color):

        st.markdown(
            f"""
    <div style="
    background:white;
    border-radius:22px;
    padding:24px 16px;
    height:145px;
    border-top:8px solid {color};
    box-shadow:0 8px 22px rgba(0,0,0,.08);
    display:flex;
    flex-direction:column;
    justify-content:space-between;
    align-items:center;
    text-align:center;">

    <div style="
    color:#6B7280;
    font-size:15px;
    font-weight:600;">
    {title}
    </div>

    <div style="
    color:{color};
    font-size:48px;
    font-weight:700;
    line-height:1;">
    {value}
    </div>

    </div>
    """,
            unsafe_allow_html=True,
        )

    # ------------------------------------------
    # KPI Cards
    # ------------------------------------------

    with c1:
        vip_card("Total Projects", total_projects, "#006747")

    with c2:
        vip_card("Scoping", scoping_projects, "#8E24AA")

    with c3:
        vip_card("Development", development_projects, "#FF9800")

    with c4:
        vip_card("UAT", uat_projects, "#F9A825")

    with c5:
        vip_card("IS Review", review_projects, "#00ACC1")

    with c6:
        vip_card("CMC", cmc_projects, "#3949AB")

    with c7:
        vip_card("LIVE", live_projects, "#00C853")

    st.markdown("<br>", unsafe_allow_html=True)
    # =====================================================
    # TEAM OVERVIEW
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:34px;
    font-weight:700;
    margin-bottom:25px;">
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
    border-radius:24px;
    padding:25px;
    text-align:center;
    border:1px solid #E5E7EB;
    box-shadow:0 12px 30px rgba(0,0,0,.08);
    min-height:210px;">

    <div style="
    width:70px;
    height:70px;
    border-radius:50%;
    background:#E8F5E9;
    margin:auto;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:34px;">
    👤
    </div>

    <div style="
    margin-top:18px;
    font-size:22px;
    font-weight:700;
    color:#006747;">
    {row["Allocation"]}
    </div>

    <div style="
    margin-top:18px;
    font-size:52px;
    font-weight:800;
    color:#111827;">
    {row["Projects"]}
    </div>

    <div style="
    margin-top:8px;
    font-size:15px;
    color:#6B7280;">
    Projects Assigned
    </div>

    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)


    # =====================================================
    # TEAM WORKLOAD
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:34px;
    font-weight:700;
    margin-bottom:20px;">
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
        marker_line_width=0
    )

    fig.update_layout(
        height=500,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=20, b=20),
        coloraxis_showscale=False,
        xaxis_title="",
        yaxis_title="Projects",
        font=dict(size=15),
        xaxis=dict(showgrid=False),
        yaxis=dict(gridcolor="#ECECEC")
    )

    st.markdown("""
    <div style="
    background:white;
    border-radius:24px;
    padding:15px;
    border:1px solid #E5E7EB;
    box-shadow:0 12px 30px rgba(0,0,0,.08);">
    </div>
    """, unsafe_allow_html=True)

    st.plotly_chart(fig, width="stretch")


    # =====================================================
    # TEAM SUMMARY
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

    st.markdown("""
    <div style="
    background:white;
    border-radius:24px;
    padding:15px;
    border:1px solid #E5E7EB;
    box-shadow:0 12px 30px rgba(0,0,0,.08);">
    </div>
    """, unsafe_allow_html=True)

    st.dataframe(
        summary,
        width="stretch",
        hide_index=True
    )


    # =====================================================
    # SMART SEARCH
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:34px;
    font-weight:700;
    margin-top:30px;
    margin-bottom:20px;">
    🔍 Smart Search
    </h2>
    """, unsafe_allow_html=True)

    project = st.text_input(
        "",
        placeholder="🔍 Search Project...",
        label_visibility="collapsed"
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

            st.success(f"✅ {len(result)} Project(s) Found")

            st.dataframe(
                result,
                width="stretch",
                hide_index=True
            )

    st.markdown("<br>", unsafe_allow_html=True)
# =====================================================
# PROJECTS
# =====================================================

elif page == "Projects":

    # =====================================================
    # HEADER
    # =====================================================

    st.markdown("""
    <div style="
    background:linear-gradient(180deg,#ffffff,#f8fbff);
    border-radius:25px;
    padding:30px;
    border:1px solid #E5E7EB;
    box-shadow:0 12px 35px rgba(0,0,0,.08);">

    <h1 style="
    color:#006747;
    margin:0;
    font-size:42px;
    font-weight:700;">
    📁 Project Portfolio
    </h1>

    <p style="
    margin-top:10px;
    color:#6B7280;
    font-size:18px;">
    SmartPay Project Management System
    </p>

    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================================
    # FILTERS
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:28px;
    font-weight:700;">
    🎯 Filters
    </h2>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        search = st.text_input(
            "Search Project",
            placeholder="🔍 Search Project..."
        )

    with c2:
        allocation = st.selectbox(
            "Team Member",
            ["All"] + sorted(df["Allocation"].dropna().unique())
        )

    with c3:
        status = st.selectbox(
            "Status",
            ["All"] + sorted(df["Status"].dropna().unique())
        )

    with c4:
        category = st.selectbox(
            "Category",
            ["All"] + sorted(df["Category"].dropna().unique())
        )

    # =====================================================
    # FILTERING
    # =====================================================

    filtered_df = df.copy()

    if search:
        filtered_df = filtered_df[
            filtered_df["Mandate"].astype(str).str.contains(
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

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================================
    # SUMMARY
    # =====================================================

    left, right = st.columns([3,1])

    with left:

        st.markdown(f"""
        <div style="
        background:#E8F5E9;
        padding:18px;
        border-radius:15px;
        color:#006747;
        font-size:20px;
        font-weight:700;">
        📌 Showing <b>{len(filtered_df)}</b> Project(s)
        </div>
        """, unsafe_allow_html=True)

    with right:

        csv = filtered_df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "⬇ Download CSV",
            csv,
            "Projects.csv",
            "text/csv",
            width="stretch"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================================
    # STATUS CARDS
    # =====================================================

    s1,s2,s3,s4,s5,s6 = st.columns(6)

    def status_card(title,value,color):

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:20px;
        padding:20px;
        border-top:7px solid {color};
        text-align:center;
        box-shadow:0 8px 20px rgba(0,0,0,.08);
        min-height:120px;">

        <div style="
        color:#6B7280;
        font-size:15px;
        font-weight:600;">
        {title}
        </div>

        <div style="
        margin-top:10px;
        color:{color};
        font-size:40px;
        font-weight:700;">
        {value}
        </div>

        </div>
        """, unsafe_allow_html=True)

    with s1:
        status_card("LIVE",
                    len(filtered_df[filtered_df["Status"].str.upper()=="LIVE"]),
                    "#00C853")

    with s2:
        status_card("UAT",
                    len(filtered_df[filtered_df["Status"].str.upper()=="UAT"]),
                    "#F9A825")

    with s3:
        status_card("Development",
                    len(filtered_df[filtered_df["Status"].str.upper()=="UNDER DEVELOPMENT"]),
                    "#FF9800")

    with s4:
        status_card("IS Review",
                    len(filtered_df[filtered_df["Status"].str.upper()=="IS REVIEW"]),
                    "#00ACC1")

    with s5:
        status_card("CMC",
                    len(filtered_df[filtered_df["Status"].str.upper()=="CMC"]),
                    "#3949AB")

    with s6:
        status_card("Scoping",
                    len(filtered_df[filtered_df["Status"].str.upper()=="UNDER SCOPING"]),
                    "#8E24AA")

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================================
    # PROJECT TABLE
    # =====================================================

    st.markdown("""
    <h2 style="
    color:#006747;
    font-size:30px;
    font-weight:700;">
    📋 Project Details
    </h2>
    """, unsafe_allow_html=True)

    st.dataframe(
        filtered_df,
        width="stretch",
        hide_index=True,
        height=700
    )
# =====================================================
# ANALYTICS
# =====================================================

elif page == "Analytics":

    st.markdown("""
    <div style="
    background:linear-gradient(180deg,#ffffff,#f8fbff);
    border-radius:24px;
    padding:28px;
    border:1px solid #E5E7EB;
    box-shadow:0 14px 35px rgba(0,0,0,.10);">

    <h1 style="
    color:#006747;
    margin:0;
    font-size:40px;
    font-weight:700;">
    📊 Analytics Dashboard
    </h1>

    <p style="
    color:#6B7280;
    font-size:17px;
    margin-top:10px;">
    SmartPay Project Insights & Team Performance
    </p>

    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    # ==========================================
    # KPI CARDS
    # ==========================================

    status = df["Status"].astype(str).str.upper().str.strip()

    total_projects = len(df)
    live_projects = len(df[status == "LIVE"])
    uat_projects = len(df[status == "UAT"])
    team_members = df["Allocation"].nunique()

    live_percent = round((live_projects / total_projects) * 100, 1) if total_projects else 0
    uat_percent = round((uat_projects / total_projects) * 100, 1) if total_projects else 0


    def analytics_card(title, value, color):

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:22px;
        padding:24px;
        text-align:center;
        border-top:7px solid {color};
        box-shadow:0 10px 30px rgba(0,0,0,.10);
        min-height:140px;">

        <div style="
        color:#6B7280;
        font-size:15px;
        font-weight:600;">
        {title}
        </div>

        <div style="
        color:{color};
        font-size:46px;
        font-weight:700;
        margin-top:18px;">
        {value}
        </div>

        </div>
        """, unsafe_allow_html=True)


    k1, k2, k3, k4 = st.columns(4)

    with k1:
        analytics_card("Total Projects", total_projects, "#006747")

    with k2:
        analytics_card("Live %", f"{live_percent}%", "#00C853")

    with k3:
        analytics_card("UAT %", f"{uat_percent}%", "#F9A825")

    with k4:
        analytics_card("Team Members", team_members, "#3949AB")

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # FILTERS
    # ==========================================

    f1, f2 = st.columns([2,2])

    with f1:
        selected_allocation = st.selectbox(
            "👤 Team Member",
            ["All"] + sorted(df["Allocation"].dropna().unique())
        )

    with f2:
        selected_status = st.selectbox(
            "📌 Status",
            ["All"] + sorted(df["Status"].dropna().unique())
        )

    analytics_df = df.copy()

    if selected_allocation != "All":
        analytics_df = analytics_df[
            analytics_df["Allocation"] == selected_allocation
        ]

    if selected_status != "All":
        analytics_df = analytics_df[
            analytics_df["Status"] == selected_status
        ]

    analytics_df["Status"] = (
        analytics_df["Status"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    status_order = [
        "LIVE",
        "UAT",
        "UNDER DEVELOPMENT",
        "IS REVIEW",
        "CMC",
        "UNDER SCOPING"
    ]

    color_map = {
        "LIVE":"#00C853",
        "UAT":"#F9A825",
        "UNDER DEVELOPMENT":"#FF9800",
        "IS REVIEW":"#00ACC1",
        "CMC":"#3949AB",
        "UNDER SCOPING":"#8E24AA"
    }

    # ==========================================
    # STATUS CHART
    # ==========================================

    st.markdown("""
    <h2 style="color:#006747;font-size:30px;">
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
        height=470,
        plot_bgcolor="white",
        paper_bgcolor="white",
        title=None,
        coloraxis_showscale=False
    )

    fig1.update_traces(
        textposition="inside",
        marker_line_width=0
    )

    st.plotly_chart(fig1, width="stretch")

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # TEAM PERFORMANCE
    # ==========================================

    st.markdown("""
    <h2 style="color:#006747;font-size:30px;">
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
        height=470,
        plot_bgcolor="white",
        paper_bgcolor="white",
        coloraxis_showscale=False,
        xaxis_title="",
        yaxis_title="Projects",
        margin=dict(l=20, r=20, t=20, b=20)
    )

    fig2.update_traces(
        textposition="outside"
    )

    st.plotly_chart(fig2, width="stretch")

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # PERSON ANALYTICS
    # ==========================================

    if selected_allocation != "All":

        left, right = st.columns([2,3])

        with left:

            st.markdown(f"""
            <h2 style="
            color:#006747;
            font-size:28px;">
            🥧 {selected_allocation}
            </h2>
            """, unsafe_allow_html=True)

            person_df = analytics_df[
                analytics_df["Allocation"] == selected_allocation
            ]

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
                height=430
            )

            st.plotly_chart(fig3, width="stretch")

        with right:

            st.markdown(f"""
            <h2 style="
            color:#006747;
            font-size:28px;">
            📋 Projects Assigned
            </h2>
            """, unsafe_allow_html=True)

            mandate_count = (
                person_df["Mandate"]
                .value_counts()
                .reset_index()
            )

            mandate_count.columns = [
                "Project",
                "Count"
            ]

            fig4 = px.bar(
                mandate_count,
                x="Project",
                y="Count",
                text="Count",
                color="Count",
                color_continuous_scale="Greens"
            )

            fig4.update_layout(
                height=430,
                xaxis_title="",
                yaxis_title="Projects",
                xaxis_tickangle=-45,
                plot_bgcolor="white",
                paper_bgcolor="white",
                coloraxis_showscale=False,
                margin=dict(l=20, r=20, t=20, b=120),
                xaxis=dict(
                    tickfont=dict(size=10)
                )
            )

            st.plotly_chart(fig4, width="stretch")

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("""
        <h2 style="
        color:#006747;
        font-size:28px;">
        📄 Project Details
        </h2>
        """, unsafe_allow_html=True)

        st.dataframe(
            person_df,
            width="stretch",
            hide_index=True
        )
# =====================================================
# PROJECT TIMELINE
# =====================================================

elif page == "Project Timeline":

    st.markdown("""
    <h1 style='color:#006747;'>
    Project Timeline
    </h1>
    """, unsafe_allow_html=True)

    st.caption("Track the current progress of SmartPay projects.")

    st.markdown("""
    <div style="
    background:#F8FAFC;
    padding:15px;
    border-radius:12px;
    border:1px solid #E5E7EB;
    margin-bottom:20px;">

    <b>Timeline Legend</b><br><br>

    🟢 <b>Completed Stage</b> &nbsp;&nbsp;&nbsp;
    🟡 <b>Current Stage</b> &nbsp;&nbsp;&nbsp;
    ⚪ <b>Pending Stage</b>

    </div>
    """, unsafe_allow_html=True)

    stages = [
        "SCOPING",
        "DEVELOPMENT",
        "UAT",
        "IS REVIEW",
        "CMC",
        "LIVE"
    ]

    # ==========================
    # SEARCH TYPE
    # ==========================

    search_type = st.radio(
        "Search By",
        ["Project", "Team Member"],
        horizontal=True
    )

    if search_type == "Project":

        selected = st.selectbox(
            "Select Project",
            ["All Projects"] + sorted(df["Mandate"].dropna().unique())
        )

    else:

        selected = st.selectbox(
            "Select Team Member",
            ["All Members"] + sorted(df["Allocation"].dropna().unique())
        )

    st.markdown("---")

    # ==========================
    # TIMELINE FUNCTION
    # ==========================

    def show_timeline(project):

        current = str(project["Status"]).upper()

        if current == "SIT":
            current = "DEVELOPMENT"

        if current not in stages:
            current = "SCOPING"

        current_index = stages.index(current)

        st.markdown(f"### 📌 {project['Mandate']}")

        html = "<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;'>"

        for i, stage in enumerate(stages):

            if i < current_index:
                color = "#16A34A"

            elif i == current_index:
                color = "#F59E0B"

            else:
                color = "#D1D5DB"

            html += f"""
            <div style='text-align:center;width:16%;'>

                <div style='
                width:30px;
                height:30px;
                background:{color};
                border-radius:50%;
                margin:auto;
                border:2px solid white;
                box-shadow:0 0 8px rgba(0,0,0,.20);'>
                </div>

                <div style='
                margin-top:8px;
                font-size:13px;
                font-weight:bold;'>
                {stage}
                </div>

            </div>
            """

        html += "</div>"

        st.components.v1.html(
            html,
            height=110,
            scrolling=False
        )

        c1, c2, c3 = st.columns(3)

        c1.metric("Project", project["Mandate"])
        c2.metric("Owner", project["Allocation"])
        c3.metric("Current Stage", current)

        st.dataframe(
            project.to_frame().T,
            hide_index=True,
            use_container_width=True
        )

        st.markdown("---")

    # ==========================
    # DISPLAY
    # ==========================

    if search_type == "Project":

        if selected == "All Projects":

            for _, project in df.iterrows():
                show_timeline(project)

        else:

            project = df[df["Mandate"] == selected].iloc[0]
            show_timeline(project)

    else:

        if selected == "All Members":

            for _, project in df.iterrows():
                show_timeline(project)

        else:

            member_df = df[df["Allocation"] == selected]

            st.success(f"{selected} is handling {len(member_df)} project(s).")

            for _, project in member_df.iterrows():
                show_timeline(project)

# =====================================================
# TEAM PERFORMANCE
# =====================================================

elif page == "Team Performance":

    st.markdown("""
    <h1 style='color:#006747;'>
    Team Performance Dashboard
    </h1>
    """, unsafe_allow_html=True)

    st.caption("Monitor workload and project distribution across SmartPay team.")

    st.markdown("---")

    # ==========================
    # Team Member Filter
    # ==========================

    member = st.selectbox(
        "Select Team Member",
        ["All"] + sorted(df["Allocation"].dropna().unique())
    )

    team_df = df.copy()

    if member != "All":
        team_df = team_df[team_df["Allocation"] == member]

    status = team_df["Status"].astype(str).str.upper().str.strip()

    # ==========================
    # KPI Cards
    # ==========================

    total = len(team_df)
    live = len(team_df[status == "LIVE"])
    uat = len(team_df[status == "UAT"])
    development = len(team_df[status == "UNDER DEVELOPMENT"])
    review = len(team_df[status == "IS REVIEW"])
    cmc = len(team_df[status == "CMC"])
    scoping = len(team_df[status == "UNDER SCOPING"])

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)

    c1.metric("Total", total)
    c2.metric("Live", live)
    c3.metric("UAT", uat)
    c4.metric("Development", development)
    c5.metric("Review", review)
    c6.metric("CMC", cmc)
    c7.metric("Under Scoping", scoping)

    st.markdown("---")

    # ==========================
    # Projects by Status
    # ==========================

    st.subheader("Projects by Status")

    status_df = (
        team_df["Status"]
        .value_counts()
        .reset_index()
    )

    status_df.columns = ["Status", "Projects"]

    fig = px.bar(
        status_df,
        x="Status",
        y="Projects",
        text="Projects",
        color="Status",
        title="Status Distribution"
    )

    st.plotly_chart(fig, width="stretch")

    # ==========================
    # Projects by Team
    # ==========================

    if member == "All":

        st.subheader("Projects by Team Member")

        allocation_df = (
            df["Allocation"]
            .value_counts()
            .reset_index()
        )

        allocation_df.columns = ["Allocation", "Projects"]

        fig2 = px.bar(
            allocation_df,
            x="Allocation",
            y="Projects",
            text="Projects",
            color="Allocation",
            title="Projects Assigned to Team Members"
        )

        st.plotly_chart(fig2, width="stretch")

    else:

        st.subheader(f"{member} Project List")

        st.dataframe(
            team_df[
                ["Mandate", "Status", "Allocation"]
            ],
            hide_index=True,
            width="stretch"
        )

    st.markdown("---")

    # ==========================
    # Top Performer
    # ==========================

    top = (
        df["Allocation"]
        .value_counts()
        .reset_index()
    )

    top.columns = ["Member", "Projects"]

    winner = top.iloc[0]

    st.success(
        f"🏆 Top Workload: {winner['Member']} "
        f"({winner['Projects']} Projects)"
    )
# =====================================================
# VOICE SEARCH
# =====================================================

elif page == "Voice Search":

    st.markdown("""
    <h1 style="color:#006747;">
    🎤 SmartPay Voice Search
    </h1>
    """, unsafe_allow_html=True)

    st.info("🎤 Speak Project, Person, Status, Category or say 'Show All Projects'")

    # Microphone (streamlit_mic_recorder)
    voice_text = listen()

    st.write("Raw Voice :", repr(voice_text))

    voice_text = normalize_voice(str(voice_text))

    st.write("Normalized :", voice_text)

    if voice_text:

        voice_text = voice_text.lower().strip()

        st.success(f"🎤 You said: {voice_text}")

        speak(f"You said {voice_text}")

        # =========================================
        # Smart Commands
        # =========================================

        if voice_text == "all":

            result = df.copy()

        elif voice_text == "live":

            result = df[
                df["Status"].astype(str).str.lower() == "live"
            ]

        elif voice_text == "uat":

            result = df[
                df["Status"].astype(str).str.lower() == "uat"
            ]

        elif voice_text == "sit":

            result = df[
                df["Status"].astype(str).str.lower() == "sit"
            ]

        elif voice_text == "under development":

            result = df[
                df["Status"].astype(str).str.lower() == "under development"
            ]

        else:

            search_cols = [
                "Mandate",
                "Allocation",
                "Status",
                "Category",
                "Update"
            ]

            result = df[
                df.apply(
                    lambda row: any(
                        voice_text in str(row[col]).lower()
                        for col in search_cols
                    ),
                    axis=1
                )
            ]

            if result.empty:

                search_values = []

                for col in search_cols:
                    search_values.extend(
                        df[col].dropna().astype(str).tolist()
                    )

                match = process.extractOne(
                    voice_text,
                    search_values,
                    scorer=fuzz.token_sort_ratio
                )

                if match and match[1] >= 70:

                    matched = match[0].lower()

                    result = df[
                        df.apply(
                            lambda row: any(
                                matched in str(row[col]).lower()
                                for col in search_cols
                            ),
                            axis=1
                        )
                    ]

        # =========================================
        # Result
        # =========================================

        if result.empty:

            st.error("❌ No Project Found")

            speak("Sorry. No matching project found.")

        else:

            st.success(f"✅ {len(result)} Project(s) Found")

            st.metric("Total Results", len(result))

            if len(result) == 1:

                first = result.iloc[0]

                status = str(first["Status"]).upper()

                if status == "LIVE":
                    badge = "#16A34A"
                elif status == "UAT":
                    badge = "#F59E0B"
                elif status == "SIT":
                    badge = "#2563EB"
                else:
                    badge = "#6B7280"

                st.markdown(
                    f"""
                    <div style="
                    background:white;
                    border-radius:18px;
                    padding:25px;
                    border-left:8px solid {badge};
                    box-shadow:0 8px 18px rgba(0,0,0,.08);
                    margin-bottom:20px;">

                    <h2 style="color:#006747;margin-top:0;">
                    {first['Mandate']}
                    </h2>

                    <table style="width:100%;font-size:16px;">

                    <tr>
                    <td><b>Status</b></td>
                    <td>{first['Status']}</td>
                    </tr>

                    <tr>
                    <td><b>Owner</b></td>
                    <td>{first['Allocation']}</td>
                    </tr>

                    <tr>
                    <td><b>Category</b></td>
                    <td>{first['Category']}</td>
                    </tr>

                    <tr>
                    <td><b>Latest Update</b></td>
                    <td>{first['Update']}</td>
                    </tr>

                    </table>

                    </div>
                    """,
                    unsafe_allow_html=True
                )

                response = (
                    f"{first['Mandate']} is currently "
                    f"{first['Status']} and allocated to "
                    f"{first['Allocation']}. "
                    f"Latest update is {first['Update']}."
                )

            else:

                st.dataframe(
                    result,
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("### Projects Found")

                c1, c2, c3 = st.columns(3)

                c1.metric(
                    "Projects",
                    len(result)
                )

                c2.metric(
                    "Owners",
                    result["Allocation"].nunique()
                )

                c3.metric(
                    "Live",
                    len(
                        result[
                            result["Status"].astype(str).str.upper() == "LIVE"
                        ]
                    )
                )

                for _, row in result.iterrows():

                    st.markdown(
                        f"""
                        <div style="
                        background:white;
                        padding:18px;
                        border-radius:15px;
                        margin-bottom:12px;
                        border:1px solid #E5E7EB;
                        box-shadow:0 4px 10px rgba(0,0,0,.06);">

                        <h4 style="color:#006747;margin:0;">
                        {row['Mandate']}
                        </h4>

                        <p style="margin-top:8px;">
                        <b>Owner:</b> {row['Allocation']}<br>
                        <b>Status:</b> {row['Status']}<br>
                        <b>Category:</b> {row['Category']}
                        </p>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                names = ", ".join(
                    result["Mandate"].astype(str).tolist()
                )

                response = (
                    f"{len(result)} projects found. "
                    f"The projects are {names}."
                )

            st.info(response)

            speak(response)
# =====================================================
# EXPORT
# =====================================================

elif page == "Export":

    st.markdown("""
    <h1 style="color:#006747;">
    Executive Report Center
    </h1>
    """, unsafe_allow_html=True)

    st.caption("Generate and download SmartPay project reports.")

    st.markdown("""
    <div style="
    background:white;
    padding:25px;
    border-radius:18px;
    border:1px solid #E5E7EB;
    box-shadow:0 6px 18px rgba(0,0,0,.08);
    margin-bottom:25px;">

    <h3 style="color:#006747;">
    Report Includes
    </h3>

    ✔ Dashboard Summary<br>
    ✔ KPI Overview<br>
    ✔ Project Details<br>
    ✔ Team Performance<br>
    ✔ Analytics Summary<br>
    ✔ Project Timeline<br><br>

    <b>Generated On:</b>
    </div>
    """, unsafe_allow_html=True)

    st.info(datetime.now().strftime("%d %B %Y  |  %I:%M %p"))

    st.markdown("---")

    # ==========================
    # CSV
    # ==========================

    csv = df.to_csv(index=False).encode("utf-8")

    # ==========================
    # Excel
    # ==========================

    excel_buffer = BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="SmartPay Projects"
        )

    excel_data = excel_buffer.getvalue()

    c1, c2, c3 = st.columns(3)

    with c1:

        st.download_button(
            "📄 Export CSV",
            csv,
            "SmartPay_Projects.csv",
            "text/csv",
            use_container_width=True
        )

    with c2:

        st.download_button(
            "📊 Export Excel",
            excel_data,
            "SmartPay_Projects.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with c3:

        if st.button(
            "📄 Export Executive PDF",
            use_container_width=True
        ):

            pdf = generate_executive_pdf(df)

            st.download_button(
                "⬇ Download Executive PDF",
                data=pdf,
                file_name="SmartPay_Executive_Report.pdf",
                mime="application/pdf",
                use_container_width=True
            )

            st.success("Executive PDF Generated Successfully.")

        st.markdown("---")

    st.subheader("Report Preview")

    st.dataframe(
        df.head(10),
        use_container_width=True,
         hide_index=True
    )

    st.success(
        f"""
        Total Projects : {len(df)}

        Report Ready For Export
        """
    )