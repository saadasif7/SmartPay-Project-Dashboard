import streamlit as st
import plotly.express as px
from datetime import datetime
from utils.excel_reader import load_data, get_weather
from utils.voice import listen
import streamlit.components.v1 as components
from rapidfuzz import process, fuzz
import plotly.express as px
from io import BytesIO
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image
from utils.pdf_report import generate_executive_pdf
import requests
from datetime import datetime

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
    # Header
    # ==========================================

    city, temp, weather = get_weather()

    left, center, right = st.columns([5, 2, 2])

    # ==========================================
    # LEFT - DASHBOARD TITLE
    # ==========================================

    with left:

        st.markdown("""
        <div style="
        background:white;
        border-radius:18px;
        padding:30px;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 18px rgba(0,0,0,.08);">

        <h1 style="
        color:#006747;
        margin:0;
        font-size:42px;
        font-weight:700;">
        SmartPay Project Dashboard
        </h1>

        <p style="
        color:#6B7280;
        margin-top:10px;
        font-size:18px;">
        Digital Banking Group | National Bank of Pakistan
        </p>

        <small style="color:#9CA3AF;">
        Enterprise Project Monitoring System
        </small>

        </div>
        """, unsafe_allow_html=True)


    # ==========================================
    # CENTER - WEATHER
    # ==========================================

    with center:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        text-align:center;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 18px rgba(0,0,0,.08);">

        <h3 style="margin:0;">🌤 {city}</h3>

        <h1 style="margin:8px 0;color:#006747;">
        {temp}°C
        </h1>

        <p style="color:gray;margin:0;">
        {weather}
        </p>

        </div>
        """, unsafe_allow_html=True)


    # ==========================================
    # RIGHT - DATE & TIME
    # ==========================================

    with right:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        text-align:center;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 18px rgba(0,0,0,.08);">

        <h4 style="color:#006747;margin:0;">
        Today
        </h4>

        <h2 style="margin-top:10px;">
        {datetime.now().strftime("%d %b %Y")}
        </h2>

        <hr>

        <h3>
        {datetime.now().strftime("%I:%M %p")}
        </h3>

        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    # ==========================================
    # KPI
    # ==========================================

    status = df["Status"].astype(str).str.upper().str.strip()

    total_projects = len(df)

    scoping_projects = len(df[status == "UNDER SCOPING"])

    development_projects = len(df[status == "UNDER DEVELOPMENT"])

    uat_projects = len(df[status == "UAT"])

    review_projects = len(df[status == "IS REVIEW"])

    cmc_projects = len(df[status == "CMC"])

    live_projects = len(df[status == "LIVE"])

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)

    # ===========================
    # Total Projects
    # ===========================

    with c1:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border-top:6px solid #006747;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">Total Projects</h5>

        <h1 style="color:#006747;font-size:38px;">
        {total_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    # ===========================
    # Under Scoping
    # ===========================

    with c2:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border-top:6px solid #8E24AA;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">Scoping</h5>

        <h1 style="color:#8E24AA;font-size:38px;">
        {scoping_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    # ===========================
    # Under Development
    # ===========================

    with c3:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border-top:6px solid #FF9800;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">Development</h5>

        <h1 style="color:#FF9800;font-size:38px;">
        {development_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    # ===========================
    # UAT
    # ===========================

    with c4:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border-top:6px solid #F9A825;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">UAT</h5>

        <h1 style="color:#F9A825;font-size:38px;">
        {uat_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    # ===========================
    # IS Review
    # ===========================

    with c5:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border-top:6px solid #00ACC1;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">IS Review</h5>

        <h1 style="color:#00ACC1;font-size:38px;">
        {review_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    # ===========================
    # CMC
    # ===========================

    with c6:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border-top:6px solid #3949AB;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">CMC</h5>

        <h1 style="color:#3949AB;font-size:38px;">
        {cmc_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    # ===========================
    # LIVE
    # ===========================

    with c7:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border-top:6px solid #00C853;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">LIVE</h5>

        <h1 style="color:#00C853;font-size:38px;">
        {live_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    # ==========================================
    # Allocation Summary
    # ==========================================

    st.markdown("## Projects by Team Member")

    allocation = (
        df["Allocation"]
        .value_counts()
        .reset_index()
    )

    allocation.columns = ["Allocation","Projects"]

    cols = st.columns(4)

    for i,row in allocation.iterrows():

        with cols[i % 4]:

            st.markdown(f"""
            <div style="
            background:white;
            border-radius:15px;
            padding:18px;
            text-align:center;
            border:1px solid #E5E7EB;
            box-shadow:0 4px 15px rgba(0,0,0,.07);">

            <h4 style="
            color:#006747;
            margin-bottom:10px;">
            {row["Allocation"]}
            </h4>

            <h2 style="
            margin:0;
            color:#111827;">
            {row["Projects"]}
            </h2>

            <p style="color:#6B7280;">
            Total Projects
            </p>

            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("## Team Workload")

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
        color_continuous_scale="Greens",
        title="Projects Assigned to Team Members"
    )

    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_x=0.5,
        height=450,
        showlegend=False
    )

    fig.update_traces(textposition="outside")

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("## Team Summary")
    summary = (
        df.groupby("Allocation")
        .agg(
            Total=("Mandate","count"),
            Live=("Status", lambda x: (x.str.upper()=="LIVE").sum()),
            UAT=("Status", lambda x: (x.str.upper()=="UAT").sum()),
        )
        .reset_index()
    )

    st.dataframe(summary, use_container_width=True, hide_index=True)
    

    # ==========================================
    # Search
    # ==========================================

    st.markdown("## Search Project")

    project = st.text_input(
        "Search Project",
        placeholder="Type Project Name...",
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

            st.error("Project not found.")

        else:

            st.success(f"{len(result)} Project(s) Found")

            st.dataframe(
                result,
                use_container_width=True,
                hide_index=True
            )
            
 

# =====================================================
# PROJECTS
# =====================================================

elif page == "Projects":

    st.markdown("""
    <div style="
    background:white;
    padding:28px;
    border-radius:18px;
    box-shadow:0 8px 18px rgba(0,0,0,.08);
    border:1px solid #E5E7EB;">

    <h1 style="
    color:#006747;
    margin-bottom:5px;
    font-size:36px;">
    Project Portfolio
    </h1>

    <p style="
    color:#6B7280;
    font-size:17px;">
    SmartPay Projects Management
    </p>

    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # Filters
    # ==========================================

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        search = st.text_input(
            "Search Project",
            placeholder="Search Mandate..."
        )

    with c2:
        allocation = st.selectbox(
            "Allocation",
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

    # ==========================================
    # Filtering
    # ==========================================

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

    # ==========================================
    # Summary
    # ==========================================

    left, right = st.columns([3,1])

    with left:

        st.success(
            f"Showing {len(filtered_df)} Project(s)"
        )

    with right:

        csv = filtered_df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "Download CSV",
            csv,
            "Projects.csv",
            "text/csv"
        )

    # ==========================================
    # Status Summary
    # ==========================================

    s1, s2, s3, s4, s5, s6 = st.columns(6)

    s1.metric(
        "Live",
        len(filtered_df[
            filtered_df["Status"].str.upper() == "LIVE"
        ])
    )

    s2.metric(
        "UAT",
        len(filtered_df[
            filtered_df["Status"].str.upper() == "UAT"
        ])
    )

    s3.metric(
        "Development",
        len(filtered_df[
            filtered_df["Status"].str.upper() == "UNDER DEVELOPMENT"
        ])
    )

    s4.metric(
        "IS Review",
        len(filtered_df[
            filtered_df["Status"].str.upper() == "IS REVIEW"
        ])
    )

    s5.metric(
        "CMC",
        len(filtered_df[
            filtered_df["Status"].str.upper() == "CMC"
        ])
    )

    s6.metric(
        "Scoping",
        len(filtered_df[
            filtered_df["Status"].str.upper() == "UNDER SCOPING"
        ])
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # Project Table
    # ==========================================

    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        height=650
    )

# =====================================================
# ANALYTICS
# =====================================================

elif page == "Analytics":

    st.markdown("""
    <h1 style='color:#006747;font-size:38px;'>
    Analytics Dashboard
    </h1>
    <p style='color:gray;font-size:17px;'>
    Project insights and team performance
    </p>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ==========================
    # FILTERS
    # ==========================

    c1, c2 = st.columns(2)

    with c1:
        selected_allocation = st.selectbox(
            "Select Team Member",
            ["All"] + sorted(df["Allocation"].dropna().unique())
        )

    with c2:
        selected_status = st.selectbox(
            "Select Status",
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
        "LIVE": "#00C853",
        "UAT": "#F9A825",
        "UNDER DEVELOPMENT": "#FF9800",
        "IS REVIEW": "#00ACC1",
        "CMC": "#3949AB",
        "UNDER SCOPING": "#8E24AA"
    }

    # ======================================
    # GRAPH 1 - STATUS
    # ======================================

    st.subheader("Projects by Status")

    status_count = (
        analytics_df["Status"]
        .value_counts()
        .reindex(status_order, fill_value=0)
        .reset_index()
    )

    status_count.columns = ["Status", "Projects"]

    fig1 = px.bar(
        status_count,
        x="Status",
        y="Projects",
        text="Projects",
        color="Status",
        color_discrete_map=color_map
    )

    fig1.update_layout(
        height=450,
        title="Project Status Distribution",
        title_x=0.5
    )

    st.plotly_chart(fig1, width="stretch")

    st.markdown("---")

    # ======================================
    # GRAPH 2 - TEAM PERFORMANCE
    # ======================================

    st.subheader("Projects by Team Member")

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
        color="Allocation"
    )

    fig2.update_layout(height=450)

    st.plotly_chart(fig2, width="stretch")

    st.markdown("---")

    # ======================================
    # GRAPH 3 - PERSON STATUS
    # ======================================

    if selected_allocation != "All":

        st.subheader(
            f"{selected_allocation} Project Status Breakdown"
        )

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
            hole=0.45,
            color="Status",
            color_discrete_map=color_map
        )

        fig3.update_layout(height=450)

        st.plotly_chart(fig3, width="stretch")

        st.markdown("---")

        # ======================================
        # GRAPH 4 - PROJECT LIST
        # ======================================

        st.subheader(
            f"{selected_allocation} Project List"
        )

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
            color="Project"
        )

        fig4.update_layout(
            height=500,
            xaxis_tickangle=-35
        )

        st.plotly_chart(fig4, width="stretch")

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