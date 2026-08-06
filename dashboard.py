import streamlit as st
import plotly.express as px
from datetime import datetime
from utils.excel_reader import load_data
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
# ==========================================
# SIDEBAR CSS
# ==========================================

st.markdown("""
<style>

/* Sidebar Background */
section[data-testid="stSidebar"]{
    background:#006747;
}

/* Radio label */
section[data-testid="stSidebar"] label{
    color:white !important;
    font-weight:600;
}

/* Remove white background */
div[role="radiogroup"] label{
    background:transparent !important;
    border:none !important;
    box-shadow:none !important;
}

/* Selected button */
div[role="radiogroup"] label[data-baseweb="radio"]{
    border-radius:12px;
    padding:10px 12px;
    margin-bottom:8px;
}

div[role="radiogroup"] label[data-baseweb="radio"]:has(input:checked){
    background:#0A8F5B !important;
    color:white !important;
}

/* Hover */
div[role="radiogroup"] label:hover{
    background:#0C7C53 !important;
}

/* Hide radio circles */
div[role="radiogroup"] input{
    display:none;
}

</style>
""", unsafe_allow_html=True)
# Logo
st.sidebar.image(
    "smartpay_logo/smartpay_logo.png",
    use_container_width=True
)

# Title
st.sidebar.markdown("""
<h2 style="
text-align:center;
color:white;
margin-top:10px;
margin-bottom:0;
font-weight:700;">
SmartPay
</h2>

<p style="
text-align:center;
color:#D1FAE5;
font-size:15px;
margin-top:5px;">
Project Dashboard
</p>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")

# Navigation
st.sidebar.markdown("""
<h3 style="
color:white;
margin-bottom:10px;">
Navigation
</h3>
""", unsafe_allow_html=True)

page = st.sidebar.radio(
    "",
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

# Information Box
st.sidebar.markdown("""
<div style="
background:rgba(255,255,255,.10);
padding:15px;
border-radius:12px;
color:white;
font-size:14px;">

<b>Department</b><br>
Digital Banking Group

<br><br>

<b>Organization</b><br>
National Bank of Pakistan

<br><br>

<b>Version</b><br>
2.0

</div>
""", unsafe_allow_html=True)

st.sidebar.markdown("<br>", unsafe_allow_html=True)

st.sidebar.success("Developed By\n\nMuhammad Saad Asif")
# =====================================================
# DASHBOARD
# =====================================================

if page == "Dashboard":

    # ==========================================
    # Header
    # ==========================================

    left, right = st.columns([5,1])

    with left:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:35px;
        border:1px solid #E5E7EB;
        box-shadow:0 8px 18px rgba(0,0,0,.08);">

        <h1 style="
        color:#006747;
        margin:0;
        font-size:42px;
        font-weight:700;
        font-family:Segoe UI;">
        SmartPay Project Dashboard
        </h1>

        <p style="
        margin-top:12px;
        color:#6B7280;
        font-size:18px;">
        Digital Banking Group | National Bank of Pakistan
        </p>

        </div>
        """, unsafe_allow_html=True)

    with right:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:20px;
        border:1px solid #E5E7EB;
        text-align:center;
        box-shadow:0 8px 18px rgba(0,0,0,.08);">

        <h4 style="color:#006747;margin:0;">Today</h4>

        <h2 style="margin-top:10px;">
        {datetime.now().strftime("%d %b %Y")}
        </h2>

        <hr>

        <h3 style="margin-top:5px;">
        {datetime.now().strftime("%I:%M %p")}
        </h3>

        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # KPI
    # ==========================================

    total_projects = len(df)
    live_projects = len(df[df["Status"].str.upper()=="LIVE"])
    uat_projects = len(df[df["Status"].str.upper()=="UAT"])
    sit_projects = len(df[df["Status"].str.upper()=="SIT"])

    c1,c2,c3,c4 = st.columns(4)

    with c1:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:22px;
        border-top:6px solid #006747;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">Total Projects</h5>

        <h1 style="color:#006747;font-size:42px;">
        {total_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    with c2:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:22px;
        border-top:6px solid #00C853;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">LIVE</h5>

        <h1 style="color:#00C853;font-size:42px;">
        {live_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    with c3:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:22px;
        border-top:6px solid #F9A825;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">UAT</h5>

        <h1 style="color:#F9A825;font-size:42px;">
        {uat_projects}
        </h1>

        </div>
        """, unsafe_allow_html=True)

    with c4:

        st.markdown(f"""
        <div style="
        background:white;
        border-radius:18px;
        padding:22px;
        border-top:6px solid #2196F3;
        box-shadow:0 6px 18px rgba(0,0,0,.08);">

        <h5 style="color:gray;">SIT</h5>

        <h1 style="color:#2196F3;font-size:42px;">
        {sit_projects}
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

    s1, s2, s3, s4 = st.columns(4)

    s1.metric(
        "Live",
        len(filtered_df[
            filtered_df["Status"].str.upper()=="LIVE"
        ])
    )

    s2.metric(
        "UAT",
        len(filtered_df[
            filtered_df["Status"].str.upper()=="UAT"
        ])
    )

    s3.metric(
        "SIT",
        len(filtered_df[
            filtered_df["Status"].str.upper()=="SIT"
        ])
    )

    s4.metric(
        "Development",
        len(filtered_df[
            filtered_df["Status"].str.contains(
                "Development",
                case=False,
                na=False
            )
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

    # ======================================
    # GRAPH 1
    # PROJECTS BY STATUS
    # ======================================

    st.subheader("Projects by Status")

    status_count = analytics_df["Status"].value_counts().reset_index()
    status_count.columns = ["Status", "Projects"]

    fig1 = px.bar(
        status_count,
        x="Status",
        y="Projects",
        text="Projects",
        color="Status",
        color_discrete_sequence=px.colors.qualitative.Set2
    )

    fig1.update_layout(
        height=450,
        title_x=0.5
    )

    st.plotly_chart(fig1, use_container_width=True)

    st.markdown("---")

    # ======================================
    # GRAPH 2
    # TEAM PERFORMANCE
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

    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")

    # ======================================
    # GRAPH 3
    # SELECTED PERSON STATUS
    # ======================================

    if selected_allocation != "All":

        st.subheader(
            f"{selected_allocation} Project Status Breakdown"
        )

        person_df = df[
            df["Allocation"] == selected_allocation
        ]

        person_status = (
            person_df["Status"]
            .value_counts()
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
            hole=.45
        )

        fig3.update_layout(height=450)

        st.plotly_chart(fig3, use_container_width=True)

        st.markdown("---")

        # ======================================
        # GRAPH 4
        # PROJECT NAMES
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

        st.plotly_chart(
            fig4,
            use_container_width=True
        )

        st.dataframe(
            person_df,
            use_container_width=True,
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

    # Team Member Filter
    member = st.selectbox(
        "Select Team Member",
        ["All"] + sorted(df["Allocation"].dropna().unique())
    )

    team_df = df.copy()

    if member != "All":
        team_df = team_df[team_df["Allocation"] == member]

    # ==========================
    # KPI Cards
    # ==========================

    total = len(team_df)
    live = len(team_df[team_df["Status"].str.upper()=="LIVE"])
    uat = len(team_df[team_df["Status"].str.upper()=="UAT"])
    sit = len(team_df[team_df["Status"].str.upper()=="SIT"])

    c1,c2,c3,c4 = st.columns(4)

    c1.metric("Total Projects", total)
    c2.metric("Live", live)
    c3.metric("UAT", uat)
    c4.metric("SIT", sit)

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

    status_df.columns=["Status","Projects"]

    fig = px.bar(
        status_df,
        x="Status",
        y="Projects",
        text="Projects",
        color="Status",
        title="Status Distribution"
    )

    st.plotly_chart(fig, use_container_width=True)

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

        allocation_df.columns=["Allocation","Projects"]

        fig2 = px.bar(
            allocation_df,
            x="Allocation",
            y="Projects",
            text="Projects",
            color="Allocation",
            title="Projects Assigned to Team Members"
        )

        st.plotly_chart(fig2, use_container_width=True)

    else:

        st.subheader(f"{member} Project List")

        st.dataframe(
            team_df[
                ["Mandate","Status","Allocation"]
            ],
            hide_index=True,
            use_container_width=True
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

    top.columns=["Member","Projects"]

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