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
import json
import base64
from zoneinfo import ZoneInfo
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





# =====================================================
# DAILY ISSUES - PERSISTENT GITHUB STORAGE
# =====================================================

DAILY_ISSUES_REPO = "saadasif7/SmartPay-Project-Dashboard"
DAILY_ISSUES_BRANCH = "main"
DAILY_ISSUES_PATH = "data/daily_issues.json"


# =====================================================
# DAILY ISSUES - CORE COLUMNS
# =====================================================

DAILY_ISSUES_CORE_COLUMNS = [
    "Issue",
    "Update",
    "End / Pending With",
    "Team Member",
    "Status"
]


# =====================================================
# GET GITHUB TOKEN
# =====================================================

def get_github_token():

    try:
        token = st.secrets.get(
            "GITHUB_TOKEN",
            ""
        )
    except Exception:
        token = ""

    return str(token).strip()


# =====================================================
# GITHUB HEADERS
# =====================================================

def github_headers():

    token = get_github_token()

    return {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2026-03-10",
        "Content-Type": "application/json"
    }


# =====================================================
# EMPTY DAILY ISSUES DATAFRAME
# =====================================================

def empty_daily_issues_df():

    return pd.DataFrame(
        columns=DAILY_ISSUES_CORE_COLUMNS
    )


# =====================================================
# NORMALIZE DAILY ISSUES DATA
# =====================================================

def normalize_daily_issues_df(df):

    if df is None:
        df = empty_daily_issues_df()

    df = df.copy()

    # -------------------------------------------------
    # Ensure Core Columns Exist
    # -------------------------------------------------

    for col in DAILY_ISSUES_CORE_COLUMNS:

        if col not in df.columns:

            if col == "Status":
                df[col] = "Active"
            else:
                df[col] = ""

    # -------------------------------------------------
    # Keep Core Columns First
    # -------------------------------------------------

    extra_columns = [
        col
        for col in df.columns
        if col not in DAILY_ISSUES_CORE_COLUMNS
    ]

    df = df[
        DAILY_ISSUES_CORE_COLUMNS +
        extra_columns
    ]

    # -------------------------------------------------
    # Fill Missing Values
    # -------------------------------------------------

    df = df.fillna("")

    # -------------------------------------------------
    # Normalize Status
    # -------------------------------------------------

    def normalize_status(value):

        value = str(value).strip().lower()

        if value == "closed":
            return "Closed"

        return "Active"

    df["Status"] = df["Status"].apply(
        normalize_status
    )

    # -------------------------------------------------
    # Convert All Values to String
    # -------------------------------------------------

    df = df.astype(str)

    # -------------------------------------------------
    # Remove Completely Blank Rows
    # -------------------------------------------------

    df = df[
        df.apply(
            lambda row:
            any(
                str(value).strip()
                for value in row
            ),
            axis=1
        )
    ].reset_index(
        drop=True
    )

    return df


# =====================================================
# LOAD DAILY ISSUES FROM GITHUB
# =====================================================

def load_daily_issues_from_github():

    token = get_github_token()

    # -------------------------------------------------
    # Token Missing
    # -------------------------------------------------

    if not token:

        return empty_daily_issues_df()

    url = (
        f"https://api.github.com/repos/"
        f"{DAILY_ISSUES_REPO}/contents/"
        f"{DAILY_ISSUES_PATH}"
    )

    try:

        # -------------------------------------------------
        # GET FILE FROM GITHUB
        # -------------------------------------------------

        response = requests.get(
            url,
            headers=github_headers(),
            params={
                "ref": DAILY_ISSUES_BRANCH
            },
            timeout=20
        )

        # -------------------------------------------------
        # File Does Not Exist
        # -------------------------------------------------

        if response.status_code == 404:

            return empty_daily_issues_df()

        # -------------------------------------------------
        # Raise Other Errors
        # -------------------------------------------------

        response.raise_for_status()

        result = response.json()

        # -------------------------------------------------
        # Get Encoded Content
        # -------------------------------------------------

        encoded_content = result.get(
            "content",
            ""
        ).replace(
            "\n",
            ""
        )

        if not encoded_content:

            return empty_daily_issues_df()

        # -------------------------------------------------
        # Decode Base64
        # -------------------------------------------------

        decoded_content = base64.b64decode(
            encoded_content
        ).decode(
            "utf-8"
        )

        # -------------------------------------------------
        # Parse JSON
        # -------------------------------------------------

        data = json.loads(
            decoded_content
        )

        if not isinstance(data, list):

            data = []

        # -------------------------------------------------
        # Convert to DataFrame
        # -------------------------------------------------

        daily_df = pd.DataFrame(data)

        # -------------------------------------------------
        # Normalize
        # -------------------------------------------------

        daily_df = normalize_daily_issues_df(
            daily_df
        )

        return daily_df

    except Exception as e:

        st.error(
            f"Daily Issues load error: {e}"
        )

        return empty_daily_issues_df()


# =====================================================
# SAVE DAILY ISSUES TO GITHUB
# =====================================================

def save_daily_issues_to_github(edited_df):

    token = get_github_token()

    # -------------------------------------------------
    # TOKEN CHECK
    # -------------------------------------------------

    if not token:

        return False, (
            "GITHUB_TOKEN is not configured. "
            "Please add it in Streamlit Secrets."
        )

    url = (
        f"https://api.github.com/repos/"
        f"{DAILY_ISSUES_REPO}/contents/"
        f"{DAILY_ISSUES_PATH}"
    )

    try:

        # =================================================
        # GET CURRENT FILE SHA
        # =================================================

        get_response = requests.get(
            url,
            headers=github_headers(),
            params={
                "ref": DAILY_ISSUES_BRANCH
            },
            timeout=20
        )

        sha = None

        if get_response.status_code == 200:

            sha = get_response.json().get(
                "sha"
            )

        elif get_response.status_code != 404:

            get_response.raise_for_status()

        # =================================================
        # PREPARE DATA
        # =================================================

        clean_df = edited_df.copy()

        # -------------------------------------------------
        # Ensure Core Columns Exist
        # -------------------------------------------------

        for col in DAILY_ISSUES_CORE_COLUMNS:

            if col not in clean_df.columns:

                if col == "Status":
                    clean_df[col] = "Active"

                else:
                    clean_df[col] = ""

        # -------------------------------------------------
        # Preserve Extra Columns
        # -------------------------------------------------

        extra_columns = [
            col
            for col in clean_df.columns
            if col not in DAILY_ISSUES_CORE_COLUMNS
        ]

        final_columns = (
            DAILY_ISSUES_CORE_COLUMNS +
            extra_columns
        )

        clean_df = clean_df[
            final_columns
        ]

        # -------------------------------------------------
        # Fill Empty Values
        # -------------------------------------------------

        clean_df = clean_df.fillna("")

        # -------------------------------------------------
        # Normalize Status
        # -------------------------------------------------

        clean_df["Status"] = clean_df[
            "Status"
        ].apply(
            lambda value:
            "Closed"
            if str(value).strip().lower() == "closed"
            else "Active"
        )

        # -------------------------------------------------
        # Convert Everything to String
        # -------------------------------------------------

        clean_df = clean_df.astype(str)

        # =================================================
        # REMOVE COMPLETELY BLANK ROWS
        # =================================================

        clean_df = clean_df[
            clean_df.apply(
                lambda row:
                any(
                    str(value).strip()
                    for value in row
                ),
                axis=1
            )
        ].reset_index(
            drop=True
        )

        # =================================================
        # CONVERT DATA TO JSON
        # =================================================

        data = clean_df.to_dict(
            orient="records"
        )

        json_content = json.dumps(
            data,
            ensure_ascii=False,
            indent=2
        )

        # =================================================
        # BASE64 ENCODE
        # =================================================

        encoded_content = base64.b64encode(
            json_content.encode(
                "utf-8"
            )
        ).decode(
            "utf-8"
        )

        # =================================================
        # GITHUB UPDATE PAYLOAD
        # =================================================

        payload = {
            "message": (
                "Update Daily Issues "
                "from SmartPay Dashboard"
            ),
            "content": encoded_content,
            "branch": DAILY_ISSUES_BRANCH
        }

        # -------------------------------------------------
        # Existing File -> SHA Required
        # -------------------------------------------------

        if sha:

            payload["sha"] = sha

        # =================================================
        # CREATE / UPDATE FILE
        # =================================================

        response = requests.put(
            url,
            headers=github_headers(),
            json=payload,
            timeout=30
        )

        response.raise_for_status()

        return True, (
            "Daily Issues saved successfully."
        )

    except requests.exceptions.HTTPError as e:

        try:
            error_details = response.json()

            message = error_details.get(
                "message",
                str(e)
            )

        except Exception:

            message = str(e)

        return False, (
            f"GitHub save error: {message}"
        )

    except Exception as e:

        return False, str(e)

# =====================================================
# EMBEDDED BACKGROUND IMAGES (base64, self-contained)
# =====================================================

HEADER_BG_B64 = "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"
SIDEBAR_BG_B64 = "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"


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

/* =====================================================
   HIDE SIDEBAR COLLAPSE ICON
===================================================== */

[data-testid="stSidebarCollapseButton"],
[data-testid="stSidebarCollapsedControl"],
[data-testid="baseButton-headerNoPadding"] {
    display:none !important;
    visibility:hidden !important;
}


/* =====================================================
   MAIN APP - SOFT LIGHT GREEN
===================================================== */

.stApp {
    background:
        linear-gradient(
            180deg,
            #EDF7F2 0%,
            #F1F9F5 45%,
            #EAF5EF 100%
        ) !important;
}


/* =====================================================
   MAIN CONTENT
===================================================== */

.block-container {
    padding-top:2.5rem !important;
}


/* =====================================================
   HEADINGS
===================================================== */

h1,
h2,
h3,
h4,
h5,
h6,
[data-testid="stHeading"],
[data-testid="stHeading"] * {
    color:#006747 !important;
    font-weight:700 !important;
}


/* =====================================================
   NORMAL TEXT
===================================================== */

p,
span {
    color:inherit !important;
}


/* =====================================================
   FIELD LABELS
===================================================== */

label,
label p,
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] p {
    color:#006747 !important;
    font-weight:bold !important;
    font-size:16px !important;
}


/* =====================================================
   SIDEBAR
===================================================== */

section[data-testid="stSidebar"] {
    background:
        linear-gradient(
            180deg,
            #004B34,
            #006747,
            #008A5A
        ) !important;

    border-right:3px solid #D4AF37;
}

section[data-testid="stSidebar"] * {
    color:white !important;
    font-family:"Segoe UI",sans-serif !important;
}


/* =====================================================
   SIDEBAR NAVIGATION TITLE
===================================================== */

section[data-testid="stSidebar"] label {
    color:#FFD700 !important;
    font-size:18px !important;
    font-weight:700 !important;
}


/* =====================================================
   RADIO BUTTONS
===================================================== */

div[role="radiogroup"] label {
    background:rgba(255,255,255,.08);
    margin-bottom:8px;
    padding:10px;
    border-radius:10px;
    transition:.3s;
}

div[role="radiogroup"] label:hover {
    background:rgba(255,255,255,.20);
}


/* =====================================================
   SELECTED PAGE
===================================================== */

div[role="radiogroup"] label[data-selected="true"] {
    background:white !important;
    color:#006747 !important;
    font-weight:bold !important;
    border-left:5px solid #FFD700;
}


/* =====================================================
   KPI CARDS
===================================================== */

div[data-testid="metric-container"] {
    background:#FFFFFF !important;

    border-left:6px solid #006747 !important;

    border-radius:12px !important;

    padding:18px !important;

    box-shadow:
        0 3px 12px
        rgba(0,103,71,.10);

    border-top:1px solid #DDEBE4 !important;
    border-right:1px solid #DDEBE4 !important;
    border-bottom:1px solid #DDEBE4 !important;
}


/* =====================================================
   KPI LABEL
===================================================== */

[data-testid="stMetricLabel"] {
    color:#006747 !important;
    font-weight:bold !important;
}


/* =====================================================
   KPI VALUE
===================================================== */

[data-testid="stMetricValue"] {
    color:#222222 !important;
    font-size:34px !important;
    font-weight:700 !important;
}


/* =====================================================
   KPI DELTA
===================================================== */

[data-testid="stMetricDelta"] {
    color:#006747 !important;
}


/* =====================================================
   BUTTONS
===================================================== */

.stButton > button {
    background:#006747 !important;

    color:white !important;

    border:none !important;

    border-radius:8px !important;

    font-weight:bold !important;

    box-shadow:
        0 2px 6px
        rgba(0,103,71,.12);
}

.stButton > button:hover {
    background:#008A5A !important;

    box-shadow:
        0 4px 10px
        rgba(0,103,71,.18);
}


/* =====================================================
   TEXT INPUT
===================================================== */

.stTextInput input {
    background:#FFFFFF !important;

    color:#222222 !important;

    border:2px solid #006747 !important;

    border-radius:8px !important;

    box-shadow:
        0 1px 5px
        rgba(0,103,71,.05);
}


/* =====================================================
   SELECT BOX
===================================================== */

div[data-baseweb="select"] > div {
    background:#FFFFFF !important;

    color:#222222 !important;

    border:2px solid #006747 !important;

    border-radius:8px !important;

    box-shadow:
        0 1px 5px
        rgba(0,103,71,.05);
}


/* =====================================================
   DATAFRAME
===================================================== */

[data-testid="stDataFrame"] {
    background:#FFFFFF !important;

    border:2px solid #006747 !important;

    border-radius:10px !important;

    box-shadow:
        0 3px 12px
        rgba(0,103,71,.08);
}


/* =====================================================
   ALERT BOXES
===================================================== */

[data-testid="stAlert"] {
    border-radius:10px !important;
}

[data-testid="stAlert"] * {
    color:#222222 !important;
}


/* =====================================================
   MOBILE RESPONSIVE
===================================================== */

@media (max-width:768px) {

    .block-container {
        padding-top:1rem !important;
        padding-left:.7rem !important;
        padding-right:.7rem !important;
        max-width:100% !important;
    }

    h1,
    h2,
    h3,
    h4,
    h5,
    h6 {
        line-height:1.15 !important;
    }

    p {
        line-height:1.25 !important;
    }

    label,
    label p,
    [data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] p {
        font-size:13px !important;
        line-height:1.15 !important;
        margin-bottom:2px !important;
    }

    .stTextInput input {
        height:38px !important;
        font-size:13px !important;
    }

    div[data-baseweb="select"] > div {
        min-height:38px !important;
        font-size:13px !important;
    }

    div[data-testid="metric-container"] {
        padding:10px !important;
        border-left-width:4px !important;
        border-radius:9px !important;
    }

    [data-testid="stMetricLabel"] {
        font-size:11px !important;
        line-height:1.1 !important;
    }

    [data-testid="stMetricValue"] {
        font-size:24px !important;
        line-height:1.1 !important;
    }

    [data-testid="stMetricDelta"] {
        font-size:10px !important;
    }

    .stButton > button {
        min-height:36px !important;
        padding:5px 8px !important;
        font-size:12px !important;
        line-height:1.15 !important;
    }

    [data-testid="stDataFrame"] {
        width:100% !important;
        overflow-x:auto !important;
    }

    [data-testid="stAlert"] {
        padding:8px !important;
        font-size:12px !important;
    }
}


/* =====================================================
   VERY SMALL MOBILE
===================================================== */

@media (max-width:480px) {

    .block-container {
        padding-top:.7rem !important;
        padding-left:.45rem !important;
        padding-right:.45rem !important;
    }

    [data-testid="stMetricValue"] {
        font-size:21px !important;
    }

    [data-testid="stMetricLabel"] {
        font-size:10px !important;
    }

    .stButton > button {
        font-size:11px !important;
        padding:4px 6px !important;
    }

    .stTextInput input {
        font-size:12px !important;
    }

    div[data-baseweb="select"] > div {
        font-size:12px !important;
    }
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
# SIDEBAR - SMARTPAY PREMIUM GREEN THEME
# =====================================================

_sidebar_css = """
<style>

/* =====================================================
   SIDEBAR BACKGROUND - FIXED WIDTH
===================================================== */

section[data-testid="stSidebar"] {

    width:250px !important;
    min-width:250px !important;
    max-width:250px !important;
    flex:0 0 250px !important;

    background:
        linear-gradient(
            180deg,
            rgba(1,46,33,.78) 0%,
            rgba(0,75,52,.55) 40%,
            rgba(0,103,71,.42) 75%,
            rgba(0,60,42,.62) 100%
        ),
        url("data:image/jpeg;base64,SIDEBAR_BG_PLACEHOLDER") !important;

    background-size: cover !important;
    background-position: 40% center !important;
    background-repeat: no-repeat !important;
    background-blend-mode: normal !important;

    overflow:hidden !important;

    border-right:
        1px solid
        rgba(255,255,255,.08) !important;
}


/* =====================================================
   SIDEBAR MAIN CONTENT - FIXED WIDTH
===================================================== */

section[data-testid="stSidebar"]
> div:first-child {

    width:250px !important;
    min-width:250px !important;
    max-width:250px !important;

    height:100vh !important;

    overflow:hidden !important;

    padding-top:0 !important;
}

section[data-testid="stSidebar"]
[data-testid="stVerticalBlock"] {

    width:100% !important;
    max-width:100% !important;

    padding-top:0 !important;
}


/* =====================================================
   LOGO - TOP POSITION (WHITE ROUNDED CARD)
===================================================== */

section[data-testid="stSidebar"]
[data-testid="stImage"] {

    background:white !important;
    border-radius:12px !important;
    padding:10px 12px !important;
    margin:12px auto 6px auto !important;
    width:94% !important;
    box-shadow:0 4px 12px rgba(0,0,0,.18) !important;
    box-sizing:border-box !important;
}

section[data-testid="stSidebar"] img {

    max-height:80px !important;

    width:auto !important;

    object-fit:contain !important;

    display:block !important;

    margin:0 auto !important;

    filter:none !important;
}


/* =====================================================
   SIDEBAR HEADINGS
===================================================== */

section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p {

    color:white !important;
}


/* =====================================================
   SMARTPAY BRAND TITLE
===================================================== */

.sidebar-brand {

    text-align:center !important;

    margin-top:-4px !important;

    margin-bottom:4px !important;
}

.sidebar-title {

    color:white !important;

    font-family:
        "Segoe UI",
        Arial,
        sans-serif !important;

    font-size:18px !important;

    font-weight:800 !important;

    letter-spacing:.4px !important;

    line-height:1.1 !important;
}

.sidebar-subtitle {

    color:#D1FAE5 !important;

    font-family:
        "Segoe UI",
        Arial,
        sans-serif !important;

    font-size:9px !important;

    font-weight:500 !important;

    letter-spacing:.8px !important;

    margin-top:2px !important;
}

.sidebar-brand-line {

    width:38px !important;

    height:2px !important;

    background:#FFD700 !important;

    border-radius:10px !important;

    margin:4px auto 2px auto !important;
}


/* =====================================================
   NAVIGATION GROUP
===================================================== */

section[data-testid="stSidebar"]
div[role="radiogroup"] {

    width:100% !important;

    max-width:100% !important;

    margin-top:3px !important;

    margin-bottom:0 !important;

    padding:4px !important;

    background:
        rgba(0,0,0,.10) !important;

    border:
        1px solid
        rgba(255,255,255,.08) !important;

    border-radius:13px !important;

    box-sizing:border-box !important;
}


/* =====================================================
   NAVIGATION BUTTONS
===================================================== */

section[data-testid="stSidebar"]
div[role="radiogroup"] label {

    width:100% !important;

    max-width:100% !important;

    background:transparent !important;

    border:
        1px solid
        transparent !important;

    border-radius:999px !important;

    padding:6px 14px !important;

    margin-bottom:3px !important;

    min-height:28px !important;

    color:#F4FFFA !important;

    font-family:
        "Segoe UI",
        Arial,
        sans-serif !important;

    font-size:12.5px !important;

    font-weight:600 !important;

    line-height:1.1 !important;

    transition:
        background .2s ease,
        transform .2s ease,
        border .2s ease !important;

    white-space:nowrap !important;

    overflow:hidden !important;

    text-overflow:ellipsis !important;

    box-sizing:border-box !important;
}


/* =====================================================
   LAST BUTTON
===================================================== */

section[data-testid="stSidebar"]
div[role="radiogroup"]
label:last-child {

    margin-bottom:0 !important;
}


/* =====================================================
   HOVER
===================================================== */

section[data-testid="stSidebar"]
div[role="radiogroup"]
label:hover {

    background:
        linear-gradient(
            90deg,
            rgba(32,164,100,.30),
            rgba(32,164,100,.12)
        ) !important;

    border:
        1px solid
        rgba(255,255,255,.10) !important;

    transform:translateX(2px) !important;
}


/* =====================================================
   SELECTED PAGE
===================================================== */

section[data-testid="stSidebar"]
div[role="radiogroup"]
label:has(input:checked) {

    background:
        linear-gradient(
            90deg,
            #0FA968,
            #008A5A
        ) !important;

    color:white !important;

    border:
        1.5px solid
        #FFD700 !important;

    box-shadow:
        0 4px 12px
        rgba(0,0,0,.18) !important;

    font-weight:700 !important;
}




/* =====================================================
   HIDE RADIO CIRCLE
===================================================== */

section[data-testid="stSidebar"]
div[role="radiogroup"] input {

    display:none !important;
}


/* =====================================================
   SIDEBAR DIVIDERS
===================================================== */

section[data-testid="stSidebar"] hr {

    border:none !important;

    border-top:
        1px solid
        rgba(255,255,255,.12) !important;

    margin-top:4px !important;

    margin-bottom:4px !important;
}


/* =====================================================
   CREATED BY
===================================================== */

.sidebar-created-by {

    text-align:center !important;

    color:#D1FAE5 !important;

    font-family:
        "Segoe UI",
        Arial,
        sans-serif !important;

    font-size:11px !important;

    font-weight:600 !important;

    letter-spacing:.2px !important;

    margin-top:6px !important;

    margin-bottom:3px !important;

    white-space:nowrap !important;
}


/* =====================================================
   MOBILE - FIXED 215px
===================================================== */

@media (max-width:768px) {

    section[data-testid="stSidebar"] {

        width:215px !important;
        min-width:215px !important;
        max-width:215px !important;
        flex:0 0 215px !important;

        overflow:hidden !important;
    }

    section[data-testid="stSidebar"]
    > div:first-child {

        width:215px !important;
        min-width:215px !important;
        max-width:215px !important;

        height:100vh !important;

        overflow-y:hidden !important;
        overflow-x:hidden !important;

        padding-top:0 !important;
    }

    section[data-testid="stSidebar"] img {

        max-height:90px !important;

        margin-top:-15px !important;

        margin-bottom:-5px !important;
    }

    section[data-testid="stSidebar"]
    div[role="radiogroup"] label {

        padding:5px 8px !important;

        margin-bottom:2px !important;

        min-height:28px !important;

        font-size:12px !important;
    }

    .sidebar-title {
        font-size:20px !important;
    }

    .sidebar-subtitle {
        font-size:11px !important;
    }

    .sidebar-created-by {

        font-size:10px !important;

        margin-top:6px !important;
    }
}


/* =====================================================
   VERY SMALL MOBILE - FIXED 200px
===================================================== */

@media (max-width:480px) {

    section[data-testid="stSidebar"] {

        width:200px !important;
        min-width:200px !important;
        max-width:200px !important;
        flex:0 0 200px !important;
    }

    section[data-testid="stSidebar"]
    > div:first-child {

        width:200px !important;
        min-width:200px !important;
        max-width:200px !important;
    }

    section[data-testid="stSidebar"] img {

        max-height:75px !important;

        margin-top:-12px !important;

        margin-bottom:-4px !important;
    }

    section[data-testid="stSidebar"]
    div[role="radiogroup"] label {

        padding:4px 7px !important;

        margin-bottom:2px !important;

        min-height:26px !important;

        font-size:11px !important;
    }

    .sidebar-title {
        font-size:18px !important;
    }

    .sidebar-subtitle {
        font-size:10px !important;
    }

    .sidebar-created-by {
        font-size:9px !important;
    }
}

</style>
""".replace("SIDEBAR_BG_PLACEHOLDER", SIDEBAR_BG_B64)

st.markdown(_sidebar_css, unsafe_allow_html=True)

# =====================================================
# LOGO
# =====================================================

st.sidebar.image(
    "smartpay_logo/smartpay_logo.png",
    use_container_width=True
)


# =====================================================
# SMARTPAY TITLE
# =====================================================

st.sidebar.markdown(
"""
<div class="sidebar-brand">
<div class="sidebar-title">🏦 SmartPay</div>
<div class="sidebar-subtitle">PROJECT DASHBOARD</div>
<div class="sidebar-brand-line"></div>
</div>
""",
unsafe_allow_html=True
)


# =====================================================
# NAVIGATION
# =====================================================

if "navigate_to" in st.session_state:

    st.session_state["page_navigation"] = (
        st.session_state["navigate_to"]
    )

    del st.session_state["navigate_to"]


page = st.sidebar.radio(
    "Menu",
   [
        "Dashboard",
        "Projects",
        "Daily Issues",
        "Analytics",
        "Project Timeline",
        "BAU Monitoring",
        "IS Issues",
        "CRPL",
        "PAYSYS",
        "Export"
    ],
    label_visibility="collapsed",
    key="page_navigation"
)


# =====================================================
# SIDEBAR FOOTER
# =====================================================

st.sidebar.markdown("---")

st.sidebar.markdown(
"""
<div class="sidebar-created-by">
✦ Created by Muhammad Saad Asif ✦
</div>
""",
unsafe_allow_html=True
)

# ==========================================
# PAGE CHANGE DETECTION
# ==========================================

if "previous_page" not in st.session_state:
    st.session_state.previous_page = page
    st.session_state.page_changed = False

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
        st.session_state.pop(key, None)

    # Mark page changed
    st.session_state.previous_page = page
    st.session_state.page_changed = True

else:
    st.session_state.page_changed = False
# =====================================================
# FORCE MAIN PAGE TO TOP AFTER NAVIGATION
# =====================================================

def scroll_main_to_top(nonce: str, duration_ms: int = 1200):
    components.html(
        f"""
        <script>
        (() => {{
            // nonce: {nonce}  <-- forces a fresh iframe on every page change
            const doc = window.parent.document;
            const START = Date.now();
            let userTookOver = false;

            function targets() {{
                const sels = [
                    'section.stMain',
                    '[data-testid="stMain"]',
                    'section.main',
                    '[data-testid="stAppViewContainer"] > section'
                ];
                const found = new Set();
                sels.forEach(s => doc.querySelectorAll(s).forEach(el => found.add(el)));
                return [...found];
            }}

            function toTop() {{
                targets().forEach(el => {{ el.scrollTop = 0; }});
                doc.documentElement.scrollTop = 0;
                doc.body.scrollTop = 0;
            }}

            // Bail out the moment the user scrolls, so we never fight them
            ['wheel', 'touchstart', 'keydown'].forEach(ev =>
                doc.addEventListener(ev, () => {{ userTookOver = true; }},
                                     {{ once: true, passive: true }})
            );

            function clamp() {{
                if (userTookOver) return;
                toTop();
                if (Date.now() - START < {duration_ms}) {{
                    window.requestAnimationFrame(clamp);
                }}
            }}

            toTop();
            clamp();
        }})();
        </script>
        """,
        height=0,
    )


if st.session_state.get("page_changed", False):
    st.session_state["scroll_nonce"] = st.session_state.get("scroll_nonce", 0) + 1
    scroll_main_to_top(f"{page}-{st.session_state['scroll_nonce']}")

def render_header(eyebrow, title, subtitle, badge, gap="0.9rem"):
    _now = datetime.now(
        ZoneInfo("Asia/Karachi")
    ).strftime("%b %d, %Y  •  %I:%M %p")

    st.markdown(f"""
<style>.block-container [data-testid="stVerticalBlock"]{{gap:{gap} !important;}}.block-container [data-testid="stMarkdownContainer"] h2{{padding-top:0 !important;padding-bottom:0 !important;margin-top:12px !important;margin-bottom:8px !important;}}.block-container [data-testid="stMarkdownContainer"] h3{{padding-top:0 !important;padding-bottom:0 !important;}}</style>
<div style="position:relative;overflow:hidden;display:flex;align-items:center;background:linear-gradient(90deg, rgba(1,30,22,.72) 0%, rgba(1,30,22,.55) 35%, rgba(1,30,22,.12) 65%, rgba(1,30,22,0) 100%), url('data:image/jpeg;base64,{HEADER_BG_B64}');background-size:cover;background-position:center 55%;background-repeat:no-repeat;border-radius:16px;padding:16px 22px;border:1px solid rgba(255,255,255,.15);box-shadow:0 7px 18px rgba(0,0,0,.12);box-sizing:border-box;margin-bottom:10px;">
<div style="position:relative;width:100%;display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:14px;">
<div>
<div style="color:#FFFFFF;font-size:12px;font-weight:800;letter-spacing:1.8px;margin-bottom:5px;text-shadow:0 1px 4px rgba(0,0,0,.85);">{eyebrow}</div>
<div style="color:white;font-size:29px;font-weight:800;line-height:1.15;margin:0;text-shadow:0 2px 8px rgba(0,0,0,.35);">{title}</div>
<div style="margin-top:5px;color:#FFFFFF;font-size:14px;font-weight:600;text-shadow:0 1px 4px rgba(0,0,0,.85);">{subtitle}</div>
<div style="margin-top:10px;display:inline-block;background:rgba(0,0,0,.45);border:1px solid rgba(255,215,0,.7);color:#FFE066;padding:4px 12px;border-radius:20px;font-size:10px;font-weight:800;letter-spacing:.4px;text-shadow:0 1px 3px rgba(0,0,0,.6);">● {badge}</div>
</div>
<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
<div style="background:rgba(0,0,0,.35);backdrop-filter:blur(6px);border:1px solid rgba(255,255,255,.35);border-radius:12px;padding:8px 14px;display:flex;align-items:center;gap:8px;">
<span style="font-size:14px;">📅</span>
<div style="color:#FFFFFF;font-size:12px;font-weight:700;white-space:nowrap;">{_now}</div>
</div>
<div style="background:rgba(0,0,0,.35);backdrop-filter:blur(6px);border:1px solid rgba(255,255,255,.35);border-radius:50%;width:34px;height:34px;display:flex;align-items:center;justify-content:center;position:relative;">
<span style="font-size:15px;">🔔</span>
<span style="position:absolute;top:5px;right:6px;width:7px;height:7px;background:#FF4D4D;border-radius:50%;border:1.5px solid #013D2B;"></span>
</div>
<div style="background:linear-gradient(135deg,#FFD700,#F5B300);color:#013D2B;font-weight:800;font-size:11px;letter-spacing:.3px;width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;border:1px solid rgba(255,255,255,.4);">DBG</div>
</div>
</div>
</div>
""", unsafe_allow_html=True)

# =====================================================
# DASHBOARD
# =====================================================

if page == "Dashboard":

    # ==========================================
    # TITLE - PREMIUM HEADER
    # ==========================================

    render_header(
        "NATIONAL BANK OF PAKISTAN",
        "💳 SmartPay Project Dashboard",
        "Digital Banking Group &nbsp;|&nbsp; Powering Digital Payments",
        "LIVE DASHBOARD",
        gap="0.55rem"
    )

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
    # KPI NAVIGATION FUNCTION
    # =====================================================

    def navigate_to_projects(status_filter):

        # Clear any previous team navigation
        st.session_state.pop(
            "selected_team_project",
            None
        )

        # Set fresh project filter
        st.session_state["project_status_filter"] = status_filter

        # Set destination
        st.session_state["navigate_to"] = "Projects"

        # Tell scroll system that page changed
        st.session_state["page_changed"] = True

        # Navigate
        st.rerun()


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
    FOCUS / CLICKED STATE
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

                navigate_to_projects("All")


        # -----------------------------------------
        # SCOPING
        # -----------------------------------------

        with c2:

            if st.button(
                f"SCOPING\n{scoping_projects}",
                key="dashboard_scoping",
                use_container_width=True
            ):

                navigate_to_projects("SCOPING")


        # -----------------------------------------
        # UAT
        # -----------------------------------------

        with c3:

            if st.button(
                f"UAT\n{uat_projects}",
                key="dashboard_uat",
                use_container_width=True
            ):

                navigate_to_projects("UAT")


        # -----------------------------------------
        # IS REVIEW
        # -----------------------------------------

        with c4:

            if st.button(
                f"IS REVIEW\n{review_projects}",
                key="dashboard_review",
                use_container_width=True
            ):

                navigate_to_projects("IS REVIEW")


        # -----------------------------------------
        # CMC
        # -----------------------------------------

        with c5:

            if st.button(
                f"CMC\n{cmc_projects}",
                key="dashboard_cmc",
                use_container_width=True
            ):

                navigate_to_projects("CMC")


        # -----------------------------------------
        # LIVE
        # -----------------------------------------

        with c6:

            if st.button(
                f"LIVE\n{live_projects}",
                key="dashboard_live",
                use_container_width=True
            ):

                navigate_to_projects("LIVE")


        # -----------------------------------------
        # BAU
        # -----------------------------------------

        with c7:

            if st.button(
                f"BAU\n{bau_projects}",
                key="dashboard_bau",
                use_container_width=True
            ):

                navigate_to_projects("BAU")
    

    # =====================================================

    # SMART SEARCH - GLOBAL VIP SEARCH
    # =====================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:24px;
        font-weight:700;
        margin-top:0;
        margin-bottom:0;
        font-family:Segoe UI,Arial,sans-serif;">
        🔍 Smart Search
    </div>
    """)


    # =====================================================
    # GLOBAL SEARCH INPUT
    # =====================================================

    global_search = st.text_input(
        "Search",
        placeholder="🔍 Search Projects, CRPL, PAYSYS, person name, issue, project name...",
        label_visibility="collapsed",
        key="global_smart_search"
    )


    # =====================================================
    # LOAD SEARCH DATA
    # =====================================================

    smartpay_search_df = df.copy()

    crpl_search_df = pd.DataFrame()
    paysys_search_df = pd.DataFrame()


    # =====================================================
    # LOAD CRPL
    # =====================================================

    try:

        crpl_file_search = os.path.join(
            os.path.dirname(__file__),
            "data",
            "CRPL_All_Data_Exactly_20.xlsx"
        )

        crpl_raw_search = pd.read_excel(
            crpl_file_search,
            sheet_name="CRPL",
            header=None
        )

        crpl_header_row = None

        for i in range(
            min(20, len(crpl_raw_search))
        ):

            row = " ".join(
                str(x).lower()
                for x in crpl_raw_search.iloc[i]
                if pd.notna(x)
            )

            if (
                "crf no" in row
                and "crf name" in row
            ):

                crpl_header_row = i
                break


        if crpl_header_row is not None:

            crpl_search_df = pd.read_excel(
                crpl_file_search,
                sheet_name="CRPL",
                header=crpl_header_row
            )

            crpl_search_df.columns = [
                str(x).strip()
                for x in crpl_search_df.columns
            ]

            crpl_search_df = (
                crpl_search_df
                .dropna(how="all")
                .reset_index(drop=True)
            )

    except Exception:

        crpl_search_df = pd.DataFrame()


    # =====================================================
    # LOAD PAYSYS
    # =====================================================

    try:

        paysys_file_search = os.path.join(
            os.path.dirname(__file__),
            "data",
            "PAYSYS_Weekly_Project_Update_04-Sep-2026.xlsx"
        )

        paysys_excel_search = pd.ExcelFile(
            paysys_file_search
        )

        paysys_sheet_search = (
            paysys_excel_search.sheet_names[0]
        )

        paysys_raw_search = pd.read_excel(
            paysys_file_search,
            sheet_name=paysys_sheet_search,
            header=None
        )

        paysys_header_row = None

        for i in range(
            min(20, len(paysys_raw_search))
        ):

            row = " ".join(
                str(x).lower()
                for x in paysys_raw_search.iloc[i]
                if pd.notna(x)
            )

            if (
                "uat/live" in row
                and "paysys response" in row
            ):

                paysys_header_row = i
                break


        if paysys_header_row is not None:

            paysys_search_df = pd.read_excel(
                paysys_file_search,
                sheet_name=paysys_sheet_search,
                header=paysys_header_row
            )

            paysys_search_df.columns = [
                str(x).strip()
                for x in paysys_search_df.columns
            ]

            paysys_search_df = (
                paysys_search_df
                .dropna(how="all")
                .reset_index(drop=True)
            )

    except Exception:

        paysys_search_df = pd.DataFrame()


    # =====================================================
    # SEARCH FUNCTION
    # =====================================================

    def global_search_dataframe(
        dataframe,
        query,
        source_keywords
    ):

        if dataframe.empty:
            return dataframe.copy()

        query_clean = str(
            query
        ).strip().lower()

        # ---------------------------------------------
        # SOURCE SEARCH
        # ---------------------------------------------

        if query_clean in source_keywords:

            return dataframe.copy()


        # ---------------------------------------------
        # SEARCH ALL COLUMNS
        # ---------------------------------------------

        search_mask = pd.Series(
            False,
            index=dataframe.index
        )

        for col in dataframe.columns:

            search_mask = (
                search_mask
                |
                dataframe[col]
                .astype(str)
                .str.contains(
                    query,
                    case=False,
                    na=False,
                    regex=False
                )
            )

        return dataframe[
            search_mask
        ].copy()


    # =====================================================
    # VIP TABLE FUNCTION
    # =====================================================

    def show_global_search_table(
        dataframe,
        source
    ):

        if dataframe.empty:
            return


        display_df = dataframe.copy()


        # =================================================
        # SOURCE-SPECIFIC STYLING
        # =================================================

        if source == "Projects":

            if "Mandate" in display_df.columns:

                display_df["Mandate"] = (
                    display_df["Mandate"]
                    .apply(
                        lambda x:
                        f'<span class="global-project-name">'
                        f'{x}</span>'
                    )
                )

            if "Status" in display_df.columns:

                def project_status_badge(
                    value
                ):

                    value = str(value).strip()

                    upper = value.upper()

                    if upper == "LIVE":

                        css = "global-stage-live"

                    elif upper == "UAT":

                        css = "global-stage-uat"

                    elif upper == "CMC":

                        css = "global-stage-cmc"

                    elif upper == "IS REVIEW":

                        css = "global-stage-review"

                    elif upper in [
                        "SCOPING",
                        "UNDER SCOPING"
                    ]:

                        css = "global-stage-scoping"

                    elif upper == "BAU":

                        css = "global-stage-bau"

                    else:

                        css = "global-stage-default"

                    return (
                        f'<span class="global-stage-badge '
                        f'{css}">{value}</span>'
                    )

                display_df["Status"] = (
                    display_df["Status"]
                    .apply(project_status_badge)
                )


        elif source == "CRPL":

            if "CRF Name" in display_df.columns:

                display_df["CRF Name"] = (
                    display_df["CRF Name"]
                    .apply(
                        lambda x:
                        f'<span class="global-project-name">'
                        f'{x}</span>'
                    )
                )

            if "Stage" in display_df.columns:

                def crpl_stage_badge(
                    value
                ):

                    value = str(value).strip()

                    upper = value.upper()

                    if upper == "HOLD":

                        css = "global-stage-hold"

                    elif upper == "WIP":

                        css = "global-stage-wip"

                    elif upper == "UAT":

                        css = "global-stage-uat"

                    elif upper in [
                        "LIVE",
                        "PRODUCTION",
                        "LIVE / PRODUCTION",
                        "LIVE/PRODUCTION"
                    ]:

                        css = "global-stage-live"

                    else:

                        css = "global-stage-default"

                    return (
                        f'<span class="global-stage-badge '
                        f'{css}">{value}</span>'
                    )

                display_df["Stage"] = (
                    display_df["Stage"]
                    .apply(crpl_stage_badge)
                )


        elif source == "PAYSYS":

            if "Project / Issue" in display_df.columns:

                display_df["Project / Issue"] = (
                    display_df["Project / Issue"]
                    .apply(
                        lambda x:
                        f'<span class="global-project-name">'
                        f'{x}</span>'
                    )
                )

            if "UAT/Live" in display_df.columns:

                def paysys_stage_badge(
                    value
                ):

                    value = str(value).strip()

                    upper = value.upper()

                    if upper == "UAT":

                        css = "global-stage-uat"

                    elif upper in [
                        "LIVE",
                        "PRODUCTION",
                        "LIVE / PRODUCTION",
                        "LIVE/PRODUCTION"
                    ]:

                        css = "global-stage-live"

                    else:

                        css = "global-stage-default"

                    return (
                        f'<span class="global-stage-badge '
                        f'{css}">{value}</span>'
                    )

                display_df["UAT/Live"] = (
                    display_df["UAT/Live"]
                    .apply(paysys_stage_badge)
                )


        # =================================================
        # HTML TABLE
        # =================================================

        table_html = display_df.to_html(
            index=False,
            escape=False,
            classes="global-search-table"
        )


        # =================================================
        # SOURCE TITLE
        # =================================================

        source_icon = {

            "Projects": "📁",
            "CRPL": "🏦",
            "PAYSYS": "💳"

        }.get(
            source,
            "📋"
        )


        st.html(
            f"""
            <div class="global-search-source-title">
                {source_icon} {source}
                <span>
                    {len(dataframe)} result(s)
                </span>
            </div>

            <div class="global-search-table-wrapper">
                {table_html}
            </div>
            """
        )


    # =====================================================
    # GLOBAL SEARCH CSS
    # =====================================================

    st.html("""
    <style>

    .global-search-source-title {
        color:#006747;
        font-size:18px;
        font-weight:700;
        margin-top:10px;
        margin-bottom:5px;
        font-family:Segoe UI,Arial,sans-serif;
    }

    .global-search-source-title span {
        color:#6B7280;
        font-size:11px;
        font-weight:600;
        margin-left:7px;
    }


    .global-search-table-wrapper {
        background:#FFFFFF;
        border:1px solid #DDE5E1;
        border-radius:16px;
        padding:6px;
        box-shadow:0 6px 20px rgba(0,103,71,0.08);
        overflow:hidden;
        margin-bottom:10px;
    }


    .global-search-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:13px;
        table-layout:fixed;
    }


    .global-search-table thead th {
        background:#006747;
        color:#FFFFFF;
        font-weight:700;
        padding:11px 10px;
        text-align:left;
    }


    .global-search-table thead th:first-child {
        border-top-left-radius:10px;
    }


    .global-search-table thead th:last-child {
        border-top-right-radius:10px;
    }


    .global-search-table tbody td {
        padding:10px;
        color:#1F2937;
        border-bottom:1px solid #E5E7EB;
        background:#FFFFFF;

        white-space:normal !important;
        word-wrap:break-word !important;
        overflow-wrap:anywhere !important;
        vertical-align:top;
        line-height:1.45;
    }


    .global-search-table tbody tr:nth-child(even) td {
        background:#F8FAFC;
    }


    .global-search-table tbody tr:hover td {
        background:#ECFDF5;
    }


    .global-project-name {
        color:#006747 !important;
        font-weight:700;
    }


    .global-stage-badge {
        display:inline-block;
        padding:4px 9px;
        border-radius:16px;
        font-size:10px;
        font-weight:800;
        white-space:nowrap;
    }


    .global-stage-live {
        background:#DCFCE7;
        color:#166534;
    }


    .global-stage-uat {
        background:#FEF3C7;
        color:#92400E;
    }


    .global-stage-wip {
        background:#DBEAFE;
        color:#1E40AF;
    }


    .global-stage-hold {
        background:#FEE2E2;
        color:#991B1B;
    }


    .global-stage-cmc {
        background:#EDE9FE;
        color:#5B21B6;
    }


    .global-stage-review {
        background:#E0F2FE;
        color:#075985;
    }


    .global-stage-scoping {
        background:#F3F4F6;
        color:#374151;
    }


    .global-stage-bau {
        background:#E0F2FE;
        color:#075985;
    }


    .global-stage-default {
        background:#F3F4F6;
        color:#374151;
    }

    </style>
    """)


    # =====================================================
    # PERFORM GLOBAL SEARCH
    # =====================================================

    if global_search:

        projects_result = global_search_dataframe(
            smartpay_search_df,
            global_search,
            {
                "project",
                "projects",
                "smartpay",
                "smartpay projects"
            }
        )


        crpl_result = global_search_dataframe(
            crpl_search_df,
            global_search,
            {
                "crpl"
            }
        )


        paysys_result = global_search_dataframe(
            paysys_search_df,
            global_search,
            {
                "paysys"
            }
        )


        total_results = (
            len(projects_result)
            + len(crpl_result)
            + len(paysys_result)
        )


        # =================================================
        # SEARCH SUMMARY
        # =================================================

        if total_results == 0:

            st.html("""
            <div style="
                background:#FEF2F2;
                border:1px solid #FECACA;
                border-left:4px solid #DC2626;
                border-radius:10px;
                padding:9px 12px;
                margin-top:7px;
                color:#991B1B;
                font-size:13px;
                font-weight:600;
                font-family:Segoe UI,Arial,sans-serif;">
                ❌ No matching records found.
            </div>
            """)

        else:

            st.html(
                f"""
                <div style="
                    background:#ECFDF5;
                    border:1px solid #A7F3D0;
                    border-left:4px solid #006747;
                    border-radius:10px;
                    padding:9px 12px;
                    margin-top:7px;
                    margin-bottom:6px;
                    color:#006747;
                    font-size:13px;
                    font-weight:700;
                    font-family:Segoe UI,Arial,sans-serif;">
                    ✅ {total_results} matching record(s) found
                </div>
                """
            )


            # =============================================
            # PROJECTS RESULTS
            # =============================================

            if not projects_result.empty:

                show_global_search_table(
                    projects_result,
                    "Projects"
                )


            # =============================================
            # CRPL RESULTS
            # =============================================

            if not crpl_result.empty:

                show_global_search_table(
                    crpl_result,
                    "CRPL"
                )


            # =============================================
            # PAYSYS RESULTS
            # =============================================

            if not paysys_result.empty:

                show_global_search_table(
                    paysys_result,
                    "PAYSYS"
                )


    # =====================================================
    # TEAM OVERVIEW - DYNAMIC CLICKABLE KPI CARDS
    # =====================================================

    st.html("""
    <div style="
        color:#006747;
        font-size:26px;
        font-weight:700;
        margin-top:0;
        margin-bottom:0;
        font-family:Segoe UI,Arial,sans-serif;">
        👥 Team Overview
    </div>
    """)


    # =====================================================
    # DYNAMIC TEAM DATA
    # =====================================================

    allocation = (
        df["Allocation"]
        .dropna()
        .astype(str)
        .str.strip()
    )

    allocation = allocation[
        allocation != ""
    ].value_counts().reset_index()

    allocation.columns = [
        "Allocation",
        "Projects"
    ]


    # =====================================================
    # TEAM CARD CSS - COMPACT SINGLE LINE
    # =====================================================

    st.html("""
    <style>

    .st-key-team_overview_kpis
    div[data-testid="stButton"] {
        width:100%;
    }

    .st-key-team_overview_kpis
    div[data-testid="stButton"] button {

        width:100% !important;
        min-height:58px !important;
        height:58px !important;

        background:linear-gradient(
            135deg,
            #013D2B,
            #006747,
            #008A5A
        ) !important;

        border:1px solid rgba(255,255,255,.15) !important;
        border-radius:10px !important;

        padding:5px 4px !important;

        box-shadow:
            0 3px 8px rgba(0,103,71,.15) !important;

        color:white !important;

        font-family:"Segoe UI",Arial,sans-serif !important;

        font-size:11px !important;
        font-weight:700 !important;

        line-height:1.15 !important;

        white-space:pre-line !important;

        text-align:center !important;

        transition:all .2s ease !important;
    }

    .st-key-team_overview_kpis
    div[data-testid="stButton"] button:hover {

        background:linear-gradient(
            135deg,
            #006747,
            #008A5A
        ) !important;

        transform:translateY(-1px);

        box-shadow:
            0 5px 12px rgba(0,103,71,.22) !important;
    }

    .st-key-team_overview_kpis
    div[data-testid="stButton"] button:focus {

        outline:none !important;

        background:#006747 !important;

        border:2px solid #FFD700 !important;

        box-shadow:
            0 0 0 2px rgba(255,215,0,.12) !important;

        color:white !important;
    }

    .st-key-team_overview_kpis
    div[data-testid="stButton"] button p {

        font-size:11px !important;
        font-weight:700 !important;
        line-height:1.2 !important;

        color:white !important;

        white-space:pre-line !important;
        text-align:center !important;
    }

    .st-key-team_overview_kpis
    div[data-testid="stHorizontalBlock"] {

        gap:4px !important;
    }

    .st-key-team_overview_kpis
    div[data-testid="stVerticalBlock"] {

        gap:0 !important;
    }

    </style>
    """)


    # =====================================================
    # DYNAMIC TEAM KPI CARDS
    # =====================================================

    with st.container(key="team_overview_kpis"):

        cols = st.columns(
            len(allocation),
            gap="small"
        )

        for i, row in allocation.iterrows():

            member = str(
                row["Allocation"]
            ).strip()

            project_count = int(
                row["Projects"]
            )

            with cols[i]:

                if st.button(
                    f"👤 {member}\n{project_count} Projects",
                    key=f"team_member_{i}",
                    use_container_width=True
                ):

                    # -----------------------------------------
                    # CLEAR OLD NAVIGATION STATES
                    # -----------------------------------------

                    st.session_state.pop(
                        "navigate_to",
                        None
                    )

                    st.session_state.pop(
                        "project_status_filter",
                        None
                    )

                    st.session_state.pop(
                        "selected_team_project",
                        None
                    )


                    # -----------------------------------------
                    # SET FRESH TEAM FILTER
                    # -----------------------------------------

                    st.session_state[
                        "selected_team_project"
                    ] = member


                    # -----------------------------------------
                    # SET DESTINATION
                    # -----------------------------------------

                    st.session_state[
                        "navigate_to"
                    ] = "Projects"


                    # -----------------------------------------
                    # MARK PAGE CHANGE
                    # -----------------------------------------

                    st.session_state[
                        "page_changed"
                    ] = True


                    # -----------------------------------------
                    # NAVIGATE
                    # -----------------------------------------

                    st.rerun()


    # =====================================================
    # TEAM SUMMARY - VIP
    # =====================================================

    st.html("""
    <div style="color:#006747;font-size:26px;font-weight:700;margin-top:0;margin-bottom:0;font-family:Segoe UI,Arial,sans-serif;">
    📋 Team Summary
    </div>
    """)

    # =====================================================
    # CREATE SUMMARY
    # =====================================================

    summary = (
        df.groupby("Allocation")
        .agg(
            Total=("Mandate", "count"),

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

            UAT=("Status", lambda x:
                (
                    x.astype(str)
                    .str.strip()
                    .str.upper()
                    == "UAT"
                ).sum()
            ),

            IS_Review=("Status", lambda x:
                (
                    x.astype(str)
                    .str.strip()
                    .str.upper()
                    == "IS REVIEW"
                ).sum()
            ),

            CMC=("Status", lambda x:
                (
                    x.astype(str)
                    .str.strip()
                    .str.upper()
                    == "CMC"
                ).sum()
            ),

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
    # COMPACT VIP TABLE CSS
    # =====================================================

    st.markdown("""
    <style>

    .vip-summary-wrapper {
        background:#FFFFFF;
        border:1px solid #DDE5E1;
        border-radius:12px;
        padding:3px;
        box-shadow:0 5px 15px rgba(0,103,71,0.07);
        overflow:hidden;
    }

    .vip-summary-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:12px;
    }

    .vip-summary-table thead th {
        background:linear-gradient(135deg,#013D2B,#006747,#008A5A);
        color:#FFFFFF;
        padding:9px 10px;
        font-weight:700;
        text-align:left;
        white-space:nowrap;
    }

    .vip-summary-table thead th:first-child {
        border-top-left-radius:9px;
    }

    .vip-summary-table thead th:last-child {
        border-top-right-radius:9px;
    }

    .vip-summary-table tbody td {
        padding:8px 10px;
        color:#1F2937;
        border-bottom:1px solid #E5E7EB;
        background:#FFFFFF;
        white-space:nowrap;
    }

    .vip-summary-table tbody tr:nth-child(even) td {
        background:#F8FAFC;
    }

    .vip-summary-table tbody tr:hover td {
        background:#ECFDF5;
    }

    .team-name {
        color:#006747 !important;
        font-weight:800;
    }

    .total-count {
        background:#F0FDF4;
        color:#006747 !important;
        font-weight:800;
        padding:3px 7px;
        border-radius:15px;
    }

    .scoping-count {
        background:#FFF7ED;
        color:#C2410C !important;
        font-weight:800;
        padding:3px 7px;
        border-radius:15px;
    }

    .uat-count {
        background:#FEF3C7;
        color:#92400E !important;
        font-weight:800;
        padding:3px 7px;
        border-radius:15px;
    }

    .review-count {
        background:#E0F2FE;
        color:#0369A1 !important;
        font-weight:800;
        padding:3px 7px;
        border-radius:15px;
    }

    .cmc-count {
        background:#F3E8FF;
        color:#7E22CE !important;
        font-weight:800;
        padding:3px 7px;
        border-radius:15px;
    }

    .live-count {
        background:#DCFCE7;
        color:#166534 !important;
        font-weight:800;
        padding:3px 7px;
        border-radius:15px;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # FORMAT SUMMARY
    # =====================================================

    summary_display = summary.copy()

    summary_display["Allocation"] = summary_display[
        "Allocation"
    ].apply(
        lambda x:
        f'<span class="team-name">👤 {x}</span>'
    )

    summary_display["Total"] = summary_display[
        "Total"
    ].apply(
        lambda x:
        f'<span class="total-count">{x}</span>'
    )

    summary_display["Scoping"] = summary_display[
        "Scoping"
    ].apply(
        lambda x:
        f'<span class="scoping-count">🟠 {x}</span>'
    )

    summary_display["UAT"] = summary_display[
        "UAT"
    ].apply(
        lambda x:
        f'<span class="uat-count">🟡 {x}</span>'
    )

    summary_display["IS Review"] = summary_display[
        "IS Review"
    ].apply(
        lambda x:
        f'<span class="review-count">🔷 {x}</span>'
    )

    summary_display["CMC"] = summary_display[
        "CMC"
    ].apply(
        lambda x:
        f'<span class="cmc-count">🟣 {x}</span>'
    )

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

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    
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

    render_header('SMARTPAY PROJECT MANAGEMENT', '📁 Project Portfolio', 'SmartPay Project Management System', 'PROJECT MANAGEMENT')


    # =====================================================
    # FILTERS - COMPACT
    # =====================================================

    st.markdown("""
    <h2 style="color:#006747;font-size:22px;font-weight:700;margin-top:2px;margin-bottom:5px;padding-top:0;">
    🎯 Filters
    </h2>
    """, unsafe_allow_html=True)


    # =====================================================
    # TEAM CARD → PROJECTS FILTER
    # =====================================================

    selected_team_project = st.session_state.pop(
        "selected_team_project",
        ""
    )

    if selected_team_project:
        st.session_state["project_filter_allocation"] = selected_team_project


    # =====================================================
    # FILTER COLUMNS
    # =====================================================

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

    left, right = st.columns([5, 1])


    with left:

        st.markdown(
            f"""
    <div style="background:#E8F5E9;padding:7px 12px;border-radius:9px;color:#006747;font-size:13px;font-weight:700;margin-top:3px;">
    📌 Showing <b>{len(filtered_df)}</b> Project(s)
    </div>
    """,
            unsafe_allow_html=True
        )


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


    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    # =====================================================
    # TEAM OVERVIEW - COMPACT CLICKABLE CARDS
    # =====================================================

    with st.container(key="team_overview_kpis"):

        allocation = (
            df["Allocation"]
            .dropna()
            .astype(str)
            .str.strip()
            .value_counts()
        )

        total_cards = len(allocation) + 1

        cols = st.columns(total_cards, gap="small")

        # -------------------------
        # ALL BUTTON
        # -------------------------
        with cols[0]:

            if st.button(
                f"📊 ALL\n{len(df)} Projects",
                key="team_all_btn",
                use_container_width=True
            ):

                st.session_state.pop(
                    "selected_team_project",
                    None
                )

                st.session_state.pop(
                    "project_filter_allocation",
                    None
                )

                st.session_state.pop(
                    "project_status_filter",
                    None
                )

                st.session_state["navigate_to"] = "Projects"
                st.session_state["page_changed"] = True

                st.rerun()

        # -------------------------
        # TEAM MEMBER BUTTONS
        # -------------------------
        for i, (member, count) in enumerate(
            allocation.items(),
            start=1
        ):

            with cols[i]:

                if st.button(
                    f"👤 {member}\n{count} Projects",
                    key=f"team_member_{i}",
                    use_container_width=True
                ):

                    st.session_state.pop(
                        "navigate_to",
                        None
                    )

                    st.session_state.pop(
                        "project_status_filter",
                        None
                    )

                    st.session_state.pop(
                        "selected_team_project",
                        None
                    )

                    st.session_state[
                        "selected_team_project"
                    ] = member

                    st.session_state[
                        "navigate_to"
                    ] = "Projects"

                    st.session_state[
                        "page_changed"
                    ] = True

                    st.rerun()


    st.html("""
    <style>

    /* =========================================
    TEAM KPI CARDS
    ========================================= */

    div[data-testid="stVerticalBlock"]:has(
        > div[data-testid="stHorizontalBlock"]
    ) {
        margin-top: 0px !important;
        margin-bottom: 0px !important;
    }


    /* =========================================
    CARD BUTTON
    ========================================= */

    div[data-testid="stButton"] > button {

        width: 100% !important;

        min-height: 52px !important;
        height: 52px !important;

        background: linear-gradient(
            135deg,
            #013D2B,
            #006747,
            #008A5A
        ) !important;

        border: 1px solid rgba(255,255,255,.35) !important;

        border-radius: 9px !important;

        padding: 4px 5px !important;

        box-shadow:
            0 2px 6px rgba(0,103,71,.14) !important;

        color: white !important;

        font-family:
            Segoe UI,
            Arial,
            sans-serif !important;

        font-size: 10px !important;

        font-weight: 700 !important;

        line-height: 1.15 !important;

        white-space: pre-line !important;

        text-align: center !important;

        margin: 0px !important;
    }


    /* =========================================
    BUTTON TEXT
    ========================================= */

    div[data-testid="stButton"] > button p {

        font-size: 10px !important;

        font-weight: 700 !important;

        line-height: 1.15 !important;

        color: white !important;

        margin: 0 !important;

        white-space: pre-line !important;

        text-align: center !important;
    }


    /* =========================================
    HOVER
    ========================================= */

    div[data-testid="stButton"] > button:hover {

        background: linear-gradient(
            135deg,
            #006747,
            #008A5A
        ) !important;

        transform: translateY(-1px);

        box-shadow:
            0 3px 7px rgba(0,103,71,.20) !important;
    }


    /* =========================================
    FOCUS
    ========================================= */

    div[data-testid="stButton"] > button:focus {

        background: #006747 !important;

        border-color: #FFD700 !important;
    }


    /* =========================================
    COMPACT COLUMNS
    ========================================= */

    div[data-testid="stHorizontalBlock"] {

        gap: 4px !important;

        margin-top: 0px !important;

        margin-bottom: 0px !important;
    }


    /* =========================================
    REMOVE EXTRA SPACING
    ========================================= */

    div[data-testid="stVerticalBlock"] {

        gap: 2px !important;
    }

    </style>
    """)
    
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
    # =====================================================
    # PROJECT TABLE - VIP STYLE
    # =====================================================

    st.markdown("""
    <h2 style="color:#006747;font-size:26px;font-weight:700;margin-top:18px;margin-bottom:9px;">
    📋 Project Details
    </h2>
    """, unsafe_allow_html=True)


    # =====================================================
    # VIP TABLE CSS
    # =====================================================

    st.markdown("""
    <style>

    .project-table-wrapper {
        background:#FFFFFF;
        border:1px solid #DDE5E1;
        border-radius:16px;
        padding:6px;
        box-shadow:0 6px 20px rgba(0,103,71,0.08);
        overflow:hidden;
    }

    .project-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:14px;
    }

    .project-table thead th {
        background:linear-gradient(135deg,#013D2B,#006747,#008A5A);
        color:#FFFFFF;
        font-weight:700;
        padding:14px 12px;
        text-align:left;
        border:none;
    }

    .project-table thead th:first-child {
        border-top-left-radius:11px;
    }

    .project-table thead th:last-child {
        border-top-right-radius:11px;
    }

    .project-table tbody td {
        padding:13px 12px;
        color:#1F2937;
        border-bottom:1px solid #E5E7EB;
        background:#FFFFFF;
    }

    .project-table tbody tr:nth-child(even) td {
        background:#F8FAFC;
    }

    .project-table tbody tr:hover td {
        background:#ECFDF5;
    }

    .project-name {
        font-weight:700;
        color:#006747 !important;
    }

    .status-badge {
        display:inline-block;
        padding:5px 11px;
        border-radius:20px;
        font-size:11px;
        font-weight:800;
        white-space:nowrap;
    }

    .status-live {
        background:#DCFCE7;
        color:#166534;
    }

    .status-uat {
        background:#FEF3C7;
        color:#92400E;
    }

    .status-development {
        background:#DBEAFE;
        color:#1E40AF;
    }

    .status-review {
        background:#CCFBF1;
        color:#115E59;
    }

    .status-cmc {
        background:#EDE9FE;
        color:#5B21B6;
    }

    .status-scoping {
        background:#F3E8FF;
        color:#7E22CE;
    }

    .status-default {
        background:#F3F4F6;
        color:#374151;
    }

    </style>
    """, unsafe_allow_html=True)


    # =====================================================
    # CREATE VIP TABLE
    # =====================================================

    display_df = filtered_df.copy()


    # =====================================================
    # DATE FORMAT
    # =====================================================

    if "Date" in display_df.columns:

        display_df["Date"] = pd.to_datetime(
            display_df["Date"],
            errors="coerce"
        ).dt.strftime("%Y-%m-%d")


    # =====================================================
    # LIVE DATE FORMAT
    # =====================================================

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


    # =====================================================
    # STATUS BADGE
    # =====================================================

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


    if "Status" in display_df.columns:

        display_df["Status"] = display_df[
            "Status"
        ].apply(status_badge)


    # =====================================================
    # PROJECT / MANDATE HIGHLIGHT
    # =====================================================

    if "Mandate" in display_df.columns:

        display_df["Mandate"] = display_df[
            "Mandate"
        ].apply(
            lambda x:
            f'<span class="project-name">📁 {x}</span>'
        )


    # =====================================================
    # HTML TABLE
    # =====================================================

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

    st.markdown(
        "<div style='height:6px;'></div>",
        unsafe_allow_html=True
    )


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
    <h2 style="color:#006747;font-size:24px;font-weight:700;margin-top:14px;margin-bottom:8px;">
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
# DAILY ISSUES
# =====================================================

elif page == "Daily Issues":

    # =================================================
    # HEADER
    # =================================================

    render_header(
        "SMARTPAY DAILY OPERATIONS",
        "📝 Daily Issues",
        "Daily Issues, Updates & Ownership Tracking",
        "DAILY ISSUE MANAGEMENT"
    )


    # =================================================
    # LOAD DATA
    # =================================================

    daily_issues_df = (
        load_daily_issues_from_github()
    )


    required_daily_columns = [
        "Issue",
        "Update",
        "End / Pending With",
        "Team Member",
        "Status"
    ]


    for col in required_daily_columns:

        if col not in daily_issues_df.columns:

            if col == "Status":

                daily_issues_df[col] = "Active"

            else:

                daily_issues_df[col] = ""


    # =================================================
    # STATUS CLEAN
    # =================================================

    daily_issues_df["Status"] = (
        daily_issues_df["Status"]
        .fillna("Active")
        .astype(str)
        .str.strip()
        .replace({
            "ACTIVE": "Active",
            "active": "Active",
            "CLOSED": "Closed",
            "closed": "Closed",
            "Close": "Closed",
            "close": "Closed"
        })
    )


    daily_issues_df.loc[
        ~daily_issues_df["Status"].isin(
            ["Active", "Closed"]
        ),
        "Status"
    ] = "Active"


    # =================================================
    # COLUMN ORDER
    # =================================================

    extra_columns = [
        col
        for col in daily_issues_df.columns
        if col not in required_daily_columns
    ]


    daily_issues_df = (
        daily_issues_df[
            required_daily_columns
            + extra_columns
        ]
        .fillna("")
    )


    # =================================================
    # DAILY ISSUES CSS
    # =================================================

    st.markdown("""
<style>

/* ==========================================
   KPI CONTAINER
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stButton"] {

    width:100% !important;
}


/* ==========================================
   KPI CARDS
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stButton"]
button {

    width:100% !important;

    min-height:96px !important;
    height:96px !important;

    background:#FFFFFF !important;

    border:1px solid #E5E7EB !important;

    border-radius:16px !important;

    padding:10px 6px !important;

    box-shadow:
        0 6px 16px
        rgba(0,0,0,.07) !important;

    color:#111827 !important;

    font-family:
        "Segoe UI",
        Arial,
        sans-serif !important;

    font-size:13px !important;

    font-weight:650 !important;

    line-height:1.3 !important;

    white-space:pre-line !important;

    text-align:center !important;

    display:flex !important;

    align-items:center !important;

    justify-content:center !important;

    transition:all .2s ease !important;
}


/* ==========================================
   BUTTON TEXT
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stButton"]
button p {

    font-family:
        "Segoe UI",
        Arial,
        sans-serif !important;

    font-size:13px !important;

    font-weight:700 !important;

    line-height:1.3 !important;

    white-space:pre-line !important;

    text-align:center !important;

    display:block !important;

    width:100% !important;

    margin:0 !important;

    padding:0 !important;

    color:#111827 !important;
}


/* ==========================================
   HOVER
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stButton"]
button:hover {

    background:#F8FAFC !important;

    transform:translateY(-2px) !important;

    box-shadow:
        0 10px 22px
        rgba(0,0,0,.11) !important;
}


/* ==========================================
   ALL
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stHorizontalBlock"]
div[data-testid="stColumn"]:nth-child(1)
div[data-testid="stButton"]
button {

    border-top:5px solid #006747 !important;
}


/* ==========================================
   ACTIVE
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stHorizontalBlock"]
div[data-testid="stColumn"]:nth-child(2)
div[data-testid="stButton"]
button {

    border-top:5px solid #00A86B !important;
}


/* ==========================================
   CLOSED
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stHorizontalBlock"]
div[data-testid="stColumn"]:nth-child(3)
div[data-testid="stButton"]
button {

    border-top:5px solid #D32F2F !important;
}


/* ==========================================
   COLUMN GAP
   ========================================== */

.st-key-daily_issue_kpis
div[data-testid="stHorizontalBlock"] {

    gap:8px !important;
}


/* ==========================================
   TABLE
   ========================================== */

.daily-issues-table-wrapper {

    background:#FFFFFF;

    border:1px solid #DDE5E1;

    border-radius:16px;

    padding:6px;

    box-shadow:
        0 6px 20px
        rgba(0,103,71,.08);

    overflow:auto;

    margin-top:5px;

    margin-bottom:10px;
}


.daily-issues-table {

    width:100%;

    border-collapse:separate;

    border-spacing:0;

    font-size:14px;

    table-layout:fixed;

    font-family:
        "Segoe UI",
        Arial,
        sans-serif;
}


.daily-issues-table thead th {

    background:
        linear-gradient(
            135deg,
            #013D2B,
            #006747,
            #008A5A
        );

    color:#FFFFFF;

    font-weight:700;

    padding:14px 12px;

    text-align:left;

    border-bottom:1px solid
        rgba(255,255,255,.20);

    line-height:1.6;
}


.daily-issues-table thead th:first-child {

    border-top-left-radius:10px;
}


.daily-issues-table thead th:last-child {

    border-top-right-radius:10px;
}


.daily-issues-table tbody td {

    padding:13px 12px;

    color:#1F2937;

    border-bottom:1px solid #E5E7EB;

    background:#FFFFFF;

    vertical-align:top;

    line-height:1.6;

    white-space:normal;

    word-wrap:break-word;

    overflow-wrap:anywhere;
}


.daily-issues-table tbody tr:nth-child(even) td {

    background:#F8FAFC;
}


.daily-issues-table tbody tr:hover td {

    background:#ECFDF5;
}


.daily-issue-name {

    color:#006747;

    font-weight:700;
}


.daily-pending-badge {

    display:inline-block;

    padding:4px 9px;

    border-radius:16px;

    background:#FEF3C7;

    color:#92400E;

    font-size:11px;

    font-weight:800;

    white-space:nowrap;
}


.daily-member-badge {

    display:inline-block;

    padding:4px 9px;

    border-radius:16px;

    background:#DCFCE7;

    color:#166534;

    font-size:11px;

    font-weight:800;

    white-space:nowrap;
}


.daily-status-active {

    display:inline-block;

    padding:4px 10px;

    border-radius:16px;

    background:#DCFCE7;

    color:#166534;

    font-size:11px;

    font-weight:800;

    white-space:nowrap;
}


.daily-status-closed {

    display:inline-block;

    padding:4px 10px;

    border-radius:16px;

    background:#FEE2E2;

    color:#991B1B;

    font-size:11px;

    font-weight:800;

    white-space:nowrap;
}


.daily-edit-info {

    background:#ECFDF5;

    border:1px solid #B7E4C7;

    border-left:5px solid #006747;

    border-radius:10px;

    padding:9px 12px;

    color:#006747;

    font-size:13px;

    font-weight:600;

    margin-bottom:8px;
}

</style>
""", unsafe_allow_html=True)


    # =================================================
    # KPI COUNTS
    # =================================================

    daily_status = (
        daily_issues_df["Status"]
        .astype(str)
        .str.strip()
        .str.lower()
    )


    total_daily_issues = len(
        daily_issues_df
    )


    active_daily_issues = (
        daily_status == "active"
    ).sum()


    closed_daily_issues = (
        daily_status == "closed"
    ).sum()


    # =================================================
    # KPI CARD STATE
    # =================================================

    if (
        "daily_issue_stage"
        not in st.session_state
    ):

        st.session_state[
            "daily_issue_stage"
        ] = "All"


    # =================================================
    # CLICKABLE KPI CARDS
    # =================================================

    with st.container(
        key="daily_issue_kpis"
    ):

        c1, c2, c3 = st.columns(3)


        # ---------------------------------------------
        # ALL
        # ---------------------------------------------

        with c1:

            if st.button(
                f"ALL\n{total_daily_issues}",
                key="daily_issue_all",
                use_container_width=True
            ):

                st.session_state[
                    "daily_issue_stage"
                ] = "All"

                st.rerun()


        # ---------------------------------------------
        # ACTIVE
        # ---------------------------------------------

        with c2:

            if st.button(
                f"ACTIVE ISSUES\n{active_daily_issues}",
                key="daily_issue_active",
                use_container_width=True
            ):

                st.session_state[
                    "daily_issue_stage"
                ] = "Active"

                st.rerun()


        # ---------------------------------------------
        # CLOSED
        # ---------------------------------------------

        with c3:

            if st.button(
                f"CLOSED ISSUES\n{closed_daily_issues}",
                key="daily_issue_closed",
                use_container_width=True
            ):

                st.session_state[
                    "daily_issue_stage"
                ] = "Closed"

                st.rerun()


    # =================================================
    # APPLY KPI FILTER
    # =================================================

    selected_daily_card = (
        st.session_state.get(
            "daily_issue_stage",
            "All"
        )
    )


    filtered_daily_df = (
        daily_issues_df.copy()
    )


    if selected_daily_card != "All":

        filtered_daily_df = (
            filtered_daily_df[
                filtered_daily_df["Status"]
                .astype(str)
                .str.strip()
                .str.lower()
                ==
                selected_daily_card.lower()
            ]
        )


    st.markdown(
        "<div style='height:8px;'></div>",
        unsafe_allow_html=True
    )


    # =================================================
    # TABLE DISPLAY
    # =================================================

    display_daily_df = (
        filtered_daily_df.copy()
    )


    if not display_daily_df.empty:

        display_daily_df["Issue"] = (
            display_daily_df["Issue"]
            .apply(
                lambda x:
                f'<span class="daily-issue-name">'
                f'{x}</span>'
            )
        )


        display_daily_df[
            "End / Pending With"
        ] = (
            display_daily_df[
                "End / Pending With"
            ]
            .apply(
                lambda x:
                (
                    f'<span class="daily-pending-badge">'
                    f'{x}</span>'
                    if str(x).strip()
                    else ""
                )
            )
        )


        display_daily_df[
            "Team Member"
        ] = (
            display_daily_df[
                "Team Member"
            ]
            .apply(
                lambda x:
                (
                    f'<span class="daily-member-badge">'
                    f'{x}</span>'
                    if str(x).strip()
                    else ""
                )
            )
        )


        # ---------------------------------------------
        # STATUS BADGE
        # ---------------------------------------------

        display_daily_df[
            "Status"
        ] = (
            display_daily_df[
                "Status"
            ]
            .apply(
                lambda x:
                (
                    '<span class="daily-status-closed">'
                    'Closed'
                    '</span>'
                    if str(x).strip().lower()
                    == "closed"
                    else
                    '<span class="daily-status-active">'
                    'Active'
                    '</span>'
                )
            )
        )


        table_html = (
            display_daily_df
            .to_html(
                index=False,
                escape=False,
                classes="daily-issues-table"
            )
        )


        st.markdown(
            f"""
<div class="daily-issues-table-wrapper">
{table_html}
</div>
""",
            unsafe_allow_html=True
        )

    else:

        st.info(
            f"No {selected_daily_card.lower()} "
            "Daily Issues available."
        )


    # =================================================
    # EDIT MODE STATE
    # =================================================

    if (
        "daily_issues_edit_mode"
        not in st.session_state
    ):

        st.session_state[
            "daily_issues_edit_mode"
        ] = False


    # =================================================
    # EDIT BUTTON
    # =================================================

    if not st.session_state[
        "daily_issues_edit_mode"
    ]:

        if st.button(
            "✏️ Edit Daily Issues",
            key="daily_issues_edit_button",
            use_container_width=True
        ):

            st.session_state[
                "daily_issues_edit_mode"
            ] = True

            st.rerun()


    # =================================================
    # EDITOR
    # =================================================

    if st.session_state[
        "daily_issues_edit_mode"
    ]:

        st.markdown(
            """
<div class="daily-edit-info">
✏️ Edit existing issues, add/delete rows,
or manage columns below.
<br>
💾 Click <b>Save Changes</b> after editing.
</div>
""",
            unsafe_allow_html=True
        )


        # =================================================
        # COLUMN MANAGEMENT
        # =================================================

        st.markdown(
            """
<h3 style="
color:#006747;
font-size:19px;
font-weight:700;
margin-top:10px;
margin-bottom:6px;">
⚙️ Manage Columns
</h3>
""",
            unsafe_allow_html=True
        )


        field_col1, field_col2 = st.columns(
            [3, 1]
        )


        # ---------------------------------------------
        # ADD FIELD
        # ---------------------------------------------

        with field_col1:

            new_daily_field = st.text_input(
                "New Column Name",
                placeholder="e.g. Vendor Update",
                key="daily_new_field_input"
            )


        with field_col2:

            st.markdown(
                "<div style='height:26px;'></div>",
                unsafe_allow_html=True
            )


            if st.button(
                "➕ Add Column",
                key="daily_add_field",
                use_container_width=True
            ):

                field_name = (
                    new_daily_field
                    .strip()
                )


                if not field_name:

                    st.warning(
                        "⚠️ Please enter a column name."
                    )

                elif (
                    field_name
                    in daily_issues_df.columns
                ):

                    st.warning(
                        "⚠️ This column already exists."
                    )

                else:

                    daily_issues_df[
                        field_name
                    ] = ""


                    # Save current extra-column
                    # structure in session

                    current_extra = [
                        col
                        for col
                        in daily_issues_df.columns
                        if col
                        not in required_daily_columns
                    ]


                    st.session_state[
                        "daily_extra_columns"
                    ] = current_extra


                    st.success(
                        f"✅ '{field_name}' "
                        "column added."
                    )

                    st.rerun()


        # ---------------------------------------------
        # REMOVE FIELD
        # ---------------------------------------------

        current_extra_columns = [
            col
            for col
            in daily_issues_df.columns
            if col
            not in required_daily_columns
        ]


        if current_extra_columns:

            remove_col1, remove_col2 = st.columns(
                [3, 1]
            )


            with remove_col1:

                remove_daily_field = st.selectbox(
                    "Remove Column",
                    current_extra_columns,
                    key="daily_remove_field_select"
                )


            with remove_col2:

                st.markdown(
                    "<div style='height:26px;'></div>",
                    unsafe_allow_html=True
                )


                if st.button(
                    "🗑 Remove Column",
                    key="daily_remove_field",
                    use_container_width=True
                ):

                    daily_issues_df = (
                        daily_issues_df
                        .drop(
                            columns=[
                                remove_daily_field
                            ]
                        )
                    )


                    remaining_extra = [
                        col
                        for col
                        in daily_issues_df.columns
                        if col
                        not in required_daily_columns
                    ]


                    st.session_state[
                        "daily_extra_columns"
                    ] = remaining_extra


                    st.success(
                        f"✅ '{remove_daily_field}' "
                        "column removed."
                    )

                    st.rerun()


        else:

            st.info(
                "No extra columns available to remove. "
                "Core columns are protected."
            )


        # =================================================
        # RESTORE SESSION EXTRA COLUMNS
        # =================================================

        session_extra = (
            st.session_state.get(
                "daily_extra_columns",
                []
            )
        )


        for col in session_extra:

            if col not in daily_issues_df.columns:

                daily_issues_df[col] = ""


        final_extra_columns = [
            col
            for col in daily_issues_df.columns
            if col not in required_daily_columns
        ]


        daily_issues_df = (
            daily_issues_df[
                required_daily_columns
                + final_extra_columns
            ]
        )


        # =================================================
        # EDITABLE TABLE
        # =================================================

        edited_daily_df = st.data_editor(

            daily_issues_df,

            use_container_width=True,

            hide_index=True,

            num_rows="dynamic",

            height=500,

            disabled=[],

            column_config={

                "Status":
                    st.column_config.SelectboxColumn(
                        "Status",
                        options=[
                            "Active",
                            "Closed"
                        ],
                        required=True
                    )
            },

            key="daily_issues_editor"
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
                key="daily_issues_save",
                use_container_width=True
            ):

                try:

                    edited_daily_df = (
                        edited_daily_df
                        .fillna("")
                    )


                    success, message = (
                        save_daily_issues_to_github(
                            edited_daily_df
                        )
                    )


                    if success:

                        st.success(
                            "✅ Daily Issues saved "
                            "successfully to GitHub."
                        )


                        st.session_state[
                            "daily_issues_edit_mode"
                        ] = False


                        if (
                            "daily_issues_editor"
                            in st.session_state
                        ):

                            del st.session_state[
                                "daily_issues_editor"
                            ]


                        st.rerun()


                    else:

                        st.error(
                            f"❌ Daily Issues save failed: "
                            f"{message}"
                        )


                except Exception as e:

                    st.error(
                        f"❌ Save error: {e}"
                    )


        # =================================================
        # CLOSE
        # =================================================

        with close_col:

            if st.button(
                "❌ Close Editor",
                key="daily_issues_close",
                use_container_width=True
            ):

                st.session_state[
                    "daily_issues_edit_mode"
                ] = False


                if (
                    "daily_issues_editor"
                    in st.session_state
                ):

                    del st.session_state[
                        "daily_issues_editor"
                    ]


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

    render_header('SMARTPAY ANALYTICS', '📊 Analytics Dashboard', 'SmartPay Project Insights & Team Performance', 'ANALYTICS & INSIGHTS')

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
            "#008A5A"
        )


    with k3:

        analytics_card(
            "UAT %",
            f"{uat_percent}%",
            "#20A464"
        )


    with k4:

        analytics_card(
            "Team Members",
            team_members,
            "#4CAF7A"
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


    # ==========================================
    # SMARTPAY GREEN THEME
    # ==========================================

    color_map = {
        "LIVE": "#006747",
        "UAT": "#008A5A",
        "IS REVIEW": "#20A464",
        "CMC": "#4CAF7A",
        "UNDER SCOPING": "#8BC9A8"
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


    # ==========================================
    # STATUS BAR CHART
    # ==========================================

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
                color="#006747",
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
                color="#006747",
                size=12
            ),
            tickfont=dict(
                color="#374151",
                size=11
            ),
            gridcolor="#DDE5E1",
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
        textfont=dict(
            color="white",
            size=11
        ),
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


    # ==========================================
    # TEAM PERFORMANCE BAR CHART
    # ==========================================

    fig2 = px.bar(
        allocation_count,
        x="Allocation",
        y="Projects",
        text="Projects",
        color="Projects",
        color_continuous_scale=[
            "#D1FAE5",
            "#8BC9A8",
            "#4CAF7A",
            "#20A464",
            "#008A5A",
            "#006747"
        ]
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
                color="#006747",
                size=12
            ),
            tickfont=dict(
                color="#374151",
                size=11
            ),
            gridcolor="#DDE5E1",
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
            color="#006747",
            size=11
        ),
        marker_line_width=0
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


        # ==========================================
        # PERSON STATUS PIE
        # ==========================================

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
            ),
            marker_line=dict(
                color="white",
                width=2
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
            # SMARTPAY GREEN STATUS COLORS
            # ======================================

            if current_status == "LIVE":

                status_bg = "#DCFCE7"
                status_color = "#166534"

            elif current_status == "UAT":

                status_bg = "#D1FAE5"
                status_color = "#006747"

            elif current_status == "CMC":

                status_bg = "#B7E4C7"
                status_color = "#087443"

            elif current_status == "IS REVIEW":

                status_bg = "#C6F6D5"
                status_color = "#15803D"

            elif current_status in [
                "SCOPING",
                "UNDER SCOPING"
            ]:

                status_bg = "#E8F5E9"
                status_color = "#2E7D5B"

            else:

                status_bg = "#F3F4F6"
                status_color = "#374151"


            # ======================================
            # PROJECT CARD
            # ======================================

            card_html = f"""
        <div style="background:#FFFFFF;border:1px solid #DDE5E1;border-radius:14px;padding:12px 14px;margin-bottom:10px;box-shadow:0 4px 12px rgba(0,103,71,.06);box-sizing:border-box;">

        <!-- PROJECT HEADER -->

        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:9px;">

        <div style="color:#006747;font-size:14px;font-weight:750;line-height:1.3;flex:1;word-break:break-word;">
        📁 {project_name}
        </div>

        <div style="background:{status_bg};color:{status_color};padding:4px 9px;border-radius:14px;font-size:9px;font-weight:800;white-space:nowrap;flex-shrink:0;">
        {current_status}
        </div>

        </div>


        <!-- PROGRESS LINE -->

        <div style="width:100%;height:6px;background:#E5E7EB;border-radius:10px;overflow:hidden;">

        <div style="width:{progress_percent}%;height:100%;background:linear-gradient(90deg,#013D2B,#006747,#008A5A);border-radius:10px;">
        </div>

        </div>


        <!-- STATUS STEPS -->

        <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-top:9px;">
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
        <div style="flex:1;text-align:center;">

        <div style="width:8px;height:8px;background:{dot_color};border-radius:50%;margin:auto;">
        </div>

        <div style="margin-top:4px;color:{text_color};font-size:8px;font-weight:700;white-space:nowrap;">
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


        # ==========================================
        # PROJECT DETAILS TABLE
        # ==========================================

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
            "#006747"
        )


    with b2:

        analytics_card(
            "BAU Owners",
            bau_owners,
            "#008A5A"
        )


    with b3:

        analytics_card(
            "BAU Updates",
            bau_updates,
            "#20A464"
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


        # =================================================
        # BAU OWNER BAR CHART
        # =================================================

        fig_bau = px.bar(
            bau_owner_count,
            x="Allocation",
            y="Projects",
            text="Projects",
            color="Projects",
            color_continuous_scale=[
                "#D1FAE5",
                "#8BC9A8",
                "#4CAF7A",
                "#20A464",
                "#008A5A",
                "#006747"
            ]
        )


        fig_bau.update_traces(
            textposition="outside",
            marker_line_width=0,
            textfont=dict(
                color="#006747",
                size=11
            )
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
                zeroline=False,
                tickfont=dict(
                    color="#374151",
                    size=11
                )
            ),

            yaxis=dict(
                title="BAU Projects",
                title_font=dict(
                    color="#006747",
                    size=11
                ),
                gridcolor="#DDE5E1",
                zeroline=False,
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

    render_header('SMARTPAY PROJECT MANAGEMENT', '📅 Project Timeline', 'Track the current progress of SmartPay projects across every delivery stage.', 'PROJECT TIMELINE')
    # ==========================================
    # LEGEND
    # ==========================================

    st.html("""
    <div style="background:#FFFFFF;border-radius:12px;padding:8px 12px;border:1px solid #DDE5E1;box-shadow:0 4px 12px rgba(0,103,71,.06);margin-bottom:7px;font-family:Segoe UI,Arial,sans-serif;">
    <div style="color:#006747;font-size:13px;font-weight:700;margin-bottom:5px;">Timeline Status</div>
    <div style="display:flex;align-items:center;gap:18px;flex-wrap:wrap;font-size:11px;font-weight:600;">
    <div><span style="display:inline-block;width:9px;height:9px;background:#16A34A;border-radius:50%;margin-right:5px;"></span><span style="color:#374151;">Completed</span></div>
    <div><span style="display:inline-block;width:9px;height:9px;background:#F59E0B;border-radius:50%;margin-right:5px;"></span><span style="color:#374151;">Current Stage</span></div>
    <div><span style="display:inline-block;width:9px;height:9px;background:#D1D5DB;border-radius:50%;margin-right:5px;"></span><span style="color:#374151;">Pending</span></div>
    </div>
    </div>
    """)


    # ==========================================
    # SEARCH
    # ==========================================

    st.html("""
    <div style="color:#006747;font-size:22px;font-weight:700;margin-top:2px;margin-bottom:3px;font-family:Segoe UI,Arial,sans-serif;">
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
        "<div style='height:1px;'></div>",
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
    # TIMELINE FUNCTION - ONE BOX PER PROJECT
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
        # PROJECT BOX
        # ==========================================

        project_html = f"""
    <div style="background:#FFFFFF;border:1px solid #D5E3DD;border-radius:14px;padding:10px;margin-bottom:8px;box-shadow:0 4px 12px rgba(0,103,71,.07);font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#6B7280;font-size:9px;font-weight:600;letter-spacing:1px;margin-bottom:2px;">
    PROJECT TIMELINE
    </div>

    <div style="color:#006747;font-size:16px;font-weight:750;line-height:1.25;word-break:break-word;margin-bottom:7px;">
    📌 {project["Mandate"]}
    </div>


    <!-- TIMELINE -->

    <div style="background:#F9FBFD;border:1px solid #E5E7EB;border-radius:11px;padding:10px 8px;overflow-x:auto;margin-bottom:7px;">

    <div style="display:flex;align-items:flex-start;min-width:580px;">
    """


        # ==========================================
        # TIMELINE STAGES
        # ==========================================

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


            connector = ""

            if i < len(stages) - 1:

                if i < current_index:
                    line_color = "#16A34A"
                else:
                    line_color = "#D1D5DB"


                connector = f"""
    <div style="flex:1;height:3px;background:{line_color};margin-top:13px;"></div>
    """


            project_html += f"""
    <div style="width:74px;text-align:center;flex-shrink:0;">

    <div style="width:26px;height:26px;border-radius:50%;background:{color};color:white;margin:auto;line-height:26px;font-size:12px;font-weight:700;box-shadow:0 3px 7px rgba(0,0,0,.10);">
    {symbol}
    </div>

    <div style="margin-top:4px;color:#374151;font-size:9px;font-weight:700;white-space:nowrap;">
    {stage}
    </div>

    </div>

    {connector}
    """


        # ==========================================
        # PROJECT INFORMATION
        # ==========================================

        project_html += f"""
    </div>
    </div>


    <div style="display:flex;gap:7px;width:100%;">

    <div style="flex:1;background:#F9FBFD;border:1px solid #E5E7EB;border-radius:9px;padding:7px 9px;min-height:50px;">

    <div style="color:#6B7280;font-size:8px;font-weight:600;">
    PROJECT
    </div>

    <div style="color:#111827;font-size:11px;font-weight:700;line-height:1.2;margin-top:3px;">
    {project["Mandate"]}
    </div>

    </div>


    <div style="flex:1;background:#F9FBFD;border:1px solid #E5E7EB;border-radius:9px;padding:7px 9px;min-height:50px;">

    <div style="color:#6B7280;font-size:8px;font-weight:600;">
    OWNER
    </div>

    <div style="color:#006747;font-size:11px;font-weight:700;margin-top:3px;">
    {project["Allocation"]}
    </div>

    </div>


    <div style="flex:1;background:#F9FBFD;border:1px solid #E5E7EB;border-radius:9px;padding:7px 9px;min-height:50px;">

    <div style="color:#92400E;font-size:8px;font-weight:600;">
    CURRENT STAGE
    </div>

    <div style="color:#F59E0B;font-size:11px;font-weight:700;margin-top:3px;">
    {current}
    </div>

    </div>

    </div>

    </div>
    """


        st.html(project_html)


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

#     st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)


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

#     st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

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


#     st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)


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


#         st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)


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

#     st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

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

    render_header('SMARTPAY BUSINESS AS USUAL', '🏦 BAU Monitoring', 'Live Business Operations Monitoring, ownership and ongoing updates.', 'BUSINESS AS USUAL')

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
    <div style="background:#FFFFFF;border-radius:14px;padding:10px 8px;height:88px;border-top:5px solid {color};border-left:1px solid #DDE5E1;border-right:1px solid #DDE5E1;border-bottom:1px solid #DDE5E1;box-shadow:0 4px 12px rgba(0,103,71,.06);display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#6B7280;font-size:11px;font-weight:700;margin-bottom:5px;">
    {title}
    </div>

    <div style="color:{color};font-size:28px;font-weight:800;line-height:1;">
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
            "#008A5A"
        )


    with k3:

        bau_card(
            "Updated Projects",
            bau_updates,
            "#20A464"
        )


    with k4:

        bau_card(
            "Categories",
            bau_categories,
            "#4CAF7A"
        )


    st.markdown(
        "<div style='height:8px;'></div>",
        unsafe_allow_html=True
    )


    # =================================================
    # FILTERS
    # =================================================

    st.html("""
    <div style="color:#006747;font-size:24px;font-weight:700;margin-top:2px;margin-bottom:5px;font-family:Segoe UI,Arial,sans-serif;">
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
    <div style="color:#006747;font-size:24px;font-weight:700;margin-top:2px;margin-bottom:7px;font-family:Segoe UI,Arial,sans-serif;">
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
    <div style="background:linear-gradient(135deg,#013D2B,#006747,#008A5A);border:1px solid rgba(255,255,255,.15);border-radius:14px;padding:10px 8px;text-align:center;box-shadow:0 4px 12px rgba(0,103,71,.12);margin-bottom:8px;font-family:Segoe UI,Arial,sans-serif;">

    <div style="font-size:20px;line-height:1;color:white;">
    👤
    </div>

    <div style="color:#D1FAE5;font-size:13px;font-weight:700;margin-top:5px;">
    {owner["Allocation"]}
    </div>

    <div style="color:white;font-size:25px;font-weight:800;margin-top:4px;line-height:1;">
    {owner["Projects"]}
    </div>

    <div style="color:#D1FAE5;font-size:10px;margin-top:3px;">
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
    font-size:26px;
    font-weight:700;
    margin-top:18px;
    margin-bottom:10px;">
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
        border-radius:16px;
        padding:6px;
        box-shadow:0 6px 20px rgba(0,103,71,.08);
        overflow-x:auto;
    }

    .bau-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:14px;
    }

    .bau-table thead th {
        background:linear-gradient(135deg,#013D2B,#006747,#008A5A);
        color:#FFFFFF;
        padding:14px 12px;
        font-weight:700;
        text-align:left;
        border:none;
    }

    .bau-table thead th:first-child {
        border-top-left-radius:11px;
    }

    .bau-table thead th:last-child {
        border-top-right-radius:11px;
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
        background:#DCFCE7;
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


    st.markdown(
        "<div style='height:4px;'></div>",
        unsafe_allow_html=True
    )


    # =================================================
    # RECENT / ONGOING UPDATES
    # =================================================

    if "Update" in filtered_bau_df.columns:

        st.markdown("""
    <h2 style="
    color:#006747;
    font-size:24px;
    font-weight:700;
    margin-top:14px;
    margin-bottom:9px;">
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
    <div style="background:#FFFFFF;border:1px solid #DDE5E1;border-left:5px solid #006747;border-radius:12px;padding:12px 15px;margin-bottom:8px;box-shadow:0 4px 12px rgba(0,103,71,.06);font-family:Segoe UI,Arial,sans-serif;">

    <div style="display:flex;justify-content:space-between;align-items:center;gap:15px;">

    <div style="color:#006747;font-size:14px;font-weight:700;">
    📁 {project_name}
    </div>

    <div style="color:#008A5A;font-size:11px;font-weight:700;">
    👤 {owner_name}
    </div>

    </div>

    <div style="color:#374151;font-size:13px;margin-top:7px;line-height:1.45;">
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
# IS ISSUES
# =====================================================

elif page == "IS Issues":

    import os
    import pandas as pd

    # =====================================================
    # LOAD IS ISSUES FILE
    # =====================================================

    file = os.path.join(
        os.path.dirname(__file__),
        "data",
        "IS Issues.xlsx"
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
                "application name" in row
                and "severity" in row
                and "issues description" in row
            ):

                header_row = i
                break

        if header_row is None:

            st.error(
                "IS Issues header row not found."
            )

            st.dataframe(
                raw.head(20)
            )

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


        df = (
            df
            .dropna(how="all")
            .reset_index(drop=True)
        )


    except Exception as e:

        st.error(
            f"IS Issues file error: {e}"
        )

        st.stop()


    # =====================================================
    # REQUIRED COLUMNS
    # =====================================================

    columns = [
        "Application Name",
        "Severity",
        "Issues Description",
        "Status of Issue",
        "DBG Remarks"
    ]


    for col in columns:

        if col not in df.columns:

            df[col] = ""


    df = df[
        columns
    ].fillna("")


    # =====================================================
    # CLEAN DATA
    # =====================================================

    for col in columns:

        df[col] = (
            df[col]
            .astype(str)
            .str.replace(
                r"\s+",
                " ",
                regex=True
            )
            .str.strip()
        )


    df["Severity"] = (
        df["Severity"]
        .str.upper()
        .str.strip()
    )


    df["Status of Issue"] = (
        df["Status of Issue"]
        .str.replace(
            r"\s+",
            " ",
            regex=True
        )
        .str.strip()
    )


    # =====================================================
    # HEADER
    # =====================================================

    render_header('INFORMATION SECURITY MONITORING', '🔐 IS Issues', 'Information Security Issues &amp; Fixation Tracking', 'SECURITY ISSUE MONITORING')


    # =====================================================
    # KPI COUNTS
    # =====================================================

    severity = (
        df["Severity"]
        .astype(str)
        .str.strip()
        .str.upper()
    )


    issue_status = (
        df["Status of Issue"]
        .astype(str)
        .str.replace(
            r"\s+",
            " ",
            regex=True
        )
        .str.strip()
        .str.upper()
    )


    total_count = len(df)

    high_count = (
        severity == "HIGH"
    ).sum()

    medium_count = (
        severity == "MEDIUM"
    ).sum()

    low_count = (
        severity == "LOW"
    ).sum()

    fixed_count = (
        issue_status == "FIXED"
    ).sum()

    fixation_count = (
        issue_status
        == "FIXATION IN PROGRESS"
    ).sum()


    # =====================================================
    # VIP CLICKABLE KPI CARDS CSS
    # =====================================================

    st.markdown("""
    <style>

    /* =====================================================
    KPI CARD CONTAINER
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stButton"] {

        width:100% !important;
    }


    /* =====================================================
    BASE CARD
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stButton"]
    button {

        width:100% !important;

        min-height:88px !important;

        height:88px !important;

        background:#FFFFFF !important;

        border:
            1px solid
            #DDE5E1 !important;

        border-radius:14px !important;

        padding:10px 6px !important;

        box-shadow:
            0 5px 16px
            rgba(0,103,71,.08) !important;

        color:#111827 !important;

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        font-size:13px !important;

        font-weight:700 !important;

        line-height:1.15 !important;

        white-space:pre-line !important;

        text-align:center !important;

        display:flex !important;

        flex-direction:column !important;

        justify-content:center !important;

        align-items:center !important;

        transition:
            all .2s ease !important;
    }


    /* =====================================================
    BUTTON TEXT
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stButton"]
    button p {

        width:100% !important;

        margin:0 !important;

        padding:0 !important;

        font-family:
            "Segoe UI",
            Arial,
            sans-serif !important;

        font-size:13px !important;

        font-weight:700 !important;

        line-height:1.25 !important;

        color:#111827 !important;

        white-space:pre-line !important;

        text-align:center !important;
    }


    /* =====================================================
    HOVER
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stButton"]
    button:hover {

        background:#F8FAFC !important;

        color:#111827 !important;

        transform:
            translateY(-2px) !important;

        box-shadow:
            0 8px 20px
            rgba(0,103,71,.14) !important;
    }


    .st-key-is_issues_kpis
    div[data-testid="stButton"]
    button:hover p {

        color:#111827 !important;
    }


    /* =====================================================
    SELECTED / FOCUS
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stButton"]
    button:focus {

        outline:none !important;

        background:#EAF5F0 !important;

        border:
            2px solid
            #006747 !important;

        box-shadow:
            0 0 0 3px
            rgba(0,103,71,.12),
            0 8px 18px
            rgba(0,103,71,.15) !important;

        color:#006747 !important;
    }


    .st-key-is_issues_kpis
    div[data-testid="stButton"]
    button:focus p {

        color:#006747 !important;

        font-weight:700 !important;
    }


    /* =====================================================
    ALL
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(1)
    div[data-testid="stButton"]
    button {

        border-top:
            4px solid
            #006747 !important;
    }


    /* =====================================================
    HIGH
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(2)
    div[data-testid="stButton"]
    button {

        border-top:
            4px solid
            #D32F2F !important;
    }


    /* =====================================================
    MEDIUM
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(3)
    div[data-testid="stButton"]
    button {

        border-top:
            4px solid
            #F9A825 !important;
    }


    /* =====================================================
    LOW
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(4)
    div[data-testid="stButton"]
    button {

        border-top:
            4px solid
            #008A5A !important;
    }


    /* =====================================================
    FIXED
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(5)
    div[data-testid="stButton"]
    button {

        border-top:
            4px solid
            #00C853 !important;
    }


    /* =====================================================
    FIXATION IN PROGRESS
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stHorizontalBlock"]:nth-child(1)
    div[data-testid="stColumn"]:nth-child(6)
    div[data-testid="stButton"]
    button {

        border-top:
            4px solid
            #FF9800 !important;
    }


    /* =====================================================
    COLUMN SPACING
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stHorizontalBlock"] {

        gap:8px !important;
    }


    /* =====================================================
    REMOVE EXTRA GAPS
    ===================================================== */

    .st-key-is_issues_kpis
    div[data-testid="stVerticalBlock"] {

        gap:0 !important;
    }


    /* =====================================================
    MOBILE
    ===================================================== */

    @media (max-width:768px) {

        .st-key-is_issues_kpis
        div[data-testid="stButton"]
        button {

            min-height:82px !important;

            height:82px !important;

            padding:8px 4px !important;

            font-size:12px !important;
        }

        .st-key-is_issues_kpis
        div[data-testid="stButton"]
        button p {

            font-size:12px !important;

            line-height:1.2 !important;
        }
    }

    </style>
    """, unsafe_allow_html=True)
    # =====================================================
    # KPI CARD CONTAINER
    # =====================================================

    with st.container(
        key="is_issues_kpis"
    ):

        c1, c2, c3, c4, c5, c6 = st.columns(
            6,
            gap="small"
        )


        # =================================================
        # ALL
        # =================================================

        with c1:

            if st.button(
                f"ALL\n{total_count}",
                key="is_all",
                use_container_width=True
            ):

                st.session_state[
                    "is_issue_stage"
                ] = "All"

                st.rerun()


        # =================================================
        # HIGH
        # =================================================

        with c2:

            if st.button(
                f"HIGH\n{high_count}",
                key="is_high",
                use_container_width=True
            ):

                st.session_state[
                    "is_issue_stage"
                ] = "HIGH"

                st.rerun()


        # =================================================
        # MEDIUM
        # =================================================

        with c3:

            if st.button(
                f"MEDIUM\n{medium_count}",
                key="is_medium",
                use_container_width=True
            ):

                st.session_state[
                    "is_issue_stage"
                ] = "MEDIUM"

                st.rerun()


        # =================================================
        # LOW
        # =================================================

        with c4:

            if st.button(
                f"LOW\n{low_count}",
                key="is_low",
                use_container_width=True
            ):

                st.session_state[
                    "is_issue_stage"
                ] = "LOW"

                st.rerun()


        # =================================================
        # FIXED
        # =================================================

        with c5:

            if st.button(
                f"FIXED\n{fixed_count}",
                key="is_fixed",
                use_container_width=True
            ):

                st.session_state[
                    "is_issue_stage"
                ] = "FIXED"

                st.rerun()


        # =================================================
        # FIXATION IN PROGRESS
        # =================================================

        with c6:

            if st.button(
                f"FIXATION IN PROGRESS\n{fixation_count}",
                key="is_fixation",
                use_container_width=True
            ):

                st.session_state[
                    "is_issue_stage"
                ] = "FIXATION IN PROGRESS"

                st.rerun()


    # =====================================================
    # ACTIVE KPI FILTER
    # =====================================================

    selected_card = st.session_state.get(
        "is_issue_stage",
        "All"
    )


    # =====================================================
    # FILTERS
    # =====================================================

    st.html("""
<div style="
color:#006747;
font-size:22px;
font-weight:700;
margin-top:0;
margin-bottom:5px;
font-family:Segoe UI,Arial,sans-serif;">
🎯 Filters
</div>
""")


    c1, c2, c3 = st.columns(3)


    with c1:

        search = st.text_input(
            "Search IS Issues",
            placeholder="🔍 Search application, issue or remarks...",
            key="is_issue_search"
        )


    with c2:

        severity_filter = st.selectbox(
            "Severity",
            [
                "All",
                "HIGH",
                "MEDIUM",
                "LOW"
            ],
            key="is_issue_severity_filter"
        )


    with c3:

        status_filter = st.selectbox(
            "Status",
            [
                "All",
                "Fixed",
                "Fixation in Progress"
            ],
            key="is_issue_status_filter"
        )


    # =====================================================
    # FILTER DATA
    # =====================================================

    filtered_df = df.copy()


    # =====================================================
    # SEARCH
    # =====================================================

    if search:

        search_columns = [
            "Application Name",
            "Severity",
            "Issues Description",
            "Status of Issue",
            "DBG Remarks"
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


    # =====================================================
    # SEVERITY FILTER
    # =====================================================

    if severity_filter != "All":

        filtered_df = filtered_df[
            filtered_df["Severity"]
            .astype(str)
            .str.upper()
            .str.strip()
            == severity_filter
        ]


    # =====================================================
    # STATUS FILTER
    # =====================================================

    if status_filter != "All":

        filtered_df = filtered_df[
            filtered_df["Status of Issue"]
            .astype(str)
            .str.replace(
                r"\s+",
                " ",
                regex=True
            )
            .str.upper()
            .str.strip()
            ==
            status_filter.upper()
        ]


    # =====================================================
    # KPI CARD FILTER
    # =====================================================

    if selected_card != "All":

        if selected_card in [
            "HIGH",
            "MEDIUM",
            "LOW"
        ]:

            filtered_df = filtered_df[
                filtered_df["Severity"]
                .astype(str)
                .str.upper()
                .str.strip()
                == selected_card
            ]


        elif selected_card == "FIXED":

            filtered_df = filtered_df[
                filtered_df["Status of Issue"]
                .astype(str)
                .str.replace(
                    r"\s+",
                    " ",
                    regex=True
                )
                .str.upper()
                .str.strip()
                == "FIXED"
            ]


        elif selected_card == "FIXATION IN PROGRESS":

            filtered_df = filtered_df[
                filtered_df["Status of Issue"]
                .astype(str)
                .str.replace(
                    r"\s+",
                    " ",
                    regex=True
                )
                .str.upper()
                .str.strip()
                == "FIXATION IN PROGRESS"
            ]


    # =====================================================
    # SHOWING COUNT
    # =====================================================

    left, right = st.columns(
        [5, 1]
    )


    with left:

        filter_name = (
            "All"
            if selected_card == "All"
            else selected_card
        )

        st.markdown(
            f"""
<div style="background:#E8F5E9;padding:7px 12px;border-radius:9px;color:#006747;font-size:13px;font-weight:700;margin-top:3px;">
🔐 Showing <b>{len(filtered_df)}</b> Issue(s)
&nbsp;•&nbsp;
Filter: <b>{filter_name}</b>
</div>
""",
            unsafe_allow_html=True
        )


    with right:

        csv = (
            filtered_df
            .to_csv(index=False)
            .encode("utf-8-sig")
        )

        st.download_button(
            "⬇ Download CSV",
            csv,
            "IS_Issues.csv",
            "text/csv",
            width="stretch"
        )


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
📋 IS Issues Details
</div>
""")


    # =====================================================
    # VIP TABLE CSS
    # =====================================================

    st.markdown("""
<style>

.is-issues-table-wrapper {

    background:#FFFFFF;

    border:
        1px solid
        #DDE5E1;

    border-radius:16px;

    padding:6px;

    box-shadow:
        0 6px 20px
        rgba(0,103,71,.08);

    overflow:hidden;
}


.is-issues-table {

    width:100%;

    border-collapse:
        separate;

    border-spacing:0;

    font-size:14px;

    table-layout:fixed;
}


.is-issues-table thead th {

    background:
        linear-gradient(
            135deg,
            #013D2B,
            #006747,
            #008A5A
        );

    color:#FFFFFF;

    font-weight:700;

    padding:14px 12px;

    text-align:left;
}


.is-issues-table thead th:first-child {

    border-top-left-radius:11px;
}


.is-issues-table thead th:last-child {

    border-top-right-radius:11px;
}


.is-issues-table tbody td {

    padding:13px 12px;

    color:#1F2937;

    border-bottom:
        1px solid
        #E5E7EB;

    background:#FFFFFF;

    white-space:
        normal !important;

    word-wrap:
        break-word !important;

    overflow-wrap:
        anywhere !important;

    vertical-align:top;

    line-height:1.6;
}


.is-issues-table tbody tr:nth-child(even) td {

    background:#F8FAFC;
}


.is-issues-table tbody tr:hover td {

    background:#ECFDF5;
}


/* =====================================================
APPLICATION
===================================================== */

.is-application {

    font-weight:700;

    color:#006747 !important;
}


/* =====================================================
SEVERITY BADGES
===================================================== */

.severity-badge {

    display:inline-block;

    padding:5px 11px;

    border-radius:20px;

    font-size:11px;

    font-weight:800;

    white-space:nowrap;
}


.severity-high {

    background:#FEE2E2;

    color:#B91C1C;
}


.severity-medium {

    background:#FEF3C7;

    color:#92400E;
}


.severity-low {

    background:#D1FAE5;

    color:#087443;
}


/* =====================================================
STATUS BADGES
===================================================== */

.issue-status {

    display:inline-block;

    padding:5px 11px;

    border-radius:20px;

    font-size:11px;

    font-weight:800;

    white-space:nowrap;
}


.issue-fixed {

    background:#DCFCE7;

    color:#166534;
}


.issue-progress {

    background:#FEF3C7;

    color:#92400E;
}


/* =====================================================
ACTION BUTTONS
===================================================== */

div.stButton > button {

    background:#006747 !important;

    color:white !important;

    border:
        1px solid
        #006747 !important;

    border-radius:10px !important;

    font-weight:700 !important;
}


div.stButton > button:hover {

    background:#008A5A !important;

    color:white !important;

    border-color:#008A5A !important;
}

</style>
""", unsafe_allow_html=True)


    # =====================================================
    # TABLE DATA
    # =====================================================

    display_df = filtered_df.copy()


    # =====================================================
    # APPLICATION BADGE
    # =====================================================

    display_df["Application Name"] = (
        display_df["Application Name"]
        .apply(
            lambda x:
            f'<span class="is-application">{x}</span>'
        )
    )


    # =====================================================
    # SEVERITY BADGE
    # =====================================================

    def severity_badge(value):

        value = str(value).strip()

        upper = value.upper()


        if upper == "HIGH":

            css = "severity-high"


        elif upper == "MEDIUM":

            css = "severity-medium"


        elif upper == "LOW":

            css = "severity-low"


        else:

            css = ""


        return (
            f'<span class="severity-badge {css}">'
            f'{value}'
            f'</span>'
        )


    display_df["Severity"] = (
        display_df["Severity"]
        .apply(severity_badge)
    )


    # =====================================================
    # STATUS BADGE
    # =====================================================

    def status_badge(value):

        value = str(value).strip()

        upper = (
            value
            .replace("  ", " ")
            .upper()
        )


        if upper == "FIXED":

            css = "issue-fixed"


        else:

            css = "issue-progress"


        return (
            f'<span class="issue-status {css}">'
            f'{value}'
            f'</span>'
        )


    display_df["Status of Issue"] = (
        display_df["Status of Issue"]
        .apply(status_badge)
    )


    # =====================================================
    # HTML TABLE
    # =====================================================

    table_html = display_df.to_html(
        index=False,
        escape=False,
        classes="is-issues-table"
    )


    st.markdown(
        f"""
<div class="is-issues-table-wrapper">
{table_html}
</div>
""",
        unsafe_allow_html=True
    )


    # =====================================================
    # EDIT IS ISSUES
    # =====================================================

    st.markdown(
        "<div style='height:4px;'></div>",
        unsafe_allow_html=True
    )


    if "is_issues_edit_mode" not in st.session_state:

        st.session_state[
            "is_issues_edit_mode"
        ] = False


    # =====================================================
    # EDIT BUTTON
    # =====================================================

    if not st.session_state[
        "is_issues_edit_mode"
    ]:

        if st.button(
            "✏️ Edit IS Issues",
            key="is_issues_edit_button",
            type="secondary",
            use_container_width=True
        ):

            st.session_state[
                "is_issues_edit_mode"
            ] = True

            st.rerun()


    # =====================================================
    # EDITOR
    # =====================================================

    if st.session_state[
        "is_issues_edit_mode"
    ]:

        st.markdown("""
<h2 style="
color:#006747;
font-size:24px;
font-weight:700;
margin-top:14px;
margin-bottom:8px;">
✏️ Edit IS Issues
</h2>
""", unsafe_allow_html=True)


        st.markdown("""
<div style="background:#ECFDF5;border:1px solid #B7E4C7;border-left:5px solid #006747;border-radius:10px;padding:9px 12px;color:#006747;font-size:13px;font-weight:600;margin-bottom:8px;">
✏️ Edit existing records, add new records, or create a new field/column.
</div>
""", unsafe_allow_html=True)


        # =================================================
        # ADD NEW FIELD
        # =================================================

        st.markdown("""
<h3 style="
color:#006747;
font-size:19px;
font-weight:700;
margin-top:10px;
margin-bottom:6px;">
➕ Add New Field
</h3>
""", unsafe_allow_html=True)


        field_col1, field_col2 = st.columns(
            [3, 1]
        )


        with field_col1:

            new_field = st.text_input(
                "New Field Name",
                placeholder="e.g. IS Remarks",
                key="is_issues_new_field_input"
            )


        with field_col2:

            st.markdown(
                "<div style='height:26px;'></div>",
                unsafe_allow_html=True
            )


            if st.button(
                "➕ Add Field",
                key="is_issues_add_field",
                use_container_width=True
            ):

                if new_field.strip():

                    new_field = (
                        new_field
                        .strip()
                    )


                    if new_field not in df.columns:

                        df[new_field] = ""


                        st.session_state[
                            "is_issues_added_fields"
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
        # SEARCH FILTER
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
        # SEVERITY FILTER
        # =================================================

        if severity_filter != "All":

            edit_df = edit_df[
                edit_df["Severity"]
                .astype(str)
                .str.upper()
                .str.strip()
                == severity_filter
            ]


        # =================================================
        # STATUS FILTER
        # =================================================

        if status_filter != "All":

            edit_df = edit_df[
                edit_df["Status of Issue"]
                .astype(str)
                .str.replace(
                    r"\s+",
                    " ",
                    regex=True
                )
                .str.upper()
                .str.strip()
                ==
                status_filter.upper()
            ]


        # =================================================
        # KPI FILTER
        # =================================================

        selected_card = st.session_state.get(
            "is_issue_stage",
            "All"
        )


        if selected_card != "All":

            if selected_card in [
                "HIGH",
                "MEDIUM",
                "LOW"
            ]:

                edit_df = edit_df[
                    edit_df["Severity"]
                    .astype(str)
                    .str.upper()
                    .str.strip()
                    == selected_card
                ]


            elif selected_card == "FIXED":

                edit_df = edit_df[
                    edit_df["Status of Issue"]
                    .astype(str)
                    .str.replace(
                        r"\s+",
                        " ",
                        regex=True
                    )
                    .str.upper()
                    .str.strip()
                    == "FIXED"
                ]


            elif selected_card == "FIXATION IN PROGRESS":

                edit_df = edit_df[
                    edit_df["Status of Issue"]
                    .astype(str)
                    .str.replace(
                        r"\s+",
                        " ",
                        regex=True
                    )
                    .str.upper()
                    .str.strip()
                    == "FIXATION IN PROGRESS"
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

            column_config={

                "Severity":
                    st.column_config.SelectboxColumn(
                        "Severity",
                        options=[
                            "HIGH",
                            "MEDIUM",
                            "LOW"
                        ]
                    ),

                "Status of Issue":
                    st.column_config.SelectboxColumn(
                        "Status of Issue",
                        options=[
                            "Fixed",
                            "Fixation in Progress"
                        ]
                    )
            },

            key="is_issues_editor"
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
                key="is_issues_save",
                use_container_width=True
            ):

                try:

                    # -------------------------------------
                    # READ ORIGINAL EXCEL
                    # -------------------------------------

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


                    # -------------------------------------
                    # ADD NEW COLUMNS
                    # -------------------------------------

                    for col in edited_df.columns:

                        if col not in original.columns:

                            original[col] = ""


                    # -------------------------------------
                    # UPDATE EXISTING ROWS
                    # -------------------------------------

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


                    # -------------------------------------
                    # SAVE EXCEL
                    # -------------------------------------

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


                    st.success(
                        "✅ IS Issues changes, new rows and new fields saved successfully!"
                    )


                    st.session_state[
                        "is_issues_edit_mode"
                    ] = False


                    if "is_issues_editor" in st.session_state:

                        del st.session_state[
                            "is_issues_editor"
                        ]


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
                key="is_issues_close_editor",
                use_container_width=True
            ):

                st.session_state[
                    "is_issues_edit_mode"
                ] = False


                if "is_issues_editor" in st.session_state:

                    del st.session_state[
                        "is_issues_editor"
                    ]


                st.rerun()
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

    render_header('CRPL MONITORING', '🏦 CRPL', 'CRPL Issues &amp; Progress Tracking', 'CRPL MONITORING')

    # =====================================================
    # FILTERS
    # =====================================================

    st.html("""
    <div style="
    color:#006747;
    font-size:22px;
    font-weight:700;
    margin-top:0;
    margin-bottom:5px;
    font-family:Segoe UI,Arial,sans-serif;">
    🎯 Filters
    </div>
    """)


    c1, c2 = st.columns(2)


    with c1:

        search = st.text_input(
            "Search CRF",
            placeholder="🔍 Search CRF No or CRF Name...",
            key="crpl_search"
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
            ],
            key="crpl_stage_filter"
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
    margin-bottom:5px;
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
        box-shadow:0 6px 20px rgba(0,103,71,.08);
        overflow-x:auto;
    }

    .crpl-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        font-size:14px;
        table-layout:fixed;
    }

    .crpl-table thead th {
        background:linear-gradient(135deg,#013D2B,#006747,#008A5A);
        color:#FFFFFF;
        font-weight:700;
        padding:14px 12px;
        text-align:left;
        border:none;
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
        background:#E8F5E9;
        color:#006747;
    }

    .stage-wip {
        background:#D1FAE5;
        color:#087443;
    }

    .stage-uat {
        background:#B7E4C7;
        color:#087443;
    }

    .stage-live {
        background:#DCFCE7;
        color:#166534;
    }

    .stage-default {
        background:#F0FDF4;
        color:#2E7D5B;
    }


    /* =====================================================
    EDITOR
    ===================================================== */

    .crpl-edit-info {
        background:#ECFDF5;
        border:1px solid #B7E4C7;
        border-left:5px solid #006747;
        border-radius:10px;
        padding:9px 12px;
        color:#006747;
        font-size:13px;
        font-weight:600;
        margin-bottom:8px;
    }


    /* =====================================================
    ADD FIELD BOX
    ===================================================== */

    .crpl-field-box {
        background:#FFFFFF;
        border:1px solid #DDE5E1;
        border-radius:12px;
        padding:10px 12px;
        box-shadow:0 4px 12px rgba(0,103,71,.06);
        margin-bottom:8px;
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
        "<div style='height:4px;'></div>",
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
    font-size:24px;
    font-weight:700;
    margin-top:14px;
    margin-bottom:8px;">
    ✏️ Edit CRPL
    </h2>
    """, unsafe_allow_html=True)


        st.markdown("""
    <div class="crpl-edit-info">
    ✏️ Edit existing records, add new records, or create a new field/column.
    </div>
    """, unsafe_allow_html=True)


        # =================================================
        # ADD NEW FIELD
        # =================================================

        st.markdown("""
    <h3 style="
    color:#006747;
    font-size:19px;
    font-weight:700;
    margin-top:10px;
    margin-bottom:6px;">
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

            st.markdown(
                "<div style='height:26px;'></div>",
                unsafe_allow_html=True
            )

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
                        "✅ CRPL changes, new rows and new fields saved successfully!"
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

    # =====================================================
    # REQUIRED COLUMNS
    # =====================================================

    columns = [
        "Project / Issue",
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

    render_header('PAYSYS MONITORING', '💳 PAYSYS', 'PAYSYS Issues &amp; Progress Tracking', 'PAYSYS MONITORING')


    # =====================================================
    # FILTERS
    # =====================================================

    st.html("""
    <div style="
    color:#006747;
    font-size:22px;
    font-weight:700;
    margin-top:0;
    margin-bottom:5px;
    font-family:Segoe UI,Arial,sans-serif;">
    🎯 Filters
    </div>
    """)


    c1, c2 = st.columns(2)


    with c1:

        search = st.text_input(
            "Search PAYSYS",
            placeholder="🔍 Search PAYSYS remarks, response or status...",
            key="paysys_search"
        )


    with c2:

        stage_filter = st.selectbox(
            "Stage",
            [
                "All",
                "UAT",
                "LIVE"
            ],
            key="paysys_stage_filter"
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
        box-shadow:0 6px 20px rgba(0,103,71,.08);
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
        background:linear-gradient(135deg,#013D2B,#006747,#008A5A);
        color:#FFFFFF;
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
        background:#D1FAE5;
        color:#087443;
    }

    .stage-live {
        background:#DCFCE7;
        color:#166534;
    }

    .stage-default {
        background:#F0FDF4;
        color:#2E7D5B;
    }


    /* =====================================================
    PAYSYS KPI / ACTION BUTTONS
    ===================================================== */

    div.stButton > button {
        background:#006747 !important;
        color:white !important;
        border:1px solid #006747 !important;
        border-radius:10px !important;
        font-weight:700 !important;
    }

    div.stButton > button:hover {
        background:#008A5A !important;
        color:white !important;
        border-color:#008A5A !important;
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
        "<div style='height:4px;'></div>",
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
    font-size:24px;
    font-weight:700;
    margin-top:14px;
    margin-bottom:8px;">
    ✏️ Edit PAYSYS
    </h2>
    """, unsafe_allow_html=True)


        st.markdown("""
    <div style="background:#ECFDF5;border:1px solid #B7E4C7;border-left:5px solid #006747;border-radius:10px;padding:9px 12px;color:#006747;font-size:13px;font-weight:600;margin-bottom:8px;">
    ✏️ Edit existing records, add new records, or create a new field/column.
    </div>
    """, unsafe_allow_html=True)


        # =================================================
        # ADD NEW FIELD
        # =================================================

        st.markdown("""
    <h3 style="
    color:#006747;
    font-size:19px;
    font-weight:700;
    margin-top:10px;
    margin-bottom:6px;">
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
                "<div style='height:26px;'></div>",
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

    render_header('SMARTPAY PROJECT MANAGEMENT', '📑 Executive Report Center', 'Generate, preview and download SmartPay project reports.', 'REPORT CENTER')

   # ==========================================
    # REPORT INFORMATION
    # ==========================================

    st.html("""
    <div style="background:#FFFFFF;border-radius:14px;padding:12px 16px;border:1px solid #DDE5E1;box-shadow:0 4px 12px rgba(0,103,71,.06);margin-bottom:8px;font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#006747;font-size:18px;font-weight:700;margin-bottom:8px;">
    📋 Report Includes
    </div>

    <div style="color:#374151;font-size:13px;line-height:1.7;">

    <span style="color:#006747;font-weight:700;">✅</span> Dashboard Summary<br>
    <span style="color:#006747;font-weight:700;">✅</span> KPI Overview<br>
    <span style="color:#006747;font-weight:700;">✅</span> Project Details<br>
    <span style="color:#006747;font-weight:700;">✅</span> Team Performance<br>
    <span style="color:#006747;font-weight:700;">✅</span> Analytics Summary<br>
    <span style="color:#006747;font-weight:700;">✅</span> Project Timeline

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
    <div style="background:#ECFDF5;border:1px solid #B7E4C7;border-radius:12px;padding:9px 14px;margin-bottom:8px;font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#006747;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;">
    Generated On
    </div>

    <div style="color:#006747;font-size:15px;font-weight:700;margin-top:3px;">
    {generated_time}
    </div>

    </div>
    """)


    # ==========================================
    # EXPORT OPTIONS
    # ==========================================

    st.html("""
    <div style="color:#006747;font-size:24px;font-weight:700;margin-top:2px;margin-bottom:5px;font-family:Segoe UI,Arial,sans-serif;">
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
                textColor=colors.HexColor("#006747"),
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
                textColor=colors.HexColor("#006747")
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
                colors.HexColor("#B7E4C7")
            ),
            (
                "BACKGROUND",
                (0, 1),
                (-1, -1),
                colors.HexColor("#ECFDF5")
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


        columns = list(
            export_pdf_df.columns
        )


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


            table_data.append(
                row_data
            )


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
                    colors.HexColor("#B7E4C7")
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


        pdf_story.append(
            pdf_table
        )


        pdf_story.append(
            Spacer(1, 0.15 * inch)
        )


        pdf_story.append(
            PageBreak()
        )


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

    pdf_doc.build(
        pdf_story
    )


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
    <div style="background:#FFFFFF;border-radius:14px;padding:12px 14px;border:1px solid #DDE5E1;border-top:4px solid #006747;box-shadow:0 4px 12px rgba(0,103,71,.06);font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#006747;font-size:16px;font-weight:700;">
    SmartPay Projects
    </div>

    <div style="color:#6B7280;font-size:12px;margin-top:4px;">
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
    <div style="background:linear-gradient(135deg,#ECFDF5,#FFFFFF);border-radius:14px;padding:12px 16px;border:1px solid #B7E4C7;border-left:4px solid #006747;box-shadow:0 4px 12px rgba(0,103,71,.06);font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#006747;font-size:18px;font-weight:700;">
    📕 Exclusive Executive PDF Report
    </div>

    <div style="color:#6B7280;font-size:12px;margin-top:4px;">
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
    <div style="background:#FFFFFF;border-radius:14px;padding:12px 14px;border:1px solid #DDE5E1;border-top:4px solid #008A5A;box-shadow:0 4px 12px rgba(0,103,71,.06);font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#008A5A;font-size:16px;font-weight:700;">
    CRPL
    </div>

    <div style="color:#6B7280;font-size:12px;margin-top:4px;">
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
    <div style="background:#FFFFFF;border-radius:14px;padding:12px 14px;border:1px solid #DDE5E1;border-top:4px solid #20A464;box-shadow:0 4px 12px rgba(0,103,71,.06);font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#20A464;font-size:16px;font-weight:700;">
    PAYSYS
    </div>

    <div style="color:#6B7280;font-size:12px;margin-top:4px;">
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
    <div style="background:linear-gradient(135deg,#ECFDF5,#FFFFFF);border-radius:14px;padding:12px 16px;border:1px solid #B7E4C7;border-left:4px solid #006747;box-shadow:0 4px 12px rgba(0,103,71,.06);font-family:Segoe UI,Arial,sans-serif;">

    <div style="color:#006747;font-size:18px;font-weight:700;">
    📊 Complete Executive Excel Report
    </div>

    <div style="color:#6B7280;font-size:12px;margin-top:4px;">
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
    <div style="color:#006747;font-size:24px;font-weight:700;margin-top:2px;margin-bottom:5px;font-family:Segoe UI,Arial,sans-serif;">
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
    <div style="color:#006747;font-size:18px;font-weight:700;margin-bottom:5px;font-family:Segoe UI,Arial,sans-serif;">
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
    <div style="color:#008A5A;font-size:18px;font-weight:700;margin-bottom:5px;font-family:Segoe UI,Arial,sans-serif;">
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
    <div style="color:#20A464;font-size:18px;font-weight:700;margin-bottom:5px;font-family:Segoe UI,Arial,sans-serif;">
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

# =====================================================
# GLOBAL FOOTER (all pages)
# =====================================================

st.markdown("""
<div style="margin-top:26px;padding:12px 20px;border-radius:14px;background:linear-gradient(135deg,#013D2B,#006747);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
<div style="color:white;font-size:12px;font-weight:600;">NBP &nbsp;|&nbsp; SmartPay Project Dashboard &nbsp;|&nbsp; Digital Banking Group</div>
<div style="color:#D1FAE5;font-size:12px;font-style:italic;">DBG Business Banking &nbsp;•&nbsp; Developed by: Muhammad Saad Asif</div>
</div>
""", unsafe_allow_html=True)