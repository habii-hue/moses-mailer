import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
from pypdf import PdfReader, PdfWriter
import io
import zipfile

# --- Page Configuration ---
st.set_page_config(page_title="MOSES Manager", page_icon="🧾", layout="wide", initial_sidebar_state="expanded")

# --- Custom CSS for Desktop UI Mimicry ---
st.markdown("""
    <style>
    /* Sidebar Branding */
    [data-testid="stSidebar"] {
        background-color: #181818;
    }
    
    /* --------------------------------------
       CUSTOM TAB NAVIGATION (Radio Overrides)
       -------------------------------------- */
    /* Hide the radio button circles completely */
    div[role="radiogroup"] label > div:first-child {
        display: none !important;
    }
    /* Style the radio button container (the tab) */
    div[role="radiogroup"] label {
        padding: 12px 20px;
        border-radius: 8px;
        margin-bottom: 5px;
        cursor: pointer;
        transition: background-color 0.2s ease;
    }
    /* Hover state for inactive tabs */
    div[role="radiogroup"] label:hover {
        background-color: #1E1E1E;
    }
    /* Active tab styling using the modern :has() selector */
    div[role="radiogroup"] label:has(input:checked) {
        background-color: #252526;
    }
    /* Make the text match the Tkinter font weight */
    div[role="radiogroup"] label p {
        font-size: 15px;
        font-weight: 700;
        margin: 0;
    }
    /* -------------------------------------- */

    /* Card Containers */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #252526;
        border-radius: 8px;
        border: none;
        padding: 5px 15px;
        margin-bottom: 20px;
    }
    /* Clean File Uploaders */
    .stFileUploader label {
        display: none;
    }
    /* System Console Log Box */
    .console-box {
        background-color: #1E1E1E;
        color: #A9B7C6;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        padding: 15px;
        border-radius: 8px;
        height: 180px;
        overflow-y: auto;
        white-space: pre-wrap;
        border: 1px solid #333;
    }
    </style>
""", unsafe_allow_html=True)

# --- Session State for Live Console Logs ---
if "split_logs" not in st.session_state:
    st.session_state.split_logs = "> System initialized. Smart PDF Scanner & Excel Engine online.\n"
if "mail_logs" not in st.session_state:
    st.session_state.mail_logs = "> System initialized. Mail Engine online.\n"

def log_split(msg):
    st.session_state.split_logs += f"> {msg}\n"
def log_mail(msg):
    st.session_state.mail_logs += f"> {msg}\n"

# ==========================================
# SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.markdown("<h1 style='color: #FF69B4; font-weight: 900; font-size: 32px; margin-bottom: -15px;'>MOSES</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: gray; font-size: 14px; margin-bottom: 30px;'>Automated Manager</p>", unsafe_allow_html=True)
    
    view = st.radio("Navigation", ["Split Receipts", "Bulk Mailer"], label_visibility="collapsed")

# ==========================================
# TAB 1: RECEIPT SPLITTER
# ==========================================
if view == "Split Receipts":
    st.markdown("<h2 style='margin-bottom: 20px;'>Receipt Splitter</h2>", unsafe_allow_html=True)
    
    with st.container(border=True):
        st.markdown("**Master PDF Document**")
        master_pdf = st.file_uploader("pdf_split", type=["pdf"], key="pdf_s")
        
        st.markdown("<br>**Member Masterlist (Excel or CSV)**", unsafe_allow_html=True)
        master_list = st.file_uploader("list_split", type=["xlsx", "csv"], key="list_s")
        
        st.markdown("<br>**Output Destination**", unsafe_allow_html=True)
        st.info("Web execution mode: Renamed files will be packaged into a downloadable ZIP archive.")
        
    action_col, _ = st.columns([1, 1]) 
    
    with action_col:
        if st.button("Extract & Rename PDFs", type="primary", use_container_width=True):
            if not master_pdf or not master_list:
                st.error("Missing files. Please upload both the PDF and the Masterlist.")
            else:
                log_split("Parsing Data...")
                
                if master_list.name.endswith('.csv'):
                    df = pd.read_csv(master_list)
                else:
                    df = pd.read_excel(master_list)
                    
                name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                contacts = df.dropna(subset=[name_col])[name_col].astype(str).tolist()
                
                log_split(f"Scanning pages inside {master_pdf.name}...")
                reader = PdfReader(master_pdf)
                zip_buffer = io.BytesIO()
                
                matched_count = 0
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for i in range(len(reader.pages)):
                        page = reader.pages[i]
                        extracted_text = page.extract_text() or ""
                        
                        found_name = None
                        for name in
