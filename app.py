import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import fitz  # PyMuPDF
import io
import zipfile
import openpyxl
import unicodedata
import re
from difflib import SequenceMatcher

# --- Page Configuration ---
st.set_page_config(page_title="MOSES · Automated Manager", page_icon="⚡", layout="wide", initial_sidebar_state="expanded")

# --- Initialize Session State ---
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "split"
if "split_logs" not in st.session_state:
    st.session_state.split_logs = "00:00:01 [INFO]  System initialized. 4-Tier Smart PDF Scanner & Excel Engine online.\n"
if "mail_logs" not in st.session_state:
    st.session_state.mail_logs = "00:00:01 [INFO]  System initialized. Multi-Attachment Mail Engine online.\n"

def render_console(placeholder, log_text, badge_text="READY"):
    placeholder.markdown(f"""
        <div class="shadcn-terminal">
            <div class="shadcn-terminal-header">
                <div class="terminal-dots">
                    <span class="dot dot-red"></span>
                    <span class="dot dot-amber"></span>
                    <span class="dot dot-green"></span>
                    <span class="terminal-title">TRACE · SYSTEM LOGS</span>
                </div>
                <span class="shadcn-badge-outline">{badge_text}</span>
            </div>
            <div class="shadcn-terminal-body" id="log-container">{log_text}</div>
        </div>
        <script>
            var box = document.getElementById("log-container");
            if (box) {{ box.scrollTop = box.scrollHeight; }}
        </script>
    """, unsafe_allow_html=True)

def log_split(msg, placeholder=None, badge="RUNNING"):
    st.session_state.split_logs += f"› {msg}\n"
    if placeholder:
        render_console(placeholder, st.session_state.split_logs, badge)

def log_mail(msg, placeholder=None, badge="RUNNING"):
    st.session_state.mail_logs += f"› {msg}\n"
    if placeholder:
        render_console(placeholder, st.session_state.mail_logs, badge)

# ==========================================
# SMART TEXT NORMALIZER & 4-TIER MATCHER
# ==========================================
def normalize_text(s):
    """Flattens newlines, strips accents (Ñ -> N), and removes punctuation/extra spaces."""
    if not s:
        return ""
    s = unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('ASCII')
    s = s.lower()
    s = re.sub(r'[^a-z0-9\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def extract_clean_page_text(page):
    """Extracts text block-by-block so multi-column Canva layouts never interleave."""
    blocks = page.get_text("blocks") or []
    blocks_sorted = sorted(blocks, key=lambda b: (round(b[1], 1), round(b[0], 1)))
    block_texts = [b[4] for b in blocks_sorted if len(b) > 4 and isinstance(b[4], str)]
    return "\n".join(block_texts) + "\n" + (page.get_text("text") or "")

def find_matching_contact(page_text, contacts, name_key=None):
    """
    4-Tier Smart Matcher:
    Tier 1: Normalized full-name substring match.
    Tier 2: Last Name + First Name prefix match.
    Tier 3: Order-independent token bag match (handles 'First Middle Last' vs 'Last, First' & column wraps).
    Tier 4: Fuzzy token match (handles 1-2 letter typos like PASUCAL vs PASCUAL or Marrgaret vs Margaret).
    """
    norm_page = normalize_text(page_text)
    if not norm_page:
        return None
    page_tokens = set(norm_page.split())
    page_words_list = norm_page.split()

    # Tier 1: Exact normalized full name
    for c in contacts:
        raw_name = str(c[name_key]).strip() if name_key else str(c).strip()
        norm_name = normalize_text(raw_name)
        if norm_name and norm_name in norm_page:
            return c

    # Tier 2: Match Last Name + First Name in sequence
    for c in contacts:
        raw_name = str(c[name_key]).strip() if name_key else str(c).strip()
        if "," in raw_name:
            last_part, given_part = raw_name.split(",", 1)
            norm_last = normalize_text(last_part)
            given_tokens = [t for t in normalize_text(given_part).split() if len(t) > 1]
            first_tokens = " ".join(given_tokens[:2]) if len(given_tokens) >= 2 else " ".join(given_tokens[:1])
            search_prefix = f"{norm_last} {first_tokens}".strip()
            if search_prefix and search_prefix in norm_page:
                return c

    # Tier 3: Order-Independent Token Match
    best_tier3 = None
    best_tier3_score = 0
    for c in contacts:
        raw_name = str(c[name_key]).strip() if name_key else str(c).strip()
        if "," in raw_name:
            last_part, given_part = raw_name.split(",", 1)
            last_tokens = [t for t in normalize_text(last_part).split() if len(t) > 1]
            given_tokens = [t for t in normalize_text(given_part).split() if len(t) > 1]
            if last_tokens and all(lt in page_tokens for lt in last_tokens):
                matched_given = [gt for gt in given_tokens if gt in page_tokens]
                required_given = 2 if len(given_tokens) >= 3 else 1
                if len(matched_given) >= required_given:
                    score = len(last_tokens) * 10 + len(matched_given)
                    if score > best_tier3_score:
                        best_tier3_score = score
                        best_tier3 = c
    if best_tier3 is not None:
        return best_tier3

    # Tier 4: Fuzzy Token Match
    def token_fuzzy_in_page(tok, threshold=0.82):
        if tok in page_tokens:
            return True
        return any(SequenceMatcher(None, tok, pw).ratio() >= threshold for pw in page_words_list if abs(len(tok) - len(pw)) <= 2)

    best_tier4 = None
    best_tier4_score = 0
    for c in contacts:
        raw_name = str(c[name_key]).strip() if name_key else str(c).strip()
        if "," in raw_name:
            last_part, given_part = raw_name.split(",", 1)
            last_tokens = [t for t in normalize_text(last_part).split() if len(t) > 1]
            given_tokens = [t for t in normalize_text(given_part).split() if len(t) > 1]
            if last_tokens and all(token_fuzzy_in_page(lt) for lt in last_tokens):
                matched_given = [gt for gt in given_tokens if token_fuzzy_in_page(gt)]
                required_given = 2 if len(given_tokens) >= 2 else 1
                if len(matched_given) >= required_given:
                    score = len(last_tokens) * 10 + len(matched_given)
                    if score > best_tier4_score:
                        best_tier4_score = score
                        best_tier4 = c
    return best_tier4

def detect_installment(text):
    """Reads the PDF heading to identify which installment receipt is on the page."""
    upper_text = (text or "").upper()
    if "FIRST INSTALLMENT" in upper_text or "1ST INSTALLMENT" in upper_text:
        return "First Installment"
    elif "SECOND INSTALLMENT" in upper_text or "2ND INSTALLMENT" in upper_text:
        return "Second Installment"
    elif "THIRD INSTALLMENT" in upper_text or "3RD INSTALLMENT" in upper_text:
        return "Third Installment"
    return None

def get_excel_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Members"
    ws.append(["Membership Type", "Name", "Section", "Amount Paid", "Email"])
    ws.append(["Normal", "Dela Cruz, Juan", "BSCS 3-1", 200, "juan.delacruz@example.com"])
    ws.append(["Reads", "Santos, Maria", "BSA 2-1", 150, "maria.santos@example.com"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# --- Dynamic Sidebar State ---
split_bg = "#26221E" if st.session_state.active_tab == "split" else "transparent"
split_border = "#3A332C" if st.session_state.active_tab == "split" else "transparent"
split_color = "#F1A651" if st.session_state.active_tab == "split" else "#C5B4A0"

mail_bg = "#26221E" if st.session_state.active_tab == "mail" else "transparent"
mail_border = "#3A332C" if st.session_state.active_tab == "mail" else "transparent"
mail_color = "#F1A651" if st.session_state.active_tab == "mail" else "#C5B4A0"

# ==========================================
# CODA.IO x SHADCN/UI DESIGN SYSTEM CSS
# ==========================================
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

    :root {{
        --background: #141210;
        --foreground: #F5F2EB;
        --card: #1C1917;
        --card-foreground: #F5F2EB;
        --muted: #26221E;
        --muted-foreground: #C5B4A0;
        --border: #2E2924;
        --input: #221F1B;
        --primary: #F1A651;
        --primary-hover: #E0923B;
        --primary-foreground: #141210;
        --accent: #9F5556;
        --ink: #FCD47C;
        --radius: 8px;
    }}

    .stApp {{
        background-color: var(--background);
        color: var(--foreground);
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        letter-spacing: -0.14px;
    }}

    header[data-testid="stHeader"] {{
        background: transparent !important;
    }}
    .stDeployButton {{
        display: none !important;
    }}

    .block-container {{
        padding-top: 2rem !important;
        padding-left: 3.2rem !important;
        padding-right: 3.2rem !important;
        max-width: 1380px !important;
        width: 95% !important;
    }}

    /* --- shadcn Sidebar --- */
    [data-testid="stSidebar"] {{
        background-color: #100E0C !important;
        border-right: 1px solid var(--border) !important;
        min-width: 260px !important;
        max-width: 260px !important;
    }}

    [data-testid="stSidebar"] .stButton > button {{
        width: 100%;
        justify-content: flex-start;
        text-align: left;
        padding: 10px 14px !important;
        border-radius: 6px !important;
        height: 40px !important;
        box-shadow: none !important;
        transition: all 0.15s ease;
    }}
    [data-testid="stSidebar"] .stButton > button p {{
        font-size: 13.5px !important;
        font-weight: 600 !important;
        letter-spacing: -0.14px !important;
        margin: 0 !important;
    }}

    /* Sidebar Item 1: Split Receipts */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button {{
        background-color: {split_bg} !important;
        border: 1px solid {split_border} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button p {{
        color: {split_color} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button:hover {{
        background-color: #26221E !important;
    }}

    /* Sidebar Item 2: Bulk Mailer */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button {{
        background-color: {mail_bg} !important;
        border: 1px solid {mail_border} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button p {{
        color: {mail_color} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button:hover {{
        background-color: #26221E !important;
    }}

    /* --- shadcn Card Container --- */
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: var(--card) !important;
        border-radius: 10px !important;
        border: 1px solid var(--border) !important;
        padding: 24px 28px !important;
        margin-bottom: 20px !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.35) !important;
    }}

    /* --- shadcn Primary Button (Coda Dominant #f1a651) --- */
    div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {{
        background-color: var(--primary) !important;
        color: #141210 !important;
        border: 1px solid #FCD47C !important;
        border-radius: 7px !important;
        height: 44px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        letter-spacing: -0.14px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.25) !important;
        transition: all 0.15s ease !important;
    }}
    div.stButton > button[kind="primary"] p, div.stDownloadButton > button[kind="primary"] p {{
        color: #141210 !important;
        font-weight: 600 !important;
    }}
    div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover {{
        background-color: var(--primary-hover) !important;
        transform: translateY(-1px);
    }}

    /* --- shadcn Secondary Outline Button (Get Template) --- */
    div.stDownloadButton > button[kind="secondary"] {{
        background-color: #221F1B !important;
        color: var(--foreground) !important;
        border: 1px solid var(--border) !important;
        border-radius: 6px !important;
        padding: 6px 14px !important;
        height: 36px !important;
        font-size: 13px !important;
        font-weight: 500 !important;
    }}
    div.stDownloadButton > button[kind="secondary"]:hover {{
        background-color: var(--muted) !important;
        border-color: var(--primary) !important;
    }}

    /* --- shadcn File Uploader Dropzone --- */
    .stFileUploader label {{
        display: none !important;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        background-color: #161412 !important;
        border: 1px dashed #3A332C !important;
        border-radius: 8px !important;
        padding: 16px 20px !important;
        min-height: 74px !important;
        transition: border-color 0.15s ease, background-color 0.15s ease !important;
    }}
    [data-testid="stFileUploaderDropzone"]:hover {{
        border-color: var(--primary) !important;
        background-color: #1E1B18 !important;
    }}

    /* --- shadcn Input --- */
    .stTextInput input {{
        background-color: #141210 !important;
        border: 1px solid var(--border) !important;
        color: var(--foreground) !important;
        border-radius: 6px !important;
        padding: 9px 13px !important;
        font-size: 13.5px !important;
        transition: border-color 0.15s ease !important;
    }}
    .stTextInput input:focus {{
        border-color: var(--primary) !important;
        box-shadow: 0 0 0 2px rgba(241, 166, 81, 0.2) !important;
    }}

    /* --- shadcn Progress Bar --- */
    .stProgress > div > div > div > div {{
        background: linear-gradient(90deg, #9F5556 0%, #F1A651 65%, #FCD47C 100%) !important;
    }}

    /* --- Custom Coda Typography & shadcn Badges/Breadcrumb --- */
    .shadcn-breadcrumb {{
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 12.5px;
        color: var(--muted-foreground);
        margin-bottom: 10px;
        font-weight: 500;
    }}
    .shadcn-breadcrumb .crumb-active {{
        color: var(--foreground);
        font-weight: 600;
    }}
    .coda-h2 {{
        font-size: 34px;
        font-weight: 700;
        line-height: 1.1;
        letter-spacing: -0.95px;
        color: var(--foreground);
        margin: 0 0 6px 0;
    }}
    .coda-subtitle {{
        font-size: 14px;
        color: var(--muted-foreground);
        margin-bottom: 22px;
        line-height: 1.45;
    }}
    .shadcn-badge {{
        display: inline-flex;
        align-items: center;
        padding: 2px 9px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 600;
        background-color: rgba(241, 166, 81, 0.14);
        color: var(--ink);
        border: 1px solid rgba(241, 166, 81, 0.32);
    }}
    .shadcn-badge-terracotta {{
        display: inline-flex;
        align-items: center;
        padding: 2px 9px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 600;
        background-color: rgba(159, 85, 86, 0.22);
        color: #E6A8A9;
        border: 1px solid rgba(159, 85, 86, 0.45);
    }}
    .shadcn-badge-outline {{
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        border-radius: 5px;
        font-size: 10.5px;
        font-family: 'JetBrains Mono', monospace;
        font-weight: 600;
        color: var(--ink);
        background-color: #26221E;
        border: 1px solid #3A332C;
    }}
    .field-label {{
        font-weight: 600;
        font-size: 14px;
        color: var(--foreground);
        margin-bottom: 3px;
        display: flex;
        align-items: center;
        gap: 8px;
    }}
    .field-desc {{
        font-size: 12.5px;
        color: var(--muted-foreground);
        margin-bottom: 8px;
        line-height: 1.4;
    }}
    .file-pill {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 3px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-family: 'JetBrains Mono', monospace;
        background-color: #141210;
        border: 1px solid var(--border);
        margin-bottom: 10px;
    }}
    .shadcn-separator {{
        border: none;
        border-top: 1px solid var(--border);
        margin: 18px 0;
    }}

    /* --- shadcn Terminal Trace Component --- */
    .shadcn-terminal {{
        background-color: #100E0C;
        border: 1px solid var(--border);
        border-radius: 10px;
        overflow: hidden;
        margin-top: 20px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
    }}
    .shadcn-terminal-header {{
        background-color: #181512;
        border-bottom: 1px solid var(--border);
        padding: 10px 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }}
    .terminal-dots {{
        display: flex;
        align-items: center;
        gap: 7px;
    }}
    .dot {{
        width: 10px;
        height: 10px;
        border-radius: 50%;
        display: inline-block;
    }}
    .dot-red {{ background-color: #9F5556; }}
    .dot-amber {{ background-color: #F1A651; }}
    .dot-green {{ background-color: #5E8C61; }}
    .terminal-title {{
        margin-left: 8px;
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px;
        font-weight: 600;
        color: var(--muted-foreground);
        letter-spacing: 0.4px;
    }}
    .shadcn-terminal-body {{
        padding: 16px 20px;
        color: #D6C7B2;
        font-family: 'JetBrains Mono', monospace;
        font-size: 12.5px;
        height: 170px;
        overflow-y: auto;
        white-space: pre-wrap;
        line-height: 1.6;
    }}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR NAVIGATION (shadcn Sidebar)
# ==========================================
with st.sidebar:
    st.markdown("""
        <div style="padding: 4px 4px 20px 4px; border-bottom: 1px solid #2E2924; margin-bottom: 18px;">
            <div style="display: flex; align-items: center; justify-content: space-between;">
                <span style="color: #F1A651; font-weight: 800; font-size: 24px; letter-spacing: -0.8px;">MOSES</span>
                <span class="shadcn-badge">v2.4</span>
            </div>
            <div style="color: #C5B4A0; font-size: 12.5px; margin-top: 2px;">Document & Mail Automation</div>
        </div>
        <div style="font-size: 11px; font-weight: 600; color: #8C7E6E; letter-spacing: 0.6px; text-transform: uppercase; padding: 0 4px 8px 4px;">
            Workspace Modules
        </div>
    """, unsafe_allow_html=True)
    
    if st.button("📄  Receipt Splitter", key="nav_split", use_container_width=True):
        st.session_state.active_tab = "split"
        st.rerun()
        
    if st.button("✉️  Bulk Mailer", key="nav_mail", use_container_width=True):
        st.session_state.active_tab = "mail"
        st.rerun()

    st.markdown("""
        <div style="margin-top: 32px; padding: 14px; background-color: #181512; border: 1px solid #2E2924; border-radius: 8px;">
            <div style="font-size: 11.5px; font-weight: 600; color: #FCD47C; margin-bottom: 6px;">⚡ 4-Tier Engine Active</div>
            <div style="font-size: 11.5px; color: #C5B4A0; line-height: 1.45;">
                Handles multi-line wraps, Ñ characters, reversed name orders, and 1st/2nd/3rd installment tags.
            </div>
        </div>
    """, unsafe_allow_html=True)

# ==========================================
# VIEW 1: SPLIT RECEIPTS
# ==========================================
if st.session_state.active_tab == "split":
    st.markdown("""
        <div class="shadcn-breadcrumb">
            <span>Workspace</span> <span>/</span> <span>MOSES</span> <span>/</span> <span class="crumb-active">Receipt Splitter</span>
            <span class="shadcn-badge" style="margin-left: 8px;">PyMuPDF C-Engine</span>
            <span class="shadcn-badge-terracotta">Multi-Installment Aware</span>
        </div>
        <h2 class="coda-h2">Receipt Splitter</h2>
        <div class="coda-subtitle">Scan multi-page Canva receipt PDFs, match member rosters via 4-tier fuzzy token resolution, and export deduplicated archives.</div>
    """, unsafe_allow_html=True)
    
    with st.container(border=True):
        # Row 1: Master PDF
        c1, c2 = st.columns([1, 1.35], gap="large")
        with c1:
            st.markdown("<div class='field-label'>Master PDF Document <span class='shadcn-badge-outline'>.PDF</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Upload the consolidated Canva receipt PDF containing 1st, 2nd, or 3rd installment pages.</div>", unsafe_allow_html=True)
            pdf_status_slot = st.empty()
        with c2:
            master_pdf = st.file_uploader("pdf_split", type=["pdf"], key="pdf_s")
        pdf_status_slot.markdown(
            f"<div class='file-pill' style='color: {'#FCD47C' if master_pdf else '#8C7E6E'};'>"
            f"{'● ' + master_pdf.name if master_pdf else '○ No file selected'}</div>",
            unsafe_allow_html=True
        )

        st.markdown("<hr class='shadcn-separator'>", unsafe_allow_html=True)

        # Row 2: Member Masterlist + Get Template
        c3, c4 = st.columns([1, 1.35], gap="large")
        with c3:
            st.markdown("<div class='field-label'>Member Masterlist <span class='shadcn-badge-outline'>.XLSX / .CSV</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Spreadsheet containing member names and email addresses.</div>", unsafe_allow_html=True)
            list_status_slot = st.empty()
            st.download_button(
                label="⬇ Get Template",
                data=get_excel_template(),
                file_name="MOSES_MasterList_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary"
            )
        with c4:
            master_list = st.file_uploader("list_split", type=["xlsx", "csv"], key="list_s")
        list_status_slot.markdown(
            f"<div class='file-pill' style='color: {'#FCD47C' if master_list else '#8C7E6E'};'>"
            f"{'● ' + master_list.name if master_list else '○ No file selected'}</div>",
            unsafe_allow_html=True
        )

        st.markdown("<hr class='shadcn-separator'>", unsafe_allow_html=True)

        # Row 3: Output Destination
        c5, c6 = st.columns([1, 1.35], gap="large")
        with c5:
            st.markdown("<div class='field-label'>Output Destination Archive <span class='shadcn-badge-outline'>.ZIP</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Automatically tags 1st, 2nd, and 3rd Installment headings so repeated names never overwrite.</div>", unsafe_allow_html=True)
        with c6:
            zip_name = st.text_input("Archive Name", value="MOSES_Extracted_Receipts.zip", label_visibility="collapsed")

    extract_clicked = st.button("Extract & Rename PDFs", type="primary", use_container_width=True)
    status_container = st.empty()
    progress_slot = st.empty()
    download_slot = st.empty()
    console_slot = st.empty()
    render_console(console_slot, st.session_state.split_logs, "IDLE")

    if extract_clicked:
        if not master_pdf or not master_list:
            log_split("[ERROR] Missing Master PDF or Member Masterlist.", console_slot, "ERROR")
            st.error("Please upload both the Master PDF Document and the Member Masterlist.")
        else:
            with status_container.status("⚙️ Initializing 4-Tier PDF Extraction Engine...", expanded=True) as status:
                try:
                    st.write("📂 Reading member masterlist spreadsheet...")
                    log_split(f"[DATA] Reading {master_list.name}...", console_slot, "PARSING")
                    if master_list.name.endswith('.csv'):
                        df = pd.read_csv(master_list)
                    else:
                        df = pd.read_excel(master_list)
                        
                    name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                    if not name_col:
                        log_split("[ERROR] Could not find a 'Name' column in spreadsheet.", console_slot, "ERROR")
                        status.update(label="❌ Extraction Failed: Missing 'Name' column", state="error", expanded=True)
                    else:
                        raw_names = df.dropna(subset=[name_col])[name_col].astype(str).tolist()
                        contacts = sorted(list(set(n.strip() for n in raw_names if n.strip())), key=len, reverse=True)
                        
                        st.write(f"📄 Loading `{master_pdf.name}` into high-speed memory buffer...")
                        log_split(f"[LOAD] Loading {master_pdf.name} into memory...", console_slot, "LOADING")
                        master_pdf.seek(0)
                        doc = fitz.open(stream=master_pdf.read(), filetype="pdf")
                        total_pages = len(doc)
                        
                        st.write(f"🔍 Scanning {total_pages} pages with 4-Tier Smart Name & Installment Matcher...")
                        log_split(f"[SCAN] Loaded {total_pages} pages. Scanning names & installment headings...", console_slot, "SCANNING")
                        
                        p_bar = progress_slot.progress(0, text=f"Starting extraction (0/{total_pages} pages)...")
                        zip_buffer = io.BytesIO()
                        matched_count = 0
                        unmatched_pages = []
                        name_counts = {}
                        used_filenames = set()
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for i in range(total_pages):
                                pct = int(((i + 1) / total_pages) * 100)
                                status.update(label=f"⏳ Processing Page {i+1} of {total_pages} ({pct}%) — Matched: {matched_count}")
                                p_bar.progress((i + 1) / total_pages, text=f"⏳ Extracting Page {i+1} of {total_pages} ({pct}%)...")
                                
                                page = doc[i]
                                extracted_text = extract_clean_page_text(page)
                                found_name = find_matching_contact(extracted_text, contacts, name_key=None)
                                
                                if found_name:
                                    safe_name = "".join(c for c in found_name if c not in r'\/*?:"<>|')
                                    installment_label = detect_installment(extracted_text)
                                    name_counts[safe_name] = name_counts.get(safe_name, 0) + 1
                                    
                                    if installment_label:
                                        base_fname = f"{safe_name} - {installment_label}"
                                    else:
                                        base_fname = f"{safe_name} - Receipt {name_counts[safe_name]}"
                                    
                                    final_fname = f"{base_fname}.pdf"
                                    counter = 2
                                    while final_fname in used_filenames:
                                        final_fname = f"{base_fname} ({counter}).pdf"
                                        counter += 1
                                    used_filenames.add(final_fname)
                                    
                                    single_doc = fitz.open()
                                    single_doc.insert_pdf(doc, from_page=i, to_page=i)
                                    pdf_bytes = single_doc.tobytes(garbage=4, deflate=True)
                                    single_doc.close()
                                    
                                    zip_file.writestr(final_fname, pdf_bytes)
                                    matched_count += 1
                                    log_split(f"[OK] Page {i+1:03d}/{total_pages} -> {final_fname}", console_slot, f"{pct}%")
                                else:
                                    unmatched_pages.append(i + 1)
                                    snippet = normalize_text(extracted_text)[:80]
                                    log_split(f"[WARN] No match on Page {i+1:03d} [Text: {snippet}...]", console_slot, f"{pct}%")
                                
                        doc.close()
                        p_bar.progress(1.0, text=f"✅ Completed! {matched_count} of {total_pages} pages extracted (100%).")
                        status.update(label=f"✅ Extraction Complete! {matched_count} of {total_pages} PDFs packaged into ZIP.", state="complete", expanded=False)
                        log_split(f"--- EXTRACTION COMPLETE: {matched_count} of {total_pages} PDFs saved (No overwrites) ---", console_slot, "COMPLETE")
                        if unmatched_pages:
                            log_split(f"[INFO] Unmatched Pages ({len(unmatched_pages)}): {unmatched_pages}", console_slot, "COMPLETE")
                        st.session_state.ready_zip = zip_buffer.getvalue()
                        st.session_state.ready_zip_name = zip_name if zip_name.endswith(".zip") else f"{zip_name}.zip"
                except Exception as e:
                    status.update(label=f"❌ Error during extraction: {str(e)}", state="error", expanded=True)
                    log_split(f"[ERROR] {str(e)}", console_slot, "ERROR")

    if "ready_zip" in st.session_state:
        download_slot.download_button(
            label=f"⬇️ Download {st.session_state.ready_zip_name}",
            data=st.session_state.ready_zip,
            file_name=st.session_state.ready_zip_name,
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

# ==========================================
# VIEW 2: BULK MAILER
# ==========================================
elif st.session_state.active_tab == "mail":
    st.markdown("""
        <div class="shadcn-breadcrumb">
            <span>Workspace</span> <span>/</span> <span>MOSES</span> <span>/</span> <span class="crumb-active">Bulk Mailer</span>
            <span class="shadcn-badge" style="margin-left: 8px;">STARTTLS :587</span>
            <span class="shadcn-badge-terracotta">Multi-PDF Bundling</span>
        </div>
        <h2 class="coda-h2">Bulk Mail Merge</h2>
        <div class="coda-subtitle">Automatically group 1st, 2nd, and 3rd installment receipts per member and dispatch consolidated HTML emails.</div>
    """, unsafe_allow_html=True)
    
    with st.container(border=True):
        mc1, mc2 = st.columns([1, 2], gap="large")
        with mc1:
            st.markdown("<div class='field-label'>Sender Email <span class='shadcn-badge-outline'>SMTP</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Official Google Workspace or Gmail address.</div>", unsafe_allow_html=True)
        with mc2:
            sender_email = st.text_input("Sender Email", placeholder="finance.officer@organization.org", label_visibility="collapsed")
            
        st.markdown("<hr class='shadcn-separator'>", unsafe_allow_html=True)
        
        mc3, mc4 = st.columns([1, 2], gap="large")
        with mc3:
            st.markdown("<div class='field-label'>App Password <span class='shadcn-badge-outline'>16-CHAR</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Google Account 2FA App Password for SMTP authentication.</div>", unsafe_allow_html=True)
        with mc4:
            app_password = st.text_input("App Password", type="password", placeholder="•••• •••• •••• ••••", label_visibility="collapsed")

    with st.container(border=True):
        ac1, ac2 = st.columns([1, 1.35], gap="large")
        with ac1:
            st.markdown("<div class='field-label'>Email Banner Image <span class='shadcn-badge-outline'>JPG / PNG</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Header graphic embedded inline at the top of the receipt email.</div>", unsafe_allow_html=True)
            b_slot = st.empty()
        with ac2:
            banner_img = st.file_uploader("banner", type=["jpg", "jpeg", "png"], key="b_img")
        b_slot.markdown(f"<div class='file-pill' style='color: {'#FCD47C' if banner_img else '#8C7E6E'};'>{'● ' + banner_img.name if banner_img else '○ No file selected'}</div>", unsafe_allow_html=True)

        st.markdown("<hr class='shadcn-separator'>", unsafe_allow_html=True)

        ac3, ac4 = st.columns([1, 1.35], gap="large")
        with ac3:
            st.markdown("<div class='field-label'>Email Signature Image <span class='shadcn-badge-outline'>JPG / PNG</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Signatory graphic embedded above the closing block.</div>", unsafe_allow_html=True)
            s_slot = st.empty()
        with ac4:
            sig_img = st.file_uploader("sig", type=["jpg", "jpeg", "png"], key="s_img")
        s_slot.markdown(f"<div class='file-pill' style='color: {'#FCD47C' if sig_img else '#8C7E6E'};'>{'● ' + sig_img.name if sig_img else '○ No file selected'}</div>", unsafe_allow_html=True)

        st.markdown("<hr class='shadcn-separator'>", unsafe_allow_html=True)

        ac5, ac6 = st.columns([1, 1.35], gap="large")
        with ac5:
            st.markdown("<div class='field-label'>Master PDF Document <span class='shadcn-badge-outline'>.PDF</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Consolidated receipt PDF to slice and bundle in-memory.</div>", unsafe_allow_html=True)
            pdf_m_slot = st.empty()
        with ac6:
            master_pdf_mail = st.file_uploader("pdf_mail", type=["pdf"], key="pdf_m")
        pdf_m_slot.markdown(f"<div class='file-pill' style='color: {'#FCD47C' if master_pdf_mail else '#8C7E6E'};'>{'● ' + master_pdf_mail.name if master_pdf_mail else '○ No file selected'}</div>", unsafe_allow_html=True)

        st.markdown("<hr class='shadcn-separator'>", unsafe_allow_html=True)

        ac7, ac8 = st.columns([1, 1.35], gap="large")
        with ac7:
            st.markdown("<div class='field-label'>Member Masterlist <span class='shadcn-badge-outline'>.XLSX / .CSV</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='field-desc'>Roster used to match names to recipient emails.</div>", unsafe_allow_html=True)
            list_m_slot = st.empty()
        with ac8:
            master_list_mail = st.file_uploader("list_mail", type=["xlsx", "csv"], key="list_m")
        list_m_slot.markdown(f"<div class='file-pill' style='color: {'#FCD47C' if master_list_mail else '#8C7E6E'};'>{'● ' + master_list_mail.name if master_list_mail else '○ No file selected'}</div>", unsafe_allow_html=True)

    send_clicked = st.button("Broadcast Emails", type="primary", use_container_width=True)
    mail_status_container = st.empty()
    mail_progress_slot = st.empty()
    mail_console_slot = st.empty()
    render_console(mail_console_slot, st.session_state.mail_logs, "IDLE")

    if send_clicked:
        if not all([sender_email, app_password, banner_img, sig_img, master_pdf_mail, master_list_mail]):
            log_mail("[ERROR] Please fill all credentials and upload all required files.", mail_console_slot, "ERROR")
            st.error("Please fill in all credentials and upload all required files.")
        else:
            with mail_status_container.status("⚙️ Initializing Multi-Attachment Mail Merge...", expanded=True) as status:
                try:
                    st.write("📂 Parsing member masterlist...")
                    log_mail("[INIT] Scanning PDF and grouping multi-installment receipts per member...", mail_console_slot, "GROUPING")
                    if master_list_mail.name.endswith('.csv'):
                        df = pd.read_csv(master_list_mail)
                    else:
                        df = pd.read_excel(master_list_mail)
                        
                    name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                    email_col = next((c for c in df.columns if 'email' in str(c).lower()), None)
                    
                    raw_contacts = df.dropna(subset=[name_col]).to_dict('records')
                    contacts = sorted(raw_contacts, key=lambda x: len(str(x[name_col]).strip()), reverse=True)
                    
                    st.write(f"📄 Loading `{master_pdf_mail.name}` and grouping multi-installment pages...")
                    master_pdf_mail.seek(0)
                    doc = fitz.open(stream=master_pdf_mail.read(), filetype="pdf")
                    total_pages = len(doc)
                    
                    grouped_recipients = {}
                    p_bar = mail_progress_slot.progress(0, text="Phase 1/2: Scanning & grouping PDF pages...")
                    
                    for i in range(total_pages):
                        pct = int(((i + 1) / total_pages) * 100)
                        status.update(label=f"⏳ Phase 1/2: Grouping PDF Pages ({i+1}/{total_pages} — {pct}%)")
                        p_bar.progress((i + 1) / total_pages, text=f"⏳ Phase 1/2: Scanning Page {i+1} of {total_pages} ({pct}%)...")
                        
                        page = doc[i]
                        extracted_text = extract_clean_page_text(page)
                        found_contact = find_matching_contact(extracted_text, contacts, name_key=name_col)
                        
                        if found_contact:
                            recipient_name = str(found_contact[name_col]).strip()
                            recipient_email = str(found_contact[email_col]).strip()
                            safe_name = "".join(c for c in recipient_name if c not in r'\/*?:"<>|')
                            installment_label = detect_installment(extracted_text)
                            
                            single_doc = fitz.open()
                            single_doc.insert_pdf(doc, from_page=i, to_page=i)
                            pdf_bytes = single_doc.tobytes(garbage=4, deflate=True)
                            single_doc.close()
                            
                            if recipient_name not in grouped_recipients:
                                grouped_recipients[recipient_name] = {
                                    "email": recipient_email,
                                    "attachments": [],
                                    "installments": []
                                }
                            
                            idx = len(grouped_recipients[recipient_name]["attachments"]) + 1
                            if installment_label:
                                att_name = f"{safe_name} - {installment_label}.pdf"
                                grouped_recipients[recipient_name]["installments"].append(installment_label)
                            else:
                                att_name = f"{safe_name} - Receipt {idx}.pdf"
                                grouped_recipients[recipient_name]["installments"].append(f"Receipt #{idx}")
                                
                            grouped_recipients[recipient_name]["attachments"].append((att_name, pdf_bytes))
                            log_mail(f"[QUEUE] Page {i+1:03d}: {att_name}", mail_console_slot, f"P1 · {pct}%")
                        else:
                            log_mail(f"[WARN] No matching name on Page {i+1:03d}. Skipping.", mail_console_slot, f"P1 · {pct}%")
                    
                    doc.close()

                    # Phase 2: Connect to SMTP and send 1 consolidated email per member
                    st.write(f"🔐 Authenticating with Gmail SMTP Server (`smtp.gmail.com:587`)...")
                    status.update(label="🔐 Phase 2/2: Authenticating with Gmail SMTP Server...")
                    log_mail(f"[SMTP] Grouped into {len(grouped_recipients)} unique recipients. Connecting...", mail_console_slot, "AUTH")
                    
                    sent_count = 0
                    banner_img.seek(0)
                    sig_img.seek(0)
                    banner_bytes = banner_img.read()
                    sig_bytes = sig_img.read()
                    banner_ext = 'jpeg' if banner_img.name.lower().endswith(('.jpg', '.jpeg')) else 'png'
                    sig_ext = 'jpeg' if sig_img.name.lower().endswith(('.jpg', '.jpeg')) else 'png'

                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.ehlo()
                        smtp.starttls()
                        smtp.ehlo()
                        smtp.login(sender_email.strip(), app_password.strip())
                        st.write("🚀 Authentication successful! Broadcasting emails...")
                        log_mail("[SMTP] Authentication successful. Broadcasting bundled emails...", mail_console_slot, "SENDING")
                        
                        total_recipients = max(1, len(grouped_recipients))
                        for idx, (recipient_name, info) in enumerate(grouped_recipients.items()):
                            pct = int(((idx + 1) / total_recipients) * 100)
                            status.update(label=f"📨 Phase 2/2: Sending to {recipient_name} ({idx+1}/{total_recipients} — {pct}%)")
                            p_bar.progress((idx + 1) / total_recipients, text=f"📨 Phase 2/2: Sending Email {idx+1} of {total_recipients} ({pct}%) — {recipient_name}")
                            
                            recipient_email = info["email"]
                            attachments = info["attachments"]
                            installments_str = ", ".join(info["installments"])
                            
                            if not recipient_email or recipient_email.lower() == 'nan':
                                log_mail(f"[SKIP] {recipient_name} (No Email in sheet)", mail_console_slot, f"P2 · {pct}%")
                                continue
                            
                            msg = EmailMessage()
                            msg['Subject'] = f"Official Receipt: Membership Fee Payment ({installments_str}) - {recipient_name}"
                            msg['From'] = sender_email.strip()
                            msg['To'] = recipient_email
                            
                            banner_cid = make_msgid()
                            sig_cid = make_msgid()
                            
                            html_body = f"""
                            <html><body>
                            <table style="border: 1px solid black; width: 100%; max-width: 650px; border-collapse: collapse;">
                                <tr><td style="padding: 0;"><img src="cid:{banner_cid[1:-1]}" style="width: 100%; display: block;"></td></tr>
                                <tr><td style="padding: 20px; font-family: 'Times New Roman', serif; font-size: 14px; line-height: 1.6;">
                                    Dear {recipient_name},<br><br>
                                    Greetings!<br><br>
                                    I hope this message finds you well.<br><br>
                                    This email serves as the <strong style="color: #9F5556;">official acknowledgment of your membership fee payment ({installments_str})</strong>.<br><br>
                                    Attached to this email {'are your personalized official receipts' if len(attachments) > 1 else 'is your personalized official receipt'} ({len(attachments)} PDF {'files' if len(attachments) > 1 else 'file'}) for your records. Kindly review the attached document(s) and ensure all details are accurate.<br><br>
                                    Should you have any concerns or discrepancies regarding your payment record or the attached receipt(s), please do not hesitate to reply to this email or reach out to the Finance Committee.<br><br>
                                    Thank you for your prompt compliance and continued support.<br><br>
                                    Best regards,<br><br>
                                    <img src="cid:{sig_cid[1:-1]}" style="max-width: 150px; display: block; margin-bottom: 5px;">
                                    Authorized Signatory<br>Finance Committee<br>MOSES Automated System
                                </td></tr>
                            </table>
                            </body></html>
                            """
                            
                            msg.set_content("Please enable HTML to view this email.")
                            msg.add_alternative(html_body, subtype='html')
                            
                            msg.get_payload()[1].add_related(banner_bytes, 'image', banner_ext, cid=banner_cid)
                            msg.get_payload()[1].add_related(sig_bytes, 'image', sig_ext, cid=sig_cid)
                            
                            for att_filename, att_bytes in attachments:
                                msg.add_attachment(att_bytes, maintype='application', subtype='pdf', filename=att_filename)
                            
                            smtp.send_message(msg)
                            sent_count += 1
                            log_mail(f"[SENT] ({sent_count}/{len(grouped_recipients)}): {recipient_name} [{len(attachments)} PDF(s): {installments_str}]", mail_console_slot, f"P2 · {pct}%")
                            
                    p_bar.progress(1.0, text=f"✅ Broadcast Complete! Sent {sent_count} consolidated emails (100%).")
                    status.update(label=f"✅ Batch Complete! Successfully emailed {sent_count} members.", state="complete", expanded=False)
                    log_mail(f"--- BATCH MAIL MERGE COMPLETE ({sent_count} members emailed) ---", mail_console_slot, "COMPLETE")
                except Exception as e:
                    status.update(label=f"❌ Critical Error: {str(e)}", state="error", expanded=True)
                    log_mail(f"[CRITICAL] {str(e)}", mail_console_slot, "ERROR")
