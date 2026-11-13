import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
from pypdf import PdfReader, PdfWriter
import io
import zipfile
import openpyxl

# --- Page Configuration ---
st.set_page_config(page_title="MOSES Manager", page_icon="🧾", layout="wide", initial_sidebar_state="expanded")

# --- Initialize Session State ---
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "split"
if "split_logs" not in st.session_state:
    st.session_state.split_logs = "> System initialized. Smart PDF Scanner & Excel Engine online.\n"
if "mail_logs" not in st.session_state:
    st.session_state.mail_logs = "> System initialized. Mail Engine online.\n"

def log_split(msg):
    st.session_state.split_logs += f"> {msg}\n"

def log_mail(msg):
    st.session_state.mail_logs += f"> {msg}\n"

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

# --- Dynamic CSS for Sidebar Tab Highlighting ---
split_bg = "#252526" if st.session_state.active_tab == "split" else "transparent"
split_color = "#FF69B4" if st.session_state.active_tab == "split" else "#888888"

mail_bg = "#252526" if st.session_state.active_tab == "mail" else "transparent"
mail_color = "#FF69B4" if st.session_state.active_tab == "mail" else "#888888"

st.markdown(f"""
    <style>
    /* Global Background & Fonts */
    .stApp {{
        background-color: #1E1E1E;
        font-family: 'Segoe UI', sans-serif;
    }}
    
    /* Hide Top Streamlit Header & Deploy Bar */
    header[data-testid="stHeader"] {{
        background: transparent !important;
    }}
    .stDeployButton {{
        display: none !important;
    }}
    
    /* EXPANDED WORKSPACE CONTAINER */
    .block-container {{
        padding-top: 2.5rem !important;
        padding-left: 3rem !important;
        padding-right: 3rem !important;
        max-width: 1400px !important;
        width: 95% !important;
    }}

    /* Sidebar Styling */
    [data-testid="stSidebar"] {{
        background-color: #181818 !important;
        border-right: 1px solid #252526;
        min-width: 250px !important;
        max-width: 250px !important;
    }}

    /* Sidebar Navigation Buttons (Zero Radio Circles!) */
    [data-testid="stSidebar"] .stButton > button {{
        width: 100%;
        justify-content: flex-start;
        text-align: left;
        padding: 14px 20px !important;
        border-radius: 8px !important;
        border: none !important;
        box-shadow: none !important;
        transition: all 0.2s ease;
    }}
    [data-testid="stSidebar"] .stButton > button p {{
        font-size: 15px !important;
        font-weight: 700 !important;
        margin: 0 !important;
    }}

    /* Tab 1 (Split Receipts) State */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button {{
        background-color: {split_bg} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button p {{
        color: {split_color} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button:hover {{
        background-color: #252526 !important;
    }}

    /* Tab 2 (Bulk Mailer) State */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button {{
        background-color: {mail_bg} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button p {{
        color: {mail_color} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button:hover {{
        background-color: #252526 !important;
    }}

    /* Card Containers - Expanded Padding */
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: #252526 !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 26px 32px !important;
        margin-bottom: 22px !important;
    }}

    /* Primary Action Button (Pink Full-Width Bar) */
    div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {{
        background-color: #FF69B4 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        height: 50px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        transition: background-color 0.2s ease !important;
    }}
    div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover {{
        background-color: #E0529C !important;
    }}

    /* Secondary Buttons (Get Template) */
    div.stDownloadButton > button[kind="secondary"] {{
        background-color: #1E1E1E !important;
        color: #E0E0E0 !important;
        border: 1px solid #3A3A3A !important;
        border-radius: 6px !important;
        padding: 8px 16px !important;
        font-size: 13px !important;
        font-weight: 600 !important;
    }}
    div.stDownloadButton > button[kind="secondary"]:hover {{
        background-color: #333333 !important;
        border-color: #FF69B4 !important;
    }}

    /* EXPANDED FILE UPLOADER DROPZONES */
    .stFileUploader label {{
        display: none !important;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        background-color: #1E1E1E !important;
        border: 1px dashed #4A4A4A !important;
        border-radius: 8px !important;
        padding: 18px 24px !important;
        min-height: 78px !important;
    }}
    [data-testid="stFileUploaderDropzone"]:hover {{
        border-color: #FF69B4 !important;
    }}

    /* Text Inputs */
    .stTextInput input {{
        background-color: #333333 !important;
        border: none !important;
        color: white !important;
        border-radius: 6px !important;
        padding: 10px 14px !important;
    }}

    /* System Console Log Box */
    .console-card {{
        background-color: #252526;
        border-radius: 12px;
        padding: 18px 26px;
        margin-top: 28px;
    }}
    .console-title {{
        font-size: 11px;
        color: #888888;
        font-weight: 800;
        letter-spacing: 0.5px;
        margin-bottom: 10px;
    }}
    .console-box {{
        background-color: transparent;
        color: #A9B7C6;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        height: 160px;
        overflow-y: auto;
        white-space: pre-wrap;
        line-height: 1.5;
    }}
    .row-title {{
        font-weight: 700;
        font-size: 15px;
        color: white;
        margin-bottom: 4px;
    }}
    .row-sub {{
        font-size: 13px;
        color: #888888;
        margin-bottom: 14px;
    }}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.markdown("<div style='color: #FF69B4; font-weight: 900; font-size: 28px; letter-spacing: 0.5px; margin-bottom: 2px;'>MOSES</div>", unsafe_allow_html=True)
    st.markdown("<div style='color: #888888; font-size: 13px; margin-bottom: 30px;'>Automated Manager</div>", unsafe_allow_html=True)
    
    if st.button("Split Receipts", key="nav_split", use_container_width=True):
        st.session_state.active_tab = "split"
        st.rerun()
        
    if st.button("Bulk Mailer", key="nav_mail", use_container_width=True):
        st.session_state.active_tab = "mail"
        st.rerun()

# ==========================================
# VIEW 1: SPLIT RECEIPTS
# ==========================================
if st.session_state.active_tab == "split":
    st.markdown("<h2 style='font-weight: 700; font-size: 28px; margin-bottom: 22px;'>Receipt Splitter</h2>", unsafe_allow_html=True)
    
    with st.container(border=True):
        # Row 1: Master PDF (Widened right column for uploader)
        c1, c2 = st.columns([1, 1.4], gap="large")
        with c1:
            st.markdown("<div class='row-title'>Master PDF Document</div>", unsafe_allow_html=True)
            pdf_status_slot = st.empty()
        with c2:
            master_pdf = st.file_uploader("pdf_split", type=["pdf"], key="pdf_s")
        pdf_status_slot.markdown(
            f"<div class='row-sub' style='color: {'#FFFFFF' if master_pdf else '#888888'};'>{master_pdf.name if master_pdf else 'No file selected'}</div>",
            unsafe_allow_html=True
        )

        st.markdown("<hr style='border-color: #333333; margin: 16px 0;'>", unsafe_allow_html=True)

        # Row 2: Member Masterlist + Get Template
        c3, c4 = st.columns([1, 1.4], gap="large")
        with c3:
            st.markdown("<div class='row-title'>Member Masterlist (Excel or CSV)</div>", unsafe_allow_html=True)
            list_status_slot = st.empty()
            st.download_button(
                label="Get Template",
                data=get_excel_template(),
                file_name="MOSES_MasterList_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary"
            )
        with c4:
            master_list = st.file_uploader("list_split", type=["xlsx", "csv"], key="list_s")
        list_status_slot.markdown(
            f"<div class='row-sub' style='color: {'#FFFFFF' if master_list else '#888888'};'>{master_list.name if master_list else 'No file selected'}</div>",
            unsafe_allow_html=True
        )

        st.markdown("<hr style='border-color: #333333; margin: 16px 0;'>", unsafe_allow_html=True)

        # Row 3: Output Destination
        c5, c6 = st.columns([1, 1.4], gap="large")
        with c5:
            st.markdown("<div class='row-title'>Output Destination Archive</div>", unsafe_allow_html=True)
            st.markdown("<div class='row-sub'>Bundled as downloadable .ZIP archive for macOS/Windows</div>", unsafe_allow_html=True)
        with c6:
            zip_name = st.text_input("Archive Name", value="MOSES_Extracted_Receipts.zip", label_visibility="collapsed")

    if st.button("Extract & Rename PDFs", type="primary", use_container_width=True):
        if not master_pdf or not master_list:
            log_split("ERROR: Missing Master PDF or Member Masterlist.")
        else:
            log_split(f"Scanning pages inside {master_pdf.name}...")
            try:
                if master_list.name.endswith('.csv'):
                    df = pd.read_csv(master_list)
                else:
                    df = pd.read_excel(master_list)
                    
                name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                if not name_col:
                    log_split("DATA ERROR: Could not find a 'Name' column in spreadsheet.")
                else:
                    contacts = df.dropna(subset=[name_col])[name_col].astype(str).tolist()
                    reader = PdfReader(master_pdf)
                    zip_buffer = io.BytesIO()
                    matched_count = 0
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for i, page in enumerate(reader.pages):
                            extracted_text = page.extract_text() or ""
                            found_name = next((n.strip() for n in contacts if n.strip().lower() in extracted_text.lower()), None)
                            
                            if found_name:
                                safe_name = "".join(c for c in found_name if c not in r'\/*?:"<>|')
                                pdf_writer = PdfWriter()
                                pdf_writer.add_page(page)
                                
                                temp_pdf = io.BytesIO()
                                pdf_writer.write(temp_pdf)
                                temp_pdf.seek(0)
                                
                                zip_file.writestr(f"{safe_name}.pdf", temp_pdf.read())
                                matched_count += 1
                                log_split(f"Matched Page {i+1} -> {safe_name}.pdf")
                            else:
                                log_split(f"WARNING: No matching name found on Page {i+1}. Skipping.")
                                
                    log_split(f"--- EXTRACTION COMPLETE: {matched_count} of {len(reader.pages)} PDFs saved ---")
                    st.session_state.ready_zip = zip_buffer.getvalue()
                    st.session_state.ready_zip_name = zip_name if zip_name.endswith(".zip") else f"{zip_name}.zip"
            except Exception as e:
                log_split(f"ERROR: {str(e)}")

    if "ready_zip" in st.session_state:
        st.download_button(
            label=f"⬇️ Download {st.session_state.ready_zip_name}",
            data=st.session_state.ready_zip,
            file_name=st.session_state.ready_zip_name,
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

    # Console Display
    st.markdown(f"""
        <div class="console-card">
            <div class="console-title">SYSTEM LOGS</div>
            <div class="console-box">{st.session_state.split_logs}</div>
        </div>
    """, unsafe_allow_html=True)

# ==========================================
# VIEW 2: BULK MAILER
# ==========================================
elif st.session_state.active_tab == "mail":
    st.markdown("<h2 style='font-weight: 700; font-size: 28px; margin-bottom: 22px;'>Bulk Mail Merge</h2>", unsafe_allow_html=True)
    
    # Credentials Card
    with st.container(border=True):
        mc1, mc2 = st.columns([1, 2.2], gap="large")
        with mc1:
            st.markdown("<div class='row-title' style='padding-top: 8px;'>Sender Email</div>", unsafe_allow_html=True)
        with mc2:
            sender_email = st.text_input("Sender Email", placeholder="finance.officer@organization.org", label_visibility="collapsed")
            
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        
        mc3, mc4 = st.columns([1, 2.2], gap="large")
        with mc3:
            st.markdown("<div class='row-title' style='padding-top: 8px;'>App Password</div>", unsafe_allow_html=True)
        with mc4:
            app_password = st.text_input("App Password", type="password", placeholder="16-character code", label_visibility="collapsed")

    # Assets & Files Card
    with st.container(border=True):
        ac1, ac2 = st.columns([1, 1.4], gap="large")
        with ac1:
            st.markdown("<div class='row-title'>Email Banner Image (JPG/PNG)</div>", unsafe_allow_html=True)
            b_slot = st.empty()
        with ac2:
            banner_img = st.file_uploader("banner", type=["jpg", "jpeg", "png"], key="b_img")
        b_slot.markdown(f"<div class='row-sub' style='color: {'#FFFFFF' if banner_img else '#888888'};'>{banner_img.name if banner_img else 'No file selected'}</div>", unsafe_allow_html=True)

        st.markdown("<hr style='border-color: #333333; margin: 16px 0;'>", unsafe_allow_html=True)

        ac3, ac4 = st.columns([1, 1.4], gap="large")
        with ac3:
            st.markdown("<div class='row-title'>Email Signature Image (JPG/PNG)</div>", unsafe_allow_html=True)
            s_slot = st.empty()
        with ac4:
            sig_img = st.file_uploader("sig", type=["jpg", "jpeg", "png"], key="s_img")
        s_slot.markdown(f"<div class='row-sub' style='color: {'#FFFFFF' if sig_img else '#888888'};'>{sig_img.name if sig_img else 'No file selected'}</div>", unsafe_allow_html=True)

        st.markdown("<hr style='border-color: #333333; margin: 16px 0;'>", unsafe_allow_html=True)

        ac5, ac6 = st.columns([1, 1.4], gap="large")
        with ac5:
            st.markdown("<div class='row-title'>Master PDF Document</div>", unsafe_allow_html=True)
            pdf_m_slot = st.empty()
        with ac6:
            master_pdf_mail = st.file_uploader("pdf_mail", type=["pdf"], key="pdf_m")
        pdf_m_slot.markdown(f"<div class='row-sub' style='color: {'#FFFFFF' if master_pdf_mail else '#888888'};'>{master_pdf_mail.name if master_pdf_mail else 'No file selected'}</div>", unsafe_allow_html=True)

        st.markdown("<hr style='border-color: #333333; margin: 16px 0;'>", unsafe_allow_html=True)

        ac7, ac8 = st.columns([1, 1.4], gap="large")
        with ac7:
            st.markdown("<div class='row-title'>Member Masterlist (Excel or CSV)</div>", unsafe_allow_html=True)
            list_m_slot = st.empty()
        with ac8:
            master_list_mail = st.file_uploader("list_mail", type=["xlsx", "csv"], key="list_m")
        list_m_slot.markdown(f"<div class='row-sub' style='color: {'#FFFFFF' if master_list_mail else '#888888'};'>{master_list_mail.name if master_list_mail else 'No file selected'}</div>", unsafe_allow_html=True)

    if st.button("Broadcast Emails", type="primary", use_container_width=True):
        if not all([sender_email, app_password, banner_img, sig_img, master_pdf_mail, master_list_mail]):
            log_mail("ERROR: Please fill all credentials and upload all required files.")
        else:
            log_mail("Connecting to Gmail Server via Port 587 (STARTTLS)...")
            try:
                if master_list_mail.name.endswith('.csv'):
                    df = pd.read_csv(master_list_mail)
                else:
                    df = pd.read_excel(master_list_mail)
                    
                name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                email_col = next((c for c in df.columns if 'email' in str(c).lower()), None)
                contacts = df.dropna(subset=[name_col]).to_dict('records')
                
                reader = PdfReader(master_pdf_mail)
                sent_count = 0

                with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                    smtp.ehlo()
                    smtp.starttls()
                    smtp.ehlo()
                    smtp.login(sender_email.strip(), app_password.strip())
                    log_mail("Authentication successful. Starting batch send...")
                    
                    for i, page in enumerate(reader.pages):
                        extracted_text = page.extract_text() or ""
                        found_contact = next(
                            (c for c in contacts if str(c[name_col]).strip().lower() in extracted_text.lower()),
                            None
                        )
                        
                        if found_contact:
                            recipient_name = str(found_contact[name_col]).strip()
                            recipient_email = str(found_contact[email_col]).strip()
                            
                            if not recipient_email or recipient_email.lower() == 'nan':
                                log_mail(f"SKIPPED: {recipient_name} (No Email in sheet)")
                                continue
                            
                            writer = PdfWriter()
                            writer.add_page(page)
                            pdf_buffer = io.BytesIO()
                            writer.write(pdf_buffer)
                            pdf_buffer.seek(0)
                            
                            msg = EmailMessage()
                            msg['Subject'] = f"Official Receipt: First Installment Payment - {recipient_name}"
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
                                    This email serves as the <strong style="color: #FF69B4;">official acknowledgment of your first installment payment</strong>.<br><br>
                                    Attached to this email is your personalized official receipt for your records. Kindly review the attached PDF document and ensure all details are accurate.<br><br>
                                    Should you have any concerns or discrepancies regarding your payment record or the attached receipt, please do not hesitate to reply to this email or reach out to the Finance Committee.<br><br>
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
                            
                            banner_ext = 'jpeg' if banner_img.name.lower().endswith(('.jpg', '.jpeg')) else 'png'
                            sig_ext = 'jpeg' if sig_img.name.lower().endswith(('.jpg', '.jpeg')) else 'png'
                            banner_img.seek(0)
                            sig_img.seek(0)
                            
                            msg.get_payload()[1].add_related(banner_img.read(), 'image', banner_ext, cid=banner_cid)
                            msg.get_payload()[1].add_related(sig_img.read(), 'image', sig_ext, cid=sig_cid)
                            msg.add_attachment(pdf_buffer.read(), maintype='application', subtype='pdf', filename=f"{recipient_name}.pdf")
                            
                            smtp.send_message(msg)
                            sent_count += 1
                            log_mail(f"Sent: {recipient_name}")
                            
                log_mail(f"--- BATCH MAIL MERGE COMPLETE ({sent_count} sent) ---")
            except Exception as e:
                log_mail(f"CRITICAL ERROR: {str(e)}")

    # Console Display
    st.markdown(f"""
        <div class="console-card">
            <div class="console-title">SYSTEM LOGS</div>
            <div class="console-box">{st.session_state.mail_logs}</div>
        </div>
    """, unsafe_allow_html=True)
