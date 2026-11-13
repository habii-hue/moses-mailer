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
    /* Force-hide the radio button circles completely */
    [data-testid="stSidebar"] label[data-baseweb="radio"] > div:first-child {
        display: none !important;
    }
    
    /* Style the radio button container (the tab) */
    [data-testid="stSidebar"] label[data-baseweb="radio"] {
        padding: 12px 20px !important;
        border-radius: 8px !important;
        margin-bottom: 5px !important;
        cursor: pointer !important;
        transition: background-color 0.2s ease !important;
        width: 100% !important;
    }
    
    /* Hover state for inactive tabs */
    [data-testid="stSidebar"] label[data-baseweb="radio"]:hover {
        background-color: #1E1E1E !important;
    }
    
    /* Active tab styling using the modern :has() selector */
    [data-testid="stSidebar"] label[data-baseweb="radio"]:has(input:checked) {
        background-color: #252526 !important;
    }
    
    /* Make the text match the Tkinter font weight */
    [data-testid="stSidebar"] label[data-baseweb="radio"] div {
        font-size: 15px !important;
        font-weight: 700 !important;
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
                        for name in contacts:
                            if name.strip().lower() in extracted_text.lower():
                                found_name = name.strip()
                                break
                        
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
                
                log_split(f"--- EXTRACTION COMPLETE: {matched_count} of {len(reader.pages)} PDFs zipped ---")
                
                st.success("Extraction Complete!")
                st.download_button(
                    label="Download ZIP Archive",
                    data=zip_buffer.getvalue(),
                    file_name="MOSES_Extracted_Receipts.zip",
                    mime="application/zip",
                    type="primary"
                )

    # Console Display
    st.markdown('<p style="font-size: 11px; color: gray; font-weight: bold; margin-top: 30px; margin-bottom: 5px;">SYSTEM LOGS</p>', unsafe_allow_html=True)
    st.markdown(f'<div class="console-box">{st.session_state.split_logs}</div>', unsafe_allow_html=True)

# ==========================================
# TAB 2: BULK MAILER
# ==========================================
elif view == "Bulk Mailer":
    st.markdown("<h2 style='margin-bottom: 20px;'>Bulk Mail Merge</h2>", unsafe_allow_html=True)
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        with st.container(border=True):
            st.markdown("**Sender Email**")
            sender_email = st.text_input("email", placeholder="erick.regala@example.com", label_visibility="collapsed")
            
            st.markdown("<br>**App Password**", unsafe_allow_html=True)
            app_password = st.text_input("pass", type="password", placeholder="16-character code", label_visibility="collapsed")

    with col_right:
        with st.container(border=True):
            st.markdown("**Email Banner (JPG/PNG)**")
            banner_img = st.file_uploader("banner", type=["jpg", "jpeg", "png"], key="b_img")
            
            st.markdown("<br>**Email Signature (JPG/PNG)**", unsafe_allow_html=True)
            sig_img = st.file_uploader("sig", type=["jpg", "jpeg", "png"], key="s_img")

    with st.container(border=True):
        st.markdown("**Master PDF Document**")
        master_pdf = st.file_uploader("pdf_mail", type=["pdf"], key="pdf_m")
        
        st.markdown("<br>**Member Masterlist (Excel or CSV)**", unsafe_allow_html=True)
        master_list = st.file_uploader("list_mail", type=["xlsx", "csv"], key="list_m")

    action_col, _ = st.columns([1, 1])
    
    with action_col:
        if st.button("Broadcast Emails", type="primary", use_container_width=True):
            if not all([sender_email, app_password, banner_img, sig_img, master_pdf, master_list]):
                st.error("Please fill in all credentials and upload all required files.")
            else:
                log_mail("Parsing Masterlist...")
                
                if master_list.name.endswith('.csv'):
                    df = pd.read_csv(master_list)
                else:
                    df = pd.read_excel(master_list)
                    
                name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                email_col = next((c for c in df.columns if 'email' in str(c).lower()), None)
                contacts = df.dropna(subset=[name_col]).to_dict('records')
                
                log_mail("Scanning PDF and connecting to SMTP Server...")
                reader = PdfReader(master_pdf)
                sent_count = 0

                try:
                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.ehlo()
                        smtp.starttls()
                        smtp.ehlo()
                        smtp.login(sender_email, app_password)
                        
                        log_mail("Authentication successful! Broadcasting...")
                        progress_bar = st.progress(0)

                        for i in range(len(reader.pages)):
                            page = reader.pages[i]
                            extracted_text = page.extract_text() or ""
                            
                            found_contact = None
                            for contact in contacts:
                                name_val = str(contact[name_col]).strip()
                                if name_val.lower() in extracted_text.lower():
                                    found_contact = contact
                                    break
                            
                            if found_contact:
                                recipient_name = str(found_contact[name_col]).strip()
                                recipient_email = str(found_contact[email_col]).strip()
                                
                                if not recipient_email or str(recipient_email).lower() == 'nan':
                                    continue
                                
                                writer = PdfWriter()
                                writer.add_page(page)
                                pdf_buffer = io.BytesIO()
                                writer.write(pdf_buffer)
                                pdf_buffer.seek(0)
                                
                                msg = EmailMessage()
                                msg['Subject'] = f"Official Receipt: First Installment Payment - {recipient_name}"
                                msg['From'] = sender_email
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
                                        This email serves as the <strong style="color: #FF69B4;">official acknowledgment of your first installment payment</strong>.<br><br>
                                        Attached to this email is your personalized official receipt for your records. Kindly review the attached PDF document and ensure all details are accurate.<br><br>
                                        Should you have any concerns regarding your payment record, please do not hesitate to reply to this email.<br><br>
                                        Thank you for your prompt compliance.<br><br>
                                        Best regards,<br><br>
                                        <img src="cid:{sig_cid[1:-1]}" style="max-width: 150px; display: block; margin-bottom: 5px;">
                                        Erick James Regala<br>President<br>MOSES Automated System
                                    </td></tr>
                                </table>
                                </body></html>
                                """
                                
                                msg.set_content("Please enable HTML to view this email.")
                                msg.add_alternative(html_body, subtype='html')
                                
                                banner_ext = 'jpeg' if banner_img.name.lower().endswith('.jpg') else 'png'
                                sig_ext = 'jpeg' if sig_img.name.lower().endswith('.jpg') else 'png'
                                banner_img.seek(0)
                                sig_img.seek(0)
                                
                                msg.get_payload()[1].add_related(banner_img.read(), 'image', banner_ext, cid=banner_cid)
                                msg.get_payload()[1].add_related(sig_img.read(), 'image', sig_ext, cid=sig_cid)
                                msg.add_attachment(pdf_buffer.read(), maintype='application', subtype='pdf', filename=f"{recipient_name}.pdf")
                                
                                smtp.send_message(msg)
                                sent_count += 1
                                log_mail(f"Sent: {recipient_name}")
                            
                            progress_bar.progress((i + 1) / len(reader.pages))
                            
                    log_mail(f"--- BATCH COMPLETE: Successfully sent {sent_count} emails ---")
                    
                except Exception as e:
                    log_mail(f"CRITICAL ERROR: {str(e)}")

    # Console Display
    st.markdown('<p style="font-size: 11px; color: gray; font-weight: bold; margin-top: 30px; margin-bottom: 5px;">SYSTEM LOGS</p>', unsafe_allow_html=True)
    st.markdown(f'<div class="console-box">{st.session_state.mail_logs}</div>', unsafe_allow_html=True)
