import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import fitz  # PyMuPDF
import io
import zipfile
import openpyxl

# --- Page Configuration ---
st.set_page_config(page_title="MOSES Manager", page_icon="🧾", layout="wide", initial_sidebar_state="expanded")

# --- Initialize Session State ---
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "split"
if "split_logs" not in st.session_state:
    st.session_state.split_logs = "> System initialized. Multi-Installment Scanner & Excel Engine online.\n"
if "mail_logs" not in st.session_state:
    st.session_state.mail_logs = "> System initialized. Multi-Attachment Mail Engine online.\n"

def render_console(placeholder, log_text):
    placeholder.markdown(f"""
        <div class="console-card">
            <div class="console-title">SYSTEM LOGS</div>
            <div class="console-box">{log_text}</div>
        </div>
    """, unsafe_allow_html=True)

def log_split(msg, placeholder=None):
    st.session_state.split_logs += f"> {msg}\n"
    if placeholder:
        render_console(placeholder, st.session_state.split_logs)

def log_mail(msg, placeholder=None):
    st.session_state.mail_logs += f"> {msg}\n"
    if placeholder:
        render_console(placeholder, st.session_state.mail_logs)

def detect_installment(text):
    """Reads the PDF heading to identify which installment receipt is on the page."""
    upper_text = text.upper()
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

# --- Dynamic CSS for Sidebar Tab Highlighting ---
split_bg = "#252526" if st.session_state.active_tab == "split" else "transparent"
split_color = "#FF69B4" if st.session_state.active_tab == "split" else "#888888"

mail_bg = "#252526" if st.session_state.active_tab == "mail" else "transparent"
mail_color = "#FF69B4" if st.session_state.active_tab == "mail" else "#888888"

st.markdown(f"""
    <style>
    .stApp {{
        background-color: #1E1E1E;
        font-family: 'Segoe UI', sans-serif;
    }}
    header[data-testid="stHeader"] {{
        background: transparent !important;
    }}
    .stDeployButton {{
        display: none !important;
    }}
    .block-container {{
        padding-top: 2.5rem !important;
        padding-left: 3rem !important;
        padding-right: 3rem !important;
        max-width: 1400px !important;
        width: 95% !important;
    }}
    [data-testid="stSidebar"] {{
        background-color: #181818 !important;
        border-right: 1px solid #252526;
        min-width: 250px !important;
        max-width: 250px !important;
    }}
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
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button {{
        background-color: {split_bg} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button p {{
        color: {split_color} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(3) .stButton > button:hover {{
        background-color: #252526 !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button {{
        background-color: {mail_bg} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button p {{
        color: {mail_color} !important;
    }}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:nth-of-type(4) .stButton > button:hover {{
        background-color: #252526 !important;
    }}
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: #252526 !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 26px 32px !important;
        margin-bottom: 22px !important;
    }}
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
    .stTextInput input {{
        background-color: #333333 !important;
        border: none !important;
        color: white !important;
        border-radius: 6px !important;
        padding: 10px 14px !important;
    }}
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

        c5, c6 = st.columns([1, 1.4], gap="large")
        with c5:
            st.markdown("<div class='row-title'>Output Destination Archive</div>", unsafe_allow_html=True)
            st.markdown("<div class='row-sub'>Bundled as downloadable .ZIP archive (Auto-tags 1st/2nd/3rd Installment)</div>", unsafe_allow_html=True)
        with c6:
            zip_name = st.text_input("Archive Name", value="MOSES_Extracted_Receipts.zip", label_visibility="collapsed")

    extract_clicked = st.button("Extract & Rename PDFs", type="primary", use_container_width=True)
    progress_slot = st.empty()
    download_slot = st.empty()
    console_slot = st.empty()
    render_console(console_slot, st.session_state.split_logs)

    if extract_clicked:
        if not master_pdf or not master_list:
            log_split("ERROR: Missing Master PDF or Member Masterlist.", console_slot)
        else:
            log_split(f"Loading {master_pdf.name} into memory...", console_slot)
            try:
                if master_list.name.endswith('.csv'):
                    df = pd.read_csv(master_list)
                else:
                    df = pd.read_excel(master_list)
                    
                name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                if not name_col:
                    log_split("DATA ERROR: Could not find a 'Name' column in spreadsheet.", console_slot)
                else:
                    # Sort names longest-first to avoid partial substring collisions
                    raw_names = df.dropna(subset=[name_col])[name_col].astype(str).tolist()
                    contacts = sorted(list(set(n.strip() for n in raw_names if n.strip())), key=len, reverse=True)
                    
                    master_pdf.seek(0)
                    doc = fitz.open(stream=master_pdf.read(), filetype="pdf")
                    total_pages = len(doc)
                    
                    log_split(f"Loaded {total_pages} pages. Scanning names & installment headings...", console_slot)
                    p_bar = progress_slot.progress(0)
                    zip_buffer = io.BytesIO()
                    matched_count = 0
                    name_counts = {}
                    used_filenames = set()
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for i in range(total_pages):
                            page = doc[i]
                            extracted_text = page.get_text("text") or ""
                            found_name = next((n for n in contacts if n.lower() in extracted_text.lower()), None)
                            
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
                                log_split(f"Matched Page {i+1}/{total_pages} -> {final_fname}", console_slot)
                            else:
                                log_split(f"WARNING: No matching name found on Page {i+1}. Skipping.", console_slot)
                            
                            p_bar.progress((i + 1) / total_pages)
                            
                    doc.close()
                    progress_slot.empty()
                    log_split(f"--- EXTRACTION COMPLETE: {matched_count} of {total_pages} PDFs saved (No overwrites) ---", console_slot)
                    st.session_state.ready_zip = zip_buffer.getvalue()
                    st.session_state.ready_zip_name = zip_name if zip_name.endswith(".zip") else f"{zip_name}.zip"
            except Exception as e:
                log_split(f"ERROR: {str(e)}", console_slot)

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
    st.markdown("<h2 style='font-weight: 700; font-size: 28px; margin-bottom: 22px;'>Bulk Mail Merge</h2>", unsafe_allow_html=True)
    
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

    send_clicked = st.button("Broadcast Emails", type="primary", use_container_width=True)
    mail_progress_slot = st.empty()
    mail_console_slot = st.empty()
    render_console(mail_console_slot, st.session_state.mail_logs)

    if send_clicked:
        if not all([sender_email, app_password, banner_img, sig_img, master_pdf_mail, master_list_mail]):
            log_mail("ERROR: Please fill all credentials and upload all required files.", mail_console_slot)
        else:
            log_mail("Scanning PDF and grouping multi-installment receipts per member...", mail_console_slot)
            try:
                if master_list_mail.name.endswith('.csv'):
                    df = pd.read_csv(master_list_mail)
                else:
                    df = pd.read_excel(master_list_mail)
                    
                name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
                email_col = next((c for c in df.columns if 'email' in str(c).lower()), None)
                
                # Sort contacts by name length descending to prevent partial name collisions
                raw_contacts = df.dropna(subset=[name_col]).to_dict('records')
                contacts = sorted(raw_contacts, key=lambda x: len(str(x[name_col]).strip()), reverse=True)
                
                master_pdf_mail.seek(0)
                doc = fitz.open(stream=master_pdf_mail.read(), filetype="pdf")
                total_pages = len(doc)
                
                # Phase 1: Group all matching PDF pages by member so 1 member = 1 email with all their PDFs
                grouped_Simulation = {}
                p_bar = mail_progress_slot.progress(0)
                
                for i in range(total_pages):
                    page = doc[i]
                    extracted_text = page.get_text("text") or ""
                    found_contact = next(
                        (c for c in contacts if str(c[name_col]).strip().lower() in extracted_text.lower()),
                        None
                    )
                    
                    if found_contact:
                        recipient_name = str(found_contact[name_col]).strip()
                        recipient_email = str(found_contact[email_col]).strip()
                        safe_name = "".join(c for c in recipient_name if c not in r'\/*?:"<>|')
                        installment_label = detect_installment(extracted_text)
                        
                        single_doc = fitz.open()
                        single_doc.insert_pdf(doc, from_page=i, to_page=i)
                        pdf_bytes = single_doc.tobytes(garbage=4, deflate=True)
                        single_doc.close()
                        
                        if recipient_name not in grouped_Simulation:
                            grouped_Simulation[recipient_name] = {
                                "email": recipient_email,
                                "attachments": [],
                                "installments": []
                            }
                        
                        idx = len(grouped_Simulation[recipient_name]["attachments"]) + 1
                        if installment_label:
                            att_name = f"{safe_name} - {installment_label}.pdf"
                            grouped_Simulation[recipient_name]["installments"].append(installment_label)
                        else:
                            att_name = f"{safe_name} - Receipt {idx}.pdf"
                            grouped_Simulation[recipient_name]["installments"].append(f"Receipt #{idx}")
                            
                        grouped_Simulation[recipient_name]["attachments"].append((att_name, pdf_bytes))
                        log_mail(f"Queued Page {i+1}: {att_name}", mail_console_slot)
                    else:
                        log_mail(f"WARNING: No matching name on Page {i+1}. Skipping.", mail_console_slot)
                        
                    p_bar.progress((i + 1) / total_pages)
                
                doc.close()

                # Phase 2: Connect to SMTP and send 1 consolidated email per member
                log_mail(f"Grouped into {len(grouped_Simulation)} unique recipients. Connecting to Gmail SMTP...", mail_console_slot)
                sent_count = 0
                banner_bytes = banner_img.read()
                sig_bytes = sig_img.read()
                banner_ext = 'jpeg' if banner_img.name.lower().endswith(('.jpg', '.jpeg')) else 'png'
                sig_ext = 'jpeg' if sig_img.name.lower().endswith(('.jpg', '.jpeg')) else 'png'

                with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                    smtp.ehlo()
                    smtp.starttls()
                    smtp.ehlo()
                    smtp.login(sender_email.strip(), app_password.strip())
                    log_mail("Authentication successful. Broadcasting bundled emails...", mail_console_slot)
                    
                    total_recipients = max(1, len(grouped_Simulation))
                    for idx, (recipient_name, info) in enumerate(grouped_Simulation.items()):
                        recipient_email = info["email"]
                        attachments = info["attachments"]
                        installments_str = ", ".join(info["installments"])
                        
                        if not recipient_email or recipient_email.lower() == 'nan':
                            log_mail(f"SKIPPED: {recipient_name} (No Email in sheet)", mail_console_slot)
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
                                This email serves as the <strong style="color: #FF69B4;">official acknowledgment of your membership fee payment ({installments_str})</strong>.<br><br>
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
                        
                        # Attach all PDFs belonging to this member
                        for att_filename, att_bytes in attachments:
                            msg.add_attachment(att_bytes, maintype='application', subtype='pdf', filename=att_filename)
                        
                        smtp.send_message(msg)
                        sent_count += 1
                        log_mail(f"Sent ({sent_count}/{len(grouped_Simulation)}): {recipient_name} [{len(attachments)} PDF(s): {installments_str}]", mail_console_slot)
                        p_bar.progress((idx + 1) / total_recipients)
                        
                mail_progress_slot.empty()
                log_mail(f"--- BATCH MAIL MERGE COMPLETE ({sent_count} members emailed) ---", mail_console_slot)
            except Exception as e:
                log_mail(f"CRITICAL ERROR: {str(e)}", mail_console_slot)
