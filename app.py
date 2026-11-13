import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
from pypdf import PdfReader, PdfWriter
import io

# --- Page Configuration ---
st.set_page_config(page_title="MOSES", page_icon="🧾", layout="centered")

st.title("MOSES")
st.markdown("### Automated Receipt Splitter & Bulk Mailer")

with st.expander("📝 How it works"):
    st.write("Upload your master PDF, your Excel masterlist, and your assets. MOSES will natively scan the PDF for names, match them to the emails in the spreadsheet, and blast them out in one click.")

st.divider()

# --- UI: File Uploads & Credentials ---
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 1. Credentials")
    sender_email = st.text_input("Sender Email", placeholder="erick.regala@example.com")
    app_password = st.text_input("App Password", type="password", placeholder="16-character code")

with col2:
    st.markdown("#### 2. Image Assets")
    banner_img = st.file_uploader("Email Banner", type=["jpg", "jpeg", "png"])
    sig_img = st.file_uploader("Email Signature", type=["jpg", "jpeg", "png"])

st.markdown("#### 3. Documents")
master_pdf = st.file_uploader("Master PDF Document", type=["pdf"])
master_list = st.file_uploader("Member Masterlist (Excel or CSV)", type=["xlsx", "csv"])

# --- Core Execution Logic ---
if st.button("Broadcast Emails", type="primary", use_container_width=True):
    if not all([sender_email, app_password, banner_img, sig_img, master_pdf, master_list]):
        st.error("Please fill in all credentials and upload all required files before broadcasting.")
    else:
        status_box = st.empty()
        progress_bar = st.progress(0)
        
        try:
            status_box.info("Parsing Masterlist...")
            
            # 1. Parse Excel / CSV
            if master_list.name.endswith('.csv'):
                df = pd.read_csv(master_list)
            else:
                df = pd.read_excel(master_list)
                
            # Smart column hunt
            name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
            email_col = next((c for c in df.columns if 'email' in str(c).lower()), None)

            if not name_col or not email_col:
                st.error("Could not find a 'Name' and 'Email' column in the uploaded spreadsheet.")
                st.stop()

            # Clean data
            contacts = df.dropna(subset=[name_col]).to_dict('records')
            
            # 2. Parse PDF and Connect to SMTP
            status_box.info("Scanning PDF and connecting to SMTP Server...")
            reader = PdfReader(master_pdf)
            total_pages = len(reader.pages)
            
            sent_count = 0

            with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(sender_email, app_password)
                
                status_box.success("Authentication successful! Broadcasting...")

                for i in range(total_pages):
                    page = reader.pages[i]
                    extracted_text = page.extract_text() or ""
                    
                    # Scan text for matching name
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
                        
                        # Extract single PDF page directly into RAM (BytesIO)
                        writer = PdfWriter()
                        writer.add_page(page)
                        pdf_buffer = io.BytesIO()
                        writer.write(pdf_buffer)
                        pdf_buffer.seek(0)
                        
                        # Construct Email
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
                        
                        # Handle Image Extensions Dynamically
                        banner_ext = 'jpeg' if banner_img.name.lower().endswith('.jpg') else 'png'
                        sig_ext = 'jpeg' if sig_img.name.lower().endswith('.jpg') else 'png'
                        
                        # Reset memory pointers before reading
                        banner_img.seek(0)
                        sig_img.seek(0)
                        
                        msg.get_payload()[1].add_related(banner_img.read(), 'image', banner_ext, cid=banner_cid)
                        msg.get_payload()[1].add_related(sig_img.read(), 'image', sig_ext, cid=sig_cid)
                        
                        # Attach the in-memory PDF
                        msg.add_attachment(pdf_buffer.read(), maintype='application', subtype='pdf', filename=f"{recipient_name}.pdf")
                        
                        smtp.send_message(msg)
                        sent_count += 1
                    
                    # Update Progress
                    progress_bar.progress((i + 1) / total_pages)
                    
            status_box.success(f"✅ Batch complete! Successfully sent {sent_count} emails.")
            
        except Exception as e:
            status_box.error(f"Critical Error: {str(e)}")
