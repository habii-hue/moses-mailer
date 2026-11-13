import os
import csv
import smtplib
import threading
from email.message import EmailMessage
from email.utils import make_msgid
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter

# --- SKLauncher-Style UI Theme Settings ---
ctk.set_appearance_mode("Dark")
ACCENT_COLOR = "#8338EC" 
HOVER_COLOR = "#5D24B3"
CARD_BG = "#252526"
MAIN_BG = "#1E1E1E"

# ==========================================
# CUSTOM UI: ANIMATION ENGINE
# ==========================================
def hex_to_rgb(hex_code):
    hex_code = hex_code.lstrip("#")
    return tuple(int(hex_code[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(rgb):
    return f"#{int(rgb[0]):02x}{int(rgb[1]):02x}{int(rgb[2]):02x}"

def interpolate_color(c1, c2, factor):
    r1, g1, b1 = hex_to_rgb(c1)
    r2, g2, b2 = hex_to_rgb(c2)
    return rgb_to_hex((
        r1 + (r2 - r1) * factor,
        g1 + (g2 - g1) * factor,
        b1 + (b2 - b1) * factor
    ))

class SmoothButton(ctk.CTkButton):
    def __init__(self, master, base_color, hover_color, press_color=None, **kwargs):
        super().__init__(master, fg_color=base_color, hover_color=base_color, **kwargs)
        self.base_color = base_color
        self.hover_color = hover_color
        self.press_color = press_color or self._darken(hover_color)
        self.current_color = base_color
        self._target_color = base_color
        self._animating = False

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Button-1>", self.on_press)
        self.bind("<ButtonRelease-1>", self.on_release)

    def _darken(self, hex_color):
        rgb = hex_to_rgb(hex_color)
        return rgb_to_hex([max(0, c - 45) for c in rgb])

    def animate(self):
        if self.current_color == self._target_color:
            self._animating = False
            return

        self.current_color = interpolate_color(self.current_color, self._target_color, 0.3)
        self.configure(fg_color=self.current_color)

        if self._color_distance(self.current_color, self._target_color) < 5:
            self.current_color = self._target_color
            self.configure(fg_color=self.current_color)
            self._animating = False
            return

        self.after(16, self.animate)

    def _color_distance(self, c1, c2):
        r1, g1, b1 = hex_to_rgb(c1)
        r2, g2, b2 = hex_to_rgb(c2)
        return abs(r1-r2) + abs(g1-g2) + abs(b1-b2)

    def set_target(self, color):
        self._target_color = color
        if not self._animating:
            self._animating = True
            self.animate()

    def on_enter(self, e):
        if self.cget("state") != "disabled": self.set_target(self.hover_color)

    def on_leave(self, e):
        if self.cget("state") != "disabled": self.set_target(self.base_color)

    def on_press(self, e):
        if self.cget("state") != "disabled":
            self.current_color = self.press_color
            self._target_color = self.press_color
            self.configure(fg_color=self.current_color)

    def on_release(self, e):
        if self.cget("state") != "disabled":
            x, y = self.winfo_pointerxy()
            bx, by = self.winfo_rootx(), self.winfo_rooty()
            w, h = self.winfo_width(), self.winfo_height()
            if bx <= x <= bx + w and by <= y <= by + h:
                self.set_target(self.hover_color)
            else:
                self.set_target(self.base_color)


# ==========================================
# MAIN APPLICATION INTERFACE
# ==========================================
class JPIAManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("JPIA-SSCR Manager")
        self.geometry("950x650")
        self.resizable(False, False)
        self.configure(fg_color=MAIN_BG)

        # File path variables
        self.master_pdf_path = ""
        self.data_path = ""
        self.output_folder = ""
        self.banner_path = ""
        self.signature_path = ""

        # --- Grid Layout ---
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.setup_sidebar()
        
        # --- Main Content Area ---
        self.main_container = ctk.CTkFrame(self, fg_color=MAIN_BG, corner_radius=0)
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_container.grid_rowconfigure(1, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        self.view_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.view_frame.grid(row=0, column=0, sticky="nsew")
        self.view_frame.grid_columnconfigure(0, weight=1)

        self.setup_split_view()
        self.setup_mail_view()

        # --- Shared Console ---
        self.console_card = ctk.CTkFrame(self.main_container, fg_color=CARD_BG, corner_radius=10)
        self.console_card.grid(row=1, column=0, sticky="sew", pady=(20, 0))
        self.console_card.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self.console_card, text="SYSTEM LOGS", font=("Segoe UI", 11, "bold"), text_color="gray").grid(row=0, column=0, sticky="w", padx=15, pady=(10, 0))
        self.console = ctk.CTkTextbox(self.console_card, height=120, fg_color="transparent", font=("Consolas", 12), text_color="#A9B7C6", state="disabled")
        self.console.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.select_view("split")
        self.log_message("System initialized. Smart PDF Scanner & Excel Engine online.")

    def log_message(self, message):
        self.console.configure(state="normal")
        self.console.insert("end", "> " + message + "\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    # ==========================================
    # SIDEBAR NAVIGATION
    # ==========================================
    def setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#181818")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(4, weight=1)

        self.lbl_logo = ctk.CTkLabel(self.sidebar, text="JPIA-SSCR", font=("Segoe UI Black", 24), text_color=ACCENT_COLOR)
        self.lbl_logo.grid(row=0, column=0, padx=20, pady=(30, 5), sticky="w")
        self.lbl_subtitle = ctk.CTkLabel(self.sidebar, text="Automated Manager", font=("Segoe UI", 12), text_color="gray")
        self.lbl_subtitle.grid(row=1, column=0, padx=20, pady=(0, 30), sticky="w")

        self.btn_nav_split = SmoothButton(self.sidebar, base_color="#181818", hover_color=CARD_BG, text="   Split Receipts", 
                                          font=("Segoe UI", 14, "bold"), text_color="white", anchor="w", height=45, 
                                          command=lambda: self.select_view("split"))
        self.btn_nav_split.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

        self.btn_nav_mail = SmoothButton(self.sidebar, base_color="#181818", hover_color=CARD_BG, text="   Bulk Mailer", 
                                         font=("Segoe UI", 14, "bold"), text_color="white", anchor="w", height=45, 
                                         command=lambda: self.select_view("mail"))
        self.btn_nav_mail.grid(row=3, column=0, sticky="ew", padx=10, pady=5)

    def select_view(self, view_name):
        self.btn_nav_split.base_color = "#181818"
        self.btn_nav_split.set_target("#181818")
        self.btn_nav_split.configure(text_color="gray")

        self.btn_nav_mail.base_color = "#181818"
        self.btn_nav_mail.set_target("#181818")
        self.btn_nav_mail.configure(text_color="gray")

        self.frame_split.grid_remove()
        self.frame_mail.grid_remove()

        if view_name == "split":
            self.frame_split.grid(row=0, column=0, sticky="nsew")
            self.btn_nav_split.base_color = CARD_BG
            self.btn_nav_split.set_target(CARD_BG)
            self.btn_nav_split.configure(text_color=ACCENT_COLOR)
        elif view_name == "mail":
            self.frame_mail.grid(row=0, column=0, sticky="nsew")
            self.btn_nav_mail.base_color = CARD_BG
            self.btn_nav_mail.set_target(CARD_BG)
            self.btn_nav_mail.configure(text_color=ACCENT_COLOR)

    # ==========================================
    # FILE PICKER & TEMPLATE GENERATION
    # ==========================================
    def create_file_row(self, parent, label_text, attr_name, row_idx, filetypes, is_dir=False, add_template_btn=False):
        lbl_title = ctk.CTkLabel(parent, text=label_text, font=("Segoe UI", 13, "bold"))
        lbl_title.grid(row=row_idx, column=0, padx=20, pady=(15, 5), sticky="w")
        
        lbl_val = ctk.CTkLabel(parent, text="No file selected", font=("Segoe UI", 12), text_color="gray")
        lbl_val.grid(row=row_idx+1, column=0, padx=20, pady=(0, 15), sticky="w")
        
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=row_idx, column=1, rowspan=2, padx=20, sticky="e")
        
        if add_template_btn:
            btn_tpl = SmoothButton(btn_frame, base_color="#1E1E1E", hover_color="#2A2A2A", text="Get Template", width=100, command=self.generate_template)
            btn_tpl.pack(side="left", padx=(0, 10))
            
        btn = SmoothButton(btn_frame, base_color="#333333", hover_color="#444444", text="Browse", width=90,
                           command=lambda: self.browse_action(attr_name, lbl_val, filetypes, is_dir))
        btn.pack(side="left")
        
        return lbl_val

    def browse_action(self, attr_name, label_widget, filetypes, is_dir):
        if is_dir:
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=filetypes)
        
        if path:
            setattr(self, attr_name, path)
            label_widget.configure(text=os.path.basename(path) or path, text_color="white")

    def generate_template(self):
        try:
            import openpyxl
        except ImportError:
            messagebox.showerror("Error", "Please install openpyxl via terminal:\npip install openpyxl")
            return
            
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="JPIA_MasterList_Template.xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if save_path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "JPIA Members"
            ws.append(["Membership Type", "Name", "Section", "Amount Paid", "Email"]) 
            ws.append(["Normal", "Buenaventura, Harvey", "BSCS 3-X", 200, "s.harvey@sscr.edu"])
            ws.append(["Reads", "Arriesgado, Elaiza Anne", "BSA 3-X", 150, "s.elaizaanne.arriesgado@sscr.edu"])
            wb.save(save_path)
            self.log_message(f"Excel Template generated: {os.path.basename(save_path)}")

    # ==========================================
    # DATA PARSING ENGINE (Excel & CSV)
    # ==========================================
    def parse_data(self):
        contacts = []
        try:
            if self.data_path.endswith('.csv'):
                with open(self.data_path, mode='r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    headers = reader.fieldnames
                    name_col = next((h for h in headers if 'name' in str(h).lower()), None)
                    email_col = next((h for h in headers if 'email' in str(h).lower()), None)
                    
                    if not name_col: raise ValueError("Could not find a 'Name' column in the file.")
                    
                    for row in reader:
                        name = row.get(name_col, "").strip()
                        email = row.get(email_col, "").strip() if email_col else ""
                        if name: contacts.append({"Name": name, "Email": email})
                        
            elif self.data_path.endswith('.xlsx'):
                import openpyxl
                wb = openpyxl.load_workbook(self.data_path, data_only=True)
                ws = wb.active
                rows = list(ws.values)
                if not rows: raise ValueError("Excel sheet is empty.")
                
                headers = [str(h).strip() if h else "" for h in rows[0]]
                name_idx = next((i for i, h in enumerate(headers) if 'name' in h.lower()), -1)
                email_idx = next((i for i, h in enumerate(headers) if 'email' in h.lower()), -1)
                
                if name_idx == -1: raise ValueError("Could not find a 'Name' column in row 1 of Excel.")
                
                for row in rows[1:]:
                    name = str(row[name_idx]).strip() if len(row) > name_idx and row[name_idx] else ""
                    email = str(row[email_idx]).strip() if email_idx != -1 and len(row) > email_idx and row[email_idx] else ""
                    if name and name.lower() != 'none':
                        contacts.append({"Name": name, "Email": email if email.lower() != 'none' else ""})
            else:
                raise ValueError("Unsupported file format. Please use .xlsx or .csv")
                
        except Exception as e:
            self.log_message(f"DATA ERROR: {str(e)}")
            return None
            
        return contacts

    # ==========================================
    # VIEW 1: SPLIT PDFs
    # ==========================================
    def setup_split_view(self):
        self.frame_split = ctk.CTkFrame(self.view_frame, fg_color="transparent")
        self.frame_split.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.frame_split, text="Receipt Splitter", font=("Segoe UI", 24, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 20))

        card1 = ctk.CTkFrame(self.frame_split, fg_color=CARD_BG, corner_radius=10)
        card1.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        card1.grid_columnconfigure(0, weight=1)

        self.lbl_pdf = self.create_file_row(card1, "Master PDF Document", "master_pdf_path", 0, [("PDF Files", "*.pdf")])
        self.lbl_data_split = self.create_file_row(card1, "Member Masterlist (Excel or CSV)", "data_path", 2, [("Data Files", "*.xlsx;*.csv")], add_template_btn=True)
        self.lbl_folder = self.create_file_row(card1, "Output Destination Folder", "output_folder", 4, None, is_dir=True)

        self.btn_split = SmoothButton(self.frame_split, base_color=ACCENT_COLOR, hover_color=HOVER_COLOR, 
                                      text="Extract & Rename PDFs", font=("Segoe UI", 14, "bold"), height=45, corner_radius=8,
                                      command=self.start_split_thread)
        self.btn_split.grid(row=2, column=0, sticky="ew")

    # ==========================================
    # VIEW 2: BULK MAILER
    # ==========================================
    def setup_mail_view(self):
        self.frame_mail = ctk.CTkFrame(self.view_frame, fg_color="transparent")
        self.frame_mail.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.frame_mail, text="Bulk Mail Merge", font=("Segoe UI", 24, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 20))

        card_cred = ctk.CTkFrame(self.frame_mail, fg_color=CARD_BG, corner_radius=10)
        card_cred.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        card_cred.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card_cred, text="Sender Email", font=("Segoe UI", 13, "bold")).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        self.entry_email = ctk.CTkEntry(card_cred, placeholder_text="s.elaizaanne.arriesgado@sscr.edu", border_width=0, fg_color="#333333")
        self.entry_email.grid(row=0, column=1, padx=20, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(card_cred, text="App Password", font=("Segoe UI", 13, "bold")).grid(row=1, column=0, padx=20, pady=(0, 20), sticky="w")
        self.entry_password = ctk.CTkEntry(card_cred, placeholder_text="16-character code", show="•", border_width=0, fg_color="#333333")
        self.entry_password.grid(row=1, column=1, padx=20, pady=(0, 20), sticky="ew")

        card_assets = ctk.CTkFrame(self.frame_mail, fg_color=CARD_BG, corner_radius=10)
        card_assets.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        card_assets.grid_columnconfigure(0, weight=1)

        self.lbl_banner = self.create_file_row(card_assets, "Email Banner Image (JPG)", "banner_path", 0, [("Images", "*.jpg;*.jpeg;*.png")])
        self.lbl_sig = self.create_file_row(card_assets, "Email Signature Image", "signature_path", 2, [("Images", "*.jpg;*.jpeg;*.png")])

        self.btn_send = SmoothButton(self.frame_mail, base_color=ACCENT_COLOR, hover_color=HOVER_COLOR, 
                                     text="Broadcast Emails", font=("Segoe UI", 14, "bold"), height=45, corner_radius=8,
                                     command=self.start_email_thread)
        self.btn_send.grid(row=3, column=0, sticky="ew")


    # ==========================================
    # THREADING & CORE LOGIC
    # ==========================================
    def start_split_thread(self):
        if not all([self.master_pdf_path, self.data_path, self.output_folder]):
            messagebox.showwarning("Missing Inputs", "Please select the Master PDF, Excel/CSV file, and Output Folder.")
            return
        self.btn_split.configure(state="disabled", text="Processing...")
        threading.Thread(target=self.process_split, daemon=True).start()

    def start_email_thread(self):
        if not all([self.entry_email.get(), self.entry_password.get(), self.data_path, self.output_folder, self.banner_path, self.signature_path]):
            messagebox.showwarning("Missing Inputs", "Please fill all credentials and ensure the Excel/CSV and Asset folders are configured.")
            return
        self.btn_send.configure(state="disabled", text="Authenticating...")
        threading.Thread(target=self.process_emails, daemon=True).start()

    def process_split(self):
        contacts = self.parse_data()
        if not contacts:
            self.btn_split.configure(state="normal", text="Extract & Rename PDFs")
            self.btn_split.set_target(ACCENT_COLOR)
            return

        names_list = [c['Name'] for c in contacts]
        self.log_message(f"Scanning pages inside {os.path.basename(self.master_pdf_path)}...")

        try:
            reader = PdfReader(self.master_pdf_path)
            total_pages = len(reader.pages)
            
            matched_count = 0
            unmatched_pages = []

            for i in range(total_pages):
                page = reader.pages[i]
                extracted_text = page.extract_text()
                if not extracted_text:
                    extracted_text = ""
                
                found_name = None
                for name in names_list:
                    if name.strip().lower() in extracted_text.lower():
                        found_name = name.strip()
                        break
                
                if found_name:
                    safe_name = "".join(c for c in found_name if c not in r'\/*?:"<>|')
                    out_path = os.path.join(self.output_folder, f"{safe_name}.pdf")
                    
                    writer = PdfWriter()
                    writer.add_page(page)
                    
                    with open(out_path, "wb") as output_pdf:
                        writer.write(output_pdf)
                    
                    self.log_message(f"Matched Page {i+1} -> {safe_name}.pdf")
                    matched_count += 1
                else:
                    self.log_message(f"WARNING: No matching name found on Page {i+1}. Skipping.")
                    unmatched_pages.append(i+1)
                    
            self.log_message(f"\n--- EXTRACTION COMPLETE: {matched_count} of {total_pages} PDFs saved ---")
            if unmatched_pages:
                self.log_message(f"Could not identify names on pages: {unmatched_pages}")
                
        except Exception as e:
            self.log_message(f"ERROR: {str(e)}")
            
        self.btn_split.configure(state="normal", text="Extract & Rename PDFs")
        self.btn_split.set_target(ACCENT_COLOR)

    def process_emails(self):
        contacts = self.parse_data()
        if not contacts:
            self.btn_send.configure(state="normal", text="Broadcast Emails")
            self.btn_send.set_target(ACCENT_COLOR)
            return

        my_email = self.entry_email.get().strip()
        app_pass = self.entry_password.get().strip()
        
        self.log_message("Connecting to Gmail Server via Port 587 (STARTTLS)...")
        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(my_email, app_pass)
                self.btn_send.configure(text="Broadcasting...")
                self.log_message("Authentication successful. Starting batch send...")
                
                for row in contacts:
                    recipient_name = row['Name']
                    recipient_email = row['Email']
                    pdf_filename = os.path.join(self.output_folder, f"{recipient_name}.pdf")
                    
                    if not os.path.exists(pdf_filename):
                        self.log_message(f"SKIPPED: {recipient_name} (No PDF found)")
                        continue
                    if not recipient_email:
                        self.log_message(f"SKIPPED: {recipient_name} (No Email in sheet)")
                        continue

                    msg = EmailMessage()
                    msg['Subject'] = f"Official Receipt: JPIA-SSCR First Installment Payment - {recipient_name}"
                    msg['From'] = my_email
                    msg['To'] = recipient_email
                    
                    banner_cid = make_msgid()
                    sig_cid = make_msgid()
                    
                    html_body = f"""
                    <html><body>
                    <table style="border: 1px solid black; width: 100%; max-width: 650px; border-collapse: collapse;">
                        <tr><td style="padding: 0;"><img src="cid:{banner_cid[1:-1]}" style="width: 100%; display: block;"></td></tr>
                        <tr><td style="padding: 20px; font-family: 'Times New Roman', serif; font-size: 14px; line-height: 1.6;">
                            Dear {recipient_name},<br><br>
                            Greetings in the name of St. Sebastian!<br><br>
                            I hope this message finds you well.<br><br>
                            This email serves as the <strong style="color: #b0008b;">official acknowledgment of your first installment payment</strong> for the Junior Philippine Institute of Accountants - SSCR (JPIA-SSCR).<br><br>
                            <div style="margin-left: 40px;">
                                Please be reminded of the installment breakdown based on your membership classification:<br>
                                - Normal Members: PHP 200.00<br>
                                - Reads: PHP 150.00
                            </div><br>
                            Attached to this email is your personalized official receipt for your records. Kindly review the attached PDF document and ensure all details are accurate.<br><br>
                            Should you have any concerns or discrepancies regarding your payment record or the attached receipt, please do not hesitate to reply to this email or reach out to the Finance Committee.<br><br>
                            Thank you for your prompt compliance and continued support to JPIA-SSCR.<br><br>
                            Best regards,<br><br>
                            <img src="cid:{sig_cid[1:-1]}" style="max-width: 150px; display: block; margin-bottom: 5px;">
                            Elaiza Anne Arriesgado<br>Vice President for Finance<br>Junior Philippine Institute of Accountants - SSCR
                        </td></tr>
                    </table>
                    </body></html>"""
                    
                    msg.set_content("Please enable HTML to view this email.")
                    msg.add_alternative(html_body, subtype='html')
                    
                    with open(self.banner_path, 'rb') as img:
                        msg.get_payload()[1].add_related(img.read(), 'image', 'jpeg', cid=banner_cid)
                    with open(self.signature_path, 'rb') as img:
                        ext = 'jpeg' if self.signature_path.lower().endswith('.jpg') else 'png'
                        msg.get_payload()[1].add_related(img.read(), 'image', ext, cid=sig_cid)
                    with open(pdf_filename, 'rb') as pdf_file:
                        msg.add_attachment(pdf_file.read(), maintype='application', subtype='pdf', filename=f"{recipient_name}.pdf")
                    
                    smtp.send_message(msg)
                    self.log_message(f"Sent: {recipient_name}")

            self.log_message("--- BATCH MAIL MERGE COMPLETE ---")
        except Exception as e:
            self.log_message(f"CRITICAL ERROR: {str(e)}")
            
        self.btn_send.configure(state="normal", text="Broadcast Emails")
        self.btn_send.set_target(ACCENT_COLOR)

if __name__ == "__main__":
    app = JPIAManagerApp()
    app.mainloop()