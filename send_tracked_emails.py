# send_tracked_emails_v3.py (with Message-ID header)
import smtplib, ssl, time, uuid, pandas as pd, os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import make_msgid # <-- Import this new utility

# ==============================================================================
# --- PRIMARY CONFIGURATION ---
# ==============================================================================
DATA_SOURCE_TYPE = 'CSV'
CSV_FILENAME = 'test.csv'
EXCEL_FILENAME = 'Master_Sales_Sheet_Cleaned.xlsx'
EXCEL_SHEET_NAME = 'Email_and_Phone'
SMTP_SERVER = "mail.otax.tech"
SMTP_PORT = 465
SENDER_EMAIL = "info@otax.tech"
TRACKING_DB_FILE = 'tracking_database.csv'
TRACKER_URL = " https://80c9c9088ed3.ngrok-free.app" 
# ==============================================================================

if TRACKER_URL == "https://YOUR_NGROK_URL_HERE.ngrok-free.app":
    print("FATAL ERROR: You must edit the script and paste your ngrok URL into the TRACKER_URL variable.")
    exit()

def create_tracked_html(body_text, tracking_id):
    # (This function does not need to change)
    tracking_pixel_html = f'<img src="{TRACKER_URL}/track/{tracking_id}" width="1" height="1" alt="">'
    html = f"""
    <html><head><style>body {{ font-family: sans-serif; }} p {{ line-height: 1.6; }}</style></head>
    <body><p>{body_text.replace(os.linesep, '<br>')}</p>{tracking_pixel_html}</body></html>"""
    return html

def main():
    # --- (All the data loading and filtering logic remains the same) ---
    print("--- Campaign Manager v3 ---")
    master_df = None
    try:
        if DATA_SOURCE_TYPE == 'CSV':
            print(f"Mode: Testing. Reading from CSV file: '{CSV_FILENAME}'...")
            master_df = pd.read_csv(CSV_FILENAME)
        elif DATA_SOURCE_TYPE == 'EXCEL':
            print(f"Mode: Production. Reading from Excel file: '{EXCEL_FILENAME}', Sheet: '{EXCEL_SHEET_NAME}'...")
            master_df = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_NAME)
        else:
            print(f"FATAL ERROR: Invalid DATA_SOURCE_TYPE: '{DATA_SOURCE_TYPE}'. Choose 'CSV' or 'EXCEL'.")
            return
        print(f"Successfully loaded {len(master_df)} total contacts from source.")
    except FileNotFoundError:
        print(f"FATAL ERROR: The source file was not found. Please check the filename.")
        return
    except Exception as e:
        print(f"FATAL ERROR: Could not read the data source. Reason: {e}")
        return

    if not os.path.exists(TRACKING_DB_FILE):
        with open(TRACKING_DB_FILE, 'w') as f: f.write("tracking_id,VAT_ID,recipient_email,sent_time\n")
    sent_df = pd.read_csv(TRACKING_DB_FILE)
    sent_vat_ids = set(sent_df['VAT_ID'])
    print(f"Found {len(sent_vat_ids)} companies that have already been contacted.")
    contacts_to_send_df = master_df[master_df['Contact_Email_0'].notna() & ~master_df['VAT_ID'].isin(sent_vat_ids)].copy()
    total_new_contacts = len(contacts_to_send_df)
    if total_new_contacts == 0:
        print("\n✅ All contacts from the selected source have been emailed. Nothing to do!")
        return
    print(f"There are {total_new_contacts} new contacts to email from this source.")

    while True:
        try:
            batch_size_str = input(f"How many emails would you like to send in this batch? (Max: {total_new_contacts}): ")
            batch_size = int(batch_size_str)
            if 0 < batch_size <= total_new_contacts: break
            else: print(f"Please enter a number between 1 and {total_new_contacts}.")
        except ValueError: print("Invalid input. Please enter a number.")
    batch_df = contacts_to_send_df.head(batch_size)
    recipients = batch_df.to_dict('records')
    SENDER_PASSWORD = input(f"Please enter the password for {SENDER_EMAIL}: ")
    context = ssl.create_default_context()
    
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            print("✅ Successfully connected to SMTP server. Starting batch...")
            email_count_this_hour, start_time = 0, time.time()
            for i, person in enumerate(recipients):
                # (Rate limiting logic is the same)
                if time.time() - start_time > 3600:
                    start_time, email_count_this_hour = time.time(), 0
                if email_count_this_hour >= 140:
                    print("\nHourly limit reached. Pausing for one hour..."); time.sleep(3600)
                    start_time, email_count_this_hour = time.time(), 0

                vat_id, recipient_email = person['VAT_ID'], person['Contact_Email_0']
                company_name = person.get('Company_Name_0', 'Valued Partner')
                progress_prefix = f"[{i+1}/{batch_size}]"
                
                message = MIMEMultipart("alternative")
                message["Subject"] = f"A proposition for {company_name}"
                message["From"] = f"OTAX Team <{SENDER_EMAIL}>"
                message["To"] = recipient_email
                # --- NEW CODE ---
                # Add a unique Message-ID header to look more like a standard email
                message["Message-ID"] = make_msgid()
                # --- END NEW CODE ---
                
                tracking_id = str(uuid.uuid4())
                body_text = f"Dear {company_name} team,\n\nWe are writing to you today with a special proposition regarding your business operations.\n\nBest regards,\nThe OTAX Team"
                html_body = create_tracked_html(body_text, tracking_id)
                message.attach(MIMEText(body_text, "plain"))
                message.attach(MIMEText(html_body, "html"))

                try:
                    server.sendmail(SENDER_EMAIL, recipient_email, message.as_string())
                    print(f"{progress_prefix} ✉️  Email sent to {recipient_email} (VAT: {vat_id})")
                    with open(TRACKING_DB_FILE, 'a', newline='') as f:
                        f.write(f'"{tracking_id}",{vat_id},"{recipient_email}","{time.strftime("%Y-%m-%d %H:%M:%S")}"\n')
                    email_count_this_hour += 1
                    time.sleep(2)
                except Exception as e:
                    print(f"{progress_prefix} 🔥 Failed to send to {recipient_email}: {e}")
            print(f"\n✅ Batch complete. Sent {len(recipients)} emails.")
    except smtplib.SMTPAuthenticationError:
        print("🔥 SMTP Login Failed. Please check your email and password.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()