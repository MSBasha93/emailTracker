# send_tracked_emails.py
import smtplib, ssl, time, uuid, pandas as pd, os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- CONFIGURATION ---
SMTP_SERVER = "mail.otax.tech"
SMTP_PORT = 465
SENDER_EMAIL = "info@otax.tech"

# !!! PASTE YOUR NGROK FORWARDING URL HERE !!!
# It must be the full URL, e.g., "https://1a2b-3c4d-5e6f.ngrok-free.app"
TRACKER_URL = "https://YOUR_NGROK_URL_HERE.ngrok-free.app" 

MASTER_SHEET = 'Master_Sales_Sheet_Cleaned.xlsx'
TRACKING_DB_FILE = 'tracking_database.csv' # This logs which ID was sent to whom

# --- SCRIPT LOGIC ---
if TRACKER_URL == "https://YOUR_NGROK_URL_HERE.ngrok-free.app":
    print("FATAL ERROR: You must edit the script and paste your ngrok URL into the TRACKER_URL variable.")
    exit()

# Ask for password at runtime for security
SENDER_PASSWORD = input(f"Please enter the password for {SENDER_EMAIL}: ")

def create_tracked_html(body_text, tracking_id):
    tracking_pixel_html = f'<img src="{TRACKER_URL}/track/{tracking_id}" width="1" height="1" alt="">'
    return f"<html><body><p>{body_text.replace(os.linesep, '<br>')}</p>{tracking_pixel_html}</body></html>"

def main():
    try:
        df = pd.read_excel(MASTER_SHEET)
        recipients = df[df['Contact_Email_0'].notna()].to_dict('records')
    except FileNotFoundError:
        print(f"Error: Master sheet '{MASTER_SHEET}' not found.")
        return

    if not os.path.exists(TRACKING_DB_FILE):
        with open(TRACKING_DB_FILE, 'w') as f:
            f.write("tracking_id,VAT_ID,recipient_email,sent_time\n")

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            print("✅ Successfully connected to SMTP server.")

            email_count_this_hour = 0
            start_time = time.time()

            for person in recipients:
                if time.time() - start_time > 3600:
                    start_time = time.time()
                    email_count_this_hour = 0
                
                if email_count_this_hour >= 140:
                    print("Hourly limit reached. Pausing for one hour...")
                    time.sleep(3600)
                    start_time = time.time()
                    email_count_this_hour = 0

                vat_id = person['VAT_ID']
                recipient_email = person['Contact_Email_0']
                company_name = person['Company_Name_0']
                
                message = MIMEMultipart("alternative")
                message["Subject"] = f"A proposition for {company_name}"
                message["From"] = SENDER_EMAIL
                message["To"] = recipient_email
                tracking_id = str(uuid.uuid4())

                body_text = f"Dear {company_name} team,\n\nWe are writing to you today with a special proposition regarding your business operations.\n\nBest regards,\nThe OTAX Team"
                html_body = create_tracked_html(body_text, tracking_id)

                message.attach(MIMEText(body_text, "plain"))
                message.attach(MIMEText(html_body, "html"))

                try:
                    server.sendmail(SENDER_EMAIL, recipient_email, message.as_string())
                    print(f"✉️ Email sent to {recipient_email} (VAT: {vat_id})")
                    
                    with open(TRACKING_DB_FILE, 'a') as f:
                        f.write(f'"{tracking_id}","{vat_id}","{recipient_email}","{time.strftime("%Y-%m-%d %H:%M:%S")}"\n')

                    email_count_this_hour += 1
                    time.sleep(2) 
                except Exception as e:
                    print(f"🔥 Failed to send to {recipient_email}: {e}")

    except smtplib.SMTPAuthenticationError:
        print("🔥 SMTP Login Failed. Wrong password?")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()