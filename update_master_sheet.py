# update_master_sheet.py
import pandas as pd

MASTER_SHEET = 'Master_Sales_Sheet_Cleaned.xlsx'
TRACKING_DB_FILE = 'tracking_database.csv'
OPENS_LOG_FILE = 'opens_log.csv'
OUTPUT_FILE = 'Master_Sheet_with_Tracking.xlsx'

def generate_report():
    print("Generating tracking report...")
    try:
        master_df = pd.read_excel(MASTER_SHEET)
        sent_df = pd.read_csv(TRACKING_DB_FILE)
        opens_df = pd.read_csv(OPENS_LOG_FILE)
    except FileNotFoundError as e:
        print(f"Error: A required file is missing: {e}. Did you run the sender script first?")
        return

    # --- Process Data ---
    # Find the very first open for each tracking ID
    opens_df['opened_time'] = pd.to_datetime(opens_df['opened_time'])
    first_opens = opens_df.sort_values('opened_time').groupby('tracking_id').first()

    # Merge sent log with the first open data
    sent_df = pd.merge(sent_df, first_opens, on='tracking_id', how='left')

    # Now, let's map this data back to the master sheet using VAT_ID
    # We only care about the latest send/open for each company
    tracking_summary = sent_df.sort_values('sent_time').groupby('VAT_ID').last().reset_index()

    # --- Update Master Sheet ---
    # Merge the tracking summary into the main dataframe
    report_df = pd.merge(master_df, tracking_summary[['VAT_ID', 'sent_time', 'opened_time']], on='VAT_ID', how='left')

    # Create clean status columns
    report_df['Send_Status'] = report_df['sent_time'].apply(lambda x: 'Sent' if pd.notna(x) else 'Not Sent')
    report_df['Open_Status'] = report_df['opened_time'].apply(lambda x: 'Opened' if pd.notna(x) else 'Not Opened')
    
    # Rename for clarity and reorder columns
    report_df.rename(columns={'sent_time': 'Last_Sent_Time', 'opened_time': 'First_Open_Time'}, inplace=True)
    
    # Move new columns to the front after VAT_ID for easy viewing
    cols = list(report_df.columns)
    new_col_order = ['VAT_ID', 'Send_Status', 'Open_Status', 'Last_Sent_Time', 'First_Open_Time'] + [c for c in cols if c not in ['VAT_ID', 'Send_Status', 'Open_Status', 'Last_Sent_Time', 'First_Open_Time']]
    report_df = report_df[new_col_order]

    report_df.to_excel(OUTPUT_FILE, index=False)
    
    total_sent = report_df['Send_Status'].value_counts().get('Sent', 0)
    total_opened = report_df['Open_Status'].value_counts().get('Opened', 0)
    open_rate = (total_opened / total_sent) * 100 if total_sent > 0 else 0
    
    print("\n--- Campaign Report ---")
    print(f"Total Companies Emailed: {total_sent}")
    print(f"Unique Companies Opened: {total_opened}")
    print(f"Open Rate:               {open_rate:.2f}%")
    print(f"\n✅ Success! Updated sheet saved as '{OUTPUT_FILE}'")

if __name__ == "__main__":
    generate_report()