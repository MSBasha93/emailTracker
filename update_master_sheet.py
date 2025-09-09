# update_master_sheet_v2.py
import pandas as pd
import os

# ==============================================================================
# --- PRIMARY CONFIGURATION ---
# ==============================================================================
# I noticed your error log showed 'test.csv'. I have changed the default
# CSV_FILENAME to match that.
DATA_SOURCE_TYPE = 'CSV'  # <-- CHANGE THIS to 'EXCEL' for the real campaign report

# --- Data Source Details ---
CSV_FILENAME = 'test.csv' # <-- Adjusted to match your error log
EXCEL_FILENAME = 'Master_Sales_Sheet_Cleaned.xlsx'
EXCEL_SHEET_NAME = 'Email_and_Phone'

# --- Log File and Output Details ---
TRACKING_DB_FILE = 'tracking_database.csv'
OPENS_LOG_FILE = 'opens_log.csv'
# ==============================================================================

def main():
    print("--- Campaign Reporting v2 (Fixed) ---")
    
    # 1. LOAD THE ORIGINAL SOURCE DATA
    source_df = None
    output_filename = ""
    try:
        if DATA_SOURCE_TYPE == 'CSV':
            print(f"Mode: Reporting on CSV test campaign from '{CSV_FILENAME}'...")
            source_df = pd.read_csv(CSV_FILENAME)
            output_filename = CSV_FILENAME.replace('.csv', '_with_tracking.csv')
        elif DATA_SOURCE_TYPE == 'EXCEL':
            print(f"Mode: Reporting on Production campaign from '{EXCEL_FILENAME}' (Sheet: '{EXCEL_SHEET_NAME}')...")
            source_df = pd.read_excel(EXCEL_FILENAME, sheet_name=EXCEL_SHEET_NAME)
            output_filename = EXCEL_FILENAME.replace('.xlsx', '_with_Tracking.xlsx')
        else:
            print(f"FATAL ERROR: Invalid DATA_SOURCE_TYPE: '{DATA_SOURCE_TYPE}'.")
            return
    except FileNotFoundError:
        print(f"FATAL ERROR: The source file '{'CSV_FILENAME' if DATA_SOURCE_TYPE == 'CSV' else EXCEL_FILENAME}' was not found.")
        return
    except Exception as e:
        print(f"FATAL ERROR: Could not read the data source. Reason: {e}")
        return

    # 2. LOAD THE TRACKING LOGS
    try:
        sent_df = pd.read_csv(TRACKING_DB_FILE)
        opens_df = pd.read_csv(OPENS_LOG_FILE)
    except FileNotFoundError as e:
        print(f"Warning: A log file is missing: {e}. Report may be incomplete.")
        if 'tracking_database' in str(e): sent_df = pd.DataFrame(columns=['tracking_id', 'VAT_ID'])
        if 'opens_log' in str(e): opens_df = pd.DataFrame(columns=['opened_time', 'tracking_id'])
    
    # 3. PROCESS THE LOGS TO GET LATEST STATUS
    # --- FIX STARTS HERE ---
    # The original script failed if opens_df was empty. This new logic handles that.

    if not opens_df.empty and 'opened_time' in opens_df.columns:
        # Process opens normally if the log has data
        opens_df['opened_time'] = pd.to_datetime(opens_df['opened_time'])
        first_opens = opens_df.sort_values('opened_time').groupby('tracking_id').first()
        # Merge the open data into the sent data. 'how=left' ensures all sent emails are kept.
        sent_df = pd.merge(sent_df, first_opens, on='tracking_id', how='left')
    else:
        # If no opens have been logged, create an empty 'opened_time' column.
        # This prevents the KeyError.
        sent_df['opened_time'] = pd.NaT # Use NaT for a null datetime value
        sent_df['user_agent'] = None

    # --- FIX ENDS HERE ---

    # Now, find the latest send/open status for each unique company (VAT_ID)
    if not sent_df.empty:
        tracking_summary = sent_df.sort_values('sent_time', ascending=False).groupby('VAT_ID').first().reset_index()
    else:
        tracking_summary = pd.DataFrame(columns=['VAT_ID', 'sent_time', 'opened_time'])

    # 4. MERGE TRACKING DATA BACK INTO THE SOURCE DATAFRAME
    report_df = pd.merge(source_df, tracking_summary[['VAT_ID', 'sent_time', 'opened_time']], on='VAT_ID', how='left')

    # Create clean, human-readable status columns
    report_df['Send_Status'] = report_df['sent_time'].apply(lambda x: 'Sent' if pd.notna(x) else 'Not Sent')
    report_df['Open_Status'] = report_df['opened_time'].apply(lambda x: 'Opened' if pd.notna(x) else 'Not Opened')
    
    report_df.rename(columns={'sent_time': 'Last_Sent_Time', 'opened_time': 'First_Open_Time'}, inplace=True)
    
    status_cols = ['Send_Status', 'Open_Status', 'Last_Sent_Time', 'First_Open_Time']
    original_cols = [col for col in source_df.columns if col not in ['VAT_ID']]
    new_col_order = ['VAT_ID'] + status_cols + original_cols
    report_df = report_df[new_col_order]

    # 5. SAVE THE FINAL REPORT
    if DATA_SOURCE_TYPE == 'CSV':
        report_df.to_csv(output_filename, index=False)
    else: # EXCEL
        report_df.to_excel(output_filename, index=False, engine='openpyxl')

    # Print a summary
    total_in_list = len(report_df)
    total_sent = report_df['Send_Status'].value_counts().get('Sent', 0)
    total_opened = report_df['Open_Status'].value_counts().get('Opened', 0)
    open_rate = (total_opened / total_sent) * 100 if total_sent > 0 else 0
    
    print("\n--- Campaign Report ---")
    print(f"Total Contacts in Source: {total_in_list}")
    print(f"Total Contacts Emailed:   {total_sent}")
    print(f"Unique Opens:             {total_opened}")
    print(f"Open Rate:                {open_rate:.2f}%")
    print(f"\n✅ Success! Updated report saved as '{output_filename}'")

if __name__ == "__main__":
    main()