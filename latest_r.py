import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time
import os
import win32com.client  # For creating draft emails in Outlook

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'
SHEETS_TO_PROCESS = ['Sheet1', 'Sheet2']  # or ['all'] for all sheets
LOGSOURCE_COLUMN = 'log source name'
IP_COLUMN = 'IP'
IN_QRADAR_COLUMN = 'In Qradar?'  # Must contain "Yes" to be processed
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL = False
DRAFT_OUTPUT_PATH = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), 'inactive_and_errors.xlsx')
ACTIVITY_THRESHOLD_DAYS = 7  # Consider log source inactive if no events in X days
REQUEST_TIMEOUT = 30
# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

# Valid timestamp range
MIN_TIMESTAMP = 0
MAX_TIMESTAMP = 2147483647


def test_qradar_connection(qradar_host, username, password):
    """Test QRadar connection and validate credentials"""
    print("🔗 Testing QRadar connection...")
    qradar_host = qradar_host.rstrip('/')
    endpoint = f"{qradar_host}/api/help/versions"
    
    try:
        resp = requests.get(
            endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )
        if resp.status_code == 200:
            print("✅ QRadar connection successful!")
            return True
        elif resp.status_code == 401:
            print("❌ Authentication failed! Check username/password.")
            return False
        else:
            print(f"⚠️ Unexpected response: {resp.status_code}")
            return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


def _empty_details():
    """Return empty details structure"""
    return {
        'qradar_id': 'N/A',
        'enabled': 'Unknown',
        'last_seen': 'N/A',
        'activity_status': 'Not Found',
        'days_since_last_event': None
    }


def safe_timestamp_conversion(timestamp_ms):
    """Safely convert timestamp to datetime string"""
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    
    try:
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)
        
        # Convert ms to seconds if needed
        if timestamp_ms > 4102444800:
            timestamp_seconds = timestamp_ms / 1000.0
        else:
            timestamp_seconds = timestamp_ms
        
        if timestamp_seconds <= MIN_TIMESTAMP or timestamp_seconds > MAX_TIMESTAMP:
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None
        
        last_event_datetime = datetime.fromtimestamp(timestamp_seconds)
        last_seen = last_event_datetime.strftime('%Y-%m-%d %H:%M:%S')
        
        time_diff = datetime.now() - last_event_datetime
        days_since_last_event = time_diff.days
        
        threshold_time = datetime.now() - timedelta(days=ACTIVITY_THRESHOLD_DAYS)
        
        if last_event_datetime > threshold_time:
            activity_status = 'Active'
        else:
            activity_status = 'Inactive'
            
        return last_seen, activity_status, days_since_last_event
        
    except Exception as e:
        # print(f"   ⚠️ Error parsing timestamp {timestamp_ms}: {e}")
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Get log source details directly from the API.
    Returns details dict or None if not found/error.
    """
    
    clean_identifier = str(identifier).replace('"', '').replace("'", "").strip()
    
    # Construct Filter
    if is_ip:
        query_filter = f'protocol_parameters contains value="{clean_identifier}"'
    else:
        query_filter = f'name ilike "%{clean_identifier}%"'
    
    ls_endpoint = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"

    try:
        resp = requests.get(
            ls_endpoint,
            params={'filter': query_filter},
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )
        
        if resp.status_code != 200:
            return {'status': f'API Error {resp.status_code}', **_empty_details()}

        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', **_empty_details()}

        # Select Best Match
        found_source = None
        
        if is_ip:
            # Strict validation for IPs
            for source in ls_data:
                params = source.get('protocol_parameters', [])
                if any(p.get('value') == clean_identifier for p in params):
                    found_source = source
                    break
            if not found_source: 
                found_source = ls_data[0]
        else:
            found_source = ls_data[0]

        ls_id = found_source.get('id')
        ls_name = found_source.get('name', identifier)
        
        # Get last event details
        last_event_time_ms = found_source.get('last_event_time')
        last_seen, activity_status, days_since_last_event = safe_timestamp_conversion(last_event_time_ms)
        enabled = found_source.get('enabled', False)
        enabled_str = 'Yes' if enabled else 'No'
            
        return {
            'status': 'Found',
            'qradar_id': str(ls_id) if ls_id is not None else '',
            'actual_name': ls_name,
            'enabled': enabled_str,
            'last_seen': last_seen,
            'activity_status': activity_status,
            'days_since_last_event': days_since_last_event
        }

    except Exception as e:
        return {'status': f'Error: {str(e)[:50]}...', **_empty_details()}


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column, in_qradar_col):
    """
    Optimized Sheet Processing:
    1. Filters DataFrame for 'Yes' rows first.
    2. Iterates ONLY relevant rows.
    3. Provides detailed console output.
    """
    print(f"\n{'='*60}")
    print(f"📋 Processing Sheet: {sheet_name}")
    print(f"{'='*60}")
    
    # Initialize columns if missing
    cols_to_init = {
        'status': 'object',
        'qradar_id': 'object', 
        'enabled': 'object',
        'last_seen': 'object',
        'activity_status': 'object',
        'days_since_last_event': 'float64',
        'remarks': 'object'
    }
    
    for col, dtype in cols_to_init.items():
        if col not in df.columns:
            df[col] = pd.Series(dtype=dtype)
    
    # ─── STEP 1: EFFICIENT FILTERING ───
    if in_qradar_col not in df.columns:
        print(f"⚠️ Column '{in_qradar_col}' not found! Skipping entire sheet.")
        return df

    # Create a mask for rows where In Qradar contains "yes" (case-insensitive)
    # Handles NaN safely
    process_mask = df[in_qradar_col].astype(str).str.lower().str.contains("yes", na=False)
    
    rows_to_process = df[process_mask]
    total_rows = len(df)
    target_count = len(rows_to_process)
    
    print(f"📊 Total Rows: {total_rows}")
    print(f"🎯 Rows marked 'Yes': {target_count} (Rows to scan)")
    
    # Mark skipped rows efficiently
    df.loc[~process_mask, 'remarks'] = "Skipped (In Qradar != Yes)"
    
    if target_count == 0:
        print("✅ No rows to process in this sheet.")
        return df

    # ─── STEP 2: ITERATE ONLY TARGET ROWS ───
    current_idx = 0
    
    for idx, row in rows_to_process.iterrows():
        current_idx += 1
        name_val = str(row[logsource_column]).strip()
        ip_val = str(row[ip_column]).strip()
        
        # Clean cleanup
        if name_val.lower() in ['nan', 'none', '', 'null']: name_val = None
        if ip_val.lower() in ['nan', 'none', '', 'null']: ip_val = None
        
        print(f"\n🔹 [{current_idx}/{target_count}] Processing Row {idx+1}")
        
        details = None
        search_method = "None"
        
        # 1. Search by Name
        if name_val:
            print(f"   🔍 Searching Name: '{name_val}' ... ", end="")
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)
            
            if details['status'] == 'Found':
                print("✅ Found!")
                search_method = "Name"
            else:
                print("⚠️ Not Found.")
        
        # 2. Fallback to IP (Only if Name failed/missing)
        if (not details or details['status'] == 'Not Found') and ip_val:
            print(f"   🔁 Fallback to IP: '{ip_val}' ... ", end="")
            details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)
            
            if details['status'] == 'Found':
                print("✅ Found!")
                search_method = "IP"
            else:
                print("❌ Not Found.")

        # Initialize if completely failed
        if not details:
            details = {'status': 'Empty/Invalid', **_empty_details()}
            
        # ─── STEP 3: UPDATE DATAFRAME & DISPLAY DETAILS ───
        
        if details['status'] == 'Found':
            # Update Data
            df.at[idx, 'status'] = 'Found'
            df.at[idx, 'remarks'] = f"Found by {search_method}"
            df.at[idx, 'qradar_id'] = details['qradar_id']
            df.at[idx, 'enabled'] = details['enabled']
            df.at[idx, 'last_seen'] = details['last_seen']
            df.at[idx, 'activity_status'] = details['activity_status']
            df.at[idx, 'days_since_last_event'] = details['days_since_last_event']
            
            # Print Details
            print(f"      📌 Log Source: {details['actual_name']}")
            print(f"      📊 Activity:   {details['activity_status']}")
            print(f"      📅 Last Event: {details['last_seen']} ({details['days_since_last_event']} days ago)")
            
        else:
            # Handle Not Found Logic
            # Smart Inference for AP/FW
            status_val = "Not Found"
            remark_val = "❌ Not found by Name or IP"
            act_val = "Not Found"
            
            if name_val and "AP" in name_val:
                status_val = "Inferred"
                remark_val = "Under WLC (Inferred)"
                act_val = "Inferred"
                print("      ℹ️  Result: Inferred as WLC (AP in name)")
            elif name_val and "FW" in name_val:
                status_val = "Inferred"
                remark_val = "Under Forti (Inferred)"
                act_val = "Inferred"
                print("      ℹ️  Result: Inferred as FortiGate (FW in name)")
            else:
                print("      ❌ Result: ABSENT in QRadar.")

            df.at[idx, 'status'] = status_val
            df.at[idx, 'remarks'] = remark_val
            df.at[idx, 'activity_status'] = act_val
            
            # Clear other fields to be clean
            df.at[idx, 'last_seen'] = "N/A"
            df.at[idx, 'days_since_last_event'] = None

        time.sleep(0.2) # Slight delay to be nice to API
    
    return df


def create_outlook_draft(attachment_path, subject, body):
    """Create Outlook draft"""
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(attachment_path)
        mail.Display()
        print(f"\n✉️  Email draft created successfully.")
    except Exception as e:
        print(f"\n❌ Failed to create Outlook draft: {e}")


def filter_and_email(df_dict, draft_path):
    """Filter results and generate report"""
    frames = []
    
    stats = {
        'total_scanned': 0,
        'found_active': 0,
        'found_inactive': 0,
        'inferred': 0,
        'not_found': 0,
        'errors': 0
    }

    for name, df in df_dict.items():
        if 'status' not in df.columns: continue
        
        # Only count rows we actually touched (status is filled)
        processed_df = df[df['status'].notna()]
        stats['total_scanned'] += len(processed_df)
        
        # Calculate Counts
        stats['found_active'] += len(processed_df[(processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Active')])
        
        inactive_mask = (processed_df['status'] == 'Found') & ((processed_df['activity_status'] == 'Inactive') | (processed_df['activity_status'] == 'No Activity'))
        stats['found_inactive'] += len(processed_df[inactive_mask])
        
        stats['inferred'] += len(processed_df[processed_df['status'] == 'Inferred'])
        stats['not_found'] += len(processed_df[processed_df['status'] == 'Not Found'])
        stats['errors'] += processed_df['status'].str.startswith('API Error', na=False).sum()

        # Build Report DataFrame (Inactive + Not Found + Errors)
        mask_report = inactive_mask | (processed_df['status'] == 'Not Found') | (processed_df['status'] == 'Inferred') | processed_df['status'].str.startswith('API Error', na=False)
        
        if mask_report.any():
            sub = processed_df[mask_report].copy()
            # Update remark for inactive specifically
            sub.loc[inactive_mask, 'remarks'] = f"Inactive > {ACTIVITY_THRESHOLD_DAYS} days"
            sub['sheet_name'] = name
            frames.append(sub)

    # Save and Email
    if not frames:
        print("✅ No issues found; skipping email.")
    else:
        try:
            result_df = pd.concat(frames, ignore_index=True)
            result_df.to_excel(draft_path, index=False)
            print(f"\n💾 Filtered report saved: {draft_path}")

            total_found = stats['found_active'] + stats['found_inactive'] + stats['inferred']
            
            body = f"""Hello,

Attached is the QRadar log source status report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.

📊 EXECUTIVE SUMMARY
----------------------------------------
Total Assets Scanned:  {stats['total_scanned']} (Rows marked 'Yes')
Total Found / Inferred: {total_found}
Total Not Found:       {stats['not_found']}
API Errors:            {stats['errors']}

📉 HEALTH BREAKDOWN
----------------------------------------
✅ Active:              {stats['found_active']}
⚠️ Inactive:            {stats['found_inactive']}  (No events in {ACTIVITY_THRESHOLD_DAYS} days)
ℹ️ Under WLC/Forti:     {stats['inferred']}

Please review the attached Excel file.

Best regards,
QRadar Automation System
"""
            create_outlook_draft(draft_path, f"QRadar Report - {len(result_df)} Items", body)
            
        except PermissionError:
             print(f"❌ ERROR: Could not save report to '{draft_path}'. Is the file open?")


def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker (Optimized)...")
    
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        return

    # Reading Excel
    print(f"\n📖 Reading Excel file: {INPUT_EXCEL_PATH}")
    try:
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
    except Exception as e:
        print(f"❌ Failed to read Excel file: {e}")
        return

    to_process = list(all_sheets.keys()) if SHEETS_TO_PROCESS == ['all'] else SHEETS_TO_PROCESS
    
    for sheet in to_process:
        if sheet in all_sheets:
            all_sheets[sheet] = process_sheet(
                all_sheets[sheet], sheet,
                QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD,
                LOGSOURCE_COLUMN, IP_COLUMN, IN_QRADAR_COLUMN
            )

    print(f"\n💾 Saving updates to original Excel...")
    try:
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
            for name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
        print("✅ Original Excel file updated successfully.")
    except PermissionError:
        print(f"❌ ERROR: Permission Denied! The file '{INPUT_EXCEL_PATH}' is OPEN. Close it and re-run.")
    except Exception as e:
        print(f"❌ Failed to save Excel file: {e}")

    filter_and_email(all_sheets, DRAFT_OUTPUT_PATH)
    print(f"\n✅ Completed!")

if __name__ == '__main__':
    main()
