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
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Get log source details directly from the API.
    Returns details dict or None if not found/error.
    """
    
    clean_identifier = str(identifier).replace('"', '').replace("'", "").strip()
    
    # Construct Filter
    if is_ip:
        # IP Search: Strict check inside protocol_parameters list
        query_filter = f'protocol_parameters contains value="{clean_identifier}"'
    else:
        # Name Search: Partial match using 'ilike' with wildcards
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
    1. CLEANS HEADERS.
    2. Filters 'Yes' rows first.
    3. Handles Disabled logic so they don't count as Inactive.
    """
    print(f"\n{'='*60}")
    print(f"📋 Processing Sheet: {sheet_name}")
    print(f"{'='*60}")
    
    # ─── FIX: CLEAN COLUMN HEADERS ───
    if not df.empty:
        df.columns = df.columns.str.strip()

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
        print(f"❌ ERROR: Column '{in_qradar_col}' not found in '{sheet_name}'!")
        print(f"   Available Columns: {list(df.columns)}")
        print("   Skipping this sheet.")
        return df

    # Create a mask for rows where In Qradar contains "yes" (case-insensitive)
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
        
        # Cleanup
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
            df.at[idx, 'qradar_id'] = details['qradar_id']
            df.at[idx, 'enabled'] = details['enabled']
            df.at[idx, 'last_seen'] = details['last_seen']
            df.at[idx, 'days_since_last_event'] = details['days_since_last_event']
            
            # --- NEW DISABLED LOGIC ---
            if details['enabled'] == 'No':
                # If disabled, we ignore its activity status
                final_status = 'Found'
                final_remark = "Disabled on QRadar"
                activity_status = "Disabled"
                
                print(f"      📌 Log Source: {details['actual_name']}")
                print(f"      ⚪ Status:     Disabled (Ignored for Inactivity)")
            else:
                # If enabled, we trust the activity status
                final_status = 'Found'
                final_remark = f"Found by {search_method}"
                activity_status = details['activity_status']
                
                print(f"      📌 Log Source: {details['actual_name']}")
                print(f"      📊 Activity:   {activity_status}")
                print(f"      📅 Last Event: {details['last_seen']} ({details['days_since_last_event']} days ago)")
                
            df.at[idx, 'status'] = final_status
            df.at[idx, 'remarks'] = final_remark
            df.at[idx, 'activity_status'] = activity_status
            
        else:
            # Handle Not Found Logic with Smart Inference
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

        time.sleep(0.2) # Slight delay
    
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
    """
    Filter inactive/error log sources and create report.
    UPDATED: Excludes 'Disabled' and 'Inferred' items from the Issue Counts.
    """
    frames = []
    
    # Global Counters
    stats = {
        'total_scanned': 0,
        'found_active': 0,
        'found_inactive': 0,
        'found_disabled': 0,  # New counter
        'inferred': 0,
        'not_found': 0,
        'errors': 0
    }

    for name, df in df_dict.items():
        if 'status' not in df.columns: continue
        
        # Only process rows that weren't skipped
        processed_df = df[df['status'].notna()]
        stats['total_scanned'] += len(processed_df)
        
        # --- Calculate Counts ---
        
        # 1. Active (Found + Active + Enabled)
        stats['found_active'] += len(processed_df[(processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Active')])
        
        # 2. Inactive (Found + Inactive + Enabled)
        # We explicitly check activity_status != Disabled here just in case, though filtering by Inactive handles it
        mask_inactive = (processed_df['status'] == 'Found') & ((processed_df['activity_status'] == 'Inactive') | (processed_df['activity_status'] == 'No Activity'))
        stats['found_inactive'] += mask_inactive.sum()
        
        # 3. Disabled (Found + Disabled)
        mask_disabled = (processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Disabled')
        stats['found_disabled'] += mask_disabled.sum()
        
        # 4. Inferred
        stats['inferred'] += len(processed_df[processed_df['status'] == 'Inferred'])
        
        # 5. Not Found
        mask_not_found = (processed_df['status'] == 'Not Found')
        stats['not_found'] += mask_not_found.sum()
        
        # 6. Errors
        mask_error = processed_df['status'].str.startswith('API Error', na=False)
        stats['errors'] += mask_error.sum()

        # --- Filter for Excel Report (ACTIONABLE ISSUES ONLY) ---
        # Exclude 'Disabled' and 'Inferred' from the attachment
        mask_report = mask_inactive | mask_not_found | mask_error
        
        if mask_report.any():
            sub = processed_df[mask_report].copy()
            # Update remark for inactive items specifically
            sub.loc[mask_inactive, 'remarks'] = f'Inactive - No events in last {ACTIVITY_THRESHOLD_DAYS} days'
            sub['sheet_name'] = name
            frames.append(sub)

    # If no "Bad" items found, skip email
    if not frames:
        print("✅ No Actionable Issues detected; skipping email.")
        return

    # Check permissions before saving report
    try:
        result_df = pd.concat(frames, ignore_index=True)
        result_df.to_excel(draft_path, index=False)
        print(f"\n💾 Filtered report saved to: {draft_path}")

        # Total Issues = Rows in the attachment (Inactive + Not Found + Errors)
        total_issues = len(result_df)

        # Email Body
        subject = f"QRadar Action Report - {total_issues} Issues Require Attention"
        
        body = f"""Hello,

Attached is the QRadar log source status report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.

⚠️ ACTION REQUIRED: {total_issues} ISSUES
(Count includes Inactive sources and Missing assets only. Disabled sources are excluded.)

📊 SCAN SUMMARY
----------------------------------------
Total Assets Scanned:  {stats['total_scanned']} (Rows where 'In Qradar' = Yes)

🔴 ISSUES FOUND (Attached):
   - Inactive:          {stats['found_inactive']} (Enabled but no events)
   - Not Found:         {stats['not_found']} (Absent in QRadar)
   - API Errors:        {stats['errors']}

🟢 HEALTHY / IGNORED (Not Attached):
   - Active:            {stats['found_active']}
   - Disabled:          {stats['found_disabled']} (Intentionally disabled on QRadar)
   - Under WLC/Forti:   {stats['inferred']}

Please review the attached Excel file for the list of items requiring attention.

Best regards,
QRadar Automation System
"""
        create_outlook_draft(draft_path, subject, body)
        
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
    except PermissionError:
        print(f"❌ ERROR: Permission Denied! The file '{INPUT_EXCEL_PATH}' is OPEN. Close it and re-run.")
        return
    except FileNotFoundError:
        print(f"❌ ERROR: File not found at '{INPUT_EXCEL_PATH}'. Check the path.")
        return
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
