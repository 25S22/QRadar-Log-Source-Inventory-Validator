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
IN_QRADAR_COLUMN = 'In Qradar?'  # The column to check for "Yes"
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
        print(f"   ⚠️ Error parsing timestamp {timestamp_ms}: {e}")
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Get log source details directly from the API.
    Uses 'ilike' with wildcards for Name to allow partial matches.
    Uses 'contains' for IP protocol parameters.
    """
    
    # Clean the identifier
    clean_identifier = str(identifier).replace('"', '').replace("'", "").strip()
    
    # Construct Filter
    if is_ip:
        # IPs are inside a list (protocol_parameters), MUST use 'contains'
        query_filter = f'protocol_parameters contains value="{clean_identifier}"'
    else:
        # Names are strings, MUST use 'ilike' with wildcards for partial match
        # This solves the 422 Error and finds "ABC123" inside "Hello @ ABC123"
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
            print(f"   ❌ Log source API error: {resp.status_code} - {resp.text}")
            return {'status': f'API Error {resp.status_code}', **_empty_details()}

        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', **_empty_details()}

        # Select Best Match
        found_source = None
        
        if is_ip:
            # Strict validation for IPs to avoid partial number matches
            for source in ls_data:
                params = source.get('protocol_parameters', [])
                if any(p.get('value') == clean_identifier for p in params):
                    found_source = source
                    break
            if not found_source: 
                found_source = ls_data[0]
        else:
            # For partial name matches, take the first result
            found_source = ls_data[0]

        ls_id = found_source.get('id')
        ls_name = found_source.get('name', identifier)
        
        # Get last event details
        last_event_time_ms = found_source.get('last_event_time')
        last_seen, activity_status, days_since_last_event = safe_timestamp_conversion(last_event_time_ms)
        enabled = found_source.get('enabled', False)
        enabled_str = 'Yes' if enabled else 'No'
            
        print(f"   📋 Found: {ls_name} | {activity_status} | {days_since_last_event if days_since_last_event is not None else 'N/A'} days ago")

        return {
            'status': 'Found',
            'qradar_id': str(ls_id) if ls_id is not None else '',
            'enabled': enabled_str,
            'last_seen': last_seen,
            'activity_status': activity_status,
            'days_since_last_event': days_since_last_event
        }

    except Exception as e:
        print(f"   ❌ Unexpected error: {e}")
        return {'status': f'Error: {str(e)[:50]}...', **_empty_details()}


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column):
    """Process a single sheet with 'In Qradar?' filtering and smart AP/FW logic"""
    print(f"\n📋 Processing sheet: {sheet_name}")
    
    # Initialize columns
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
    
    total = len(df)
    processed_count = 0
    skipped_count = 0
    
    print(f"Total rows in sheet: {total}")
    
    for idx, row in df.iterrows():
        
        # ─── NEW FILTERING LOGIC ───
        # Check if 'In Qradar?' column exists and filter
        should_process = True
        if IN_QRADAR_COLUMN in df.columns:
            in_qradar_val = str(row[IN_QRADAR_COLUMN]).strip()
            # Only process if it contains "Yes" (case-insensitive)
            # This handles 'Yes', 'yes', 'Yes-M', 'Yes-E', etc.
            if "yes" not in in_qradar_val.lower():
                should_process = False
                df.at[idx, 'remarks'] = "Skipped (In Qradar? != Yes)"
                skipped_count += 1
        
        if not should_process:
            continue
            
        processed_count += 1
        print(f"[{processed_count}] Processing row {idx + 1}...", end='\r')
        
        name_val = str(row[logsource_column]).strip()
        details = None
        
        # 1. Name Lookup (Partial Match allowed via 'ilike')
        if name_val and name_val.lower() not in ['nan', 'none', '', 'null']:
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)
        
        # 2. IP Fallback (Only if Name failed)
        if not details or details['status'] == 'Not Found':
            ip_val = str(row[ip_column]).strip()
            if ip_val and ip_val.lower() not in ['nan', 'none', '', 'null']:
                print(f"\n   🔁 Fallback to IP: '{ip_val}'")
                details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)
        
        # Initialize details if still missing
        if not details:
            details = {'status': 'Empty/Invalid', **_empty_details()}
        
        # ─── UPDATE LOGIC ───
        
        # Determine Final Status and Remarks
        if details['status'] == 'Found':
            final_status = 'Found'
            final_remark = "Found"
            activity_status = details['activity_status']
        else:
            # 3. SMART LOGIC for AP/FW when Not Found
            # Check if Name contains 'AP' or 'FW' to infer status
            if "AP" in name_val:
                final_status = 'Inferred'
                final_remark = "Under WLC (Inferred from Name)"
                activity_status = "Inferred"
                print(f"\n   ℹ️  Inferred: Under WLC")
            elif "FW" in name_val:
                final_status = 'Inferred'
                final_remark = "Under Forti (Inferred from Name)"
                activity_status = "Inferred"
                print(f"\n   ℹ️  Inferred: Under Forti")
            else:
                final_status = 'Not Found'
                final_remark = "❌ Not found by Name or IP - Please Check!"
                activity_status = "Not Found"
                print(f"\n   ❌ Not Found")

        # Write to DataFrame
        for k, v in details.items():
            if k in df.columns:
                df.at[idx, k] = v
        
        # Overwrite fields based on final logic
        df.at[idx, 'status'] = final_status
        df.at[idx, 'remarks'] = final_remark
        if activity_status != details['activity_status']:
            df.at[idx, 'activity_status'] = activity_status

        time.sleep(0.2)
    
    print(f"\n📊 Sheet {sheet_name} completed. Processed: {processed_count}, Skipped: {skipped_count}")
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
        print(f"✉️ Draft created: {attachment_path}")
    except Exception as e:
        print(f"❌ Failed to create Outlook draft: {e}")


def filter_and_email(df_dict, draft_path):
    """Filter inactive/error log sources and create report with accurate counts"""
    frames = []
    
    # Global Counters
    count_total_processed = 0 # Rows where 'In Qradar' was Yes
    count_found_active = 0
    count_found_inactive = 0
    count_inferred = 0
    count_not_found = 0
    count_api_errors = 0

    for name, df in df_dict.items():
        if 'status' in df.columns:
            
            # Count only rows that were processed (status is not null)
            processed_mask = df['status'].notna()
            count_total_processed += processed_mask.sum()
            
            # --- Logic to Categorize Rows ---
            # 1. API Errors
            is_error = df['status'].str.startswith('API Error', na=False)
            count_api_errors += is_error.sum()
            
            # 2. Not Found (Strictly those marked Not Found, excluding Inferred)
            is_not_found = (df['status'] == 'Not Found')
            count_not_found += is_not_found.sum()
            
            # 3. Inferred (AP/FW logic)
            is_inferred = (df['status'] == 'Inferred')
            count_inferred += is_inferred.sum()
            
            # 4. Found - Check Activity
            is_found = (df['status'] == 'Found')
            is_inactive = is_found & ((df['activity_status'] == 'Inactive') | (df['activity_status'] == 'No Activity'))
            is_active = is_found & (df['activity_status'] == 'Active')
            
            count_found_inactive += is_inactive.sum()
            count_found_active += is_active.sum()

            # --- Filter for Excel Report (Issues Only) ---
            # Include: Inactive, Not Found, Errors, and Inferred (for visibility)
            mask_report = is_inactive | is_not_found | is_error | is_inferred
            
            if mask_report.any():
                sub = df[mask_report].copy()
                
                # Update remarks for inactive items specifically
                sub.loc[is_inactive, 'remarks'] = f'Inactive - No events in last {ACTIVITY_THRESHOLD_DAYS} days'
                
                sub['sheet_name'] = name
                frames.append(sub)

    # Calculate Total Found (Active + Inactive + Inferred)
    total_found_and_inferred = count_found_active + count_found_inactive + count_inferred

    if not frames:
        print("✅ No issues found; skipping email.")
    else:
        result_df = pd.concat(frames, ignore_index=True)
        result_df.to_excel(draft_path, index=False)
        print(f"💾 Report saved to: {draft_path}")

        # Email Body
        subject = f"QRadar Report - {len(result_df)} Items Flagged"
        body = f"""Hello,

Attached is the QRadar log source status report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.

📊 EXECUTIVE SUMMARY
----------------------------------------
Total Assets Processed:  {count_total_processed} (Only 'In Qradar' = Yes)
Total Found / Inferred: {total_found_and_inferred}
Total Not Found:       {count_not_found}
API Errors:            {count_api_errors}

📉 HEALTH BREAKDOWN
----------------------------------------
✅ Active:              {count_found_active}
⚠️ Inactive:            {count_found_inactive}  (No events in {ACTIVITY_THRESHOLD_DAYS} days)
ℹ️ Under WLC/Forti:     {count_inferred}      (Inferred from Name)

Please review the attached Excel file for the detailed list of inactive and missing log sources.

Best regards,
QRadar Automation System
"""
        create_outlook_draft(draft_path, subject, body)


def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker...")
    
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        return

    try:
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
    except Exception as e:
        print(f"❌ Failed to read Excel: {e}")
        return

    to_process = list(all_sheets.keys()) if SHEETS_TO_PROCESS == ['all'] else SHEETS_TO_PROCESS
    
    for sheet in to_process:
        if sheet in all_sheets:
            all_sheets[sheet] = process_sheet(
                all_sheets[sheet], sheet,
                QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD,
                LOGSOURCE_COLUMN, IP_COLUMN
            )

    print(f"\n💾 Saving original Excel...")
    try:
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
            for name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
    except Exception as e:
        print(f"❌ Failed to save Excel: {e}")

    filter_and_email(all_sheets, DRAFT_OUTPUT_PATH)
    print(f"\n✅ Completed!")

if __name__ == '__main__':
    main()
