import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time
import os
import json
import win32com.client  # For creating draft emails in Outlook
import numpy as np

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'
SHEETS_TO_PROCESS = ['Sheet1', 'Sheet2']  # or ['all'] for all sheets
LOGSOURCE_COLUMN = 'log source name'
IP_COLUMN = 'IP'
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL = False
DRAFT_OUTPUT_PATH = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), 'inactive_and_errors.xlsx')
ACTIVITY_THRESHOLD_DAYS = 7  # Consider log source inactive if no events in X days
REQUEST_TIMEOUT = 30
MAX_SEARCH_RETRIES = 20  # Increased from 10
SEARCH_RETRY_DELAY = 3   # Increased from 2

# Valid timestamp range (avoid dates before 1970 and after 2038 for 32-bit systems)
MIN_TIMESTAMP = 0  # Unix epoch start
MAX_TIMESTAMP = 2147483647  # 2038-01-19 (32-bit limit)
# ─── END CONFIGURATION ─────────────────────────────────────────────────────────


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
            headers={'Accept': 'application/json', 'Version': '14.0'}  # Added API version
        )
        
        if resp.status_code == 200:
            print("✅ QRadar connection successful!")
            version_info = resp.json()
            print(f"   QRadar Version: {version_info[0].get('version', 'Unknown')}")
            return True
        elif resp.status_code == 401:
            print("❌ Authentication failed! Check username/password.")
            return False
        else:
            print(f"⚠️ Unexpected response: {resp.status_code} - {resp.text}")
            return False
            
    except requests.exceptions.Timeout:
        print("❌ Connection timeout! Check network connectivity.")
        return False
    except requests.exceptions.ConnectionError:
        print("❌ Connection error! Check QRadar host URL.")
        return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


def _empty_details():
    """
    Return empty details structure with proper data types.
    FIX: returns None for days_since_last_event so it doesn't show as '0'.
    """
    return {
        'qradar_id': 'N/A',
        'enabled': 'Unknown',
        'last_seen': 'N/A',
        'activity_status': 'Not Found',
        'days_since_last_event': None  # Changed from 0 to None to avoid confusion
    }


def safe_timestamp_conversion(timestamp_ms):
    """
    Safely convert timestamp to datetime string with proper validation
    QRadar returns timestamps in milliseconds, so we need to divide by 1000
    Returns: (last_seen_str, activity_status, days_since_last_event)
    """
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    
    try:
        # Convert to int if it's a float
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)
        
        # Check if timestamp looks like milliseconds (> year 2100 in seconds)
        # If it's a very large number, it's likely milliseconds
        if timestamp_ms > 4102444800:  # Year 2100 in seconds
            timestamp_seconds = timestamp_ms / 1000.0
        else:
            timestamp_seconds = timestamp_ms
        
        # Validate timestamp is within reasonable range (after conversion)
        if timestamp_seconds <= MIN_TIMESTAMP or timestamp_seconds > MAX_TIMESTAMP:
            print(f"   ⚠️ Timestamp out of valid range: {timestamp_seconds}")
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None
        
        # Convert to datetime
        last_event_datetime = datetime.fromtimestamp(timestamp_seconds)
        last_seen = last_event_datetime.strftime('%Y-%m-%d %H:%M:%S')
        
        # Calculate days since last event
        current_time = datetime.now()
        time_diff = current_time - last_event_datetime
        days_since_last_event = time_diff.days
        
        # Check if recent enough to be considered active
        threshold_time = datetime.now() - timedelta(days=ACTIVITY_THRESHOLD_DAYS)
        
        if last_event_datetime > threshold_time:
            activity_status = 'Active'
        else:
            activity_status = 'Inactive'
            
        return last_seen, activity_status, days_since_last_event
        
    except (ValueError, TypeError, OSError, OverflowError) as e:
        print(f"   ⚠️ Error parsing timestamp {timestamp_ms}: {e}")
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Get log source details directly from the API.
    FIXED: Uses 'protocol_parameters contains value' for IP lookup to resolve Error 422.
    """
    
    # 1. CLEAN THE IDENTIFIER
    # Remove quotes and whitespace to prevent syntax errors in the query
    clean_identifier = str(identifier).replace('"', '').replace("'", "").strip()
    
    # 2. CONSTRUCT THE CORRECT FILTER
    # 'ip_address' is not a valid field. We must search inside protocol_parameters.
    if is_ip:
        query_filter = f'protocol_parameters contains value="{clean_identifier}"'
    else:
        query_filter = f'name="{clean_identifier}"'
    
    ls_endpoint = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"

    try:
        # Get log source details with all fields
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

        # 3. VERIFY RESULT (For IP Searches)
        # Because 'contains value' is broad, we double-check the result
        found_source = None
        
        if is_ip:
            for source in ls_data:
                # Look inside protocol parameters to find the exact IP match
                params = source.get('protocol_parameters', [])
                if any(p.get('value') == clean_identifier for p in params):
                    found_source = source
                    break
            # Fallback: if we didn't find exact match in loop, take the first one returned
            if not found_source: 
                found_source = ls_data[0]
        else:
            found_source = ls_data[0]

        ls_id = found_source.get('id')
        ls_name = found_source.get('name', identifier)
        
        print(f"   📋 Found log source: {ls_name} (ID: {ls_id})")

        # Get last_event_time directly from the API response (in milliseconds)
        last_event_time_ms = found_source.get('last_event_time')
        
        # Use safe timestamp conversion
        last_seen, activity_status, days_since_last_event = safe_timestamp_conversion(last_event_time_ms)
        
        # Get additional useful fields from the API
        enabled = found_source.get('enabled', False)
        
        # Get status string for enabled - ensure it's a string
        enabled_str = 'Yes' if enabled else 'No'
            
        print(f"   📊 Last Event: {last_seen} | Status: {activity_status} | Enabled: {enabled_str} | Days Since: {days_since_last_event}")

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
    """Process a single sheet with enhanced logging and error handling"""
    print(f"\n📋 Processing sheet: {sheet_name}")
    
    # Validate columns exist
    missing_cols = []
    for col in [logsource_column, ip_column]:
        if col not in df.columns:
            missing_cols.append(col)
    
    if missing_cols:
        print(f"❌ Missing columns in {sheet_name}: {missing_cols}")
        print(f"   Available columns: {list(df.columns)}")
        return df
    
    # Add result columns if they don't exist with proper data types
    result_columns_config = {
        'status': 'object',
        'qradar_id': 'object', 
        'enabled': 'object',
        'last_seen': 'object',
        'activity_status': 'object',
        'days_since_last_event': 'float64', # Changed to float to allow NaN (empty) values
        'remarks': 'object' # Added remarks column
    }
    
    for col, dtype in result_columns_config.items():
        if col not in df.columns:
            if dtype == 'object':
                df[col] = pd.Series(dtype=object)
            elif dtype == 'int64':
                df[col] = pd.Series(dtype='int64')
            elif dtype == 'float64':
                df[col] = pd.Series(dtype='float64')
    
    total = len(df)
    processed = 0
    found_count = 0
    
    print(f"Found {total} rows to process...")
    
    for idx, row in df.iterrows():
        processed += 1
        print(f"\n[{processed}/{total}] Processing row {idx + 1}")
        
        name_val = str(row[logsource_column]).strip()
        details = None
        
        # Try lookup by name first
        if name_val and name_val.lower() not in ['nan', 'none', '', 'null']:
            print(f"   🔍 Lookup by name: '{name_val}'")
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)
        
        # Fallback to IP lookup if name lookup failed
        if not details or details['status'] == 'Not Found':
            ip_val = str(row[ip_column]).strip()
            if ip_val and ip_val.lower() not in ['nan', 'none', '', 'null']:
                print(f"   🔁 Fallback to IP: '{ip_val}'")
                details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)
        
        # Use empty details if nothing found
        if not details:
            details = {'status': 'Empty/Invalid', **_empty_details()}
        
        # ─── UPDATE LOGIC ───
        
        # Update DataFrame with details
        for k, v in details.items():
            df.loc[idx, k] = v
        
        # Set Remarks and specific error messages
        if details['status'] == 'Found':
            found_count += 1
            df.loc[idx, 'remarks'] = "Found"
            icon = "✅"
            days_msg = f"{details.get('days_since_last_event', 'N/A')}"
        else:
            # THIS IS THE FIX: Explicit alert message
            df.loc[idx, 'remarks'] = "Not found by host name or IP - please check!"
            df.loc[idx, 'activity_status'] = "Not Found" # Force status
            icon = "❌"
            days_msg = "N/A"
            
        print(f"   {icon} Result: {details['status']} | Activity: {details['activity_status']} | Days Since: {days_msg}")
        
        # Add delay to avoid overwhelming QRadar
        time.sleep(0.5)
    
    print(f"\n📊 Sheet {sheet_name} completed: {found_count}/{total} log sources found")
    return df


def create_outlook_draft(attachment_path, subject, body):
    """Create Outlook draft with error handling"""
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(attachment_path)
        mail.Display()  # Pop up the draft window
        print(f"✉️ Draft created and displayed: {attachment_path}")
    except Exception as e:
        print(f"❌ Failed to create Outlook draft: {e}")
        print(f"   Email would have been: {subject}")
        print(f"   Attachment: {attachment_path}")


def filter_and_email(df_dict, draft_path):
    """Filter inactive/error log sources and create email draft"""
    frames = []
    
    for name, df in df_dict.items():
        if 'status' in df.columns and 'activity_status' in df.columns:
            # Filter inactive sources
            mask_inactive = (df['activity_status'] == 'Inactive') | (df['activity_status'] == 'No Activity')
            
            # Filter API errors
            mask_errors = df['status'].str.startswith('API Error', na=False)
            
            # Filter not found
            mask_not_found = (df['status'] == 'Not Found') | (df['activity_status'] == 'Not Found')

            # Add inactive sources
            if mask_inactive.any():
                sub = df[mask_inactive].copy()
                # Only add remark if it's not already set to the "Not Found" message
                sub.loc[sub['remarks'].isnull(), 'remark'] = 'Inactive or no activity detected'
                sub['sheet_name'] = name
                frames.append(sub)

            # Add error sources
            if mask_errors.any():
                sub_err = df[mask_errors].copy()
                sub_err['sheet_name'] = name
                frames.append(sub_err)
                
            # Add not found sources
            if mask_not_found.any():
                sub_nf = df[mask_not_found].copy()
                sub_nf['sheet_name'] = name
                frames.append(sub_nf)

    if not frames:
        print("✅ No inactive, error, or not found log sources; skipping email.")
        return

    result_df = pd.concat(frames, ignore_index=True)
    total = len(result_df)
    inactive_count = ((result_df['activity_status'] == 'Inactive') | (result_df['activity_status'] == 'No Activity')).sum()
    error_count = result_df['status'].str.startswith('API Error', na=False).sum()
    not_found_count = ((result_df['status'] == 'Not Found') | (result_df['activity_status'] == 'Not Found')).sum()

    # Save filtered results
    result_df.to_excel(draft_path, index=False)
    print(f"💾 Filtered report saved to: {draft_path}")

    # Create email
    subject = f"QRadar Log Source Status Report - {total} Issues Found"
    body = f"""Hello,

Attached is the QRadar log source status report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.

Summary:
- Total flagged log sources: {total}
- Inactive/No Activity: {inactive_count}
- API Errors: {error_count}
- Not Found: {not_found_count}

Please review the attached Excel file for detailed information and take appropriate action.

Activity threshold: {ACTIVITY_THRESHOLD_DAYS} days
Report includes log sources with no events in the last {ACTIVITY_THRESHOLD_DAYS} days or API/lookup errors.

Best regards,
QRadar Automation System
"""
    
    create_outlook_draft(draft_path, subject, body)


def main():
    """Main execution function"""
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker (Enhanced Version)...")
    print(f"⚙️ Configuration:")
    print(f"   - QRadar Host: {QRADAR_HOST}")
    print(f"   - Activity Threshold: {ACTIVITY_THRESHOLD_DAYS} days")
    print(f"   - Request Timeout: {REQUEST_TIMEOUT}s")
    print(f"   - SSL Verification: {VERIFY_SSL}")
    
    # Test connection
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        print("❌ Connection test failed. Please check your configuration.")
        return

    # Read Excel file
    print(f"\n📖 Reading Excel file: {INPUT_EXCEL_PATH}")
    try:
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
        sheets = list(all_sheets.keys())
        print(f"📄 Sheets found: {sheets}")
    except Exception as e:
        print(f"❌ Failed to read Excel file: {e}")
        return

    # Process sheets
    to_process = sheets if SHEETS_TO_PROCESS == ['all'] else SHEETS_TO_PROCESS
    print(f"📋 Sheets to process: {to_process}")
    
    for sheet in to_process:
        if sheet in all_sheets:
            print(f"\n{'='*50}")
            all_sheets[sheet] = process_sheet(
                all_sheets[sheet], sheet,
                QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD,
                LOGSOURCE_COLUMN, IP_COLUMN
            )
        else:
            print(f"⚠️ Sheet '{sheet}' not found in Excel file. Skipping...")

    # Save updated Excel
    print(f"\n💾 Saving updates to original Excel file...")
    try:
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
            for name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
        print("✅ Original Excel file updated successfully.")
    except Exception as e:
        print(f"❌ Failed to save Excel file: {e}")

    # Generate filtered report and email
    print(f"\n📧 Generating filtered report...")
    filter_and_email(all_sheets, DRAFT_OUTPUT_PATH)

    # Final summary
    print(f"\n📊 FINAL SUMMARY:")
    print("=" * 60)
    
    total_processed = 0
    total_found = 0
    total_active = 0
    total_inactive = 0
    total_errors = 0
    
    for sheet in to_process:
        if sheet in all_sheets:
            df = all_sheets[sheet]
            if 'status' in df.columns:
                sheet_total = len(df)
                sheet_found = (df['status'] == 'Found').sum()
                sheet_active = (df['activity_status'] == 'Active').sum()
                sheet_inactive = ((df['activity_status'] == 'Inactive') | (df['activity_status'] == 'No Activity')).sum()
                sheet_errors = (df['status'].str.startswith('API Error', na=False) | (df['status'] == 'Not Found')).sum()
                
                print(f"📋 {sheet}:")
                print(f"   Total: {sheet_total}, Found: {sheet_found}, Active: {sheet_active}, Inactive: {sheet_inactive}, Errors: {sheet_errors}")
                
                total_processed += sheet_total
                total_found += sheet_found
                total_active += sheet_active
                total_inactive += sheet_inactive
                total_errors += sheet_errors
    
    print("=" * 60)
    print(f"🎯 OVERALL TOTALS:")
    print(f"   Processed: {total_processed}")
    print(f"   Found: {total_found}")
    print(f"   Active: {total_active}")
    print(f"   Inactive: {total_inactive}")
    print(f"   Errors: {total_errors}")
    
    if total_processed > 0:
        success_rate = (total_found / total_processed) * 100
        active_rate = (total_active / total_found) * 100 if total_found > 0 else 0
        print(f"   Success Rate: {success_rate:.1f}%")
        print(f"   Active Rate: {active_rate:.1f}%")
    
    print(f"\n✅ QRadar Log Source Checker completed!")
    print(f"📁 Updated Excel: {INPUT_EXCEL_PATH}")
    print(f"📧 Filtered Report: {DRAFT_OUTPUT_PATH}")


if __name__ == '__main__':
    main()
