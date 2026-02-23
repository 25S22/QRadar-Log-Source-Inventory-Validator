import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time
import os
import win32com.client  # For creating draft emails in Outlook

# Ensure charts generate in the background without opening UI windows
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt 

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
    2. Performs a hard reset on specific columns to prevent stale data.
    3. Filters 'Yes' rows first.
    4. Handles Disabled logic so they don't count as Inactive.
    """
    print(f"\n{'='*60}")
    print(f"📋 Processing Sheet: {sheet_name}")
    print(f"{'='*60}")
    
    # ─── FIX: CLEAN COLUMN HEADERS ───
    if not df.empty:
        df.columns = df.columns.str.strip()

    if in_qradar_col not in df.columns:
        print(f"❌ ERROR: Column '{in_qradar_col}' not found in '{sheet_name}'!")
        print(f"   Available Columns: {list(df.columns)}")
        print("   Skipping this sheet.")
        return df

    # ─── HARD RESET / PRE-RUN CLEANSE ───
    # Empties all the columns populated by this script so every run starts completely fresh.
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
        if col in df.columns:
            df[col] = None  # Wipes existing data
        else:
            df[col] = pd.Series(dtype=dtype)

    # ─── STEP 1: EFFICIENT FILTERING ───
    # Create a mask for rows where In Qradar contains "yes" (case-insensitive)
    process_mask = df[in_qradar_col].astype(str).str.lower().str.contains("yes", na=False)
    
    rows_to_process = df[process_mask]
    total_rows = len(df)
    target_count = len(rows_to_process)
    
    print(f"📊 Total Rows: {total_rows}")
    print(f"🎯 Rows marked 'Yes': {target_count} (Rows to scan)")
    
    # Mark skipped rows clearly (the rest of their fields remain empty/None from the Hard Reset)
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
                print(f"⚠️ {details['status']}")
        
        # 2. Fallback to IP (Triggers if Name was not found OR if there was an API Error)
        if (not details or details['status'] != 'Found') and ip_val:
            print(f"   🔁 Fallback to IP: '{ip_val}' ... ", end="")
            details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)
            
            if details['status'] == 'Found':
                print("✅ Found!")
                search_method = "IP"
            else:
                print(f"❌ {details['status']}")

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
            status_val = details['status'] 
            remark_val = f"❌ {status_val}"
            act_val = "Error" if "Error" in status_val else "Not Found"
            
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
                print(f"      ❌ Result: {status_val.upper()}")

            df.at[idx, 'status'] = status_val
            df.at[idx, 'remarks'] = remark_val
            df.at[idx, 'activity_status'] = act_val
            
            # Clear other fields to be clean
            df.at[idx, 'last_seen'] = "N/A"
            df.at[idx, 'days_since_last_event'] = None

        time.sleep(0.2) # Slight delay
    
    return df


def generate_pie_chart(data_dict, title, filename):
    """
    Generates a pie chart based on the provided data dictionary, 
    saves it locally, and returns the file path.
    """
    # Filter out zero values so the chart is clean
    filtered_data = {k: v for k, v in data_dict.items() if v > 0}
    
    if not filtered_data:
        return None

    labels = list(filtered_data.keys())
    sizes = list(filtered_data.values())
    
    # Specific colors for clear reporting visibility in cyber defense
    color_map = {
        'Active': '#28a745',             # Green
        'Inactive': '#dc3545',           # Red
        'Not Found': '#6c757d',          # Grey
        'API Errors': '#fd7e14',         # Orange
        'Disabled/Inferred': '#17a2b8'   # Teal
    }
    
    colors = []
    for label in labels:
        colors.append(color_map.get(label, '#cccccc'))

    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.pie(
        sizes, 
        labels=labels, 
        colors=colors, 
        autopct='%1.1f%%', 
        startangle=140, 
        textprops={'fontsize': 9}, 
        wedgeprops={'edgecolor': 'white'}
    )
    ax.axis('equal') 
    
    plt.title(title, pad=15, fontsize=11, fontweight='bold')
    
    filepath = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), filename)
    plt.savefig(filepath, bbox_inches='tight', dpi=100)
    plt.close()
    
    return filepath


def create_html_outlook_draft(attachment_path, subject, html_body, image_paths):
    """
    Create Outlook HTML draft with embedded images and ephemeral cleanup.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        
        # Attach the Excel report
        if os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)

        # Attach and embed images via HTML Content-IDs
        for cid, img_path in image_paths.items():
            if img_path and os.path.exists(img_path):
                attachment = mail.Attachments.Add(img_path)
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
        
        mail.HTMLBody = html_body
        mail.Display()
        print(f"\n✉️  Email draft created successfully.")
        
        # EPHEMERAL CLEANUP: Delete the generated PNG files so they don't clutter your drive
        for cid, img_path in image_paths.items():
            if img_path and os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception as e:
                    print(f"⚠️ Could not delete temporary image {img_path}: {e}")
                
    except Exception as e:
        print(f"\n❌ Failed to create Outlook draft: {e}")


def filter_and_email(processed_sheets_only, draft_path):
    """
    Filter inactive/error log sources and create an HTML report.
    UPDATED: Excludes 'Disabled' and 'Inferred' items from the Issue Counts,
    generates separate Excel tabs per sheet, and embeds pie charts.
    """
    report_frames = {}
    sheet_stats = {}
    images_to_embed = {}
    
    global_stats = { 
        'Active': 0, 
        'Inactive': 0, 
        'Not Found': 0, 
        'API Errors': 0, 
        'Disabled/Inferred': 0 
    }

    for name, df in processed_sheets_only.items():
        if 'status' not in df.columns: 
            continue
            
        # Only process rows that weren't skipped
        processed_df = df[df['status'].notna()]
        
        if len(processed_df) == 0: 
            continue

        # --- Calculate Counts for Charts ---
        active_count = len(processed_df[(processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Active')])
        
        mask_inactive = (processed_df['status'] == 'Found') & ((processed_df['activity_status'] == 'Inactive') | (processed_df['activity_status'] == 'No Activity'))
        inactive_count = mask_inactive.sum()
        
        disabled_count = len(processed_df[(processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Disabled')])
        inferred_count = len(processed_df[processed_df['status'] == 'Inferred'])
        
        mask_not_found = (processed_df['status'] == 'Not Found')
        not_found_count = mask_not_found.sum()
        
        mask_error = processed_df['status'].astype(str).str.startswith('API Error', na=False)
        error_count = mask_error.sum()

        sheet_counts = {
            'Active': active_count,
            'Inactive': inactive_count,
            'Not Found': not_found_count,
            'API Errors': error_count,
            'Disabled/Inferred': disabled_count + inferred_count
        }
        
        sheet_stats[name] = sheet_counts
        
        for k in global_stats: 
            global_stats[k] += sheet_counts[k]

        # --- Filter for Excel Report (ACTIONABLE ISSUES ONLY) ---
        # Exclude 'Disabled' and 'Inferred' from the attachment
        mask_report = mask_inactive | mask_not_found | mask_error
        
        if mask_report.any():
            sub = processed_df[mask_report].copy()
            # Update remark for inactive items specifically
            sub.loc[mask_inactive, 'remarks'] = f'Inactive - No events in last {ACTIVITY_THRESHOLD_DAYS} days'
            report_frames[name] = sub 

    # If no "Bad" items found, skip email
    if not report_frames:
        print("✅ No Actionable Issues detected; skipping email.")
        return

    # Write actionable issues into individual tabs based on their original sheet names
    try:
        with pd.ExcelWriter(draft_path, engine='openpyxl') as writer:
            for sheet_name, df in report_frames.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"\n💾 Filtered report saved to: {draft_path}")
    except PermissionError:
        print(f"❌ ERROR: Could not save report to '{draft_path}'. Is the file open?")
        return

    # Generate Chart Images
    overall_cid = "overall_chart"
    overall_path = generate_pie_chart(global_stats, "Overall Inventory Status", "temp_overall.png")
    if overall_path:
        images_to_embed[overall_cid] = overall_path
    
    for name, counts in sheet_stats.items():
        cid = f"chart_{name.replace(' ', '_')}"
        chart_path = generate_pie_chart(counts, f"{name} Status", f"temp_{name}.png")
        if chart_path:
            images_to_embed[cid] = chart_path

    # Total Issues = Inactive + Not Found + Errors
    total_issues = global_stats['Inactive'] + global_stats['Not Found'] + global_stats['API Errors']

    # Build the HTML Body
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #0056b3;">QRadar Action Report - {total_issues} Issues Require Attention</h2>
        <p>Attached is the QRadar log source status report generated on <b>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</b>.</p>
        
        <p>⚠️ <b>ACTION REQUIRED: {total_issues} ISSUES</b><br>
        (Count includes Inactive sources and Missing assets only. Disabled sources are excluded.)</p>

        <table style="width: 100%; max-width: 600px; border-collapse: collapse; margin-bottom: 20px;">
            <tr style="background-color: #f8f9fa;">
                <td style="padding: 10px; border: 1px solid #dee2e6;"><b>Total Assets Scanned:</b></td>
                <td style="padding: 10px; border: 1px solid #dee2e6;">{sum(global_stats.values())}</td>
            </tr>
            <tr>
                <td style="padding: 10px; border: 1px solid #dee2e6; color: #dc3545;"><b>🔴 Actionable Issues:</b></td>
                <td style="padding: 10px; border: 1px solid #dee2e6;"><b>{total_issues}</b></td>
            </tr>
        </table>

        <h3>Overall System Health</h3>
    """
    
    if overall_cid in images_to_embed:
        html_body += f'<img src="cid:{overall_cid}"><br>'
        
    html_body += """
        <hr style="border: 1px solid #eee; margin: 30px 0;">
        <h3>Breakdown by Processed Sheets</h3>
    """
    
    for name, counts in sheet_stats.items():
        cid = f"chart_{name.replace(' ', '_')}"
        html_body += f"""
        <div style="margin-bottom: 30px;">
            <h4 style="margin-bottom: 5px; color: #444;">{name}</h4>
            <ul style="list-style-type: none; padding-left: 0; margin-bottom: 10px;">
                <li><span style="color: #dc3545; font-weight: bold;">Inactive:</span> {counts['Inactive']}</li>
                <li><span style="color: #6c757d; font-weight: bold;">Not Found:</span> {counts['Not Found']}</li>
                <li><span style="color: #fd7e14; font-weight: bold;">API Errors:</span> {counts['API Errors']}</li>
                <li><span style="color: #28a745; font-weight: bold;">Active:</span> {counts['Active']}</li>
            </ul>
        """
        if cid in images_to_embed:
            html_body += f'<img src="cid:{cid}">'
        html_body += "</div>"

    html_body += """
        <br><p style="font-size: 12px; color: #777;">Automated Cyber Defense Reporting</p>
    </body>
    </html>
    """
    
    subject = f"QRadar Action Report - {total_issues} Issues Require Attention"
    create_html_outlook_draft(draft_path, subject, html_body, images_to_embed)


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

    # ─── STRICT SHEET ISOLATION ─── 
    # Pass only the targeted sheets to the reporting function so old data doesn't sneak in
    processed_sheets_only = {k: v for k, v in all_sheets.items() if k in to_process}
    
    filter_and_email(processed_sheets_only, DRAFT_OUTPUT_PATH)
    
    print(f"\n✅ Completed!")

if __name__ == '__main__':
    main()
