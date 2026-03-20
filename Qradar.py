import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time
import os
import tempfile
import win32com.client
import openpyxl
from openpyxl.styles import PatternFill
import concurrent.futures
import logging
import traceback

# Ensure charts generate in the background without opening UI windows
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# ─── LOGGING SETUP ─────────────────────────────────────────────────────────────
# Logs warnings and above to console. Change level to logging.DEBUG for verbose output.
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
INPUT_EXCEL_PATH    = r'C:\path\to\your\input.xlsx'
SHEETS_TO_PROCESS   = ['Sheet1', 'Sheet2']   # or ['all'] for all sheets
LOGSOURCE_COLUMN    = 'log source name'
IP_COLUMN           = 'IP'
IN_QRADAR_COLUMN    = 'In Qradar?'           # Defines if it should be scanned, skipped, or pended

QRADAR_HOST         = 'https://your-qradar-host'
QRADAR_USERNAME     = 'your-username'
QRADAR_PASSWORD     = 'your-password'
VERIFY_SSL          = False

DRAFT_OUTPUT_PATH        = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), 'inactive_and_errors.xlsx')
ACTIVITY_THRESHOLD_DAYS  = 7     # Consider log source inactive if no events in X days
REQUEST_TIMEOUT          = 30
MAX_WORKERS              = 10    # Number of simultaneous API requests to QRadar

# Expected Log Source Types (Fuzzy Matching Applied)
# You only need to put the core keywords here now.
# "Microsoft Security" will catch "Microsoft Windows Security Event Log".
EXPECTED_LS_TYPES = ['Microsoft Security', 'Linux OS']

# ─── GROUP THRESHOLD CONFIGURATION ────────────────────────────────────────────
# Set GROUP_COLUMN to the exact Excel column header name that contains group names.
# Leave as None to disable this feature entirely.
#
# Example: GROUP_COLUMN = 'Group'
GROUP_COLUMN = None

# Map group names (case-insensitive) to their inactivity threshold in days.
# Any group not listed here falls back to ACTIVITY_THRESHOLD_DAYS above.
# This config is only read when GROUP_COLUMN is not None.
#
# Example:
# GROUP_THRESHOLDS = {
#     'Critical Assets': 3,
#     'DMZ Servers':     5,
#     'Branch Offices':  14,
# }
GROUP_THRESHOLDS = {}

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

# Valid timestamp range for QRadar API responses
MIN_TIMESTAMP = 0
MAX_TIMESTAMP = 2147483647

# Global Cache to prevent API hammering when resolving Log Source Type IDs
LOG_SOURCE_TYPES_CACHE = {}

# ─── MAPI PROPERTY CONSTANT ────────────────────────────────────────────────────
# PR_ATTACH_CONTENT_ID (0x3712001F) is a standard MAPI property tag used by the
# local Outlook COM object to link an attachment to an HTML <img src="cid:..."> tag.
# This is a pure string identifier for the local COM call — no network request is
# made to any external address at any point. The "http://schemas..." prefix is
# just the MAPI schema URI naming convention, identical to an XML namespace string.
_MAPI_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"


# ─── GROUP THRESHOLD RESOLVER ──────────────────────────────────────────────────
def resolve_threshold(group_name=None):
    """
    Resolves the correct inactivity threshold (in days) for a given group name.

    - If GROUP_COLUMN is None (feature disabled): always returns ACTIVITY_THRESHOLD_DAYS.
    - If group_name is empty / null: falls back to ACTIVITY_THRESHOLD_DAYS.
    - Matches against GROUP_THRESHOLDS keys case-insensitively (strips whitespace).
    - No match found: falls back to ACTIVITY_THRESHOLD_DAYS.
    """
    if not GROUP_COLUMN:
        return ACTIVITY_THRESHOLD_DAYS

    if not group_name or str(group_name).strip().lower() in ('nan', 'none', '', 'null'):
        return ACTIVITY_THRESHOLD_DAYS

    group_clean = str(group_name).strip().lower()

    for key, value in GROUP_THRESHOLDS.items():
        if str(key).strip().lower() == group_clean:
            return value

    return ACTIVITY_THRESHOLD_DAYS


# ─── CHART STATS HELPER ────────────────────────────────────────────────────────
def _chart_stats(stats_dict):
    """
    Returns a copy of a stats dict suitable for pie charts only:
    Inferred is folded into Active so it shows as healthy/green on the chart.
    The original stats_dict is NOT mutated — text breakdowns still show Inferred separately.
    """
    d = dict(stats_dict)
    d['Active'] = d.get('Active', 0) + d.pop('Inferred', 0)
    return d


def test_qradar_connection(qradar_host, username, password):
    """
    Test QRadar connection and validate credentials before starting heavy processing.
    """
    print("🔗 Testing QRadar connection...")
    qradar_host = qradar_host.rstrip('/')
    endpoint    = f"{qradar_host}/api/help/versions"

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

    except requests.exceptions.Timeout:
        print("❌ Connection timed out. Check QRADAR_HOST and network connectivity.")
        return False
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Connection error: {e}")
        return False
    except Exception as e:
        logger.error("Unexpected error during connection test:\n%s", traceback.format_exc())
        print(f"❌ Connection failed: {e}")
        return False


def fetch_log_source_types(qradar_host, username, password):
    """
    PRE-FETCH CACHE: Downloads the master dictionary of Log Source Type IDs
    to their string Names so worker threads don't have to constantly query the API.
    Must be called and fully populated before any worker threads are spawned.
    """
    print("📥 Fetching Log Source Types Dictionary into memory...")
    qradar_host = qradar_host.rstrip('/')
    endpoint    = f"{qradar_host}/api/config/event_sources/log_source_management/log_source_types"

    try:
        resp = requests.get(
            endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )

        if resp.status_code == 200:
            types_data = resp.json()
            for t in types_data:
                ls_id   = t.get('id')
                ls_name = t.get('name')
                if ls_id is not None and ls_name is not None:
                    LOG_SOURCE_TYPES_CACHE[ls_id] = ls_name
            print(f"✅ Successfully cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
        else:
            print(f"⚠️ Failed to fetch Log Source Types. API returned {resp.status_code}.")

    except requests.exceptions.Timeout:
        print("❌ Timed out while fetching Log Source Types.")
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Connection error while fetching Log Source Types: {e}")
    except Exception as e:
        logger.error("Unexpected error in fetch_log_source_types:\n%s", traceback.format_exc())
        print(f"❌ Error fetching Log Source Types: {e}")


def _empty_details():
    """
    Return an empty details structure to ensure dictionary keys
    always exist even if the log source is not found.
    """
    return {
        'qradar_id':             'N/A',
        'enabled':               'Unknown',
        'last_seen':             'N/A',
        'activity_status':       'Not Found',
        'days_since_last_event': None,
        'actual_name':           'N/A',
        'ls_type':               'N/A',
        'is_older_expected':     False
    }


def safe_timestamp_conversion(timestamp_ms, threshold=None):
    """
    Safely convert the QRadar epoch timestamp to a human-readable datetime string
    and calculate how many days have passed since the last event.

    threshold: optional override in days. If None, uses ACTIVITY_THRESHOLD_DAYS.
    """
    effective_threshold = threshold if threshold is not None else ACTIVITY_THRESHOLD_DAYS

    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None

    try:
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)

        # Convert milliseconds to seconds if needed
        if timestamp_ms > 4102444800:
            timestamp_seconds = timestamp_ms / 1000.0
        else:
            timestamp_seconds = timestamp_ms

        # Validate timestamp boundaries
        if timestamp_seconds <= MIN_TIMESTAMP or timestamp_seconds > MAX_TIMESTAMP:
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None

        last_event_datetime = datetime.fromtimestamp(timestamp_seconds)
        last_seen           = last_event_datetime.strftime('%Y-%m-%d %H:%M:%S')

        time_diff              = datetime.now() - last_event_datetime
        days_since_last_event  = time_diff.days

        threshold_time = datetime.now() - timedelta(days=effective_threshold)

        activity_status = 'Active' if last_event_datetime > threshold_time else 'Inactive'

        return last_seen, activity_status, days_since_last_event

    except Exception as e:
        logger.error("Timestamp conversion error for value %s:\n%s", timestamp_ms, traceback.format_exc())
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False, threshold=None):
    """
    Get log source details directly from the QRadar API.
    Implements FUZZY MATCHING and ABSOLUTE TYPE-PRIORITY OVERRIDE to guarantee
    expected log source types are chosen over anything else.
    """
    clean_identifier = str(identifier).replace('"', '').replace("'", "").strip()

    if is_ip:
        query_filter = f'protocol_parameters contains value="{clean_identifier}" or name ilike "%{clean_identifier}%"'
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

        # ─── 1. PRE-FILTER VALID MATCHES ───
        valid_sources = []
        if is_ip:
            for src in ls_data:
                params   = src.get('protocol_parameters', [])
                in_params = any(p.get('value') == clean_identifier for p in params)
                in_name   = clean_identifier.lower() in str(src.get('name', '')).lower()
                if in_params or in_name:
                    valid_sources.append(src)
        else:
            valid_sources = ls_data

        if not valid_sources:
            return {'status': 'Not Found', **_empty_details()}

        # ─── 2. FUZZY MATCH & TYPE SEPARATION ───
        expected_sources   = []
        unexpected_sources = []

        for src in valid_sources:
            type_name      = LOG_SOURCE_TYPES_CACHE.get(src.get('type_id'), "")
            api_name_clean = str(type_name).lower()

            is_match = False
            for exp in EXPECTED_LS_TYPES:
                exp_words = str(exp).lower().split()
                if all(w in api_name_clean for w in exp_words):
                    is_match = True
                    break

            if is_match:
                expected_sources.append(src)
            else:
                unexpected_sources.append(src)

        def get_best_source(source_list):
            if not source_list:
                return None
            enabled  = [s for s in source_list if s.get('enabled') is True]
            disabled = [s for s in source_list if s.get('enabled') is False]
            enabled.sort(key=lambda x: x.get('last_event_time') or 0, reverse=True)
            disabled.sort(key=lambda x: x.get('last_event_time') or 0, reverse=True)
            return enabled[0] if enabled else disabled[0]

        # ─── 3. ABSOLUTE TYPE-PRIORITY OVERRIDE ───
        found_source      = None
        is_older_expected = False

        absolute_max_time = max([s.get('last_event_time') or 0 for s in valid_sources])

        if expected_sources:
            found_source = get_best_source(expected_sources)
            if (found_source.get('last_event_time') or 0) < absolute_max_time:
                is_older_expected = True
        else:
            found_source = get_best_source(unexpected_sources)

        # ─── 4. EXTRACT FINAL DETAILS ───
        ls_id   = found_source.get('id')
        ls_name = found_source.get('name', identifier)
        type_id = found_source.get('type_id')

        ls_type_name = LOG_SOURCE_TYPES_CACHE.get(type_id, f"Unknown Type ID: {type_id}")

        last_event_time_ms                            = found_source.get('last_event_time')
        last_seen, activity_status, days_since_last_event = safe_timestamp_conversion(
            last_event_time_ms, threshold=threshold
        )

        enabled_str = 'Yes' if found_source.get('enabled', False) else 'No'

        return {
            'status':                'Found',
            'qradar_id':             str(ls_id) if ls_id is not None else '',
            'actual_name':           ls_name,
            'ls_type':               ls_type_name,
            'enabled':               enabled_str,
            'last_seen':             last_seen,
            'activity_status':       activity_status,
            'days_since_last_event': days_since_last_event,
            'is_older_expected':     is_older_expected
        }

    except requests.exceptions.Timeout:
        logger.warning("Request timed out for identifier: %s", identifier)
        return {'status': 'Timeout', **_empty_details()}
    except requests.exceptions.ConnectionError as e:
        logger.error("Connection error for identifier %s: %s", identifier, e)
        return {'status': 'Connection Error', **_empty_details()}
    except Exception as e:
        logger.error("Unexpected error for identifier %s:\n%s", identifier, traceback.format_exc())
        return {'status': f'Error: {str(e)[:50]}', **_empty_details()}


def process_single_row(idx, name_val, ip_val, qradar_host, username, password, group_val=None):
    """
    WORKER FUNCTION FOR THREADING.
    group_val: raw group string from the Excel column (or None if feature disabled).
    """
    if name_val and str(name_val).lower() in ['nan', 'none', '', 'null']: name_val = None
    if ip_val   and str(ip_val).lower()   in ['nan', 'none', '', 'null']: ip_val   = None

    effective_threshold = resolve_threshold(group_val)

    details       = None
    search_method = "None"

    # 1. Primary Search by Name
    if name_val:
        details = get_log_source_details(
            qradar_host, username, password, name_val,
            is_ip=False, threshold=effective_threshold
        )
        if details['status'] == 'Found':
            search_method = "Name"

    # 2. Fallback Search by IP
    if (not details or details['status'] != 'Found') and ip_val:
        details = get_log_source_details(
            qradar_host, username, password, ip_val,
            is_ip=True, threshold=effective_threshold
        )
        if details['status'] == 'Found':
            search_method = "IP"

    if not details:
        details = {'status': 'Empty/Invalid', **_empty_details()}

    return idx, name_val, details, search_method


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column, in_qradar_col):
    """
    Optimised & Threaded Sheet Processing.
    """
    print(f"\n{'='*60}")
    print(f"📋 Processing Sheet: {sheet_name} (Multi-threaded execution)")
    print(f"{'='*60}")

    if not df.empty:
        df.columns = df.columns.str.strip()

    # ─── FIX 6: Guard ALL required columns upfront ───
    required_columns = [in_qradar_col, logsource_column, ip_column]
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ ERROR: The following required column(s) are missing from sheet '{sheet_name}': {missing}")
        print(f"   Available columns: {list(df.columns)}")
        print(f"   Skipping sheet.")
        return df

    # ─── HARD RESET / PRE-RUN CLEANSE ───
    cols_to_init = {
        'status':                'object',
        'qradar_id':             'object',
        'enabled':               'object',
        'last_seen':             'object',
        'activity_status':       'object',
        'days_since_last_event': 'float64',
        'remarks':               'object',
        'QRadar Actual Name':    'object',
        'Log Source Type':       'object',
        'Is Older Expected':     'bool'
    }

    for col, dtype in cols_to_init.items():
        if col in df.columns:
            df[col] = None
        else:
            df[col] = pd.Series(dtype=dtype)

    # ─── FIX 2: Build a reliable Excel row index map before the loop ───
    # Maps DataFrame index → actual Excel row number (1-based header + 1 offset)
    row_index_map = {
        df_idx: excel_row
        for excel_row, df_idx in enumerate(df.index, start=2)
    }

    # ─── EFFICIENT FILTERING MASKS ───
    in_qradar_series = df[in_qradar_col].astype(str).str.lower()
    process_mask     = in_qradar_series.str.contains("yes",                 na=False)
    pending_mask     = in_qradar_series.str.contains("pending-maintenance", na=False)

    rows_to_process = df[process_mask]
    total_rows      = len(df)
    target_count    = len(rows_to_process)
    pending_count   = pending_mask.sum()

    # Determine if group feature is usable for this specific sheet
    group_feature_active = bool(GROUP_COLUMN and GROUP_COLUMN in df.columns)
    if GROUP_COLUMN and not group_feature_active:
        print(f"⚠️  GROUP_COLUMN '{GROUP_COLUMN}' not found in sheet '{sheet_name}' — using global threshold.")

    print(f"📊 Total Rows: {total_rows} | 🎯 'Yes' (To Scan): {target_count} | ⏳ Pending: {pending_count}")
    if group_feature_active:
        print(f"🏷️  Group-based thresholds ACTIVE (column: '{GROUP_COLUMN}')")

    skipped_mask = ~(process_mask | pending_mask)
    df.loc[skipped_mask, 'remarks'] = "Skipped (Not Yes or Pending)"

    if pending_count > 0:
        df.loc[pending_mask, 'status']          = 'Pending-Maintenance'
        df.loc[pending_mask, 'activity_status'] = 'Pending-Maintenance'
        df.loc[pending_mask, 'remarks']         = 'Pending Maintenance (Not Scanned)'
        df.loc[pending_mask, 'last_seen']       = 'N/A'

    if target_count == 0:
        return df

    # ─── MULTI-THREADING EXECUTION ───
    processed_count = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

        futures = {
            executor.submit(
                process_single_row,
                idx,
                str(row[logsource_column]).strip(),
                str(row[ip_column]).strip(),
                qradar_host,
                username,
                password,
                str(row[GROUP_COLUMN]).strip() if group_feature_active else None
            ): idx for idx, row in rows_to_process.iterrows()
        }

        for future in concurrent.futures.as_completed(futures):
            processed_count += 1

            # ─── FIX 5: Catch worker thread exceptions — don't let one crash kill the run ───
            try:
                idx, name_val, details, search_method = future.result()
            except Exception as worker_exc:
                original_idx = futures[future]
                logger.error(
                    "Worker thread crashed for row index %s:\n%s",
                    original_idx, traceback.format_exc()
                )
                print(f"\n⚠️  [{processed_count}/{target_count}] Worker crashed for row {original_idx}: {worker_exc}")
                df.at[original_idx, 'status']        = 'Worker Error'
                df.at[original_idx, 'remarks']       = f'Thread exception: {str(worker_exc)[:80]}'
                df.at[original_idx, 'activity_status'] = 'Error'
                df.at[original_idx, 'last_seen']     = 'N/A'
                continue

            print(f"\n🔹 [{processed_count}/{target_count}] Resolving: {name_val or 'Unknown'} -> {details['status']}")

            df.at[idx, 'QRadar Actual Name'] = details['actual_name']
            df.at[idx, 'Log Source Type']    = details['ls_type']
            df.at[idx, 'Is Older Expected']  = details.get('is_older_expected', False)

            if details['status'] == 'Found':
                df.at[idx, 'qradar_id']             = details['qradar_id']
                df.at[idx, 'enabled']               = details['enabled']
                df.at[idx, 'last_seen']             = details['last_seen']
                df.at[idx, 'days_since_last_event'] = details['days_since_last_event']

                base_remark = f"Found by {search_method}"

                if details.get('is_older_expected'):
                    base_remark += " | ⚠️ Bypassed newer unexpected source"
                    print(f"      🚨 WARNING: Bypassed a newer unexpected log source to lock onto this expected one!")

                if details['enabled'] == 'No':
                    df.at[idx, 'status']          = 'Found'
                    df.at[idx, 'remarks']         = "Disabled on QRadar"
                    df.at[idx, 'activity_status'] = "Disabled"
                    print(f"      📌 Log Source: {details['actual_name']} [{details['ls_type']}]")
                    print(f"      ⚪ Status:     Disabled (Ignored for Inactivity)")
                else:
                    df.at[idx, 'status']          = 'Found'
                    df.at[idx, 'remarks']         = base_remark
                    df.at[idx, 'activity_status'] = details['activity_status']
                    print(f"      📌 Log Source: {details['actual_name']} [{details['ls_type']}]")
                    print(f"      📊 Activity:   {details['activity_status']}")
                    print(f"      📅 Last Event: {details['last_seen']} ({details['days_since_last_event']} days ago)")

            else:
                status_val = details['status']
                remark_val = f"❌ {status_val}"

                if "Error" in status_val or "Timeout" in status_val:
                    act_val = "Error"
                else:
                    act_val = "Not Found"

                if name_val and "AP" in name_val:
                    status_val = "Inferred"
                    remark_val = "Under WLC (Inferred)"
                    act_val    = "Inferred"
                    print("      ℹ️  Result: Inferred as WLC (AP in name)")

                elif name_val and "FW" in name_val:
                    status_val = "Inferred"
                    remark_val = "Under Forti (Inferred)"
                    act_val    = "Inferred"
                    print("      ℹ️  Result: Inferred as FortiGate (FW in name)")

                else:
                    print(f"      ❌ Result: {status_val.upper()}")

                df.at[idx, 'status']                = status_val
                df.at[idx, 'remarks']               = remark_val
                df.at[idx, 'activity_status']       = act_val
                df.at[idx, 'last_seen']             = "N/A"
                df.at[idx, 'days_since_last_event'] = None

    return df


def generate_pie_chart(data_dict, title, prefix='qradar_chart'):
    """
    Generates a pie chart, saves it to a secure temp file, and returns the file path.
    Uses tempfile so names are unpredictable and isolated per run — no stale file risk.
    """
    filtered_data = {k: v for k, v in data_dict.items() if v > 0}
    if not filtered_data:
        return None

    labels = list(filtered_data.keys())
    sizes  = list(filtered_data.values())

    color_map = {
        'Active':              '#28a745',
        'Inactive':            '#dc3545',
        'Not Found':           '#6c757d',
        'API Errors':          '#fd7e14',
        'Disabled':            '#17a2b8',
        'Inferred':            '#6f42c1',
        'Pending-Maintenance': '#007bff'
    }

    colors = [color_map.get(label, '#cccccc') for label in labels]

    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    ax.pie(
        sizes, labels=labels, colors=colors, autopct='%1.1f%%',
        startangle=140, textprops={'fontsize': 9}, wedgeprops={'edgecolor': 'white'}
    )
    ax.axis('equal')
    plt.title(title, pad=15, fontsize=11, fontweight='bold')

    # ─── FIX 3 & 4 (chart): Use tempfile — no predictable names, no path traversal ───
    tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False, prefix=f'{prefix}_')
    filepath = tmp.name
    tmp.close()

    plt.savefig(filepath, bbox_inches='tight', dpi=100)
    plt.close()

    return filepath


def create_html_outlook_draft(attachment_path, subject, html_body, image_paths):
    """
    Interfaces with Outlook via the local COM object to create an HTML draft,
    embeds images via Content-ID, and guarantees temp file cleanup via try/finally.

    NOTE: _MAPI_PR_ATTACH_CONTENT_ID is a pure local MAPI property tag string
    used by the COM interface. No network request is made to any external address.
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail    = outlook.CreateItem(0)
        mail.Subject = subject

        if os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)

        for cid, img_path in image_paths.items():
            if img_path and os.path.exists(img_path):
                attachment = mail.Attachments.Add(img_path)
                # Sets the Content-ID on the attachment so Outlook renders it
                # inline via <img src="cid:...">. This is a local COM property
                # assignment — no external connection is made.
                attachment.PropertyAccessor.SetProperty(_MAPI_PR_ATTACH_CONTENT_ID, cid)

        mail.HTMLBody = html_body
        mail.Display()
        print(f"\n✉️  Email draft created successfully.")

    except Exception as e:
        logger.error("Failed to create Outlook draft:\n%s", traceback.format_exc())
        print(f"\n❌ Failed to create Outlook draft: {e}")

    finally:
        # ─── FIX 1: Always clean up temp chart files, even if Outlook fails ───
        for cid, img_path in image_paths.items():
            if img_path and os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception as cleanup_err:
                    print(f"⚠️ Could not delete temporary image {img_path}: {cleanup_err}")


def _build_email_html(global_stats, sheet_stats, total_issues, images_to_embed):
    """
    Builds the full HTML email body with a polished dark SOC report aesthetic.
    All dynamic values are plain numeric or pre-validated strings — no raw
    user input is interpolated into the HTML body.
    """
    run_time      = datetime.now().strftime('%d %B %Y  •  %H:%M:%S')
    total_scanned = sum(global_stats.values()) - global_stats['Pending-Maintenance']

    # Severity badge
    if total_issues == 0:
        badge_bg, badge_label = '#1a7a4a', 'ALL CLEAR'
    elif total_issues <= 10:
        badge_bg, badge_label = '#c87800', f'{total_issues} ISSUES'
    else:
        badge_bg, badge_label = '#c0392b', f'{total_issues} ISSUES'

    def metric_card(label, value, accent):
        return f"""
        <td style="padding:5px;">
          <div style="background:#1e2535;border-left:4px solid {accent};border-radius:6px;
                      padding:14px 16px;min-width:100px;text-align:center;">
            <div style="font-size:26px;font-weight:700;color:{accent};
                        letter-spacing:-1px;line-height:1;">{value}</div>
            <div style="font-size:10px;color:#7a86a0;margin-top:5px;
                        text-transform:uppercase;letter-spacing:0.6px;">{label}</div>
          </div>
        </td>"""

    overall_cards = (
        metric_card('Active',         global_stats['Active'] + global_stats['Inferred'], '#28a745') +
        metric_card('Inactive',       global_stats['Inactive'],                          '#dc3545') +
        metric_card('Not Found',      global_stats['Not Found'],                         '#6c757d') +
        metric_card('Disabled',       global_stats['Disabled'],                          '#17a2b8') +
        metric_card('Inferred',       global_stats['Inferred'],                          '#6f42c1') +
        metric_card('API Errors',     global_stats['API Errors'],                        '#fd7e14') +
        metric_card('Pending Maint.', global_stats['Pending-Maintenance'],               '#007bff')
    )

    def stat_row(label, value, color):
        return f"""
        <tr>
          <td style="padding:7px 14px;border-bottom:1px solid #22293d;
                     color:#8090b0;font-size:12px;">{label}</td>
          <td style="padding:7px 14px;border-bottom:1px solid #22293d;
                     font-weight:700;color:{color};font-size:13px;
                     text-align:right;">{value}</td>
        </tr>"""

    sheet_blocks = ""
    for sheet_name, counts in sheet_stats.items():
        cid        = f"chart_{sheet_name.replace(' ', '_')}"
        chart_html = (
            f'<img src="cid:{cid}" '
            f'style="display:block;max-width:300px;margin:12px auto 0;">'
        ) if cid in images_to_embed else ''

        # sheet_name here is a key we generated internally from Excel sheet names —
        # it is numeric/status data only, not raw user input from QRadar.
        sheet_blocks += f"""
        <div style="background:#161d2e;border:1px solid #252f47;border-radius:10px;
                    padding:20px;margin-bottom:18px;">
          <div style="font-size:14px;font-weight:700;color:#d0d8f0;
                      border-bottom:1px solid #252f47;padding-bottom:10px;
                      margin-bottom:12px;letter-spacing:0.2px;">
            📋 {sheet_name}
          </div>
          <table style="width:100%;border-collapse:collapse;">
            {stat_row('Active (incl. Inferred)', counts['Active'] + counts['Inferred'], '#28a745')}
            {stat_row('Inactive',                counts['Inactive'],                    '#dc3545')}
            {stat_row('Not Found',               counts['Not Found'],                   '#adb5bd')}
            {stat_row('Disabled',                counts['Disabled'],                    '#17a2b8')}
            {stat_row('Inferred',                counts['Inferred'],                    '#6f42c1')}
            {stat_row('API Errors',              counts['API Errors'],                  '#fd7e14')}
            {stat_row('Pending Maintenance',     counts['Pending-Maintenance'],         '#007bff')}
          </table>
          {chart_html}
        </div>"""

    overall_chart_html = ""
    if "overall_chart" in images_to_embed:
        overall_chart_html = (
            '<img src="cid:overall_chart" '
            'style="display:block;max-width:340px;margin:16px auto 4px;">'
        )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0d1117;
             font-family:'Segoe UI',Helvetica,Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#0d1117;padding:24px 0;">
<tr><td align="center">
<table width="660" cellpadding="0" cellspacing="0"
       style="max-width:660px;width:100%;">

  <!-- ═══ HEADER ═══ -->
  <tr>
    <td style="background:linear-gradient(135deg,#0c1628 0%,#172040 60%,#1a2a50 100%);
               border-radius:12px 12px 0 0;padding:28px 32px 24px;
               border-bottom:2px solid #2a3d6b;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <div style="font-size:10px;color:#3a5a9a;letter-spacing:3px;
                        text-transform:uppercase;margin-bottom:8px;">
              Automated SOC Intelligence Report
            </div>
            <div style="font-size:24px;font-weight:700;color:#e8ecf4;
                        letter-spacing:-0.5px;line-height:1.25;">
              QRadar Log Source<br>Status Report
            </div>
            <div style="margin-top:10px;font-size:11px;color:#4a6590;
                        letter-spacing:0.3px;">
              {run_time}
            </div>
          </td>
          <td align="right" valign="top">
            <div style="background:{badge_bg};color:#fff;font-size:12px;
                        font-weight:700;padding:9px 18px;border-radius:20px;
                        letter-spacing:0.8px;white-space:nowrap;
                        box-shadow:0 2px 8px rgba(0,0,0,0.4);">
              ⚠&nbsp; {badge_label}
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ═══ ACTION BANNER ═══ -->
  <tr>
    <td style="background:#1c0b0b;border-left:4px solid #c0392b;
               padding:13px 24px;">
      <span style="color:#e74c3c;font-size:13px;font-weight:700;">
        ACTION REQUIRED:&nbsp;
      </span>
      <span style="color:#b8b8c8;font-size:13px;">
        {total_issues} issues (Inactive&nbsp;+&nbsp;Not Found&nbsp;+&nbsp;API Errors)
        require immediate attention. The attached Excel contains
        <strong style="color:#e8ecf4;">only actionable items</strong>.
      </span>
    </td>
  </tr>

  <!-- ═══ OVERALL METRICS ═══ -->
  <tr>
    <td style="background:#131929;padding:22px 24px 18px;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:14px;">
        Overall Summary &nbsp;·&nbsp; {total_scanned} Assets Scanned
      </div>
      <table cellpadding="0" cellspacing="0">
        <tr>{overall_cards}</tr>
      </table>
    </td>
  </tr>

  <!-- ═══ OVERALL PIE CHART ═══ -->
  <tr>
    <td style="background:#131929;padding:4px 24px 22px;text-align:center;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:6px;">
        System Health Overview
      </div>
      {overall_chart_html}
    </td>
  </tr>

  <!-- ═══ DIVIDER ═══ -->
  <tr>
    <td style="padding:0;">
      <div style="height:1px;background:linear-gradient(90deg,
                  #0d1117,#2a3d6b 30%,#2a3d6b 70%,#0d1117);"></div>
    </td>
  </tr>

  <!-- ═══ PER-SHEET BREAKDOWN ═══ -->
  <tr>
    <td style="background:#131929;padding:22px 24px;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:16px;">
        Breakdown by Sheet
      </div>
      {sheet_blocks}
    </td>
  </tr>

  <!-- ═══ FOOTER ═══ -->
  <tr>
    <td style="background:#0a0f1a;border-radius:0 0 12px 12px;
               padding:16px 32px;text-align:center;
               border-top:1px solid #141c2e;">
      <div style="font-size:10px;color:#2a3a5a;letter-spacing:0.5px;">
        Automated Cyber Defense Reporting &nbsp;·&nbsp; Generated {run_time}
      </div>
    </td>
  </tr>

</table>
</td></tr>
</table>

</body>
</html>"""

    return html


def filter_and_email(processed_sheets_only, draft_path):
    """
    Calculates final SOC metrics across all categories, generates charts,
    saves the strictly ACTIONABLE inventory as an Excel attachment,
    and builds the polished HTML email draft.
    """
    report_frames   = {}
    sheet_stats     = {}
    images_to_embed = {}

    global_stats = {
        'Active': 0, 'Inactive': 0, 'Not Found': 0,
        'API Errors': 0, 'Disabled': 0, 'Inferred': 0,
        'Pending-Maintenance': 0
    }

    for name, df in processed_sheets_only.items():
        if 'status' not in df.columns:
            continue
        processed_df = df[df['status'].notna()].copy()
        if len(processed_df) == 0:
            continue

        active_count   = len(processed_df[(processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Active')])
        mask_inactive  = (processed_df['status'] == 'Found') & (processed_df['activity_status'].isin(['Inactive', 'No Activity']))
        inactive_count = mask_inactive.sum()
        disabled_count = len(processed_df[(processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Disabled')])
        inferred_count = len(processed_df[processed_df['status'] == 'Inferred'])
        mask_not_found = processed_df['status'] == 'Not Found'
        not_found_count = mask_not_found.sum()
        mask_error     = processed_df['status'].astype(str).str.startswith('API Error', na=False)
        error_count    = mask_error.sum()
        pending_count  = len(processed_df[processed_df['status'] == 'Pending-Maintenance'])

        sheet_counts = {
            'Active': active_count, 'Inactive': inactive_count, 'Not Found': not_found_count,
            'API Errors': error_count, 'Disabled': disabled_count, 'Inferred': inferred_count,
            'Pending-Maintenance': pending_count
        }

        sheet_stats[name] = sheet_counts
        for k in global_stats:
            global_stats[k] += sheet_counts[k]

        mask_report = mask_inactive | mask_not_found | mask_error

        if mask_report.any():
            sub = processed_df[mask_report].copy()

            # ─── FIX 4: Build remarks from actual days_since_last_event per row ───
            def build_inactive_remark(row):
                days = row.get('days_since_last_event')
                if pd.notna(days):
                    return f'Inactive - No events in last {int(days)} days'
                return 'Inactive - No events recorded'

            for idx in sub[mask_inactive.loc[sub.index]].index:
                sub.at[idx, 'remarks'] = build_inactive_remark(sub.loc[idx])

            report_frames[name] = sub

    if not report_frames:
        print("✅ No Actionable Issues detected; skipping email.")
        return

    try:
        with pd.ExcelWriter(draft_path, engine='openpyxl') as writer:
            for sheet_name, df in report_frames.items():
                if 'Is Older Expected' in df.columns:
                    df = df.drop(columns=['Is Older Expected'])
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except PermissionError:
        print(f"❌ ERROR: Could not save report to '{draft_path}'. Is the file open?")
        return

    # ── Generate charts via tempfile — no predictable names ──
    overall_path = generate_pie_chart(
        _chart_stats(global_stats), "Overall Inventory Status", prefix='qradar_overall'
    )
    if overall_path:
        images_to_embed["overall_chart"] = overall_path

    for name, counts in sheet_stats.items():
        cid        = f"chart_{name.replace(' ', '_')}"
        chart_path = generate_pie_chart(
            _chart_stats(counts), f"{name} Status", prefix=f'qradar_{name}'
        )
        if chart_path:
            images_to_embed[cid] = chart_path

    total_issues = global_stats['Inactive'] + global_stats['Not Found'] + global_stats['API Errors']

    html_body = _build_email_html(global_stats, sheet_stats, total_issues, images_to_embed)
    subject   = f"QRadar Action Report — {total_issues} Issues Require Attention"

    create_html_outlook_draft(draft_path, subject, html_body, images_to_embed)


def save_surgical_updates_to_excel(filepath, processed_dataframes):
    """
    SURGICAL SAVE & HIGHLIGHTER: Opens the original workbook with openpyxl and updates
    only the specific status columns. Evaluates conditional highlighting for expected/rogue Types.
    Uses the row_index_map built during process_sheet to guarantee correct row targeting.
    """
    try:
        wb = openpyxl.load_workbook(filepath)

        red_fill    = PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

        cols_to_update = [
            'status', 'qradar_id', 'enabled', 'last_seen', 'activity_status',
            'days_since_last_event', 'remarks', 'QRadar Actual Name', 'Log Source Type'
        ]

        for sheet_name, df in processed_dataframes.items():
            if sheet_name not in wb.sheetnames:
                continue

            ws         = wb[sheet_name]
            header_row = 1
            col_map    = {}

            for col_idx in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=header_row, column=col_idx).value
                if cell_val is not None:
                    col_map[str(cell_val).strip()] = col_idx

            for c in cols_to_update:
                if c not in col_map:
                    new_col_idx    = ws.max_column + 1
                    col_map[c]     = new_col_idx
                    ws.cell(row=header_row, column=new_col_idx).value = c

            # ─── FIX 2: Build reliable row map from the DataFrame's actual index ───
            row_index_map = {
                df_idx: excel_row
                for excel_row, df_idx in enumerate(df.index, start=2)
            }

            for idx, row in df.iterrows():
                excel_row         = row_index_map[idx]   # Safe — always correct row
                is_older_expected = row.get('Is Older Expected', False)

                for c in cols_to_update:
                    val = row[c]
                    if pd.isna(val):
                        val = ""

                    target_cell       = ws.cell(row=excel_row, column=col_map[c])
                    target_cell.value = val

                    if c == 'Log Source Type' and val not in ("", "N/A"):
                        is_match       = False
                        api_name_clean = str(val).lower()
                        for exp in EXPECTED_LS_TYPES:
                            exp_words = str(exp).lower().split()
                            if all(w in api_name_clean for w in exp_words):
                                is_match = True
                                break

                        if is_older_expected:
                            target_cell.fill = red_fill
                        elif not is_match:
                            target_cell.fill = yellow_fill

        wb.save(filepath)
        print("✅ Original Excel file updated surgically (Formulas, original data, and Conditional Formatting applied!).")

    except PermissionError:
        print(f"❌ ERROR: Permission Denied! The file '{filepath}' is OPEN. Close it and re-run to save changes.")
    except Exception as e:
        logger.error("Failed to surgically save Excel file:\n%s", traceback.format_exc())
        print(f"❌ Failed to surgically save Excel file: {e}")


def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker (Multi-threaded & Surgical)...")

    if GROUP_COLUMN:
        print(f"🏷️  Group Threshold Feature: ENABLED  "
              f"(column: '{GROUP_COLUMN}', {len(GROUP_THRESHOLDS)} custom threshold(s) defined)")
    else:
        print(f"🏷️  Group Threshold Feature: DISABLED  "
              f"(global threshold: {ACTIVITY_THRESHOLD_DAYS}d for all rows)")

    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        return

    # ── Must be fully populated before any worker threads are spawned ──
    fetch_log_source_types(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD)

    print(f"\n📖 Reading Excel file: {INPUT_EXCEL_PATH}")
    try:
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
    except Exception as e:
        logger.error("Failed to read Excel file:\n%s", traceback.format_exc())
        print(f"❌ Failed to read Excel file: {e}")
        return

    # Case-insensitive 'all' guard
    if [s.lower() for s in SHEETS_TO_PROCESS] == ['all']:
        to_process = list(all_sheets.keys())
    else:
        to_process = SHEETS_TO_PROCESS

    # ─── FIX 3: Warn on missing sheet names immediately ───
    for sheet in to_process:
        if sheet not in all_sheets:
            print(f"⚠️  WARNING: Sheet '{sheet}' not found in workbook.")
            print(f"   Available sheets: {list(all_sheets.keys())}")

    for sheet in to_process:
        if sheet in all_sheets:
            all_sheets[sheet] = process_sheet(
                all_sheets[sheet], sheet, QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD,
                LOGSOURCE_COLUMN, IP_COLUMN, IN_QRADAR_COLUMN
            )

    print(f"\n💾 Saving updates to original Excel...")
    processed_sheets_only = {k: v for k, v in all_sheets.items() if k in to_process}

    save_surgical_updates_to_excel(INPUT_EXCEL_PATH, processed_sheets_only)
    filter_and_email(processed_sheets_only, DRAFT_OUTPUT_PATH)

    print(f"\n✅ Completed!")


if __name__ == '__main__':
    main()
