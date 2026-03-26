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
IN_QRADAR_COLUMN    = 'In Qradar?'

QRADAR_HOST         = 'https://your-qradar-host'
QRADAR_USERNAME     = 'your-username'
QRADAR_PASSWORD     = 'your-password'
VERIFY_SSL          = False

DRAFT_OUTPUT_PATH        = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), 'inactive_and_errors.xlsx')
ACTIVITY_THRESHOLD_DAYS  = 7
REQUEST_TIMEOUT          = 30
MAX_WORKERS              = 10

EXPECTED_LS_TYPES = ['Microsoft Security', 'Linux OS']

# ─── GROUP THRESHOLD CONFIGURATION ────────────────────────────────────────────
GROUP_COLUMN     = None
GROUP_THRESHOLDS = {}

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

MIN_TIMESTAMP = 0
MAX_TIMESTAMP = 2147483647

LOG_SOURCE_TYPES_CACHE = {}

_MAPI_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"


# ─── GROUP THRESHOLD RESOLVER ──────────────────────────────────────────────────
def resolve_threshold(group_name=None):
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
    d = dict(stats_dict)
    d['Active'] = d.get('Active', 0) + d.pop('Inferred', 0)
    return d


def test_qradar_connection(qradar_host, username, password):
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
        print("❌ Connection timed out.")
        return False
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Connection error: {e}")
        return False
    except Exception as e:
        logger.error("Unexpected error during connection test:\n%s", traceback.format_exc())
        print(f"❌ Connection failed: {e}")
        return False


def fetch_log_source_types(qradar_host, username, password):
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
            for t in resp.json():
                ls_id   = t.get('id')
                ls_name = t.get('name')
                if ls_id is not None and ls_name is not None:
                    LOG_SOURCE_TYPES_CACHE[ls_id] = ls_name
            print(f"✅ Cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
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
    effective_threshold = threshold if threshold is not None else ACTIVITY_THRESHOLD_DAYS
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    try:
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)
        if timestamp_ms > 4102444800:
            timestamp_seconds = timestamp_ms / 1000.0
        else:
            timestamp_seconds = timestamp_ms
        if timestamp_seconds <= MIN_TIMESTAMP or timestamp_seconds > MAX_TIMESTAMP:
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None
        last_event_datetime   = datetime.fromtimestamp(timestamp_seconds)
        last_seen             = last_event_datetime.strftime('%Y-%m-%d %H:%M:%S')
        time_diff             = datetime.now() - last_event_datetime
        days_since_last_event = time_diff.days
        threshold_time        = datetime.now() - timedelta(days=effective_threshold)
        activity_status       = 'Active' if last_event_datetime > threshold_time else 'Inactive'
        return last_seen, activity_status, days_since_last_event
    except Exception:
        logger.error("Timestamp conversion error:\n%s", traceback.format_exc())
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False, threshold=None):
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

        valid_sources = []
        if is_ip:
            for src in ls_data:
                params    = src.get('protocol_parameters', [])
                in_params = any(p.get('value') == clean_identifier for p in params)
                in_name   = clean_identifier.lower() in str(src.get('name', '')).lower()
                if in_params or in_name:
                    valid_sources.append(src)
        else:
            valid_sources = ls_data

        if not valid_sources:
            return {'status': 'Not Found', **_empty_details()}

        expected_sources   = []
        unexpected_sources = []

        for src in valid_sources:
            type_name      = LOG_SOURCE_TYPES_CACHE.get(src.get('type_id'), "")
            api_name_clean = str(type_name).lower()
            is_match = any(
                all(w in api_name_clean for w in str(exp).lower().split())
                for exp in EXPECTED_LS_TYPES
            )
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

        found_source      = None
        is_older_expected = False
        absolute_max_time = max([s.get('last_event_time') or 0 for s in valid_sources])

        if expected_sources:
            found_source = get_best_source(expected_sources)
            if (found_source.get('last_event_time') or 0) < absolute_max_time:
                is_older_expected = True
        else:
            found_source = get_best_source(unexpected_sources)

        ls_id        = found_source.get('id')
        ls_name      = found_source.get('name', identifier)
        type_id      = found_source.get('type_id')
        ls_type_name = LOG_SOURCE_TYPES_CACHE.get(type_id, f"Unknown Type ID: {type_id}")

        last_event_time_ms = found_source.get('last_event_time')
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
    if name_val and str(name_val).lower() in ['nan', 'none', '', 'null']: name_val = None
    if ip_val   and str(ip_val).lower()   in ['nan', 'none', '', 'null']: ip_val   = None

    effective_threshold = resolve_threshold(group_val)
    details             = None
    search_method       = "None"

    if name_val:
        details = get_log_source_details(
            qradar_host, username, password, name_val,
            is_ip=False, threshold=effective_threshold
        )
        if details['status'] == 'Found':
            search_method = "Name"

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


def process_sheet(df, sheet_name, qradar_host, username, password,
                  logsource_column, ip_column, in_qradar_col):
    print(f"\n{'='*60}")
    print(f"📋 Processing Sheet: {sheet_name}")
    print(f"{'='*60}")

    if not df.empty:
        df.columns = df.columns.str.strip()

    required_columns = [in_qradar_col, logsource_column, ip_column]
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ Missing columns {missing} in sheet '{sheet_name}'. Skipping.")
        return df

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

    row_index_map = {
        df_idx: excel_row
        for excel_row, df_idx in enumerate(df.index, start=2)
    }

    in_qradar_series = df[in_qradar_col].astype(str).str.lower()
    process_mask     = in_qradar_series.str.contains("yes",                 na=False)
    pending_mask     = in_qradar_series.str.contains("pending-maintenance", na=False)

    rows_to_process = df[process_mask]
    total_rows      = len(df)
    target_count    = len(rows_to_process)
    pending_count   = pending_mask.sum()

    group_feature_active = bool(GROUP_COLUMN and GROUP_COLUMN in df.columns)
    if GROUP_COLUMN and not group_feature_active:
        print(f"⚠️  GROUP_COLUMN '{GROUP_COLUMN}' not found — using global threshold.")

    print(f"📊 Total: {total_rows} | To Scan: {target_count} | Pending: {pending_count}")

    skipped_mask = ~(process_mask | pending_mask)
    df.loc[skipped_mask, 'remarks'] = "Skipped (Not Yes or Pending)"

    if pending_count > 0:
        df.loc[pending_mask, 'status']          = 'Pending-Maintenance'
        df.loc[pending_mask, 'activity_status'] = 'Pending-Maintenance'
        df.loc[pending_mask, 'remarks']         = 'Pending Maintenance (Not Scanned)'
        df.loc[pending_mask, 'last_seen']       = 'N/A'

    if target_count == 0:
        return df

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
            try:
                idx, name_val, details, search_method = future.result()
            except Exception as worker_exc:
                original_idx = futures[future]
                logger.error("Worker crashed for row %s:\n%s", original_idx, traceback.format_exc())
                print(f"\n⚠️  [{processed_count}/{target_count}] Worker crashed: {worker_exc}")
                df.at[original_idx, 'status']          = 'Worker Error'
                df.at[original_idx, 'remarks']         = f'Thread exception: {str(worker_exc)[:80]}'
                df.at[original_idx, 'activity_status'] = 'Error'
                df.at[original_idx, 'last_seen']       = 'N/A'
                continue

            print(f"\n🔹 [{processed_count}/{target_count}] {name_val or 'Unknown'} -> {details['status']}")

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
                    print(f"      🚨 Bypassed newer unexpected log source!")

                if details['enabled'] == 'No':
                    df.at[idx, 'status']          = 'Found'
                    df.at[idx, 'remarks']         = "Disabled on QRadar"
                    df.at[idx, 'activity_status'] = "Disabled"
                    print(f"      ⚪ Status: Disabled")
                else:
                    df.at[idx, 'status']          = 'Found'
                    df.at[idx, 'remarks']         = base_remark
                    df.at[idx, 'activity_status'] = details['activity_status']
                    print(f"      📊 Activity: {details['activity_status']}")
                    print(f"      📅 Last Event: {details['last_seen']} ({details['days_since_last_event']} days ago)")
            else:
                status_val = details['status']
                remark_val = f"❌ {status_val}"
                act_val    = "Error" if "Error" in status_val or "Timeout" in status_val else "Not Found"

                if name_val and "AP" in name_val:
                    status_val = "Inferred"
                    remark_val = "Under WLC (Inferred)"
                    act_val    = "Inferred"
                    print("      ℹ️  Inferred as WLC")
                elif name_val and "FW" in name_val:
                    status_val = "Inferred"
                    remark_val = "Under Forti (Inferred)"
                    act_val    = "Inferred"
                    print("      ℹ️  Inferred as FortiGate")
                else:
                    print(f"      ❌ Result: {status_val.upper()}")

                df.at[idx, 'status']                = status_val
                df.at[idx, 'remarks']               = remark_val
                df.at[idx, 'activity_status']       = act_val
                df.at[idx, 'last_seen']             = "N/A"
                df.at[idx, 'days_since_last_event'] = None

    return df


def generate_pie_chart(data_dict, title, prefix='qradar_chart'):
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
        startangle=140, textprops={'fontsize': 9},
        wedgeprops={'edgecolor': 'white'}
    )
    ax.axis('equal')
    plt.title(title, pad=15, fontsize=11, fontweight='bold')

    tmp      = tempfile.NamedTemporaryFile(suffix='.png', delete=False, prefix=f'{prefix}_')
    filepath = tmp.name
    tmp.close()
    plt.savefig(filepath, bbox_inches='tight', dpi=100)
    plt.close()
    return filepath


def create_html_outlook_draft(attachment_path, subject, html_body, image_paths):
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail    = outlook.CreateItem(0)
        mail.Subject = subject

        if os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)

        for cid, img_path in image_paths.items():
            if img_path and os.path.exists(img_path):
                attachment = mail.Attachments.Add(img_path)
                attachment.PropertyAccessor.SetProperty(_MAPI_PR_ATTACH_CONTENT_ID, cid)

        mail.HTMLBody = html_body
        mail.Display()
        print(f"\n✉️  Email draft created successfully.")

    except Exception as e:
        logger.error("Failed to create Outlook draft:\n%s", traceback.format_exc())
        print(f"\n❌ Failed to create Outlook draft: {e}")

    finally:
        for cid, img_path in image_paths.items():
            if img_path and os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception as cleanup_err:
                    print(f"⚠️ Could not delete temp image {img_path}: {cleanup_err}")


# ─── STATUS METADATA ───────────────────────────────────────────────────────────
_STATUS_META = {
    'Inactive':   {'bg': '#2a0d0d', 'accent': '#dc3545', 'badge_bg': '#dc3545', 'label': 'INACTIVE',   'icon': '🔴'},
    'No Activity':{'bg': '#2a0d0d', 'accent': '#dc3545', 'badge_bg': '#dc3545', 'label': 'NO ACTIVITY','icon': '🔴'},
    'Not Found':  {'bg': '#1a1c22', 'accent': '#6c757d', 'badge_bg': '#6c757d', 'label': 'NOT FOUND',  'icon': '⚫'},
    'Error':      {'bg': '#1f1608', 'accent': '#fd7e14', 'badge_bg': '#fd7e14', 'label': 'API ERROR',  'icon': '🟠'},
}

def _get_status_meta(activity_status):
    s = str(activity_status).strip()
    for key, meta in _STATUS_META.items():
        if key.lower() in s.lower():
            return meta
    return {'bg': '#1a1c22', 'accent': '#6c757d', 'badge_bg': '#6c757d',
            'label': s.upper()[:12], 'icon': '⚫'}


def _build_actionable_table(report_df, logsource_col, ip_col):
    """
    Builds a compact inline HTML table for one sheet's actionable rows.
    Shows: hostname, IP, QRadar ID, status, last event time.
    Rows are colour-coded by severity tier.
    """
    if report_df is None or len(report_df) == 0:
        return '<p style="color:#3a4a6a;font-size:12px;font-style:italic;padding:8px 0;">No actionable items.</p>'

    rows_html = ''
    for i, (_, row) in enumerate(report_df.iterrows()):
        hostname      = str(row.get(logsource_col, 'N/A') or 'N/A')
        ip_val        = str(row.get(ip_col,        'N/A') or 'N/A')
        qradar_id     = str(row.get('qradar_id',   'N/A') or 'N/A')
        last_seen     = str(row.get('last_seen',   'N/A') or 'N/A')
        act_status    = str(row.get('activity_status', 'Unknown') or 'Unknown')

        meta    = _get_status_meta(act_status)
        row_bg  = meta['bg'] if i % 2 == 0 else '#131929'

        days = row.get('days_since_last_event')
        if pd.notna(days) and days is not None:
            try:
                days_str = (f"<span style='color:#888;font-size:10px;display:block;'>"
                            f"{'Today' if int(days) == 0 else str(int(days)) + 'd ago'}</span>")
            except Exception:
                days_str = ''
        else:
            days_str = ''

        # Truncate long hostnames gracefully
        hostname_display = hostname if len(hostname) <= 38 else hostname[:36] + '…'

        rows_html += f"""
        <tr style="background:{row_bg};">
          <td style="padding:8px 12px;border-bottom:1px solid #1e2840;
                     font-size:11px;color:#c8d0e8;max-width:220px;">
            <span title="{hostname}">{hostname_display}</span>
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #1e2840;
                     font-size:11px;color:#8090b0;text-align:center;
                     white-space:nowrap;">{ip_val}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #1e2840;
                     font-size:11px;color:#6a80b0;text-align:center;
                     white-space:nowrap;">{qradar_id}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #1e2840;
                     font-size:11px;color:#8090b0;text-align:center;">
            {last_seen}{days_str}
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #1e2840;
                     text-align:center;">
            <span style="background:{meta['badge_bg']};color:#fff;font-size:9px;
                         font-weight:700;padding:3px 9px;border-radius:8px;
                         letter-spacing:0.3px;white-space:nowrap;">
              {meta['icon']}&nbsp;{meta['label']}
            </span>
          </td>
        </tr>"""

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border-collapse:collapse;margin-top:12px;">
      <tr style="background:#1a2240;">
        <th style="padding:7px 12px;text-align:left;font-size:10px;color:#6a80b0;
                   font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                   border-bottom:2px solid #2a3d6b;">Hostname / Log Source</th>
        <th style="padding:7px 12px;text-align:center;font-size:10px;color:#6a80b0;
                   font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                   border-bottom:2px solid #2a3d6b;">IP Address</th>
        <th style="padding:7px 12px;text-align:center;font-size:10px;color:#6a80b0;
                   font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                   border-bottom:2px solid #2a3d6b;">QRadar ID</th>
        <th style="padding:7px 12px;text-align:center;font-size:10px;color:#6a80b0;
                   font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                   border-bottom:2px solid #2a3d6b;">Last Event</th>
        <th style="padding:7px 12px;text-align:center;font-size:10px;color:#6a80b0;
                   font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                   border-bottom:2px solid #2a3d6b;">Status</th>
      </tr>
      {rows_html}
    </table>"""


def _build_email_html(global_stats, sheet_stats, total_issues,
                      images_to_embed, report_frames,
                      logsource_col, ip_col):
    """
    Builds the full HTML email body — dark SOC aesthetic matching the
    onboarding tracker. Each sheet section now includes:
      • stat counters + mini pie chart (unchanged)
      • inline actionable table: hostname · IP · QRadar ID · last event · status
    """
    run_time      = datetime.now().strftime('%d %B %Y  •  %H:%M:%S')
    total_scanned = sum(global_stats.values()) - global_stats['Pending-Maintenance']

    # ── Header severity badge ──
    if total_issues == 0:
        badge_bg, badge_label = '#1a7a4a', '✔&nbsp; ALL CLEAR'
    elif total_issues <= 10:
        badge_bg, badge_label = '#c87800', f'⚠&nbsp; {total_issues} ISSUES'
    else:
        badge_bg, badge_label = '#c0392b', f'⚠&nbsp; {total_issues} ISSUES'

    # ── Overall metric cards ──
    def metric_card(label, value, accent, sub=''):
        sub_html = (f'<div style="font-size:9px;color:#4a5a7a;margin-top:3px;">'
                    f'{sub}</div>') if sub else ''
        return f"""
        <td style="padding:5px;">
          <div style="background:#1e2535;border-left:4px solid {accent};
                      border-radius:6px;padding:14px 16px;min-width:100px;
                      text-align:center;">
            <div style="font-size:26px;font-weight:700;color:{accent};
                        letter-spacing:-1px;line-height:1;">{value}</div>
            <div style="font-size:10px;color:#7a86a0;margin-top:5px;
                        text-transform:uppercase;letter-spacing:0.6px;">{label}</div>
            {sub_html}
          </div>
        </td>"""

    overall_cards = (
        metric_card('Active',         global_stats['Active'] + global_stats['Inferred'],
                    '#28a745', 'incl. inferred') +
        metric_card('Inactive',       global_stats['Inactive'],        '#dc3545') +
        metric_card('Not Found',      global_stats['Not Found'],       '#6c757d') +
        metric_card('Disabled',       global_stats['Disabled'],        '#17a2b8') +
        metric_card('Inferred',       global_stats['Inferred'],        '#6f42c1') +
        metric_card('API Errors',     global_stats['API Errors'],      '#fd7e14') +
        metric_card('Pending Maint.', global_stats['Pending-Maintenance'], '#007bff')
    )

    # ── Per-sheet blocks ──
    def stat_pill(label, value, color):
        if value == 0:
            return ''
        return f"""
        <span style="display:inline-block;background:{color}22;border:1px solid {color}55;
                     color:{color};font-size:10px;font-weight:700;padding:3px 10px;
                     border-radius:10px;margin:3px 4px 3px 0;letter-spacing:0.3px;">
          {label}&nbsp;{value}
        </span>"""

    sheet_blocks = ''
    for sheet_name, counts in sheet_stats.items():
        cid         = f"chart_{sheet_name.replace(' ', '_')}"
        sheet_total = sum(counts.values()) - counts['Pending-Maintenance']

        # Pill summary row
        pills = (
            stat_pill('Active',     counts['Active'] + counts['Inferred'], '#28a745') +
            stat_pill('Inactive',   counts['Inactive'],                    '#dc3545') +
            stat_pill('Not Found',  counts['Not Found'],                   '#6c757d') +
            stat_pill('Disabled',   counts['Disabled'],                    '#17a2b8') +
            stat_pill('API Errors', counts['API Errors'],                  '#fd7e14') +
            stat_pill('Pending',    counts['Pending-Maintenance'],         '#007bff')
        )

        chart_html = (
            f'<img src="cid:{cid}" '
            f'style="display:block;max-width:280px;margin:14px auto 0;">'
        ) if cid in images_to_embed else ''

        # Inline actionable table for this sheet
        report_df       = report_frames.get(sheet_name)
        actionable_html = _build_actionable_table(report_df, logsource_col, ip_col)

        issue_count  = counts['Inactive'] + counts['Not Found'] + counts['API Errors']
        section_accent = '#dc3545' if issue_count > 0 else '#28a745'

        sheet_blocks += f"""
        <div style="background:#161d2e;border:1px solid #252f47;border-radius:10px;
                    margin-bottom:22px;overflow:hidden;">

          <!-- Sheet header -->
          <div style="background:{section_accent}18;border-left:4px solid {section_accent};
                      padding:13px 18px;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td>
                  <span style="font-size:14px;font-weight:700;color:#d0d8f0;">
                    📋&nbsp;{sheet_name}
                  </span>
                  <span style="font-size:10px;color:#6a80b0;margin-left:10px;">
                    {sheet_total} scanned
                  </span>
                </td>
                <td align="right">
                  {'<span style="background:#dc354522;color:#dc3545;font-size:10px;font-weight:700;padding:4px 12px;border-radius:10px;border:1px solid #dc354555;">' + str(issue_count) + ' ISSUE' + ('S' if issue_count != 1 else '') + '</span>' if issue_count > 0 else '<span style="background:#28a74522;color:#28a745;font-size:10px;font-weight:700;padding:4px 12px;border-radius:10px;border:1px solid #28a74555;">ALL CLEAR</span>'}
                </td>
              </tr>
            </table>
          </div>

          <!-- Stat pills + mini chart -->
          <div style="padding:14px 18px 10px;">
            <div style="margin-bottom:10px;">{pills}</div>
            {chart_html}
          </div>

          <!-- Actionable items table -->
          <div style="border-top:1px solid #252f47;padding:14px 18px 16px;">
            <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                        letter-spacing:2px;margin-bottom:2px;">
              Actionable Items
            </div>
            <div style="font-size:10px;color:#3a4a6a;margin-bottom:8px;">
              Inactive · Not Found · API Errors only — full data in attached Excel.
            </div>
            {actionable_html}
          </div>
        </div>"""

    # ── Overall pie chart ──
    overall_chart_html = (
        '<img src="cid:overall_chart" '
        'style="display:block;max-width:340px;margin:16px auto 4px;">'
    ) if "overall_chart" in images_to_embed else ''

    return f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0d1117;
             font-family:'Segoe UI',Helvetica,Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#0d1117;padding:24px 0;">
<tr><td align="center">
<table width="720" cellpadding="0" cellspacing="0"
       style="max-width:720px;width:100%;">

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
            <div style="margin-top:10px;font-size:11px;color:#4a6590;">
              {run_time}
            </div>
          </td>
          <td align="right" valign="top">
            <div style="background:{badge_bg};color:#fff;font-size:12px;
                        font-weight:700;padding:9px 18px;border-radius:20px;
                        letter-spacing:0.8px;white-space:nowrap;
                        box-shadow:0 2px 8px rgba(0,0,0,0.4);">
              {badge_label}
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
        {total_issues} issues (Inactive + Not Found + API Errors) require
        attention. The attached Excel contains
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
                  letter-spacing:2px;margin-bottom:6px;">System Health Overview</div>
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
                  letter-spacing:2px;margin-bottom:4px;">Breakdown by Sheet</div>
      <div style="font-size:10px;color:#3a4a6a;margin-bottom:18px;">
        Each section shows a stat summary, health chart, and inline table
        of assets requiring attention.
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


def filter_and_email(processed_sheets_only, draft_path,
                     logsource_col=LOGSOURCE_COLUMN, ip_col=IP_COLUMN):
    """
    Calculates final SOC metrics, generates charts, saves the actionable
    Excel attachment, and builds the polished HTML email draft.

    report_frames is now passed into _build_email_html so the inline
    per-sheet actionable tables can be rendered in the email body.
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

        active_count    = len(processed_df[
            (processed_df['status'] == 'Found') &
            (processed_df['activity_status'] == 'Active')
        ])
        mask_inactive   = (
            (processed_df['status'] == 'Found') &
            (processed_df['activity_status'].isin(['Inactive', 'No Activity']))
        )
        inactive_count  = mask_inactive.sum()
        disabled_count  = len(processed_df[
            (processed_df['status'] == 'Found') &
            (processed_df['activity_status'] == 'Disabled')
        ])
        inferred_count  = len(processed_df[processed_df['status'] == 'Inferred'])
        mask_not_found  = processed_df['status'] == 'Not Found'
        not_found_count = mask_not_found.sum()
        mask_error      = processed_df['status'].astype(str).str.startswith('API Error', na=False)
        error_count     = mask_error.sum()
        pending_count   = len(processed_df[processed_df['status'] == 'Pending-Maintenance'])

        sheet_counts = {
            'Active': active_count, 'Inactive': inactive_count,
            'Not Found': not_found_count, 'API Errors': error_count,
            'Disabled': disabled_count, 'Inferred': inferred_count,
            'Pending-Maintenance': pending_count
        }

        sheet_stats[name] = sheet_counts
        for k in global_stats:
            global_stats[k] += sheet_counts[k]

        mask_report = mask_inactive | mask_not_found | mask_error

        if mask_report.any():
            sub = processed_df[mask_report].copy()

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

    # ── Save actionable Excel ──
    try:
        with pd.ExcelWriter(draft_path, engine='openpyxl') as writer:
            for sheet_name, df in report_frames.items():
                out_df = df.copy()
                if 'Is Older Expected' in out_df.columns:
                    out_df = out_df.drop(columns=['Is Older Expected'])
                out_df.to_excel(writer, sheet_name=sheet_name, index=False)
    except PermissionError:
        print(f"❌ Could not save report to '{draft_path}'. Is the file open?")
        return

    # ── Generate charts ──
    overall_path = generate_pie_chart(
        _chart_stats(global_stats), "Overall Inventory Status",
        prefix='qradar_overall'
    )
    if overall_path:
        images_to_embed["overall_chart"] = overall_path

    for name, counts in sheet_stats.items():
        cid        = f"chart_{name.replace(' ', '_')}"
        chart_path = generate_pie_chart(
            _chart_stats(counts), f"{name} Status",
            prefix=f'qradar_{name}'
        )
        if chart_path:
            images_to_embed[cid] = chart_path

    total_issues = (global_stats['Inactive'] +
                    global_stats['Not Found'] +
                    global_stats['API Errors'])

    html_body = _build_email_html(
        global_stats, sheet_stats, total_issues,
        images_to_embed, report_frames,
        logsource_col, ip_col
    )

    subject = f"QRadar Action Report — {total_issues} Issues Require Attention"
    create_html_outlook_draft(draft_path, subject, html_body, images_to_embed)


def save_surgical_updates_to_excel(filepath, processed_dataframes):
    """
    Opens the original workbook with openpyxl and updates only the specific
    status columns. Applies conditional highlighting for expected/rogue types.
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

            row_index_map = {
                df_idx: excel_row
                for excel_row, df_idx in enumerate(df.index, start=2)
            }

            for idx, row in df.iterrows():
                excel_row         = row_index_map[idx]
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
        print("✅ Original Excel file updated surgically.")

    except PermissionError:
        print(f"❌ Permission Denied! '{filepath}' is open. Close it and re-run.")
    except Exception as e:
        logger.error("Failed to surgically save Excel:\n%s", traceback.format_exc())
        print(f"❌ Failed to surgically save Excel: {e}")


def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker (Multi-threaded & Surgical)...")

    if GROUP_COLUMN:
        print(f"🏷️  Group Threshold: ENABLED  "
              f"(column: '{GROUP_COLUMN}', {len(GROUP_THRESHOLDS)} threshold(s))")
    else:
        print(f"🏷️  Group Threshold: DISABLED  (global: {ACTIVITY_THRESHOLD_DAYS}d)")

    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        return

    fetch_log_source_types(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD)

    print(f"\n📖 Reading Excel: {INPUT_EXCEL_PATH}")
    try:
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
    except Exception as e:
        logger.error("Failed to read Excel:\n%s", traceback.format_exc())
        print(f"❌ Failed to read Excel: {e}")
        return

    if [s.lower() for s in SHEETS_TO_PROCESS] == ['all']:
        to_process = list(all_sheets.keys())
    else:
        to_process = SHEETS_TO_PROCESS

    for sheet in to_process:
        if sheet not in all_sheets:
            print(f"⚠️  Sheet '{sheet}' not found. Available: {list(all_sheets.keys())}")

    for sheet in to_process:
        if sheet in all_sheets:
            all_sheets[sheet] = process_sheet(
                all_sheets[sheet], sheet,
                QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD,
                LOGSOURCE_COLUMN, IP_COLUMN, IN_QRADAR_COLUMN
            )

    print(f"\n💾 Saving updates to original Excel...")
    processed_sheets_only = {k: v for k, v in all_sheets.items() if k in to_process}

    save_surgical_updates_to_excel(INPUT_EXCEL_PATH, processed_sheets_only)
    filter_and_email(
        processed_sheets_only, DRAFT_OUTPUT_PATH,
        logsource_col=LOGSOURCE_COLUMN, ip_col=IP_COLUMN
    )

    print(f"\n✅ Completed!")


if __name__ == '__main__':
    main()
