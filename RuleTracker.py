import os
import shutil
import time
import tempfile
import traceback
import logging
from datetime import datetime, timedelta

import requests
import urllib3
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import win32com.client

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


QRADAR_HOST     = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL      = False

TRACKER_VALIDATION_WINDOW_DAYS = 7
FULL_HISTORY_LOOKBACK_DAYS     = None
RULE_DEAD_THRESHOLD            = 0

TRACKER_EXCEL_PATH                 = ''
TRACKER_HEADER_ROW                 = 1
AUTO_UPDATE_TRACKER_TESTED         = True
AUTO_CLEAR_INVESTIGATION_ON_TESTED = True
TRACKER_BACKUP_BEFORE_WRITE        = True

OUTPUT_DIR = r'C:\path\to\your\output'

REPORT_TITLE = 'Rules Under Investigation & Dead'

EMAIL_TABLE_ROW_CAP = 25
EMAIL_HIGH_IMPORTANCE_DEAD_THRESHOLD = 15
DEAD_RULE_STALE_AGE_DAYS = 365

REQUEST_TIMEOUT    = 30
MAX_RETRIES        = 3
RETRY_DELAY_BASE   = 1.5
API_PAGE_SIZE      = 9999
MAX_PAGES          = 150
QRADAR_API_VERSION = '14.0'

OUTPUT_EXCEL = os.path.join(OUTPUT_DIR, 'qradar_rules_under_investigation_and_dead.xlsx')

_MAPI_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

_RULE_NAME_COL_ALIASES     = ['alert rule name', 'rule name', 'name']
_INVESTIGATION_COL_ALIASES = ['under investigation', 'investigation', 'investigating']
_TESTED_COL_ALIASES        = ['tested', 'testing completed', 'test completed', 'testing']
_TRUE_STRINGS = {'true', 'yes', 'y', '1', 'x', '✓', 'done', 'complete', 'completed'}


def _api_get_page(path, range_start, range_end, params=None, label='request'):
    url     = f"{QRADAR_HOST.rstrip('/')}{path}"
    headers = {
        'Accept':  'application/json',
        'Version': QRADAR_API_VERSION,
        'Range':   f'items={range_start}-{range_end}',
    }
    last_err = None

    for attempt in range(MAX_RETRIES):
        if attempt > 0:
            wait = RETRY_DELAY_BASE * (2 ** (attempt - 1))
            logger.warning("Retry %d/%d for '%s' — waiting %.1fs after %s",
                           attempt, MAX_RETRIES - 1, label, wait,
                           type(last_err).__name__)
            time.sleep(wait)
        try:
            resp = requests.get(
                url,
                params=params,
                auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
                verify=VERIFY_SSL,
                timeout=REQUEST_TIMEOUT,
                headers=headers,
            )
            if resp.status_code in (200, 206):
                total = None
                cr = resp.headers.get('Content-Range', '')
                if cr:
                    try:
                        total = int(cr.split('/')[-1].strip())
                    except Exception:
                        pass
                return resp.json(), total
            elif resp.status_code == 401:
                raise RuntimeError("Authentication failed — check QRADAR_USERNAME / PASSWORD")
            else:
                last_err = RuntimeError(f"HTTP {resp.status_code}")
                logger.warning("HTTP %d for '%s' (attempt %d/%d)",
                               resp.status_code, label, attempt + 1, MAX_RETRIES)
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_err = exc
            logger.warning("%s on attempt %d for '%s'",
                           type(exc).__name__, attempt + 1, label)
        except RuntimeError:
            raise
        except Exception:
            logger.error("Non-retriable error for '%s':\n%s", label, traceback.format_exc())
            raise

    raise RuntimeError(f"All {MAX_RETRIES} attempts failed for '{label}': {last_err}")


def _api_get(path, params=None, label='request'):
    items, total = _api_get_page(path, 0, API_PAGE_SIZE, params=params, label=label)
    if total is not None and total > API_PAGE_SIZE + 1:
        logger.warning(
            "Pagination cap hit for '%s': %d items on server, only %d fetched. "
            "Use _api_get_all for this endpoint.", label, total, API_PAGE_SIZE + 1
        )
    return items


def _api_get_all(path, params=None, label='request', page_size=None, max_pages=MAX_PAGES):
    page_size = page_size or (API_PAGE_SIZE + 1)
    all_items = []
    start = 0
    for page_num in range(max_pages):
        end = start + page_size - 1
        items, total = _api_get_page(
            path, start, end, params=params,
            label=f'{label} (items {start}-{end})'
        )
        if not items:
            break
        all_items.extend(items)
        if total is not None and len(all_items) >= total:
            break
        if len(items) < page_size:
            break
        start += page_size
    else:
        logger.warning(
            "Hit max_pages=%d for '%s' — data may be incomplete. Raise MAX_PAGES "
            "if your deployment has more history than that.", max_pages, label
        )
    return all_items


def validate_config():
    problems = []

    if QRADAR_HOST.rstrip('/') in ('https://your-qradar-host', ''):
        problems.append("QRADAR_HOST is still the placeholder — set it to your real console URL.")
    if QRADAR_USERNAME in ('your-username', '') or QRADAR_PASSWORD in ('your-password', ''):
        problems.append("QRADAR_USERNAME / QRADAR_PASSWORD look like the placeholders.")
    if OUTPUT_DIR == r'C:\path\to\your\output':
        problems.append("OUTPUT_DIR is still the placeholder path.")

    if TRACKER_VALIDATION_WINDOW_DAYS <= 0:
        problems.append("TRACKER_VALIDATION_WINDOW_DAYS must be a positive number of days.")

    if FULL_HISTORY_LOOKBACK_DAYS is not None:
        if FULL_HISTORY_LOOKBACK_DAYS <= 0:
            problems.append("FULL_HISTORY_LOOKBACK_DAYS must be None (unbounded) or a positive number of days.")
        elif FULL_HISTORY_LOOKBACK_DAYS <= TRACKER_VALIDATION_WINDOW_DAYS:
            problems.append(
                f"FULL_HISTORY_LOOKBACK_DAYS ({FULL_HISTORY_LOOKBACK_DAYS}) must be greater than "
                f"TRACKER_VALIDATION_WINDOW_DAYS ({TRACKER_VALIDATION_WINDOW_DAYS})."
            )

    if TRACKER_EXCEL_PATH and not os.path.exists(TRACKER_EXCEL_PATH):
        problems.append(
            f"TRACKER_EXCEL_PATH is set to '{TRACKER_EXCEL_PATH}' but that file doesn't exist yet."
        )

    if problems:
        print("⚠  Configuration issues found:")
        for p in problems:
            print(f"   - {p}")
        print()

    fatal = any(
        ('placeholder' in p and 'QRADAR_' in p) or 'must be' in p
        for p in problems
    )
    return not fatal


def test_connection():
    print("🔗 Testing QRadar connection...")
    try:
        result = _api_get('/api/help/versions', label='connection test')
        if result:
            print("✅ QRadar connection successful!")
            return True
        print("⚠  Unexpected empty response from /api/help/versions")
        return False
    except RuntimeError as e:
        print(f"❌ {e}")
        return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


def fetch_all_rules():
    print("📥 Fetching correlation rules...")
    try:
        data = _api_get_all('/api/analytics/rules', label='analytics rules')
        rules = [
            r for r in data
            if r.get('enabled', False) is True
            and 'BB' not in str(r.get('type', '')).upper()
            and 'BUILDING_BLOCK' not in str(r.get('type', '')).upper()
        ]
        print(f"   ✅ {len(rules)} enabled correlation rules (of {len(data)} total).")
        return rules
    except Exception as e:
        print(f"   ❌ Failed to fetch rules: {e}")
        return []


def fetch_all_offenses():
    if FULL_HISTORY_LOOKBACK_DAYS:
        cutoff_ms  = int((datetime.now() - timedelta(days=FULL_HISTORY_LOOKBACK_DAYS)).timestamp() * 1000)
        filter_str = f'start_time >= {cutoff_ms}'
        print(f"📥 Fetching offenses from the last {FULL_HISTORY_LOOKBACK_DAYS} days (paginated)...")
    else:
        cutoff_ms  = None
        filter_str = None
        print("📥 Fetching ALL offenses currently retained by QRadar (no date filter, paginated)...")

    fields_str = 'id,rules,start_time,magnitude,status,severity'
    params = {'fields': fields_str}
    if filter_str:
        params['filter'] = filter_str

    try:
        data = _api_get_all('/api/siem/offenses', params=params, label='offenses')
        print(f"   ✅ {len(data)} offenses fetched.")
        return data
    except Exception as e:
        if filter_str:
            print(f"   ⚠  Filtered offense fetch failed ({e}). Trying unfiltered fetch...")
            try:
                data = _api_get_all('/api/siem/offenses', params={'fields': fields_str},
                                     label='offenses (unfiltered)')
                data = [o for o in data if (o.get('start_time') or 0) >= cutoff_ms]
                print(f"   ✅ {len(data)} offenses after in-memory filter.")
                return data
            except Exception as e2:
                print(f"   ❌ Offense fetch failed entirely: {e2} — every rule will show as Dead this run.")
                return []
        print(f"   ❌ Offense fetch failed: {e} — every rule will show as Dead this run.")
        return []


def filter_offenses_by_days(offenses, days):
    cutoff_ms = int((datetime.now() - timedelta(days=days)).timestamp() * 1000)
    return [o for o in offenses if (o.get('start_time') or 0) >= cutoff_ms]


def _extract_rule_ids(offense):
    ids = []
    for r in offense.get('rules', []):
        if isinstance(r, dict):
            rid = r.get('id')
        elif isinstance(r, (int, str)):
            rid = r
        else:
            rid = None
        if rid is not None:
            try:
                ids.append(int(rid))
            except (TypeError, ValueError):
                pass
    return ids


def _html_escape(text):
    if text is None:
        return ''
    text = str(text)
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))


def _csv_formula_safe(text):
    if text is None:
        return ''
    text = str(text)
    if text and text[0] in ('=', '+', '-', '@', '\t', '\r'):
        return "'" + text
    return text


def _native(val):
    try:
        if pd.isna(val):
            return ''
    except (TypeError, ValueError):
        pass
    if isinstance(val, pd.Timestamp):
        return val.to_pydatetime()
    if hasattr(val, 'item'):
        try:
            return val.item()
        except Exception:
            pass
    return val


def _format_epoch_ms(val):
    if not val:
        return '—'
    try:
        return datetime.fromtimestamp(int(val) / 1000).strftime('%Y-%m-%d')
    except (TypeError, ValueError, OSError, OverflowError):
        return '—'


def _age_days_from_epoch(val):
    if not val:
        return None
    try:
        created = datetime.fromtimestamp(int(val) / 1000)
        return (datetime.now() - created).days
    except (TypeError, ValueError, OSError, OverflowError):
        return None


def _get_rule_description(rule):
    for field in ('notes', 'description', 'rule_description', 'comment'):
        val = rule.get(field)
        if val and isinstance(val, str) and val.strip():
            return val.strip()
    return 'Not exposed via API — review in QRadar Console (Offenses → Rules) for full trigger logic.'


def _find_col(columns, aliases):
    lower_map = {str(c).strip().lower(): c for c in columns}
    for alias in aliases:
        if alias in lower_map:
            return lower_map[alias]
    return None


def _to_bool(val):
    try:
        if pd.isna(val):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val != 0
    return str(val).strip().lower() in _TRUE_STRINGS


def _normalize_name(s):
    return ' '.join(str(s).strip().lower().split())


def load_tracker_file(path):
    if not path:
        return None
    if not os.path.exists(path):
        logger.warning("Tracker file not found at '%s'.", path)
        return None
    try:
        df = pd.read_excel(path, engine='openpyxl', sheet_name=0)
        sheet_name = pd.ExcelFile(path, engine='openpyxl').sheet_names[0]
    except Exception as e:
        logger.warning("Could not read tracker file '%s' (%s) — skipping it this run.", path, e)
        return None

    name_col   = _find_col(df.columns, _RULE_NAME_COL_ALIASES)
    inv_col    = _find_col(df.columns, _INVESTIGATION_COL_ALIASES)
    tested_col = _find_col(df.columns, _TESTED_COL_ALIASES)

    if inv_col is None and tested_col is None and not df.empty:
        logger.warning("Tracker file '%s' has neither an 'Under Investigation' nor a 'Tested' column.", path)
    if name_col is None and not df.empty:
        logger.warning(
            "Tracker file '%s' has no recognizable Rule Name column — cross-referencing against "
            "QRadar and the auto-mark-tested logic are disabled this run; only raw counts will be reported.",
            path
        )

    return {
        'raw':                 df,
        'sheet_name':          sheet_name,
        'name_col':            name_col,
        'investigation_col':   inv_col,
        'tested_col':          tested_col,
        'changed_cells':       [],
        'auto_marked_rows':    [],
        'auto_marked_indices': set(),
        'unmatched_count':     None,
        'under_investigation_count': 0,
        'tested_count':        0,
    }


def reconcile_tracker_with_qradar(tracker_info, master_df):
    df = tracker_info['raw']
    name_col   = tracker_info['name_col']
    inv_col    = tracker_info['investigation_col']
    tested_col = tracker_info['tested_col']

    lookup = {}
    if not master_df.empty:
        for _, r in master_df.iterrows():
            key = _normalize_name(r['rule_name'])
            if key not in lookup:
                lookup[key] = r

    matched_flags, recent_list, total_list = [], [], []
    changed_cells, auto_marked_rows, auto_marked_indices = [], [], set()

    for idx, row in df.iterrows():
        matched_rule = lookup.get(_normalize_name(row[name_col])) if name_col else None

        is_investigating = _to_bool(row[inv_col]) if inv_col else False
        is_tested        = _to_bool(row[tested_col]) if tested_col else False
        recent = int(matched_rule['offenses_recent']) if matched_rule is not None else None
        total  = int(matched_rule['offenses_total']) if matched_rule is not None else None

        matched_flags.append(matched_rule is not None)
        recent_list.append(recent)
        total_list.append(total)

        should_auto_mark = (
            AUTO_UPDATE_TRACKER_TESTED
            and matched_rule is not None
            and is_investigating
            and not is_tested
            and tested_col is not None
            and recent is not None
            and recent > 0
        )
        if should_auto_mark:
            df.at[idx, tested_col] = True
            changed_cells.append((idx, tested_col, True))
            if AUTO_CLEAR_INVESTIGATION_ON_TESTED and inv_col:
                df.at[idx, inv_col] = False
                changed_cells.append((idx, inv_col, False))
            auto_marked_indices.add(idx)
            auto_marked_rows.append({
                'rule_name': str(row[name_col]),
                'offenses_recent': recent,
            })

    df['_matched_to_qradar'] = matched_flags
    df['_offenses_recent']   = recent_list
    df['_offenses_total']    = total_list

    tracker_info['raw'] = df
    tracker_info['changed_cells'] = changed_cells
    tracker_info['auto_marked_rows'] = auto_marked_rows
    tracker_info['auto_marked_indices'] = auto_marked_indices
    tracker_info['unmatched_count'] = int(len(matched_flags) - sum(matched_flags)) if name_col else None
    tracker_info['under_investigation_count'] = int(df[inv_col].apply(_to_bool).sum()) if inv_col else 0
    tracker_info['tested_count'] = int(df[tested_col].apply(_to_bool).sum()) if tested_col else 0
    return tracker_info


def persist_tracker_updates(path, tracker_info):
    changed_cells = tracker_info.get('changed_cells') or []
    if not changed_cells:
        return 'no_changes'
    if not AUTO_UPDATE_TRACKER_TESTED:
        return 'disabled'

    try:
        if TRACKER_BACKUP_BEFORE_WRITE:
            backup_path = path.replace('.xlsx', f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            shutil.copy2(path, backup_path)

        wb = openpyxl.load_workbook(path)
        ws = wb[tracker_info['sheet_name']]
        header_cells = ws[TRACKER_HEADER_ROW]
        col_letter_by_name = {str(c.value).strip(): get_column_letter(c.column) for c in header_cells if c.value}

        for idx, col_name, value in changed_cells:
            col_letter = col_letter_by_name.get(col_name)
            if not col_letter:
                continue
            excel_row = idx + TRACKER_HEADER_ROW + 1
            ws[f'{col_letter}{excel_row}'] = value

        wb.save(path)
        return 'written'
    except PermissionError:
        logger.warning("Tracker file '%s' is open elsewhere — could not persist updates this run.", path)
        return 'locked'
    except Exception as e:
        logger.error("Could not write updates back to tracker file '%s': %s", path, e)
        return 'error'


def build_master_rule_table(rules, offenses_all):
    full_counts = {}
    for o in offenses_all:
        for rid in _extract_rule_ids(o):
            full_counts[rid] = full_counts.get(rid, 0) + 1

    recent_offenses = filter_offenses_by_days(offenses_all, TRACKER_VALIDATION_WINDOW_DAYS)
    recent_counts = {}
    for o in recent_offenses:
        for rid in _extract_rule_ids(o):
            recent_counts[rid] = recent_counts.get(rid, 0) + 1

    rows = []
    for rule in rules:
        rid    = rule.get('id')
        total  = full_counts.get(rid, 0)
        recent = recent_counts.get(rid, 0)
        rows.append({
            'rule_id':          rid,
            'rule_name':        _csv_formula_safe(rule.get('name', f'Rule {rid}')),
            'rule_type':        rule.get('type', 'UNKNOWN'),
            'origin':           rule.get('origin', 'UNKNOWN'),
            'owner':            rule.get('owner', 'Unknown'),
            'created':          _format_epoch_ms(rule.get('creation_date')),
            'created_age_days': _age_days_from_epoch(rule.get('creation_date')),
            'modified':         _format_epoch_ms(rule.get('modification_date')),
            'description':      _csv_formula_safe(_get_rule_description(rule)),
            'offenses_total':   total,
            'offenses_recent':  recent,
            'is_dead':          total <= RULE_DEAD_THRESHOLD,
        })
    return pd.DataFrame(rows)


_FILLS = {
    'red':         PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid'),
    'orange':      PatternFill(start_color='FFBF47', end_color='FFBF47', fill_type='solid'),
    'green':       PatternFill(start_color='A8E6CF', end_color='A8E6CF', fill_type='solid'),
    'blue':        PatternFill(start_color='74B9FF', end_color='74B9FF', fill_type='solid'),
    'header':      PatternFill(start_color='2D2257', end_color='2D2257', fill_type='solid'),
    'zebra_a':     PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid'),
    'zebra_b':     PatternFill(start_color='F3F1FB', end_color='F3F1FB', fill_type='solid'),
    'stale':       PatternFill(start_color='FFE2E2', end_color='FFE2E2', fill_type='solid'),
    'auto_marked': PatternFill(start_color='C6F6D5', end_color='C6F6D5', fill_type='solid'),
}
_HDR_FONT   = Font(bold=True, color='E8E0FF', size=10)
_BOLD       = Font(bold=True, size=10)
_STALE_FONT = Font(bold=True, size=10, color='B91C1C')
_CENTRE     = Alignment(horizontal='center', vertical='center')
_WRAP       = Alignment(horizontal='left',   vertical='top', wrap_text=True)
_THIN_SIDE  = Side(style='thin', color='DCD7EE')
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)


def _write_sheet_header(ws, columns, col_widths):
    ws.append(columns)
    for col_idx, (cell, width) in enumerate(zip(ws[1], col_widths), start=1):
        cell.fill      = _FILLS['header']
        cell.font      = _HDR_FONT
        cell.alignment = _CENTRE
        cell.border    = _THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'


def _apply_zebra_and_borders(ws, first_data_row, last_data_row, num_cols):
    for r in range(first_data_row, last_data_row + 1):
        fill = _FILLS['zebra_a'] if (r - first_data_row) % 2 == 0 else _FILLS['zebra_b']
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            if cell.fill.start_color.rgb in (None, '00000000'):
                cell.fill = fill
            cell.border = _THIN_BORDER


def _pct(n, total):
    return f"{(n / total * 100):.1f}%" if total else "0.0%"


def _write_summary_sheet(wb, dead_df, tracker_info, total_rules):
    ws = wb.active
    ws.title = 'Executive Summary'

    ws['A1'] = f'QRadar — {REPORT_TITLE}'
    ws['A1'].font = Font(bold=True, size=16, color='2D2257')
    ws['A2'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws['A2'].font = Font(italic=True, size=10, color='7C6FA0')

    hist_note = ('unbounded — all offenses currently retained by QRadar'
                 if not FULL_HISTORY_LOOKBACK_DAYS else f'last {FULL_HISTORY_LOOKBACK_DAYS} days')
    ws['A3'] = f"Full-history window: {hist_note}   ·   Tracker validation window: last {TRACKER_VALIDATION_WINDOW_DAYS} day(s)"
    ws['A3'].font = Font(size=10, color='7C6FA0')

    ws['A5'] = 'Total enabled correlation rules analyzed:'
    ws['B5'] = int(total_rules)
    ws['B5'].font = _BOLD

    row = 7
    ws.cell(row=row, column=1, value='Metric').fill = _FILLS['header']
    ws.cell(row=row, column=1).font = _HDR_FONT
    for c, label in ((2, 'Count'), (3, '% of Rules')):
        ws.cell(row=row, column=c, value=label).fill = _FILLS['header']
        ws.cell(row=row, column=c).font = _HDR_FONT
        ws.cell(row=row, column=c).alignment = _CENTRE
    row += 1

    dead_n = len(dead_df)
    metrics = [
        ('Dead — Never Fired', dead_n, _pct(dead_n, total_rules), 'red'),
        ('Under Investigation', tracker_info['under_investigation_count'] if tracker_info else 'N/A', '', 'orange'),
        ('Testing Completed',   tracker_info['tested_count']              if tracker_info else 'N/A', '', 'blue'),
    ]
    for label, value, pct, fill_key in metrics:
        ws.cell(row=row, column=1, value=label).font = _BOLD
        c1 = ws.cell(row=row, column=2, value=value)
        c1.alignment = _CENTRE
        if value != 'N/A':
            c1.fill = _FILLS[fill_key]
        ws.cell(row=row, column=3, value=pct).alignment = _CENTRE
        row += 1
    row += 1

    if tracker_info is None:
        tracker_note = (
            f"Investigation/Testing tracker not configured or not found "
            f"(TRACKER_EXCEL_PATH = '{TRACKER_EXCEL_PATH or '(blank)'}')."
        )
    else:
        crossref = 'enabled (matched by Rule Name)' if tracker_info['name_col'] else 'disabled (no Rule Name column found)'
        auto_n = len(tracker_info['auto_marked_rows'])
        unmatched = tracker_info['unmatched_count']
        tracker_note = (
            f"Tracker source: {TRACKER_EXCEL_PATH}  ·  {len(tracker_info['raw'])} row(s) read  ·  "
            f"Cross-referencing: {crossref}  ·  Auto-marked Tested this run: {auto_n}"
            + (f"  ·  Unmatched rows: {unmatched}" if unmatched is not None else "")
        )
    ws.cell(row=row, column=1, value=tracker_note).font = Font(italic=True, size=9, color='7C6FA0')
    ws.cell(row=row, column=1).alignment = _WRAP
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 40
    row += 2

    caveat = (
        'Rule descriptions reflect whatever metadata QRadar exposes via REST API; full boolean '
        'trigger logic is not reliably exposed in most versions — cross-reference the Rules console. '
        '"Dead" is bounded by whatever offense history QRadar currently retains (or by '
        'FULL_HISTORY_LOOKBACK_DAYS, if set). Auto-marking Tested uses recent offense activity as a '
        'proxy for "fired since you started investigating" — it does not confirm the offense was '
        'caused by your specific test versus unrelated real traffic.'
    )
    ws.cell(row=row, column=1, value=caveat).font = Font(italic=True, size=9, color='7C6FA0')
    ws.cell(row=row, column=1).alignment = _WRAP
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 60

    for i, w in enumerate([46, 14, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_dead_sheet(wb, dead_df):
    ws = wb.create_sheet('Dead Rules')
    cols   = ['Rule ID', 'Rule Name', 'Type', 'Origin', 'Owner', 'Created', 'Age (Days)',
              'Modified', 'Description / Trigger Logic']
    widths = [9, 42, 14, 10, 14, 12, 11, 12, 60]
    _write_sheet_header(ws, cols, widths)

    if dead_df.empty:
        ws.append(['No dead rules — every enabled rule has fired at least once.'])
        return

    sorted_df = dead_df.sort_values('created_age_days', ascending=False, na_position='last')
    desc_idx = len(cols)
    age_idx  = cols.index('Age (Days)') + 1
    first_row = ws.max_row + 1

    for _, r in sorted_df.iterrows():
        age = r['created_age_days']
        ws.append([int(r['rule_id']), str(r['rule_name']), str(r['rule_type']), str(r['origin']),
                   str(r['owner']), str(r['created']), age if age is not None else '—',
                   str(r['modified']), str(r['description'])])
        r_idx = ws.max_row
        ws.cell(row=r_idx, column=desc_idx).alignment = _WRAP
        if age is not None and age >= DEAD_RULE_STALE_AGE_DAYS:
            ws.cell(row=r_idx, column=age_idx).fill = _FILLS['stale']
            ws.cell(row=r_idx, column=age_idx).font = _STALE_FONT

    _apply_zebra_and_borders(ws, first_row, ws.max_row, len(cols))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"


def _write_tracker_sheet(wb, tracker_info):
    ws = wb.create_sheet('Investigation & Testing Tracker')
    df = tracker_info['raw']
    original_cols = [c for c in df.columns if not str(c).startswith('_')]
    has_match_info = '_matched_to_qradar' in df.columns
    extra_cols = (
        ['Matched to QRadar Rule?', f'Offenses (last {TRACKER_VALIDATION_WINDOW_DAYS}d)',
         'Offenses (Total)', 'Auto-Marked Tested This Run?']
        if has_match_info else []
    )
    cols = [str(c) for c in original_cols] + extra_cols
    widths = [max(14, min(50, len(str(c)) + 4)) for c in original_cols] + [20, 18, 16, 20][:len(extra_cols)]
    _write_sheet_header(ws, cols, widths)

    if df.empty:
        ws.append(['Tracker file has no rows.'])
        return

    inv_col, tested_col = tracker_info['investigation_col'], tracker_info['tested_col']
    auto_marked_indices = tracker_info.get('auto_marked_indices', set())
    first_row = ws.max_row + 1

    for idx, r in df.iterrows():
        values = [_native(r[c]) for c in original_cols]
        if has_match_info:
            recent, total = r['_offenses_recent'], r['_offenses_total']
            values += [
                'Yes' if r['_matched_to_qradar'] else 'No',
                int(recent) if pd.notna(recent) else '—',
                int(total) if pd.notna(total) else '—',
                'Yes' if idx in auto_marked_indices else 'No',
            ]
        ws.append(values)
        r_idx = ws.max_row

        if idx in auto_marked_indices:
            for c_idx in range(1, len(cols) + 1):
                ws.cell(row=r_idx, column=c_idx).fill = _FILLS['auto_marked']

        for c_idx, c in enumerate(original_cols, start=1):
            if c == inv_col and _to_bool(r[c]):
                ws.cell(row=r_idx, column=c_idx).fill = _FILLS['orange']
            elif c == tested_col and _to_bool(r[c]):
                ws.cell(row=r_idx, column=c_idx).fill = _FILLS['blue']

    _apply_zebra_and_borders(ws, first_row, ws.max_row, len(cols))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"


def save_excel_report(dead_df, tracker_info, total_rules, path):
    try:
        wb = openpyxl.Workbook()
        _write_summary_sheet(wb, dead_df, tracker_info, total_rules)
        _write_dead_sheet(wb, dead_df)
        if tracker_info is not None:
            _write_tracker_sheet(wb, tracker_info)

        try:
            wb.save(path)
            print(f"✅ Excel saved → {path}")
            return path
        except PermissionError:
            fallback = path.replace('.xlsx', f"_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            print(f"⚠  '{path}' is open elsewhere — saving to '{fallback}' instead.")
            wb.save(fallback)
            print(f"✅ Excel saved → {fallback}")
            return fallback
    except Exception as e:
        logger.error("Excel save failed:\n%s", traceback.format_exc())
        print(f"❌ Excel save error: {e}")
        return None


_CHART_BG    = '#0a0a10'
_CHART_GRID  = '#201c33'
_CHART_TEXT  = '#cfc8ea'
_CHART_TITLE = '#a78bfa'
_CHART_AXIS  = '#332c50'


def generate_status_chart(dead_count, tracker_info, total_rules):
    labels  = ['Dead']
    values  = [dead_count]
    colours = ['#f87171']

    if tracker_info is not None:
        labels  += ['Under\nInvestigation', 'Testing\nCompleted']
        values  += [tracker_info['under_investigation_count'], tracker_info['tested_count']]
        colours += ['#f5b155', '#5b8cff']

    fig, ax = plt.subplots(figsize=(6.0, 4.0), facecolor=_CHART_BG)
    ax.set_facecolor(_CHART_BG)
    ax.yaxis.grid(True, color=_CHART_GRID, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)

    bars = ax.bar(labels, values, color=colours, edgecolor=_CHART_BG, linewidth=1.2, zorder=3, width=0.5)
    max_val = max(values + [1])
    for bar, val in zip(bars, values):
        pct = f"{val / total_rules * 100:.1f}% of rules" if total_rules else ""
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_val * 0.03, str(val),
                ha='center', va='bottom', color=_CHART_TEXT, fontsize=11, fontweight='bold')
        if pct:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() * 0.5, pct,
                    ha='center', va='center', color=_CHART_BG, fontsize=7.5)

    ax.set_ylim(0, max_val * 1.25)
    ax.set_title('Rule Status Overview', color=_CHART_TITLE, fontsize=12, fontweight='700', pad=14)
    ax.tick_params(colors=_CHART_TEXT, labelsize=9.5)
    for spine in ax.spines.values():
        spine.set_edgecolor(_CHART_AXIS)
        spine.set_linewidth(0.6)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout(pad=1.0)

    tmp  = tempfile.NamedTemporaryFile(suffix='.png', delete=False, prefix='qr_status_')
    path = tmp.name
    tmp.close()
    plt.savefig(path, bbox_inches='tight', dpi=120, facecolor=_CHART_BG, edgecolor='none')
    plt.close(fig)
    return path


_C = {
    'page':   '#0a0a10',
    'card':   '#111017',
    'border': '#26223a',
    'purple': '#8b5cf6',
    'violet': '#a78bfa',
    'dim':    '#8d87a8',
    'text':   '#f4f2fb',
    'red':    '#f87171',
    'amber':  '#f5b155',
    'blue':   '#5b8cff',
}


def _dot(color):
    return f'<span style="color:{color};font-size:10px;vertical-align:middle;">●</span>'


def _build_dead_table_html(dead_df):
    C = _C
    if dead_df.empty:
        return (f'<p style="color:{C["blue"]};font-size:12px;font-weight:600;padding:8px 0;margin:0;">'
                f'{_dot(C["blue"])} No dead rules — every enabled rule has fired at least once.</p>')

    shown = dead_df.sort_values('created_age_days', ascending=False, na_position='last').head(EMAIL_TABLE_ROW_CAP)
    rows_html = ''
    for i, (_, row) in enumerate(shown.iterrows()):
        bg = C['page'] if i % 2 == 0 else C['card']
        name = _html_escape(row['rule_name'])
        name_short = name[:42] + '…' if len(name) > 42 else name
        age = row['created_age_days']
        age_txt = f"{age}d" if age is not None else '—'
        rows_html += f"""
        <tr bgcolor="{bg}" style="background-color:{bg};">
          <td bgcolor="{bg}" style="padding:9px 12px;border-left:2px solid {C['red']};font-size:12px;color:{C['text']};" title="{name}">{name_short}</td>
          <td bgcolor="{bg}" style="padding:9px 12px;font-size:11px;color:{C['dim']};text-align:center;">{_html_escape(row['rule_type'])}</td>
          <td bgcolor="{bg}" style="padding:9px 12px;font-size:11px;color:{C['dim']};text-align:center;">{_html_escape(row['owner'])}</td>
          <td bgcolor="{bg}" style="padding:9px 12px;font-size:11px;color:{C['dim']};text-align:center;">{_html_escape(row['created'])}</td>
          <td bgcolor="{bg}" style="padding:9px 12px;font-size:11px;color:{C['red']};text-align:center;font-weight:600;">{age_txt}</td>
        </tr>"""

    more_note = ''
    if len(dead_df) > EMAIL_TABLE_ROW_CAP:
        more_note = (f'<div style="font-size:10px;color:{C["dim"]};margin-top:8px;">'
                     f'+ {len(dead_df) - EMAIL_TABLE_ROW_CAP} more — see attached Excel for the full list.</div>')

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="{C['page']}" style="border-collapse:collapse;margin-top:10px;background-color:{C['page']};">
      <thead><tr>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:left;font-size:9px;color:{C['red']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Rule Name</th>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:center;font-size:9px;color:{C['red']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Type</th>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:center;font-size:9px;color:{C['red']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Owner</th>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:center;font-size:9px;color:{C['red']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Created</th>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:center;font-size:9px;color:{C['red']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Age</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>{more_note}"""


def _build_under_investigation_table_html(tracker_info):
    C = _C
    if tracker_info is None or not tracker_info.get('investigation_col'):
        return (f'<p style="color:{C["dim"]};font-size:12px;padding:8px 0;margin:0;">'
                f'Tracker not configured — set TRACKER_EXCEL_PATH to see active investigations here.</p>')

    df = tracker_info['raw']
    inv_col, name_col = tracker_info['investigation_col'], tracker_info['name_col']
    has_match = '_matched_to_qradar' in df.columns

    mask = df[inv_col].apply(_to_bool)
    shown = df[mask]
    if shown.empty:
        return (f'<p style="color:{C["blue"]};font-size:12px;font-weight:600;padding:8px 0;margin:0;">'
                f'{_dot(C["blue"])} Nothing currently under investigation.</p>')

    shown = shown.head(EMAIL_TABLE_ROW_CAP)
    rows_html = ''
    for i, (_, row) in enumerate(shown.iterrows()):
        bg = C['page'] if i % 2 == 0 else C['card']
        name = _html_escape(row[name_col]) if name_col else '(no name column)'
        name_short = name[:42] + '…' if len(name) > 42 else name
        recent = row['_offenses_recent'] if has_match else None
        if recent is not None and pd.notna(recent):
            fired_txt, fired_color = ('Yes', C['blue']) if recent > 0 else ('No', C['dim'])
            recent_txt = str(int(recent))
        else:
            fired_txt, fired_color, recent_txt = '—', C['dim'], '—'
        rows_html += f"""
        <tr bgcolor="{bg}" style="background-color:{bg};">
          <td bgcolor="{bg}" style="padding:9px 12px;border-left:2px solid {C['amber']};font-size:12px;color:{C['text']};" title="{name}">{name_short}</td>
          <td bgcolor="{bg}" style="padding:9px 12px;font-size:11px;color:{fired_color};text-align:center;font-weight:600;">{fired_txt}</td>
          <td bgcolor="{bg}" style="padding:9px 12px;font-size:11px;color:{C['dim']};text-align:center;">{recent_txt}</td>
        </tr>"""

    more_note = ''
    if mask.sum() > EMAIL_TABLE_ROW_CAP:
        more_note = (f'<div style="font-size:10px;color:{C["dim"]};margin-top:8px;">'
                     f'+ {int(mask.sum()) - EMAIL_TABLE_ROW_CAP} more — see attached Excel for the full list.</div>')

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="{C['page']}" style="border-collapse:collapse;margin-top:10px;background-color:{C['page']};">
      <thead><tr>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:left;font-size:9px;color:{C['amber']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Rule Name</th>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:center;font-size:9px;color:{C['amber']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Fired Recently?</th>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:center;font-size:9px;color:{C['amber']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Offenses (last {TRACKER_VALIDATION_WINDOW_DAYS}d)</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>{more_note}"""


def _build_auto_validated_table_html(tracker_info):
    C = _C
    rows = tracker_info.get('auto_marked_rows') if tracker_info else None
    if not rows:
        return None

    shown = rows[:EMAIL_TABLE_ROW_CAP]
    rows_html = ''
    for i, r in enumerate(shown):
        bg = C['page'] if i % 2 == 0 else C['card']
        name = _html_escape(r['rule_name'])
        name_short = name[:48] + '…' if len(name) > 48 else name
        rows_html += f"""
        <tr bgcolor="{bg}" style="background-color:{bg};">
          <td bgcolor="{bg}" style="padding:9px 12px;border-left:2px solid {C['blue']};font-size:12px;color:{C['text']};" title="{name}">{name_short}</td>
          <td bgcolor="{bg}" style="padding:9px 12px;font-size:11px;color:{C['blue']};text-align:center;font-weight:600;">{r['offenses_recent']}</td>
        </tr>"""

    more_note = ''
    if len(rows) > EMAIL_TABLE_ROW_CAP:
        more_note = (f'<div style="font-size:10px;color:{C["dim"]};margin-top:8px;">'
                     f'+ {len(rows) - EMAIL_TABLE_ROW_CAP} more — see attached Excel.</div>')

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="{C['page']}" style="border-collapse:collapse;margin-top:10px;background-color:{C['page']};">
      <thead><tr>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:left;font-size:9px;color:{C['blue']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Rule Name</th>
        <th bgcolor="{C['page']}" style="padding:8px 12px;text-align:center;font-size:9px;color:{C['blue']};font-weight:700;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid {C['border']};">Offenses (last {TRACKER_VALIDATION_WINDOW_DAYS}d)</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>{more_note}"""


def build_email_html(dead_df, tracker_info, total_rules, chart_cid):
    C = _C
    run_time = datetime.now().strftime('%d %b %Y  ·  %H:%M:%S')

    dead_count   = len(dead_df)
    inv_count    = tracker_info['under_investigation_count'] if tracker_info else None
    tested_count = tracker_info['tested_count']              if tracker_info else None
    auto_marked  = tracker_info['auto_marked_rows']          if tracker_info else []

    if dead_count > 0:
        hdr_txt, hdr_color = f'{dead_count} DEAD RULE(S) TO REVIEW', C['red']
    elif auto_marked:
        hdr_txt, hdr_color = f'{len(auto_marked)} AUTO-VALIDATED THIS RUN', C['blue']
    else:
        hdr_txt, hdr_color = 'STEADY STATE', C['violet']

    def badge(color, txt):
        return (f'<span style="color:{color};font-size:10px;font-weight:700;'
                f'padding:5px 13px;border-radius:20px;letter-spacing:0.6px;border:1px solid {color}80;'
                f'white-space:nowrap;">{txt}</span>')

    def metric_card(label, value, color, note=''):
        display_val = value if value is not None else '—'
        note_html = (f'<div style="font-size:9px;color:{C["dim"]};margin-top:4px;">{note}</div>'
                     if note else '')
        return (f'<td width="33%" bgcolor="{C["page"]}" style="padding:5px;background-color:{C["page"]};">'
                f'<table width="100%" cellpadding="0" cellspacing="0" bgcolor="{C["card"]}" '
                f'style="border-collapse:collapse;background-color:{C["card"]};border:1px solid {C["border"]};'
                f'border-top:2px solid {color};border-radius:8px;">'
                f'<tr><td bgcolor="{C["card"]}" style="background-color:{C["card"]};padding:16px 8px;text-align:center;">'
                f'<div style="font-size:28px;font-weight:800;color:{color};line-height:1;letter-spacing:-1px;">{display_val}</div>'
                f'<div style="font-size:9px;color:{C["dim"]};margin-top:6px;text-transform:uppercase;letter-spacing:1px;">{label}</div>'
                f'{note_html}</td></tr></table></td>')

    dead_pct = f"{dead_count / total_rules * 100:.1f}% of rules" if total_rules else ''
    inv_note = '' if tracker_info else 'not configured'
    tested_note = f"+{len(auto_marked)} auto-marked this run" if auto_marked else ('' if tracker_info else 'not configured')

    headline_metrics = (
        metric_card('Dead Rules', dead_count, C['red'], dead_pct)
        + metric_card('Under Investigation', inv_count, C['amber'], inv_note)
        + metric_card('Testing Completed', tested_count, C['blue'], tested_note)
    )

    chart_html = (f'<img src="cid:{chart_cid}" alt="Rule status chart" '
                  f'style="display:block;max-width:100%;margin:18px auto 0;">') if chart_cid else ''
    dead_table_html   = _build_dead_table_html(dead_df)
    inv_table_html    = _build_under_investigation_table_html(tracker_info)
    auto_table_html   = _build_auto_validated_table_html(tracker_info)

    def section_header(color, title, subtitle):
        return f"""
  <tr><td bgcolor="{C['page']}" style="background-color:{C['page']};padding:28px 0 4px;border-top:1px solid {C['border']};">
    <span style="font-size:13px;font-weight:700;color:{C['text']};">{_dot(color)}&nbsp;&nbsp;{title}</span>
    <div style="font-size:10px;color:{C['dim']};margin-top:4px;">{subtitle}</div>
  </td></tr>"""

    auto_section = ''
    if auto_table_html:
        auto_section = (
            section_header(C['blue'], 'Auto-Validated This Run',
                            'Marked Under Investigation and fired since — Tested has been set to Yes in your tracker file.')
            + f'<tr><td bgcolor="{C["page"]}" style="background-color:{C["page"]};padding:0 0 20px;">{auto_table_html}</td></tr>'
        )

    hist_note = ('all offenses currently retained by QRadar' if not FULL_HISTORY_LOOKBACK_DAYS
                 else f'the last {FULL_HISTORY_LOOKBACK_DAYS} days')

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="dark">
<meta name="supported-color-schemes" content="dark">
<title>{REPORT_TITLE}</title>
<!--[if mso]>
<style type="text/css">
table, td {{ font-family:Segoe UI, Arial, sans-serif !important; }}
</style>
<![endif]-->
</head>
<body bgcolor="{C['page']}" style="margin:0;padding:0;background-color:{C['page']};font-family:'Segoe UI',Helvetica,Arial,sans-serif;">
<table role="presentation" width="100%" bgcolor="{C['page']}" cellpadding="0" cellspacing="0" style="background-color:{C['page']};">
<tr>
<td align="center" bgcolor="{C['page']}" style="background-color:{C['page']};padding:32px 16px;">

<table role="presentation" width="640" cellpadding="0" cellspacing="0" bgcolor="{C['page']}" style="width:640px;max-width:640px;background-color:{C['page']};">

  <tr><td bgcolor="{C['page']}" style="background-color:{C['page']};padding:0 0 18px;border-bottom:2px solid {C['purple']};">
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="{C['page']}" style="background-color:{C['page']};"><tr>
      <td bgcolor="{C['page']}" style="background-color:{C['page']};">
        <div style="font-size:9px;color:{C['dim']};letter-spacing:3px;text-transform:uppercase;margin-bottom:8px;">
          QRadar
        </div>
        <div style="font-size:23px;font-weight:800;color:{C['text']};letter-spacing:-0.4px;line-height:1.25;">
          {REPORT_TITLE}
        </div>
        <div style="margin-top:8px;font-size:11px;color:{C['dim']};">{run_time}</div>
      </td>
      <td align="right" valign="top" bgcolor="{C['page']}" style="background-color:{C['page']};">{badge(hdr_color, hdr_txt)}</td>
    </tr></table>
  </td></tr>

  <tr><td bgcolor="{C['page']}" style="background-color:{C['page']};padding:20px 0 4px;">
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="{C['page']}" style="background-color:{C['page']};"><tr>{headline_metrics}</tr></table>
  </td></tr>

  <tr><td bgcolor="{C['page']}" style="background-color:{C['page']};padding:0 0 4px;text-align:center;">{chart_html}</td></tr>

  {section_header(C['red'], 'Dead Rules — Never Fired',
                   f"Zero offenses across {hist_note}. Candidates for review or removal.")}
  <tr><td bgcolor="{C['page']}" style="background-color:{C['page']};padding:0 0 8px;">{dead_table_html}</td></tr>

  {section_header(C['amber'], 'Rules Under Investigation',
                   f"Currently flagged in your tracker — 'Fired Recently?' checks the last {TRACKER_VALIDATION_WINDOW_DAYS} day(s).")}
  <tr><td bgcolor="{C['page']}" style="background-color:{C['page']};padding:0 0 8px;">{inv_table_html}</td></tr>

  {auto_section}

  <tr><td bgcolor="{C['page']}" style="background-color:{C['page']};padding:18px 0 4px;border-top:1px solid {C['border']};">
    <div style="font-size:9px;color:{C['dim']};letter-spacing:0.3px;line-height:1.6;">
      QRadar Rule Status Auditor &nbsp;·&nbsp; Auto-generated {run_time}<br>
      Full rule details, and the Investigation/Testing tracker (if configured), are in the attached Excel workbook.
    </div>
  </td></tr>

</table>

</td>
</tr>
</table>
</body></html>"""


def create_outlook_draft(excel_path, subject, html_body, images, high_importance=False):
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail    = outlook.CreateItem(0)
        mail.Subject = subject
        if high_importance:
            mail.Importance = 2

        if excel_path and os.path.exists(excel_path):
            mail.Attachments.Add(excel_path)

        for cid, img_path in images.items():
            if img_path and os.path.exists(img_path):
                att = mail.Attachments.Add(img_path)
                att.PropertyAccessor.SetProperty(_MAPI_PR_ATTACH_CONTENT_ID, cid)

        mail.HTMLBody = html_body
        mail.Display()
        print("\n✉  Outlook draft created.")
    except Exception as e:
        logger.error("Outlook draft failed:\n%s", traceback.format_exc())
        print(f"\n❌ Outlook draft error: {e}")
    finally:
        for cid, img_path in images.items():
            if img_path and os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass


def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("=" * 62)
    print(f"  QRadar — {REPORT_TITLE}")
    print("=" * 62)
    print(f"  Host                     : {QRADAR_HOST}")
    print(f"  Tracker validation window: last {TRACKER_VALIDATION_WINDOW_DAYS} day(s)")
    print(f"  Full history             : "
          f"{'unbounded (all retained offenses)' if not FULL_HISTORY_LOOKBACK_DAYS else f'{FULL_HISTORY_LOOKBACK_DAYS} days'}")
    print(f"  Tracker file             : {TRACKER_EXCEL_PATH or '(not configured)'}")
    print(f"  Auto-mark Tested         : {AUTO_UPDATE_TRACKER_TESTED}")
    print(f"  Retry config             : {MAX_RETRIES} attempts, {RETRY_DELAY_BASE}s base backoff")
    print("=" * 62)

    if not validate_config():
        print("❌ Fix the configuration issues above before running.")
        return

    if not test_connection():
        return

    rules = fetch_all_rules()
    if not rules:
        print("\n❌ No enabled rules returned — check credentials, host URL, and permissions.")
        return

    offenses_all = fetch_all_offenses()
    if not offenses_all:
        print("\n⚠  WARNING: no offenses were retrieved — check the warnings above before "
              "treating every rule below as 'Dead'.")

    print("\n🔍 Classifying rules...")
    master_df = build_master_rule_table(rules, offenses_all)
    dead_df   = master_df[master_df['is_dead']].copy()

    print(f"   Dead — never fired: {len(dead_df)}")

    undocumented = int(master_df['description'].str.contains('Not exposed via API', na=False).sum())
    if undocumented:
        print(f"\n   ℹ  Trigger-logic descriptions are unavailable via the API for {undocumented} "
              f"of {len(master_df)} rule(s).")

    print("\n📋 Loading investigation/testing tracker...")
    tracker_info = load_tracker_file(TRACKER_EXCEL_PATH) if TRACKER_EXCEL_PATH else None
    if tracker_info is None:
        print("   ℹ  Tracker not configured or unavailable this run.")
    else:
        tracker_info = reconcile_tracker_with_qradar(tracker_info, master_df)
        print(f"   ✅ {tracker_info['under_investigation_count']} under investigation, "
              f"{tracker_info['tested_count']} testing completed "
              f"(of {len(tracker_info['raw'])} tracked row(s)).")
        if tracker_info['name_col'] and tracker_info['unmatched_count']:
            print(f"   ⚠  {tracker_info['unmatched_count']} tracker row(s) didn't match any enabled QRadar rule by name.")

        if tracker_info['changed_cells']:
            result = persist_tracker_updates(TRACKER_EXCEL_PATH, tracker_info)
            n_auto = len(tracker_info['auto_marked_rows'])
            if result == 'written':
                print(f"   ✏  Auto-marked {n_auto} rule(s) as Tested — written back to {TRACKER_EXCEL_PATH} "
                      f"(backup saved alongside it).")
            elif result == 'locked':
                print(f"   ⚠  {n_auto} rule(s) qualify to be auto-marked Tested, but the tracker file is "
                      f"open elsewhere — close it and re-run to persist.")
            elif result == 'disabled':
                print(f"   ℹ  {n_auto} rule(s) qualify to be auto-marked Tested "
                      f"(AUTO_UPDATE_TRACKER_TESTED=False — not written back).")
            elif result == 'error':
                print("   ❌ Could not write updates back to the tracker file — see log for details.")

    print(f"\n💾 Saving Excel report → {OUTPUT_EXCEL}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    saved_excel_path = save_excel_report(dead_df, tracker_info, len(master_df), OUTPUT_EXCEL)

    print("\n📊 Generating status chart...")
    chart_path = generate_status_chart(len(dead_df), tracker_info, len(master_df))

    print("\n✉  Building email draft...")
    html_body = build_email_html(dead_df, tracker_info, len(master_df),
                                  chart_cid='rule_status' if chart_path else None)

    subject_parts = [f"{len(dead_df)} dead"]
    if tracker_info is not None:
        subject_parts.append(f"{tracker_info['under_investigation_count']} under investigation")
        subject_parts.append(f"{tracker_info['tested_count']} testing completed")
    subject = f"{REPORT_TITLE} — " + ", ".join(subject_parts)

    high_importance = len(dead_df) >= EMAIL_HIGH_IMPORTANCE_DEAD_THRESHOLD

    images = {'rule_status': chart_path} if chart_path else {}
    create_outlook_draft(saved_excel_path, subject, html_body, images, high_importance=high_importance)
    print("\n✅ Done!")


if __name__ == '__main__':
    main()
