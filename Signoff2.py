"""
QRadar Signoff Auto-Draft  v3.1
=================================
CHANGES FROM v3.0
──────────────────────────────────────────────
EMAIL BODY
  • Removed explicit dark backgrounds from body/container — email is now
    transparent and blends with Outlook's dark-mode theme (no floating card)
  • meta color-scheme: dark added for Outlook dark-mode signal
  • Status banners retain their colored backgrounds (intentional accents)
  • All table row backgrounds removed — subtle borders only
  • Text colours light enough to read on any dark background

DASHBOARD — FILTER / DELETE
  • Status filter pills removed from table header
  • OS-group dropdown removed from table header (OS group still visible as badge in rows)
  • Quick-delete (🗑) button added to every row — no modal, no reason required;
    10-second undo bar appears immediately after click
  • Edit (✎) button still opens the full modal for overrides/notes

DASHBOARD — VISUAL REFRESH
  • Cleaner card design: subtle gradient top-border accent, better spacing
  • Table: alternating micro-tint rows, sticky header, sharper typography
  • Header bar simplified and tightened
  • Charts section unchanged (pure-SVG, air-gapped)
  • Period selector preserved as-is

AUTO-OPEN
  • generate_dashboard() calls os.startfile(DASHBOARD_PATH) at the end of
    every run so the dashboard opens in the default browser automatically
"""

import os
import re
import json
import uuid
import time
import urllib3
import requests
import win32com.client

from datetime import datetime, timedelta

# ─── PATH AUTO-CONFIGURATION ─────────────────────────────────────────────────
_SCRIPT_DIR       = os.path.dirname(os.path.abspath(__file__))
RUN_LOG_PATH      = os.path.join(_SCRIPT_DIR, 'signoff_runner.log')
LOCKFILE_PATH     = os.path.join(_SCRIPT_DIR, 'signoff.lock')
SIGNOFF_DATA_PATH = os.path.join(_SCRIPT_DIR, 'signoff_data.json')
DASHBOARD_PATH    = os.path.join(_SCRIPT_DIR, 'signoff_dashboard.html')

# ─── QRADAR CREDENTIALS ──────────────────────────────────────────────────────
QRADAR_HOST     = os.environ.get('QRADAR_HOST',     'https://your-qradar-host')
QRADAR_USERNAME = os.environ.get('QRADAR_USERNAME', 'your-username')
QRADAR_PASSWORD = os.environ.get('QRADAR_PASSWORD', 'your-password')
VERIFY_SSL      = False

# ─── SUBJECT MATCHING ────────────────────────────────────────────────────────
SUBJECT_KEYWORD   = 'Security Signoff'
SUBJECT_SEPARATOR = '|'

# ─── SCAN WINDOWS ────────────────────────────────────────────────────────────
LOOKBACK_DAYS              = 30
SENT_SCAN_DAYS             = 90
REVALIDATION_COOLDOWN_DAYS = 0
ACTIVE_SKIP_DAYS           = 30

# ─── SENDER / DL GUARDS ──────────────────────────────────────────────────────
ALLOWED_SENDERS    = []
YOUR_EMAIL_ADDRESS = 'youremail@yourorg.com'
TRIGGER_DL         = '@SOC-DL@yourorg.com'

# ─── ESCALATION ──────────────────────────────────────────────────────────────
ESCALATION_TO      = ['onboarding-owner@yourorg.com']
ESCALATION_CC      = ['@SOC-DL@yourorg.com']
ESCALATION_CONTACT = '@xyz'

# ─── OUTLOOK FOLDERS ─────────────────────────────────────────────────────────
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'

# ─── OS TYPE VALIDATION ──────────────────────────────────────────────────────
OS_TYPE_GROUPS = {
    'Windows': {'required': ['Microsoft Security', 'WinCollect']},
    'Linux':   {'required': ['Linux OS']},
}

# ─── SUBJECT OUTCOME TAGS ────────────────────────────────────────────────────
TAG_ACTIVE     = '[Processed-Active]'
TAG_PARTIAL    = '[Processed-Partial]'
TAG_NOT_FOUND  = '[Processed-NotFound]'
REVALIDATABLE_TAGS = {TAG_PARTIAL, TAG_NOT_FOUND}

# ─── CONSTANTS ───────────────────────────────────────────────────────────────
ACTIVITY_THRESHOLD_DAYS = 7
_MIN_TS                 = 0
_MAX_TS                 = 2147483647
LOG_SOURCE_TYPES_CACHE  = {}
STATUS_PRIORITY         = {'not_found': 2, 'partial': 1, 'active': 0}
REQUEST_TIMEOUT         = 30

_runtime_drafted_hosts: set = set()

_OVERRIDES_PATH = SIGNOFF_DATA_PATH.replace('.json', '_overrides.json')
VERSION         = '3.1'
_API_RETRIES    = 3
_API_PAGE_SIZE  = 50


# ═══════════════════════════════════════════════════════════════════════════════
#  LOGGING & LOCKFILE
# ═══════════════════════════════════════════════════════════════════════════════

def _log(msg: str) -> None:
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(RUN_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except (IOError, PermissionError) as e:
        print(f"  [log-write-fail: {e}]")


def acquire_lock() -> bool:
    if os.path.exists(LOCKFILE_PATH):
        _log("WARN: Lockfile exists — another instance may be running. Exiting.")
        return False
    try:
        with open(LOCKFILE_PATH, 'w') as f:
            f.write(str(os.getpid()))
        return True
    except Exception as e:
        _log(f"ERROR: Could not create lockfile: {e}")
        return False


def release_lock() -> None:
    try:
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
    except Exception as e:
        _log(f"WARN: Could not remove lockfile: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
#  DATETIME HELPER
# ═══════════════════════════════════════════════════════════════════════════════

def _com_dt_to_py(com_dt) -> datetime | None:
    if com_dt is None:
        return None
    try:
        return datetime(
            com_dt.year, com_dt.month, com_dt.day,
            com_dt.hour, com_dt.minute, com_dt.second,
        )
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════════════════════════
#  DATA FILE
# ═══════════════════════════════════════════════════════════════════════════════

def _load_data() -> dict:
    if not os.path.exists(SIGNOFF_DATA_PATH):
        return {"records": [], "overrides": {}}
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, list):
            _log("  INFO: Legacy data format detected — migrating to v3 schema")
            return {"records": data, "overrides": {}}
        if "records" not in data:
            data["records"] = []
        if "overrides" not in data:
            data["overrides"] = {}
        return data
    except Exception as e:
        _log(f"WARN: Could not load data file: {e}")
        return {"records": [], "overrides": {}}


def _save_data(data: dict) -> None:
    try:
        with open(SIGNOFF_DATA_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False, default=str)
    except Exception as e:
        _log(f"WARN: Could not write data file: {e}")


def append_record(record: dict) -> None:
    data = _load_data()
    data["records"].append(record)
    _save_data(data)


def load_overrides() -> dict:
    if os.path.exists(_OVERRIDES_PATH):
        try:
            with open(_OVERRIDES_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            _log(f"WARN: Could not read overrides file: {e}")
    return _load_data().get("overrides", {})


# ═══════════════════════════════════════════════════════════════════════════════
#  CROSS-RUN DEDUPLICATION
# ═══════════════════════════════════════════════════════════════════════════════

def _was_active_recently(hostname_frozenset: frozenset) -> tuple[bool, str]:
    if ACTIVE_SKIP_DAYS <= 0:
        return False, ''
    cutoff = datetime.now() - timedelta(days=ACTIVE_SKIP_DAYS)
    data   = _load_data()
    for rec in reversed(data["records"]):
        try:
            rec_ts  = datetime.fromisoformat(rec.get("timestamp", ""))
            rec_hn  = frozenset(h["hostname"] for h in rec.get("host_results", []))
            rec_sta = rec.get("overall_status", "")
        except Exception:
            continue
        if rec_hn == hostname_frozenset and rec_sta == "active" and rec_ts >= cutoff:
            return True, rec_ts.strftime('%Y-%m-%d')
    return False, ''


# ═══════════════════════════════════════════════════════════════════════════════
#  QRADAR API — RETRY + PAGINATION
# ═══════════════════════════════════════════════════════════════════════════════

def _qradar_get(path: str, params: dict | None = None,
                paginate: bool = False) -> tuple[int, list | dict]:
    url  = f"{QRADAR_HOST.rstrip('/')}{path}"
    hdrs = {'Accept': 'application/json', 'Version': '14.0'}
    auth = (QRADAR_USERNAME, QRADAR_PASSWORD)

    def _one_get(extra: dict) -> requests.Response | None:
        merged = {**hdrs, **extra}
        for attempt in range(1, _API_RETRIES + 1):
            try:
                r = requests.get(url, params=params, auth=auth,
                                 verify=VERIFY_SSL, timeout=REQUEST_TIMEOUT,
                                 headers=merged)
                if r.status_code < 500:
                    return r
                wait = 2 ** attempt
                _log(f"   QRadar HTTP {r.status_code} — retry {attempt}/{_API_RETRIES} in {wait}s")
                time.sleep(wait)
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                wait = 2 ** attempt
                _log(f"   QRadar error: {e} — retry {attempt}/{_API_RETRIES} in {wait}s")
                time.sleep(wait)
        return None

    if not paginate:
        r = _one_get({})
        if r is None:
            return 503, []
        try:
            return r.status_code, r.json()
        except Exception:
            return r.status_code, []

    all_items: list = []
    start = 0
    while True:
        r = _one_get({'Range': f'items={start}-{start + _API_PAGE_SIZE - 1}'})
        if r is None or r.status_code not in (200, 206):
            code = r.status_code if r else 503
            _log(f"   WARN: Pagination stopped at offset {start} — HTTP {code}")
            return code, all_items
        try:
            page = r.json()
        except Exception:
            break
        if not page:
            break
        all_items.extend(page)
        cr = r.headers.get('Content-Range', '')
        if cr:
            try:
                total = int(cr.split('/')[-1])
                if start + _API_PAGE_SIZE >= total:
                    break
            except Exception:
                pass
        if len(page) < _API_PAGE_SIZE:
            break
        start += _API_PAGE_SIZE
    return 200, all_items


def test_qradar_connection() -> bool:
    _log("Testing QRadar connection...")
    code, _ = _qradar_get('/api/help/versions')
    if code == 200:
        _log("QRadar connection OK.")
        return True
    if code == 401:
        _log("ERROR: Authentication failed. Check QRADAR_USERNAME / QRADAR_PASSWORD env vars.")
        return False
    _log(f"WARN: Unexpected response HTTP {code}")
    return False


def fetch_log_source_types() -> None:
    _log("Fetching Log Source Types (paginated)...")
    code, data = _qradar_get(
        '/api/config/event_sources/log_source_management/log_source_types',
        paginate=True,
    )
    if code == 200 and isinstance(data, list):
        for t in data:
            ls_id, ls_name = t.get('id'), t.get('name')
            if ls_id is not None and ls_name is not None:
                LOG_SOURCE_TYPES_CACHE[ls_id] = ls_name
        _log(f"Cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
    else:
        _log(f"WARN: Failed to fetch Log Source Types — HTTP {code}")


# ═══════════════════════════════════════════════════════════════════════════════
#  QRADAR QUERIES — STRICTLY READ-ONLY
# ═══════════════════════════════════════════════════════════════════════════════

def _safe_timestamp(timestamp_ms) -> tuple:
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    try:
        ts = int(timestamp_ms)
        epoch_s = ts / 1000.0 if ts > 4102444800 else ts
        if not (_MIN_TS < epoch_s <= _MAX_TS):
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None
        last_event_dt = datetime.fromtimestamp(epoch_s)
        days_ago      = (datetime.now() - last_event_dt).days
        activity      = 'Active' if days_ago <= ACTIVITY_THRESHOLD_DAYS else 'Inactive'
        return last_event_dt.strftime('%Y-%m-%d %H:%M:%S'), activity, days_ago
    except Exception:
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def query_all_log_sources_readonly(hostname: str) -> dict:
    clean = re.sub(r"[\"']", '', hostname).strip()
    code, ls_data = _qradar_get(
        '/api/config/event_sources/log_source_management/log_sources',
        params={'filter': f'name ilike "%{clean}%"'},
        paginate=True,
    )
    if code != 200:
        return {'status': f'API Error {code}', 'sources': []}
    if not ls_data:
        return {'status': 'Not Found', 'sources': []}
    sources = []
    for src in ls_data:
        type_id      = src.get('type_id')
        ls_type_name = LOG_SOURCE_TYPES_CACHE.get(type_id, f'Unknown Type ID: {type_id}')
        last_seen, activity, days_ago = _safe_timestamp(src.get('last_event_time'))
        sources.append({
            'name':      src.get('name', hostname),
            'ls_type':   ls_type_name,
            'enabled':   src.get('enabled', False),
            'last_seen': last_seen,
            'activity':  activity,
            'days_ago':  days_ago,
        })
    return {'status': 'Found', 'sources': sources}


def detect_os_group(sources: list) -> tuple:
    if not OS_TYPE_GROUPS:
        return None, None
    for group_name, rules in OS_TYPE_GROUPS.items():
        required = rules.get('required', [])
        if not required:
            continue
        sig_words = str(required[0]).lower().split()
        if any(
            all(w in str(s.get('ls_type', '')).lower() for w in sig_words)
            for s in sources
        ):
            return group_name, rules
    return None, None


def validate_expected_types(all_sources_result: dict, required_types: list) -> list:
    results = []
    sources = all_sources_result.get('sources', [])
    for expected_kw in required_types:
        exp_words = str(expected_kw).lower().split()
        matched   = [
            s for s in sources
            if all(w in str(s.get('ls_type', '')).lower() for w in exp_words)
        ]
        if not matched:
            results.append({
                'expected': expected_kw, 'found': False,
                'ls_type': None, 'ls_name': None,
                'last_seen': None, 'days_ago': None,
            })
            continue
        me = sorted([s for s in matched if s.get('enabled')],
                    key=lambda x: (x.get('days_ago') is None, x.get('days_ago') or 99999))
        md = sorted([s for s in matched if not s.get('enabled')],
                    key=lambda x: (x.get('days_ago') is None, x.get('days_ago') or 99999))
        best = me[0] if me else md[0]
        results.append({
            'expected':  expected_kw, 'found': True,
            'ls_type':   best.get('ls_type'), 'ls_name': best.get('name'),
            'last_seen': best.get('last_seen'), 'days_ago': best.get('days_ago'),
        })
    return results


# ═══════════════════════════════════════════════════════════════════════════════
#  SUBJECT PARSING
# ═══════════════════════════════════════════════════════════════════════════════

def passes_subject_guards(subject: str) -> tuple:
    if not subject:
        return False, "empty subject"
    s, sl = subject.strip(), subject.strip().lower()
    if any(sl.startswith(p) for p in ('re:', 'fw:', 'fwd:')):
        return False, f"reply/forward prefix: '{s[:30]}'"
    if '[processed' in sl:
        return False, "subject bears an outcome tag"
    if SUBJECT_SEPARATOR not in s:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found"
    left = s.split(SUBJECT_SEPARATOR)[0].strip().lower()
    if SUBJECT_KEYWORD.lower() not in left:
        return False, f"keyword '{SUBJECT_KEYWORD}' not found left of separator"
    return True, "ok"


def extract_hostnames(subject: str) -> list:
    parts  = subject.split(SUBJECT_SEPARATOR)
    seen, result = set(), []
    for p in parts[1:]:
        hn = ' '.join(p.split()).strip()
        if hn and hn not in seen:
            seen.add(hn)
            result.append(hn)
    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  SENDER VALIDATION
# ═══════════════════════════════════════════════════════════════════════════════

def is_sender_allowed(sender_address: str) -> bool:
    if not ALLOWED_SENDERS:
        return True
    if not sender_address:
        return False
    sc = sender_address.strip().lower()
    for entry in ALLOWED_SENDERS:
        ec = entry.strip().lower()
        if ec.startswith('@') and sc.endswith(ec):
            return True
        if ec == sc:
            return True
    return False


# ═══════════════════════════════════════════════════════════════════════════════
#  BODY DL CHECK
# ═══════════════════════════════════════════════════════════════════════════════

def body_contains_dl(mail_item) -> bool:
    if not TRIGGER_DL.strip():
        return True
    dl_lower = TRIGGER_DL.strip().lower()
    try:
        if dl_lower in (mail_item.Body or '').lower():
            return True
        if dl_lower in (mail_item.HTMLBody or '').lower():
            return True
        return False
    except Exception as e:
        _log(f"   WARN: Body DL check failed: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════════
#  CONVERSATION STATUS
# ═══════════════════════════════════════════════════════════════════════════════

def check_conversation_status(mail_item, sent_folder, drafts_folder) -> tuple:
    try:
        conv_id = mail_item.ConversationID
    except Exception:
        return None, None

    last_tag = None
    last_dt  = None

    def _extract_tag(subject: str) -> str | None:
        s = (subject or '').lower()
        if TAG_NOT_FOUND.lower() in s: return TAG_NOT_FOUND
        if TAG_PARTIAL.lower()   in s: return TAG_PARTIAL
        if TAG_ACTIVE.lower()    in s: return TAG_ACTIVE
        if '[processed]'         in s: return 'legacy'
        return None

    def _update(tag, item_dt):
        nonlocal last_tag, last_dt
        if tag and (last_dt is None or (item_dt and item_dt > last_dt)):
            last_tag, last_dt = tag, item_dt

    def _scan_folder(folder, items_iter, dt_attr: str):
        for item in items_iter:
            try:
                if item.ConversationID == conv_id:
                    _update(_extract_tag(item.Subject),
                            _com_dt_to_py(getattr(item, dt_attr, None)))
            except Exception:
                continue

    cutoff_iso = (datetime.now() - timedelta(days=SENT_SCAN_DAYS)
                  ).strftime('%Y-%m-%d')
    try:
        restricted = sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff_iso}'")
        _scan_folder(sent_folder, restricted, 'SentOn')
    except Exception:
        try:
            _scan_folder(sent_folder, sent_folder.Items, 'SentOn')
        except Exception as e:
            _log(f"   WARN: Could not scan Sent Items: {e}")

    try:
        _scan_folder(drafts_folder, drafts_folder.Items, 'LastModificationTime')
    except Exception as e:
        _log(f"   WARN: Could not scan Drafts: {e}")

    return last_tag, last_dt


# ═══════════════════════════════════════════════════════════════════════════════
#  HTML REPLY BUILDER — EMAIL  (v3.1: transparent background, Outlook dark-mode native)
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
#  HTML REPLY BUILDER — EMAIL  (v3.2)
#
#  Design principles:
#   • NO background on <body> or outer container — transparent so the email
#     sits flush on Outlook's dark-mode canvas (no floating card artefact)
#   • All text is solid and vivid — no muted rgba washes
#   • Status banners use deep solid fills (readable on ANY background)
#   • Host headers have a prominent 4px coloured left-border + solid colour tag
#   • Table rows: high-contrast coloured icons + text, visible dividers
#   • Footer text is intentionally softer (purposeful hierarchy, not accidental haze)
# ═══════════════════════════════════════════════════════════════════════════════

def _host_section_html(hostname: str, host_status: str,
                        type_validation, os_group, sources: list) -> str:

    # ── Palette ────────────────────────────────────────────────────────────────
    SOLID  = {'active': '#22c55e', 'partial': '#f59e0b', 'not_found': '#ef4444'}
    DEEP   = {'active': '#14532d', 'partial': '#78350f', 'not_found': '#7f1d1d'}
    LABEL  = {'active': 'Active',  'partial': 'Partial', 'not_found': 'Not Found'}

    col   = SOLID.get(host_status, '#888888')
    deep  = DEEP.get(host_status,  '#1a1a1a')
    lbl   = LABEL.get(host_status, host_status)

    # ── Host header bar ────────────────────────────────────────────────────────
    os_badge = (
        f'<span style="font-family:Arial,sans-serif;font-size:10px;color:#94a3b8;'
        f'margin-left:10px;font-weight:400">{os_group}</span>'
    ) if os_group else ''

    status_pill = (
        f'<span style="background:{deep};color:{col};font-size:10px;font-weight:700;'
        f'letter-spacing:.4px;padding:2px 9px;border-radius:3px;'
        f'margin-left:10px;text-transform:uppercase;'
        f'border:1px solid {col}">{lbl}</span>'
    )

    hdr = (
        f'<div style="margin:22px 0 0;padding:10px 16px;'
        f'border-left:4px solid {col};">'
        f'<span style="font-family:Consolas,\'Courier New\',monospace;'
        f'font-size:13px;font-weight:700;color:{col}">{hostname}</span>'
        f'{status_pill}{os_badge}'
        f'</div>'
    )

    # ── Not found ──────────────────────────────────────────────────────────────
    if host_status == 'not_found':
        return hdr + (
            f'<div style="padding:8px 16px 4px;border-left:4px solid {col};">'
            f'<span style="font-family:Arial,sans-serif;font-size:12px;'
            f'color:#ef4444;font-weight:600">'
            f'&#10006;&nbsp; Not found in QRadar log source inventory.</span>'
            f'</div>'
        )

    # ── Shared table helpers ───────────────────────────────────────────────────
    TH = lambda t: (
        f'<th style="padding:7px 12px;border-bottom:1px solid #374151;'
        f'text-align:left;color:#94a3b8;font-size:10px;font-family:Arial,sans-serif;'
        f'font-weight:700;text-transform:uppercase;letter-spacing:.5px;'
        f'background:#111827">{t}</th>'
    )

    TABLE_OPEN = (
        f'<div style="border-left:4px solid {col};margin-bottom:4px">'
        f'<table style="width:100%;border-collapse:collapse;'
        f'border:1px solid #374151;font-family:Arial,sans-serif;font-size:12px">'
        f'<thead><tr><th style="width:28px;background:#111827;'
        f'border-bottom:1px solid #374151"></th>'
        + TH('Log Source Type') + TH('Log Source Name')
        + TH('Last Event') + TH('Status')
        + '</tr></thead><tbody>'
    )
    TABLE_CLOSE = '</tbody></table></div>'

    # ── Type-validation mode ───────────────────────────────────────────────────
    if type_validation is not None:
        rows = ''
        for r in type_validation:
            da_str = ''
            if r['days_ago'] is not None:
                label  = 'Today' if r['days_ago'] == 0 else f"{r['days_ago']}d ago"
                da_str = (
                    f'&nbsp;<span style="color:#6b7280;font-size:10px">({label})</span>'
                )

            if not r['found']:
                note = (
                    f"{ESCALATION_CONTACT}&nbsp;please onboard this log source."
                    if ESCALATION_CONTACT.strip() else
                    "Not found — please onboard this log source."
                )
                rows += (
                    f'<tr style="background:#1c0a0a">'
                    f'<td style="text-align:center;padding:8px 4px;'
                    f'border-bottom:1px solid #374151;color:#ef4444;font-size:14px">&#10006;</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#ef4444;font-weight:700">{r["expected"]}</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#6b7280">&#8212;</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#6b7280">&#8212;</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#ef4444;font-style:italic">{note}</td>'
                    f'</tr>'
                )
            elif r['days_ago'] is None:
                note = (
                    f"{ESCALATION_CONTACT}&nbsp;no events received — please investigate."
                    if ESCALATION_CONTACT.strip() else
                    "No events received — please investigate."
                )
                rows += (
                    f'<tr style="background:#1c1205">'
                    f'<td style="text-align:center;padding:8px 4px;'
                    f'border-bottom:1px solid #374151;color:#f59e0b;font-size:14px">&#9888;</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#f59e0b;font-weight:700">{r["expected"]}</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#d1d5db">{r.get("ls_name","N/A")}</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#f59e0b;font-style:italic">No events recorded</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#f59e0b;font-style:italic">{note}</td>'
                    f'</tr>'
                )
            else:
                rows += (
                    f'<tr style="background:#0a1c10">'
                    f'<td style="text-align:center;padding:8px 4px;'
                    f'border-bottom:1px solid #374151;color:#22c55e;font-size:14px">&#10004;</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#ffffff;font-weight:700">{r["expected"]}</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#d1d5db">{r.get("ls_name","N/A")}</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#d1d5db">{r.get("last_seen","N/A")}{da_str}</td>'
                    f'<td style="padding:8px 12px;border-bottom:1px solid #374151;'
                    f'color:#22c55e;font-weight:700">&#10004;&nbsp;Confirmed</td>'
                    f'</tr>'
                )
        return hdr + TABLE_OPEN + rows + TABLE_CLOSE

    # ── Simple mode (no OS group match) ───────────────────────────────────────
    best = None
    if sources:
        en = sorted(
            [s for s in sources if s.get('enabled')],
            key=lambda x: (x.get('days_ago') is None, x.get('days_ago') or 99999)
        )
        di = sorted(
            [s for s in sources if not s.get('enabled')],
            key=lambda x: (x.get('days_ago') is None, x.get('days_ago') or 99999)
        )
        best = en[0] if en else (di[0] if di else None)

    if best:
        da  = best.get('days_ago')
        dsp = 'Today' if da == 0 else (f"{da}d ago" if da is not None else 'N/A')
        dcolor = '#22c55e' if (da is not None and da <= 7) else '#f59e0b'

        def simple_row(label, value, val_style='color:#d1d5db'):
            return (
                f'<tr>'
                f'<td style="padding:7px 14px;color:#94a3b8;font-size:11px;'
                f'font-family:Arial,sans-serif;border-bottom:1px solid #1f2937;'
                f'width:130px;font-weight:600;text-transform:uppercase;letter-spacing:.4px">'
                f'{label}</td>'
                f'<td style="padding:7px 14px;font-size:12px;font-family:Arial,sans-serif;'
                f'border-bottom:1px solid #1f2937;{val_style}">{value}</td>'
                f'</tr>'
            )

        return hdr + (
            f'<div style="border-left:4px solid {col};margin-bottom:4px">'
            f'<table style="width:100%;max-width:520px;border-collapse:collapse;'
            f'border:1px solid #374151;background:#0d1117">'
            + simple_row('Log Source',
                         f'<strong style="color:#ffffff;font-size:13px">'
                         f'{best.get("name","N/A")}</strong>')
            + simple_row('Source Type', best.get("ls_type","N/A"))
            + simple_row('Last Event',
                         f'<span style="color:#d1d5db">{best.get("last_seen","N/A")}</span>'
                         f'&nbsp;<span style="color:{dcolor};font-weight:700;font-size:11px">'
                         f'({dsp})</span>')
            + '</table></div>'
        )

    return hdr


def build_reply_for_all_hosts(hostname_qr_pairs: list) -> tuple:
    """
    v3.2 email design:
     - Transparent body (no background) — blends with Outlook dark mode
     - Solid vivid status banner with deep coloured fill + bright accent border
     - White body text — fully solid, not hazy
     - Per-host sections with prominent 4px left-border + status pill badge
     - High-contrast table rows (dark tinted rows, vivid status colours)
    """
    run_time = datetime.now().strftime('%d %B %Y, %H:%M')
    host_sections, host_tracking, statuses = [], [], []

    for hostname, qr_result in hostname_qr_pairs:
        host_status, type_validation, os_group = _status_for_host(hostname, qr_result)
        sources = qr_result.get('sources', [])
        statuses.append(host_status)
        host_sections.append(
            _host_section_html(hostname, host_status, type_validation, os_group, sources)
        )
        best_da = best_seen = None
        if sources:
            ev = [s for s in sources if s.get('enabled') and s.get('days_ago') is not None]
            if ev:
                best      = min(ev, key=lambda x: x['days_ago'])
                best_da   = best['days_ago']
                best_seen = best['last_seen']
        host_tracking.append({
            'hostname':  hostname,
            'status':    host_status,
            'os_group':  os_group,
            'last_seen': best_seen,
            'days_ago':  best_da,
        })

    # FIX from v3.0: max() with STATUS_PRIORITY (not_found=2 > partial=1 > active=0)
    overall_status = max(statuses, key=lambda s: STATUS_PRIORITY.get(s, 0)) \
        if statuses else 'not_found'

    n_hosts  = len(hostname_qr_pairs)
    n_ok     = sum(1 for s in statuses if s == 'active')
    n_issues = n_hosts - n_ok

    # ── Banner config: deep solid fill + vivid top-border ─────────────────────
    BANNERS = {
        'active': {
            'bg':    '#14532d',          # deep solid green
            'bdr':   '#22c55e',          # vivid green border
            'icon':  '&#10004;',
            'text':  '#ffffff',
            'label': 'All Hosts Confirmed Reporting on SIEM',
        },
        'partial': {
            'bg':    '#78350f',
            'bdr':   '#f59e0b',
            'icon':  '&#9888;',
            'text':  '#ffffff',
            'label': (f'{n_issues} of {n_hosts} Host{"s" if n_hosts > 1 else ""}'
                      f' Require Attention'),
        },
        'not_found': {
            'bg':    '#7f1d1d',
            'bdr':   '#ef4444',
            'icon':  '&#10006;',
            'text':  '#ffffff',
            'label': (f'{"Some hosts not" if n_ok else "No hosts"} found in QRadar'),
        },
    }
    B = BANNERS.get(overall_status, BANNERS['partial'])

    # ── Summary sentence ───────────────────────────────────────────────────────
    if overall_status == 'active' and n_hosts == 1:
        summary = (f'<b style="color:#ffffff">{hostname_qr_pairs[0][0]}</b>'
                   f' is confirmed reporting on our SIEM.')
    elif overall_status == 'active':
        names   = ', '.join(
            f'<b style="color:#ffffff">{h}</b>' for h, _ in hostname_qr_pairs
        )
        summary = f'All {n_hosts} hosts ({names}) are confirmed active on our SIEM.'
    else:
        summary = (
            f'<b style="color:#ffffff">{n_ok} of {n_hosts}'
            f' host{"s" if n_hosts > 1 else ""}</b> confirmed active. '
            f'Issues are highlighted per host below.'
        )

    # ── Assemble email ────────────────────────────────────────────────────────
    #    body background intentionally omitted — renders transparent on
    #    Outlook's dark-mode canvas
    body = (
        '<html>'
        '<head>'
        '<meta name="color-scheme" content="dark light">'
        '<meta name="supported-color-scheme" content="dark light">'
        '</head>'
        '<body style="'
        'font-family:Arial,\'Segoe UI\',sans-serif;'
        'font-size:13px;'
        'line-height:1.6;'
        'color:#ffffff;'          # solid white — no haze
        'margin:0;padding:16px 0;">'

        '<div style="max-width:700px;">'

        # Greeting
        '<p style="margin:0 0 18px;font-size:13px;color:#ffffff">Hi,</p>'

        # ── Status banner ──────────────────────────────────────────────────────
        f'<div style="'
        f'background:{B["bg"]};'
        f'border:1px solid {B["bdr"]};'
        f'border-left:5px solid {B["bdr"]};'
        f'border-radius:4px;'
        f'padding:12px 18px;'
        f'margin-bottom:16px;">'
        f'<span style="font-size:14px;font-weight:700;color:{B["text"]}">'
        f'{B["icon"]}&nbsp;&nbsp;{B["label"]}'
        f'</span>'
        f'</div>'

        # Summary line
        f'<p style="margin:0 0 6px;font-size:13px;color:#d1d5db">{summary}</p>'

        # Per-host sections
        + ''.join(host_sections) +

        # ── Footer ────────────────────────────────────────────────────────────
        '<div style="margin-top:28px;padding-top:14px;border-top:1px solid #374151">'
        f'<p style="margin:0;font-size:11px;color:#6b7280">'
        f'Automated SIEM monitoring response &mdash; QRadar checked {run_time}.</p>'
        '</div>'

        '<p style="margin:18px 0 0;font-size:13px;color:#ffffff">'
        'Regards,<br>'
        '<strong style="color:#ffffff">Cyberdefence</strong>'
        '</p>'

        '</div>'
        '</body></html>'
    )

    return body, overall_status, host_tracking

def _status_for_host(hostname: str, qr_result: dict) -> tuple:
    status  = qr_result.get('status')
    sources = qr_result.get('sources', [])
    if status != 'Found' or not sources:
        return 'not_found', None, None
    if not OS_TYPE_GROUPS:
        return 'active', None, None
    group_name, group_rules = detect_os_group(sources)
    if group_name is None:
        _log(f"      OS group undetected for {hostname} — raw found, no type validation")
        return 'active', None, None
    validation  = validate_expected_types(qr_result, group_rules.get('required', []))
    any_problem = any(
        not r['found'] or (r['found'] and r['days_ago'] is None)
        for r in validation
    )
    return ('partial' if any_problem else 'active'), validation, group_name


# ═══════════════════════════════════════════════════════════════════════════════
#  DRAFT CREATOR
# ═══════════════════════════════════════════════════════════════════════════════

def create_draft_reply(mail_item, html_body: str, overall_status: str,
                       hostnames_str: str, is_revalidation: bool = False) -> bool:
    tag_map = {'active': TAG_ACTIVE, 'partial': TAG_PARTIAL, 'not_found': TAG_NOT_FOUND}
    tag    = tag_map.get(overall_status, TAG_ACTIVE)
    prefix = '[Revalidated] ' if is_revalidation else ''

    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"{prefix}{tag} {mail_item.Subject}"

        use_escalation = (
            overall_status in ('partial', 'not_found')
            and (ESCALATION_TO or ESCALATION_CC)
        )
        if use_escalation:
            reply.To = '; '.join(ESCALATION_TO) if ESCALATION_TO else ''
            reply.CC = '; '.join(ESCALATION_CC) if ESCALATION_CC else ''
            _log(f"      Escalation routing  To: {reply.To or '(none)'}  CC: {reply.CC or '(none)'}")
        else:
            _log(f"      ReplyAll (Active or escalation lists empty)")

        reply.Save()
        _log(f"      Draft saved [{tag}] — {hostnames_str}"
             f"{' [revalidation]' if is_revalidation else ''}")
        return True

    except Exception as e:
        _log(f"      ERROR: Draft creation failed for {hostnames_str}: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════════
#  OUTLOOK SETUP
# ═══════════════════════════════════════════════════════════════════════════════

def get_outlook_folders() -> tuple:
    try:
        outlook    = win32com.client.Dispatch('Outlook.Application')
        ns         = outlook.GetNamespace('MAPI')
        main_inbox = ns.GetDefaultFolder(6)
        drafts     = ns.GetDefaultFolder(16)
        sent       = ns.GetDefaultFolder(5)

        if SIGNOFF_FOLDER_NAME:
            try:
                inbox = main_inbox.Folders[SIGNOFF_FOLDER_NAME]
                _log(f"Monitoring: Inbox\\{SIGNOFF_FOLDER_NAME}")
            except Exception:
                _log(f"WARN: Subfolder '{SIGNOFF_FOLDER_NAME}' not found — falling back to Inbox.")
                inbox = main_inbox
        else:
            inbox = main_inbox
            _log("Monitoring: Full Inbox")

        return inbox, drafts, sent
    except Exception as e:
        _log(f"ERROR: Could not connect to Outlook: {e}")
        return None, None, None


# ═══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD GENERATOR  (v3.1: auto-opens on run completion)
# ═══════════════════════════════════════════════════════════════════════════════

def generate_dashboard() -> None:
    data = _load_data()
    try:
        html = _build_dashboard_html(data["records"])
        with open(DASHBOARD_PATH, 'w', encoding='utf-8') as f:
            f.write(html)
        _log(f"Dashboard written → {DASHBOARD_PATH}  ({len(data['records'])} records)")
    except Exception as e:
        _log(f"WARN: Dashboard generation failed: {e}")
        return

    # v3.1: auto-open dashboard in the default browser
    try:
        os.startfile(DASHBOARD_PATH)
        _log("Dashboard opened in browser.")
    except Exception as e:
        _log(f"INFO: Could not auto-open dashboard: {e}")


def _build_dashboard_html(records: list) -> str:
    data_json  = json.dumps(records, ensure_ascii=False, default=str)
    generated  = datetime.now().strftime('%d %B %Y at %H:%M')
    rcd        = str(REVALIDATION_COOLDOWN_DAYS)
    rwd        = str(LOOKBACK_DAYS)
    return (
        _DASHBOARD_TMPL
        .replace('%%DATA_JSON%%',  data_json)
        .replace('%%GENERATED%%',  generated)
        .replace('%%REVAL_COOL%%', rcd)
        .replace('%%LOOKBACK%%',   rwd)
        .replace('%%VERSION%%',    VERSION)
        .replace('%%OVPATH%%',     _OVERRIDES_PATH)
    )


# ─── Dashboard HTML template ──────────────────────────────────────────────────
# v3.1 changes:
#   • Status filter pills REMOVED from table header
#   • OS-group dropdown REMOVED from table header
#   • Quick-delete (🗑) button added to every row — no modal, 10-second undo
#   • Cleaner card design: gradient accent top-border, tighter spacing
#   • Table: sticky header, subtle alternating rows, sharper type
#   • Modal kept for full edit/override/note workflow
# ─────────────────────────────────────────────────────────────────────────────
_DASHBOARD_TMPL = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>SIEM Signoff Dashboard v%%VERSION%%</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg0:#07111e;--bg1:#0d1929;--bg2:#112236;--bg3:#162b42;
  --bdr:#1e3352;--bdr2:#162544;
  --t0:#f0f6ff;--t1:#c5d4ee;--t2:#7e9cbf;--t3:#435e7a;
  --green:#22c55e;--gd:#15803d;--gbg:#071812;
  --amber:#f59e0b;--ad:#b45309;--abg:#1a1106;
  --red:#f87171;--rd:#b91c1c;--rbg:#1a0808;
  --blue:#60a5fa;--bd:#1d4ed8;--bbg:#06122a;
  --purple:#a78bfa;--pbg:#130f28;
  --mono:'Consolas','Cascadia Code','Courier New',monospace;
  --r:10px;--r2:6px;
  --shadow:0 1px 3px rgba(0,0,0,.4),0 4px 16px rgba(0,0,0,.3);
}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg0);color:var(--t1);font-size:13px;line-height:1.5;min-height:100vh}

/* ── Header ── */
.hdr{
  background:var(--bg1);
  border-bottom:1px solid var(--bdr);
  padding:11px 24px;
  display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:100;gap:12px;flex-wrap:wrap;
}
.hdr-l{display:flex;align-items:center;gap:10px}
.hdr h1{font-size:14px;font-weight:600;color:var(--t0);letter-spacing:-.01em}
.badge{font-size:10px;padding:2px 8px;border-radius:20px;font-weight:600;letter-spacing:.3px}
.bqr{background:var(--pbg);color:var(--purple);border:1px solid #4c2d8a}
.bv {background:var(--bbg);color:var(--blue);border:1px solid #1e3a7a}
.hdr-r{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.hdr-ts{font-size:11px;color:var(--t3)}
.btn{
  background:var(--bg2);border:1px solid var(--bdr);color:var(--t2);
  padding:5px 13px;border-radius:var(--r2);cursor:pointer;
  font-size:12px;font-family:inherit;transition:all .14s;white-space:nowrap;
}
.btn:hover{border-color:var(--blue);color:var(--t0);background:var(--bg3)}
.btn-sm{padding:4px 10px;font-size:11px}
.btn-pri{background:var(--bbg);border-color:#1e4080;color:var(--blue)}
.btn-pri:hover{background:var(--blue);color:#fff;border-color:var(--blue)}
.btn-del{
  background:transparent;border:1px solid transparent;
  color:var(--t3);padding:3px 7px;border-radius:var(--r2);
  cursor:pointer;font-size:13px;font-family:inherit;
  transition:all .14s;line-height:1;
}
.btn-del:hover{background:var(--rbg);border-color:var(--rd);color:var(--red)}
.btn-edit{
  background:transparent;border:1px solid transparent;
  color:var(--t3);padding:3px 7px;border-radius:var(--r2);
  cursor:pointer;font-size:13px;font-family:inherit;
  transition:all .14s;line-height:1;
}
.btn-edit:hover{background:var(--bbg);border-color:#1e4080;color:var(--blue)}

/* ── Layout ── */
.main{padding:18px 24px;max-width:1440px}

/* ── Period bar ── */
.pbar{display:flex;gap:4px;margin-bottom:16px;align-items:center;flex-wrap:wrap}
.pbar .sep{width:1px;height:16px;background:var(--bdr);margin:0 6px}
.pbtn{
  background:transparent;border:1px solid var(--bdr);color:var(--t3);
  padding:4px 13px;border-radius:20px;cursor:pointer;font-size:12px;
  font-family:inherit;transition:all .14s;
}
.pbtn:hover{border-color:var(--blue);color:var(--t1)}
.pbtn.active{background:var(--bbg);border-color:#1e4080;color:var(--blue);font-weight:500}

/* ── KPI Cards ── */
.cards{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:16px}
.card{
  background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);
  padding:14px 16px;position:relative;overflow:hidden;
}
.card::before{
  content:'';position:absolute;top:0;left:0;right:0;height:2px;
  border-radius:var(--r) var(--r) 0 0;
}
.card.cb::before{background:linear-gradient(90deg,var(--blue),#3b82f620)}
.card.cg::before{background:linear-gradient(90deg,var(--green),#22c55e20)}
.card.ca::before{background:linear-gradient(90deg,var(--amber),#f59e0b20)}
.card.cr::before{background:linear-gradient(90deg,var(--red),#f8717120)}
.card.cp::before{background:linear-gradient(90deg,var(--purple),#a78bfa20)}
.card-lbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--t3);margin-bottom:8px}
.card-val{font-size:28px;font-weight:700;line-height:1;color:var(--t0);letter-spacing:-.02em}
.card-sub{font-size:11px;color:var(--t3);margin-top:4px}
.card.cg .card-val{color:var(--green)}
.card.ca .card-val{color:var(--amber)}
.card.cr .card-val{color:var(--red)}
.card.cb .card-val{color:var(--blue)}
.card.cp .card-val{color:var(--purple)}

/* ── Charts ── */
.charts{display:grid;grid-template-columns:220px 1fr;gap:10px;margin-bottom:16px}
.cbox{background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);padding:16px}
.clbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--t3);margin-bottom:12px}
.leg{display:flex;flex-direction:column;gap:7px;margin-top:12px}
.leg-r{display:flex;align-items:center;gap:8px;font-size:11px;color:var(--t2)}
.leg-d{width:8px;height:8px;border-radius:2px;flex-shrink:0}

/* ── Table wrapper ── */
.twrap{background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);overflow:hidden}
.ttbar{
  display:flex;align-items:center;gap:10px;
  padding:12px 16px;border-bottom:1px solid var(--bdr);flex-wrap:wrap;
}
.ttitle{font-size:11px;font-weight:600;color:var(--t0);letter-spacing:-.01em}
.trcount{font-size:11px;color:var(--t3)}
.sp1{flex:1}
.srch{
  background:var(--bg0);border:1px solid var(--bdr);border-radius:var(--r2);
  color:var(--t0);padding:5px 11px;font-size:12px;width:200px;font-family:inherit;
  transition:border-color .14s;
}
.srch:focus{outline:none;border-color:var(--blue)}
.srch::placeholder{color:var(--t3)}

/* ── Table ── */
table{width:100%;border-collapse:collapse}
thead{position:sticky;top:49px;z-index:10}
th{
  background:var(--bg0);color:var(--t3);font-size:10px;font-weight:600;
  text-align:left;padding:9px 13px;border-bottom:1px solid var(--bdr);
  text-transform:uppercase;letter-spacing:.5px;cursor:pointer;
  white-space:nowrap;user-select:none;transition:color .12s;
}
th:hover{color:var(--t1)}
th.sa::after{content:' ▲';color:var(--blue);font-size:9px}
th.sd::after{content:' ▼';color:var(--blue);font-size:9px}
td{padding:10px 13px;border-bottom:1px solid var(--bdr2);vertical-align:middle;font-size:12px}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--bg2)}
tr.drow td{opacity:.35;text-decoration:line-through}
.hn{font-family:var(--mono);font-size:12px;font-weight:500;color:var(--t0)}
.hndel{color:var(--t3)}
.sb{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;letter-spacing:.3px;white-space:nowrap}
.sa1{background:rgba(34,197,94,.1);color:var(--green);border:1px solid var(--gd)}
.sp3{background:rgba(245,158,11,.1);color:var(--amber);border:1px solid var(--ad)}
.sn {background:rgba(248,113,113,.1);color:var(--red);border:1px solid var(--rd)}
.bovr{background:rgba(96,165,250,.1);color:var(--blue);font-size:9px;padding:1px 5px;border-radius:3px;margin-left:5px;font-weight:600;border:1px solid #1e4080}
.bdel{background:var(--rbg);color:var(--red);font-size:9px;padding:1px 5px;border-radius:3px;margin-left:5px;font-weight:600;border:1px solid var(--rd)}
.obadge{background:rgba(167,139,250,.08);color:var(--purple);font-size:10px;padding:2px 8px;border-radius:4px;border:1px solid #4c2d8a;white-space:nowrap}
.tup{color:var(--green);font-size:11px}
.tdn{color:var(--red);font-size:11px}
.teq{color:var(--t3);font-size:11px}
.rvsoon{color:var(--amber);font-weight:600;font-size:11px}
.rvok{color:var(--t3);font-size:11px}
.ract{display:flex;gap:3px;align-items:center}
.note-i{cursor:help;color:var(--amber);font-size:11px;margin-left:4px;vertical-align:middle}
.nodata{text-align:center;padding:44px;color:var(--t3);font-size:13px}

/* ── Pager ── */
.pager{
  display:flex;align-items:center;justify-content:flex-end;gap:6px;
  padding:10px 16px;border-top:1px solid var(--bdr);font-size:12px;color:var(--t2);
}
.pginfo{font-size:11px;color:var(--t3);margin-right:4px}

/* ── Modal ── */
.mbg{position:fixed;inset:0;background:rgba(0,0,0,.7);backdrop-filter:blur(3px);z-index:500;display:none;align-items:center;justify-content:center}
.mbg.open{display:flex}
.modal{background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);padding:24px;width:440px;max-width:94vw;max-height:88vh;overflow-y:auto;box-shadow:var(--shadow)}
.modal h2{font-size:14px;font-weight:600;color:var(--t0);margin-bottom:16px;letter-spacing:-.01em}
.frow{margin-bottom:14px}
label{display:block;font-size:10px;font-weight:600;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px}
select,textarea,input[type=text]{
  background:var(--bg0);border:1px solid var(--bdr);border-radius:var(--r2);
  color:var(--t0);padding:7px 10px;font-size:13px;font-family:inherit;width:100%;
  transition:border-color .14s;
}
select:focus,textarea:focus,input[type=text]:focus{outline:none;border-color:var(--blue)}
textarea{resize:vertical;min-height:68px}
.mact{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}
.btn-mex{background:var(--rbg);border:1px solid var(--rd);color:var(--red);padding:5px 13px;border-radius:var(--r2);cursor:pointer;font-size:12px;font-family:inherit;transition:all .14s}
.btn-mex:hover{background:var(--red);color:#fff}
.err{border-color:var(--red)!important}

/* ── Toast & Undo ── */
.toast{
  position:fixed;bottom:20px;right:20px;
  background:var(--bg2);border:1px solid var(--bdr);border-radius:var(--r);
  padding:9px 15px;font-size:12px;z-index:600;
  box-shadow:var(--shadow);display:none;color:var(--t1);
}
.toast.show{display:block}
.ubar{
  position:fixed;bottom:20px;left:50%;transform:translateX(-50%);
  background:var(--bg2);border:1px solid var(--rd);border-radius:var(--r);
  padding:10px 18px;font-size:12px;z-index:600;
  display:none;align-items:center;gap:12px;
  box-shadow:var(--shadow);white-space:nowrap;color:var(--t1);
}
.ubar.show{display:flex}
.ubtn{color:var(--amber);cursor:pointer;font-weight:600;text-decoration:underline}
.utmr{font-size:11px;color:var(--t3);font-family:var(--mono)}

@media(max-width:900px){.cards{grid-template-columns:repeat(3,1fr)}.charts{grid-template-columns:1fr}}
@media(max-width:600px){.cards{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-l">
    <h1>SIEM Signoff Dashboard</h1>
    <span class="badge bqr">QRadar</span>
    <span class="badge bv">v%%VERSION%%</span>
  </div>
  <div class="hdr-r">
    <span class="hdr-ts">Generated %%GENERATED%%</span>
    <button class="btn btn-sm" onclick="exportCSV()">&#8595; Export CSV</button>
    <button class="btn btn-sm" onclick="exportOvr()">&#8595; Overrides JSON</button>
  </div>
</div>

<div class="main">

  <!-- Period selector -->
  <div class="pbar">
    <button class="pbtn" onclick="setPeriod(7)"  data-p="7">7 days</button>
    <button class="pbtn" onclick="setPeriod(15)" data-p="15">15 days</button>
    <button class="pbtn active" onclick="setPeriod(30)" data-p="30">30 days</button>
    <button class="pbtn" onclick="setPeriod(0)"  data-p="0">All time</button>
    <div class="sep"></div>
    <span style="font-size:11px;color:var(--t3)">Lookback %%LOOKBACK%%d &nbsp;&bull;&nbsp; Reval cooldown %%REVAL_COOL%%d</span>
  </div>

  <!-- KPI cards -->
  <div class="cards">
    <div class="card cb"><div class="card-lbl">Signoff Emails</div><div class="card-val" id="ct0">—</div><div class="card-sub" id="ct0s">—</div></div>
    <div class="card cg"><div class="card-lbl">Active</div><div class="card-val" id="ct1">—</div><div class="card-sub" id="ct1s">—</div></div>
    <div class="card ca"><div class="card-lbl">Partial</div><div class="card-val" id="ct2">—</div><div class="card-sub" id="ct2s">—</div></div>
    <div class="card cr"><div class="card-lbl">Not Found</div><div class="card-val" id="ct3">—</div><div class="card-sub" id="ct3s">—</div></div>
    <div class="card cp"><div class="card-lbl">Exceptions</div><div class="card-val" id="ct4">—</div><div class="card-sub">deleted / overridden</div></div>
  </div>

  <!-- Charts -->
  <div class="charts">
    <div class="cbox">
      <div class="clbl">Status distribution</div>
      <svg id="dsvg" viewBox="0 0 200 170" width="100%" role="img" aria-label="Donut chart"></svg>
      <div class="leg" id="dleg"></div>
    </div>
    <div class="cbox">
      <div class="clbl">Signoffs over time &mdash; <span style="color:var(--red)">not-found</span> &rarr; <span style="color:var(--amber)">partial</span> &rarr; <span style="color:var(--green)">active</span></div>
      <svg id="bsvg" viewBox="0 0 530 185" width="100%" role="img" aria-label="Bar chart" style="overflow:visible"></svg>
    </div>
  </div>

  <!-- Host table -->
  <div class="twrap">
    <div class="ttbar">
      <span class="ttitle">Host Registry</span>
      <span class="trcount" id="trcount"></span>
      <div class="sp1"></div>
      <input class="srch" type="text" placeholder="&#128269; Search hostname…" oninput="setSrch(this.value)" id="srchbox">
    </div>
    <div style="overflow-x:auto">
      <table>
        <thead><tr>
          <th onclick="srt('hostname')">Hostname</th>
          <th onclick="srt('os_group')">OS Group</th>
          <th onclick="srt('eff_status')">Status</th>
          <th onclick="srt('last_checked')">Last Checked</th>
          <th onclick="srt('days_ago')">Last QRadar Event</th>
          <th onclick="srt('checks')" style="text-align:center">Checks</th>
          <th onclick="srt('next_rv')">Next Revalidation</th>
          <th>Trend</th>
          <th style="width:70px"></th>
        </tr></thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
    <div class="pager" id="pager"></div>
  </div>
</div>

<!-- Edit modal -->
<div class="mbg" id="mbg">
  <div class="modal">
    <h2 id="mtitle">Edit host</h2>
    <div class="frow">
      <label>Status override</label>
      <select id="mst">
        <option value="">— use script result —</option>
        <option value="active">Active</option>
        <option value="partial">Partial</option>
        <option value="not_found">Not Found</option>
      </select>
    </div>
    <div class="frow">
      <label>Note <span style="font-weight:400;color:var(--t3)">(shown as tooltip)</span></label>
      <textarea id="mnote" placeholder="Freeform note…"></textarea>
    </div>
    <div class="frow">
      <label>Exception reason <span style="color:var(--red)">*</span> <span style="font-weight:400;color:var(--t3)">required for exception</span></label>
      <textarea id="mreason" placeholder="e.g. Decommissioned — approved by SOC-Lead 2026-01-15"></textarea>
    </div>
    <div class="mact">
      <button class="btn" onclick="closeM()">Cancel</button>
      <button class="btn-mex" onclick="delFromM()" id="mdelbtn">Mark as Exception</button>
      <button class="btn btn-pri" onclick="saveM()">Save</button>
    </div>
  </div>
</div>

<!-- Undo bar -->
<div class="ubar" id="ubar">
  <span id="umsg">Host removed.</span>
  <span class="ubtn" onclick="undoDel()">Undo</span>
  <span class="utmr" id="utmr">10s</span>
</div>

<div class="toast" id="toast"></div>

<script>
// ── Injected data ────────────────────────────────────────────────────────────
const ALL = %%DATA_JSON%%;
const RVCOOL = %%REVAL_COOL%%;
const OVR_PATH = '%%OVPATH%%';

// ── localStorage ─────────────────────────────────────────────────────────────
const LS = 'siem_ovr_v3';
const PS = 25;

function loadOvr()      { try { return JSON.parse(localStorage.getItem(LS)||'{}'); } catch { return {}; } }
function saveOvr(o)     { localStorage.setItem(LS, JSON.stringify(o)); }
function getO(hn)       { return loadOvr()[hn] || {}; }
function setO(hn,patch) { const a=loadOvr(); a[hn]=Object.assign(a[hn]||{},patch,{ts:new Date().toISOString()}); saveOvr(a); }

// ── Status helpers ────────────────────────────────────────────────────────────
const SLBL  = {active:'Active',partial:'Partial',not_found:'Not Found'};
const SCLS  = {active:'sa1',partial:'sp3',not_found:'sn'};
const SRANK = {active:0,partial:1,not_found:2};

// ── State ─────────────────────────────────────────────────────────────────────
let period=30, srch='';
let scol='last_checked', sasc=false;
let cp=1, allH=[], editHN=null;
let undoTmr=null, undoPend=null, undoSecs=10;

// ── Host map ──────────────────────────────────────────────────────────────────
function buildMap(recs) {
  const m = {};
  recs.forEach(r => r.host_results.forEach(h => {
    const k = h.hostname;
    if (!m[k]) m[k] = {hostname:k,os_group:h.os_group||'—',status:h.status,
                       last_checked:r.timestamp,days_ago:h.days_ago,
                       checks:0,prev_status:null};
    const e = m[k];
    if (r.timestamp > e.last_checked) {
      e.prev_status=e.status; e.status=h.status;
      e.last_checked=r.timestamp; e.days_ago=h.days_ago;
      e.os_group=h.os_group||e.os_group||'—';
    }
    e.checks++;
  }));
  return Object.values(m);
}

function applyOvr(hosts) {
  const o = loadOvr();
  return hosts.map(h => {
    const x = o[h.hostname] || {};
    return {...h,deleted:!!x.deleted,status_override:x.status_override||null,
            note:x.note||'',exception_reason:x.exception_reason||'',
            eff_status:x.status_override||h.status};
  });
}

function filtRecs(d) {
  if (!d) return ALL;
  const c = new Date(); c.setDate(c.getDate()-d);
  return ALL.filter(r => new Date(r.timestamp)>=c);
}

// ── Controls ──────────────────────────────────────────────────────────────────
function setPeriod(d) {
  period=d;
  document.querySelectorAll('.pbtn').forEach(b=>b.classList.toggle('active',+b.dataset.p===d));
  cp=1; render();
}
function setSrch(v) { srch=v.toLowerCase(); cp=1; drawTable(); }
function srt(col) {
  if (scol===col) sasc=!sasc; else { scol=col; sasc=(col==='hostname'); }
  document.querySelectorAll('th').forEach(t=>t.classList.remove('sa','sd'));
  const cols=['hostname','os_group','eff_status','last_checked','days_ago','checks','next_rv','_t','_a'];
  const idx=cols.indexOf(col), ths=document.querySelectorAll('th');
  if(idx>=0&&ths[idx]) ths[idx].classList.add(sasc?'sa':'sd');
  drawTable();
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function rel(iso) {
  const d=(Date.now()-new Date(iso))/1000;
  if(d<60) return 'Just now';
  if(d<3600) return Math.floor(d/60)+'m ago';
  if(d<86400) return Math.floor(d/3600)+'h ago';
  const n=Math.floor(d/86400); return n===1?'Yesterday':n+'d ago';
}
function nextRV(h) {
  if (h.deleted) return '<span class="rvok">N/A</span>';
  if (h.eff_status==='active') return '<span class="rvok">—</span>';
  if (RVCOOL===0) return '<span class="rvsoon">Every run</span>';
  const el=(Date.now()-new Date(h.last_checked))/86400000;
  const rem=Math.max(0,RVCOOL-el);
  if (rem<=0) return '<span class="rvsoon">Due now</span>';
  return `<span class="${Math.ceil(rem)<=1?'rvsoon':'rvok'}">in ${Math.ceil(rem)}d</span>`;
}
function trend(h) {
  if (!h.prev_status||h.prev_status===h.status) return '<span class="teq">—</span>';
  return SRANK[h.status]<SRANK[h.prev_status]
    ?'<span class="tup">▲ Improved</span>':'<span class="tdn">▼ Degraded</span>';
}

// ── Main render ───────────────────────────────────────────────────────────────
function render() {
  const recs=filtRecs(period), hosts=applyOvr(buildMap(recs));
  const nd=hosts.filter(h=>!h.deleted);
  const a=nd.filter(h=>h.eff_status==='active').length;
  const p=nd.filter(h=>h.eff_status==='partial').length;
  const n=nd.filter(h=>h.eff_status==='not_found').length;
  const ex=hosts.filter(h=>h.deleted).length;
  const u=nd.length;
  const pct=v=>u?Math.round(v/u*100)+'%':'—';
  document.getElementById('ct0').textContent=recs.length;
  document.getElementById('ct0s').textContent=u+' unique host'+(u!==1?'s':'');
  document.getElementById('ct1').textContent=a; document.getElementById('ct1s').textContent=pct(a)+' of hosts';
  document.getElementById('ct2').textContent=p; document.getElementById('ct2s').textContent=pct(p)+' of hosts';
  document.getElementById('ct3').textContent=n; document.getElementById('ct3s').textContent=pct(n)+' of hosts';
  document.getElementById('ct4').textContent=ex;
  renderDonut(a,p,n); renderBar(recs);
  allH=hosts; drawTable();
}

// ── Donut ─────────────────────────────────────────────────────────────────────
function renderDonut(a,p,n) {
  const tot=a+p+n, sv=document.getElementById('dsvg');
  if (!tot) { sv.innerHTML='<text x="100" y="90" text-anchor="middle" fill="#435e7a" font-size="12" font-family="Segoe UI,sans-serif">No data</text>'; document.getElementById('dleg').innerHTML=''; return; }
  const cx=100,cy=82,R=62,ri=44;
  const vs=[{v:a,c:'#22c55e'},{v:p,c:'#f59e0b'},{v:n,c:'#f87171'}];
  let ang=-Math.PI/2, arcs='';
  vs.forEach(({v,c})=>{
    if(!v) return;
    const sw=2*Math.PI*(v/tot);
    const x1=cx+R*Math.cos(ang),y1=cy+R*Math.sin(ang); ang+=sw;
    const x2=cx+R*Math.cos(ang),y2=cy+R*Math.sin(ang);
    const xi1=cx+ri*Math.cos(ang-sw),yi1=cy+ri*Math.sin(ang-sw);
    const xi2=cx+ri*Math.cos(ang),yi2=cy+ri*Math.sin(ang);
    const lg=sw>Math.PI?1:0;
    arcs+=`<path d="M${x1},${y1} A${R},${R} 0 ${lg},1 ${x2},${y2} L${xi2},${yi2} A${ri},${ri} 0 ${lg},0 ${xi1},${yi1} Z" fill="${c}" opacity="0.9"/>`;
  });
  const pct=Math.round(a/tot*100);
  sv.innerHTML=arcs+`<text x="${cx}" y="${cy-4}" text-anchor="middle" font-size="22" font-weight="700" fill="#f0f6ff" font-family="Segoe UI,sans-serif">${a}</text><text x="${cx}" y="${cy+14}" text-anchor="middle" font-size="10" fill="#7e9cbf" font-family="Segoe UI,sans-serif">active (${pct}%)</text>`;
  document.getElementById('dleg').innerHTML=[['#22c55e','Active',a],['#f59e0b','Partial',p],['#f87171','Not Found',n]].map(([c,l,v])=>`<div class="leg-r"><div class="leg-d" style="background:${c}"></div><span style="flex:1">${l}</span><span style="color:var(--t0);font-weight:600">${v}</span></div>`).join('');
}

// ── Bar chart ─────────────────────────────────────────────────────────────────
function renderBar(recs) {
  const days=period||30, bkts={};
  for(let i=days-1;i>=0;i--){ const d=new Date(); d.setDate(d.getDate()-i); bkts[d.toISOString().slice(0,10)]={a:0,p:0,n:0}; }
  recs.forEach(r=>{ const d=r.timestamp.slice(0,10); if(!bkts[d]) return; if(r.overall_status==='active')bkts[d].a++; else if(r.overall_status==='partial')bkts[d].p++; else bkts[d].n++; });
  const keys=Object.keys(bkts), mv=Math.max(1,...Object.values(bkts).map(b=>b.a+b.p+b.n));
  const W=530,H=185,PL=24,PR=8,PT=10,PB=32;
  const aW=W-PL-PR, bw=Math.max(3,Math.floor(aW/keys.length)-1), gap=aW/keys.length;
  const yS=v=>(H-PT-PB)*(v/mv);
  let bars='',labs='',grid='';
  for(let g=0;g<=4;g++){
    const y=PT+(H-PT-PB)*(1-g/4);
    grid+=`<line x1="${PL}" y1="${y}" x2="${W-PR}" y2="${y}" stroke="#1e3352" stroke-width="0.6"/>`;
    grid+=`<text x="${PL-3}" y="${y+3}" text-anchor="end" font-size="8" fill="#435e7a" font-family="Segoe UI,sans-serif">${Math.round(mv*g/4)}</text>`;
  }
  keys.forEach((k,i)=>{
    const {a,p,n}=bkts[k], x=PL+i*gap;
    let y=H-PB;
    if(n){const h=Math.max(1,Math.round(yS(n)));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="rgba(248,113,113,.3)" stroke="#f87171" stroke-width="0.6" rx="1"/>`;y-=h;}
    if(p){const h=Math.max(1,Math.round(yS(p)));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="rgba(245,158,11,.3)" stroke="#f59e0b" stroke-width="0.6" rx="1"/>`;y-=h;}
    if(a){const h=Math.max(1,Math.round(yS(a)));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="rgba(34,197,94,.3)" stroke="#22c55e" stroke-width="0.6" rx="1"/>`;}
    const sk=keys.length>20?Math.ceil(keys.length/12):1;
    if(i%sk===0) labs+=`<text x="${x+bw/2}" y="${H-PB+11}" text-anchor="middle" font-size="8" fill="#435e7a" font-family="Segoe UI,sans-serif">${k.slice(5)}</text>`;
  });
  const ly=H-3;
  const leg=`<rect x="${PL}" y="${ly-5}" width="7" height="5" fill="rgba(34,197,94,.3)" stroke="#22c55e" stroke-width=".6"/><text x="${PL+9}" y="${ly}" font-size="8" fill="#7e9cbf" font-family="Segoe UI,sans-serif">Active</text><rect x="${PL+52}" y="${ly-5}" width="7" height="5" fill="rgba(245,158,11,.3)" stroke="#f59e0b" stroke-width=".6"/><text x="${PL+61}" y="${ly}" font-size="8" fill="#7e9cbf" font-family="Segoe UI,sans-serif">Partial</text><rect x="${PL+110}" y="${ly-5}" width="7" height="5" fill="rgba(248,113,113,.3)" stroke="#f87171" stroke-width=".6"/><text x="${PL+119}" y="${ly}" font-size="8" fill="#7e9cbf" font-family="Segoe UI,sans-serif">Not Found</text>`;
  document.getElementById('bsvg').innerHTML=grid+bars+labs+leg;
}

// ── Table ─────────────────────────────────────────────────────────────────────
function drawTable() {
  let rows=allH.filter(h=>{
    if(srch && !h.hostname.toLowerCase().includes(srch)) return false;
    return true;
  });
  rows.sort((a,b)=>{
    let av=a[scol]??'zzz',bv=b[scol]??'zzz';
    if(scol==='eff_status'){av=SRANK[av]??1;bv=SRANK[bv]??1;}
    if(scol==='next_rv'){
      av=a.eff_status==='active'?9999:(Date.now()-new Date(a.last_checked))/86400000;
      bv=b.eff_status==='active'?9999:(Date.now()-new Date(b.last_checked))/86400000;
    }
    if(typeof av==='string')av=av.toLowerCase();
    if(typeof bv==='string')bv=bv.toLowerCase();
    return (av<bv?-1:av>bv?1:0)*(sasc?1:-1);
  });
  const tot=rows.length, pgs=Math.max(1,Math.ceil(tot/PS));
  if(cp>pgs)cp=pgs;
  const sl=rows.slice((cp-1)*PS,cp*PS);
  const tb=document.getElementById('tbody');
  document.getElementById('trcount').textContent=tot+' host'+(tot!==1?'s':'');
  if(!sl.length){
    tb.innerHTML='<tr><td colspan="9" class="nodata">No hosts found.</td></tr>';
    document.getElementById('pager').innerHTML=''; return;
  }
  tb.innerHTML=sl.map(h=>{
    const da=h.days_ago!=null?(h.days_ago===0?'Today':`${h.days_ago}d ago`):'—';
    const dac=h.days_ago!=null&&h.days_ago>7?'color:var(--red)':'color:var(--t2)';
    const og=h.os_group&&h.os_group!=='—'?`<span class="obadge">${h.os_group}</span>`:'<span style="color:var(--t3)">—</span>';
    const ob=h.status_override?'<span class="bovr">override</span>':'';
    const db=h.deleted?'<span class="bdel">exception</span>':'';
    const ni=h.note?`<span class="note-i" title="${h.note.replace(/"/g,'&quot;')}">&#9432;</span>`:'';
    const hn=`<span class="hn${h.deleted?' hndel':''}">${h.hostname}</span>${ni}${ob}${db}`;
    const esc=h.hostname.replace(/'/g,"\\'");
    return `<tr${h.deleted?' class="drow"':''}>
      <td>${hn}</td>
      <td>${og}</td>
      <td><span class="sb ${SCLS[h.eff_status]||'sp3'}">${SLBL[h.eff_status]||h.eff_status}</span></td>
      <td style="color:var(--t2)">${rel(h.last_checked)}</td>
      <td style="${dac}">${da}</td>
      <td style="color:var(--t3);text-align:center">${h.checks}</td>
      <td>${nextRV(h)}</td>
      <td>${trend(h)}</td>
      <td>
        <div class="ract">
          ${h.deleted?'':`<button class="btn-del" title="Delete bad read" onclick="quickDel('${esc}')">&#128465;</button>`}
          <button class="btn-edit" title="Edit / override" onclick="openM('${esc}')">&#9998;</button>
        </div>
      </td>
    </tr>`;
  }).join('');
  // Pager
  const pg=document.getElementById('pager');
  if(pgs<=1){pg.innerHTML=`<span class="pginfo">Showing all ${tot} host${tot!==1?'s':''}</span>`;return;}
  let bs='';
  for(let i=1;i<=pgs;i++) bs+=`<button class="btn btn-sm${i===cp?' btn-pri':''}" onclick="goP(${i})">${i}</button>`;
  pg.innerHTML=`<span class="pginfo">${(cp-1)*PS+1}–${Math.min(cp*PS,tot)} of ${tot}</span>${bs}`;
}
function goP(n){cp=n;drawTable();}

// ── Quick delete (no modal, no reason required) ───────────────────────────────
function quickDel(hn) {
  setO(hn,{deleted:true,exception_reason:'Removed via dashboard quick-delete'});
  undoPend=hn; showUndo(hn); render();
}

// ── Modal ─────────────────────────────────────────────────────────────────────
function openM(hn) {
  editHN=hn; const o=getO(hn);
  document.getElementById('mtitle').textContent='Edit: '+hn;
  document.getElementById('mst').value=o.status_override||'';
  document.getElementById('mnote').value=o.note||'';
  document.getElementById('mreason').value=o.exception_reason||'';
  document.getElementById('mreason').classList.remove('err');
  document.getElementById('mdelbtn').style.display=o.deleted?'none':'';
  document.getElementById('mbg').classList.add('open');
}
function closeM(){document.getElementById('mbg').classList.remove('open');editHN=null;}
function saveM() {
  if(!editHN) return;
  setO(editHN,{status_override:document.getElementById('mst').value||null,
               note:document.getElementById('mnote').value.trim(),
               exception_reason:document.getElementById('mreason').value.trim(),
               deleted:getO(editHN).deleted||false});
  closeM(); render(); toast('Saved: '+editHN);
}
function delFromM() {
  if(!editHN) return;
  const r=document.getElementById('mreason').value.trim();
  if(!r){document.getElementById('mreason').classList.add('err');return;}
  const hn=editHN;
  setO(hn,{deleted:true,exception_reason:r,
           note:document.getElementById('mnote').value.trim(),
           status_override:document.getElementById('mst').value||null});
  closeM(); undoPend=hn; showUndo(hn); render();
}

// ── Undo ──────────────────────────────────────────────────────────────────────
function showUndo(hn) {
  clearInterval(undoTmr); undoSecs=10;
  document.getElementById('umsg').textContent=`"${hn}" removed.`;
  document.getElementById('utmr').textContent=undoSecs+'s';
  document.getElementById('ubar').classList.add('show');
  undoTmr=setInterval(()=>{
    undoSecs--;
    document.getElementById('utmr').textContent=undoSecs+'s';
    if(undoSecs<=0){clearInterval(undoTmr);document.getElementById('ubar').classList.remove('show');undoPend=null;}
  },1000);
}
function undoDel() {
  if(!undoPend) return;
  setO(undoPend,{deleted:false});
  clearInterval(undoTmr);
  document.getElementById('ubar').classList.remove('show');
  render(); toast('"'+undoPend+'" restored.');
  undoPend=null;
}

// ── Toast ─────────────────────────────────────────────────────────────────────
let ttmr=null;
function toast(msg) {
  clearTimeout(ttmr); const el=document.getElementById('toast');
  el.textContent=msg; el.classList.add('show');
  ttmr=setTimeout(()=>el.classList.remove('show'),2800);
}

// ── Export ────────────────────────────────────────────────────────────────────
function exportCSV() {
  const hosts=applyOvr(buildMap(filtRecs(period)));
  const rows=[['Hostname','OS Group','Status','Status Override','Last Checked','Days Since Event','Checks','Exception','Note','Exception Reason']];
  hosts.forEach(h=>rows.push([h.hostname,h.os_group||'',SLBL[h.eff_status]||h.eff_status,
    h.status_override?SLBL[h.status_override]:'',h.last_checked,h.days_ago??'',
    h.checks,h.deleted?'YES':'',h.note||'',h.exception_reason||'']));
  const csv=rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\r\n');
  dl(new Blob([csv],{type:'text/csv'}),'signoff_export.csv');
}
function exportOvr() {
  dl(new Blob([JSON.stringify(loadOvr(),null,2)],{type:'application/json'}),'signoff_overrides.json');
  toast('Place file at: '+OVR_PATH);
}
function dl(blob,name){const u=URL.createObjectURL(blob);Object.assign(document.createElement('a'),{href:u,download:name}).click();URL.revokeObjectURL(u);}

// ── Keyboard ──────────────────────────────────────────────────────────────────
document.addEventListener('keydown',e=>{
  if(e.key==='Escape'){closeM();document.getElementById('ubar').classList.remove('show');}
  if(e.key==='ArrowRight') goP(Math.min(cp+1,Math.ceil(allH.length/PS)));
  if(e.key==='ArrowLeft')  goP(Math.max(cp-1,1));
});
document.getElementById('mbg').addEventListener('click',e=>{if(e.target===document.getElementById('mbg'))closeM();});

render();
</script>
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _log("=" * 64)
    _log(f"QRadar Signoff Auto-Draft v{VERSION}")
    _log(f"   Script dir      : {_SCRIPT_DIR}")
    _log(f"   Data file       : {SIGNOFF_DATA_PATH}")
    _log(f"   Dashboard       : {DASHBOARD_PATH}")
    _log(f"   Overrides       : {_OVERRIDES_PATH}")
    _log(f"   Lookback        : {LOOKBACK_DAYS}d  |  Sent scan: {SENT_SCAN_DAYS}d")
    _log(f"   Reval cooldown  : {REVALIDATION_COOLDOWN_DAYS}d  "
         f"|  Active skip: {ACTIVE_SKIP_DAYS}d")
    _log(f"   Trigger DL      : '{TRIGGER_DL}'")
    _log(f"   Escalation To   : {ESCALATION_TO}")
    _log(f"   Escalation CC   : {ESCALATION_CC}")
    _log(f"   MODE            : DRAFT ONLY — nothing is ever sent automatically")
    _log("=" * 64)

    if not acquire_lock():
        return

    try:
        overrides      = load_overrides()
        deleted_hosts  = {hn for hn, o in overrides.items() if o.get('deleted')}
        if deleted_hosts:
            _log(f"Exception hosts (skipped on this run): {sorted(deleted_hosts)}")

        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        if not test_qradar_connection():
            _log("ERROR: QRadar unreachable — exiting. All emails left untouched.")
            return
        fetch_log_source_types()

        cutoff_str = (datetime.now() - timedelta(days=LOOKBACK_DAYS)
                      ).strftime('%Y-%m-%d')
        try:
            inbox_items = list(inbox.Items.Restrict(
                f"[ReceivedTime] >= '{cutoff_str}'"
            ))
        except Exception:
            _log("WARN: Restrict failed — iterating full folder (locale fallback)")
            cutoff_dt   = datetime.now() - timedelta(days=LOOKBACK_DAYS)
            inbox_items = [
                item for item in inbox.Items
                if _com_dt_to_py(getattr(item, 'ReceivedTime', None)) or datetime.min
                >= cutoff_dt
            ]

        _log(f"\nInbox scan: {len(inbox_items)} email(s) in last {LOOKBACK_DAYS}d window.")

        processed = skipped = drafted = revalidated = 0

        for mail_item in inbox_items:
            try:
                if mail_item.Class != 43:
                    continue
            except Exception:
                continue

            subject = ''
            try:
                subject = mail_item.Subject or ''
            except Exception:
                continue

            passed, reason = passes_subject_guards(subject)
            if not passed:
                skipped += 1
                _log(f"   SKIP (subject — {reason}): '{subject[:60]}'")
                continue

            try:
                sender = mail_item.SenderEmailAddress or ''
            except Exception:
                sender = ''

            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                continue

            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"   SKIP (sender not allowed — {sender}): '{subject[:60]}'")
                continue

            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"   SKIP ('{TRIGGER_DL}' not in body): '{subject[:60]}'")
                continue

            hostnames = extract_hostnames(subject)
            if not hostnames:
                skipped += 1
                _log(f"   SKIP (no hostnames): '{subject[:60]}'")
                continue

            active_hn = [h for h in hostnames if h not in deleted_hosts]
            if not active_hn:
                skipped += 1
                _log(f"   SKIP (all hosts are exceptions): {hostnames}")
                continue
            if len(active_hn) < len(hostnames):
                skipped_hn = [h for h in hostnames if h in deleted_hosts]
                _log(f"   NOTE: Skipping exception host(s) {skipped_hn}")
            hostnames = active_hn

            hn_fset = frozenset(hostnames)

            if hn_fset in _runtime_drafted_hosts:
                skipped += 1
                _log(f"   SKIP (already drafted in this run): {hostnames}")
                continue

            was_active, active_date = _was_active_recently(hn_fset)
            if was_active:
                skipped += 1
                _log(f"   SKIP (Active in data file on {active_date} "
                     f"within {ACTIVE_SKIP_DAYS}d window)")
                continue

            _log(f"\nCandidate: '{subject[:70]}'")
            _log(f"   Sender : {sender}")
            _log(f"   Hosts  : {hostnames}")

            last_tag, last_dt = check_conversation_status(mail_item, sent, drafts)
            cooldown_cutoff   = datetime.now() - timedelta(days=REVALIDATION_COOLDOWN_DAYS)

            if last_tag is None:
                is_revalidation = False
                _log(f"   -> New signoff (no prior tag in Sent/Drafts)")

            elif last_tag in (TAG_ACTIVE, 'legacy'):
                skipped += 1
                dt_s = last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'
                _log(f"   SKIP (Sent/Drafts: Active tag from {dt_s})")
                continue

            elif last_tag in REVALIDATABLE_TAGS:
                if REVALIDATION_COOLDOWN_DAYS > 0 and last_dt and last_dt >= cooldown_cutoff:
                    days_ago = (datetime.now() - last_dt).days
                    skipped += 1
                    _log(f"   SKIP (cooldown — last checked {days_ago}d ago, "
                         f"cooldown is {REVALIDATION_COOLDOWN_DAYS}d)")
                    continue
                dt_s = last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'
                _log(f"   -> Revalidating {last_tag} from {dt_s}")
                is_revalidation = True

            else:
                skipped += 1
                _log(f"   SKIP (unrecognised tag state: '{last_tag}')")
                continue

            _log(f"   Querying QRadar for {len(hostnames)} host(s)...")
            hostname_qr_pairs = []
            for hn in hostnames:
                qr = query_all_log_sources_readonly(hn)
                _log(f"      {hn}: {qr['status']} ({len(qr.get('sources',[]))} sources)")
                hostname_qr_pairs.append((hn, qr))

            body, overall_status, host_tracking = build_reply_for_all_hosts(
                hostname_qr_pairs
            )
            _log(f"   Overall: {overall_status.upper()}")

            hostnames_str = ' | '.join(hostnames)
            success = create_draft_reply(
                mail_item, body, overall_status,
                hostnames_str=hostnames_str,
                is_revalidation=is_revalidation,
            )

            if success:
                drafted += 1
                if is_revalidation:
                    revalidated += 1
                _runtime_drafted_hosts.add(hn_fset)
                append_record({
                    'id':              str(uuid.uuid4()),
                    'version':         VERSION,
                    'timestamp':       datetime.now().isoformat(timespec='seconds'),
                    'subject':         subject,
                    'sender':          sender,
                    'is_revalidation': is_revalidation,
                    'overall_status':  overall_status,
                    'host_results':    host_tracking,
                })

            processed += 1

        # Dashboard written + auto-opened here
        generate_dashboard()

        _log(f"\n{'=' * 64}")
        _log(f"Run complete — v{VERSION}")
        _log(f"   Processed   : {processed}  Drafted: {drafted} "
             f"({revalidated} revalidation{'s' if revalidated != 1 else ''})")
        _log(f"   Skipped     : {skipped}")
        _log(f"\n   File paths (paste into config to hardcode):")
        _log(f"   RUN_LOG_PATH      = r'{RUN_LOG_PATH}'")
        _log(f"   LOCKFILE_PATH     = r'{LOCKFILE_PATH}'")
        _log(f"   SIGNOFF_DATA_PATH = r'{SIGNOFF_DATA_PATH}'")
        _log(f"   DASHBOARD_PATH    = r'{DASHBOARD_PATH}'")
        _log(f"{'=' * 64}\n")

    finally:
        release_lock()


if __name__ == '__main__':
    main()
