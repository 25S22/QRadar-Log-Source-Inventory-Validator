"""
QRadar Signoff Auto-Draft
─────────────────────────
Scans a dedicated Outlook folder for SIEM signoff emails, queries QRadar
for each hostname listed in the subject line, and saves a formatted draft
reply to Outlook Drafts for manual review and sending.

Multi-hostname format:  Security Signoff | HOST1 | HOST2 | HOST3
Routing:
  All hosts Active  →  ReplyAll (original requestor + recipients)
  Any Partial/NotFound → Escalation list only (ESCALATION_TO / ESCALATION_CC)

Revalidation:
  Partial and Not Found threads are automatically re-checked after
  REVALIDATION_COOLDOWN_DAYS — no manual intervention needed.
  Confirmed Active threads are permanently skipped.

THIS SCRIPT IS DRAFT-ONLY. reply.Save() is called, NEVER reply.Send().
"""

import json
import os
import urllib3
import uuid
import win32com.client

import requests
from datetime import datetime, timedelta

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
# Credentials — set as Windows environment variables, never hardcode in source.
QRADAR_HOST     = os.environ.get('QRADAR_HOST',     'https://your-qradar-host')
QRADAR_USERNAME = os.environ.get('QRADAR_USERNAME', 'your-username')
QRADAR_PASSWORD = os.environ.get('QRADAR_PASSWORD', 'your-password')
VERIFY_SSL      = False

# Subject matching
# Format: "<SUBJECT_KEYWORD> <SUBJECT_SEPARATOR> HOST1 <SUBJECT_SEPARATOR> HOST2 ..."
# Example: "Security Signoff | HOSTNAME-01 | HOSTNAME-02"
SUBJECT_KEYWORD   = 'Security Signoff'
SUBJECT_SEPARATOR = '|'

# Normal signoff scan window — how far back to look for NEW (unprocessed) emails.
LOOKBACK_HOURS = 24

# Sender allowlist — exact addresses or @domain wildcards. Empty list = allow all.
ALLOWED_SENDERS = []

# Your reply-from address — prevents processing your own sent items on shared mailboxes.
YOUR_EMAIL_ADDRESS = 'youremail@yourorg.com'

# DL string that must appear in the email body. Case-insensitive. '' = disabled.
TRIGGER_DL = '@SOC-DL@yourorg.com'

# ─── ESCALATION RECIPIENTS ─────────────────────────────────────────────────────
# Partial / Not Found → ONLY these addresses. The original requestor is excluded.
# Active (all hosts confirmed) → standard ReplyAll, no override.
#
# Supports exact addresses. Add as many as needed.
ESCALATION_TO = [
    'onboarding-owner@yourorg.com',
]
ESCALATION_CC = [
    '@SOC-DL@yourorg.com',
    # 'siem-lead@yourorg.com',
    # 'extra-analyst@yourorg.com',
]

# ─── REVALIDATION ──────────────────────────────────────────────────────────────
# REVALIDATION_WINDOW_DAYS  : How far back to scan for previously-flagged threads.
# REVALIDATION_COOLDOWN_DAYS: Min gap between re-checks of the same thread.
REVALIDATION_WINDOW_DAYS   = 14
REVALIDATION_COOLDOWN_DAYS = 3

# ─── OUTLOOK FOLDERS ───────────────────────────────────────────────────────────
# Set SIGNOFF_FOLDER_NAME to the Outlook rule subfolder under Inbox.
# Set to None to scan the full Inbox (not recommended for production).
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'
DRAFTS_FOLDER_NAME  = 'Drafts'
SENT_FOLDER_NAME    = 'Sent Items'

# ─── FILE PATHS ────────────────────────────────────────────────────────────────
RUN_LOG_PATH      = r'C:\path\to\signoff_runner.log'
LOCKFILE_PATH     = r'C:\path\to\signoff.lock'
# Dashboard data — one JSON file, appended on every run. Point the dashboard
# HTML to this same file using the file-picker on first load.
SIGNOFF_DATA_PATH = r'C:\path\to\signoff_data.json'

REQUEST_TIMEOUT = 30

# ─── OS TYPE VALIDATION ────────────────────────────────────────────────────────
# First entry in each 'required' list is the OS signature (used for detection).
# All entries in 'required' are validated and shown as individual table rows.
# Set to {} to disable type validation (simple found/not-found reply).
OS_TYPE_GROUPS = {
    'Windows': {
        'required': ['Microsoft Security', 'WinCollect'],
    },
    'Linux': {
        'required': ['Linux OS'],
    },
}

# ─── SUBJECT OUTCOME TAGS ──────────────────────────────────────────────────────
# Written into draft subject lines to record outcome without external storage.
# The subject guard checks '[processed' (no closing bracket) to catch all variants.
TAG_ACTIVE     = '[Processed-Active]'
TAG_PARTIAL    = '[Processed-Partial]'
TAG_NOT_FOUND  = '[Processed-NotFound]'
REVALIDATABLE_TAGS = {TAG_PARTIAL, TAG_NOT_FOUND}

# ─── CONSTANTS ─────────────────────────────────────────────────────────────────
ACTIVITY_THRESHOLD_DAYS  = 7
_MIN_TS                  = 0
_MAX_TS                  = 2147483647
LOG_SOURCE_TYPES_CACHE   = {}
STATUS_PRIORITY          = {'not_found': 2, 'partial': 1, 'active': 0}

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────


# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _log(message):
    """Appends a timestamped line to the run log and prints to console."""
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    try:
        with open(RUN_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception as e:
        print(f"⚠️  Could not write to log: {e}")


def acquire_lock():
    """Prevents two instances running simultaneously via a PID lockfile."""
    if os.path.exists(LOCKFILE_PATH):
        _log("⚠️  Lockfile exists — another instance may be running. Exiting.")
        return False
    try:
        with open(LOCKFILE_PATH, 'w') as f:
            f.write(str(os.getpid()))
        return True
    except Exception as e:
        _log(f"❌ Could not create lockfile: {e}")
        return False


def release_lock():
    """Removes lockfile on clean exit."""
    try:
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
    except Exception as e:
        _log(f"⚠️  Could not remove lockfile: {e}")


def _com_dt_to_py(com_dt):
    """
    Converts a pywintypes COM datetime to a naive Python datetime (local time).
    Returns None on failure. Used for all win32com date comparisons.
    """
    if com_dt is None:
        return None
    try:
        return datetime(
            com_dt.year, com_dt.month, com_dt.day,
            com_dt.hour, com_dt.minute, com_dt.second,
        )
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════════════════
# QRADAR CONNECTION
# ══════════════════════════════════════════════════════════════════════════════

def test_qradar_connection():
    """
    Validates QRadar connectivity before any email processing begins.
    Failure aborts the entire run — all emails remain untouched for next run.
    """
    _log("🔗 Testing QRadar connection...")
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/help/versions"
    try:
        resp = requests.get(
            endpoint,
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'},
        )
        if resp.status_code == 200:
            _log("✅ QRadar connection successful.")
            return True
        elif resp.status_code == 401:
            _log("❌ Auth failed — check QRADAR_USERNAME / QRADAR_PASSWORD env vars.")
            return False
        _log(f"⚠️  Unexpected HTTP {resp.status_code} from QRadar.")
        return False
    except Exception as e:
        _log(f"❌ Connection failed: {e}")
        return False


def fetch_log_source_types():
    """Pre-fetches Log Source Type ID → Name into the module-level cache."""
    _log("📥 Fetching Log Source Types into cache...")
    endpoint = (
        f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
        f"/log_source_management/log_source_types"
    )
    try:
        resp = requests.get(
            endpoint,
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'},
        )
        if resp.status_code == 200:
            for t in resp.json():
                ls_id, ls_name = t.get('id'), t.get('name')
                if ls_id is not None and ls_name is not None:
                    LOG_SOURCE_TYPES_CACHE[ls_id] = ls_name
            _log(f"✅ Cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
        else:
            _log(f"⚠️  Failed to fetch Log Source Types: HTTP {resp.status_code}")
    except Exception as e:
        _log(f"❌ Error fetching Log Source Types: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# QRADAR QUERIES — STRICTLY READ-ONLY (GET requests only)
# ══════════════════════════════════════════════════════════════════════════════

def _safe_timestamp(timestamp_ms):
    """
    Converts a QRadar epoch-ms (or epoch-s) timestamp to readable string,
    activity status, and days-since value. Read-only, no side effects.
    """
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    try:
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)
        # QRadar sometimes returns seconds, sometimes milliseconds
        epoch_s = timestamp_ms / 1000.0 if timestamp_ms > 4102444800 else timestamp_ms
        if epoch_s <= _MIN_TS or epoch_s > _MAX_TS:
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None
        last_event_dt = datetime.fromtimestamp(epoch_s)
        days_ago      = (datetime.now() - last_event_dt).days
        threshold_dt  = datetime.now() - timedelta(days=ACTIVITY_THRESHOLD_DAYS)
        activity      = 'Active' if last_event_dt > threshold_dt else 'Inactive'
        return last_event_dt.strftime('%Y-%m-%d %H:%M:%S'), activity, days_ago
    except Exception:
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def query_all_log_sources_readonly(hostname):
    """
    STRICTLY READ-ONLY — fetches ALL log sources matching hostname.
    Only HTTP GET. Nothing in QRadar is created, modified, or deleted.

    Returns:
        {
            'status':  'Found' | 'Not Found' | 'API Error ...',
            'sources': [ {name, ls_type, enabled, last_seen, activity, days_ago}, ... ]
        }
    """
    clean = str(hostname).replace('"', '').replace("'", "").strip()
    endpoint = (
        f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
        f"/log_source_management/log_sources"
    )
    try:
        resp = requests.get(
            endpoint,
            params={'filter': f'name ilike "%{clean}%"'},
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'},
        )
        if resp.status_code != 200:
            return {'status': f'API Error {resp.status_code}', 'sources': []}
        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', 'sources': []}
        sources = []
        for src in ls_data:
            type_id      = src.get('type_id')
            ls_type_name = LOG_SOURCE_TYPES_CACHE.get(type_id, f'Unknown TypeID:{type_id}')
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
    except Exception as e:
        return {'status': f'Error: {str(e)[:80]}', 'sources': []}


def validate_expected_types(all_sources_result, required_types):
    """
    Checks each required type keyword against returned QRadar sources.
    Fuzzy keyword matching — every word in the keyword must appear in the type name.
    Returns one result dict per required type (shown as one table row each).
    """
    results  = []
    sources  = all_sources_result.get('sources', [])
    for expected_kw in required_types:
        exp_words = str(expected_kw).lower().split()
        matched   = [
            s for s in sources
            if all(w in str(s.get('ls_type', '')).lower() for w in exp_words)
        ]
        if not matched:
            results.append({'expected': expected_kw, 'found': False,
                            'ls_type': None, 'ls_name': None,
                            'last_seen': None, 'days_ago': None})
            continue
        # Prefer enabled sources; sort by recency within each group
        enabled  = sorted([s for s in matched if s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        disabled = sorted([s for s in matched if not s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        best = enabled[0] if enabled else disabled[0]
        results.append({'expected': expected_kw, 'found': True,
                        'ls_type':  best.get('ls_type'),
                        'ls_name':  best.get('name'),
                        'last_seen': best.get('last_seen'),
                        'days_ago': best.get('days_ago')})
    return results


def detect_os_group(sources):
    """
    Detects OS group by fuzzy-matching the first 'required' entry (the OS
    signature) against all returned source types. First match in dict order wins.
    Returns (group_name, group_rules) or (None, None).
    """
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


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL GUARDS
# ══════════════════════════════════════════════════════════════════════════════

def is_sender_allowed(sender_address):
    """
    Validates sender against ALLOWED_SENDERS.
    @domain entries match any address ending with that domain. Empty = allow all.
    """
    if not ALLOWED_SENDERS:
        return True
    if not sender_address:
        return False
    sender_clean = sender_address.strip().lower()
    for entry in ALLOWED_SENDERS:
        e = entry.strip().lower()
        if e.startswith('@') and sender_clean.endswith(e):
            return True
        if not e.startswith('@') and sender_clean == e:
            return True
    return False


def passes_subject_guards(subject):
    """
    All guards must pass before an email is processed.

    Note: '[processed' (no closing bracket) catches ALL outcome tag variants:
    [Processed-Active], [Processed-Partial], [Processed-NotFound], [Processed].
    Using '[processed]' (with bracket) would silently miss the new variants.
    """
    if not subject:
        return False, "empty subject"
    s   = subject.strip()
    sl  = s.lower()
    if any(sl.startswith(p) for p in ('re:', 'fw:', 'fwd:')):
        return False, f"reply/forward prefix in: '{s[:30]}'"
    if '[processed' in sl:
        return False, "subject already carries an outcome tag"
    if SUBJECT_SEPARATOR not in s:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found"
    left = s.split(SUBJECT_SEPARATOR)[0].strip().lower()
    if SUBJECT_KEYWORD.lower() not in left:
        return False, f"keyword '{SUBJECT_KEYWORD}' not found left of separator"
    return True, "ok"


def extract_hostnames(subject):
    """
    Extracts one or more hostnames from the subject line.

    Format: "Security Signoff | HOST1 | HOST2 | HOST3"
             ──────────────── ^ first sep  ^ subsequent seps
    Returns a list of stripped, non-empty hostname strings.
    Empty list if nothing is found right of the first separator.
    """
    parts = subject.split(SUBJECT_SEPARATOR, 1)
    if len(parts) < 2:
        return []
    remainder = parts[1]
    # Everything right of the first separator is a pipe-delimited hostname list
    return [h.strip() for h in remainder.split(SUBJECT_SEPARATOR) if h.strip()]


def body_contains_dl(mail_item):
    """
    Returns True if TRIGGER_DL appears in the plain-text or HTML body.
    Falls back to HTML body if plain text is empty. '' = disabled.
    """
    if not TRIGGER_DL.strip():
        return True
    dl_lower = TRIGGER_DL.strip().lower()
    try:
        if dl_lower in (mail_item.Body or '').strip().lower():
            return True
        if dl_lower in (mail_item.HTMLBody or '').strip().lower():
            return True
        return False
    except Exception as e:
        _log(f"⚠️  Body DL check failed: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# CONVERSATION STATUS (State Machine via Subject Tags)
# ══════════════════════════════════════════════════════════════════════════════

def check_conversation_status(mail_item, sent_folder, drafts_folder):
    """
    Scans Sent Items and Drafts for any prior reply in this conversation thread
    bearing a recognised outcome tag, without any external state file.

    Returns:
        (TAG_ACTIVE,    dt) → confirmed active — permanent skip
        (TAG_PARTIAL,   dt) → partial — revalidate after cooldown
        (TAG_NOT_FOUND, dt) → not found — revalidate after cooldown
        ('legacy',      dt) → old [Processed] without suffix — treat as Active
        (None,         None)→ no prior reply — treat as brand-new signoff

    Sent Items scan is restricted to REVALIDATION_WINDOW_DAYS for performance.
    """
    conv_id  = mail_item.ConversationID
    last_tag = None
    last_dt  = None

    def _tag(subject):
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

    # Restrict Sent Items scan to the revalidation window for speed
    reval_cutoff = (datetime.now() - timedelta(days=REVALIDATION_WINDOW_DAYS)
                    ).strftime('%m/%d/%Y %I:%M %p')
    try:
        for item in sent_folder.Items.Restrict(f"[SentOn] >= '{reval_cutoff}'"):
            try:
                if item.ConversationID == conv_id:
                    _update(_tag(item.Subject), _com_dt_to_py(item.SentOn))
            except Exception:
                continue
    except Exception as e:
        _log(f"⚠️  Could not scan Sent Items: {e}")

    try:
        for item in drafts_folder.Items:
            try:
                if item.ConversationID == conv_id:
                    _update(_tag(item.Subject),
                            _com_dt_to_py(item.LastModificationTime))
            except Exception:
                continue
    except Exception as e:
        _log(f"⚠️  Could not scan Drafts: {e}")

    return last_tag, last_dt


# ══════════════════════════════════════════════════════════════════════════════
# HTML EMAIL BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _build_host_html_section(hostname, qradar_result):
    """
    Builds the HTML block for a single hostname result.

    Returns:
        html_section : str   — <div> block for this host
        host_status  : str   — 'active' | 'partial' | 'not_found'
        type_records : list  — serialisable list for JSON dashboard data
        os_group     : str|None
    """
    status  = qradar_result.get('status')
    sources = qradar_result.get('sources', [])

    # ── Not Found ─────────────────────────────────────────────────────────────
    if status != 'Found' or not sources:
        section = f"""
        <div style="margin-bottom:20px;border:1px solid #f5c6c6;border-radius:8px;
                    overflow:hidden;">
          <div style="background:#c0392b;color:#fff;padding:9px 14px;
                      font-size:13px;font-weight:700;letter-spacing:0.3px;">
            ✖&nbsp; {hostname} — Not Found in QRadar
          </div>
          <div style="padding:12px 14px;font-size:12px;color:#555;">
            <b>{hostname}</b> was not found in the QRadar log source inventory.
            Please ensure the asset is onboarded and configured correctly.
          </div>
        </div>"""
        return section, 'not_found', [], None

    # ── OS detection for type validation ──────────────────────────────────────
    group_name, group_rules = detect_os_group(sources)
    type_records = []

    if OS_TYPE_GROUPS and group_name:
        validation   = validate_expected_types(qradar_result,
                                               group_rules.get('required', []))
        any_missing  = any(not r['found'] for r in validation)
        any_silent   = any(r['found'] and r['days_ago'] is None for r in validation)
        any_problem  = any_missing or any_silent
        host_status  = 'partial' if any_problem else 'active'

        # Banner
        os_label = f' ({group_name})'
        if not any_problem:
            banner_bg  = '#1a7a4a'
            banner_txt = f'✔&nbsp; {hostname}{os_label} — Confirmed Reporting on SIEM'
        elif any_missing:
            banner_bg  = '#c87800'
            n_found    = sum(1 for r in validation if r['found'])
            banner_txt = (f'⚠&nbsp; {hostname}{os_label} — '
                          f'{n_found}/{len(validation)} required log sources found')
        else:
            banner_bg  = '#c87800'
            banner_txt = (f'⚠&nbsp; {hostname}{os_label} — '
                          f'Log sources present but not yet reporting')

        # Rows
        rows = ''
        for r in validation:
            if not r['found']:
                icon, row_bg, status_cell = (
                    '✖', '#fff5f5',
                    '<span style="color:#c0392b;font-weight:600;">Missing — requires onboarding</span>'
                )
            elif r['days_ago'] is None:
                icon, row_bg, status_cell = (
                    '⚠', '#fffbf0',
                    '<span style="color:#c87800;font-weight:600;">No events recorded yet</span>'
                )
            else:
                d_str = 'Today' if r['days_ago'] == 0 else f"{r['days_ago']}d ago"
                icon, row_bg, status_cell = (
                    '✔', '#f0faf4',
                    f'<span style="color:#1a7a4a;font-weight:600;">Active</span>'
                    f'&nbsp;<span style="color:#888;font-size:11px;">({d_str})</span>'
                )
            rows += f"""
            <tr style="background:{row_bg};">
              <td style="padding:7px 10px;font-size:12px;color:{'#c0392b' if not r['found'] else '#c87800' if r['days_ago'] is None else '#1a7a4a'};
                         font-weight:700;text-align:center;width:22px;">{icon}</td>
              <td style="padding:7px 10px;font-size:12px;font-weight:600;color:#333;">
                {r['expected']}</td>
              <td style="padding:7px 10px;font-size:12px;color:#555;">
                {r.get('ls_name') or '—'}</td>
              <td style="padding:7px 10px;font-size:12px;color:#555;">
                {r.get('last_seen') or '—'}</td>
              <td style="padding:7px 10px;font-size:12px;">{status_cell}</td>
            </tr>"""

            # Record for JSON
            type_records.append({
                'expected': r['expected'],
                'found':    r['found'],
                'days_ago': r['days_ago'],
            })

        detail = f"""
        <table style="width:100%;border-collapse:collapse;">
          <tr style="background:#f5f5f5;">
            <th style="padding:6px 10px;font-size:11px;color:#888;font-weight:600;
                       text-align:left;border-bottom:1px solid #ddd;width:22px;"></th>
            <th style="padding:6px 10px;font-size:11px;color:#888;font-weight:600;
                       text-align:left;border-bottom:1px solid #ddd;">Log Source Type</th>
            <th style="padding:6px 10px;font-size:11px;color:#888;font-weight:600;
                       text-align:left;border-bottom:1px solid #ddd;">Log Source Name</th>
            <th style="padding:6px 10px;font-size:11px;color:#888;font-weight:600;
                       text-align:left;border-bottom:1px solid #ddd;">Last Event</th>
            <th style="padding:6px 10px;font-size:11px;color:#888;font-weight:600;
                       text-align:left;border-bottom:1px solid #ddd;">Status</th>
          </tr>
          {rows}
        </table>"""

    else:
        # ── Simple mode: OS_TYPE_GROUPS empty or OS undetected ────────────────
        if OS_TYPE_GROUPS and not group_name:
            _log(f"      ⚠️  OS undetected for {hostname} — showing best source, "
                 f"no type validation.")

        enabled  = sorted([s for s in sources if s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        disabled = sorted([s for s in sources if not s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        best = enabled[0] if enabled else (disabled[0] if disabled else None)

        host_status  = 'active'
        banner_bg    = '#1a7a4a'
        banner_txt   = f'✔&nbsp; {hostname} — Confirmed Reporting on SIEM'
        group_name   = None

        if best:
            days_val = best.get('days_ago')
            days_str = 'Today' if days_val == 0 else f"{days_val} days ago" if days_val is not None else 'N/A'
            detail   = f"""
            <table style="width:100%;border-collapse:collapse;">
              <tr>
                <td style="padding:7px 10px;font-size:12px;color:#555;width:160px;
                           border-bottom:1px solid #eee;">Log Source Name</td>
                <td style="padding:7px 10px;font-size:12px;font-weight:600;color:#222;
                           border-bottom:1px solid #eee;">{best.get('name','N/A')}</td>
              </tr>
              <tr>
                <td style="padding:7px 10px;font-size:12px;color:#555;
                           border-bottom:1px solid #eee;">Log Source Type</td>
                <td style="padding:7px 10px;font-size:12px;color:#333;
                           border-bottom:1px solid #eee;">{best.get('ls_type','N/A')}</td>
              </tr>
              <tr>
                <td style="padding:7px 10px;font-size:12px;color:#555;">Last Event</td>
                <td style="padding:7px 10px;font-size:12px;color:#333;">
                  {best.get('last_seen','N/A')}
                  &nbsp;<span style="color:#888;font-size:11px;">({days_str})</span>
                </td>
              </tr>
            </table>"""
        else:
            detail = ''

    section = f"""
    <div style="margin-bottom:20px;border:1px solid #e0e0e0;border-radius:8px;
                overflow:hidden;">
      <div style="background:{banner_bg};color:#fff;padding:9px 14px;
                  font-size:13px;font-weight:700;letter-spacing:0.3px;">
        {banner_txt}
      </div>
      <div style="padding:0;">{detail}</div>
    </div>"""

    return section, host_status, type_records, group_name


def _build_full_reply_html(hostname_list, host_sections, host_statuses, run_time):
    """
    Wraps all per-host sections into a complete HTML email body.
    Includes a summary badge bar at the top for quick multi-host overview.
    """
    badge_cfg = {
        'active':    ('#1a7a4a', '✔'),
        'partial':   ('#c87800', '⚠'),
        'not_found': ('#c0392b', '✖'),
    }
    badges = ''
    for hn, hs in zip(hostname_list, host_statuses):
        bg, icon = badge_cfg.get(hs, ('#555', '?'))
        badges += (
            f'<span style="display:inline-block;background:{bg};color:#fff;'
            f'padding:4px 12px;border-radius:12px;font-size:11px;font-weight:600;'
            f'margin:0 4px 6px 0;">{icon}&nbsp;{hn}</span>'
        )

    count_label = f"{len(hostname_list)} host{'s' if len(hostname_list) != 1 else ''} checked"
    sections_html = '\n'.join(host_sections)

    return f"""
<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;color:#222;font-size:13px;
             line-height:1.6;margin:0;padding:0;">
  <div style="max-width:680px;padding:20px 0;">
    <p style="margin:0 0 14px 0;">Hi,</p>
    <p style="margin:0 0 10px 0;color:#555;font-size:12px;">
      Results for your SIEM Security Signoff request — {count_label}.
    </p>
    <div style="margin-bottom:18px;">{badges}</div>
    {sections_html}
    <p style="margin:20px 0 4px 0;color:#888;font-size:11px;">
      This is an automated response from the SIEM monitoring system.<br>
      Checked against QRadar on {run_time}.
    </p>
    <p style="margin:14px 0 0 0;">Regards,<br>
    <span style="font-weight:700;">Cyberdefence</span></p>
  </div>
</body>
</html>"""


def build_all_hosts_reply(hostname_list):
    """
    Main reply orchestrator — queries QRadar for every hostname in the list,
    builds a combined HTML body, and determines the overall signoff outcome.

    Overall status follows worst-case priority: not_found > partial > active.

    Returns:
        html_body     : str  — complete HTML email
        overall_status: str  — 'active' | 'partial' | 'not_found'
        host_records  : list — serialisable per-host data for JSON dashboard
    """
    run_time     = datetime.now().strftime('%d %B %Y, %H:%M')
    host_sections = []
    host_statuses = []
    host_records  = []
    overall_status = 'active'

    for hostname in hostname_list:
        _log(f"      🔍 [{hostname}] Querying QRadar...")
        qr = query_all_log_sources_readonly(hostname)
        _log(f"      📊 [{hostname}] Status: {qr['status']} | "
             f"Sources: {len(qr.get('sources', []))}")

        section, host_status, type_records, os_group = _build_host_html_section(
            hostname, qr
        )
        host_sections.append(section)
        host_statuses.append(host_status)

        # Escalate overall status to worst-case
        if STATUS_PRIORITY.get(host_status, 0) > STATUS_PRIORITY.get(overall_status, 0):
            overall_status = host_status

        host_records.append({
            'hostname':     hostname,
            'status':       host_status,
            'os_group':     os_group,
            'type_results': type_records,
        })
        _log(f"      🏷️  [{hostname}] Host status: {host_status.upper()}")

    html_body = _build_full_reply_html(
        hostname_list, host_sections, host_statuses, run_time
    )
    return html_body, overall_status, host_records


# ══════════════════════════════════════════════════════════════════════════════
# JSON DATA WRITER (Dashboard Feed)
# ══════════════════════════════════════════════════════════════════════════════

def write_signoff_record(email_subject, sender, host_records,
                         overall_status, is_revalidation, prior_status):
    """
    Appends one signoff run record to SIGNOFF_DATA_PATH for the dashboard.
    Reads existing data, appends, and writes back atomically.
    Silently logs failures — data write errors never abort the draft creation.
    """
    record = {
        'run_id':          f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}",
        'timestamp':       datetime.now().isoformat(),
        'email_subject':   email_subject,
        'sender':          sender,
        'is_revalidation': is_revalidation,
        'prior_status':    prior_status,
        'overall_status':  overall_status,
        'hosts':           host_records,
    }

    data = {'schema_version': 2, 'entries': []}
    if os.path.exists(SIGNOFF_DATA_PATH):
        try:
            with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            _log(f"⚠️  Could not read data file (will overwrite): {e}")

    data['entries'].append(record)

    try:
        with open(SIGNOFF_DATA_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        _log(f"      💾 Dashboard record written ({overall_status})")
    except Exception as e:
        _log(f"⚠️  Could not write dashboard data: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# DRAFT CREATOR
# ══════════════════════════════════════════════════════════════════════════════

def create_draft_reply(mail_item, html_body, hostname_list,
                       overall_status, is_revalidation=False):
    """
    Saves a reply draft to Outlook Drafts. NEVER calls reply.Send().

    Routing:
      overall_status == 'active'               → ReplyAll (original recipients kept)
      overall_status in ('partial','not_found') → ESCALATION_TO / ESCALATION_CC only

    Subject tagging:
      TAG_ACTIVE / TAG_PARTIAL / TAG_NOT_FOUND written into subject so future
      runs can read prior outcome from Sent Items without any external state.
      Revalidation runs additionally prepend '[Revalidated] ' for reviewer clarity.
    """
    tag_map = {'active': TAG_ACTIVE, 'partial': TAG_PARTIAL, 'not_found': TAG_NOT_FOUND}
    tag     = tag_map.get(overall_status, TAG_ACTIVE)
    prefix  = '[Revalidated] ' if is_revalidation else ''

    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"{prefix}{tag} {mail_item.Subject}"

        if overall_status in ('partial', 'not_found'):
            # Override recipients — escalation list only.
            # The original requestor's thread receives NO reply for problem cases.
            if ESCALATION_TO:
                reply.To = '; '.join(ESCALATION_TO)
            if ESCALATION_CC:
                reply.CC = '; '.join(ESCALATION_CC)
            _log(f"      📧 Escalation routing | "
                 f"To: {reply.To} | CC: {reply.CC or '(none)'}")
        else:
            _log(f"      📧 ReplyAll (Active result)")

        reply.Save()
        hosts_label = ', '.join(hostname_list)
        _log(f"      ✅ Draft saved [{tag}] for: {hosts_label}"
             f"{' — REVALIDATION' if is_revalidation else ''}")
        return True

    except Exception as e:
        _log(f"      ❌ Failed to create draft: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# OUTLOOK SETUP
# ══════════════════════════════════════════════════════════════════════════════

def get_outlook_folders():
    """
    Connects to the running Outlook instance and returns (inbox, drafts, sent).
    Returns (None, None, None) if Outlook is not open — run aborts cleanly.
    """
    try:
        outlook    = win32com.client.Dispatch('Outlook.Application')
        ns         = outlook.GetNamespace('MAPI')
        main_inbox = ns.GetDefaultFolder(6)   # 6  = Inbox
        drafts     = ns.GetDefaultFolder(16)  # 16 = Drafts
        sent       = ns.GetDefaultFolder(5)   # 5  = Sent Items

        if SIGNOFF_FOLDER_NAME:
            try:
                inbox = main_inbox.Folders[SIGNOFF_FOLDER_NAME]
                _log(f"📁 Subfolder: Inbox\\{SIGNOFF_FOLDER_NAME}")
            except Exception:
                _log(f"⚠️  Subfolder '{SIGNOFF_FOLDER_NAME}' not found — using full Inbox.")
                inbox = main_inbox
        else:
            inbox = main_inbox
            _log("📁 Monitoring: Full Inbox (no subfolder configured)")

        return inbox, drafts, sent
    except Exception as e:
        _log(f"❌ Could not connect to Outlook: {e}. Is Outlook open and logged in?")
        return None, None, None


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _log("=" * 65)
    _log("🚀 QRadar Signoff Auto-Draft starting...")
    _log(f"   Normal window  : {LOOKBACK_HOURS}h")
    _log(f"   Reval window   : {REVALIDATION_WINDOW_DAYS}d | "
         f"Cooldown: {REVALIDATION_COOLDOWN_DAYS}d")
    _log(f"   Folder         : {SIGNOFF_FOLDER_NAME or 'Full Inbox'}")
    _log(f"   Escalation To  : {ESCALATION_TO or 'ReplyAll fallback'}")
    _log(f"   Escalation CC  : {ESCALATION_CC or '(none)'}")
    _log(f"   Dashboard data : {SIGNOFF_DATA_PATH}")
    _log(f"   MODE           : DRAFT ONLY — nothing is sent automatically")

    if not acquire_lock():
        return

    try:
        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        if not test_qradar_connection():
            _log("❌ QRadar unreachable — aborting. All emails left untouched.")
            return

        fetch_log_source_types()

        # Scan window = broader of normal window vs revalidation window so a
        # single inbox pass covers both fresh signoffs and aged problem cases.
        scan_hours      = max(LOOKBACK_HOURS, REVALIDATION_WINDOW_DAYS * 24)
        normal_cutoff   = datetime.now() - timedelta(hours=LOOKBACK_HOURS)
        cooldown_cutoff = datetime.now() - timedelta(days=REVALIDATION_COOLDOWN_DAYS)
        cutoff_str      = (datetime.now() - timedelta(hours=scan_hours)
                           ).strftime('%m/%d/%Y %I:%M %p')

        inbox_items = list(inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'"))
        _log(f"\n📬 Scan: {len(inbox_items)} email(s) in last {scan_hours // 24}d")

        processed = skipped = drafted = revalidated = 0

        for mail_item in inbox_items:
            try:
                if mail_item.Class != 43:   # 43 = olMail
                    continue
            except Exception:
                continue

            try:
                subject = mail_item.Subject or ''
            except Exception:
                continue

            # ── Subject guards ───────────────────────────────────────────────
            passed, reason = passes_subject_guards(subject)
            if not passed:
                skipped += 1
                _log(f"   ⏭️  SKIP ({reason}): '{subject[:60]}'")
                continue

            # ── Sender guards ────────────────────────────────────────────────
            try:
                sender = mail_item.SenderEmailAddress or ''
            except Exception:
                sender = ''

            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                _log(f"   ⏭️  SKIP (own address): '{subject[:60]}'")
                continue

            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"   ⏭️  SKIP (sender not in allowlist — {sender}): '{subject[:60]}'")
                continue

            # ── Body DL check ────────────────────────────────────────────────
            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"   ⏭️  SKIP ('{TRIGGER_DL}' not in body): '{subject[:60]}'")
                continue

            # ── Hostname extraction ──────────────────────────────────────────
            hostname_list = extract_hostnames(subject)
            if not hostname_list:
                skipped += 1
                _log(f"   ⏭️  SKIP (no hostnames found after separator): '{subject[:60]}'")
                continue

            received_dt       = _com_dt_to_py(mail_item.ReceivedTime)
            in_normal_window  = received_dt is not None and received_dt >= normal_cutoff

            _log(f"\n🔹 Candidate: '{subject[:70]}'")
            _log(f"      Sender   : {sender}")
            _log(f"      Hosts    : {hostname_list} ({len(hostname_list)} host(s))")
            _log(f"      Received : {received_dt} "
                 f"({'normal' if in_normal_window else 'reval'} window)")

            # ── Conversation state ───────────────────────────────────────────
            last_tag, last_dt = check_conversation_status(mail_item, sent, drafts)

            if last_tag is None:
                if not in_normal_window:
                    skipped += 1
                    _log(f"      ⏭️  SKIP (old email, no prior tag — out of scope)")
                    continue
                is_revalidation = False
                _log(f"      🆕 New signoff")

            elif last_tag in (TAG_ACTIVE, 'legacy'):
                skipped += 1
                last_str = last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'
                _log(f"      ⏭️  SKIP (Active confirmed on {last_str} — permanent)")
                continue

            elif last_tag in REVALIDATABLE_TAGS:
                if last_dt and last_dt >= cooldown_cutoff:
                    skipped += 1
                    days_ago = (datetime.now() - last_dt).days
                    _log(f"      ⏭️  SKIP (cooldown — last checked {days_ago}d ago, "
                         f"min gap {REVALIDATION_COOLDOWN_DAYS}d)")
                    continue
                last_str = last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'
                _log(f"      🔄 Revalidating {last_tag} from {last_str}")
                is_revalidation = True

            else:
                skipped += 1
                _log(f"      ⏭️  SKIP (unrecognised tag state: '{last_tag}')")
                continue

            # ── Query QRadar + build reply ───────────────────────────────────
            html_body, overall_status, host_records = build_all_hosts_reply(hostname_list)
            _log(f"      🏷️  Overall outcome: {overall_status.upper()}")

            # ── Save draft ───────────────────────────────────────────────────
            success = create_draft_reply(
                mail_item, html_body, hostname_list,
                overall_status=overall_status,
                is_revalidation=is_revalidation,
            )

            if success:
                drafted += 1
                if is_revalidation:
                    revalidated += 1
                # Write dashboard data — non-critical, never aborts on failure
                write_signoff_record(
                    email_subject  = subject,
                    sender         = sender,
                    host_records   = host_records,
                    overall_status = overall_status,
                    is_revalidation= is_revalidation,
                    prior_status   = last_tag,
                )
            processed += 1

        _log(f"\n{'=' * 65}")
        _log(f"✅ Run complete — "
             f"{processed} processed | {drafted} drafted "
             f"({revalidated} reval) | {skipped} skipped")
        _log(f"   Drafts in Drafts folder — review and send manually.")
        _log(f"{'=' * 65}\n")

    finally:
        release_lock()


if __name__ == '__main__':
    main()
