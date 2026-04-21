"""
QRadar Signoff Auto-Draft
─────────────────────────
Scans Outlook for SIEM signoff emails, queries QRadar, saves draft replies.

New in this version:
  • Paths auto-created on first run — printed at end so you can hardcode them.
  • Runtime dedup — same hostname set processed once per run, no repeat drafts.
  • Cross-run hostname dedup — data file used as secondary state store when
    ConversationID matching fails (e.g. separate email threads for same hosts).
  • Dashboard HTML auto-generated and opened in browser at end of every run.
  • Multi-hostname conversation-status fix — falls back to data-file lookup.

THIS SCRIPT IS DRAFT-ONLY. reply.Save() is called, NEVER reply.Send().
"""

import json
import os
import urllib3
import uuid
import webbrowser
import win32com.client
import requests

from datetime import datetime, timedelta

# ─── PATH AUTO-CONFIGURATION ────────────────────────────────────────────────
# All files are created next to this script on first run.
# Paths are printed at the end of each run — paste them here to hardcode.
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
LOOKBACK_DAYS              = 30   # Inbox scan depth
SENT_SCAN_DAYS             = 90   # Sent Items scan for Active tag detection
REVALIDATION_COOLDOWN_DAYS = 0    # 0 = re-check every run; raise to 1-2 later
# Cross-run dedup: if data file shows Active for the SAME hostname set
# within this many days, skip without querying QRadar again.
ACTIVE_SKIP_DAYS           = 30

# ─── SENDER / DL GUARDS ──────────────────────────────────────────────────────
ALLOWED_SENDERS    = []                          # empty = allow all
YOUR_EMAIL_ADDRESS = 'youremail@yourorg.com'
TRIGGER_DL         = '@SOC-DL@yourorg.com'       # '' = disabled

# ─── ESCALATION ──────────────────────────────────────────────────────────────
ESCALATION_TO = ['onboarding-owner@yourorg.com']
ESCALATION_CC = ['@SOC-DL@yourorg.com']

# ─── OUTLOOK FOLDERS ─────────────────────────────────────────────────────────
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'   # set None for full Inbox

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

# Runtime dedup — hostname frozensets drafted in THIS execution only
_runtime_drafted_hosts: set = set()


# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _log(message):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    try:
        with open(RUN_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception as e:
        print(f"WARNING: Could not write to log: {e}")


def _ensure_paths():
    """Creates all required files if they don't already exist."""
    for path in (RUN_LOG_PATH, SIGNOFF_DATA_PATH):
        if not os.path.exists(path):
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    if path.endswith('.json'):
                        json.dump({'schema_version': 2, 'entries': []}, f, indent=2)
                _log(f"Created: {path}")
            except Exception as e:
                _log(f"WARNING: Could not create {path}: {e}")


def _print_paths():
    """Prints all managed file paths at end of run for config copy-paste."""
    _log("")
    _log("-" * 65)
    _log("FILE PATHS — paste into script config to hardcode:")
    _log(f"  RUN_LOG_PATH      = r'{RUN_LOG_PATH}'")
    _log(f"  LOCKFILE_PATH     = r'{LOCKFILE_PATH}'")
    _log(f"  SIGNOFF_DATA_PATH = r'{SIGNOFF_DATA_PATH}'")
    _log(f"  DASHBOARD_PATH    = r'{DASHBOARD_PATH}'")
    _log("-" * 65)


def acquire_lock():
    if os.path.exists(LOCKFILE_PATH):
        _log("WARNING: Lockfile exists — another instance may be running. Exiting.")
        return False
    try:
        with open(LOCKFILE_PATH, 'w') as f:
            f.write(str(os.getpid()))
        return True
    except Exception as e:
        _log(f"ERROR: Could not create lockfile: {e}")
        return False


def release_lock():
    try:
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
    except Exception as e:
        _log(f"WARNING: Could not remove lockfile: {e}")


def _com_dt_to_py(com_dt):
    if com_dt is None:
        return None
    try:
        return datetime(com_dt.year, com_dt.month, com_dt.day,
                        com_dt.hour, com_dt.minute, com_dt.second)
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════════════════
# RUNTIME + CROSS-RUN DEDUP
# ══════════════════════════════════════════════════════════════════════════════

def _host_key(hostname_list):
    """Canonical, order-insensitive key for a hostname set."""
    return frozenset(h.upper().strip() for h in hostname_list)


def is_drafted_this_run(hostname_list):
    """True if this exact hostname set was already processed this run."""
    return _host_key(hostname_list) in _runtime_drafted_hosts


def mark_drafted_this_run(hostname_list):
    _runtime_drafted_hosts.add(_host_key(hostname_list))


def get_prior_status_from_data(hostname_list):
    """
    Fallback dedup for when ConversationID matching fails — typically when
    multiple separate email threads cover the exact same hostname set.

    Looks up signoff_data.json for the most recent entry with a matching
    hostname set (order-insensitive). Returns (status_str, datetime) or
    (None, None) if not found.
    """
    if not os.path.exists(SIGNOFF_DATA_PATH):
        return None, None
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return None, None

    key         = _host_key(hostname_list)
    best_status = None
    best_dt     = None

    for entry in data.get('entries', []):
        if entry.get('manually_resolved'):
            continue
        entry_hosts = [h.get('hostname', '') for h in entry.get('hosts', [])]
        if _host_key(entry_hosts) != key:
            continue
        try:
            entry_dt = datetime.fromisoformat(entry['timestamp'])
        except Exception:
            continue
        if best_dt is None or entry_dt > best_dt:
            best_status = entry.get('overall_status')
            best_dt     = entry_dt

    return best_status, best_dt


# ══════════════════════════════════════════════════════════════════════════════
# QRADAR
# ══════════════════════════════════════════════════════════════════════════════

def test_qradar_connection():
    _log("Testing QRadar connection...")
    try:
        resp = requests.get(
            f"{QRADAR_HOST.rstrip('/')}/api/help/versions",
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL, timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'},
        )
        if resp.status_code == 200:
            _log("QRadar connection OK.")
            return True
        if resp.status_code == 401:
            _log("ERROR: Auth failed — check QRADAR_USERNAME / QRADAR_PASSWORD.")
            return False
        _log(f"WARNING: Unexpected HTTP {resp.status_code}")
        return False
    except Exception as e:
        _log(f"ERROR: Connection failed: {e}")
        return False


def fetch_log_source_types():
    _log("Fetching Log Source Types...")
    try:
        resp = requests.get(
            f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
            f"/log_source_management/log_source_types",
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL, timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'},
        )
        if resp.status_code == 200:
            for t in resp.json():
                if t.get('id') is not None:
                    LOG_SOURCE_TYPES_CACHE[t['id']] = t.get('name', '')
            _log(f"Cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
        else:
            _log(f"WARNING: HTTP {resp.status_code} fetching Log Source Types.")
    except Exception as e:
        _log(f"ERROR: {e}")


def _safe_timestamp(ts):
    if not ts:
        return 'No events recorded', 'No Activity', None
    try:
        ts = int(ts) if isinstance(ts, float) else ts
        s  = ts / 1000.0 if ts > 4102444800 else ts
        if not (_MIN_TS < s <= _MAX_TS):
            return f'Invalid: {ts}', 'Unknown', None
        dt   = datetime.fromtimestamp(s)
        days = (datetime.now() - dt).days
        act  = 'Active' if dt > datetime.now() - timedelta(days=ACTIVITY_THRESHOLD_DAYS) else 'Inactive'
        return dt.strftime('%Y-%m-%d %H:%M:%S'), act, days
    except Exception:
        return f'Invalid: {ts}', 'Unknown', None


def query_all_log_sources_readonly(hostname):
    clean = str(hostname).replace('"', '').replace("'", "").strip()
    try:
        resp = requests.get(
            f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
            f"/log_source_management/log_sources",
            params={'filter': f'name ilike "%{clean}%"'},
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL, timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'},
        )
        if resp.status_code != 200:
            return {'status': f'API Error {resp.status_code}', 'sources': []}
        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', 'sources': []}
        sources = []
        for src in ls_data:
            tid = src.get('type_id')
            last_seen, activity, days_ago = _safe_timestamp(src.get('last_event_time'))
            sources.append({
                'name':      src.get('name', hostname),
                'ls_type':   LOG_SOURCE_TYPES_CACHE.get(tid, f'Unknown TypeID:{tid}'),
                'enabled':   src.get('enabled', False),
                'last_seen': last_seen,
                'activity':  activity,
                'days_ago':  days_ago,
            })
        return {'status': 'Found', 'sources': sources}
    except Exception as e:
        return {'status': f'Error: {str(e)[:80]}', 'sources': []}


def validate_expected_types(result, required_types):
    sources = result.get('sources', [])
    out = []
    for kw in required_types:
        words   = kw.lower().split()
        matched = [s for s in sources
                   if all(w in s.get('ls_type', '').lower() for w in words)]
        if not matched:
            out.append({'expected': kw, 'found': False,
                        'ls_type': None, 'ls_name': None,
                        'last_seen': None, 'days_ago': None})
            continue
        enabled  = sorted([s for s in matched if s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        disabled = sorted([s for s in matched if not s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        best = (enabled or disabled)[0]
        out.append({'expected': kw, 'found': True,
                    'ls_type': best.get('ls_type'), 'ls_name': best.get('name'),
                    'last_seen': best.get('last_seen'), 'days_ago': best.get('days_ago')})
    return out


def detect_os_group(sources):
    if not OS_TYPE_GROUPS:
        return None, None
    for gname, rules in OS_TYPE_GROUPS.items():
        req = rules.get('required', [])
        if req:
            sig = req[0].lower().split()
            if any(all(w in s.get('ls_type', '').lower() for w in sig) for s in sources):
                return gname, rules
    return None, None


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL GUARDS
# ══════════════════════════════════════════════════════════════════════════════

def is_sender_allowed(addr):
    if not ALLOWED_SENDERS:
        return True
    if not addr:
        return False
    a = addr.strip().lower()
    for e in ALLOWED_SENDERS:
        e = e.strip().lower()
        if e.startswith('@') and a.endswith(e):
            return True
        if not e.startswith('@') and a == e:
            return True
    return False


def passes_subject_guards(subject):
    """
    Guards: outcome tag present → skip. No separator → skip. Keyword absent → skip.
    RE/FW prefixes stripped for keyword matching only — not rejected outright.
    Rationale: [processed tag + conversation state handle all dedup; rejecting
    RE/FW was creating gaps when threads surface replies before originals.
    """
    if not subject:
        return False, "empty subject"
    s  = subject.strip()
    sl = s.lower()
    if '[processed' in sl:
        return False, "already tagged"
    if SUBJECT_SEPARATOR not in s:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found"
    left = s.split(SUBJECT_SEPARATOR)[0].strip()
    for pfx in ('re:', 'fw:', 'fwd:'):
        if left.lower().startswith(pfx):
            left = left[len(pfx):].strip()
    if SUBJECT_KEYWORD.lower() not in left.lower():
        return False, f"keyword '{SUBJECT_KEYWORD}' not found"
    return True, "ok"


def extract_hostnames(subject):
    """Extracts hostnames right of the first separator. Works with RE:/FW: prefixes."""
    parts = subject.split(SUBJECT_SEPARATOR, 1)
    if len(parts) < 2:
        return []
    return [h.strip() for h in parts[1].split(SUBJECT_SEPARATOR) if h.strip()]


def body_contains_dl(mail_item):
    if not TRIGGER_DL.strip():
        return True
    dl = TRIGGER_DL.strip().lower()
    try:
        return (dl in (mail_item.Body or '').lower() or
                dl in (mail_item.HTMLBody or '').lower())
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════════════
# CONVERSATION STATUS
# ══════════════════════════════════════════════════════════════════════════════

def _tag_from_subject(subject):
    s = (subject or '').lower()
    if TAG_NOT_FOUND.lower() in s: return TAG_NOT_FOUND
    if TAG_PARTIAL.lower()   in s: return TAG_PARTIAL
    if TAG_ACTIVE.lower()    in s: return TAG_ACTIVE
    if '[processed]'         in s: return 'legacy'
    return None


def check_conversation_status(mail_item, sent_folder, drafts_folder, hostname_list):
    """
    Two-layer dedup:

    Layer 1 — ConversationID scan of Sent + Drafts.
      Handles same-thread re-runs perfectly.

    Layer 2 — data file hostname-set lookup.
      Catches cross-thread cases: separate email threads for the same
      hostnames have different ConversationIDs so Layer 1 misses them.
      This was the root cause of the multi-hostname repeated-draft bug.

    Returns (tag, datetime) of most recent matching outcome, or (None, None).
    """
    conv_id  = mail_item.ConversationID
    last_tag = None
    last_dt  = None

    def _update(tag, dt):
        nonlocal last_tag, last_dt
        if tag and (last_dt is None or (dt and dt > last_dt)):
            last_tag, last_dt = tag, dt

    # Layer 1a — Sent Items (bounded to SENT_SCAN_DAYS for performance)
    cutoff = (datetime.now() - timedelta(days=SENT_SCAN_DAYS)).strftime('%m/%d/%Y %I:%M %p')
    try:
        for item in sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff}'"):
            try:
                if item.ConversationID == conv_id:
                    _update(_tag_from_subject(item.Subject), _com_dt_to_py(item.SentOn))
            except Exception:
                continue
    except Exception as e:
        _log(f"      WARNING: Sent scan error: {e}")

    # Layer 1b — Drafts (no time bound; Drafts folder is small)
    try:
        for item in drafts_folder.Items:
            try:
                if item.ConversationID == conv_id:
                    _update(_tag_from_subject(item.Subject),
                            _com_dt_to_py(item.LastModificationTime))
            except Exception:
                continue
    except Exception as e:
        _log(f"      WARNING: Drafts scan error: {e}")

    # Layer 2 — data file fallback for cross-thread hostname dedup
    if last_tag is None and hostname_list:
        data_status, data_dt = get_prior_status_from_data(hostname_list)
        if data_status and data_dt:
            tag_map = {
                'active':    TAG_ACTIVE,
                'partial':   TAG_PARTIAL,
                'not_found': TAG_NOT_FOUND,
            }
            fallback_tag = tag_map.get(data_status)
            if fallback_tag:
                _log(f"      INFO: Layer-2 fallback — data file: '{data_status}' on "
                     f"{data_dt.strftime('%Y-%m-%d')} for {hostname_list}")
                _update(fallback_tag, data_dt)

    return last_tag, last_dt


# ══════════════════════════════════════════════════════════════════════════════
# HTML EMAIL BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _build_host_html_section(hostname, qradar_result):
    status  = qradar_result.get('status')
    sources = qradar_result.get('sources', [])

    if status != 'Found' or not sources:
        section = f"""
        <div style="margin-bottom:20px;border:1px solid #f5c6c6;border-radius:8px;overflow:hidden;">
          <div style="background:#c0392b;color:#fff;padding:9px 14px;font-size:13px;font-weight:700;">
            &#x2716;&nbsp; {hostname} &mdash; Not Found in QRadar
          </div>
          <div style="padding:12px 14px;font-size:12px;color:#555;">
            <b>{hostname}</b> was not found in QRadar. Please ensure the asset is onboarded.
          </div>
        </div>"""
        return section, 'not_found', [], None

    group_name, group_rules = detect_os_group(sources)
    type_records = []

    if OS_TYPE_GROUPS and group_name:
        validation  = validate_expected_types(qradar_result, group_rules.get('required', []))
        any_missing = any(not r['found'] for r in validation)
        any_silent  = any(r['found'] and r['days_ago'] is None for r in validation)
        any_problem = any_missing or any_silent
        host_status = 'partial' if any_problem else 'active'
        os_label    = f' ({group_name})'

        if not any_problem:
            banner_bg, banner_txt = '#1a7a4a', f'&#x2714;&nbsp; {hostname}{os_label} &mdash; Confirmed Reporting on SIEM'
        elif any_missing:
            n = sum(1 for r in validation if r['found'])
            banner_bg  = '#c87800'
            banner_txt = f'&#x26A0;&nbsp; {hostname}{os_label} &mdash; {n}/{len(validation)} required log sources found'
        else:
            banner_bg  = '#c87800'
            banner_txt = f'&#x26A0;&nbsp; {hostname}{os_label} &mdash; Log sources present but not yet reporting'

        rows = ''
        for r in validation:
            if not r['found']:
                icon, row_bg, status_cell, icon_color = '&#x2716;', '#fff5f5', '<span style="color:#c0392b;font-weight:600;">Missing &mdash; requires onboarding</span>', '#c0392b'
            elif r['days_ago'] is None:
                icon, row_bg, status_cell, icon_color = '&#x26A0;', '#fffbf0', '<span style="color:#c87800;font-weight:600;">No events recorded yet</span>', '#c87800'
            else:
                d_str = 'Today' if r['days_ago'] == 0 else f"{r['days_ago']}d ago"
                icon, row_bg, icon_color = '&#x2714;', '#f0faf4', '#1a7a4a'
                status_cell = (f'<span style="color:#1a7a4a;font-weight:600;">Active</span>'
                               f'&nbsp;<span style="color:#888;font-size:11px;">({d_str})</span>')
            rows += f"""
            <tr style="background:{row_bg};">
              <td style="padding:7px 10px;font-size:12px;color:{icon_color};font-weight:700;text-align:center;width:22px;">{icon}</td>
              <td style="padding:7px 10px;font-size:12px;font-weight:600;color:#333;">{r['expected']}</td>
              <td style="padding:7px 10px;font-size:12px;color:#555;">{r.get('ls_name') or '&mdash;'}</td>
              <td style="padding:7px 10px;font-size:12px;color:#555;">{r.get('last_seen') or '&mdash;'}</td>
              <td style="padding:7px 10px;font-size:12px;">{status_cell}</td>
            </tr>"""
            type_records.append({'expected': r['expected'], 'found': r['found'], 'days_ago': r['days_ago']})

        detail = f"""
        <table style="width:100%;border-collapse:collapse;">
          <tr style="background:#f5f5f5;">
            <th style="padding:6px 10px;font-size:11px;color:#888;text-align:left;border-bottom:1px solid #ddd;width:22px;"></th>
            <th style="padding:6px 10px;font-size:11px;color:#888;text-align:left;border-bottom:1px solid #ddd;">Log Source Type</th>
            <th style="padding:6px 10px;font-size:11px;color:#888;text-align:left;border-bottom:1px solid #ddd;">Log Source Name</th>
            <th style="padding:6px 10px;font-size:11px;color:#888;text-align:left;border-bottom:1px solid #ddd;">Last Event</th>
            <th style="padding:6px 10px;font-size:11px;color:#888;text-align:left;border-bottom:1px solid #ddd;">Status</th>
          </tr>{rows}
        </table>"""
    else:
        if OS_TYPE_GROUPS and not group_name:
            _log(f"      WARNING: OS undetected for {hostname} — simple mode.")
        enabled  = sorted([s for s in sources if s.get('enabled')],     key=lambda x: x.get('days_ago') or 99999)
        disabled = sorted([s for s in sources if not s.get('enabled')], key=lambda x: x.get('days_ago') or 99999)
        best        = (enabled or disabled or [None])[0]
        host_status = 'active'
        banner_bg   = '#1a7a4a'
        banner_txt  = f'&#x2714;&nbsp; {hostname} &mdash; Confirmed Reporting on SIEM'
        group_name  = None
        if best:
            dv     = best.get('days_ago')
            ds     = 'Today' if dv == 0 else (f"{dv} days ago" if dv is not None else 'N/A')
            detail = f"""
            <table style="width:100%;border-collapse:collapse;">
              <tr><td style="padding:7px 10px;font-size:12px;color:#555;width:160px;border-bottom:1px solid #eee;">Log Source Name</td>
                  <td style="padding:7px 10px;font-size:12px;font-weight:600;color:#222;border-bottom:1px solid #eee;">{best.get('name','N/A')}</td></tr>
              <tr><td style="padding:7px 10px;font-size:12px;color:#555;border-bottom:1px solid #eee;">Log Source Type</td>
                  <td style="padding:7px 10px;font-size:12px;color:#333;border-bottom:1px solid #eee;">{best.get('ls_type','N/A')}</td></tr>
              <tr><td style="padding:7px 10px;font-size:12px;color:#555;">Last Event</td>
                  <td style="padding:7px 10px;font-size:12px;color:#333;">{best.get('last_seen','N/A')}
                  &nbsp;<span style="color:#888;font-size:11px;">({ds})</span></td></tr>
            </table>"""
        else:
            detail = ''

    section = f"""
    <div style="margin-bottom:20px;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">
      <div style="background:{banner_bg};color:#fff;padding:9px 14px;font-size:13px;font-weight:700;">{banner_txt}</div>
      <div style="padding:0;">{detail}</div>
    </div>"""
    return section, host_status, type_records, group_name


def _build_full_reply_html(hostname_list, host_sections, host_statuses, run_time):
    badge_cfg = {'active': ('#1a7a4a', '&#x2714;'), 'partial': ('#c87800', '&#x26A0;'), 'not_found': ('#c0392b', '&#x2716;')}
    badges = ''.join(
        f'<span style="display:inline-block;background:{badge_cfg.get(hs,("#555","?"))[0]};color:#fff;'
        f'padding:4px 12px;border-radius:12px;font-size:11px;font-weight:600;margin:0 4px 6px 0;">'
        f'{badge_cfg.get(hs,("#555","?"))[1]}&nbsp;{hn}</span>'
        for hn, hs in zip(hostname_list, host_statuses)
    )
    count_label = f"{len(hostname_list)} host{'s' if len(hostname_list)!=1 else ''} checked"
    return f"""<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;color:#222;font-size:13px;line-height:1.6;margin:0;padding:0;">
  <div style="max-width:680px;padding:20px 0;">
    <p style="margin:0 0 14px 0;">Hi,</p>
    <p style="margin:0 0 10px 0;color:#555;font-size:12px;">Results for your SIEM Security Signoff request &mdash; {count_label}.</p>
    <div style="margin-bottom:18px;">{badges}</div>
    {''.join(host_sections)}
    <p style="margin:20px 0 4px 0;color:#888;font-size:11px;">
      Automated response from the SIEM monitoring system.<br>Checked against QRadar on {run_time}.
    </p>
    <p style="margin:14px 0 0 0;">Regards,<br><span style="font-weight:700;">Cyberdefence</span></p>
  </div>
</body></html>"""


def build_all_hosts_reply(hostname_list):
    run_time       = datetime.now().strftime('%d %B %Y, %H:%M')
    host_sections  = []
    host_statuses  = []
    host_records   = []
    overall_status = 'active'

    for hostname in hostname_list:
        _log(f"      Querying QRadar for [{hostname}]...")
        qr = query_all_log_sources_readonly(hostname)
        _log(f"      [{hostname}] {qr['status']} | {len(qr.get('sources',[]))} sources")
        section, host_status, type_records, os_group = _build_host_html_section(hostname, qr)
        host_sections.append(section)
        host_statuses.append(host_status)
        if STATUS_PRIORITY.get(host_status, 0) > STATUS_PRIORITY.get(overall_status, 0):
            overall_status = host_status
        host_records.append({'hostname': hostname, 'status': host_status,
                             'os_group': os_group, 'type_results': type_records})
        _log(f"      [{hostname}] -> {host_status.upper()}")

    return _build_full_reply_html(hostname_list, host_sections, host_statuses, run_time), overall_status, host_records


# ══════════════════════════════════════════════════════════════════════════════
# DATA STORE
# ══════════════════════════════════════════════════════════════════════════════

def write_signoff_record(email_subject, sender, host_records,
                         overall_status, is_revalidation, prior_status):
    record = {
        'run_id':            f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}",
        'timestamp':         datetime.now().isoformat(),
        'email_subject':     email_subject,
        'sender':            sender,
        'is_revalidation':   is_revalidation,
        'prior_status':      prior_status,
        'overall_status':    overall_status,
        'hosts':             host_records,
        'notes':             '',
        'manually_resolved': False,
    }
    data = {'schema_version': 2, 'entries': []}
    if os.path.exists(SIGNOFF_DATA_PATH):
        try:
            with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            _log(f"WARNING: Data read error (will overwrite): {e}")
    data['entries'].append(record)
    try:
        with open(SIGNOFF_DATA_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        _log(f"      Record written ({overall_status})")
    except Exception as e:
        _log(f"WARNING: Data write error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# DRAFT CREATOR
# ══════════════════════════════════════════════════════════════════════════════

def create_draft_reply(mail_item, html_body, hostname_list, overall_status, is_revalidation=False):
    tag_map = {'active': TAG_ACTIVE, 'partial': TAG_PARTIAL, 'not_found': TAG_NOT_FOUND}
    tag     = tag_map.get(overall_status, TAG_ACTIVE)
    prefix  = '[Revalidated] ' if is_revalidation else ''
    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"{prefix}{tag} {mail_item.Subject}"
        if overall_status in ('partial', 'not_found'):
            if ESCALATION_TO: reply.To = '; '.join(ESCALATION_TO)
            if ESCALATION_CC: reply.CC = '; '.join(ESCALATION_CC)
            _log(f"      Escalation routing -> To:{reply.To} | CC:{reply.CC or '(none)'}")
        else:
            _log(f"      ReplyAll (Active)")
        reply.Save()
        _log(f"      Draft saved [{tag}]{' — REVAL' if is_revalidation else ''}")
        return True
    except Exception as e:
        _log(f"      ERROR: Draft failed: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def generate_dashboard():
    """
    Reads signoff_data.json, embeds it in a self-contained HTML dashboard,
    writes signoff_dashboard.html next to the script, then opens it in the
    default browser. No server needed.
    """
    _log("Generating dashboard...")
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        _log(f"WARNING: Dashboard skipped — could not read data: {e}")
        return

    json_blob = json.dumps(data, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SIEM Signoff Dashboard</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
  *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0;}}
  :root{{
    --bg:#0d0f14;--surface:#151820;--surface2:#1c2030;--border:#252a38;
    --green:#00e676;--amber:#ffab40;--red:#ef5350;--blue:#448aff;
    --text:#e8eaf0;--muted:#6b7280;
    --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;
  }}
  body{{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;min-height:100vh;}}

  header{{background:var(--surface);border-bottom:1px solid var(--border);padding:18px 32px;display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:100;}}
  header h1{{font-family:var(--mono);font-size:15px;font-weight:600;letter-spacing:1px;color:var(--green);}}
  .hdr-sub{{color:var(--muted);font-size:12px;margin-top:2px;}}
  .spacer{{flex:1;}}

  .controls{{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}}
  .pill-group{{display:flex;border:1px solid var(--border);border-radius:6px;overflow:hidden;}}
  .pill-group button{{background:transparent;border:none;color:var(--muted);font-family:var(--mono);font-size:11px;padding:6px 14px;cursor:pointer;transition:all .15s;}}
  .pill-group button.active{{background:var(--green);color:#000;font-weight:600;}}
  .pill-group button:not(.active):hover{{background:var(--surface2);color:var(--text);}}
  input[type=text]{{background:var(--surface2);border:1px solid var(--border);border-radius:6px;color:var(--text);font-family:var(--mono);font-size:12px;padding:6px 12px;width:220px;outline:none;}}
  input[type=text]:focus{{border-color:var(--blue);}}

  main{{padding:28px 32px;max-width:1400px;margin:0 auto;}}

  .stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:28px;}}
  .card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:20px 24px;position:relative;overflow:hidden;}}
  .card::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;}}
  .card.green::before{{background:var(--green);}} .card.amber::before{{background:var(--amber);}}
  .card.red::before{{background:var(--red);}}     .card.blue::before{{background:var(--blue);}}
  .card .num{{font-family:var(--mono);font-size:36px;font-weight:600;line-height:1;margin-bottom:6px;}}
  .card.green .num{{color:var(--green);}} .card.amber .num{{color:var(--amber);}}
  .card.red .num{{color:var(--red);}}     .card.blue .num{{color:var(--blue);}}
  .card .clabel{{color:var(--muted);font-size:11px;letter-spacing:.5px;text-transform:uppercase;}}
  .card .sub2{{color:var(--muted);font-size:11px;margin-top:4px;}}

  .charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:28px;}}
  @media(max-width:700px){{.charts-grid{{grid-template-columns:1fr;}}}}
  .chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 20px;}}
  .chart-card h3{{font-family:var(--mono);font-size:12px;color:var(--muted);margin-bottom:14px;letter-spacing:.5px;}}
  .chart-row{{display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:12px;}}
  .chart-row .ck{{width:90px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
  .bar-track{{flex:1;background:var(--surface2);border-radius:3px;height:8px;}}
  .bar-fill{{height:100%;border-radius:3px;transition:width .4s;}}
  .chart-row .cv{{width:30px;text-align:right;color:var(--muted);font-family:var(--mono);font-size:11px;}}

  .table-wrap{{background:var(--surface);border:1px solid var(--border);border-radius:10px;overflow:hidden;}}
  .table-header{{padding:14px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:12px;}}
  .table-header h2{{font-size:13px;font-weight:600;font-family:var(--mono);}}
  table{{width:100%;border-collapse:collapse;}}
  th{{background:var(--surface2);color:var(--muted);font-family:var(--mono);font-size:11px;font-weight:600;letter-spacing:.5px;text-align:left;padding:10px 16px;border-bottom:1px solid var(--border);cursor:pointer;user-select:none;white-space:nowrap;}}
  th:hover{{color:var(--text);}}
  td{{padding:11px 16px;border-bottom:1px solid var(--border);vertical-align:middle;font-size:12px;}}
  tr:last-child td{{border-bottom:none;}}
  tr:hover td{{background:var(--surface2);}}

  .badge{{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-family:var(--mono);font-weight:600;white-space:nowrap;}}
  .ba{{background:#00e67622;color:var(--green);border:1px solid #00e67644;}}
  .bp{{background:#ffab4022;color:var(--amber);border:1px solid #ffab4044;}}
  .bn{{background:#ef535022;color:var(--red);border:1px solid #ef535044;}}
  .br{{background:#448aff22;color:var(--blue);border:1px solid #448aff44;}}
  .bx{{background:#6b728022;color:var(--muted);border:1px solid #6b728044;}}

  .hp{{display:inline-block;background:var(--surface2);border:1px solid var(--border);border-radius:4px;padding:2px 8px;font-family:var(--mono);font-size:11px;margin:2px 2px 2px 0;}}

  .btn{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:5px;font-size:11px;padding:4px 10px;cursor:pointer;font-family:var(--mono);transition:all .15s;}}
  .btn:hover{{border-color:var(--blue);color:var(--blue);}}
  .btn-save{{border-color:var(--green);color:var(--green);}}
  .btn-save:hover{{background:var(--green);color:#000;}}
  .btn-exp{{border-color:var(--amber);color:var(--amber);}}
  .btn-exp:hover{{background:var(--amber);color:#000;}}

  .modal-bg{{display:none;position:fixed;inset:0;background:#000a;z-index:200;align-items:center;justify-content:center;}}
  .modal-bg.open{{display:flex;}}
  .modal{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:28px 32px;width:520px;max-width:95vw;max-height:85vh;overflow-y:auto;}}
  .modal h3{{font-family:var(--mono);font-size:14px;margin-bottom:18px;color:var(--blue);}}
  .fr{{margin-bottom:14px;}}
  .fr label{{display:block;color:var(--muted);font-size:11px;margin-bottom:5px;text-transform:uppercase;letter-spacing:.5px;}}
  .fr textarea,.fr select{{width:100%;background:var(--surface2);border:1px solid var(--border);border-radius:6px;color:var(--text);font-family:var(--mono);font-size:12px;padding:8px 12px;outline:none;resize:vertical;}}
  .fr textarea:focus,.fr select:focus{{border-color:var(--blue);}}
  .modal-actions{{display:flex;gap:10px;margin-top:20px;justify-content:flex-end;}}

  .pag{{display:flex;gap:6px;align-items:center;padding:12px 16px;}}
  .pag button{{background:var(--surface2);border:1px solid var(--border);color:var(--muted);border-radius:4px;padding:4px 10px;font-size:12px;cursor:pointer;font-family:var(--mono);}}
  .pag button.active{{background:var(--blue);border-color:var(--blue);color:#fff;}}
  .pag button:disabled{{opacity:.4;cursor:default;}}
  .pag .info{{color:var(--muted);font-size:11px;margin-left:auto;font-family:var(--mono);}}
  .empty{{padding:48px;text-align:center;color:var(--muted);font-family:var(--mono);font-size:13px;}}
</style>
</head>
<body>
<header>
  <div>
    <h1>SIEM SIGNOFF DASHBOARD</h1>
    <div class="hdr-sub" id="lastUpdated"></div>
  </div>
  <div class="spacer"></div>
  <div class="controls">
    <div class="pill-group" id="pg">
      <button onclick="setPeriod('week',this)" class="active">7D</button>
      <button onclick="setPeriod('month',this)">30D</button>
      <button onclick="setPeriod('quarter',this)">90D</button>
      <button onclick="setPeriod('all',this)">ALL</button>
    </div>
    <input type="text" id="search" placeholder="Search host / sender..." oninput="render()">
    <button class="btn btn-exp" onclick="exportData()">&#x2B07; Export JSON</button>
  </div>
</header>
<main>
  <div class="stats" id="statsCards"></div>
  <div class="charts-grid">
    <div class="chart-card"><h3>STATUS BREAKDOWN</h3><div id="statusChart"></div></div>
    <div class="chart-card"><h3>TOP HOSTNAMES (BY REQUESTS)</h3><div id="hostChart"></div></div>
  </div>
  <div class="table-wrap">
    <div class="table-header">
      <h2>SIGNOFF LOG</h2>
      <div class="spacer"></div>
    </div>
    <table>
      <thead><tr>
        <th onclick="sortBy('timestamp')">TIMESTAMP &#x21D5;</th>
        <th onclick="sortBy('hosts')">HOSTNAMES</th>
        <th onclick="sortBy('overall_status')">STATUS &#x21D5;</th>
        <th onclick="sortBy('sender')">SENDER &#x21D5;</th>
        <th>FLAGS</th>
        <th>NOTES</th>
        <th>EDIT</th>
      </tr></thead>
      <tbody id="logBody"></tbody>
    </table>
    <div class="pag" id="pag"></div>
  </div>
</main>

<div class="modal-bg" id="mb" onclick="if(event.target===this)closeMod()">
  <div class="modal">
    <h3>EDIT RECORD</h3>
    <div class="fr">
      <label>Override Status</label>
      <select id="eStatus">
        <option value="">&#x2014; keep current &#x2014;</option>
        <option value="active">Active</option>
        <option value="partial">Partial</option>
        <option value="not_found">Not Found</option>
      </select>
    </div>
    <div class="fr">
      <label>Mark as Manually Resolved</label>
      <select id="eResolved">
        <option value="false">No</option>
        <option value="true">Yes &#x2014; resolved, exclude from dedup</option>
      </select>
    </div>
    <div class="fr">
      <label>Notes (ticket ID, action taken...)</label>
      <textarea id="eNotes" rows="4"></textarea>
    </div>
    <div class="modal-actions">
      <button class="btn" onclick="closeMod()">Cancel</button>
      <button class="btn btn-save" onclick="saveEdit()">Save</button>
    </div>
  </div>
</div>

<script>
const RAW = {json_blob};
let D = JSON.parse(JSON.stringify(RAW));
let period='week', sf='timestamp', sasc=false, page=1, eidx=null;
const PS=15;

function cutoff(){{
  const m={{week:7,month:30,quarter:90,all:36500}};
  return new Date(Date.now()-m[period]*86400000);
}}
function filtered(){{
  const q=document.getElementById('search').value.toLowerCase();
  const c=cutoff();
  return D.entries.filter(e=>{{
    if(new Date(e.timestamp)<c) return false;
    if(q){{
      const h=(e.hosts||[]).map(x=>x.hostname||'').join(' ').toLowerCase();
      if(!h.includes(q)&&!(e.sender||'').toLowerCase().includes(q)&&!(e.notes||'').toLowerCase().includes(q)) return false;
    }}
    return true;
  }});
}}
function srt(arr){{
  return [...arr].sort((a,b)=>{{
    let av,bv;
    if(sf==='timestamp'){{av=a.timestamp;bv=b.timestamp;}}
    else if(sf==='overall_status'){{av=a.overall_status||'';bv=b.overall_status||'';}}
    else if(sf==='sender'){{av=a.sender||'';bv=b.sender||'';}}
    else if(sf==='hosts'){{av=(a.hosts||[]).length;bv=(b.hosts||[]).length;}}
    else{{av='';bv='';}}
    return (av<bv?-1:av>bv?1:0)*(sasc?1:-1);
  }});
}}
function badge(status,resolved,reval){{
  if(resolved) return '<span class="badge bx">RESOLVED</span>';
  const m={{active:'<span class="badge ba">ACTIVE</span>',partial:'<span class="badge bp">PARTIAL</span>',not_found:'<span class="badge bn">NOT FOUND</span>'}};
  return (m[status]||`<span class="badge">${{status}}</span>`)+(reval?' <span class="badge br">REVAL</span>':'');
}}
function fmt(iso){{
  if(!iso) return '&#x2014;';
  const d=new Date(iso);
  return d.toLocaleDateString('en-GB',{{day:'2-digit',month:'short',year:'numeric'}})+' '+d.toLocaleTimeString('en-GB',{{hour:'2-digit',minute:'2-digit'}});
}}
function barChart(el,items,cfn){{
  const mx=Math.max(...items.map(i=>i.v),1);
  el.innerHTML=items.map(i=>`<div class="chart-row">
    <span class="ck" title="${{i.k}}">${{i.k}}</span>
    <div class="bar-track"><div class="bar-fill" style="width:${{Math.round(i.v/mx*100)}}%;background:${{cfn(i.k)}}"></div></div>
    <span class="cv">${{i.v}}</span></div>`).join('');
}}
function renderStats(e){{
  const t=e.length,a=e.filter(x=>x.overall_status==='active'&&!x.manually_resolved).length,
    p=e.filter(x=>x.overall_status==='partial'&&!x.manually_resolved).length,
    n=e.filter(x=>x.overall_status==='not_found'&&!x.manually_resolved).length,
    rv=e.filter(x=>x.is_revalidation).length;
  document.getElementById('statsCards').innerHTML=`
    <div class="card blue"><div class="num">${{t}}</div><div class="clabel">Total Signoffs</div><div class="sub2">${{rv}} revalidations</div></div>
    <div class="card green"><div class="num">${{a}}</div><div class="clabel">Active</div><div class="sub2">${{t?Math.round(a/t*100):0}}% of period</div></div>
    <div class="card amber"><div class="num">${{p}}</div><div class="clabel">Partial</div><div class="sub2">Missing log sources</div></div>
    <div class="card red"><div class="num">${{n}}</div><div class="clabel">Not Found</div><div class="sub2">${{e.filter(x=>x.manually_resolved).length}} resolved</div></div>`;
}}
function renderCharts(e){{
  const sm={{}};
  e.forEach(x=>{{const k=x.manually_resolved?'resolved':(x.overall_status||'?');sm[k]=(sm[k]||0)+1;}});
  const sc={{'active':'var(--green)','partial':'var(--amber)','not_found':'var(--red)','resolved':'var(--muted)'}};
  barChart(document.getElementById('statusChart'),Object.entries(sm).map(([k,v])=>{{return{{k,v}}}}).sort((a,b)=>b.v-a.v),k=>sc[k]||'var(--blue)');
  const hm={{}};
  e.forEach(x=>(x.hosts||[]).forEach(h=>{{if(h.hostname)hm[h.hostname]=(hm[h.hostname]||0)+1;}}));
  barChart(document.getElementById('hostChart'),Object.entries(hm).map(([k,v])=>{{return{{k,v}}}}).sort((a,b)=>b.v-a.v).slice(0,10),()=>'var(--blue)');
}}
function render(){{
  const e=srt(filtered());
  const tot=e.length,pages=Math.max(1,Math.ceil(tot/PS));
  if(page>pages)page=1;
  const sl=e.slice((page-1)*PS,page*PS);
  const body=document.getElementById('logBody');
  if(!sl.length){{body.innerHTML=`<tr><td colspan="7"><div class="empty">No records in this period.</div></td></tr>`;}}
  else{{
    body.innerHTML=sl.map((x,i)=>{{
      const gi=D.entries.indexOf(e[(page-1)*PS+i]);
      const hosts=(x.hosts||[]).map(h=>`<span class="hp">${{h.hostname||'?'}}</span>`).join('');
      const sender=(x.sender||'').split('@')[0]||'&#x2014;';
      const notes=x.notes?`<span title="${{x.notes.replace(/"/g,'&quot;')}}" style="color:var(--blue);cursor:help;">&#x1F4DD;</span>`:'';
      return`<tr>
        <td style="font-family:var(--mono);font-size:11px;white-space:nowrap;color:var(--muted)">${{fmt(x.timestamp)}}</td>
        <td>${{hosts}}</td>
        <td>${{badge(x.overall_status,x.manually_resolved,x.is_revalidation)}}</td>
        <td style="font-family:var(--mono);font-size:11px;color:var(--muted)">${{sender}}</td>
        <td style="font-size:11px;color:var(--muted)">${{x.prior_status?'prior: '+x.prior_status.replace('[Processed-','').replace(']','').toLowerCase():'new'}}</td>
        <td>${{notes}}</td>
        <td><button class="btn" onclick="openMod(${{gi}})">Edit</button></td>
      </tr>`;
    }}).join('');
  }}
  // Pagination
  let ph=`<button onclick="goP(${{page-1}})" ${{page===1?'disabled':''}}>&#x2039; Prev</button>`;
  const s=Math.max(1,page-2),en=Math.min(pages,page+2);
  if(s>1)ph+=`<button onclick="goP(1)">1</button>${{s>2?'<span style="color:var(--muted);padding:0 4px">&#x2026;</span>':''}}`;
  for(let p=s;p<=en;p++)ph+=`<button onclick="goP(${{p}})" class="${{p===page?'active':''}}">${{p}}</button>`;
  if(en<pages)ph+=`${{en<pages-1?'<span style="color:var(--muted);padding:0 4px">&#x2026;</span>':''}}<button onclick="goP(${{pages}})">${{pages}}</button>`;
  ph+=`<button onclick="goP(${{page+1}})" ${{page===pages?'disabled':''}}>Next &#x203A;</button><span class="info">${{tot}} records</span>`;
  document.getElementById('pag').innerHTML=ph;
  renderStats(filtered());
  renderCharts(filtered());
}}
function goP(p){{page=p;render();}}
function sortBy(f){{if(sf===f)sasc=!sasc;else{{sf=f;sasc=false;}}render();}}
function setPeriod(p,btn){{
  period=p;page=1;
  document.querySelectorAll('#pg button').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');render();
}}
function openMod(i){{
  eidx=i;const e=D.entries[i];
  document.getElementById('eStatus').value=e.overall_status||'';
  document.getElementById('eResolved').value=e.manually_resolved?'true':'false';
  document.getElementById('eNotes').value=e.notes||'';
  document.getElementById('mb').classList.add('open');
}}
function closeMod(){{document.getElementById('mb').classList.remove('open');eidx=null;}}
function saveEdit(){{
  if(eidx===null)return;
  const e=D.entries[eidx],s=document.getElementById('eStatus').value;
  if(s)e.overall_status=s;
  e.manually_resolved=document.getElementById('eResolved').value==='true';
  e.notes=document.getElementById('eNotes').value.trim();
  closeMod();render();
}}
function exportData(){{
  const b=new Blob([JSON.stringify(D,null,2)],{{type:'application/json'}});
  Object.assign(document.createElement('a'),{{href:URL.createObjectURL(b),download:'signoff_data.json'}}).click();
}}
const last=D.entries.length?D.entries[D.entries.length-1].timestamp:null;
document.getElementById('lastUpdated').textContent=last?'Last updated: '+fmt(last):'No data yet';
render();
</script>
</body>
</html>"""

    try:
        with open(DASHBOARD_PATH, 'w', encoding='utf-8') as f:
            f.write(html)
        _log(f"Dashboard written: {DASHBOARD_PATH}")
        webbrowser.open(f"file:///{DASHBOARD_PATH.replace(os.sep, '/')}")
    except Exception as e:
        _log(f"WARNING: Dashboard error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# OUTLOOK SETUP
# ══════════════════════════════════════════════════════════════════════════════

def get_outlook_folders():
    try:
        outlook    = win32com.client.Dispatch('Outlook.Application')
        ns         = outlook.GetNamespace('MAPI')
        main_inbox = ns.GetDefaultFolder(6)
        drafts     = ns.GetDefaultFolder(16)
        sent       = ns.GetDefaultFolder(5)
        if SIGNOFF_FOLDER_NAME:
            try:
                inbox = main_inbox.Folders[SIGNOFF_FOLDER_NAME]
                _log(f"Folder: Inbox\\{SIGNOFF_FOLDER_NAME}")
            except Exception:
                _log(f"WARNING: '{SIGNOFF_FOLDER_NAME}' not found — using full Inbox.")
                inbox = main_inbox
        else:
            inbox = main_inbox
        return inbox, drafts, sent
    except Exception as e:
        _log(f"ERROR: Outlook connect failed: {e}")
        return None, None, None


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _ensure_paths()

    _log("=" * 65)
    _log("QRadar Signoff Auto-Draft starting...")
    _log(f"  Inbox: {LOOKBACK_DAYS}d | Sent: {SENT_SCAN_DAYS}d | "
         f"Cooldown: {'off' if not REVALIDATION_COOLDOWN_DAYS else f'{REVALIDATION_COOLDOWN_DAYS}d'} | "
         f"Active-skip: {ACTIVE_SKIP_DAYS}d")
    _log("  Runtime dedup: ON | Cross-thread dedup: ON | Mode: DRAFT ONLY")

    if not acquire_lock():
        return

    try:
        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        if not test_qradar_connection():
            _log("ERROR: QRadar unreachable — aborting. Emails untouched.")
            return

        fetch_log_source_types()

        cutoff_str     = (datetime.now() - timedelta(days=LOOKBACK_DAYS)).strftime('%m/%d/%Y %I:%M %p')
        cooldown_cutoff = (
            datetime.now() - timedelta(days=REVALIDATION_COOLDOWN_DAYS)
            if REVALIDATION_COOLDOWN_DAYS > 0 else None
        )
        active_skip_cutoff = datetime.now() - timedelta(days=ACTIVE_SKIP_DAYS)

        inbox_items = list(inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'"))
        _log(f"\n{len(inbox_items)} email(s) in last {LOOKBACK_DAYS}d")

        processed = skipped = drafted = revalidated = 0

        for mail_item in inbox_items:
            try:
                if mail_item.Class != 43:
                    continue
            except Exception:
                continue

            try:
                subject = mail_item.Subject or ''
            except Exception:
                continue

            # Subject guards
            passed, reason = passes_subject_guards(subject)
            if not passed:
                skipped += 1
                _log(f"  SKIP ({reason}): '{subject[:60]}'")
                continue

            # Sender guards
            try:
                sender = mail_item.SenderEmailAddress or ''
            except Exception:
                sender = ''

            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                continue
            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"  SKIP (sender not allowed): '{subject[:60]}'")
                continue
            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"  SKIP (DL not in body): '{subject[:60]}'")
                continue

            hostname_list = extract_hostnames(subject)
            if not hostname_list:
                skipped += 1
                _log(f"  SKIP (no hostnames): '{subject[:60]}'")
                continue

            # ── Runtime dedup ────────────────────────────────────────────────
            # Prevents 3-4 emails about the same hosts creating 3-4 drafts.
            if is_drafted_this_run(hostname_list):
                skipped += 1
                _log(f"  SKIP (runtime dedup — already handled {hostname_list} this run)")
                continue

            _log(f"\n  Candidate: '{subject[:70]}'")
            _log(f"    Sender: {sender} | Hosts: {hostname_list}")

            # ── Conversation + cross-thread state check ──────────────────────
            last_tag, last_dt = check_conversation_status(
                mail_item, sent, drafts, hostname_list
            )

            if last_tag in (TAG_ACTIVE, 'legacy'):
                if last_dt and last_dt >= active_skip_cutoff:
                    skipped += 1
                    _log(f"    SKIP (Active on {last_dt.strftime('%Y-%m-%d')} — within {ACTIVE_SKIP_DAYS}d)")
                    continue
                _log(f"    Active result >  {ACTIVE_SKIP_DAYS}d old — allowing recheck")
                is_revalidation = True

            elif last_tag in REVALIDATABLE_TAGS:
                if cooldown_cutoff and last_dt and last_dt >= cooldown_cutoff:
                    skipped += 1
                    _log(f"    SKIP (cooldown: {(datetime.now()-last_dt).days}d ago)")
                    continue
                _log(f"    Revalidating {last_tag} from {last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'}")
                is_revalidation = True

            elif last_tag is None:
                _log(f"    New signoff")
                is_revalidation = False

            else:
                skipped += 1
                _log(f"    SKIP (unknown tag: {last_tag})")
                continue

            # ── QRadar query ─────────────────────────────────────────────────
            html_body, overall_status, host_records = build_all_hosts_reply(hostname_list)
            _log(f"    Overall: {overall_status.upper()}")

            success = create_draft_reply(
                mail_item, html_body, hostname_list,
                overall_status=overall_status,
                is_revalidation=is_revalidation,
            )

            if success:
                drafted += 1
                if is_revalidation:
                    revalidated += 1
                mark_drafted_this_run(hostname_list)   # prevents re-draft in same run
                write_signoff_record(
                    email_subject   = subject,
                    sender          = sender,
                    host_records    = host_records,
                    overall_status  = overall_status,
                    is_revalidation = is_revalidation,
                    prior_status    = last_tag,
                )
            processed += 1

        _log(f"\n{'='*65}")
        _log(f"Done — {processed} processed | {drafted} drafted ({revalidated} reval) | {skipped} skipped")

    finally:
        release_lock()
        _print_paths()
        generate_dashboard()


if __name__ == '__main__':
    main()
