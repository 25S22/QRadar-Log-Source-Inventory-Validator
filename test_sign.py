"""
QRadar Signoff Runner  v3.0
────────────────────────────────────────────────────────────────────────────
Scans Outlook for SIEM signoff emails, queries QRadar, saves HTML draft replies,
and writes results to signoff_data.json.

Key behaviours
  • First-email-only policy — any subject starting with RE: / FW: / FWD: is
    skipped immediately.  No prefix-stripping, no chain chasing.
  • Draft-only — reply.Save() is called, NEVER reply.Send().
  • Atomic JSON writes — data file is never left in a corrupt half-written state.
  • Single-instance lock — a lockfile prevents concurrent runs.
  • Runtime dedup — the same hostname set is never drafted twice per run.
  • Conversation dedup — Sent + Drafts folders are scanned for prior outcomes.

Dashboard
  Run signoff_dashboard.py separately to view / edit results.
────────────────────────────────────────────────────────────────────────────
"""

import json
import os
import tempfile
import uuid
import urllib3
import win32com.client
import requests

from datetime import datetime, timedelta


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION  — edit this block, nothing else needs touching
# ══════════════════════════════════════════════════════════════════════════════

# ─── Paths (auto-configured relative to this script) ──────────────────────────
_DIR              = os.path.dirname(os.path.abspath(__file__))
RUN_LOG_PATH      = os.path.join(_DIR, 'signoff_runner.log')
LOCKFILE_PATH     = os.path.join(_DIR, 'signoff.lock')
SIGNOFF_DATA_PATH = os.path.join(_DIR, 'signoff_data.json')

# ─── QRadar ───────────────────────────────────────────────────────────────────
QRADAR_HOST     = os.environ.get('QRADAR_HOST',     'https://your-qradar-host')
QRADAR_USERNAME = os.environ.get('QRADAR_USERNAME', 'your-username')
QRADAR_PASSWORD = os.environ.get('QRADAR_PASSWORD', 'your-password')
VERIFY_SSL      = False          # set True + supply a CA bundle in production

# ─── Subject matching ─────────────────────────────────────────────────────────
SUBJECT_KEYWORD   = 'Security Signoff'   # must appear left of SUBJECT_SEPARATOR
SUBJECT_SEPARATOR = '|'                  # separates keyword from hostname list

# ─── Scan windows ─────────────────────────────────────────────────────────────
LOOKBACK_DAYS    = 30   # how many days back to scan the Inbox
SENT_SCAN_DAYS   = 90   # how many days back to scan Sent + Drafts for prior outcomes
ACTIVE_SKIP_DAYS = 30   # skip re-drafting if an Active result exists within this window

# ─── Sender guards ────────────────────────────────────────────────────────────
# Leave ALLOWED_SENDERS empty ([]) to accept any sender.
# Entries starting with '@' match the whole domain — e.g. '@yourorg.com'.
ALLOWED_SENDERS    = []
YOUR_EMAIL_ADDRESS = 'youremail@yourorg.com'   # your own address — never process self-sent

# TRIGGER_DL: a string that must appear in the email body.
# Set to '' to disable this check entirely.
TRIGGER_DL = '@SOC-DL@yourorg.com'

# ─── Escalation routing (Partial / Not-Found drafts) ──────────────────────────
ESCALATION_TO = ['onboarding-owner@yourorg.com']
ESCALATION_CC = ['@SOC-DL@yourorg.com']

# ─── Outlook ──────────────────────────────────────────────────────────────────
# Sub-folder of Inbox to scan.  Set to None to scan the full Inbox.
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'

# ─── OS type validation ───────────────────────────────────────────────────────
# Map OS group names to required QRadar log source type keywords.
# The first keyword in 'required' is used for OS detection (type name contains it).
# Remove or empty this dict to skip OS validation and use simple mode.
OS_TYPE_GROUPS = {
    'Windows': {'required': ['Microsoft Security', 'WinCollect']},
    'Linux':   {'required': ['Linux OS']},
}

# ─── Outcome subject tags ─────────────────────────────────────────────────────
TAG_ACTIVE    = '[Processed-Active]'
TAG_PARTIAL   = '[Processed-Partial]'
TAG_NOT_FOUND = '[Processed-NotFound]'
REVALIDATABLE_TAGS = {TAG_PARTIAL, TAG_NOT_FOUND}

# ─── QRadar API ───────────────────────────────────────────────────────────────
ACTIVITY_THRESHOLD_DAYS = 7     # days since last event to call a source "Active"
REQUEST_TIMEOUT         = 30    # seconds
_MIN_TS                 = 0
_MAX_TS                 = 2_147_483_647

# ══════════════════════════════════════════════════════════════════════════════
# MODULE-LEVEL STATE
# ══════════════════════════════════════════════════════════════════════════════

LOG_SOURCE_TYPES_CACHE: dict = {}
STATUS_PRIORITY = {'not_found': 2, 'partial': 1, 'active': 0}
_runtime_drafted_hosts: set = set()


# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _log(msg: str) -> None:
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(RUN_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception as exc:
        print(f"WARNING: Log write failed — {exc}")


def _atomic_write_json(path: str, data: dict) -> None:
    """
    Write JSON atomically: dump to a sibling .tmp file, then os.replace() it
    into position.  A crash mid-write leaves the original file intact.
    """
    dir_ = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(dir=dir_, suffix='.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _load_data() -> dict:
    """Load signoff_data.json; return empty schema if missing or corrupt."""
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as exc:
        _log(f"WARNING: Data file unreadable ({exc}) — starting fresh.")
        return {'schema_version': 3, 'entries': []}


def _ensure_data_file() -> None:
    """Initialise signoff_data.json with an empty schema if it does not exist."""
    if not os.path.exists(SIGNOFF_DATA_PATH):
        _atomic_write_json(SIGNOFF_DATA_PATH, {'schema_version': 3, 'entries': []})
        _log(f"Created: {SIGNOFF_DATA_PATH}")


def acquire_lock() -> bool:
    if os.path.exists(LOCKFILE_PATH):
        _log("WARNING: Lockfile present — another instance may be running.  Exiting.")
        return False
    try:
        with open(LOCKFILE_PATH, 'w') as f:
            f.write(str(os.getpid()))
        return True
    except Exception as exc:
        _log(f"ERROR: Cannot create lockfile: {exc}")
        return False


def release_lock() -> None:
    try:
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
    except Exception as exc:
        _log(f"WARNING: Could not remove lockfile: {exc}")


def _com_dt_to_py(com_dt) -> datetime | None:
    if com_dt is None:
        return None
    try:
        return datetime(com_dt.year, com_dt.month, com_dt.day,
                        com_dt.hour, com_dt.minute, com_dt.second)
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════════════════
# RUNTIME DEDUP
# ══════════════════════════════════════════════════════════════════════════════

def _host_key(hostname_list: list) -> frozenset:
    return frozenset(h.upper().strip() for h in hostname_list)


def is_drafted_this_run(hostname_list: list) -> bool:
    return _host_key(hostname_list) in _runtime_drafted_hosts


def mark_drafted_this_run(hostname_list: list) -> None:
    _runtime_drafted_hosts.add(_host_key(hostname_list))


# ══════════════════════════════════════════════════════════════════════════════
# QRADAR
# ══════════════════════════════════════════════════════════════════════════════

def _qradar_get(path: str, params: dict = None) -> requests.Response:
    return requests.get(
        f"{QRADAR_HOST.rstrip('/')}{path}",
        params=params,
        auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
        verify=VERIFY_SSL,
        timeout=REQUEST_TIMEOUT,
        headers={'Accept': 'application/json', 'Version': '14.0'},
    )


def test_qradar_connection() -> bool:
    _log("Testing QRadar connection...")
    try:
        r = _qradar_get('/api/help/versions')
        if r.status_code == 200:
            _log("QRadar connection OK.")
            return True
        if r.status_code == 401:
            _log("ERROR: QRadar auth failed — check QRADAR_USERNAME / QRADAR_PASSWORD.")
        else:
            _log(f"WARNING: QRadar returned HTTP {r.status_code}")
        return False
    except Exception as exc:
        _log(f"ERROR: QRadar unreachable — {exc}")
        return False


def fetch_log_source_types() -> None:
    _log("Fetching Log Source Types...")
    try:
        r = _qradar_get('/api/config/event_sources/log_source_management/log_source_types')
        if r.status_code == 200:
            for t in r.json():
                if t.get('id') is not None:
                    LOG_SOURCE_TYPES_CACHE[t['id']] = t.get('name', '')
            _log(f"Cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
        else:
            _log(f"WARNING: HTTP {r.status_code} fetching Log Source Types.")
    except Exception as exc:
        _log(f"ERROR: {exc}")


def _safe_timestamp(ts) -> tuple:
    """Returns (formatted_str, activity_str, days_ago_int|None)."""
    if not ts:
        return 'No events recorded', 'No Activity', None
    try:
        s = int(ts) / 1000.0 if int(ts) > 4_102_444_800 else int(ts)
        if not (_MIN_TS < s <= _MAX_TS):
            return f'Invalid: {ts}', 'Unknown', None
        dt   = datetime.fromtimestamp(s)
        days = (datetime.now() - dt).days
        act  = 'Active' if days <= ACTIVITY_THRESHOLD_DAYS else 'Inactive'
        return dt.strftime('%Y-%m-%d %H:%M:%S'), act, days
    except Exception:
        return f'Invalid: {ts}', 'Unknown', None


def query_log_sources(hostname: str) -> dict:
    """Query QRadar for all log sources whose name contains the hostname."""
    clean = hostname.replace('"', '').replace("'", '').strip()
    try:
        r = _qradar_get(
            '/api/config/event_sources/log_source_management/log_sources',
            params={'filter': f'name ilike "%{clean}%"'},
        )
        if r.status_code != 200:
            return {'status': f'API Error {r.status_code}', 'sources': []}
        raw = r.json()
        if not raw:
            return {'status': 'Not Found', 'sources': []}
        sources = []
        for src in raw:
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
    except Exception as exc:
        return {'status': f'Error: {str(exc)[:80]}', 'sources': []}


def validate_required_types(result: dict, required_types: list) -> list:
    """Check which required log source type keywords are present in the result."""
    sources = result.get('sources', [])
    out = []
    for kw in required_types:
        words   = kw.lower().split()
        matched = [s for s in sources if all(w in s.get('ls_type', '').lower() for w in words)]
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


def detect_os_group(sources: list) -> tuple:
    """Returns (group_name, group_rules) or (None, None) if undetected."""
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

def is_sender_allowed(addr: str) -> bool:
    if not ALLOWED_SENDERS:
        return True
    a = (addr or '').strip().lower()
    for entry in ALLOWED_SENDERS:
        e = entry.strip().lower()
        if e.startswith('@'):
            if a.split('@', 1)[-1] == e[1:]:
                return True
        elif a == e:
            return True
    return False


def passes_subject_guards(subject: str) -> tuple:
    """
    Returns (True, 'ok') or (False, reason_string).

    Rules applied in order:
      1. RE: / FW: / FWD: prefix → skip  (first-email-only policy)
      2. Already carries a [Processed*] tag → skip
      3. Separator not present → skip
      4. Keyword not found left of the separator → skip
    """
    if not subject:
        return False, "empty subject"

    s     = subject.strip()
    lower = s.lower()

    # Rule 1 — first-email-only: reject any reply or forward
    for prefix in ('re:', 'fw:', 'fwd:'):
        if lower.startswith(prefix):
            return False, f"reply/forward ({prefix.rstrip(':')})"

    # Rule 2 — skip already-processed emails
    if '[processed' in lower:
        return False, "already tagged"

    # Rule 3 — separator must be present
    if SUBJECT_SEPARATOR not in s:
        return False, f"no '{SUBJECT_SEPARATOR}' separator"

    # Rule 4 — keyword must appear left of the separator
    left = s.split(SUBJECT_SEPARATOR)[0].strip()
    if SUBJECT_KEYWORD.lower() not in left.lower():
        return False, f"keyword '{SUBJECT_KEYWORD}' not in '{left}'"

    return True, "ok"


def extract_hostnames(subject: str) -> list:
    """Return list of hostnames parsed from the right-hand side of the separator."""
    parts = subject.split(SUBJECT_SEPARATOR, 1)
    if len(parts) < 2:
        return []
    return [h.strip() for h in parts[1].split(SUBJECT_SEPARATOR) if h.strip()]


def body_contains_dl(mail_item) -> bool:
    """Return True if TRIGGER_DL appears in the email body (or check is disabled)."""
    if not TRIGGER_DL.strip():
        return True
    dl = TRIGGER_DL.strip().lower()
    try:
        return (dl in (mail_item.Body or '').lower() or
                dl in (mail_item.HTMLBody or '').lower())
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════════════
# CONVERSATION STATUS  (dedup — avoid re-drafting for the same conversation)
# ══════════════════════════════════════════════════════════════════════════════

def _tag_from_subject(subj: str) -> str | None:
    s = (subj or '').lower()
    if '[processed-notfound]' in s: return TAG_NOT_FOUND
    if '[processed-partial]'  in s: return TAG_PARTIAL
    if '[processed-active]'   in s: return TAG_ACTIVE
    return None


def check_conversation_status(mail_item, sent_folder, drafts_folder) -> tuple:
    """
    Scan Sent Items and Drafts for a previous reply in this conversation thread.
    Returns (tag, datetime) of the most recent outcome, or (None, None) if none found.
    """
    conv_id  = mail_item.ConversationID
    last_tag = None
    last_dt  = None

    def _update(tag, dt):
        nonlocal last_tag, last_dt
        if tag and (last_dt is None or (dt and dt > last_dt)):
            last_tag, last_dt = tag, dt

    cutoff = (datetime.now() - timedelta(days=SENT_SCAN_DAYS)).strftime('%m/%d/%Y %I:%M %p')

    # Sent Items
    try:
        for item in sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff}'"):
            try:
                if item.ConversationID == conv_id:
                    _update(_tag_from_subject(item.Subject), _com_dt_to_py(item.SentOn))
            except Exception:
                continue
    except Exception as exc:
        _log(f"      WARNING: Sent scan error: {exc}")

    # Drafts (same date window to keep large mailboxes fast)
    try:
        for item in drafts_folder.Items.Restrict(f"[LastModificationTime] >= '{cutoff}'"):
            try:
                if item.ConversationID == conv_id:
                    _update(_tag_from_subject(item.Subject),
                            _com_dt_to_py(item.LastModificationTime))
            except Exception:
                continue
    except Exception as exc:
        _log(f"      WARNING: Drafts scan error: {exc}")

    return last_tag, last_dt


# ══════════════════════════════════════════════════════════════════════════════
# HTML EMAIL BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def _host_section(hostname: str, qr: dict) -> tuple:
    """
    Build the per-host HTML block.
    Returns (html_str, host_status, type_records_list, os_group_name).
    host_status ∈ {'active', 'partial', 'not_found'}
    """
    sources = qr.get('sources', [])

    # ── Not found ────────────────────────────────────────────────────────────
    if qr.get('status') != 'Found' or not sources:
        html = f"""
<div style="margin-bottom:18px;border:1px solid #f5c6c6;border-radius:6px;overflow:hidden;">
  <div style="background:#c0392b;color:#fff;padding:8px 14px;font-size:13px;font-weight:700;">
    &#x2716;&nbsp;{hostname} &mdash; Not Found in QRadar
  </div>
  <div style="padding:10px 14px;font-size:12px;color:#555;">
    <strong>{hostname}</strong> was not found in QRadar.
    Please ensure the asset is onboarded before re-submitting the signoff.
  </div>
</div>"""
        return html, 'not_found', [], None

    group_name, group_rules = detect_os_group(sources)
    type_records = []

    # ── OS-group mode ────────────────────────────────────────────────────────
    if OS_TYPE_GROUPS and group_name:
        validation  = validate_required_types(qr, group_rules.get('required', []))
        any_missing = any(not r['found'] for r in validation)
        any_silent  = any(r['found'] and r['days_ago'] is None for r in validation)
        host_status = 'partial' if (any_missing or any_silent) else 'active'

        if host_status == 'active':
            banner_bg  = '#1a7a4a'
            banner_txt = f'&#x2714;&nbsp;{hostname} ({group_name}) &mdash; Confirmed Reporting on SIEM'
        elif any_missing:
            n = sum(1 for r in validation if r['found'])
            banner_bg  = '#c87800'
            banner_txt = (f'&#x26A0;&nbsp;{hostname} ({group_name}) &mdash; '
                          f'{n}/{len(validation)} required log sources found')
        else:
            banner_bg  = '#c87800'
            banner_txt = (f'&#x26A0;&nbsp;{hostname} ({group_name}) &mdash; '
                          f'Log sources present but no events recorded yet')

        rows = ''
        for r in validation:
            if not r['found']:
                icon, bg, ic = '&#x2716;', '#fff5f5', '#c0392b'
                cell = '<span style="color:#c0392b;font-weight:600;">Missing &mdash; requires onboarding</span>'
            elif r['days_ago'] is None:
                icon, bg, ic = '&#x26A0;', '#fffbf0', '#c87800'
                cell = '<span style="color:#c87800;font-weight:600;">No events recorded yet</span>'
            else:
                d    = 'Today' if r['days_ago'] == 0 else f"{r['days_ago']}d ago"
                icon, bg, ic = '&#x2714;', '#f0faf4', '#1a7a4a'
                cell = (f'<span style="color:#1a7a4a;font-weight:600;">Active</span>'
                        f'&nbsp;<span style="color:#888;font-size:11px;">({d})</span>')
            rows += f"""
<tr style="background:{bg};">
  <td style="padding:6px 10px;color:{ic};font-weight:700;text-align:center;width:22px;">{icon}</td>
  <td style="padding:6px 10px;font-size:12px;font-weight:600;color:#333;">{r['expected']}</td>
  <td style="padding:6px 10px;font-size:12px;color:#555;">{r.get('ls_name') or '&mdash;'}</td>
  <td style="padding:6px 10px;font-size:12px;color:#555;">{r.get('last_seen') or '&mdash;'}</td>
  <td style="padding:6px 10px;font-size:12px;">{cell}</td>
</tr>"""
            type_records.append({
                'expected': r['expected'],
                'found':    r['found'],
                'days_ago': r['days_ago'],
            })

        detail = f"""
<table style="width:100%;border-collapse:collapse;">
  <tr style="background:#f5f5f5;">
    <th style="padding:5px 10px;font-size:11px;color:#888;text-align:left;
               border-bottom:1px solid #ddd;width:22px;"></th>
    <th style="padding:5px 10px;font-size:11px;color:#888;text-align:left;
               border-bottom:1px solid #ddd;">Log Source Type</th>
    <th style="padding:5px 10px;font-size:11px;color:#888;text-align:left;
               border-bottom:1px solid #ddd;">Log Source Name</th>
    <th style="padding:5px 10px;font-size:11px;color:#888;text-align:left;
               border-bottom:1px solid #ddd;">Last Event</th>
    <th style="padding:5px 10px;font-size:11px;color:#888;text-align:left;
               border-bottom:1px solid #ddd;">Status</th>
  </tr>{rows}
</table>"""

    # ── Simple mode (no OS group detected) ───────────────────────────────────
    else:
        if OS_TYPE_GROUPS:
            _log(f"      WARNING: OS group undetected for {hostname} — using simple mode.")
        enabled  = sorted([s for s in sources if s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        disabled = sorted([s for s in sources if not s.get('enabled')],
                          key=lambda x: x.get('days_ago') or 99999)
        best        = (enabled or disabled or [None])[0]
        host_status = 'active'
        banner_bg   = '#1a7a4a'
        banner_txt  = f'&#x2714;&nbsp;{hostname} &mdash; Confirmed Reporting on SIEM'
        group_name  = None

        if best:
            dv     = best.get('days_ago')
            ds     = 'Today' if dv == 0 else (f"{dv} days ago" if dv is not None else 'N/A')
            detail = f"""
<table style="width:100%;border-collapse:collapse;font-size:12px;">
  <tr><td style="padding:6px 10px;color:#555;width:160px;
                 border-bottom:1px solid #eee;">Log Source Name</td>
      <td style="padding:6px 10px;font-weight:600;color:#222;
                 border-bottom:1px solid #eee;">{best.get('name', 'N/A')}</td></tr>
  <tr><td style="padding:6px 10px;color:#555;
                 border-bottom:1px solid #eee;">Log Source Type</td>
      <td style="padding:6px 10px;color:#333;
                 border-bottom:1px solid #eee;">{best.get('ls_type', 'N/A')}</td></tr>
  <tr><td style="padding:6px 10px;color:#555;">Last Event</td>
      <td style="padding:6px 10px;color:#333;">
        {best.get('last_seen', 'N/A')}
        &nbsp;<span style="color:#888;font-size:11px;">({ds})</span>
      </td></tr>
</table>"""
        else:
            detail = ''

    section = f"""
<div style="margin-bottom:18px;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">
  <div style="background:{banner_bg};color:#fff;padding:8px 14px;
              font-size:13px;font-weight:700;">{banner_txt}</div>
  <div>{detail}</div>
</div>"""
    return section, host_status, type_records, group_name


def build_reply_html(hostname_list: list) -> tuple:
    """
    Query QRadar for every host and build the full HTML reply body.
    Returns (html_str, overall_status, host_records_list).
    overall_status ∈ {'active', 'partial', 'not_found'}
    """
    run_time      = datetime.now().strftime('%d %B %Y, %H:%M')
    sections      = []
    host_statuses = []
    host_records  = []
    overall       = 'active'

    badge_bg   = {'active': '#1a7a4a', 'partial': '#c87800', 'not_found': '#c0392b'}
    badge_icon = {'active': '&#x2714;', 'partial': '&#x26A0;', 'not_found': '&#x2716;'}

    for hostname in hostname_list:
        _log(f"      Querying [{hostname}]...")
        qr = query_log_sources(hostname)
        _log(f"      [{hostname}] {qr['status']} | {len(qr.get('sources', []))} sources")
        section, hs, tr, og = _host_section(hostname, qr)
        sections.append(section)
        host_statuses.append(hs)
        if STATUS_PRIORITY.get(hs, 0) > STATUS_PRIORITY.get(overall, 0):
            overall = hs
        host_records.append({
            'hostname':    hostname,
            'status':      hs,
            'os_group':    og,
            'type_results': tr,
        })
        _log(f"      [{hostname}] → {hs.upper()}")

    badges = ''.join(
        f'<span style="display:inline-block;background:{badge_bg.get(hs,"#555")};'
        f'color:#fff;padding:3px 12px;border-radius:12px;'
        f'font-size:11px;font-weight:600;margin:0 4px 6px 0;">'
        f'{badge_icon.get(hs,"?")} {hn}</span>'
        for hn, hs in zip(hostname_list, host_statuses)
    )
    count = f"{len(hostname_list)} host{'s' if len(hostname_list) != 1 else ''} checked"

    html = f"""<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;color:#222;
             font-size:13px;line-height:1.6;margin:0;padding:0;">
  <div style="max-width:700px;padding:20px 0;">
    <p style="margin:0 0 14px 0;">Hi,</p>
    <p style="margin:0 0 10px 0;color:#555;font-size:12px;">
      Results for your SIEM Security Signoff request &mdash; {count}.
    </p>
    <div style="margin-bottom:18px;">{badges}</div>
    {''.join(sections)}
    <p style="margin:20px 0 4px 0;color:#888;font-size:11px;">
      Automated response from the SIEM monitoring system.<br>
      Checked against QRadar on {run_time}.
    </p>
    <p style="margin:14px 0 0 0;">Regards,<br><strong>Cyberdefence</strong></p>
  </div>
</body></html>"""

    return html, overall, host_records


# ══════════════════════════════════════════════════════════════════════════════
# DATA STORE
# ══════════════════════════════════════════════════════════════════════════════

def write_record(email_subject: str, sender: str, host_records: list,
                 overall_status: str, is_revalidation: bool,
                 prior_status: str | None) -> None:
    data = _load_data()
    data.setdefault('entries', [])
    data['entries'].append({
        'id':                f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}",
        'timestamp':         datetime.now().isoformat(),
        'email_subject':     email_subject,
        'sender':            sender,
        'overall_status':    overall_status,
        'is_revalidation':   is_revalidation,
        'prior_status':      prior_status,
        'hosts':             host_records,
        'manually_resolved': False,
        'notes':             '',
    })
    try:
        _atomic_write_json(SIGNOFF_DATA_PATH, data)
        _log(f"      Record saved ({overall_status})")
    except Exception as exc:
        _log(f"WARNING: Data write failed: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# DRAFT CREATOR
# ══════════════════════════════════════════════════════════════════════════════

def create_draft(mail_item, html_body: str, overall_status: str,
                 is_revalidation: bool = False) -> bool:
    """
    Create and Save a draft reply.  Never calls reply.Send().
    Escalation recipients are applied for Partial and Not-Found outcomes.
    """
    tag_map = {'active': TAG_ACTIVE, 'partial': TAG_PARTIAL, 'not_found': TAG_NOT_FOUND}
    tag     = tag_map.get(overall_status, TAG_ACTIVE)
    prefix  = '[Revalidated] ' if is_revalidation else ''
    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"{prefix}{tag} {mail_item.Subject}"
        if overall_status in ('partial', 'not_found'):
            if ESCALATION_TO:
                reply.To = '; '.join(ESCALATION_TO)
            if ESCALATION_CC:
                reply.CC = '; '.join(ESCALATION_CC)
            _log(f"      Escalation → To: {reply.To}  |  CC: {reply.CC or '(none)'}")
        else:
            _log("      ReplyAll (Active)")
        reply.Save()  # DRAFT ONLY — never reply.Send()
        _log(f"      Draft saved [{tag}]{' — REVAL' if is_revalidation else ''}")
        return True
    except Exception as exc:
        _log(f"      ERROR: Draft creation failed: {exc}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# OUTLOOK
# ══════════════════════════════════════════════════════════════════════════════

def get_outlook_folders():
    """Connect to Outlook and return (inbox, drafts, sent) folder objects."""
    try:
        outlook    = win32com.client.Dispatch('Outlook.Application')
        ns         = outlook.GetNamespace('MAPI')
        main_inbox = ns.GetDefaultFolder(6)    # olFolderInbox
        drafts     = ns.GetDefaultFolder(16)   # olFolderDrafts
        sent       = ns.GetDefaultFolder(5)    # olFolderSentMail
        if SIGNOFF_FOLDER_NAME:
            try:
                inbox = main_inbox.Folders[SIGNOFF_FOLDER_NAME]
                _log(f"Folder: Inbox\\{SIGNOFF_FOLDER_NAME}")
            except Exception:
                _log(f"WARNING: '{SIGNOFF_FOLDER_NAME}' sub-folder not found — scanning full Inbox.")
                inbox = main_inbox
        else:
            inbox = main_inbox
        return inbox, drafts, sent
    except Exception as exc:
        _log(f"ERROR: Outlook connection failed: {exc}")
        return None, None, None


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _ensure_data_file()

    _log('=' * 65)
    _log('QRadar Signoff Runner  v3.0')
    _log(f'  Inbox scan : last {LOOKBACK_DAYS}d')
    _log(f'  Sent scan  : last {SENT_SCAN_DAYS}d')
    _log(f'  Active-skip: {ACTIVE_SKIP_DAYS}d')
    _log(f'  Policy     : first-email-only | draft-only')
    _log(f'  Data file  : {SIGNOFF_DATA_PATH}')
    _log(f'  Log file   : {RUN_LOG_PATH}')
    _log('=' * 65)

    if not acquire_lock():
        return

    try:
        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        if not test_qradar_connection():
            _log("ERROR: QRadar unreachable — aborting.  No emails processed.")
            return

        fetch_log_source_types()

        cutoff_str         = (datetime.now() - timedelta(days=LOOKBACK_DAYS)).strftime('%m/%d/%Y %I:%M %p')
        active_skip_cutoff = datetime.now() - timedelta(days=ACTIVE_SKIP_DAYS)

        items = list(inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'"))
        _log(f"\n{len(items)} email(s) found in last {LOOKBACK_DAYS}d\n{'─'*40}")

        processed = skipped = drafted = revalidated = 0

        for mail in items:
            # Only process mail items (Class 43)
            try:
                if mail.Class != 43:
                    continue
            except Exception:
                continue

            try:
                subject = mail.Subject or ''
                sender  = mail.SenderEmailAddress or ''
            except Exception:
                continue

            # ── Subject guards ────────────────────────────────────────────────
            ok, reason = passes_subject_guards(subject)
            if not ok:
                skipped += 1
                _log(f"  SKIP ({reason}): {subject[:70]!r}")
                continue

            # ── Sender guards ─────────────────────────────────────────────────
            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                continue
            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"  SKIP (sender not in allowlist): {sender}")
                continue
            if not body_contains_dl(mail):
                skipped += 1
                _log(f"  SKIP (trigger DL not in body): {subject[:60]!r}")
                continue

            # ── Parse hostnames ───────────────────────────────────────────────
            hostname_list = extract_hostnames(subject)
            if not hostname_list:
                skipped += 1
                _log(f"  SKIP (no hostnames parsed): {subject[:60]!r}")
                continue

            # ── Runtime dedup ─────────────────────────────────────────────────
            if is_drafted_this_run(hostname_list):
                skipped += 1
                _log(f"  SKIP (runtime dedup — {hostname_list} already processed this run)")
                continue

            _log(f"\n  Candidate : {subject[:70]!r}")
            _log(f"  Sender    : {sender}")
            _log(f"  Hosts     : {hostname_list}")

            # ── Conversation dedup ────────────────────────────────────────────
            last_tag, last_dt = check_conversation_status(mail, sent, drafts)

            if last_tag == TAG_ACTIVE:
                if last_dt and last_dt >= active_skip_cutoff:
                    skipped += 1
                    _log(f"  SKIP (Active on {last_dt.strftime('%Y-%m-%d')} — within {ACTIVE_SKIP_DAYS}d window)")
                    continue
                _log(f"  Active result is >{ACTIVE_SKIP_DAYS}d old — revalidating")
                is_reval = True
            elif last_tag in REVALIDATABLE_TAGS:
                _log(f"  Revalidating prior {last_tag} result")
                is_reval = True
            elif last_tag is None:
                _log("  New signoff — no prior result found")
                is_reval = False
            else:
                skipped += 1
                _log(f"  SKIP (unrecognised tag: {last_tag})")
                continue

            # ── QRadar query + draft ──────────────────────────────────────────
            html_body, overall_status, host_records = build_reply_html(hostname_list)
            _log(f"  Overall   : {overall_status.upper()}")

            success = create_draft(mail, html_body, overall_status, is_revalidation=is_reval)
            if success:
                drafted     += 1
                revalidated += int(is_reval)
                mark_drafted_this_run(hostname_list)
                write_record(
                    email_subject   = subject,
                    sender          = sender,
                    host_records    = host_records,
                    overall_status  = overall_status,
                    is_revalidation = is_reval,
                    prior_status    = last_tag,
                )
            processed += 1

        _log(f"\n{'='*65}")
        _log(f"Run complete — {processed} processed | {drafted} drafted "
             f"({revalidated} revalidation{'s' if revalidated != 1 else ''}) | "
             f"{skipped} skipped")
        _log(f"Data : {SIGNOFF_DATA_PATH}")
        _log("Tip  : run signoff_dashboard.py to view results in your browser.")

    finally:
        release_lock()


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        _log("\nInterrupted by user.")
        release_lock()
