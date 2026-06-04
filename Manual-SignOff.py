"""
QRadar Manual Signoff  v1.0
────────────────────────────────────────────────────────────────────────────
Terminal-driven signoff checker — no Outlook scanning required.
Drop this file alongside signoff_runner.py; it shares the same
signoff_data.json so signoff_dashboard.py shows manual runs too.

Usage
  python manual_signoff.py

  When prompted, enter one or more hostnames separated by  |
      server01 | server02 | server03

Outputs
  • HTML report  →  manual_signoff_YYYYMMDD-HHMMSS.html  (auto-opened)
  • JSON record  →  signoff_data.json  (shared with dashboard)
  • Appended log →  signoff_runner.log
────────────────────────────────────────────────────────────────────────────
"""

import json
import os
import sys
import tempfile
import uuid
import urllib3
import webbrowser
import requests

from datetime import datetime


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION  — keep in sync with signoff_runner.py
# ══════════════════════════════════════════════════════════════════════════════

# ─── Paths ────────────────────────────────────────────────────────────────────
_DIR              = os.path.dirname(os.path.abspath(__file__))
RUN_LOG_PATH      = os.path.join(_DIR, 'signoff_runner.log')       # shared log
LOCKFILE_PATH     = os.path.join(_DIR, 'manual_signoff.lock')      # own lockfile
SIGNOFF_DATA_PATH = os.path.join(_DIR, 'signoff_data.json')        # shared data

# ─── QRadar ───────────────────────────────────────────────────────────────────
QRADAR_HOST     = os.environ.get('QRADAR_HOST',     'https://your-qradar-host')
QRADAR_USERNAME = os.environ.get('QRADAR_USERNAME', 'your-username')
QRADAR_PASSWORD = os.environ.get('QRADAR_PASSWORD', 'your-password')
VERIFY_SSL      = False          # set True + supply a CA bundle in production

# ─── OS type validation ───────────────────────────────────────────────────────
OS_TYPE_GROUPS = {
    'Windows': {'required': ['Microsoft Security', 'WinCollect']},
    'Linux':   {'required': ['Linux OS']},
}

# ─── Outcome tags (must match signoff_runner.py) ──────────────────────────────
TAG_ACTIVE    = '[Processed-Active]'
TAG_PARTIAL   = '[Processed-Partial]'
TAG_NOT_FOUND = '[Processed-NotFound]'

# ─── QRadar API ───────────────────────────────────────────────────────────────
ACTIVITY_THRESHOLD_DAYS = 7
REQUEST_TIMEOUT         = 30
_MIN_TS                 = 0
_MAX_TS                 = 2_147_483_647


# ══════════════════════════════════════════════════════════════════════════════
# MODULE-LEVEL STATE
# ══════════════════════════════════════════════════════════════════════════════

LOG_SOURCE_TYPES_CACHE: dict = {}
STATUS_PRIORITY = {'not_found': 2, 'partial': 1, 'active': 0}

# ── ANSI colours ──────────────────────────────────────────────────────────────
_C = {
    'green':  '\033[92m',
    'yellow': '\033[93m',
    'red':    '\033[91m',
    'cyan':   '\033[96m',
    'bold':   '\033[1m',
    'dim':    '\033[2m',
    'reset':  '\033[0m',
}
# Enable VT sequences on Windows 10+ cmd; silently disable colour if that fails
if sys.platform == 'win32':
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleMode(
            ctypes.windll.kernel32.GetStdHandle(-11), 7)
    except Exception:
        _C = {k: '' for k in _C}


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
    """Crash-safe JSON write: dump to .tmp, then os.replace() into place."""
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
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as exc:
        _log(f"WARNING: Data file unreadable ({exc}) — starting fresh.")
        return {'schema_version': 3, 'entries': []}


def _ensure_data_file() -> None:
    if not os.path.exists(SIGNOFF_DATA_PATH):
        _atomic_write_json(SIGNOFF_DATA_PATH, {'schema_version': 3, 'entries': []})
        _log(f"Created: {SIGNOFF_DATA_PATH}")


def acquire_lock() -> bool:
    if os.path.exists(LOCKFILE_PATH):
        _log("WARNING: Lockfile present — another instance may be running. Exiting.")
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
        r = _qradar_get(
            '/api/config/event_sources/log_source_management/log_source_types')
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


def detect_os_group(sources: list) -> tuple:
    """Returns (group_name, group_rules) or (None, None) if undetected."""
    for gname, rules in OS_TYPE_GROUPS.items():
        req = rules.get('required', [])
        if req:
            sig = req[0].lower().split()
            if any(all(w in s.get('ls_type', '').lower() for w in sig)
                   for s in sources):
                return gname, rules
    return None, None


# ══════════════════════════════════════════════════════════════════════════════
# HTML REPORT BUILDER  (identical logic to signoff_runner.py)
# ══════════════════════════════════════════════════════════════════════════════

def _host_section(hostname: str, qr: dict) -> tuple:
    """
    Build the per-host HTML block.
    Returns (html_str, host_status, type_records_list, os_group_name).
    host_status ∈ {'active', 'partial', 'not_found'}
    """
    sources = qr.get('sources', [])

    # ── Not found ─────────────────────────────────────────────────────────────
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

    # ── OS-group mode ──────────────────────────────────────────────────────────
    if OS_TYPE_GROUPS and group_name:
        validation  = validate_required_types(qr, group_rules.get('required', []))
        any_missing = any(not r['found'] for r in validation)
        any_silent  = any(r['found'] and r['days_ago'] is None for r in validation)
        host_status = 'partial' if (any_missing or any_silent) else 'active'

        if host_status == 'active':
            banner_bg  = '#1a7a4a'
            banner_txt = (f'&#x2714;&nbsp;{hostname} ({group_name}) '
                          f'&mdash; Confirmed Reporting on SIEM')
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
                cell = ('<span style="color:#c0392b;font-weight:600;">'
                        'Missing &mdash; requires onboarding</span>')
            elif r['days_ago'] is None:
                icon, bg, ic = '&#x26A0;', '#fffbf0', '#c87800'
                cell = ('<span style="color:#c87800;font-weight:600;">'
                        'No events recorded yet</span>')
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

    # ── Simple mode (no OS group detected) ────────────────────────────────────
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


def build_reply_html(hostname_list: list, analyst: str = '') -> tuple:
    """
    Query QRadar for every host and build the full HTML report.
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
            'hostname':     hostname,
            'status':       hs,
            'os_group':     og,
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
    by    = f' &nbsp;&middot;&nbsp; Submitted by: <strong>{analyst}</strong>' if analyst else ''

    # Overall status banner for the report header
    overall_colors = {'active': '#1a7a4a', 'partial': '#c87800', 'not_found': '#c0392b'}
    overall_labels = {
        'active':    '&#x2714; ALL ACTIVE',
        'partial':   '&#x26A0; PARTIAL',
        'not_found': '&#x2716; NOT FOUND',
    }
    ov_bg  = overall_colors.get(overall, '#555')
    ov_lbl = overall_labels.get(overall, overall.upper())

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SIEM Manual Signoff — {run_time}</title>
  <style>
    body {{
      font-family: 'Segoe UI', Arial, sans-serif;
      color: #222;
      font-size: 13px;
      line-height: 1.6;
      background: #f4f6f9;
      margin: 0;
      padding: 24px;
    }}
    .card {{
      background: #fff;
      border-radius: 8px;
      box-shadow: 0 1px 4px rgba(0,0,0,.12);
      max-width: 760px;
      margin: 0 auto;
      padding: 28px 32px;
    }}
    .report-header {{
      border-bottom: 2px solid #eee;
      margin-bottom: 20px;
      padding-bottom: 14px;
    }}
    .report-title {{
      font-size: 18px;
      font-weight: 700;
      color: #1a1a2e;
      margin: 0 0 4px 0;
    }}
    .report-meta {{
      font-size: 12px;
      color: #888;
    }}
    .overall-badge {{
      display: inline-block;
      background: {ov_bg};
      color: #fff;
      padding: 5px 18px;
      border-radius: 20px;
      font-size: 13px;
      font-weight: 700;
      margin-bottom: 16px;
    }}
    .footer {{
      margin-top: 24px;
      padding-top: 12px;
      border-top: 1px solid #eee;
      font-size: 11px;
      color: #aaa;
    }}
  </style>
</head>
<body>
  <div class="card">
    <div class="report-header">
      <p class="report-title">&#x1F6E1; SIEM Manual Signoff Report</p>
      <p class="report-meta">
        {run_time}{by}
        &nbsp;&middot;&nbsp; {count}
      </p>
    </div>

    <div class="overall-badge">{ov_lbl}</div>

    <div style="margin-bottom:18px;">{badges}</div>

    {''.join(sections)}

    <div class="footer">
      Manual signoff — QRadar checked on {run_time}.<br>
      View all results in <strong>signoff_dashboard.py</strong>.
    </div>
  </div>
</body>
</html>"""

    return html, overall, host_records


# ══════════════════════════════════════════════════════════════════════════════
# DATA STORE  (schema matches signoff_runner.py — dashboard reads both)
# ══════════════════════════════════════════════════════════════════════════════

def write_record(label: str, analyst: str, notes: str,
                 host_records: list, overall_status: str) -> None:
    """Append one entry to signoff_data.json using the same schema as v3.0."""
    data = _load_data()
    data.setdefault('entries', [])
    data['entries'].append({
        'id':                f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}",
        'timestamp':         datetime.now().isoformat(),
        'email_subject':     label,        # "Manual Signoff | host1 | host2"
        'sender':            analyst,      # analyst name / 'Manual Entry'
        'overall_status':    overall_status,
        'is_revalidation':   False,
        'prior_status':      None,
        'hosts':             host_records,
        'manually_resolved': False,
        'notes':             notes,
    })
    try:
        _atomic_write_json(SIGNOFF_DATA_PATH, data)
        _log(f"      Record saved ({overall_status}) → {SIGNOFF_DATA_PATH}")
    except Exception as exc:
        _log(f"WARNING: Data write failed: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# HTML OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

def save_html_report(html: str) -> str | None:
    """Save HTML to a timestamped file next to this script and open in browser."""
    ts       = datetime.now().strftime('%Y%m%d-%H%M%S')
    filename = f"manual_signoff_{ts}.html"
    out_path = os.path.join(_DIR, filename)
    try:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html)
        _log(f"HTML report saved: {out_path}")
        try:
            # file:/// URI needs forward slashes on all platforms
            uri = 'file:///' + out_path.replace('\\', '/')
            webbrowser.open(uri)
            _log("Report opened in browser.")
        except Exception as exc:
            _log(f"NOTE: Could not auto-open browser ({exc}). Open the file manually.")
        return out_path
    except Exception as exc:
        _log(f"ERROR: Could not save HTML report: {exc}")
        return None


# ══════════════════════════════════════════════════════════════════════════════
# TERMINAL UI
# ══════════════════════════════════════════════════════════════════════════════

_STATUS_COLOUR = {
    'active':    _C['green'],
    'partial':   _C['yellow'],
    'not_found': _C['red'],
}
_STATUS_LABEL = {
    'active':    'ACTIVE',
    'partial':   'PARTIAL',
    'not_found': 'NOT FOUND',
}


def _cprint(text: str, colour: str = '') -> None:
    """Print with optional ANSI colour."""
    print(f"{colour}{text}{_C['reset']}" if colour else text)


def print_banner() -> None:
    _cprint(f"\n{'═'*65}", _C['cyan'])
    _cprint(f"  QRadar Manual Signoff  v1.0", _C['bold'])
    _cprint(f"{'═'*65}", _C['cyan'])


def prompt_hostnames() -> list:
    """Prompt until at least one valid hostname is entered."""
    print(f"\n{_C['bold']}  Enter hostname(s) separated by  |{_C['reset']}")
    print(f"  {_C['dim']}Example: server01 | server02 | server03{_C['reset']}")
    print(f"  {_C['dim']}Press Ctrl-C to exit.{_C['reset']}\n")

    while True:
        try:
            raw = input("  Hostnames ▶  ").strip()
        except KeyboardInterrupt:
            print()
            return []

        hosts = [h.strip() for h in raw.split('|') if h.strip()]
        if hosts:
            return hosts
        print(f"  {_C['yellow']}⚠  No hostnames parsed — try again.{_C['reset']}\n")


def prompt_analyst() -> tuple:
    """Prompt for optional analyst name and notes."""
    try:
        analyst = input(
            f"\n  {_C['dim']}Your name (optional — press Enter to skip): {_C['reset']}"
        ).strip() or 'Manual Entry'
        notes = input(
            f"  {_C['dim']}Notes (optional — press Enter to skip):    {_C['reset']}"
        ).strip()
    except KeyboardInterrupt:
        print()
        analyst, notes = 'Manual Entry', ''
    return analyst, notes


def print_summary(hostname_list: list, host_records: list,
                  overall_status: str, out_path: str | None) -> None:
    """Print a clean result table to the terminal."""
    ov_colour = _STATUS_COLOUR.get(overall_status, '')
    ov_label  = _STATUS_LABEL.get(overall_status, overall_status.upper())

    print(f"\n{'─'*65}")
    _cprint(f"  RESULT SUMMARY", _C['bold'])
    print(f"{'─'*65}")

    col_w = max((len(h) for h in hostname_list), default=20) + 4
    for rec in host_records:
        hn     = rec['hostname']
        st     = rec['status']
        colour = _STATUS_COLOUR.get(st, '')
        label  = _STATUS_LABEL.get(st, st.upper())

        # Show days-ago for the most recent active type (if available)
        days_info = ''
        for tr in rec.get('type_results', []):
            if tr.get('found') and tr.get('days_ago') is not None:
                d = tr['days_ago']
                days_info = f"  (last event: {'Today' if d == 0 else f'{d}d ago'})"
                break

        _cprint(f"  {hn:<{col_w}} {label}{days_info}", colour)

    print(f"{'─'*65}")
    _cprint(f"  Overall : {ov_label}", ov_colour + _C['bold'])
    if out_path:
        _cprint(f"  Report  : {out_path}", _C['dim'])
    _cprint(f"  Data    : {SIGNOFF_DATA_PATH}", _C['dim'])
    print(f"{'─'*65}\n")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def run_once() -> None:
    """Perform a single manual signoff query."""
    hostname_list = prompt_hostnames()
    if not hostname_list:
        return

    analyst, notes = prompt_analyst()

    _log(f"\n  Analyst   : {analyst}")
    _log(f"  Hosts     : {hostname_list}")

    # ── Query QRadar + build HTML ──────────────────────────────────────────────
    html, overall_status, host_records = build_reply_html(hostname_list, analyst)
    _log(f"  Overall   : {overall_status.upper()}")

    # ── Save HTML report (opens browser) ──────────────────────────────────────
    out_path = save_html_report(html)

    # ── Persist to signoff_data.json for dashboard ────────────────────────────
    label = f"Manual Signoff | {'|'.join(hostname_list)}"
    write_record(
        label          = label,
        analyst        = analyst,
        notes          = notes,
        host_records   = host_records,
        overall_status = overall_status,
    )

    # ── Terminal summary ───────────────────────────────────────────────────────
    print_summary(hostname_list, host_records, overall_status, out_path)


def main() -> None:
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _ensure_data_file()
    print_banner()

    _log(f"  Data file  : {SIGNOFF_DATA_PATH}")
    _log(f"  Log file   : {RUN_LOG_PATH}")

    if not acquire_lock():
        return

    try:
        if not test_qradar_connection():
            _log("ERROR: QRadar unreachable — aborting.")
            return

        fetch_log_source_types()

        # ── Multi-run loop ─────────────────────────────────────────────────────
        while True:
            run_once()

            try:
                again = input(
                    f"  Check another? {_C['dim']}[y/N]{_C['reset']}  "
                ).strip().lower()
            except KeyboardInterrupt:
                print()
                break
            if again not in ('y', 'yes'):
                break
            print()

    finally:
        release_lock()

    _cprint("\n  Done. Run signoff_dashboard.py to view all results.\n", _C['dim'])


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        _log("\nInterrupted by user.")
        release_lock()
