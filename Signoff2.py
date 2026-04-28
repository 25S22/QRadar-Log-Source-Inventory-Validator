"""
QRadar Signoff Auto-Draft  v2.2
────────────────────────────────────────────────────────────────────────────
Scans Outlook for SIEM signoff emails, queries QRadar, saves draft replies.

v2.2 changes (over v2.1):
  • Fix 1: Multi-level prefix stripping — RE: FW: RE: chains now fully stripped
            before keyword matching.
  • Fix 2: Drafts folder date-restricted — same SENT_SCAN_DAYS window applied
            via .Restrict() so large mailboxes don't slow the scan to a crawl.
  • Fix 3: Clean Ctrl+C / crash exit — KeyboardInterrupt caught at __main__
            level; HTTP server shut down immediately instead of hanging until
            the DASHBOARD_SERVE_MINUTES timeout expires.
  • Fix 4: Atomic JSON writes — all data-file writes go through
            _atomic_write_json() which writes to a .tmp file then os.replace()
            so a mid-write crash can never corrupt signoff_data.json.
  • Fix 5: Dashboard save button — persistent "Save Changes" button in the
            header (greyed + disabled when clean, green + pulsing when dirty);
            discard fully reverts to last persisted snapshot; gi mapping is
            stable across filtered / sorted / paginated views.

DRAFT-ONLY: reply.Save() is called, NEVER reply.Send().
"""

import json
import os
import socket
import tempfile
import threading
import time
import urllib3
import uuid
import webbrowser
import win32com.client
import requests

from datetime import datetime, timedelta
from http.server import HTTPServer, BaseHTTPRequestHandler


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

# ─── HISTORICAL IMPORT ───────────────────────────────────────────────────────
# ONE-TIME USE: set True to ingest already-tagged legacy emails into the data
# file for management reporting.  Reset to False immediately after that run.
HISTORICAL_RUN_MODE = False

# ─── DASHBOARD SERVER ────────────────────────────────────────────────────────
DASHBOARD_SERVER_PORT   = 8745   # preferred local port; auto-picks if in use
DASHBOARD_SERVE_MINUTES = 30     # minutes to keep server alive; 0 = file:// only (no save)

# ─── CONSTANTS ───────────────────────────────────────────────────────────────
ACTIVITY_THRESHOLD_DAYS = 7
_MIN_TS                 = 0
_MAX_TS                 = 2147483647
LOG_SOURCE_TYPES_CACHE  = {}
STATUS_PRIORITY         = {'not_found': 2, 'partial': 1, 'active': 0}
REQUEST_TIMEOUT         = 30

_runtime_drafted_hosts: set  = set()
_http_server_instance        = None
_http_server_port: int       = 0


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


def _atomic_write_json(path, data):
    """
    Write JSON atomically: dump to a sibling .tmp file then os.replace() it
    into place.  A crash mid-write leaves the original file intact.
    """
    dir_ = os.path.dirname(os.path.abspath(path))
    tmp_fd, tmp_path = tempfile.mkstemp(dir=dir_, suffix='.tmp')
    try:
        with os.fdopen(tmp_fd, 'w', encoding='utf-8') as tf:
            json.dump(data, tf, indent=2, ensure_ascii=False)
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD HTTP SERVER
# ══════════════════════════════════════════════════════════════════════════════

def _find_free_port(preferred=8745):
    """Returns preferred port if free, otherwise picks a random free port."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(('127.0.0.1', preferred))
        s.close()
        return preferred
    except OSError:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(('127.0.0.1', 0))
        port = s.getsockname()[1]
        s.close()
        return port


class _DashboardHandler(BaseHTTPRequestHandler):
    """Minimal HTTP handler: serves dashboard HTML + data JSON, accepts saves."""

    def do_GET(self):
        if self.path in ('/', '/index.html'):
            self._serve_file(DASHBOARD_PATH, 'text/html; charset=utf-8')
        elif self.path == '/data.json':
            self._serve_file(SIGNOFF_DATA_PATH, 'application/json')
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == '/save':
            try:
                length = int(self.headers.get('Content-Length', 0))
                body   = self.rfile.read(length)
                data   = json.loads(body.decode('utf-8'))
                # FIX 4: atomic write — original file untouched if this crashes
                _atomic_write_json(SIGNOFF_DATA_PATH, data)
                n = len(data.get('entries', []))
                _log(f"[Dashboard] Saved via HTTP — {n} entries in {SIGNOFF_DATA_PATH}")
                # Regenerate the HTML so embedded data is fresh on next file:// open
                _write_dashboard_html(_http_server_port)
                self._json_ok({'ok': True, 'entries': n})
            except Exception as e:
                _log(f"[Dashboard] Save error: {e}")
                self._json_ok({'error': str(e)}, code=500)
        else:
            self.send_response(404)
            self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def _serve_file(self, path, content_type):
        try:
            with open(path, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', content_type)
            self.send_header('Cache-Control', 'no-cache')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(content)
        except FileNotFoundError:
            self.send_response(404)
            self.end_headers()
        except Exception as e:
            self.send_response(500)
            self.end_headers()
            self.wfile.write(str(e).encode())

    def _json_ok(self, obj, code=200):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        pass  # suppress per-request console noise


def _start_dashboard_server():
    """Start the local HTTP server (once only). Returns port number or 0 if disabled."""
    global _http_server_instance, _http_server_port

    if DASHBOARD_SERVE_MINUTES <= 0:
        return 0
    if _http_server_instance:
        return _http_server_port

    port   = _find_free_port(DASHBOARD_SERVER_PORT)
    server = HTTPServer(('127.0.0.1', port), _DashboardHandler)
    _http_server_instance = server
    _http_server_port     = port

    # Serve thread — daemon=False keeps process alive until server shuts down
    threading.Thread(
        target=server.serve_forever,
        name='DashboardServer',
        daemon=False,
    ).start()

    # Timer thread — shuts server down after configured minutes
    def _auto_shutdown(srv, mins):
        time.sleep(mins * 60)
        _log(f"[Dashboard] Server timeout ({mins} min) — shutting down. Process will exit.")
        srv.shutdown()

    threading.Thread(
        target=_auto_shutdown,
        args=(server, DASHBOARD_SERVE_MINUTES),
        name='DashboardTimer',
        daemon=True,
    ).start()

    _log(f"[Dashboard] Server running: http://127.0.0.1:{port} "
         f"(auto-closes in {DASHBOARD_SERVE_MINUTES} min)")
    return port


# ══════════════════════════════════════════════════════════════════════════════
# RUNTIME + CROSS-RUN DEDUP
# ══════════════════════════════════════════════════════════════════════════════

def _host_key(hostname_list):
    return frozenset(h.upper().strip() for h in hostname_list)


def is_drafted_this_run(hostname_list):
    return _host_key(hostname_list) in _runtime_drafted_hosts


def mark_drafted_this_run(hostname_list):
    _runtime_drafted_hosts.add(_host_key(hostname_list))


def get_prior_status_from_data(hostname_list):
    """
    Cross-run dedup: look up the most recent live (non-historical) entry for
    this hostname set in signoff_data.json.  Returns (status_str, datetime)
    or (None, None) if not found.
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
        if entry.get('historical'):          # skip historical imports
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
# HISTORICAL IMPORT HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _historical_already_imported(email_subject):
    """True if an entry with this exact subject was already written as historical."""
    if not os.path.exists(SIGNOFF_DATA_PATH):
        return False
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return any(
            e.get('historical') and e.get('email_subject') == email_subject
            for e in data.get('entries', [])
        )
    except Exception:
        return False


def write_historical_record(email_subject, sender, hostname_list, status, received_dt=None):
    """Write a minimal record for a previously-processed (already-tagged) email."""
    host_records = [
        {'hostname': h, 'status': status, 'os_group': None, 'type_results': []}
        for h in hostname_list
    ]
    # Use original email date so the management timeline is accurate
    timestamp = received_dt.isoformat() if received_dt else datetime.now().isoformat()

    data = {'schema_version': 2, 'entries': []}
    if os.path.exists(SIGNOFF_DATA_PATH):
        try:
            with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            _log(f"WARNING: Data read error (will overwrite): {e}")

    record = {
        'run_id':            f"hist-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}",
        'timestamp':         timestamp,
        'email_subject':     email_subject,
        'sender':            sender,
        'is_revalidation':   False,
        'prior_status':      None,
        'overall_status':    status,
        'hosts':             host_records,
        'notes':             'Historical import',
        'manually_resolved': False,
        'historical':        True,
    }
    data['entries'].append(record)
    try:
        # FIX 4: atomic write
        _atomic_write_json(SIGNOFF_DATA_PATH, data)
        _log(f"      Historical record written ({status})")
    except Exception as e:
        _log(f"WARNING: Data write error: {e}")


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
        if e.startswith('@'):
            # exact domain match — '@org.com' must not match 'user@badorg.com'
            if a.split('@', 1)[-1] == e[1:]:
                return True
        elif a == e:
            return True
    return False


def passes_subject_guards(subject):
    """
    Guards: outcome tag present → skip (unless HISTORICAL_RUN_MODE).
    No separator → skip.  Keyword absent → skip.
    RE/FW/FWD prefixes stripped in a loop until none remain (FIX 1).
    Keyword matching is case-insensitive.
    """
    if not subject:
        return False, "empty subject"
    s  = subject.strip()
    sl = s.lower()
    if '[processed' in sl and not HISTORICAL_RUN_MODE:
        return False, "already tagged"
    if SUBJECT_SEPARATOR not in s:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found"
    left = s.split(SUBJECT_SEPARATOR)[0].strip()

    # FIX 1: loop until all RE:/FW:/FWD: prefixes are stripped
    _prefixes = ('re:', 'fw:', 'fwd:')
    while any(left.lower().startswith(p) for p in _prefixes):
        for pfx in _prefixes:
            if left.lower().startswith(pfx):
                left = left[len(pfx):].strip()
                break

    if SUBJECT_KEYWORD.lower() not in left.lower():
        return False, f"keyword '{SUBJECT_KEYWORD}' not found"
    return True, "ok"


def extract_hostnames(subject):
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
      Layer 2 — data file hostname-set lookup (cross-thread fallback).
    Returns (tag, datetime) of most recent outcome, or (None, None).
    """
    conv_id  = mail_item.ConversationID
    last_tag = None
    last_dt  = None

    def _update(tag, dt):
        nonlocal last_tag, last_dt
        if tag and (last_dt is None or (dt and dt > last_dt)):
            last_tag, last_dt = tag, dt

    # Layer 1a — Sent Items
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

    # Layer 1b — Drafts (FIX 2: date-restricted, same window as Sent scan)
    drafts_cutoff = (datetime.now() - timedelta(days=SENT_SCAN_DAYS)).strftime('%m/%d/%Y %I:%M %p')
    try:
        for item in drafts_folder.Items.Restrict(f"[LastModificationTime] >= '{drafts_cutoff}'"):
            try:
                if item.ConversationID == conv_id:
                    _update(_tag_from_subject(item.Subject),
                            _com_dt_to_py(item.LastModificationTime))
            except Exception:
                continue
    except Exception as e:
        _log(f"      WARNING: Drafts scan error: {e}")

    # Layer 2 — data file fallback
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
        'historical':        False,
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
        # FIX 4: atomic write
        _atomic_write_json(SIGNOFF_DATA_PATH, data)
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
# DASHBOARD — HTML WRITER  (paste your _write_dashboard_html body here)
# ══════════════════════════════════════════════════════════════════════════════

def _write_dashboard_html(port: int):
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        _log(f"WARNING: Dashboard HTML generation skipped — could not read data: {e}")
        return

    json_blob        = json.dumps(data, ensure_ascii=False)
    save_origin      = f'http://127.0.0.1:{port}' if port else ''
    server_active_js = 'true' if port else 'false'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SIEM Signoff Dashboard</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap');
  *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0;}}
  :root{{
    --bg:#080b10;--surface:#0e1219;--surface2:#141820;--surface3:#1a2030;
    --border:#1e2535;--border2:#252d40;
    --green:#00d68f;--green-dim:#00d68f33;
    --amber:#f0a500;--amber-dim:#f0a50033;
    --red:#ff4d6d;--red-dim:#ff4d6d33;
    --blue:#4d9fff;--blue-dim:#4d9fff33;
    --purple:#b57bee;--purple-dim:#b57bee33;
    --text:#dde3f0;--muted:#5a6480;--muted2:#3d4560;
    --mono:'JetBrains Mono',monospace;--sans:'DM Sans',sans-serif;
    --radius:10px;--transition:all .18s ease;
  }}
  html{{scroll-behavior:smooth;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;
        min-height:100vh;overflow-x:hidden;}}

  /* ── Scrollbar ─────────────────────────────────────────────── */
  ::-webkit-scrollbar{{width:6px;height:6px;}}
  ::-webkit-scrollbar-track{{background:var(--bg);}}
  ::-webkit-scrollbar-thumb{{background:var(--border2);border-radius:3px;}}

  /* ── Header ─────────────────────────────────────────────────── */
  header{{
    background:linear-gradient(180deg,var(--surface) 0%,rgba(14,18,25,.95) 100%);
    border-bottom:1px solid var(--border);padding:0 32px;
    display:flex;align-items:center;gap:20px;height:62px;
    position:sticky;top:0;z-index:200;backdrop-filter:blur(12px);
  }}
  .logo{{display:flex;align-items:center;gap:10px;}}
  .logo-dot{{width:8px;height:8px;border-radius:50%;background:var(--green);
              box-shadow:0 0 8px var(--green);animation:pulse-dot 2.4s ease-in-out infinite;}}
  @keyframes pulse-dot{{0%,100%{{opacity:1;}}50%{{opacity:.4;}}}}
  .logo h1{{font-family:var(--mono);font-size:12px;font-weight:700;
             letter-spacing:2px;color:var(--text);text-transform:uppercase;}}
  .logo-sub{{font-family:var(--mono);font-size:10px;color:var(--muted);letter-spacing:1px;margin-top:2px;}}
  .spacer{{flex:1;}}
  .hdr-controls{{display:flex;gap:8px;align-items:center;}}

  /* ── Buttons ─────────────────────────────────────────────────── */
  .btn{{
    background:transparent;border:1px solid var(--border2);color:var(--muted);
    border-radius:7px;font-size:11px;padding:6px 14px;cursor:pointer;
    font-family:var(--mono);font-weight:600;letter-spacing:.3px;
    transition:var(--transition);white-space:nowrap;
  }}
  .btn:hover:not(:disabled){{border-color:var(--blue);color:var(--blue);background:var(--blue-dim);}}
  .btn:disabled{{opacity:.35;cursor:not-allowed;}}

  .btn-save{{border-color:var(--green);color:var(--green);}}
  .btn-save:hover:not(:disabled){{background:var(--green);color:#000;box-shadow:0 0 14px var(--green-dim);}}
  .btn-save.dirty{{
    background:var(--green-dim);border-color:var(--green);color:var(--green);
    animation:save-pulse 1.8s ease-in-out infinite;
  }}
  @keyframes save-pulse{{0%,100%{{box-shadow:0 0 0 0 var(--green-dim);}}
    50%{{box-shadow:0 0 0 5px transparent;}}}}

  .btn-discard{{border-color:var(--red);color:var(--red);}}
  .btn-discard:hover:not(:disabled){{background:var(--red-dim);}}
  .btn-del{{border-color:var(--red-dim);color:var(--red);}}
  .btn-del:hover{{background:var(--red);color:#fff;border-color:var(--red);}}
  .btn-exp{{border-color:var(--amber-dim);color:var(--amber);}}
  .btn-exp:hover{{background:var(--amber-dim);border-color:var(--amber);}}

  /* ── Pill period selector ───────────────────────────────────── */
  .pill-group{{display:flex;border:1px solid var(--border2);border-radius:7px;overflow:hidden;}}
  .pill-group button{{
    background:transparent;border:none;color:var(--muted);
    font-family:var(--mono);font-size:11px;font-weight:600;
    padding:6px 13px;cursor:pointer;transition:var(--transition);letter-spacing:.3px;
  }}
  .pill-group button.active{{background:var(--blue);color:#fff;}}
  .pill-group button:not(.active):hover{{background:var(--surface3);color:var(--text);}}

  /* ── Search ──────────────────────────────────────────────────── */
  .search-wrap{{position:relative;}}
  .search-wrap svg{{position:absolute;left:10px;top:50%;transform:translateY(-50%);
                    color:var(--muted);pointer-events:none;}}
  input[type=text]{{
    background:var(--surface2);border:1px solid var(--border2);border-radius:7px;
    color:var(--text);font-family:var(--mono);font-size:11px;
    padding:6px 12px 6px 32px;width:210px;outline:none;transition:var(--transition);
  }}
  input[type=text]:focus{{border-color:var(--blue);background:var(--surface3);width:240px;}}

  /* ── Dirty banner ────────────────────────────────────────────── */
  #dirtyBanner{{
    display:none;background:linear-gradient(90deg,#1a1200,#100800);
    border-bottom:1px solid var(--amber);padding:8px 32px;
    align-items:center;gap:10px;font-family:var(--mono);font-size:11px;color:var(--amber);
  }}
  #dirtyBanner.show{{display:flex;}}
  #dirtyBanner .dbmsg{{flex:1;}}

  /* ── Main ────────────────────────────────────────────────────── */
  main{{padding:28px 32px;max-width:1440px;margin:0 auto;}}

  /* ── Stat cards ──────────────────────────────────────────────── */
  .stats{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:26px;}}
  @media(max-width:900px){{.stats{{grid-template-columns:repeat(2,1fr);}}}}
  .card{{
    background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
    padding:20px 22px;position:relative;overflow:hidden;cursor:default;
    transition:var(--transition);
  }}
  .card:hover{{border-color:var(--border2);transform:translateY(-1px);}}
  .card-accent{{position:absolute;top:0;left:0;right:0;height:2px;}}
  .card.g .card-accent{{background:linear-gradient(90deg,var(--green),transparent);}}
  .card.a .card-accent{{background:linear-gradient(90deg,var(--amber),transparent);}}
  .card.r .card-accent{{background:linear-gradient(90deg,var(--red),transparent);}}
  .card.b .card-accent{{background:linear-gradient(90deg,var(--blue),transparent);}}
  .card.p .card-accent{{background:linear-gradient(90deg,var(--purple),transparent);}}
  .card .num{{
    font-family:var(--mono);font-size:38px;font-weight:700;
    line-height:1;margin-bottom:8px;letter-spacing:-1px;
  }}
  .card.g .num{{color:var(--green);}} .card.a .num{{color:var(--amber);}}
  .card.r .num{{color:var(--red);}}   .card.b .num{{color:var(--blue);}}
  .card.p .num{{color:var(--purple);}}
  .card .clabel{{color:var(--muted);font-size:10px;letter-spacing:1px;
                 text-transform:uppercase;font-weight:600;}}
  .card .sub{{color:var(--muted2);font-size:11px;margin-top:5px;}}

  /* ── Charts ──────────────────────────────────────────────────── */
  .charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:26px;}}
  @media(max-width:700px){{.charts-grid{{grid-template-columns:1fr;}}}}
  .chart-card{{
    background:var(--surface);border:1px solid var(--border);
    border-radius:var(--radius);padding:20px 22px;
  }}
  .chart-card h3{{
    font-family:var(--mono);font-size:10px;color:var(--muted);
    margin-bottom:16px;letter-spacing:1.5px;text-transform:uppercase;
  }}
  .chart-row{{display:flex;align-items:center;gap:10px;margin-bottom:8px;}}
  .chart-row .ck{{
    width:96px;color:var(--text);font-family:var(--mono);font-size:10px;
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  }}
  .bar-track{{flex:1;background:var(--surface3);border-radius:4px;height:6px;overflow:hidden;}}
  .bar-fill{{height:100%;border-radius:4px;transition:width .5s cubic-bezier(.4,0,.2,1);}}
  .chart-row .cv{{
    width:26px;text-align:right;color:var(--muted);
    font-family:var(--mono);font-size:10px;
  }}

  /* ── Table wrapper ───────────────────────────────────────────── */
  .table-wrap{{
    background:var(--surface);border:1px solid var(--border);
    border-radius:var(--radius);overflow:hidden;
  }}
  .table-header{{
    padding:14px 20px;border-bottom:1px solid var(--border);
    display:flex;align-items:center;gap:12px;
  }}
  .table-header h2{{font-family:var(--mono);font-size:11px;font-weight:700;
                    letter-spacing:1.5px;text-transform:uppercase;color:var(--muted);}}
  .rec-count{{
    background:var(--surface3);border:1px solid var(--border2);
    border-radius:20px;padding:2px 9px;font-family:var(--mono);font-size:10px;color:var(--muted);
  }}
  table{{width:100%;border-collapse:collapse;}}
  th{{
    background:var(--surface2);color:var(--muted);font-family:var(--mono);
    font-size:10px;font-weight:600;letter-spacing:1px;text-align:left;
    padding:10px 16px;border-bottom:1px solid var(--border);
    cursor:pointer;user-select:none;white-space:nowrap;transition:var(--transition);
    text-transform:uppercase;
  }}
  th:hover{{color:var(--text);}}
  th .sort-arrow{{opacity:.3;margin-left:4px;font-size:9px;}}
  th.sorted .sort-arrow{{opacity:1;color:var(--blue);}}
  td{{
    padding:11px 16px;border-bottom:1px solid var(--border);
    vertical-align:middle;font-size:12px;transition:background .1s;
  }}
  tr:last-child td{{border-bottom:none;}}
  tr:hover td{{background:var(--surface2);}}

  /* ── Badges ──────────────────────────────────────────────────── */
  .badge{{
    display:inline-flex;align-items:center;gap:4px;
    padding:3px 9px;border-radius:20px;font-size:10px;
    font-family:var(--mono);font-weight:700;white-space:nowrap;
    letter-spacing:.3px;
  }}
  .ba{{background:var(--green-dim);color:var(--green);border:1px solid #00d68f44;}}
  .bp{{background:var(--amber-dim);color:var(--amber);border:1px solid #f0a50044;}}
  .bn{{background:var(--red-dim);color:var(--red);border:1px solid #ff4d6d44;}}
  .br{{background:var(--blue-dim);color:var(--blue);border:1px solid #4d9fff44;}}
  .bh{{background:var(--purple-dim);color:var(--purple);border:1px solid #b57bee44;}}
  .bx{{background:var(--muted2);color:var(--muted);border:1px solid var(--border2);}}

  .hp{{
    display:inline-block;background:var(--surface3);border:1px solid var(--border2);
    border-radius:5px;padding:2px 8px;font-family:var(--mono);font-size:10px;
    margin:2px 2px 2px 0;color:var(--text);
  }}

  /* ── Pagination ──────────────────────────────────────────────── */
  .pag{{
    display:flex;gap:5px;align-items:center;
    padding:12px 16px;flex-wrap:wrap;border-top:1px solid var(--border);
  }}
  .pag button{{
    background:var(--surface2);border:1px solid var(--border);color:var(--muted);
    border-radius:5px;padding:4px 10px;font-size:11px;cursor:pointer;
    font-family:var(--mono);transition:var(--transition);
  }}
  .pag button.active{{background:var(--blue);border-color:var(--blue);color:#fff;}}
  .pag button:disabled{{opacity:.3;cursor:not-allowed;}}
  .pag button:not(.active):not(:disabled):hover{{border-color:var(--blue);color:var(--blue);}}
  .pag-info{{color:var(--muted);font-size:10px;margin-left:auto;font-family:var(--mono);}}
  .empty{{
    padding:56px;text-align:center;color:var(--muted);
    font-family:var(--mono);font-size:12px;letter-spacing:.5px;
  }}
  .empty-icon{{font-size:32px;margin-bottom:12px;opacity:.4;}}

  /* ── Notes cell ──────────────────────────────────────────────── */
  .note-chip{{
    display:inline-block;background:var(--blue-dim);border:1px solid var(--blue-dim);
    border-radius:4px;padding:2px 8px;font-size:10px;color:var(--blue);
    font-family:var(--mono);cursor:help;max-width:160px;
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  }}

  /* ── Modal ───────────────────────────────────────────────────── */
  .modal-bg{{
    display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);
    z-index:300;align-items:center;justify-content:center;backdrop-filter:blur(4px);
  }}
  .modal-bg.open{{display:flex;}}
  .modal{{
    background:var(--surface);border:1px solid var(--border2);
    border-radius:14px;padding:30px 32px;width:500px;
    max-width:95vw;max-height:88vh;overflow-y:auto;
    animation:modal-in .2s cubic-bezier(.4,0,.2,1);
  }}
  @keyframes modal-in{{from{{opacity:0;transform:translateY(-10px) scale(.97);}}to{{opacity:1;transform:none;}}}}
  .modal-title{{
    font-family:var(--mono);font-size:12px;font-weight:700;color:var(--blue);
    letter-spacing:1.5px;text-transform:uppercase;margin-bottom:6px;
  }}
  .modal-subtitle{{font-size:11px;color:var(--muted);margin-bottom:22px;}}
  .fr{{margin-bottom:16px;}}
  .fr label{{
    display:block;color:var(--muted);font-size:10px;font-weight:600;
    margin-bottom:6px;text-transform:uppercase;letter-spacing:1px;font-family:var(--mono);
  }}
  .fr select,.fr textarea{{
    width:100%;background:var(--surface2);border:1px solid var(--border2);
    border-radius:7px;color:var(--text);font-family:var(--mono);font-size:12px;
    padding:9px 12px;outline:none;resize:vertical;transition:var(--transition);
  }}
  .fr select:focus,.fr textarea:focus{{border-color:var(--blue);background:var(--surface3);}}
  .fr select option{{background:var(--surface2);}}
  .modal-actions{{display:flex;gap:8px;margin-top:22px;justify-content:flex-end;}}

  /* ── Toast ───────────────────────────────────────────────────── */
  .toast-stack{{position:fixed;bottom:24px;right:24px;display:flex;flex-direction:column;gap:8px;z-index:999;}}
  .toast{{
    background:var(--surface);border:1px solid var(--border2);border-radius:8px;
    padding:11px 18px;font-family:var(--mono);font-size:11px;
    opacity:0;transform:translateX(12px);transition:all .25s;color:var(--text);
    pointer-events:none;max-width:320px;
  }}
  .toast.show{{opacity:1;transform:translateX(0);}}
  .toast.t-ok{{border-left:3px solid var(--green);color:var(--green);}}
  .toast.t-err{{border-left:3px solid var(--red);color:var(--red);}}
  .toast.t-warn{{border-left:3px solid var(--amber);color:var(--amber);}}
  .toast.t-info{{border-left:3px solid var(--blue);color:var(--blue);}}
</style>
</head>
<body>

<!-- ── Unsaved-changes banner ──────────────────────────────── -->
<div id="dirtyBanner">
  <span class="dbmsg">&#x26A0;&nbsp; You have unsaved changes — they exist only in this browser tab.</span>
  <button class="btn btn-discard" onclick="discardChanges()">Discard</button>
</div>

<header>
  <div class="logo">
    <div class="logo-dot"></div>
    <div>
      <h1>SIEM Signoff Dashboard</h1>
      <div class="logo-sub" id="lastUpdated"></div>
    </div>
  </div>
  <div class="spacer"></div>
  <div class="hdr-controls">
    <div class="pill-group" id="pg">
      <button onclick="setPeriod('week',this)" class="active">7D</button>
      <button onclick="setPeriod('month',this)">30D</button>
      <button onclick="setPeriod('quarter',this)">90D</button>
      <button onclick="setPeriod('all',this)">ALL</button>
    </div>
    <div class="search-wrap">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
        <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
      </svg>
      <input type="text" id="search" placeholder="Host / sender / notes..." oninput="render()">
    </div>
    <button class="btn btn-exp" onclick="exportJSON()">&#x2B07; Export</button>
    <button class="btn btn-save" id="saveBtn" onclick="saveChanges()" disabled>
      &#x2713; Save Changes
    </button>
  </div>
</header>

<main>
  <div class="stats" id="statsCards"></div>
  <div class="charts-grid">
    <div class="chart-card"><h3>Status Breakdown</h3><div id="statusChart"></div></div>
    <div class="chart-card"><h3>Top Hostnames by Requests</h3><div id="hostChart"></div></div>
  </div>
  <div class="table-wrap">
    <div class="table-header">
      <h2>Signoff Log</h2>
      <span class="rec-count" id="recCount">0 records</span>
      <div class="spacer"></div>
    </div>
    <table>
      <thead><tr>
        <th onclick="sortBy('timestamp')" id="th-timestamp">
          Timestamp<span class="sort-arrow">&#x25BC;</span>
        </th>
        <th>Hostnames</th>
        <th onclick="sortBy('overall_status')" id="th-overall_status">
          Status<span class="sort-arrow">&#x25BC;</span>
        </th>
        <th onclick="sortBy('sender')" id="th-sender">
          Sender<span class="sort-arrow">&#x25BC;</span>
        </th>
        <th>Flags</th>
        <th>Notes</th>
        <th style="min-width:110px;text-align:right;">Actions</th>
      </tr></thead>
      <tbody id="logBody"></tbody>
    </table>
    <div class="pag" id="pag"></div>
  </div>
</main>

<!-- ── Edit modal ─────────────────────────────────────────────── -->
<div class="modal-bg" id="mb" onclick="if(event.target===this)closeMod()">
  <div class="modal">
    <div class="modal-title">Edit Record</div>
    <div class="modal-subtitle" id="modal-host-label"></div>
    <div class="fr">
      <label>Override Status</label>
      <select id="eStatus">
        <option value="">— keep current —</option>
        <option value="active">Active</option>
        <option value="partial">Partial</option>
        <option value="not_found">Not Found</option>
      </select>
    </div>
    <div class="fr">
      <label>Manually Resolved</label>
      <select id="eResolved">
        <option value="false">No — still open</option>
        <option value="true">Yes — resolved, exclude from dedup</option>
      </select>
    </div>
    <div class="fr">
      <label>Notes (ticket ID, actions taken...)</label>
      <textarea id="eNotes" rows="4" placeholder="e.g. TICKET-1234 — onboarding initiated..."></textarea>
    </div>
    <div class="modal-actions">
      <button class="btn" onclick="closeMod()">Cancel</button>
      <button class="btn btn-save" onclick="saveEdit()">&#x2713; Apply Changes</button>
    </div>
  </div>
</div>

<div class="toast-stack" id="toastStack"></div>

<script>
// ═══════════════════════════════════════════════
// Bootstrap
// ═══════════════════════════════════════════════
const EMBEDDED      = {json_blob};
const SERVER_ACTIVE = {server_active_js};
const SAVE_URL      = '{save_origin}/save';
const DATA_URL      = '{save_origin}/data.json';

let RAW;          // last persisted snapshot
let D;            // working copy — only differs from RAW when dirty
let _dirty = false;

let sf = 'timestamp', sasc = false, page = 1, eidx = null;
const PS = 15;
let _period = 'week';

async function init() {{
  if (SERVER_ACTIVE) {{
    try {{
      const r = await fetch(DATA_URL, {{cache:'no-cache'}});
      if (r.ok) {{
        RAW = await r.json();
        D   = deepClone(RAW);
        render();
        return;
      }}
    }} catch(e) {{
      console.warn('Live fetch failed, using embedded snapshot:', e);
    }}
  }}
  RAW = EMBEDDED;
  D   = deepClone(RAW);
  render();
}}

function deepClone(o) {{ return JSON.parse(JSON.stringify(o)); }}

// ═══════════════════════════════════════════════
// Dirty-state management
// ═══════════════════════════════════════════════
function markDirty() {{
  _dirty = true;
  document.getElementById('saveBanner')?.classList?.add('show');
  document.getElementById('dirtyBanner').classList.add('show');
  const btn = document.getElementById('saveBtn');
  btn.disabled = false;
  btn.classList.add('dirty');
}}

function markClean() {{
  _dirty = false;
  document.getElementById('dirtyBanner').classList.remove('show');
  const btn = document.getElementById('saveBtn');
  btn.disabled = true;
  btn.classList.remove('dirty');
}}

// ═══════════════════════════════════════════════
// Save / discard
// ═══════════════════════════════════════════════
async function saveChanges() {{
  if (!_dirty) return;
  const btn = document.getElementById('saveBtn');
  btn.textContent = 'Saving…';
  btn.disabled = true;

  if (SERVER_ACTIVE) {{
    try {{
      const r = await fetch(SAVE_URL, {{
        method: 'POST',
        headers: {{'Content-Type':'application/json'}},
        body: JSON.stringify(D),
      }});
      if (r.ok) {{
        RAW = deepClone(D);
        markClean();
        btn.textContent = '✓ Save Changes';
        showToast('Saved to signoff_data.json', 'ok');
        return;
      }}
      throw new Error('HTTP ' + r.status);
    }} catch(e) {{
      btn.disabled = false;
      btn.classList.add('dirty');
      btn.textContent = '✓ Save Changes';
      showToast('Server save failed — downloaded JSON instead. Replace signoff_data.json manually.', 'warn');
      _downloadJSON();
      return;
    }}
  }}

  // file:// mode — download
  _downloadJSON();
  RAW = deepClone(D);
  markClean();
  btn.textContent = '✓ Save Changes';
  showToast('Downloaded JSON — replace your signoff_data.json with this file', 'warn');
}}

function discardChanges() {{
  if (!confirm('Discard all unsaved changes and revert to the last saved state?')) return;
  D = deepClone(RAW);
  markClean();
  page = 1;
  render();
  showToast('Changes discarded', 'warn');
}}

// ═══════════════════════════════════════════════
// Delete
// ═══════════════════════════════════════════════
function deleteRow(gi) {{
  const entry = D.entries[gi];
  const label = (entry.hosts||[]).map(h=>h.hostname||'?').join(', ') || 'this record';
  if (!confirm(`Permanently delete the record for "${{label}}"?\\n\\nClick Save Changes to write this to disk.`)) return;
  D.entries.splice(gi, 1);
  markDirty();
  // Stay on same page if possible
  const pages = Math.max(1, Math.ceil(filtered().length / PS));
  if (page > pages) page = pages;
  render();
  showToast('Row deleted — click Save Changes to persist', 'warn');
}}

// ═══════════════════════════════════════════════
// Edit modal
// ═══════════════════════════════════════════════
function openMod(gi) {{
  eidx = gi;
  const e = D.entries[gi];
  const hosts = (e.hosts||[]).map(h=>h.hostname||'?').join(', ');
  document.getElementById('modal-host-label').textContent = hosts || e.email_subject || '';
  document.getElementById('eStatus').value   = e.overall_status || '';
  document.getElementById('eResolved').value = e.manually_resolved ? 'true' : 'false';
  document.getElementById('eNotes').value    = e.notes || '';
  document.getElementById('mb').classList.add('open');
}}

function closeMod() {{
  document.getElementById('mb').classList.remove('open');
  eidx = null;
}}

function saveEdit() {{
  if (eidx === null) return;
  const e  = D.entries[eidx];
  const ns = document.getElementById('eStatus').value;
  if (ns) e.overall_status = ns;
  e.manually_resolved = document.getElementById('eResolved').value === 'true';
  e.notes = document.getElementById('eNotes').value.trim();
  closeMod();
  markDirty();
  render();
  showToast('Edit applied — click Save Changes to persist', 'info');
}}

// ═══════════════════════════════════════════════
// Filtering / sorting / pagination
// ═══════════════════════════════════════════════
function periodCutoff() {{
  const days = {{week:7, month:30, quarter:90, all:36500}};
  return new Date(Date.now() - (days[_period]||7) * 86400000);
}}

function filtered() {{
  const q = (document.getElementById('search').value||'').toLowerCase().trim();
  const c = periodCutoff();
  return D.entries.filter(e => {{
    if (new Date(e.timestamp) < c) return false;
    if (!q) return true;
    const hosts  = (e.hosts||[]).map(x=>x.hostname||'').join(' ').toLowerCase();
    const sender = (e.sender||'').toLowerCase();
    const notes  = (e.notes||'').toLowerCase();
    const subj   = (e.email_subject||'').toLowerCase();
    return hosts.includes(q) || sender.includes(q) || notes.includes(q) || subj.includes(q);
  }});
}}

function srt(arr) {{
  return [...arr].sort((a, b) => {{
    let av, bv;
    if      (sf==='timestamp')      {{ av=a.timestamp||'';      bv=b.timestamp||''; }}
    else if (sf==='overall_status') {{ av=a.overall_status||''; bv=b.overall_status||''; }}
    else if (sf==='sender')         {{ av=a.sender||'';         bv=b.sender||''; }}
    else                            {{ av=''; bv=''; }}
    if (av < bv) return sasc ? -1 :  1;
    if (av > bv) return sasc ?  1 : -1;
    return 0;
  }});
}}

// ═══════════════════════════════════════════════
// Badges & formatting
// ═══════════════════════════════════════════════
function statusBadge(status, resolved, reval, hist) {{
  if (resolved) return '<span class="badge bx">&#x2714; Resolved</span>';
  if (hist)     return '<span class="badge bh">&#x2605; Historical</span>';
  const icons = {{active:'&#x2714;', partial:'&#x26A0;', not_found:'&#x2716;'}};
  const cls   = {{active:'ba',       partial:'bp',       not_found:'bn'}};
  const label = {{active:'Active',   partial:'Partial',  not_found:'Not Found'}};
  const base  = `<span class="badge ${{cls[status]||'bx'}}">${{icons[status]||'?'}}&nbsp;${{label[status]||status}}</span>`;
  return base + (reval ? ' <span class="badge br">&#x21BB; Reval</span>' : '');
}}

function fmtDate(iso) {{
  if (!iso) return '&mdash;';
  const d = new Date(iso);
  const date = d.toLocaleDateString('en-GB',{{day:'2-digit',month:'short',year:'numeric'}});
  const time = d.toLocaleTimeString('en-GB',{{hour:'2-digit',minute:'2-digit'}});
  return `<span style="color:var(--text)">${{date}}</span> <span style="color:var(--muted);font-size:10px">${{time}}</span>`;
}}

function fmtSender(sender) {{
  if (!sender) return '<span style="color:var(--muted2)">—</span>';
  const parts = sender.split('@');
  if (parts.length < 2) return `<span style="font-family:var(--mono);font-size:11px">${{sender}}</span>`;
  return `<span style="font-family:var(--mono);font-size:11px;color:var(--text)">${{parts[0]}}</span>` +
         `<span style="font-family:var(--mono);font-size:10px;color:var(--muted)">@${{parts[1]}}</span>`;
}}

// ═══════════════════════════════════════════════
// Render
// ═══════════════════════════════════════════════
function render() {{
  const e     = srt(filtered());
  const tot   = e.length;
  const pages = Math.max(1, Math.ceil(tot / PS));
  if (page > pages) page = 1;
  const sl = e.slice((page-1)*PS, page*PS);

  document.getElementById('recCount').textContent =
    `${{tot}} record${{tot!==1?'s':''}}`;

  // Update sort header highlights
  ['timestamp','overall_status','sender'].forEach(f => {{
    const th = document.getElementById('th-'+f);
    if (!th) return;
    th.classList.toggle('sorted', sf===f);
    const arrow = th.querySelector('.sort-arrow');
    if (arrow) arrow.textContent = (sf===f && sasc) ? '▲' : '▼';
  }});

  const body = document.getElementById('logBody');
  if (!sl.length) {{
    body.innerHTML = `<tr><td colspan="7">
      <div class="empty"><div class="empty-icon">&#x1F50D;</div>No records match this filter.</div>
    </td></tr>`;
  }} else {{
    body.innerHTML = sl.map((x, i) => {{
      // gi = true index in D.entries, stable even after filtering/sorting
      const gi    = D.entries.indexOf(e[(page-1)*PS + i]);
      const hosts = (x.hosts||[]).map(h =>
        `<span class="hp" title="${{h.hostname||''}}">${{h.hostname||'?'}}</span>`
      ).join('');

      const flagParts = [];
      if (x.historical)      flagParts.push('<span style="color:var(--purple);font-size:10px;font-family:var(--mono)">imported</span>');
      else if (x.prior_status) flagParts.push(`<span style="color:var(--muted);font-size:10px;font-family:var(--mono)">prior:&nbsp;${{x.prior_status.replace(/\[Processed-?/i,'').replace(']','').toLowerCase()}}</span>`);
      else                   flagParts.push('<span style="color:var(--muted2);font-size:10px;font-family:var(--mono)">new</span>');

      const noteCell = x.notes
        ? `<span class="note-chip" title="${{x.notes.replace(/"/g,'&quot;')}}">${{x.notes}}</span>`
        : '<span style="color:var(--muted2);font-size:10px">—</span>';

      return `<tr>
        <td style="white-space:nowrap;font-family:var(--mono);font-size:11px;">${{fmtDate(x.timestamp)}}</td>
        <td>${{hosts}}</td>
        <td>${{statusBadge(x.overall_status, x.manually_resolved, x.is_revalidation, x.historical)}}</td>
        <td style="white-space:nowrap;">${{fmtSender(x.sender)}}</td>
        <td>${{flagParts.join('')}}</td>
        <td>${{noteCell}}</td>
        <td style="text-align:right;white-space:nowrap;">
          <button class="btn" onclick="openMod(${{gi}})">Edit</button>
          <button class="btn btn-del" onclick="deleteRow(${{gi}})" title="Delete this record">&#x2715;</button>
        </td>
      </tr>`;
    }}).join('');
  }}

  // Pagination
  let ph = `<button onclick="goP(${{page-1}})" ${{page===1?'disabled':''}}>&#x2039; Prev</button>`;
  const s = Math.max(1,page-2), en = Math.min(pages,page+2);
  if (s>1) ph += `<button onclick="goP(1)">1</button>${{s>2?'<span style="color:var(--muted);padding:0 4px">…</span>':''}}`;
  for (let p=s;p<=en;p++) ph += `<button onclick="goP(${{p}})" class="${{p===page?'active':''}}">${{p}}</button>`;
  if (en<pages) ph += `${{en<pages-1?'<span style="color:var(--muted);padding:0 4px">…</span>':''}}<button onclick="goP(${{pages}})">${{pages}}</button>`;
  ph += `<button onclick="goP(${{page+1}})" ${{page===pages?'disabled':''}}>Next &#x203A;</button>
         <span class="pag-info">${{tot}} record${{tot!==1?'s':''}}&nbsp;·&nbsp;page ${{page}}/${{pages}}</span>`;
  document.getElementById('pag').innerHTML = ph;

  renderStats(filtered());
  renderCharts(filtered());
}}

function goP(p) {{ page=p; render(); window.scrollTo({{top:0,behavior:'smooth'}}); }}
function sortBy(f) {{ sf===f ? sasc=!sasc : (sf=f, sasc=false); render(); }}
function setPeriod(p,btn) {{
  _period=p; page=1;
  document.querySelectorAll('#pg button').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  render();
}}

// ═══════════════════════════════════════════════
// Stats cards
// ═══════════════════════════════════════════════
function renderStats(e) {{
  const total  = e.length;
  const active = e.filter(x=>x.overall_status==='active'    && !x.manually_resolved && !x.historical).length;
  const partial= e.filter(x=>x.overall_status==='partial'   && !x.manually_resolved && !x.historical).length;
  const notfnd = e.filter(x=>x.overall_status==='not_found' && !x.manually_resolved && !x.historical).length;
  const hist   = e.filter(x=>x.historical).length;
  const revals = e.filter(x=>x.is_revalidation).length;
  const pct    = total ? Math.round(active/total*100) : 0;

  document.getElementById('statsCards').innerHTML = `
    <div class="card b"><div class="card-accent"></div>
      <div class="num">${{total}}</div>
      <div class="clabel">Total Signoffs</div>
      <div class="sub">${{revals}} revalidation${{revals!==1?'s':''}}</div></div>
    <div class="card g"><div class="card-accent"></div>
      <div class="num">${{active}}</div>
      <div class="clabel">Active</div>
      <div class="sub">${{pct}}% of period</div></div>
    <div class="card a"><div class="card-accent"></div>
      <div class="num">${{partial}}</div>
      <div class="clabel">Partial</div>
      <div class="sub">Missing log sources</div></div>
    <div class="card r"><div class="card-accent"></div>
      <div class="num">${{notfnd}}</div>
      <div class="clabel">Not Found</div>
      <div class="sub">${{e.filter(x=>x.manually_resolved).length}} resolved</div></div>
    <div class="card p"><div class="card-accent"></div>
      <div class="num">${{hist}}</div>
      <div class="clabel">Historical</div>
      <div class="sub">Legacy imported</div></div>`;
}}

// ═══════════════════════════════════════════════
// Charts
// ═══════════════════════════════════════════════
function renderCharts(e) {{
  // Status chart
  const sm = {{}};
  e.forEach(x => {{
    const k = x.manually_resolved ? 'resolved' : (x.historical ? 'historical' : (x.overall_status||'unknown'));
    sm[k] = (sm[k]||0) + 1;
  }});
  const colorMap = {{
    active:'var(--green)', partial:'var(--amber)', not_found:'var(--red)',
    resolved:'var(--muted)', historical:'var(--purple)', unknown:'var(--border2)',
  }};
  barChart(
    document.getElementById('statusChart'),
    Object.entries(sm).map(([k,v])=>{{return{{k,v}}}}).sort((a,b)=>b.v-a.v),
    k => colorMap[k] || 'var(--blue)'
  );

  // Host chart — count by hostname across all entries in period
  const hm = {{}};
  e.forEach(x => (x.hosts||[]).forEach(h => {{
    if (h.hostname) hm[h.hostname] = (hm[h.hostname]||0) + 1;
  }}));
  barChart(
    document.getElementById('hostChart'),
    Object.entries(hm).map(([k,v])=>{{return{{k,v}}}}).sort((a,b)=>b.v-a.v).slice(0,10),
    () => 'var(--blue)'
  );
}}

function barChart(el, items, colorFn) {{
  if (!items.length) {{ el.innerHTML = '<div style="color:var(--muted);font-family:var(--mono);font-size:11px;text-align:center;padding:20px 0;">No data</div>'; return; }}
  const mx = Math.max(...items.map(i=>i.v), 1);
  el.innerHTML = items.map(i => `
    <div class="chart-row">
      <span class="ck" title="${{i.k}}">${{i.k}}</span>
      <div class="bar-track">
        <div class="bar-fill" style="width:${{Math.round(i.v/mx*100)}}%;background:${{colorFn(i.k)}}"></div>
      </div>
      <span class="cv">${{i.v}}</span>
    </div>`).join('');
}}

// ═══════════════════════════════════════════════
// Export
// ═══════════════════════════════════════════════
function _downloadJSON() {{
  const b = new Blob([JSON.stringify(D,null,2)],{{type:'application/json'}});
  Object.assign(document.createElement('a'),{{
    href:URL.createObjectURL(b), download:'signoff_data.json'
  }}).click();
}}
function exportJSON() {{
  _downloadJSON();
  showToast('Exported current view as JSON', 'info');
}}

// ═══════════════════════════════════════════════
// Toast
// ═══════════════════════════════════════════════
function showToast(msg, type='ok') {{
  const stack = document.getElementById('toastStack');
  const el    = document.createElement('div');
  el.className = `toast t-${{type}}`;
  el.textContent = msg;
  stack.appendChild(el);
  requestAnimationFrame(() => {{ el.classList.add('show'); }});
  setTimeout(() => {{
    el.classList.remove('show');
    setTimeout(() => el.remove(), 300);
  }}, 4000);
}}

// ═══════════════════════════════════════════════
// Boot
// ═══════════════════════════════════════════════
const _last = EMBEDDED.entries?.length
  ? EMBEDDED.entries[EMBEDDED.entries.length-1]?.timestamp
  : null;
document.getElementById('lastUpdated').textContent =
  _last
    ? 'Last run: ' + new Date(_last).toLocaleDateString('en-GB',{{day:'2-digit',month:'short',year:'numeric'}})
    : 'No data yet';

init();
</script>
</body>
</html>"""

    try:
        with open(DASHBOARD_PATH, 'w', encoding='utf-8') as f:
            f.write(html)
    except Exception as e:
        _log(f"WARNING: Dashboard HTML write error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════════════════

def generate_dashboard():
    """Start HTTP server (if enabled), write HTML, open browser."""
    _log("Generating dashboard...")
    port = _start_dashboard_server()    # starts server + timer; returns 0 if disabled
    _write_dashboard_html(port)

    if port:
        url = f'http://127.0.0.1:{port}/'
        _log(f"Dashboard URL : {url}")
        _log(f"  Edits will auto-save to signoff_data.json (server up for {DASHBOARD_SERVE_MINUTES} min).")
    else:
        url = "file:///" + DASHBOARD_PATH.replace(os.sep, '/')
        _log(f"Dashboard URL : {url}")
        _log("  Note: server disabled (DASHBOARD_SERVE_MINUTES=0). "
             "Saves will download JSON — replace signoff_data.json manually.")

    try:
        webbrowser.open(url)
    except Exception as e:
        _log(f"WARNING: Could not open browser: {e}")


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
    _log("QRadar Signoff Auto-Draft v2.2 starting...")
    _log(f"  Inbox: {LOOKBACK_DAYS}d | Sent: {SENT_SCAN_DAYS}d | "
         f"Cooldown: {'off' if not REVALIDATION_COOLDOWN_DAYS else f'{REVALIDATION_COOLDOWN_DAYS}d'} | "
         f"Active-skip: {ACTIVE_SKIP_DAYS}d")
    _log("  Runtime dedup: ON | Cross-thread dedup: ON | Mode: DRAFT ONLY")

    if HISTORICAL_RUN_MODE:
        _log("")
        _log("  *** HISTORICAL RUN MODE ENABLED ***")
        _log("  *** Already-tagged emails will be imported (no QRadar query, no draft). ***")
        _log("  *** Remember to set HISTORICAL_RUN_MODE = False after this run! ***")
        _log("")

    if not acquire_lock():
        return

    try:
        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        # In historical-only mode we skip QRadar entirely; for mixed runs we still connect.
        if not HISTORICAL_RUN_MODE:
            if not test_qradar_connection():
                _log("ERROR: QRadar unreachable — aborting. Emails untouched.")
                return
            fetch_log_source_types()
        else:
            _log("Historical mode — skipping QRadar connectivity check.")

        cutoff_str      = (datetime.now() - timedelta(days=LOOKBACK_DAYS)).strftime('%m/%d/%Y %I:%M %p')
        cooldown_cutoff = (
            datetime.now() - timedelta(days=REVALIDATION_COOLDOWN_DAYS)
            if REVALIDATION_COOLDOWN_DAYS > 0 else None
        )
        active_skip_cutoff = datetime.now() - timedelta(days=ACTIVE_SKIP_DAYS)

        inbox_items = list(inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'"))
        _log(f"\n{len(inbox_items)} email(s) in last {LOOKBACK_DAYS}d")

        processed = skipped = drafted = revalidated = historical_count = 0

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

            # Subject guards (allows tagged in HISTORICAL_RUN_MODE)
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

            # ── Historical import path ────────────────────────────────────
            # Triggered when the email already carries a [Processed*] tag,
            # which means a previous version of this script handled it.
            has_processed_tag = '[processed' in subject.lower()
            if HISTORICAL_RUN_MODE and has_processed_tag:
                if _historical_already_imported(subject):
                    skipped += 1
                    _log(f"  HIST-SKIP (already imported): '{subject[:60]}'")
                    continue

                existing_tag = _tag_from_subject(subject)
                tag_to_status = {
                    TAG_ACTIVE:    'active',
                    TAG_PARTIAL:   'partial',
                    TAG_NOT_FOUND: 'not_found',
                    'legacy':      'active',
                }
                hist_status = tag_to_status.get(existing_tag, 'active')

                try:
                    received_dt = _com_dt_to_py(mail_item.ReceivedTime)
                except Exception:
                    received_dt = None

                _log(f"  HISTORICAL [{hist_status.upper()}]: '{subject[:70]}'")
                _log(f"    Sender: {sender} | Hosts: {hostname_list} | "
                     f"Date: {received_dt.strftime('%Y-%m-%d') if received_dt else 'unknown'}")

                write_historical_record(
                    email_subject=subject,
                    sender=sender,
                    hostname_list=hostname_list,
                    status=hist_status,
                    received_dt=received_dt,
                )
                historical_count += 1
                processed        += 1
                continue

            # ── Runtime dedup ─────────────────────────────────────────────
            if is_drafted_this_run(hostname_list):
                skipped += 1
                _log(f"  SKIP (runtime dedup — already handled {hostname_list} this run)")
                continue

            _log(f"\n  Candidate: '{subject[:70]}'")
            _log(f"    Sender: {sender} | Hosts: {hostname_list}")

            # ── Conversation + cross-thread state ─────────────────────────
            last_tag, last_dt = check_conversation_status(
                mail_item, sent, drafts, hostname_list
            )

            if last_tag in (TAG_ACTIVE, 'legacy'):
                if last_dt and last_dt >= active_skip_cutoff:
                    skipped += 1
                    _log(f"    SKIP (Active on {last_dt.strftime('%Y-%m-%d')} — within {ACTIVE_SKIP_DAYS}d)")
                    continue
                _log(f"    Active result > {ACTIVE_SKIP_DAYS}d old — allowing recheck")
                is_revalidation = True

            elif last_tag in REVALIDATABLE_TAGS:
                if cooldown_cutoff and last_dt and last_dt >= cooldown_cutoff:
                    skipped += 1
                    _log(f"    SKIP (cooldown: {(datetime.now()-last_dt).days}d ago)")
                    continue
                _log(f"    Revalidating {last_tag} from "
                     f"{last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'}")
                is_revalidation = True

            elif last_tag is None:
                _log(f"    New signoff")
                is_revalidation = False

            else:
                skipped += 1
                _log(f"    SKIP (unknown tag: {last_tag})")
                continue

            # ── QRadar query + draft ──────────────────────────────────────
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
                mark_drafted_this_run(hostname_list)
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
        _log(f"Done — {processed} processed | {drafted} drafted "
             f"({revalidated} reval) | {skipped} skipped"
             + (f" | {historical_count} historical imported" if historical_count else ""))

        if HISTORICAL_RUN_MODE and historical_count > 0:
            _log("")
            _log(f"  *** Historical import complete: {historical_count} record(s) added. ***")
            _log(f"  *** ACTION REQUIRED: Set HISTORICAL_RUN_MODE = False before next run! ***")

    finally:
        release_lock()
        _print_paths()
        generate_dashboard()


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT  (FIX 3: clean Ctrl+C / crash exit)
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        _log("\nInterrupted by user — shutting down server if running.")
        if _http_server_instance:
            try:
                _http_server_instance.shutdown()
            except Exception:
                pass
