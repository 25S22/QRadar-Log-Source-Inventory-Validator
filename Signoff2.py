"""
QRadar Signoff Auto-Draft  v3.0
=================================
WHAT CHANGED / FIXED vs ALL PREVIOUS VERSIONS
──────────────────────────────────────────────
CONFIGURATION
  • Exact config format from reference document (LOOKBACK_DAYS, SENT_SCAN_DAYS,
    REVALIDATION_COOLDOWN_DAYS, ACTIVE_SKIP_DAYS, STATUS_PRIORITY, _runtime_drafted_hosts)
  • All paths auto-derived from _SCRIPT_DIR — zero manual config required on first run
  • _OVERRIDES_PATH derived from SIGNOFF_DATA_PATH (no extra config key needed)

DEDUPLICATION — three independent layers
  1. _runtime_drafted_hosts (frozenset): blocks same hostname-set being drafted
     twice in one execution (e.g. script stuck in loop or large inbox)
  2. ACTIVE_SKIP_DAYS: reads data file — if hostname-set was Active within N days,
     skip entirely without touching Outlook or QRadar
  3. Outlook Sent/Drafts tag scan (SENT_SCAN_DAYS): detects outcome tags in
     conversation thread for partial/not_found cooldown decisions

LOGICAL FLAWS FIXED
  • STATUS_PRIORITY {'not_found':2,'partial':1,'active':0} — worst = max(),
    not min(); all previous versions used the wrong comparator
  • Outlook Restrict: tries ISO format first, falls back to manual iteration
    (locale-safe) — previous versions hardcoded US locale date string
  • ConversationID access wrapped in try/except for every item (meeting
    responses and NDRs in inbox don't expose this property and previously
    caused silent AttributeError loops)
  • Empty ESCALATION_TO/CC with partial/not_found: falls back to ReplyAll
    instead of saving a draft with blank To: (Outlook raises COMException)
  • extract_hostnames: collapses inner whitespace, deduplicates order-preserving
    (duplicate hostnames in subject caused double QRadar queries)
  • _safe_timestamp: explicit float→int cast before epoch heuristic — Python
    float precision caused false "epoch > 4102444800" on some QRadar versions
  • detect_os_group: skips groups with empty 'required' list instead of
    IndexError on required[0]
  • validate_expected_types: 'days_ago is None' now correctly treated as
    "found but silent" not "not found" in every code path
  • build_reply_for_all_hosts: uses STATUS_PRIORITY max() correctly
  • SENT_SCAN_DAYS respected in Restrict filter (was REVALIDATION_WINDOW_DAYS)
  • Dashboard edits/deletes now persistent via localStorage — survive refresh,
    browser close, and dashboard HTML regeneration (data re-embedded each run,
    localStorage state layered on top)
  • Delete requires mandatory exception reason; enforced both in JS and noted
    in export so the script side can log it
  • Undo delete: 10-second grace period with live countdown
  • Bar chart stacking order: not_found (bottom) → partial → active (top),
    so most-visible colour (green) correctly represents good results on top
  • Pure-SVG charts — no external CDN dependency (works air-gapped)
  • Pagination (25 rows / page) with keyboard arrow-key navigation
  • Next-revalidation countdown column calculated from last_checked + cooldown
  • Status filter pills + OS-group dropdown + free-text search combined
  • Export CSV and Export Overrides JSON on dashboard toolbar
  • Overrides JSON exported from dashboard can be placed at _OVERRIDES_PATH
    for the script to skip exception hosts on next run
  • VERSION written into every tracking record for audit traceability
  • QRadar API: retry (3 attempts, exponential back-off) + Range-header
    pagination (50 items/page) via centralised _qradar_get()
  • Sent Items Restrict restricted to SENT_SCAN_DAYS not full mailbox scan
  • Tags in Drafts: LastModificationTime used (SentOn undefined on drafts)
  • _log: handles IOError/PermissionError on log file without crashing run
  • release_lock: always called via finally — previous bare try/except could
    leave lock in place on unexpected SystemExit
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
ESCALATION_TO      = ['onboarding-owner@yourorg.com']
ESCALATION_CC      = ['@SOC-DL@yourorg.com']
ESCALATION_CONTACT = '@xyz'   # shown inline in escalation rows

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

# ─── DERIVED PATHS (not in config block — auto-derived) ──────────────────────
_OVERRIDES_PATH = SIGNOFF_DATA_PATH.replace('.json', '_overrides.json')
VERSION         = '3.0'
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
    """Converts pywintypes COM datetime to naive Python datetime (local time)."""
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
#  DATA FILE  (records + overrides in one JSON)
# ═══════════════════════════════════════════════════════════════════════════════

def _load_data() -> dict:
    """
    Loads the combined data file.
    Schema: {"records": [...], "overrides": {...}}
    Returns a default-empty structure if missing or corrupt.
    """
    if not os.path.exists(SIGNOFF_DATA_PATH):
        return {"records": [], "overrides": {}}
    try:
        with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Handle legacy format (plain list of records from v1/v2)
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
    """
    Reads overrides from _OVERRIDES_PATH (dashboard export).
    Falls back to overrides embedded in SIGNOFF_DATA_PATH.
    Schema: {"HOSTNAME": {"deleted": bool, "status_override": str|null,
                          "note": str, "exception_reason": str}}
    """
    # Prefer the dedicated overrides file (exported from dashboard)
    if os.path.exists(_OVERRIDES_PATH):
        try:
            with open(_OVERRIDES_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            _log(f"WARN: Could not read overrides file: {e}")
    # Fall back to embedded overrides in main data file
    return _load_data().get("overrides", {})


# ═══════════════════════════════════════════════════════════════════════════════
#  CROSS-RUN DEDUPLICATION  (data-file based)
# ═══════════════════════════════════════════════════════════════════════════════

def _was_active_recently(hostname_frozenset: frozenset) -> tuple[bool, str]:
    """
    FIX: STATUS_PRIORITY uses max() (higher number = worse status).
    Checks data file records to see if the SAME hostname set was processed
    as Active within ACTIVE_SKIP_DAYS. Returns (True, date_str) or (False, '').
    """
    if ACTIVE_SKIP_DAYS <= 0:
        return False, ''
    cutoff = datetime.now() - timedelta(days=ACTIVE_SKIP_DAYS)
    data   = _load_data()
    for rec in reversed(data["records"]):   # most recent first
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
    """
    Centralised QRadar GET with:
      - 3 retries, exponential back-off (2s, 4s, 8s)
      - Range-header pagination (50 items/page) when paginate=True
      - Returns (status_code, body) — body is [] on total failure
    FIX: retry only on 5xx and connection errors, not on 4xx (auth fails etc.)
    """
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
                    return r               # success or 4xx — no retry
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

    # Paginated fetch via Range header
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
        # Parse Content-Range: items START-END/TOTAL
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
    """Paginates through log source types and populates the module-level cache."""
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
#  QRADAR QUERIES — STRICTLY READ-ONLY (HTTP GET ONLY)
# ═══════════════════════════════════════════════════════════════════════════════

def _safe_timestamp(timestamp_ms) -> tuple:
    """
    FIX: explicit float→int before epoch heuristic prevents float precision
    causing a valid ms timestamp to appear > 4102444800 on some QRadar builds.
    """
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    try:
        ts = int(timestamp_ms)                         # FIX: cast first
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
    """
    READ-ONLY paginated fetch of every log source matching hostname.
    Only HTTP GET. Nothing in QRadar is modified.
    """
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
    """
    FIX: skips groups with an empty 'required' list instead of IndexError on
    required[0]. First matching group by key-insertion order wins.
    """
    if not OS_TYPE_GROUPS:
        return None, None
    for group_name, rules in OS_TYPE_GROUPS.items():
        required = rules.get('required', [])
        if not required:          # FIX: guard added
            continue
        sig_words = str(required[0]).lower().split()
        if any(
            all(w in str(s.get('ls_type', '')).lower() for w in sig_words)
            for s in sources
        ):
            return group_name, rules
    return None, None


def validate_expected_types(all_sources_result: dict, required_types: list) -> list:
    """
    FIX: days_ago is None now correctly means "found but silent" (no events),
    not "not found". This affects both row colouring and needs_cc logic.
    """
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
#  SUBJECT PARSING — MULTI-HOSTNAME
# ═══════════════════════════════════════════════════════════════════════════════

def passes_subject_guards(subject: str) -> tuple:
    """
    '[processed' without a closing bracket catches ALL tag variants:
    [Processed-Active], [Processed-Partial], [Processed-NotFound],
    and the legacy [Processed] format from v1.
    """
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
    """
    FIX: collapses inner whitespace (handles tabs / double spaces),
    deduplicates order-preserving (prevents double QRadar queries).

    "Security Signoff | HOST1 | HOST2 | HOST1 |  |" → ['HOST1', 'HOST2']
    "Security Signoff |  HOST1  | HOST2"             → ['HOST1', 'HOST2']
    """
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
#  CONVERSATION STATUS  (Outlook Sent Items + Drafts)
# ═══════════════════════════════════════════════════════════════════════════════

def check_conversation_status(mail_item, sent_folder, drafts_folder) -> tuple:
    """
    FIX 1: ConversationID access wrapped in try/except for every item —
            meeting responses / NDRs in inbox don't expose this property.
    FIX 2: Restrict tries ISO date string first; falls back to full scan
            if COM raises (locale mismatch on non-US Windows installs).
    FIX 3: SENT_SCAN_DAYS used for Restrict cutoff (was REVALIDATION_WINDOW_DAYS).
    FIX 4: Drafts use LastModificationTime (SentOn is undefined on unsent items).
    """
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
                if item.ConversationID == conv_id:    # FIX 1
                    _update(_extract_tag(item.Subject),
                            _com_dt_to_py(getattr(item, dt_attr, None)))
            except Exception:
                continue

    # ── Sent Items ────────────────────────────────────────────────────────────
    cutoff_iso = (datetime.now() - timedelta(days=SENT_SCAN_DAYS)
                  ).strftime('%Y-%m-%d')              # FIX 2 & 3
    try:
        restricted = sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff_iso}'")
        _scan_folder(sent_folder, restricted, 'SentOn')
    except Exception:
        try:                                           # FIX 2: locale fallback
            _scan_folder(sent_folder, sent_folder.Items, 'SentOn')
        except Exception as e:
            _log(f"   WARN: Could not scan Sent Items: {e}")

    # ── Drafts ────────────────────────────────────────────────────────────────
    try:
        _scan_folder(drafts_folder, drafts_folder.Items, 'LastModificationTime')  # FIX 4
    except Exception as e:
        _log(f"   WARN: Could not scan Drafts: {e}")

    return last_tag, last_dt


# ═══════════════════════════════════════════════════════════════════════════════
#  HTML REPLY BUILDER — MULTI-HOSTNAME
# ═══════════════════════════════════════════════════════════════════════════════

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
    # FIX: 'found but silent' (days_ago is None) correctly treated as problem
    any_problem = any(
        not r['found'] or (r['found'] and r['days_ago'] is None)
        for r in validation
    )
    return ('partial' if any_problem else 'active'), validation, group_name


def _host_section_html(hostname: str, host_status: str,
                        type_validation, os_group, sources: list) -> str:
    COLORS = {'active': '#22c55e', 'partial': '#f59e0b', 'not_found': '#ef4444'}
    col = COLORS.get(host_status, '#888')
    osl = f' &mdash; {os_group}' if os_group else ''
    hdr = (
        f'<div style="margin:18px 0 5px;padding:7px 14px;'
        f'background:{col}12;border-left:3px solid {col};border-radius:0 6px 6px 0;">'
        f'<span style="font-family:Consolas,monospace;font-weight:700;'
        f'color:{col};font-size:12px">{hostname}</span>'
        f'<span style="color:#8898b8;font-size:11px">{osl}</span></div>'
    )
    if host_status == 'not_found':
        return hdr + (
            '<p style="font-size:12px;color:#ef4444;margin:4px 0 0 14px;">'
            'Not found in QRadar log source inventory.</p>'
        )

    if type_validation is not None:
        rows = ''
        for r in type_validation:
            ds = (
                f' <span style="color:#5a6f90;font-size:10px">'
                f"({'Today' if r['days_ago']==0 else str(r['days_ago'])+'d ago'})"
                f'</span>'
                if r['days_ago'] is not None else ''
            )
            if not r['found']:
                note = (f"{ESCALATION_CONTACT} please onboard this log source."
                        if ESCALATION_CONTACT.strip() else
                        "Not found — please onboard this log source.")
                rows += (
                    f'<tr style="background:#2d1010"><td style="color:#ef4444;font-weight:700;'
                    f'text-align:center;padding:7px 10px;border-bottom:1px solid #3d1515;'
                    f'font-size:11px;width:20px">&#10006;</td>'
                    f'<td style="color:#ef4444;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #3d1515;font-size:11px">{r["expected"]}</td>'
                    f'<td style="color:#ef4444;padding:7px 10px;border-bottom:1px solid #3d1515;'
                    f'font-size:11px">&#8212;</td>'
                    f'<td style="color:#ef4444;padding:7px 10px;border-bottom:1px solid #3d1515;'
                    f'font-size:11px">&#8212;</td>'
                    f'<td style="color:#ef4444;font-style:italic;padding:7px 10px;'
                    f'border-bottom:1px solid #3d1515;font-size:11px">{note}</td></tr>'
                )
            elif r['days_ago'] is None:
                note = (f"{ESCALATION_CONTACT} no events received yet — please investigate."
                        if ESCALATION_CONTACT.strip() else
                        "No events received yet — please investigate.")
                rows += (
                    f'<tr style="background:#2d1f08"><td style="color:#f59e0b;font-weight:700;'
                    f'text-align:center;padding:7px 10px;border-bottom:1px solid #3d2d10;'
                    f'font-size:11px;width:20px">&#9888;</td>'
                    f'<td style="color:#f59e0b;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #3d2d10;font-size:11px">{r["expected"]}</td>'
                    f'<td style="color:#c8d4eb;padding:7px 10px;border-bottom:1px solid #3d2d10;'
                    f'font-size:11px">{r.get("ls_name","N/A")}</td>'
                    f'<td style="color:#f59e0b;font-style:italic;padding:7px 10px;'
                    f'border-bottom:1px solid #3d2d10;font-size:11px">No events recorded</td>'
                    f'<td style="color:#f59e0b;font-style:italic;padding:7px 10px;'
                    f'border-bottom:1px solid #3d2d10;font-size:11px">{note}</td></tr>'
                )
            else:
                rows += (
                    f'<tr style="background:#0d2e1a"><td style="color:#22c55e;font-weight:700;'
                    f'text-align:center;padding:7px 10px;border-bottom:1px solid #163d23;'
                    f'font-size:11px;width:20px">&#10004;</td>'
                    f'<td style="color:#e2e8f0;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #163d23;font-size:11px">{r["expected"]}</td>'
                    f'<td style="color:#c8d4eb;padding:7px 10px;border-bottom:1px solid #163d23;'
                    f'font-size:11px">{r.get("ls_name","N/A")}</td>'
                    f'<td style="color:#c8d4eb;padding:7px 10px;border-bottom:1px solid #163d23;'
                    f'font-size:11px">{r.get("last_seen","N/A")}{ds}</td>'
                    f'<td style="color:#22c55e;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #163d23;font-size:11px">Confirmed</td></tr>'
                )
        th = lambda t: (f'<th style="padding:6px 10px;border-bottom:1px solid #2a3650;'
                        f'text-align:left;color:#8898b8;font-size:10px;'
                        f'text-transform:uppercase;font-weight:600">{t}</th>')
        return hdr + (
            f'<table style="width:100%;border-collapse:collapse;margin:4px 0 0;'
            f'border:1px solid #2a3650;border-radius:4px;overflow:hidden;font-size:11px">'
            f'<tr style="background:#111827"><th style="width:20px;padding:6px 10px;'
            f'border-bottom:1px solid #2a3650"></th>'
            + th('Log Source Type') + th('Log Source Name') + th('Last Event') + th('Status')
            + f'</tr>{rows}</table>'
        )

    # Simple mode
    best = None
    if sources:
        en = sorted([s for s in sources if s.get('enabled')],
                    key=lambda x: (x.get('days_ago') is None, x.get('days_ago') or 99999))
        di = sorted([s for s in sources if not s.get('enabled')],
                    key=lambda x: (x.get('days_ago') is None, x.get('days_ago') or 99999))
        best = en[0] if en else (di[0] if di else None)
    if best:
        da  = best.get('days_ago')
        dsp = 'Today' if da == 0 else (f"{da}d ago" if da is not None else 'N/A')
        td  = lambda l, v: (
            f'<tr><td style="padding:6px 12px;color:#8898b8;font-size:12px;'
            f'border-bottom:1px solid #1e2d45;width:140px">{l}</td>'
            f'<td style="padding:6px 12px;font-size:12px;border-bottom:1px solid #1e2d45;'
            f'color:#c8d4eb">{v}</td></tr>'
        )
        return hdr + (
            f'<table style="width:100%;max-width:480px;border-collapse:collapse;'
            f'margin:4px 0 0;border:1px solid #2a3650;border-radius:4px;overflow:hidden">'
            + td('Log Source Name',
                 f'<b style="color:#f0f4ff">{best.get("name","N/A")}</b>')
            + td('Log Source Type', best.get("ls_type","N/A"))
            + td('Last Event',
                 f'{best.get("last_seen","N/A")} '
                 f'<span style="color:#5a6f90;font-size:11px">({dsp})</span>')
            + '</table>'
        )
    return hdr


def build_reply_for_all_hosts(hostname_qr_pairs: list) -> tuple:
    """
    FIX: uses STATUS_PRIORITY with max() — higher value = worse status.
    'not_found':2 > 'partial':1 > 'active':0  ← correct with max().
    Previous version used min() with an inverted dict which was wrong.

    FIX: empty ESCALATION_TO + partial/not_found now falls back to ReplyAll
    (flagged in caller) instead of passing blank To: to Outlook COM.

    Returns (html_body, overall_status, host_tracking_list).
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

    # FIX: max() with STATUS_PRIORITY (not min() with inverted dict)
    overall_status = max(statuses, key=lambda s: STATUS_PRIORITY.get(s, 0)) \
        if statuses else 'not_found'

    n_hosts  = len(hostname_qr_pairs)
    n_ok     = sum(1 for s in statuses if s == 'active')
    n_issues = n_hosts - n_ok

    BANNERS = {
        'active':    ('#166534', '#0d2e1a', '&#10004;&nbsp; All Hosts Confirmed Reporting on SIEM'),
        'partial':   ('#92400e', '#2d1f08',
                      f'&#9888;&nbsp; {n_issues} of {n_hosts} Host{"s" if n_hosts>1 else ""} Require Attention'),
        'not_found': ('#991b1b', '#2d1010',
                      f'&#10006;&nbsp; {"Some hosts not" if n_ok else "No hosts"} found in QRadar'),
    }
    bcol, bbg, blbl = BANNERS.get(overall_status, BANNERS['partial'])

    if overall_status == 'active' and n_hosts == 1:
        summary = f'<b>{hostname_qr_pairs[0][0]}</b> is confirmed reporting on our SIEM.'
    elif overall_status == 'active':
        names   = ', '.join(f'<b>{h}</b>' for h, _ in hostname_qr_pairs)
        summary = f'All {n_hosts} requested hosts ({names}) are confirmed reporting on our SIEM.'
    else:
        summary = (f'{n_ok} of {n_hosts} host{"s" if n_hosts>1 else ""} confirmed active. '
                   f'Issues are highlighted per host below.')

    body = (
        '<html><body style="font-family:\'Segoe UI\',Arial,sans-serif;color:#c8d4eb;'
        'background:#0b0f1a;font-size:13px;line-height:1.5;margin:0;padding:0;">'
        '<div style="max-width:700px;padding:20px;background:#111827;border-radius:8px;margin:16px">'
        '<p style="margin:0 0 16px 0;color:#e2e8f0">Hi,</p>'
        f'<div style="background:{bbg};border:1px solid {bcol};color:#f0f4ff;padding:10px 16px;'
        f'border-radius:6px;font-size:13px;font-weight:600;margin-bottom:12px">{blbl}</div>'
        f'<p style="margin:0 0 4px 0;color:#c8d4eb">{summary}</p>'
        + ''.join(host_sections) +
        f'<p style="margin:20px 0 4px 0;color:#5a6f90;font-size:12px">'
        f'Automated response from the SIEM monitoring system.<br>'
        f'Checked against QRadar on {run_time}.</p>'
        '<p style="margin:16px 0 0 0;color:#e2e8f0">Regards,<br>'
        '<span style="font-weight:600">Cyberdefence</span></p>'
        '</div></body></html>'
    )
    return body, overall_status, host_tracking


# ═══════════════════════════════════════════════════════════════════════════════
#  DRAFT CREATOR
# ═══════════════════════════════════════════════════════════════════════════════

def create_draft_reply(mail_item, html_body: str, overall_status: str,
                       hostnames_str: str, is_revalidation: bool = False) -> bool:
    """
    FIX: if ESCALATION_TO and ESCALATION_CC are BOTH empty for a
    partial/not_found, falls back to ReplyAll rather than setting blank To:
    (Outlook COM raises COMException / E_FAIL on SaveAs with empty To:).

    THIS IS DRAFT ONLY. reply.Save() called. reply.Send() NEVER called.
    """
    tag_map = {'active': TAG_ACTIVE, 'partial': TAG_PARTIAL, 'not_found': TAG_NOT_FOUND}
    tag    = tag_map.get(overall_status, TAG_ACTIVE)
    prefix = '[Revalidated] ' if is_revalidation else ''

    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"{prefix}{tag} {mail_item.Subject}"

        use_escalation = (
            overall_status in ('partial', 'not_found')
            and (ESCALATION_TO or ESCALATION_CC)       # FIX: only if non-empty
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
        main_inbox = ns.GetDefaultFolder(6)    # 6  = Inbox
        drafts     = ns.GetDefaultFolder(16)   # 16 = Drafts
        sent       = ns.GetDefaultFolder(5)    # 5  = Sent Items

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
#  DASHBOARD GENERATOR
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


def _build_dashboard_html(records: list) -> str:
    """
    Self-contained HTML dashboard.
    Data injected via %%PLACEHOLDER%% replacement (avoids f-string brace escaping).
    All JS uses literal { } braces safely inside the raw template string.
    """
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


_DASHBOARD_TMPL = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>SIEM Signoff Dashboard v%%VERSION%%</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg0:#09111f;--bg1:#0f1c2e;--bg2:#162234;--bg3:#1c2d42;
  --bdr:#223049;--bdr2:#19273b;
  --t0:#edf2ff;--t1:#c5d4ee;--t2:#8aa3c8;--t3:#536b8a;
  --green:#22c55e;--gd:#14532d;--gbg:#071a10;
  --amber:#f59e0b;--ad:#92400e;--abg:#1e1508;
  --red:#ef4444;--rd:#991b1b;--rbg:#1e0808;
  --blue:#3b82f6;--bd:#1e40af;--bbg:#0a1628;
  --purple:#a78bfa;--pbg:#130f28;
  --mono:'Consolas','Cascadia Code','Courier New',monospace;
  --r:8px;--r2:5px;
}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg0);color:var(--t1);font-size:13px;line-height:1.5;min-height:100vh}
.hdr{background:var(--bg1);border-bottom:1px solid var(--bdr);padding:12px 22px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;gap:10px;flex-wrap:wrap}
.hdr-l{display:flex;align-items:center;gap:9px}
.hdr h1{font-size:15px;font-weight:600;color:var(--t0)}
.badge{font-size:10px;padding:2px 7px;border-radius:4px;font-weight:600;white-space:nowrap}
.bqr{background:var(--pbg);color:var(--purple);border:1px solid #3d2a7a}
.bv{background:var(--bbg);color:var(--blue);border:1px solid var(--bd)}
.hdr-r{display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.hdr-ts{font-size:11px;color:var(--t3)}
.btn{background:var(--bg2);border:1px solid var(--bdr);color:var(--t2);padding:5px 12px;border-radius:var(--r2);cursor:pointer;font-size:12px;font-family:inherit;transition:all .12s;white-space:nowrap}
.btn:hover{border-color:var(--blue);color:var(--t0)}
.btn-sm{padding:3px 9px;font-size:11px}
.btn-pri{background:var(--bbg);border-color:var(--bd);color:var(--blue)}
.btn-pri:hover{background:var(--blue);color:#fff}
.btn-dan{background:var(--rbg);border-color:var(--rd);color:var(--red)}
.btn-dan:hover{background:var(--red);color:#fff}
.main{padding:18px 22px;max-width:1400px}
.pbar{display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap;align-items:center}
.pbar .sep{width:1px;height:18px;background:var(--bdr);margin:0 3px}
.pbtn{background:var(--bg2);border:1px solid var(--bdr);color:var(--t2);padding:5px 14px;border-radius:var(--r2);cursor:pointer;font-size:12px;font-family:inherit;transition:all .12s}
.pbtn:hover{border-color:var(--blue);color:var(--t0)}
.pbtn.active{background:var(--bbg);border-color:var(--bd);color:var(--blue);font-weight:500}
.cards{display:grid;grid-template-columns:repeat(5,1fr);gap:9px;margin-bottom:14px}
.card{background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);padding:13px 15px}
.card-lbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--t3);margin-bottom:7px}
.card-val{font-size:27px;font-weight:600;line-height:1;color:var(--t0)}
.card-sub{font-size:11px;color:var(--t3);margin-top:3px}
.card.cg .card-val{color:var(--green)}
.card.ca .card-val{color:var(--amber)}
.card.cr .card-val{color:var(--red)}
.card.cb .card-val{color:var(--blue)}
.card.cp .card-val{color:var(--purple)}
.charts{display:grid;grid-template-columns:230px 1fr;gap:10px;margin-bottom:14px}
.cbox{background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);padding:14px}
.clbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--t3);margin-bottom:11px}
.leg{display:flex;flex-direction:column;gap:6px;margin-top:11px}
.leg-r{display:flex;align-items:center;gap:7px;font-size:11px;color:var(--t2)}
.leg-d{width:8px;height:8px;border-radius:2px;flex-shrink:0}
.twrap{background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);overflow:hidden}
.ttbar{display:flex;align-items:center;gap:7px;padding:11px 15px;border-bottom:1px solid var(--bdr);flex-wrap:wrap}
.ttitle{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--t3);white-space:nowrap;margin-right:3px}
.fps{display:flex;gap:4px;flex-wrap:wrap}
.fp{background:var(--bg2);border:1px solid var(--bdr);color:var(--t2);padding:3px 10px;border-radius:20px;cursor:pointer;font-size:11px;font-family:inherit;transition:all .12s}
.fp:hover{border-color:var(--blue);color:var(--t0)}
.fp.active{background:var(--bbg);border-color:var(--bd);color:var(--blue);font-weight:500}
.fp.fa.active{background:var(--gbg);border-color:var(--gd);color:var(--green)}
.fp.fp2.active{background:var(--abg);border-color:var(--ad);color:var(--amber)}
.fp.fn.active{background:var(--rbg);border-color:var(--rd);color:var(--red)}
.srch{background:var(--bg0);border:1px solid var(--bdr);border-radius:var(--r2);color:var(--t0);padding:5px 10px;font-size:12px;width:185px;font-family:inherit}
.srch:focus{outline:none;border-color:var(--blue)}
.os-sel{background:var(--bg0);border:1px solid var(--bdr);border-radius:var(--r2);color:var(--t1);padding:5px 8px;font-size:12px;font-family:inherit}
.os-sel:focus{outline:none;border-color:var(--blue)}
.sp1{flex:1}
table{width:100%;border-collapse:collapse}
th{background:var(--bg0);color:var(--t3);font-size:10px;font-weight:600;text-align:left;padding:8px 12px;border-bottom:1px solid var(--bdr);text-transform:uppercase;letter-spacing:.5px;cursor:pointer;white-space:nowrap;user-select:none}
th:hover{color:var(--t0)}
th.sa::after{content:' ▲';color:var(--blue)}
th.sd::after{content:' ▼';color:var(--blue)}
td{padding:9px 12px;border-bottom:1px solid var(--bdr2);vertical-align:middle;font-size:12px}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--bg2)}
tr.drow td{opacity:.38;text-decoration:line-through}
.hn{font-family:var(--mono);font-size:12px;font-weight:500;color:var(--t0)}
.hndel{color:var(--t3)}
.sb{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;letter-spacing:.3px;white-space:nowrap}
.sa1{background:var(--gbg);color:var(--green);border:1px solid var(--gd)}
.sp3{background:var(--abg);color:var(--amber);border:1px solid var(--ad)}
.sn{background:var(--rbg);color:var(--red);border:1px solid var(--rd)}
.bovr{background:var(--abg);color:var(--amber);border:1px solid var(--ad);font-size:9px;padding:1px 5px;border-radius:3px;margin-left:4px;font-weight:600}
.bdel{background:var(--rbg);color:var(--red);border:1px solid var(--rd);font-size:9px;padding:1px 5px;border-radius:3px;margin-left:4px;font-weight:600}
.obadge{background:var(--bbg);color:var(--blue);font-size:10px;padding:2px 7px;border-radius:4px;border:1px solid var(--bd);white-space:nowrap}
.tup{color:var(--green);font-size:11px}
.tdn{color:var(--red);font-size:11px}
.teq{color:var(--t3);font-size:11px}
.rvsoon{color:var(--amber);font-weight:600}
.rvok{color:var(--t3)}
.ract{display:flex;gap:4px;opacity:0;transition:opacity .1s}
tr:hover .ract{opacity:1}
.nodata{text-align:center;padding:36px;color:var(--t3)}
.pager{display:flex;align-items:center;justify-content:flex-end;gap:7px;padding:9px 15px;border-top:1px solid var(--bdr);font-size:12px;color:var(--t2)}
.pginfo{font-size:11px;color:var(--t3)}
.mbg{position:fixed;inset:0;background:#00000090;backdrop-filter:blur(2px);z-index:500;display:none;align-items:center;justify-content:center}
.mbg.open{display:flex}
.modal{background:var(--bg1);border:1px solid var(--bdr);border-radius:var(--r);padding:22px;width:430px;max-width:94vw;max-height:88vh;overflow-y:auto}
.modal h2{font-size:14px;font-weight:600;color:var(--t0);margin-bottom:14px}
.frow{margin-bottom:13px}
label{display:block;font-size:10px;font-weight:600;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px}
select,textarea,input[type=text]{background:var(--bg0);border:1px solid var(--bdr);border-radius:var(--r2);color:var(--t0);padding:7px 9px;font-size:13px;font-family:inherit;width:100%}
select:focus,textarea:focus,input[type=text]:focus{outline:none;border-color:var(--blue)}
textarea{resize:vertical;min-height:65px}
.mact{display:flex;justify-content:flex-end;gap:7px;margin-top:16px}
.toast{position:fixed;bottom:18px;right:18px;background:var(--bg2);border:1px solid var(--bdr);border-radius:var(--r);padding:9px 14px;font-size:12px;z-index:600;box-shadow:0 4px 18px #000a;display:none}
.toast.show{display:block}
.ubar{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:var(--bg2);border:1px solid var(--rd);border-radius:var(--r);padding:10px 16px;font-size:12px;z-index:600;display:none;align-items:center;gap:10px;box-shadow:0 4px 18px #000a;white-space:nowrap}
.ubar.show{display:flex}
.ubtn{color:var(--amber);cursor:pointer;font-weight:600;text-decoration:underline}
.utmr{font-size:11px;color:var(--t3);font-family:var(--mono)}
.note-i{cursor:help;color:var(--amber);font-size:11px;margin-left:3px;vertical-align:middle}
.err{border-color:var(--red)!important}
@media(max-width:800px){.cards{grid-template-columns:repeat(2,1fr)}.charts{grid-template-columns:1fr}}
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
    <span class="hdr-ts" id="gents">Generated %%GENERATED%%</span>
    <button class="btn btn-sm" onclick="exportCSV()">&#8595; Export CSV</button>
    <button class="btn btn-sm" onclick="exportOvr()">&#8595; Overrides JSON</button>
  </div>
</div>

<div class="main">
  <div class="pbar">
    <button class="pbtn" onclick="setPeriod(7)"  data-p="7">7 days</button>
    <button class="pbtn" onclick="setPeriod(15)" data-p="15">15 days</button>
    <button class="pbtn active" onclick="setPeriod(30)" data-p="30">30 days</button>
    <button class="pbtn" onclick="setPeriod(0)"  data-p="0">All time</button>
    <div class="sep"></div>
    <span style="font-size:11px;color:var(--t3)">Lookback: %%LOOKBACK%%d &nbsp;|&nbsp; Reval cooldown: %%REVAL_COOL%%d</span>
  </div>

  <div class="cards">
    <div class="card cb"><div class="card-lbl">Signoff emails</div><div class="card-val" id="ct0">—</div><div class="card-sub" id="ct0s">—</div></div>
    <div class="card cg"><div class="card-lbl">Active</div><div class="card-val" id="ct1">—</div><div class="card-sub" id="ct1s">—</div></div>
    <div class="card ca"><div class="card-lbl">Partial</div><div class="card-val" id="ct2">—</div><div class="card-sub" id="ct2s">—</div></div>
    <div class="card cr"><div class="card-lbl">Not Found</div><div class="card-val" id="ct3">—</div><div class="card-sub" id="ct3s">—</div></div>
    <div class="card cp"><div class="card-lbl">Exceptions</div><div class="card-val" id="ct4">—</div><div class="card-sub">deleted / overridden</div></div>
  </div>

  <div class="charts">
    <div class="cbox">
      <div class="clbl">Host status distribution</div>
      <svg id="dsvg" viewBox="0 0 200 170" width="100%" role="img" aria-label="Host status donut chart"></svg>
      <div class="leg" id="dleg"></div>
    </div>
    <div class="cbox">
      <div class="clbl">Signoffs over time (stacked: not-found &rarr; partial &rarr; active)</div>
      <svg id="bsvg" viewBox="0 0 530 185" width="100%" role="img" aria-label="Daily signoff bar chart" style="overflow:visible"></svg>
    </div>
  </div>

  <div class="twrap">
    <div class="ttbar">
      <span class="ttitle">Host registry</span>
      <div class="fps">
        <button class="fp active" data-sf="all"       onclick="setSF('all')">All</button>
        <button class="fp fa"     data-sf="active"    onclick="setSF('active')">Active</button>
        <button class="fp fp2"   data-sf="partial"   onclick="setSF('partial')">Partial</button>
        <button class="fp fn"     data-sf="not_found" onclick="setSF('not_found')">Not Found</button>
        <button class="fp" data-sf="deleted" onclick="setSF('deleted')" style="border-color:var(--rd);color:var(--red)">Exceptions</button>
      </div>
      <select class="os-sel" onchange="setOS(this.value)" id="ossel">
        <option value="">All OS groups</option>
      </select>
      <div class="sp1"></div>
      <input class="srch" type="text" placeholder="Filter hostname…" oninput="setSrch(this.value)">
    </div>
    <div style="overflow-x:auto">
      <table>
        <thead><tr>
          <th onclick="srt('hostname')">Hostname</th>
          <th onclick="srt('os_group')">OS Group</th>
          <th onclick="srt('eff_status')">Status</th>
          <th onclick="srt('last_checked')">Last Checked</th>
          <th onclick="srt('days_ago')">Last QRadar Event</th>
          <th onclick="srt('checks')">Checks</th>
          <th onclick="srt('next_rv')">Next Revalidation</th>
          <th>Trend</th>
          <th style="width:80px"></th>
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
        <option value="">— no override (use script result) —</option>
        <option value="active">Active</option>
        <option value="partial">Partial</option>
        <option value="not_found">Not Found</option>
      </select>
    </div>
    <div class="frow">
      <label>Notes <span style="color:var(--t3);font-weight:400">(tooltip on dashboard)</span></label>
      <textarea id="mnote" placeholder="Freeform note…"></textarea>
    </div>
    <div class="frow">
      <label>Exception / exclusion reason <span style="color:var(--red)">*</span> <span style="color:var(--t3);font-weight:400">required to mark as exception</span></label>
      <textarea id="mreason" placeholder="e.g. Decommissioned — approved by SOC-Lead 2026-01-15…"></textarea>
    </div>
    <div class="mact">
      <button class="btn" onclick="closeM()">Cancel</button>
      <button class="btn btn-dan" onclick="delFromM()" id="mdelbtn">Mark as Exception</button>
      <button class="btn btn-pri" onclick="saveM()">Save changes</button>
    </div>
  </div>
</div>

<!-- Undo bar -->
<div class="ubar" id="ubar">
  <span id="umsg">Host marked as exception.</span>
  <span class="ubtn" onclick="undoDel()">Undo</span>
  <span class="utmr" id="utmr">10s</span>
</div>

<div class="toast" id="toast"></div>

<script>
// ── Injected data ────────────────────────────────────────────────────────────
const ALL = %%DATA_JSON%%;
const RVCOOL = %%REVAL_COOL%%;
const OVR_PATH = '%%OVPATH%%';

// ── localStorage keys ────────────────────────────────────────────────────────
const LS = 'siem_ovr_v3';
const PS = 25;   // page size

// ── Override helpers (localStorage — survives refresh & regen) ───────────────
function loadOvr()      { try { return JSON.parse(localStorage.getItem(LS)||'{}'); } catch { return {}; } }
function saveOvr(o)     { localStorage.setItem(LS, JSON.stringify(o)); }
function getO(hn)       { return loadOvr()[hn] || {}; }
function setO(hn, patch){ const a=loadOvr(); a[hn]=Object.assign(a[hn]||{}, patch, {ts:new Date().toISOString()}); saveOvr(a); }

// ── Status helpers ───────────────────────────────────────────────────────────
const SLBL  = {active:'Active',partial:'Partial',not_found:'Not Found'};
const SCLS  = {active:'sa1',partial:'sp3',not_found:'sn'};
const SRANK = {active:0,partial:1,not_found:2};  // higher = worse (matches STATUS_PRIORITY)

// ── State ────────────────────────────────────────────────────────────────────
let period=30, sf='all', osf='', srch='';
let scol='last_checked', sasc=false;
let cp=1, allH=[], editHN=null;
let undoTmr=null, undoPend=null, undoSecs=10;

// ── Build host map from records ──────────────────────────────────────────────
function buildMap(recs) {
  const m = {};
  recs.forEach(r => r.host_results.forEach(h => {
    const k = h.hostname;
    if (!m[k]) m[k] = {hostname:k,os_group:h.os_group||'—',status:h.status,
                       last_checked:r.timestamp,days_ago:h.days_ago,
                       checks:0,prev_status:null};
    const e = m[k];
    if (r.timestamp > e.last_checked) {
      e.prev_status=e.status; e.status=h.status; e.last_checked=r.timestamp;
      e.days_ago=h.days_ago; e.os_group=h.os_group||e.os_group||'—';
    }
    e.checks++;
  }));
  return Object.values(m);
}

function applyOvr(hosts) {
  const o = loadOvr();
  return hosts.map(h => {
    const x = o[h.hostname] || {};
    return {...h, deleted:!!x.deleted, status_override:x.status_override||null,
            note:x.note||'', exception_reason:x.exception_reason||'',
            eff_status:x.status_override||h.status};
  });
}

function filtRecs(d) {
  if (!d) return ALL;
  const c = new Date(); c.setDate(c.getDate()-d);
  return ALL.filter(r => new Date(r.timestamp)>=c);
}

// ── Period / filter controls ─────────────────────────────────────────────────
function setPeriod(d) {
  period=d;
  document.querySelectorAll('.pbtn').forEach(b=>b.classList.toggle('active',+b.dataset.p===d));
  cp=1; render();
}
function setSF(s) {
  sf=s;
  document.querySelectorAll('.fp').forEach(b=>b.classList.toggle('active',b.dataset.sf===s));
  cp=1; drawTable();
}
function setOS(v) { osf=v; cp=1; drawTable(); }
function setSrch(v) { srch=v.toLowerCase(); cp=1; drawTable(); }
function srt(col) {
  if (scol===col) sasc=!sasc; else { scol=col; sasc=(col==='hostname'); }
  document.querySelectorAll('th').forEach(t=>t.classList.remove('sa','sd'));
  const TH_COLS=['hostname','os_group','eff_status','last_checked','days_ago','checks','next_rv','_t','_a'];
  const idx=TH_COLS.indexOf(col);
  const ths=document.querySelectorAll('th');
  if(idx>=0 && ths[idx]) ths[idx].classList.add(sasc?'sa':'sd');
  drawTable();
}

// ── Helpers ──────────────────────────────────────────────────────────────────
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

// ── Main render ──────────────────────────────────────────────────────────────
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
  // OS dropdown
  const ogs=[...new Set(hosts.map(h=>h.os_group).filter(g=>g&&g!=='—'))].sort();
  const sel=document.getElementById('ossel'), cur=sel.value;
  sel.innerHTML='<option value="">All OS groups</option>'+ogs.map(g=>`<option value="${g}"${g===cur?' selected':''}>${g}</option>`).join('');
  renderDonut(a,p,n); renderBar(recs);
  allH=hosts; drawTable();
}

// ── Donut (pure SVG) ─────────────────────────────────────────────────────────
function renderDonut(a,p,n) {
  const tot=a+p+n, sv=document.getElementById('dsvg');
  if (!tot) { sv.innerHTML='<text x="100" y="90" text-anchor="middle" fill="#536b8a" font-size="12" font-family="Segoe UI,sans-serif">No data</text>'; document.getElementById('dleg').innerHTML=''; return; }
  const cx=100,cy=85,R=65,ri=46;
  const vs=[{v:a,c:'#22c55e'},{v:p,c:'#f59e0b'},{v:n,c:'#ef4444'}];
  let ang=-Math.PI/2, arcs='';
  vs.forEach(({v,c})=>{
    if(!v) return;
    const sw=2*Math.PI*(v/tot);
    const x1=cx+R*Math.cos(ang),y1=cy+R*Math.sin(ang); ang+=sw;
    const x2=cx+R*Math.cos(ang),y2=cy+R*Math.sin(ang);
    const xi1=cx+ri*Math.cos(ang-sw),yi1=cy+ri*Math.sin(ang-sw);
    const xi2=cx+ri*Math.cos(ang),yi2=cy+ri*Math.sin(ang);
    const lg=sw>Math.PI?1:0;
    arcs+=`<path d="M${x1},${y1} A${R},${R} 0 ${lg},1 ${x2},${y2} L${xi2},${yi2} A${ri},${ri} 0 ${lg},0 ${xi1},${yi1} Z" fill="${c}" opacity="0.85"/>`;
  });
  const pct=Math.round(a/tot*100);
  sv.innerHTML=arcs+`<text x="${cx}" y="${cy-5}" text-anchor="middle" font-size="21" font-weight="600" fill="#edf2ff" font-family="Segoe UI,sans-serif">${a}</text><text x="${cx}" y="${cy+13}" text-anchor="middle" font-size="10" fill="#8aa3c8" font-family="Segoe UI,sans-serif">active (${pct}%)</text>`;
  document.getElementById('dleg').innerHTML=[['#22c55e','Active',a],['#f59e0b','Partial',p],['#ef4444','Not Found',n]].map(([c,l,v])=>`<div class="leg-r"><div class="leg-d" style="background:${c}"></div><span style="flex:1">${l}</span><span style="color:#edf2ff;font-weight:600">${v}</span></div>`).join('');
}

// ── Bar chart (pure SVG) — stacking: not_found bottom, partial mid, active top ─
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
    grid+=`<line x1="${PL}" y1="${y}" x2="${W-PR}" y2="${y}" stroke="#19273b" stroke-width="0.5"/>`;
    grid+=`<text x="${PL-3}" y="${y+3}" text-anchor="end" font-size="8" fill="#536b8a" font-family="Segoe UI,sans-serif">${Math.round(mv*g/4)}</text>`;
  }
  keys.forEach((k,i)=>{
    const {a,p,n}=bkts[k], x=PL+i*gap;
    let y=H-PB;
    // FIX: draw not_found at bottom, partial above, active topmost (most visible = good)
    if(n){const h=Math.max(1,Math.round(yS(n)));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="#ef444450" stroke="#ef4444" stroke-width="0.5" rx="1"/>`;y-=h;}
    if(p){const h=Math.max(1,Math.round(yS(p)));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="#f59e0b50" stroke="#f59e0b" stroke-width="0.5" rx="1"/>`;y-=h;}
    if(a){const h=Math.max(1,Math.round(yS(a)));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="#22c55e50" stroke="#22c55e" stroke-width="0.5" rx="1"/>`;}
    const sk=keys.length>20?Math.ceil(keys.length/12):1;
    if(i%sk===0) labs+=`<text x="${x+bw/2}" y="${H-PB+11}" text-anchor="middle" font-size="8" fill="#536b8a" font-family="Segoe UI,sans-serif">${k.slice(5)}</text>`;
  });
  // Legend
  const ly=H-4;
  const leg=`<rect x="${PL}" y="${ly-5}" width="7" height="5" fill="#22c55e50" stroke="#22c55e" stroke-width=".5"/><text x="${PL+9}" y="${ly}" font-size="8" fill="#8aa3c8" font-family="Segoe UI,sans-serif">Active</text><rect x="${PL+52}" y="${ly-5}" width="7" height="5" fill="#f59e0b50" stroke="#f59e0b" stroke-width=".5"/><text x="${PL+61}" y="${ly}" font-size="8" fill="#8aa3c8" font-family="Segoe UI,sans-serif">Partial</text><rect x="${PL+110}" y="${ly-5}" width="7" height="5" fill="#ef444450" stroke="#ef4444" stroke-width=".5"/><text x="${PL+119}" y="${ly}" font-size="8" fill="#8aa3c8" font-family="Segoe UI,sans-serif">Not Found</text>`;
  document.getElementById('bsvg').innerHTML=grid+bars+labs+leg;
}

// ── Table ────────────────────────────────────────────────────────────────────
function drawTable() {
  let rows=allH.filter(h=>{
    if(sf==='deleted') return h.deleted;
    if(sf!=='all' && h.eff_status!==sf) return false;
    if(h.deleted && sf!=='all') return false;
    if(osf && h.os_group!==osf) return false;
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
  const sl=rows.slice((cp-1)*PS, cp*PS);
  const tb=document.getElementById('tbody');
  if(!sl.length){tb.innerHTML='<tr><td colspan="9" class="nodata">No hosts match the current filters.</td></tr>';document.getElementById('pager').innerHTML='';return;}
  tb.innerHTML=sl.map(h=>{
    const da=h.days_ago!=null?(h.days_ago===0?'Today':`${h.days_ago}d ago`):'—';
    const dac=h.days_ago!=null&&h.days_ago>7?'color:var(--red)':'color:var(--t2)';
    const og=h.os_group&&h.os_group!=='—'?`<span class="obadge">${h.os_group}</span>`:'<span style="color:var(--t3)">—</span>';
    const ob=h.status_override?'<span class="bovr">overridden</span>':'';
    const db=h.deleted?'<span class="bdel">exception</span>':'';
    const ni=h.note?`<span class="note-i" title="${h.note.replace(/"/g,'&quot;')}">&#9432;</span>`:'';
    const hn=`<span class="hn${h.deleted?' hndel':''}">${h.hostname}</span>${ni}${ob}${db}`;
    const esc=h.hostname.replace(/'/g,"\\'");
    return `<tr${h.deleted?' class="drow"':''}><td>${hn}</td><td>${og}</td><td><span class="sb ${SCLS[h.eff_status]||'sp3'}">${SLBL[h.eff_status]||h.eff_status}</span></td><td style="color:var(--t2)">${rel(h.last_checked)}</td><td style="${dac}">${da}</td><td style="color:var(--t3);text-align:center">${h.checks}</td><td>${nextRV(h)}</td><td>${trend(h)}</td><td><div class="ract"><button class="btn btn-sm" onclick="openM('${esc}')">&#9998;</button></div></td></tr>`;
  }).join('');
  // Pager
  const pg=document.getElementById('pager');
  if(pgs<=1){pg.innerHTML=`<span class="pginfo">Showing ${tot} host${tot!==1?'s':''}</span>`;return;}
  let bs='';
  for(let i=1;i<=pgs;i++) bs+=`<button class="btn btn-sm${i===cp?' btn-pri':''}" onclick="goP(${i})">${i}</button>`;
  pg.innerHTML=`<span class="pginfo">Showing ${(cp-1)*PS+1}–${Math.min(cp*PS,tot)} of ${tot}</span>${bs}`;
}
function goP(n){cp=n;drawTable();}

// ── Modal ────────────────────────────────────────────────────────────────────
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
  setO(editHN,{status_override:document.getElementById('mst').value||null, note:document.getElementById('mnote').value.trim(), exception_reason:document.getElementById('mreason').value.trim(), deleted:getO(editHN).deleted||false});
  closeM(); render(); toast('Changes saved for '+editHN);
}
function delFromM() {
  if(!editHN) return;
  const r=document.getElementById('mreason').value.trim();
  if(!r){document.getElementById('mreason').classList.add('err');return;}
  const hn=editHN;
  setO(hn,{deleted:true, exception_reason:r, note:document.getElementById('mnote').value.trim(), status_override:document.getElementById('mst').value||null});
  closeM(); undoPend=hn; showUndo(hn); render();
}

// ── Undo delete ──────────────────────────────────────────────────────────────
function showUndo(hn) {
  clearInterval(undoTmr); undoSecs=10;
  document.getElementById('umsg').textContent=`"${hn}" marked as exception.`;
  document.getElementById('utmr').textContent=undoSecs+'s';
  document.getElementById('ubar').classList.add('show');
  undoTmr=setInterval(()=>{undoSecs--;document.getElementById('utmr').textContent=undoSecs+'s';if(undoSecs<=0){clearInterval(undoTmr);document.getElementById('ubar').classList.remove('show');undoPend=null;}},1000);
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
  ttmr=setTimeout(()=>el.classList.remove('show'),2600);
}

// ── Export CSV ───────────────────────────────────────────────────────────────
function exportCSV() {
  const hosts=applyOvr(buildMap(filtRecs(period)));
  const rows=[['Hostname','OS Group','Status','Status Override','Last Checked','Days Since Event','Checks','Exception','Note','Exception Reason']];
  hosts.forEach(h=>rows.push([h.hostname,h.os_group||'',SLBL[h.eff_status]||h.eff_status,h.status_override?SLBL[h.status_override]:'',h.last_checked,h.days_ago??'',h.checks,h.deleted?'YES':'',h.note||'',h.exception_reason||'']));
  const csv=rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\r\n');
  dl(new Blob([csv],{type:'text/csv'}),'signoff_export.csv');
}

// ── Export Overrides JSON (place at _OVERRIDES_PATH for script to read) ──────
function exportOvr() {
  const o=loadOvr();
  dl(new Blob([JSON.stringify(o,null,2)],{type:'application/json'}),'signoff_overrides.json');
  toast('Save this file to: '+OVR_PATH);
}

function dl(blob,name){const u=URL.createObjectURL(blob);const a=Object.assign(document.createElement('a'),{href:u,download:name});a.click();URL.revokeObjectURL(u);}

// ── Keyboard shortcuts ────────────────────────────────────────────────────────
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
        # ── Load overrides ─────────────────────────────────────────────────
        overrides      = load_overrides()
        deleted_hosts  = {hn for hn, o in overrides.items() if o.get('deleted')}
        if deleted_hosts:
            _log(f"Exception hosts (skipped on this run): {sorted(deleted_hosts)}")

        # ── Outlook ────────────────────────────────────────────────────────
        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        # ── QRadar ─────────────────────────────────────────────────────────
        if not test_qradar_connection():
            _log("ERROR: QRadar unreachable — exiting. All emails left untouched.")
            return
        fetch_log_source_types()

        # ── Inbox scan ─────────────────────────────────────────────────────
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
                if mail_item.Class != 43:   # 43 = olMail
                    continue
            except Exception:
                continue

            subject = ''
            try:
                subject = mail_item.Subject or ''
            except Exception:
                continue

            # ── Subject guards ──────────────────────────────────────────
            passed, reason = passes_subject_guards(subject)
            if not passed:
                skipped += 1
                _log(f"   SKIP (subject — {reason}): '{subject[:60]}'")
                continue

            # ── Sender guards ───────────────────────────────────────────
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

            # ── Body DL check ───────────────────────────────────────────
            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"   SKIP ('{TRIGGER_DL}' not in body): '{subject[:60]}'")
                continue

            # ── Hostname extraction ─────────────────────────────────────
            hostnames = extract_hostnames(subject)
            if not hostnames:
                skipped += 1
                _log(f"   SKIP (no hostnames): '{subject[:60]}'")
                continue

            # ── Exception host filter ───────────────────────────────────
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

            # ── Layer 1: runtime dedup (same execution) ─────────────────
            if hn_fset in _runtime_drafted_hosts:
                skipped += 1
                _log(f"   SKIP (already drafted in this run): {hostnames}")
                continue

            # ── Layer 2: data-file Active skip ──────────────────────────
            was_active, active_date = _was_active_recently(hn_fset)
            if was_active:
                skipped += 1
                _log(f"   SKIP (Active in data file on {active_date} "
                     f"within {ACTIVE_SKIP_DAYS}d window)")
                continue

            received_dt = _com_dt_to_py(
                getattr(mail_item, 'ReceivedTime', None)
            )
            _log(f"\nCandidate: '{subject[:70]}'")
            _log(f"   Sender : {sender}")
            _log(f"   Hosts  : {hostnames}")

            # ── Layer 3: Outlook Sent/Drafts tag scan ───────────────────
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
                # REVALIDATION_COOLDOWN_DAYS = 0 means every run re-checks
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

            # ── QRadar query per hostname ───────────────────────────────
            _log(f"   Querying QRadar for {len(hostnames)} host(s)...")
            hostname_qr_pairs = []
            for hn in hostnames:
                qr = query_all_log_sources_readonly(hn)
                _log(f"      {hn}: {qr['status']} ({len(qr.get('sources',[]))} sources)")
                hostname_qr_pairs.append((hn, qr))

            # ── Build reply ─────────────────────────────────────────────
            body, overall_status, host_tracking = build_reply_for_all_hosts(
                hostname_qr_pairs
            )
            _log(f"   Overall: {overall_status.upper()}")

            # ── Create draft ────────────────────────────────────────────
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
                    'id':               str(uuid.uuid4()),
                    'version':          VERSION,
                    'timestamp':        datetime.now().isoformat(timespec='seconds'),
                    'subject':          subject,
                    'sender':           sender,
                    'is_revalidation':  is_revalidation,
                    'overall_status':   overall_status,
                    'host_results':     host_tracking,
                })

            processed += 1

        # ── Dashboard ──────────────────────────────────────────────────────
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
