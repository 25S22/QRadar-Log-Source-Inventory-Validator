"""
QRadar Signoff Runner  v3.2
────────────────────────────────────────────────────────────────────────────
Scans Outlook for SIEM signoff emails, queries QRadar, and replies with the
result — either as a draft for review, or sent immediately when configured.

Key behaviours
  • First-email-only policy — any subject starting with RE: / FW: / FWD: is
    skipped immediately.  No prefix-stripping, no chain chasing.
  • Tolerant subject parsing (v3.2) — the keyword is matched regardless of
    spacing/hyphenation ("Sign-off", "Sign off"), extra words between the
    keyword and the hostnames are ignored ("Security Signoff for | ..."),
    hostnames may be separated by | , ; / or a spaced dash, and environment
    tags such as "OnPrem" / "Azure" / "Prod" are dropped before lookup.
    Run `python signoff_runner.py --selftest` to see the parser in action.
  • Draft-by-default — reply.Save() unless AUTO_SEND_ON_ACTIVE=True, in which
    case a fully-Active result (ALL hosts confirmed) is sent immediately.
    Partial / Not-Found results are ALWAYS drafted for human review and
    escalation — this is never overridden by AUTO_SEND_ON_ACTIVE.
  • Atomic JSON writes — data file is never left in a corrupt half-written state.
  • Single-instance lock — a lockfile prevents concurrent runs.
  • Runtime dedup — the same hostname set is never drafted/sent twice per run.
  • Conversation dedup — Sent + Drafts folders are scanned for prior outcomes.

Dashboard
  Run signoff_dashboard.py separately to view / edit results.
────────────────────────────────────────────────────────────────────────────
"""

import json
import os
import re
import sys
import tempfile
import time
import uuid
import urllib3
import win32com.client
import pywintypes
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

# ─── Subject parsing tolerance  (NEW in v3.2 — purely additive) ───────────────
# Extra characters that ALSO separate hostnames, for senders who don't use '|'.
# A plain '-' is handled separately: it only splits when surrounded by spaces
# (" SRV1 - SRV2 "), so embedded dashes in names like PRD-SQL-01 stay intact.
EXTRA_HOST_SEPARATORS = [',', ';', '/', '\n', '\r']

# Tokens/words dropped before the QRadar lookup: environment and platform tags
# plus the usual subject-line filler.  Matching ignores case, spaces, dots and
# dashes — so one entry 'onprem' covers "OnPrem", "On-Prem" and "on prem".
# A word is only dropped on an EXACT match, so 'AZURE01' survives 'azure'.
SUBJECT_NOISE_WORDS = [
    # placement / platform
    'onprem', 'onpremise', 'onpremises', 'premise', 'premises', 'cloud',
    'azure', 'aws', 'gcp', 'oci', 'vmware', 'vm', 'physical', 'virtual',
    # environment
    'prod', 'production', 'nonprod', 'preprod', 'dev', 'development',
    'test', 'testing', 'uat', 'sit', 'qa', 'stage', 'staging', 'dr', 'poc',
    # operating system
    'windows', 'win', 'linux', 'unix', 'rhel', 'ubuntu', 'centos', 'aix',
    # generic nouns
    'server', 'servers', 'host', 'hosts', 'hostname', 'hostnames', 'name',
    'names', 'machine', 'machines', 'asset', 'assets', 'device', 'devices',
    'node', 'nodes', 'box', 'boxes', 'ip', 'fqdn',
    # request filler
    'new', 'newly', 'onboard', 'onboarded', 'onboarding', 'request',
    'requests', 'req', 'required', 'please', 'kindly', 'thanks', 'thank',
    'regards', 'urgent', 'asap', 'fyi', 'eom', 'external', 'internal',
    'siem', 'qradar', 'security', 'signoff', 'signoffs', 'sign', 'off',
    'confirmation', 'confirm', 'validation', 'validate', 'verification',
    'verify', 'check', 'status', 'update', 'and', 'for', 'of', 'the',
    'on', 'in', 'to', 'from', 'with', 'tbd', 'na', 'nil', 'xxx', 'xxxx',
]

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

# ─── Send behaviour ─────────────────────────────────────────────────────────────
# When True: a fully-Active result (every host in the subject line confirmed
# reporting) is sent immediately via reply.Send() instead of being saved as
# a draft.  Partial / Not-Found results are ALWAYS drafted for escalation —
# this flag has no effect on them, by design.
AUTO_SEND_ON_ACTIVE = False

# Safety net for rollout: when True, an Active result that WOULD be sent is
# instead logged (so you can see exactly what would have gone out) and still
# saved as a draft.  Run with this on for a few cycles before flipping it off.
AUTO_SEND_DRY_RUN = True

# ─── Outlook ──────────────────────────────────────────────────────────────────
# Sub-folder of Inbox to scan.  Set to None to scan the full Inbox.
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'

# Named MAPI profile to log on with. Leave '' to use whatever profile is
# already loaded (the normal case when Outlook is already open under your
# own session). Only set this for scheduled-task setups where the profile
# needs to be selected explicitly.
OUTLOOK_PROFILE_NAME = os.environ.get('OUTLOOK_PROFILE_NAME', '')

# How many times to retry connecting to Outlook before giving up, and how
# long to wait between attempts. Useful if this script is launched at the
# same time Outlook itself is starting (e.g. both at login).
OUTLOOK_CONNECT_RETRIES     = 3
OUTLOOK_CONNECT_RETRY_DELAY = 5   # seconds

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

# Tokens thrown away by the most recent extract_hostnames() call, purely so the
# run log can show WHY something was ignored. Reset on every call.
_last_parse_dropped: list = []


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


def _items_since(items_collection, cutoff_dt: datetime, sort_property: str, get_dt) -> list:
    """
    Return every item in `items_collection` dated on/after `cutoff_dt`, WITHOUT
    using Outlook's string-based Restrict()/Find() for the date comparison.

    ROOT CAUSE OF "0 emails found" BUGS: Restrict() filters like
    "[ReceivedTime] >= '09/01/2026 12:00 AM'" are parsed by Outlook using the
    *Windows regional short-date format* of the account the script runs
    under — not necessarily MM/DD/YYYY. If that machine's Region settings use
    DD/MM/YYYY, YYYY-MM-DD, or anything else, the filter string is silently
    misread and matches nothing — no error, no exception, just an empty
    result set. This is a long-standing, well-documented Outlook COM quirk.

    Fix: sort the collection by the same field (descending) and compare real
    Python datetime objects instead of building a locale-dependent string.
    Sorting first lets us break out early once we pass the cutoff, so this
    stays fast even on large folders. If Sort() itself fails for any reason,
    we fall back to scanning every item (slower, but still correct) rather
    than risking a false "0 found" ever showing up again.
    """
    sorted_ok = True
    try:
        items_collection.Sort(f"[{sort_property}]", True)  # True = descending, newest first
    except Exception as exc:
        sorted_ok = False
        _log(f"      WARNING: Sort by [{sort_property}] failed ({exc}) — "
             f"scanning every item instead (slower, still correct).")

    matched = []
    for item in items_collection:
        try:
            dt = get_dt(item)
        except Exception:
            continue
        if dt is None:
            continue
        if dt < cutoff_dt:
            if sorted_ok:
                break   # collection is sorted newest-first — nothing later can qualify
            continue    # unsorted fallback — keep scanning the rest
        matched.append(item)
    return matched


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


def _qradar_get_with_retry(path: str, params: dict = None, headers: dict = None,
                           max_retries: int = 3) -> requests.Response | None:
    """
    GET with exponential backoff on network errors, 429 (rate limit), and 5xx.
    Returns None if every attempt fails — callers must handle that case.
    """
    headers = headers or {'Accept': 'application/json', 'Version': '14.0'}
    backoff = 1.0
    for attempt in range(1, max_retries + 1):
        try:
            r = requests.get(
                f"{QRADAR_HOST.rstrip('/')}{path}",
                params=params,
                auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
                verify=VERIFY_SSL,
                timeout=REQUEST_TIMEOUT,
                headers=headers,
            )
        except requests.exceptions.RequestException as exc:
            _log(f"      WARNING: QRadar request error (attempt {attempt}/{max_retries}): {exc}")
            if attempt == max_retries:
                return None
            time.sleep(backoff)
            backoff *= 2
            continue

        if r.status_code == 429:
            retry_after = float(r.headers.get('Retry-After', backoff))
            _log(f"      WARNING: QRadar rate-limited (429) — retrying in {retry_after:.0f}s")
            time.sleep(retry_after)
            backoff *= 2
            continue

        if r.status_code >= 500 and attempt < max_retries:
            _log(f"      WARNING: QRadar HTTP {r.status_code} — retrying in {backoff:.0f}s")
            time.sleep(backoff)
            backoff *= 2
            continue

        return r
    return None


def _qradar_get_all(path: str, params: dict = None, page_size: int = 100) -> list:
    """
    Fetch a full QRadar list endpoint using Range-header pagination.
    QRadar list endpoints cap results per request (commonly 50-1000 depending
    on config) — without this, a broad hostname substring match could
    silently return only a partial result set.
    """
    results = []
    start   = 0
    while True:
        headers = {
            'Accept': 'application/json',
            'Version': '14.0',
            'Range': f'items={start}-{start + page_size - 1}',
        }
        r = _qradar_get_with_retry(path, params=params, headers=headers)
        if r is None or r.status_code not in (200, 206):
            if r is not None:
                _log(f"      WARNING: QRadar HTTP {r.status_code} paginating {path}")
            break
        chunk = r.json()
        if not chunk:
            break
        results.extend(chunk)
        if len(chunk) < page_size:
            break
        start += page_size
    return results


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
        raw = _qradar_get_all(
            '/api/config/event_sources/log_source_management/log_sources',
            params={'filter': f'name ilike "%{clean}%"'},
        )
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

def _get_sender_smtp(mail_item) -> str:
    """
    Resolve the sender's real SMTP address.

    For on-prem/hybrid Exchange accounts, mail_item.SenderEmailAddress often
    returns an X.500 legacyExchangeDN (e.g. "/O=ORG/OU=.../cn=recipients/...")
    instead of an smtp address. That silently breaks ALLOWED_SENDERS and
    self-sent domain matching, since neither ever equals an '@domain' string.
    """
    try:
        if getattr(mail_item, 'SenderEmailType', '') == 'EX':
            smtp = mail_item.Sender.GetExchangeUser().PrimarySmtpAddress
            if smtp:
                return smtp
    except Exception:
        pass
    return mail_item.SenderEmailAddress or ''


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


# ══════════════════════════════════════════════════════════════════════════════
# SUBJECT PARSING   (rebuilt in v3.2)
# ══════════════════════════════════════════════════════════════════════════════
#
# Real-world subject lines that must all resolve to the same hostname list:
#
#   Security Signoff | SRV01
#   Security Signoff for | SRV01
#   Security Sign-off Request | OnPrem | SRV01
#   Security Signoff | SRV01, SRV02; SRV03
#   Security Signoff - SRV01 - SRV02            (spaced dash = separator)
#   Security Signoff | PRD-SQL-01               (embedded dash = NOT a separator)
#   Security Signoff for SRV01 and SRV02        (no separator at all)
#
# Strategy:
#   1. Locate the keyword by comparing *alphanumerics only*, so spacing and
#      hyphenation ("Sign-off", "Sign off", "SIGNOFF") never matter.
#   2. Everything to the right of the keyword is the payload.
#   3. Split the payload on | , ; / newline and spaced/en/em dashes.
#   4. Drop noise tokens (OnPrem, Azure, Prod, "for", "please", ...).
#   5. Validate what's left looks like a hostname, de-duplicate, return.
#
# Safety rail: when the payload contained NO explicit separator, a multi-token
# result is narrowed to tokens that actually look like hostnames (they contain
# a digit, dot, or dash).  That stops free-text subjects such as
# "Security Signoff for the new servers" turning prose into QRadar queries.
# ──────────────────────────────────────────────────────────────────────────────

# Reply / forward prefixes, including numbered ones (RE[2]:) and the common
# non-English equivalents that Exchange adds for other Outlook languages.
_RE_REPLY_PREFIX = re.compile(
    r'^\s*(re|fw|fwd|rv|aw|wg|tr|sv|vs|res|enc|antw|antwort)\s*(\[\d+\])?\s*:',
    re.IGNORECASE,
)

# A dash only separates hosts when it is spaced, or when it is an en/em dash.
_RE_DASH_SEPARATOR = re.compile(r'\s+[-]+\s+|\s*[\u2013\u2014]+\s*')

# Characters trimmed from the edge of a token: quotes, brackets, punctuation.
_TOKEN_TRIM_CHARS = ' \t\'"`()[]{}<>.,;:!?*_'

# A hostname may contain letters, digits, dots, dashes, underscores — and a
# space only in the multi-word fallback path below.
_RE_HOSTNAME_SHAPE = re.compile(r'^[A-Za-z0-9][A-Za-z0-9._\- ]*$')

# Some teams submit the IP rather than the name — QRadar log source names very
# often contain it, so accept it as a lookup target.
_RE_IPV4 = re.compile(r'^\d{1,3}(?:\.\d{1,3}){3}$')

_SPLIT_SENTINEL = '\x00'


def _norm_word(text: str) -> str:
    """Lowercase and strip everything that isn't a letter or digit."""
    return re.sub(r'[^a-z0-9]', '', (text or '').lower())


_NOISE_SET = {_norm_word(w) for w in SUBJECT_NOISE_WORDS if _norm_word(w)}


def _squash(text: str) -> tuple:
    """
    Return (alphanumerics-only lowercase string, index map back to `text`).
    Lets us find 'securitysignoff' inside 'Security  Sign-Off' and still know
    exactly where the match ends in the original string.
    """
    squashed, index_map = [], []
    for i, ch in enumerate(text or ''):
        if ch.isalnum():
            squashed.append(ch.lower())
            index_map.append(i)
    return ''.join(squashed), index_map


def _keyword_span(subject: str):
    """Return (start, end) offsets of SUBJECT_KEYWORD in `subject`, or None."""
    key = _norm_word(SUBJECT_KEYWORD)
    if not key:
        return (0, 0)
    squashed, index_map = _squash(subject)
    pos = squashed.find(key)
    if pos < 0:
        return None
    return index_map[pos], index_map[pos + len(key) - 1] + 1


def _clean_token(token: str) -> str:
    return (token or '').strip().strip(_TOKEN_TRIM_CHARS).strip()


def _is_noise(token: str) -> bool:
    n = _norm_word(token)
    return (not n) or n in _NOISE_SET


def _is_valid_hostname(token: str) -> bool:
    """Loose sanity check — is this plausibly a machine name?"""
    if not token or len(token) > 120:
        return False
    if '@' in token:                                  # an email address, not a host
        return False
    if _RE_IPV4.match(token):
        return all(0 <= int(o) <= 255 for o in token.split('.'))
    if not _RE_HOSTNAME_SHAPE.match(token):
        return False
    if not any(c.isalpha() for c in token):           # kills dates and bare numbers
        return False
    if len(_norm_word(token)) < 3:                    # kills stray initials
        return False
    return True


def _is_strong_hostname(token: str) -> bool:
    """Stricter check used where the sender gave us no explicit separator."""
    if not token or ' ' in token:
        return False
    return any(c.isdigit() for c in token) or '.' in token or '-' in token


def _split_payload(payload: str) -> tuple:
    """Return (raw_tokens, had_explicit_separator)."""
    text = payload or ''
    had_sep = False

    for sep in [SUBJECT_SEPARATOR] + list(EXTRA_HOST_SEPARATORS):
        if sep and sep in text:
            had_sep = True
            text = text.replace(sep, _SPLIT_SENTINEL)

    if _RE_DASH_SEPARATOR.search(text):
        had_sep = True
        text = _RE_DASH_SEPARATOR.sub(_SPLIT_SENTINEL, text)

    return [t.strip() for t in text.split(_SPLIT_SENTINEL)], had_sep


def _tokens_from_multiword(token: str) -> list:
    """
    Handle a token that still contains spaces, e.g. "OnPrem SRV01" or
    "Prod Windows SRV01".  Noise words are dropped; if any survivor looks
    like a real hostname we keep only those, otherwise we keep the whole
    remaining phrase (some estates really do use spaced names).
    """
    words = [_clean_token(w) for w in token.split()]
    kept  = [w for w in words if w and not _is_noise(w)]
    if not kept:
        return []
    strong = [w for w in kept if _is_strong_hostname(w)]
    if strong:
        return strong
    return [' '.join(kept)]


def extract_hostnames(subject: str) -> list:
    """
    Return the list of hostnames parsed from `subject`, in the order given,
    de-duplicated case-insensitively, with environment/filler tags removed.
    Tokens that were thrown away are recorded in `_last_parse_dropped` so the
    run log can explain itself.
    """
    global _last_parse_dropped
    _last_parse_dropped = []

    span = _keyword_span(subject or '')
    if span is None:
        return []
    payload = (subject or '')[span[1]:]

    raw_tokens, had_sep = _split_payload(payload)

    candidates = []
    for raw in raw_tokens:
        token = _clean_token(raw)
        if not token:
            continue
        if _is_noise(token):
            _last_parse_dropped.append(token)
            continue
        if ' ' in token:
            pieces = _tokens_from_multiword(token)
            if not pieces:
                _last_parse_dropped.append(token)
            candidates.extend(pieces)
        else:
            candidates.append(token)

    hostnames, seen = [], set()
    for cand in candidates:
        cand = _clean_token(cand)
        if not cand or _is_noise(cand):
            if cand:
                _last_parse_dropped.append(cand)
            continue
        if not _is_valid_hostname(cand):
            _last_parse_dropped.append(cand)
            continue
        key = cand.upper()
        if key in seen:
            continue
        seen.add(key)
        hostnames.append(cand)

    # No explicit separator anywhere — be conservative with free text.
    if not had_sep and len(hostnames) > 1:
        strong = [h for h in hostnames if _is_strong_hostname(h)]
        _last_parse_dropped.extend(h for h in hostnames if h not in strong)
        hostnames = strong

    return hostnames


def passes_subject_guards(subject: str) -> tuple:
    """
    Returns (True, 'ok') or (False, reason_string).

    Rules applied in order:
      1. RE: / FW: / FWD: prefix → skip  (first-email-only policy)
      2. Already carries a [Processed*] tag → skip
      3. Keyword not found (spacing/hyphenation-insensitive) → skip
      4. Keyword sits to the RIGHT of the first separator → skip
      5. Nothing at all after the keyword → skip

    Note: unlike v3.1 the separator is no longer mandatory — a subject such as
    "Security Signoff for SRV01" is accepted, because extract_hostnames() can
    resolve it safely.  Emails whose payload yields no hostname are still
    skipped by the caller.
    """
    if not subject or not subject.strip():
        return False, "empty subject"

    s     = subject.strip()
    lower = s.lower()

    # Rule 1 — first-email-only: reject any reply or forward
    m = _RE_REPLY_PREFIX.match(s)
    if m:
        return False, f"reply/forward ({m.group(1).lower()})"

    # Rule 2 — skip already-processed emails
    if '[processed' in lower:
        return False, "already tagged"

    # Rule 3 — keyword must be present somewhere
    span = _keyword_span(s)
    if span is None:
        return False, f"keyword '{SUBJECT_KEYWORD}' not found"

    # Rule 4 — keyword must sit left of the primary separator when one is used
    if SUBJECT_SEPARATOR and SUBJECT_SEPARATOR in s:
        if span[0] > s.index(SUBJECT_SEPARATOR):
            return False, f"keyword '{SUBJECT_KEYWORD}' not left of '{SUBJECT_SEPARATOR}'"

    # Rule 5 — something usable must follow the keyword. "Security Signoff for |"
    # and "Security Signoff for the new servers" both fail here: every word left
    # over is filler, so there is nothing to look up.
    tail = s[span[1]:]
    for sep in [SUBJECT_SEPARATOR] + list(EXTRA_HOST_SEPARATORS):
        if sep:
            tail = tail.replace(sep, ' ')
    tail_words = [w for w in (_clean_token(w) for w in tail.split()) if w]
    if not tail_words:
        return False, "nothing after keyword"
    if all(_is_noise(w) for w in tail_words):
        return False, f"nothing usable after keyword ({' '.join(tail_words)!r} is all filler)"

    return True, "ok"


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

    cutoff = datetime.now() - timedelta(days=SENT_SCAN_DAYS)

    # Sent Items
    try:
        sent_candidates = _items_since(sent_folder.Items, cutoff, 'SentOn',
                                       lambda it: _com_dt_to_py(it.SentOn))
        for item in sent_candidates:
            try:
                if item.ConversationID == conv_id:
                    _update(_tag_from_subject(item.Subject), _com_dt_to_py(item.SentOn))
            except Exception:
                continue
    except Exception as exc:
        _log(f"      WARNING: Sent scan error: {exc}")

    # Drafts (same date window to keep large mailboxes fast)
    try:
        draft_candidates = _items_since(drafts_folder.Items, cutoff, 'LastModificationTime',
                                        lambda it: _com_dt_to_py(it.LastModificationTime))
        for item in draft_candidates:
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
                 prior_status: str | None, was_sent: bool = False,
                 ignored_tokens: list = None) -> None:
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
        'was_sent':          was_sent,   # True = auto-sent, False = saved as draft
        'hosts':             host_records,
        'ignored_tokens':    ignored_tokens or [],   # v3.2 — noise dropped from the subject
        'manually_resolved': False,
        'notes':             '',
    })
    try:
        _atomic_write_json(SIGNOFF_DATA_PATH, data)
        _log(f"      Record saved ({overall_status})")
    except Exception as exc:
        _log(f"WARNING: Data write failed: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# REPLY FINALIZER  (draft, or send when configured)
# ══════════════════════════════════════════════════════════════════════════════

def finalize_reply(mail_item, html_body: str, overall_status: str,
                   is_revalidation: bool = False) -> tuple:
    """
    Build the ReplyAll, tag the subject, and either Send() or Save() it.

      • overall_status in ('partial', 'not_found'):
          ALWAYS drafted with escalation recipients applied. AUTO_SEND_ON_ACTIVE
          has no effect here — these outcomes always need human review.
      • overall_status == 'active':
          Sent immediately via reply.Send() only if AUTO_SEND_ON_ACTIVE is True
          and AUTO_SEND_DRY_RUN is False. Otherwise saved as a draft (dry-run
          mode logs what WOULD have been sent so you can validate before
          flipping AUTO_SEND_DRY_RUN off).

    Returns (success: bool, was_sent: bool).
    """
    tag_map = {'active': TAG_ACTIVE, 'partial': TAG_PARTIAL, 'not_found': TAG_NOT_FOUND}
    tag     = tag_map.get(overall_status, TAG_ACTIVE)
    prefix  = '[Revalidated] ' if is_revalidation else ''

    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"{prefix}{tag} {mail_item.Subject}"

        # Partial / Not-Found — escalate, always draft, never auto-send.
        if overall_status in ('partial', 'not_found'):
            if ESCALATION_TO:
                reply.To = '; '.join(ESCALATION_TO)
            if ESCALATION_CC:
                reply.CC = '; '.join(ESCALATION_CC)
            reply.Save()
            _log(f"      Draft saved [{tag}]{' — REVAL' if is_revalidation else ''} "
                 f"(escalation → To: {reply.To}  |  CC: {reply.CC or '(none)'})")
            return True, False

        # Active — every host confirmed reporting.
        if AUTO_SEND_ON_ACTIVE and AUTO_SEND_DRY_RUN:
            _log(f"      DRY-RUN: would SEND [{tag}] → {reply.To} "
                 f"(AUTO_SEND_DRY_RUN=True — saving as draft instead)")
            reply.Save()
            return True, False

        if AUTO_SEND_ON_ACTIVE:
            reply.Send()  # sent immediately — no human review for this outcome
            _log(f"      SENT [{tag}]{' — REVAL' if is_revalidation else ''} → {reply.To}")
            return True, True

        reply.Save()
        _log(f"      Draft saved [{tag}]{' — REVAL' if is_revalidation else ''}")
        return True, False

    except Exception as exc:
        _log(f"      ERROR: Reply creation failed: {exc}")
        return False, False


# ══════════════════════════════════════════════════════════════════════════════
# OUTLOOK
# ══════════════════════════════════════════════════════════════════════════════

# Known pywin32 COM error HRESULTs mapped to an actionable hint. This turns a
# bare "outlook connection failed" into something you can actually act on
# instead of guessing. Not exhaustive — extend as you hit new codes.
_OUTLOOK_COM_ERROR_HINTS = {
    -2147221005: ("Class not registered — Outlook isn't installed for this "
                  "account, or there's a 32-bit/64-bit mismatch between "
                  "Python and Outlook. Match Python's bitness to Outlook's."),
    -2147023174: ("RPC server unavailable — Outlook is not running or is "
                  "still starting up. OUTLOOK_CONNECT_RETRIES will retry "
                  "this automatically; increase it if Outlook is slow to load."),
    -2147024891: ("Access denied — likely the Outlook Object Model Guard or "
                  "antivirus mail-scan add-in blocking programmatic access. "
                  "This usually shows as an on-screen security prompt, which "
                  "blocks unattended runs. See Implementation Notes."),
    -2147352567: ("Outlook rejected the call — often a MAPI profile issue. "
                  "Set OUTLOOK_PROFILE_NAME, or open Outlook manually once "
                  "under the account this script runs as."),
}


def _connect_outlook_application():
    """
    Dispatch Outlook.Application, preferring early binding (gencache) for
    better attribute/type checking, falling back to late binding if the
    gen_py cache is stale (common after an Office update).
    """
    try:
        return win32com.client.gencache.EnsureDispatch('Outlook.Application')
    except AttributeError:
        _log("WARNING: win32com gen_py cache looks stale — using late binding. "
             "(Fix permanently by deleting the folder under "
             "%LOCALAPPDATA%\\Temp\\gen_py and re-running.)")
        return win32com.client.Dispatch('Outlook.Application')


def get_outlook_folders():
    """
    Connect to Outlook and return (inbox, drafts, sent) folder objects.
    Retries transient failures (e.g. Outlook still starting) up to
    OUTLOOK_CONNECT_RETRIES times before giving up.
    """
    last_exc = None
    for attempt in range(1, OUTLOOK_CONNECT_RETRIES + 1):
        try:
            outlook = _connect_outlook_application()
            ns      = outlook.GetNamespace('MAPI')

            if OUTLOOK_PROFILE_NAME:
                ns.Logon(OUTLOOK_PROFILE_NAME, None, False, True)

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

        except pywintypes.com_error as com_exc:
            last_exc = com_exc
            hresult  = com_exc.args[0] if com_exc.args else None
            hint     = _OUTLOOK_COM_ERROR_HINTS.get(
                hresult, "See Implementation Notes for general Outlook COM troubleshooting steps.")
            _log(f"ERROR: Outlook COM connection failed (attempt {attempt}/{OUTLOOK_CONNECT_RETRIES}, "
                 f"HRESULT {hresult}): {com_exc}. {hint}")
        except Exception as exc:
            last_exc = exc
            _log(f"ERROR: Outlook connection failed (attempt {attempt}/{OUTLOOK_CONNECT_RETRIES}): {exc}")

        if attempt < OUTLOOK_CONNECT_RETRIES:
            time.sleep(OUTLOOK_CONNECT_RETRY_DELAY)

    _log(f"ERROR: Giving up on Outlook after {OUTLOOK_CONNECT_RETRIES} attempts. Last error: {last_exc}")
    return None, None, None


# ══════════════════════════════════════════════════════════════════════════════
# SELF-TEST  —  python signoff_runner.py --selftest
# ══════════════════════════════════════════════════════════════════════════════

_SELFTEST_CASES = [
    # (subject, expected_hostnames  |  None = expected to be skipped by guards)
    ('Security Signoff | SRV01',                                   ['SRV01']),
    ('Security Signoff for | SRV01',                               ['SRV01']),
    ('Security Signoff for |',                                     None),
    ('Security Sign-off Request | OnPrem | ABCPRDSQL01',           ['ABCPRDSQL01']),
    ('Security Sign off | On-Prem | PRD-SQL-01',                   ['PRD-SQL-01']),
    ('Security Signoff | OnPrem| SRV01, SRV02; SRV03',             ['SRV01', 'SRV02', 'SRV03']),
    ('Security Signoff - SRV01 - SRV02',                           ['SRV01', 'SRV02']),
    ('Security Signoff | PRD-SQL-01',                              ['PRD-SQL-01']),
    ('Security Signoff | Azure | web-app-03.corp.local',           ['web-app-03.corp.local']),
    ('Security Signoff | 10.20.30.40',                             ['10.20.30.40']),
    ('Security Signoff | SRV01 | SRV02 | Thanks',                   ['SRV01', 'SRV02']),
    ('Security Signoff for SRV01 and SRV02',                       ['SRV01', 'SRV02']),
    ('Security Signoff | Prod Windows SRV01 / Linux SRV02',        ['SRV01', 'SRV02']),
    ('Security SignOff Request | SRV01 | SRV01',                   ['SRV01']),
    ('Security Signoff | OnPrem | Server Name',                    None),
    ('Security Signoff for the new servers',                       None),
    ('RE: Security Signoff | SRV01',                               None),
    ('FWD: Security Signoff | SRV01',                              None),
    ('RE[2]: Security Signoff | SRV01',                            None),
    ('[Processed-Active] Security Signoff | SRV01',                None),
    ('SRV01 | Security Signoff',                                   None),
    ('Weekly patching report | SRV01',                             None),
]


def _selftest() -> int:
    """Run the subject parser over known-good/known-bad subjects. No Outlook, no QRadar."""
    print('=' * 78)
    print('Subject parser self-test — SUBJECT_KEYWORD = '
          f'{SUBJECT_KEYWORD!r}, SUBJECT_SEPARATOR = {SUBJECT_SEPARATOR!r}')
    print('=' * 78)
    failures = 0
    for subject, expected in _SELFTEST_CASES:
        ok, reason = passes_subject_guards(subject)
        if not ok:
            actual, shown = None, f'SKIP ({reason})'
        else:
            actual = extract_hostnames(subject)
            shown  = f'{actual}' + (f'   [ignored: {_last_parse_dropped}]'
                                    if _last_parse_dropped else '')
        passed = (actual == expected)
        failures += int(not passed)
        print(f"  {'PASS' if passed else 'FAIL'}  {subject!r}\n        → {shown}")
        if not passed:
            print(f"        expected: {expected if expected is not None else 'SKIP'}")
    print('=' * 78)
    print(f"{len(_SELFTEST_CASES) - failures}/{len(_SELFTEST_CASES)} passed")
    return 1 if failures else 0


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _ensure_data_file()

    if AUTO_SEND_ON_ACTIVE and AUTO_SEND_DRY_RUN:
        send_policy = 'draft-only | AUTO-SEND enabled but in DRY-RUN (no real sends)'
    elif AUTO_SEND_ON_ACTIVE:
        send_policy = 'Active results SENT immediately | Partial/NotFound always drafted'
    else:
        send_policy = 'draft-only'

    _log('=' * 65)
    _log('QRadar Signoff Runner  v3.2')
    _log(f'  Inbox scan : last {LOOKBACK_DAYS}d')
    _log(f'  Sent scan  : last {SENT_SCAN_DAYS}d')
    _log(f'  Active-skip: {ACTIVE_SKIP_DAYS}d')
    _log(f'  Policy     : first-email-only | {send_policy}')
    _log(f"  Separators : '{SUBJECT_SEPARATOR}' "
         f"{[s for s in EXTRA_HOST_SEPARATORS if s.strip()]} + spaced dash")
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

        lookback_cutoff    = datetime.now() - timedelta(days=LOOKBACK_DAYS)
        active_skip_cutoff = datetime.now() - timedelta(days=ACTIVE_SKIP_DAYS)

        items = _items_since(inbox.Items, lookback_cutoff, 'ReceivedTime',
                             lambda it: _com_dt_to_py(it.ReceivedTime))
        _log(f"\n{len(items)} email(s) found in last {LOOKBACK_DAYS}d\n{'─'*40}")

        processed = skipped = drafted = sent_count = revalidated = 0

        for mail in items:
            # Only process mail items (Class 43)
            try:
                if mail.Class != 43:
                    continue
            except Exception:
                continue

            try:
                subject = mail.Subject or ''
                sender  = _get_sender_smtp(mail)
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
            hostname_list  = extract_hostnames(subject)
            ignored_tokens = list(_last_parse_dropped)
            if not hostname_list:
                skipped += 1
                _log(f"  SKIP (no hostnames parsed): {subject[:60]!r}"
                     + (f"  [ignored: {ignored_tokens}]" if ignored_tokens else ''))
                continue

            # ── Runtime dedup ─────────────────────────────────────────────────
            if is_drafted_this_run(hostname_list):
                skipped += 1
                _log(f"  SKIP (runtime dedup — {hostname_list} already processed this run)")
                continue

            _log(f"\n  Candidate : {subject[:70]!r}")
            _log(f"  Sender    : {sender}")
            _log(f"  Hosts     : {hostname_list}")
            if ignored_tokens:
                _log(f"  Ignored   : {ignored_tokens}")

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

            success, was_sent = finalize_reply(mail, html_body, overall_status, is_revalidation=is_reval)
            if success:
                drafted     += int(not was_sent)
                sent_count  += int(was_sent)
                revalidated += int(is_reval)
                mark_drafted_this_run(hostname_list)
                write_record(
                    email_subject   = subject,
                    sender          = sender,
                    host_records    = host_records,
                    overall_status  = overall_status,
                    is_revalidation = is_reval,
                    prior_status    = last_tag,
                    was_sent        = was_sent,
                    ignored_tokens  = ignored_tokens,
                )
            processed += 1

        _log(f"\n{'='*65}")
        _log(f"Run complete — {processed} processed | {drafted} drafted | {sent_count} sent "
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
    if '--selftest' in sys.argv:
        sys.exit(_selftest())
    try:
        main()
    except KeyboardInterrupt:
        _log("\nInterrupted by user.")
        release_lock()
