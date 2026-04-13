"""
QRadar Security Signoff Auto-Draft  (v2.0)
==========================================
Scans an Outlook folder for SIEM signoff requests, queries QRadar read-only,
and saves a formatted Outlook draft.  Nothing is ever auto-sent.

Two scan modes per execution
──────────────────────────────────────────────────────────────────────────────
  SCAN 1 — New Signoff Scan
      Processes emails received within LOOKBACK_HOURS.
      Active result   → ReplyAll draft (confirms to the original requester).
      Partial/NotFound→ New escalation draft addressed ONLY to the internal
                        SOC DL (ESCALATION_TO + CCs).  Original requester
                        is intentionally excluded from escalations.

  SCAN 2 — Recheck Scan
      Reads the tracking JSON for partial/not-found entries whose
      next_recheck timestamp has elapsed and re-queries QRadar independently.
      No original email is needed — fully standalone.
      Resolved (now active) → [SIEM Resolved] escalation draft.
      Still pending         → draft only if RECHECK_NOTIFY_IF_STILL_PENDING=True.

Tracking file
──────────────────────────────────────────────────────────────────────────────
  A JSON file at RECHECK_TRACKING_FILE persists per-ConversationID state:
    result_type   : 'active' | 'partial' | 'not_found'
    next_recheck  : ISO datetime for next scheduled re-query (None if active)
    attempts      : cumulative recheck count
  Active entries → skip permanently.
  Partial/NotFound entries → skip until next_recheck; then flip to 'recheck'.
  Sent Items / Drafts folder scan retained as a legacy fallback for emails
  processed before the tracking file existed.

All drafts must be reviewed and sent manually.
"""

from __future__ import annotations

import json
import os
import requests
import urllib3
import win32com.client

from datetime import datetime, timedelta
from typing   import Optional

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

QRADAR_HOST     = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL      = False

# Subject matching — hostname extracted from right of SUBJECT_SEPARATOR
SUBJECT_KEYWORD   = 'Security Signoff'
SUBJECT_SEPARATOR = '|'

# New-email scan window
LOOKBACK_HOURS = 24

# Sender allowlist (exact address or @domain; empty list = allow all)
ALLOWED_SENDERS    : list[str] = []
YOUR_EMAIL_ADDRESS = 'youremail@yourorg.com'

# Body trigger — script only acts if this string appears in the email body
TRIGGER_DL = '@SOC-DL@yourorg.com'

# ── Outlook folder names ──────────────────────────────────────────────────────
# SIGNOFF_FOLDER_NAME: dedicated subfolder under Inbox where your Outlook rule
#   routes signoff emails.  Set to None to scan the full Inbox.
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'   # None = full Inbox
DRAFTS_FOLDER_NAME  = 'Drafts'
SENT_FOLDER_NAME    = 'Sent Items'

# ── File paths ────────────────────────────────────────────────────────────────
RUN_LOG_PATH  = r'C:\path\to\your\signoff_runner.log'
LOCKFILE_PATH = r'C:\path\to\your\signoff.lock'

REQUEST_TIMEOUT = 30   # seconds

# ── OS-based log source type validation ──────────────────────────────────────
# First entry in each group's 'required' list is the OS signature.
# Leave empty ({}) to disable type validation and use simple found/not-found.
OS_TYPE_GROUPS: dict = {
    'Windows': {
        'required': ['Microsoft Security', 'WinCollect'],
    },
    'Linux': {
        'required': ['Linux OS'],
    },
}

# ── Onboarding escalation ─────────────────────────────────────────────────────
# Added to CC and mentioned in the reply body when types are missing or silent.
ONBOARD_REQUEST_CC   = 'onboarding-owner@yourorg.com'
ONBOARD_REQUEST_NAME = '@xyz'

# ── Escalation routing (partial / not-found results) ─────────────────────────
# For partial or not-found results a NEW email is drafted — NOT a ReplyAll.
# The original signoff requester is intentionally excluded so that internal
# escalation stays separate from the requester-facing acknowledgement flow.
#
# ESCALATION_TO     : Primary TO recipient(s).  Separate multiple with ';'.
# ESCALATION_CC_LIST: Additional CCs beyond ONBOARD_REQUEST_CC.
ESCALATION_TO      = '@SOC-DL@yourorg.com'
ESCALATION_CC_LIST : list[str] = [
    # 'security-manager@yourorg.com',
    # 'team-lead@yourorg.com',
]

# ── Recheck configuration ─────────────────────────────────────────────────────
# Hosts returning partial or not-found are re-queried automatically once
# RECHECK_INTERVAL_DAYS has elapsed.  State is persisted in a JSON file.
#
# RECHECK_ENABLED              : Set False to disable the recheck mechanism.
# RECHECK_TRACKING_FILE        : Path to the JSON tracking file.
# RECHECK_INTERVAL_DAYS        : Days between re-queries per host.
# RECHECK_MAX_ATTEMPTS         : Stop rechecking after N total attempts (0=∞).
# RECHECK_NOTIFY_IF_STILL_PENDING:
#   True  → create a reminder draft on every recheck cycle (even unresolved).
#   False → create a draft ONLY when the host transitions to active/resolved.
RECHECK_ENABLED                    = True
RECHECK_TRACKING_FILE              = r'C:\path\to\your\signoff_tracking.json'
RECHECK_INTERVAL_DAYS              = 7
RECHECK_MAX_ATTEMPTS               = 4
RECHECK_NOTIFY_IF_STILL_PENDING    = False

# ── Internal constants ────────────────────────────────────────────────────────
ACTIVITY_THRESHOLD_DAYS : int  = 7
_MIN_TS                 : int  = 0
_MAX_TS                 : int  = 2_147_483_647
LOG_SOURCE_TYPES_CACHE  : dict = {}


# ═══════════════════════════════════════════════════════════════════════════════
#  LOGGING
# ═══════════════════════════════════════════════════════════════════════════════

def _log(message: str) -> None:
    """Appends a timestamped entry to the run log and echoes to stdout."""
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    try:
        with open(RUN_LOG_PATH, 'a', encoding='utf-8') as fh:
            fh.write(line + '\n')
    except Exception as exc:
        print(f"⚠️  Could not write to log file: {exc}")


# ═══════════════════════════════════════════════════════════════════════════════
#  LOCKFILE
# ═══════════════════════════════════════════════════════════════════════════════

def acquire_lock() -> bool:
    if os.path.exists(LOCKFILE_PATH):
        _log("⚠️  Lockfile exists — another instance may be running. Exiting.")
        return False
    try:
        with open(LOCKFILE_PATH, 'w') as fh:
            fh.write(str(os.getpid()))
        return True
    except Exception as exc:
        _log(f"❌ Could not create lockfile: {exc}")
        return False


def release_lock() -> None:
    try:
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
    except Exception as exc:
        _log(f"⚠️  Could not remove lockfile: {exc}")


# ═══════════════════════════════════════════════════════════════════════════════
#  RECHECK TRACKING
# ═══════════════════════════════════════════════════════════════════════════════

def load_tracking() -> dict:
    """
    Loads recheck state from JSON.  Returns {} if disabled, missing, or corrupt.
    """
    if not RECHECK_ENABLED:
        return {}
    if not os.path.exists(RECHECK_TRACKING_FILE):
        return {}
    try:
        with open(RECHECK_TRACKING_FILE, 'r', encoding='utf-8') as fh:
            return json.load(fh)
    except Exception as exc:
        _log(f"⚠️  Could not load tracking file: {exc} — starting with empty state.")
        return {}


def save_tracking(tracking: dict) -> None:
    """
    Persists tracking state to JSON.  Creates parent directories if needed.
    Called once at the very end of each run to minimise partial-write risk.
    """
    if not RECHECK_ENABLED:
        return
    try:
        parent = os.path.dirname(RECHECK_TRACKING_FILE)
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(RECHECK_TRACKING_FILE, 'w', encoding='utf-8') as fh:
            json.dump(tracking, fh, indent=2, default=str)
    except Exception as exc:
        _log(f"❌ Could not save tracking file: {exc}")


def update_tracking_entry(
    tracking:    dict,
    conv_id:     str,
    hostname:    str,
    result_type: str,           # 'active' | 'partial' | 'not_found'
    subject:     str = '',
    sender:      str = '',
) -> None:
    """
    Inserts or updates a tracking entry in-place.

    next_recheck is set only for partial/not_found entries.
    Active entries carry next_recheck=None and are permanently retired.
    attempts is always incremented so cumulative recheck count is preserved.
    """
    now      = datetime.now()
    existing = tracking.get(conv_id, {})
    attempts = existing.get('attempts', 0) + 1

    next_recheck: Optional[str] = None
    if result_type in ('partial', 'not_found'):
        next_recheck = (now + timedelta(days=RECHECK_INTERVAL_DAYS)).isoformat()

    tracking[conv_id] = {
        'hostname':        hostname,
        'result_type':     result_type,
        'first_processed': existing.get('first_processed', now.isoformat()),
        'last_checked':    now.isoformat(),
        'next_recheck':    next_recheck,
        'attempts':        attempts,
        'subject':         subject or existing.get('subject', ''),
        'sender':          sender  or existing.get('sender',  ''),
    }


def _conv_id_in_folder(conv_id: str, folder) -> bool:
    """
    Scans an Outlook folder for any item matching conv_id.
    Used as a legacy fallback for emails processed before tracking existed.
    Exits early on first match for performance.
    """
    try:
        for item in folder.Items:
            try:
                if getattr(item, 'ConversationID', None) == conv_id:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def check_dedup_status(
    conv_id:      str,
    tracking:     dict,
    sent_folder,
    drafts_folder,
) -> tuple[str, str]:
    """
    Decides the correct action for an email this run.

    Lookup priority:
      1. Tracking file (authoritative for all recheck decisions).
      2. Sent Items + Drafts scan (legacy fallback for pre-tracking runs only).

    Returns one of:
      ('process', reason) — treat as a new signoff.
      ('recheck', reason) — treat as a scheduled re-query of a prior partial/not-found.
      ('skip',    reason) — do not process this email.
    """
    entry = tracking.get(conv_id)

    if entry:
        result_type = entry.get('result_type', 'active')
        attempts    = entry.get('attempts', 0)
        hostname    = entry.get('hostname', 'unknown')

        # Active entries are permanently retired — skip forever.
        if result_type == 'active':
            return 'skip', f'confirmed active in tracking ({hostname})'

        if result_type in ('partial', 'not_found'):
            # Honour the max-attempts cap.
            if RECHECK_MAX_ATTEMPTS > 0 and attempts >= RECHECK_MAX_ATTEMPTS:
                return 'skip', (
                    f'max recheck attempts ({RECHECK_MAX_ATTEMPTS}) reached '
                    f'for {hostname}'
                )

            # Check whether the recheck window has elapsed.
            next_str = entry.get('next_recheck') or ''
            if next_str:
                try:
                    next_dt = datetime.fromisoformat(next_str)
                    if datetime.now() < next_dt:
                        delta = next_dt - datetime.now()
                        return 'skip', (
                            f'recheck not due for ~{delta.days} day(s) ({hostname})'
                        )
                except ValueError:
                    pass   # Malformed date — fall through to recheck

            return 'recheck', (
                f'recheck due — attempt #{attempts + 1} '
                f'(prev: {result_type}, {hostname})'
            )

    # ── Legacy fallback ───────────────────────────────────────────────────────
    # Only reached when conv_id is absent from tracking (first run with new code,
    # or emails processed before the tracking file was introduced).
    # If found in Sent Items, we don't know the original result_type, so we
    # conservatively skip to avoid duplicate drafts.  The entry is NOT written
    # to tracking here — it remains a passive guard only.
    if _conv_id_in_folder(conv_id, sent_folder):
        return 'skip', 'reply found in Sent Items (legacy fallback)'
    if _conv_id_in_folder(conv_id, drafts_folder):
        return 'skip', 'draft already in Drafts (legacy fallback)'

    return 'process', 'new signoff request'


# ═══════════════════════════════════════════════════════════════════════════════
#  QRADAR CONNECTION
# ═══════════════════════════════════════════════════════════════════════════════

def test_qradar_connection() -> bool:
    """Validates QRadar reachability and credentials before any email work."""
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
        if resp.status_code == 401:
            _log("❌ Authentication failed — check credentials.")
            return False
        _log(f"⚠️  Unexpected QRadar response: {resp.status_code}")
        return False
    except Exception as exc:
        _log(f"❌ QRadar connection failed: {exc}")
        return False


def fetch_log_source_types() -> None:
    """Pre-fetches the Log Source Type ID → Name mapping into the module cache."""
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
                ls_id   = t.get('id')
                ls_name = t.get('name')
                if ls_id is not None and ls_name is not None:
                    LOG_SOURCE_TYPES_CACHE[ls_id] = ls_name
            _log(f"✅ Cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
        else:
            _log(f"⚠️  Failed to fetch Log Source Types: {resp.status_code}")
    except Exception as exc:
        _log(f"❌ Error fetching Log Source Types: {exc}")


# ═══════════════════════════════════════════════════════════════════════════════
#  QRADAR QUERIES — READ ONLY
# ═══════════════════════════════════════════════════════════════════════════════

def _safe_timestamp(timestamp_ms) -> tuple[str, str, Optional[int]]:
    """
    Converts a QRadar epoch-ms timestamp to (human_string, activity, days_ago).
    days_ago=None means the source exists but has never sent an event.
    """
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    try:
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)
        epoch_s  = timestamp_ms / 1000.0 if timestamp_ms > 4_102_444_800 else float(timestamp_ms)
        if epoch_s <= _MIN_TS or epoch_s > _MAX_TS:
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None
        dt        = datetime.fromtimestamp(epoch_s)
        last_seen = dt.strftime('%Y-%m-%d %H:%M:%S')
        days_ago  = (datetime.now() - dt).days
        threshold = datetime.now() - timedelta(days=ACTIVITY_THRESHOLD_DAYS)
        activity  = 'Active' if dt > threshold else 'Inactive'
        return last_seen, activity, days_ago
    except Exception:
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def query_all_log_sources_readonly(hostname: str) -> dict:
    """
    STRICTLY READ-ONLY — HTTP GET only.  QRadar is never modified.

    Returns ALL log sources whose name fuzzy-matches the hostname.
    Unified return format used by all downstream functions:
      {
        'status':  'Found' | 'Not Found' | 'API Error ...',
        'sources': [
          { 'name', 'ls_type', 'enabled', 'last_seen', 'activity', 'days_ago' },
          ...
        ]
      }
    """
    clean    = str(hostname).replace('"', '').replace("'", '').strip()
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
            ls_type_name = LOG_SOURCE_TYPES_CACHE.get(
                type_id, f'Unknown Type ID: {type_id}'
            )
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

    except Exception as exc:
        return {'status': f'Error: {str(exc)[:80]}', 'sources': []}


def validate_expected_types(all_sources_result: dict, required_types: list) -> list:
    """
    Fuzzy-matches each required type keyword against returned QRadar sources.
    Returns one result dict per required type, regardless of found/not-found state.
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

        # Prefer enabled sources; within each group sort by most recent activity.
        enabled  = [s for s in matched if s.get('enabled')]
        disabled = [s for s in matched if not s.get('enabled')]
        enabled.sort(key=lambda x: x.get('days_ago') or 99_999)
        disabled.sort(key=lambda x: x.get('days_ago') or 99_999)
        best = enabled[0] if enabled else disabled[0]

        results.append({
            'expected':  expected_kw,
            'found':     True,
            'ls_type':   best.get('ls_type'),
            'ls_name':   best.get('name'),
            'last_seen': best.get('last_seen'),
            'days_ago':  best.get('days_ago'),   # None = found but zero events yet
        })

    return results


def detect_os_group(sources: list) -> tuple[Optional[str], Optional[dict]]:
    """
    Identifies which OS_TYPE_GROUPS group this device belongs to.
    The first entry in each group's 'required' list is its OS signature.
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


def classify_qradar_result(qradar_result: dict) -> tuple[str, bool]:
    """
    Derives (result_type, needs_cc) cheaply, without building HTML.

    result_type : 'active' | 'partial' | 'not_found'
    needs_cc    : True when escalation CCs should be added to the draft.

    Used to prepare recheck_info before the single HTML build call, avoiding
    a second call to build_reply_body solely for classification purposes.
    """
    status  = qradar_result.get('status', '')
    sources = qradar_result.get('sources', [])

    if status != 'Found' or not sources:
        return 'not_found', False

    if not OS_TYPE_GROUPS:
        return 'active', False

    group_name, group_rules = detect_os_group(sources)
    if group_name is None:
        # Sources exist but no OS signature matched — show raw, no type penalty.
        return 'active', False

    validation = validate_expected_types(qradar_result, group_rules.get('required', []))
    needs_cc   = any(
        not r['found'] or (r['found'] and r['days_ago'] is None)
        for r in validation
    )
    return ('partial' if needs_cc else 'active'), needs_cc


# ═══════════════════════════════════════════════════════════════════════════════
#  SENDER VALIDATION
# ═══════════════════════════════════════════════════════════════════════════════

def is_sender_allowed(sender_address: str) -> bool:
    """Validates sender against ALLOWED_SENDERS.  Empty list = allow all."""
    if not ALLOWED_SENDERS:
        return True
    if not sender_address:
        return False
    sender_clean = sender_address.strip().lower()
    for entry in ALLOWED_SENDERS:
        e = entry.strip().lower()
        if e.startswith('@') and sender_clean.endswith(e):
            return True
        if sender_clean == e:
            return True
    return False


# ═══════════════════════════════════════════════════════════════════════════════
#  SUBJECT GUARDS
# ═══════════════════════════════════════════════════════════════════════════════

def passes_subject_guards(subject: str) -> tuple[bool, str]:
    """All four guards must pass.  Returns (passed, reason_string)."""
    if not subject:
        return False, 'empty subject'
    subj  = subject.strip()
    lower = subj.lower()
    if any(lower.startswith(p) for p in ('re:', 'fw:', 'fwd:')):
        return False, f"reply/forward prefix: '{subj[:30]}'"
    if '[processed]' in lower:
        return False, 'subject contains [Processed] tag'
    if SUBJECT_SEPARATOR not in subj:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found"
    left = subj.split(SUBJECT_SEPARATOR)[0].strip().lower()
    if SUBJECT_KEYWORD.lower() not in left:
        return False, f"keyword '{SUBJECT_KEYWORD}' not found left of separator"
    return True, 'ok'


def extract_hostname(subject: str) -> Optional[str]:
    """Extracts the hostname from the right side of SUBJECT_SEPARATOR."""
    parts    = subject.split(SUBJECT_SEPARATOR, 1)
    hostname = parts[1].strip() if len(parts) > 1 else ''
    return hostname or None


# ═══════════════════════════════════════════════════════════════════════════════
#  HTML BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def _build_context_block(original_context: Optional[dict]) -> str:
    """
    Renders a shaded 'Original Request' reference block.
    Included in first-time escalation emails so the SOC DL knows which
    signoff request triggered the escalation.
    """
    if not original_context:
        return ''
    return (
        "<div style=\"background:#f5f5f5;border-left:3px solid #bbb;"
        "padding:8px 14px;margin:0 0 16px 0;border-radius:0 4px 4px 0;"
        "font-size:12px;color:#555;\">"
        "<b>Original Signoff Request</b><br>"
        f"From:&nbsp;&nbsp;&nbsp;&nbsp; {original_context.get('sender', 'N/A')}<br>"
        f"Subject:&nbsp; {original_context.get('subject', 'N/A')}<br>"
        f"Received: {original_context.get('received_time', 'N/A')}"
        "</div>"
    )


def _build_recheck_block(recheck_info: Optional[dict]) -> str:
    """
    Renders a coloured recheck-notification banner above the main result table.
    Green when the host is now resolved; amber when still pending.
    """
    if not recheck_info:
        return ''

    attempt     = recheck_info.get('attempt', 1)
    prev_type   = recheck_info.get('prev_type', 'unknown').replace('_', ' ')
    is_resolved = recheck_info.get('is_resolved', False)

    if is_resolved:
        colour = '#1a6e36'
        label  = '🔄&nbsp; RESOLVED — Host is now reporting on SIEM'
        note   = (
            f"Scheduled recheck (attempt #{attempt}). "
            f"Previous status was &quot;{prev_type}&quot;."
        )
    else:
        colour = '#c87800'
        label  = f'🔄&nbsp; RECHECK UPDATE — Attempt #{attempt}'
        note   = (
            f"Status remains &quot;{prev_type}&quot; as of this recheck. "
            f"Awaiting resolution."
        )

    return (
        f"<div style=\"background:{colour};color:#fff;padding:8px 16px;"
        f"border-radius:6px;font-size:12px;font-weight:600;margin-bottom:10px;"
        f"letter-spacing:0.2px;\">{label}</div>"
        f"<p style=\"font-size:12px;color:#666;margin:0 0 16px 0;"
        f"font-style:italic;\">{note}</p>"
    )


def _build_reply_html(
    hostname:         str,
    qradar_result:    dict,
    type_validation:  Optional[list] = None,
    os_group:         Optional[str]  = None,
    original_context: Optional[dict] = None,
    recheck_info:     Optional[dict] = None,
) -> str:
    """
    Assembles the full HTML body for any result state.

    States handled:
      Not Found          → red banner
      All types confirmed → green banner, all rows green
      Any type silent    → amber banner, silent rows amber + CC person called out
      Any type missing   → amber banner, missing rows red  + CC person called out
      Simple mode (no OS_TYPE_GROUPS) → green banner, single-source detail table
    """
    status   = qradar_result.get('status') if qradar_result else 'Not Found'
    sources  = qradar_result.get('sources', []) if qradar_result else []
    run_time = datetime.now().strftime('%d %B %Y, %H:%M')

    recheck_block = _build_recheck_block(recheck_info)
    context_block = _build_context_block(original_context)

    # ── NOT FOUND ─────────────────────────────────────────────────────────────
    if status != 'Found' or not sources:
        body_inner = (
            f"{recheck_block}{context_block}"
            "<div style=\"background:#c0392b;color:#fff;padding:10px 16px;"
            "border-radius:6px;font-size:13px;font-weight:600;margin-bottom:16px;\">"
            "✖&nbsp; Not Found in QRadar</div>"
            f"<p style=\"margin:0 0 4px 0;\"><b>{hostname}</b> was <b>not found</b> "
            "in the QRadar log source inventory. Please ensure the asset is onboarded "
            "and configured correctly in SIEM.</p>"
        )
        banner_colour = None   # Not used in this path
        detail_block  = ''

    # ── TYPE VALIDATION MODE ──────────────────────────────────────────────────
    elif type_validation is not None:
        os_label    = f' ({os_group} device)' if os_group else ''
        any_missing = any(not r['found'] for r in type_validation)
        any_silent  = any(r['found'] and r['days_ago'] is None for r in type_validation)
        any_problem = any_missing or any_silent

        if not any_problem:
            banner_colour = '#1a7a4a'
            banner_label  = '✔&nbsp; Confirmed Reporting on SIEM'
            summary_line  = (
                f"<b>{hostname}</b> is reporting on our SIEM{os_label}. "
                "All required log sources are present and active."
            )
        else:
            banner_colour = '#c87800'
            found_n       = sum(1 for r in type_validation if r['found'])
            total_n       = len(type_validation)
            if any_missing:
                banner_label = '⚠&nbsp; Partial — Log Sources Missing'
                summary_line = (
                    f"<b>{hostname}</b>{os_label} — {found_n} of {total_n} required "
                    "log sources found on SIEM. Missing sources are highlighted below."
                )
            else:
                banner_label = '⚠&nbsp; Partial — Log Sources Not Reporting'
                summary_line = (
                    f"<b>{hostname}</b>{os_label} — all {total_n} required log sources "
                    "are onboarded but one or more have not sent any events yet."
                )

        type_rows = ''
        for r in type_validation:
            days_str = (
                "<span style='color:#888;font-size:11px;'>"
                + ('Today' if r['days_ago'] == 0 else f"{r['days_ago']} days ago")
                + "</span>"
                if r['days_ago'] is not None else ''
            )

            if not r['found']:
                note = (
                    f"{ONBOARD_REQUEST_NAME}, please onboard this log source on SIEM."
                    if ONBOARD_REQUEST_NAME.strip() else
                    "Not found — please onboard."
                )
                type_rows += (
                    "<tr style=\"background:#fff5f5;\">"
                    "<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    "font-size:12px;color:#c0392b;font-weight:700;width:24px;"
                    "text-align:center;\">✖</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#c0392b;font-weight:600;\">{r['expected']}</td>"
                    "<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    "font-size:12px;color:#c0392b;\">—</td>"
                    "<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    "font-size:12px;color:#c0392b;\">—</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#c0392b;font-style:italic;\">{note}</td>"
                    "</tr>"
                )

            elif r['days_ago'] is None:
                # Source exists in QRadar but has never sent an event.
                note = (
                    f"{ONBOARD_REQUEST_NAME}, no events received yet — "
                    "please investigate."
                    if ONBOARD_REQUEST_NAME.strip() else
                    "No events received yet — please investigate."
                )
                type_rows += (
                    "<tr style=\"background:#fffbf0;\">"
                    "<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    "font-size:12px;color:#c87800;font-weight:700;width:24px;"
                    "text-align:center;\">⚠</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#c87800;font-weight:600;\">{r['expected']}</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#555;\">{r.get('ls_name', 'N/A')}</td>"
                    "<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    "font-size:12px;color:#c87800;font-style:italic;\">"
                    "No events recorded</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#c87800;font-style:italic;\">{note}</td>"
                    "</tr>"
                )

            else:
                # Source found and actively reporting.
                type_rows += (
                    "<tr style=\"background:#f0faf4;\">"
                    "<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    "font-size:12px;color:#1a7a4a;font-weight:700;width:24px;"
                    "text-align:center;\">✔</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#333;font-weight:600;\">{r['expected']}</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#555;\">{r.get('ls_name', 'N/A')}</td>"
                    f"<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    f"font-size:12px;color:#555;\">{r.get('last_seen', 'N/A')} {days_str}</td>"
                    "<td style=\"padding:8px 12px;border-bottom:1px solid #e8e8e8;"
                    "font-size:12px;color:#1a7a4a;font-weight:600;\">Confirmed</td>"
                    "</tr>"
                )

        detail_block = (
            "<table style=\"width:100%;max-width:680px;border-collapse:collapse;"
            "margin-top:16px;border:1px solid #e0e0e0;border-radius:6px;"
            "overflow:hidden;\">"
            "<tr style=\"background:#f5f5f5;\">"
            "<th style=\"padding:7px 12px;font-size:11px;color:#777;font-weight:600;"
            "text-align:left;border-bottom:2px solid #ddd;width:24px;\"></th>"
            "<th style=\"padding:7px 12px;font-size:11px;color:#777;font-weight:600;"
            "text-align:left;border-bottom:2px solid #ddd;\">Log Source Type</th>"
            "<th style=\"padding:7px 12px;font-size:11px;color:#777;font-weight:600;"
            "text-align:left;border-bottom:2px solid #ddd;\">Log Source Name</th>"
            "<th style=\"padding:7px 12px;font-size:11px;color:#777;font-weight:600;"
            "text-align:left;border-bottom:2px solid #ddd;\">Last Event</th>"
            "<th style=\"padding:7px 12px;font-size:11px;color:#777;font-weight:600;"
            "text-align:left;border-bottom:2px solid #ddd;\">Status</th>"
            "</tr>"
            f"{type_rows}</table>"
        )

        body_inner = (
            f"{recheck_block}{context_block}"
            f"<div style=\"background:{banner_colour};color:#fff;padding:10px 16px;"
            f"border-radius:6px;font-size:13px;font-weight:600;margin-bottom:16px;"
            f"letter-spacing:0.2px;\">{banner_label}</div>"
            f"<p style=\"margin:0 0 4px 0;\">{summary_line}</p>"
            f"{detail_block}"
        )

    # ── SIMPLE MODE (OS_TYPE_GROUPS empty or OS undetected) ───────────────────
    else:
        banner_colour = '#1a7a4a'
        banner_label  = '✔&nbsp; Confirmed Reporting on SIEM'
        summary_line  = f"<b>{hostname}</b> is reporting on our SIEM."

        enabled  = [s for s in sources if s.get('enabled')]
        disabled = [s for s in sources if not s.get('enabled')]
        enabled.sort(key=lambda x: x.get('days_ago') or 99_999)
        best_src = enabled[0] if enabled else (disabled[0] if disabled else None)

        if best_src:
            days_val  = best_src.get('days_ago')
            days_disp = (
                'Today'               if days_val == 0     else
                f"{days_val} days ago" if days_val is not None else 'N/A'
            )
            detail_block = (
                "<table style=\"width:100%;max-width:480px;border-collapse:collapse;"
                "margin-top:16px;border:1px solid #e0e0e0;border-radius:6px;"
                "overflow:hidden;\">"
                "<tr><td style=\"padding:7px 12px;color:#555;font-size:13px;"
                "border-bottom:1px solid #eee;width:160px;\">Log Source Name</td>"
                f"<td style=\"padding:7px 12px;font-size:13px;border-bottom:1px solid #eee;"
                f"font-weight:600;color:#222;\">{best_src.get('name', 'N/A')}</td></tr>"
                "<tr><td style=\"padding:7px 12px;color:#555;font-size:13px;"
                "border-bottom:1px solid #eee;\">Log Source Type</td>"
                f"<td style=\"padding:7px 12px;font-size:13px;border-bottom:1px solid #eee;"
                f"color:#333;\">{best_src.get('ls_type', 'N/A')}</td></tr>"
                "<tr><td style=\"padding:7px 12px;color:#555;font-size:13px;\">"
                "Last Event</td>"
                f"<td style=\"padding:7px 12px;font-size:13px;color:#333;\">"
                f"{best_src.get('last_seen', 'N/A')}"
                f"&nbsp;<span style=\"color:#888;font-size:12px;\">({days_disp})</span>"
                "</td></tr></table>"
            )
        else:
            detail_block = ''

        body_inner = (
            f"{recheck_block}{context_block}"
            f"<div style=\"background:{banner_colour};color:#fff;padding:10px 16px;"
            f"border-radius:6px;font-size:13px;font-weight:600;margin-bottom:16px;"
            f"letter-spacing:0.2px;\">{banner_label}</div>"
            f"<p style=\"margin:0 0 4px 0;\">{summary_line}</p>"
            f"{detail_block}"
        )

    return (
        "<html><body style=\"font-family:'Segoe UI',Arial,sans-serif;color:#222;"
        "font-size:13px;line-height:1.6;margin:0;padding:0;\">"
        "<div style=\"max-width:700px;padding:20px 0;\">"
        "<p style=\"margin:0 0 16px 0;\">Hi,</p>"
        f"{body_inner}"
        "<p style=\"margin:20px 0 4px 0;color:#555;font-size:12px;\">"
        "This is an automated response from the SIEM monitoring system.<br>"
        f"Checked against QRadar on {run_time}.</p>"
        "<p style=\"margin:16px 0 0 0;\">Regards,<br>"
        "<span style=\"font-weight:600;\">Cyberdefence</span></p>"
        "</div></body></html>"
    )


def build_reply_body(
    hostname:         str,
    qradar_result:    dict,
    original_context: Optional[dict] = None,
    recheck_info:     Optional[dict] = None,
) -> tuple[str, bool, str]:
    """
    Routes to the correct HTML builder variant and returns
    (html_body, needs_cc, result_type).

    result_type : 'active' | 'partial' | 'not_found'
    needs_cc    : True when ONBOARD_REQUEST_CC should be included.
    original_context : Injected into escalation emails as an 'Original Request' block.
    recheck_info     : Injected as a coloured recheck-status banner.
    """
    status  = qradar_result.get('status')
    sources = qradar_result.get('sources', [])

    if status != 'Found' or not sources:
        html = _build_reply_html(
            hostname, qradar_result,
            original_context=original_context,
            recheck_info=recheck_info,
        )
        return html, False, 'not_found'

    if not OS_TYPE_GROUPS:
        html = _build_reply_html(
            hostname, qradar_result,
            original_context=original_context,
            recheck_info=recheck_info,
        )
        return html, False, 'active'

    group_name, group_rules = detect_os_group(sources)

    if group_name is None:
        _log(f"      ⚠️  OS group undetected for {hostname} — showing raw sources.")
        html = _build_reply_html(
            hostname, qradar_result,
            original_context=original_context,
            recheck_info=recheck_info,
        )
        return html, False, 'active'

    validation = validate_expected_types(qradar_result, group_rules.get('required', []))
    needs_cc   = any(
        not r['found'] or (r['found'] and r['days_ago'] is None)
        for r in validation
    )
    result_type = 'partial' if needs_cc else 'active'

    html = _build_reply_html(
        hostname, qradar_result,
        type_validation=validation,
        os_group=group_name,
        original_context=original_context,
        recheck_info=recheck_info,
    )
    return html, needs_cc, result_type


# ═══════════════════════════════════════════════════════════════════════════════
#  DRAFT BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def _build_escalation_cc() -> str:
    """Returns the full semicolon-delimited CC string for escalation drafts."""
    parts = []
    if ONBOARD_REQUEST_CC.strip():
        parts.append(ONBOARD_REQUEST_CC.strip())
    parts.extend(a.strip() for a in ESCALATION_CC_LIST if a.strip())
    return '; '.join(parts)


def create_active_reply_draft(mail_item, html_body: str, hostname: str) -> bool:
    """
    Creates a ReplyAll draft for confirmed-active results.
    The original signoff requester receives a direct confirmation reply.
    Subject prefixed with [Processed] to prevent re-processing on future scans.
    """
    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"[Processed] {mail_item.Subject}"
        reply.Save()
        _log(f"      ✅ Active reply draft saved for: {hostname}")
        return True
    except Exception as exc:
        _log(f"      ❌ Failed to create active reply draft for {hostname}: {exc}")
        return False


def create_escalation_draft(
    outlook_app,
    html_body:   str,
    hostname:    str,
    subject_tag: str = '[SIEM Escalation]',
) -> bool:
    """
    Creates a NEW outgoing draft for partial and not-found results.

    Addressed to ESCALATION_TO with CC to ONBOARD_REQUEST_CC + ESCALATION_CC_LIST.
    The original signoff requester is intentionally excluded — escalations are
    internal SOC actions, not requester-facing acknowledgements.

    Also used by the recheck scan for resolved ([SIEM Resolved]) notifications.
    Outlook.CreateItem(0) is used rather than ReplyAll so that no original thread
    is contaminated and routing is fully under config control.
    """
    try:
        mail         = outlook_app.CreateItem(0)   # 0 = olMailItem
        mail.To      = ESCALATION_TO
        cc_str       = _build_escalation_cc()
        if cc_str:
            mail.CC  = cc_str
        mail.Subject  = f"{subject_tag} {hostname}"
        mail.HTMLBody = html_body
        mail.Save()
        _log(f"      ✅ Escalation draft ({subject_tag}) saved for: {hostname}")
        return True
    except Exception as exc:
        _log(f"      ❌ Failed to create escalation draft for {hostname}: {exc}")
        return False


# ═══════════════════════════════════════════════════════════════════════════════
#  RECHECK SCAN
# ═══════════════════════════════════════════════════════════════════════════════

def recheck_pending_entries(
    outlook_app,
    tracking:           dict,
    processed_conv_ids: set,
) -> int:
    """
    Standalone recheck scan — runs after the new-email scan each execution.

    Iterates all tracking entries where:
      - result_type in ('partial', 'not_found')
      - next_recheck timestamp has elapsed
      - attempts < RECHECK_MAX_ATTEMPTS (or RECHECK_MAX_ATTEMPTS == 0)
      - conv_id not in processed_conv_ids (not already handled this run)

    For each due entry:
      1. Re-queries QRadar by the stored hostname (no original email needed).
      2. If now active   → creates [SIEM Resolved] escalation draft.
      3. If still pending → creates [SIEM Recheck] draft only if
                            RECHECK_NOTIFY_IF_STILL_PENDING is True.
      4. Updates the tracking entry (next_recheck advanced; active entries retired).

    Returns the number of recheck drafts created.
    """
    if not RECHECK_ENABLED:
        return 0

    _log("\n🔁 [SCAN 2] Recheck scan — checking pending partial/not-found entries...")
    now         = datetime.now()
    drafts_made = 0
    due_entries : list[tuple[str, dict]] = []

    for conv_id, entry in tracking.items():
        if conv_id in processed_conv_ids:
            continue   # Already handled by new-email scan this run

        result_type = entry.get('result_type', 'active')
        if result_type not in ('partial', 'not_found'):
            continue

        attempts = entry.get('attempts', 0)
        if RECHECK_MAX_ATTEMPTS > 0 and attempts >= RECHECK_MAX_ATTEMPTS:
            continue

        next_str = entry.get('next_recheck') or ''
        if next_str:
            try:
                next_dt = datetime.fromisoformat(next_str)
                if now < next_dt:
                    continue   # Not due yet
            except ValueError:
                pass   # Malformed date — treat as immediately due

        due_entries.append((conv_id, entry))

    if not due_entries:
        _log("   ✅ No recheck entries due at this time.")
        return 0

    _log(f"   {len(due_entries)} entry/entries due for recheck.")

    for conv_id, entry in due_entries:
        hostname    = entry.get('hostname', '')
        prev_type   = entry.get('result_type', 'unknown')
        attempt_num = entry.get('attempts', 0) + 1

        _log(f"\n🔁 Rechecking: {hostname}  (attempt #{attempt_num}, prev: {prev_type})")

        if not hostname:
            _log("      ⚠️  Hostname missing from tracking entry — skipping.")
            continue

        # ── Re-query QRadar ───────────────────────────────────────────────────
        qradar_result              = query_all_log_sources_readonly(hostname)
        new_result_type, needs_cc  = classify_qradar_result(qradar_result)
        is_resolved                = (new_result_type == 'active')

        _log(
            f"      📊 Status: {qradar_result['status']} | "
            f"Result: {new_result_type} | "
            f"Sources: {len(qradar_result.get('sources', []))}"
        )

        # ── Build recheck context ─────────────────────────────────────────────
        recheck_info = {
            'attempt':     attempt_num,
            'prev_type':   prev_type,
            'is_resolved': is_resolved,
        }

        # ── Decide whether to create a draft ──────────────────────────────────
        should_draft = is_resolved or RECHECK_NOTIFY_IF_STILL_PENDING

        if should_draft:
            html, _, _ = build_reply_body(
                hostname,
                qradar_result,
                original_context=None,       # Original email unavailable here
                recheck_info=recheck_info,
            )
            subject_tag = '[SIEM Resolved]' if is_resolved else '[SIEM Recheck]'
            success = create_escalation_draft(
                outlook_app, html, hostname, subject_tag=subject_tag,
            )
            if success:
                drafts_made += 1
        else:
            _log(
                f"      ℹ️  Still {new_result_type} — no draft "
                f"(RECHECK_NOTIFY_IF_STILL_PENDING=False). "
                f"Next recheck in {RECHECK_INTERVAL_DAYS}d."
            )

        # ── Update tracking ───────────────────────────────────────────────────
        # Active entries get next_recheck=None and are permanently retired.
        # Partial/not-found entries have next_recheck advanced by the interval.
        # We write unconditionally (regardless of draft success) so attempts are
        # always incremented and the recheck window advances — this prevents
        # a broken Outlook session from causing an infinite retry storm.
        update_tracking_entry(
            tracking,
            conv_id,
            hostname    = hostname,
            result_type = new_result_type,
            subject     = entry.get('subject', ''),
            sender      = entry.get('sender',  ''),
        )
        # update_tracking_entry increments from existing count but entry has not
        # yet been mutated in this call, so the result is attempt_num.  No
        # manual override needed — the value matches.

    return drafts_made


# ═══════════════════════════════════════════════════════════════════════════════
#  OUTLOOK SETUP
# ═══════════════════════════════════════════════════════════════════════════════

def get_outlook_connection() -> tuple:
    """
    Connects to the running Outlook instance and resolves the required folders.
    Returns (outlook_app, inbox, drafts, sent) or (None, None, None, None).

    The outlook_app object is returned so that create_escalation_draft() can
    call outlook_app.CreateItem(0) for new (non-reply) draft emails.
    """
    try:
        outlook    = win32com.client.Dispatch('Outlook.Application')
        ns         = outlook.GetNamespace('MAPI')
        main_inbox = ns.GetDefaultFolder(6)    # 6  = Inbox
        drafts     = ns.GetDefaultFolder(16)   # 16 = Drafts
        sent       = ns.GetDefaultFolder(5)    # 5  = Sent Items

        if SIGNOFF_FOLDER_NAME:
            try:
                inbox = main_inbox.Folders[SIGNOFF_FOLDER_NAME]
                _log(f"📁 Monitoring subfolder: Inbox\\{SIGNOFF_FOLDER_NAME}")
            except Exception:
                _log(
                    f"⚠️  Subfolder '{SIGNOFF_FOLDER_NAME}' not found — "
                    f"falling back to full Inbox. Create the Outlook rule first."
                )
                inbox = main_inbox
        else:
            inbox = main_inbox
            _log("📁 Monitoring: Full Inbox (no subfolder configured)")

        return outlook, inbox, drafts, sent

    except Exception as exc:
        _log(f"❌ Could not connect to Outlook: {exc}. Is Outlook open and logged in?")
        return None, None, None, None


# ═══════════════════════════════════════════════════════════════════════════════
#  BODY DL CHECK
# ═══════════════════════════════════════════════════════════════════════════════

def body_contains_dl(mail_item) -> bool:
    """
    Returns True if TRIGGER_DL appears anywhere in the email body.
    Checks plain text first; falls back to HTML body.  Case-insensitive.
    Returns True unconditionally when TRIGGER_DL is empty (check disabled).
    """
    if not TRIGGER_DL.strip():
        return True
    dl_lower = TRIGGER_DL.strip().lower()
    try:
        if dl_lower in (mail_item.Body or '').lower():
            return True
        if dl_lower in (mail_item.HTMLBody or '').lower():
            return True
        return False
    except Exception as exc:
        _log(f"⚠️  Could not read email body for DL check: {exc}")
        return False


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _log('=' * 68)
    _log('🚀 QRadar Signoff Auto-Draft  v2.0  starting...')
    _log(f"   Lookback   : {LOOKBACK_HOURS}h  |  Keyword: '{SUBJECT_KEYWORD}'")
    _log(f"   Separator  : '{SUBJECT_SEPARATOR}'  |  "
         f"Senders: {'ALL' if not ALLOWED_SENDERS else len(ALLOWED_SENDERS)}")
    _log(f"   Trigger DL : '{TRIGGER_DL}'")
    _log(f"   Folder     : '{SIGNOFF_FOLDER_NAME or 'Full Inbox'}'")
    _log(
        f"   Recheck    : {'ENABLED' if RECHECK_ENABLED else 'DISABLED'}"
        f" | Interval: {RECHECK_INTERVAL_DAYS}d"
        f" | Max attempts: {RECHECK_MAX_ATTEMPTS if RECHECK_MAX_ATTEMPTS else '∞'}"
    )
    _log(f"   Escalation : TO → {ESCALATION_TO}")
    _log(f"   MODE       : DRAFT ONLY — nothing is ever sent automatically")

    if not acquire_lock():
        return

    # Load tracking before the try/finally so save_tracking always runs in finally.
    tracking = load_tracking()
    _log(f"   Tracking   : {len(tracking)} entries loaded.")

    try:
        # ── Outlook ───────────────────────────────────────────────────────────
        outlook, inbox, drafts, sent = get_outlook_connection()
        if outlook is None:
            return

        # ── QRadar ────────────────────────────────────────────────────────────
        if not test_qradar_connection():
            _log("❌ QRadar unreachable — exiting. All emails left untouched.")
            return
        fetch_log_source_types()

        # ═════════════════════════════════════════════════════════════════════
        # SCAN 1 — New Signoff Emails
        # ═════════════════════════════════════════════════════════════════════
        cutoff_time = datetime.now() - timedelta(hours=LOOKBACK_HOURS)
        _log(
            f"\n📬 [SCAN 1] New signoffs since "
            f"{cutoff_time.strftime('%Y-%m-%d %H:%M:%S')}..."
        )

        cutoff_str  = cutoff_time.strftime('%m/%d/%Y %I:%M %p')
        inbox_items = list(
            inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'")
        )
        _log(f"   {len(inbox_items)} email(s) in window.")

        processed          : int      = 0
        skipped            : int      = 0
        drafted            : int      = 0
        processed_conv_ids : set[str] = set()   # Shared with recheck scan

        for mail_item in inbox_items:

            # Skip non-mail items (calendar invites, meeting requests, etc.).
            try:
                if mail_item.Class != 43:    # 43 = olMail
                    continue
            except Exception:
                continue

            subject = ''
            try:
                subject = mail_item.Subject or ''
            except Exception:
                continue

            # ── Subject guards ────────────────────────────────────────────────
            passed, reason = passes_subject_guards(subject)
            if not passed:
                skipped += 1
                _log(f"   ⏭️  SKIP (subject — {reason}): '{subject[:60]}'")
                continue

            # ── Sender guards ─────────────────────────────────────────────────
            sender = ''
            try:
                sender = mail_item.SenderEmailAddress or ''
            except Exception:
                pass

            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                _log(f"   ⏭️  SKIP (own address): '{subject[:60]}'")
                continue

            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"   ⏭️  SKIP (sender not in allowlist — {sender})")
                continue

            # ── Body DL check ─────────────────────────────────────────────────
            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"   ⏭️  SKIP ('{TRIGGER_DL}' not in body): '{subject[:60]}'")
                continue

            # ── Hostname extraction ───────────────────────────────────────────
            hostname = extract_hostname(subject)
            if not hostname:
                skipped += 1
                _log(f"   ⏭️  SKIP (empty hostname after separator): '{subject[:60]}'")
                continue

            _log(f"\n🔹 Candidate: '{subject[:70]}'")
            _log(f"      Sender  : {sender}")
            _log(f"      Hostname: {hostname}")

            # ── Deduplication / recheck decision ──────────────────────────────
            conv_id = ''
            try:
                conv_id = mail_item.ConversationID or ''
            except Exception:
                pass

            action, action_reason = check_dedup_status(
                conv_id, tracking, sent, drafts
            )

            if action == 'skip':
                skipped += 1
                _log(f"      ⏭️  SKIP ({action_reason})")
                continue

            is_recheck = (action == 'recheck')
            if is_recheck:
                _log(f"      🔁 RECHECK — {action_reason}")

            processed_conv_ids.add(conv_id)

            # ── QRadar query ──────────────────────────────────────────────────
            _log("      🔍 Querying QRadar...")
            qradar_result = query_all_log_sources_readonly(hostname)
            _log(
                f"      📊 Status: {qradar_result['status']} | "
                f"Sources: {len(qradar_result.get('sources', []))}"
            )

            # ── Classify result (needed before HTML build for recheck context) ─
            result_type, _ = classify_qradar_result(qradar_result)

            # ── Build recheck_info (recheck path only) ────────────────────────
            recheck_info: Optional[dict] = None
            if is_recheck:
                prev_entry  = tracking.get(conv_id, {})
                recheck_info = {
                    'attempt':     prev_entry.get('attempts', 0) + 1,
                    'prev_type':   prev_entry.get('result_type', 'unknown'),
                    'is_resolved': result_type == 'active',
                }

            # ── Build original_context (escalation path, first-time only) ─────
            # Included in escalation emails so the SOC DL has full request context.
            original_context: Optional[dict] = None
            if result_type in ('partial', 'not_found') and not is_recheck:
                try:
                    rt = mail_item.ReceivedTime
                    original_context = {
                        'sender':        sender,
                        'subject':       subject,
                        'received_time': rt.strftime('%Y-%m-%d %H:%M') if rt else 'N/A',
                    }
                except Exception:
                    pass

            # ── Build HTML ────────────────────────────────────────────────────
            html, needs_cc, result_type = build_reply_body(
                hostname,
                qradar_result,
                original_context=original_context,
                recheck_info=recheck_info,
            )

            # ── Create draft ──────────────────────────────────────────────────
            #
            # Active result (new or recheck-resolved):
            #   • New signoff → ReplyAll to original thread (requester gets confirmation).
            #   • Recheck resolved → Escalation draft [SIEM Resolved] to SOC DL
            #     (the original requester's thread would be stale; DL needs the update).
            #
            # Partial / Not Found:
            #   → New escalation draft to ESCALATION_TO ONLY.
            #     Original requester is deliberately excluded — this is an internal action.
            if result_type == 'active' and not is_recheck:
                success = create_active_reply_draft(mail_item, html, hostname)
            elif result_type == 'active' and is_recheck:
                success = create_escalation_draft(
                    outlook, html, hostname, subject_tag='[SIEM Resolved]'
                )
            else:
                success = create_escalation_draft(
                    outlook, html, hostname, subject_tag='[SIEM Escalation]'
                )

            # ── Update tracking ───────────────────────────────────────────────
            # Written only on successful draft creation so that a broken Outlook
            # session does not silently suppress future recheck attempts.
            if success:
                drafted += 1
                update_tracking_entry(
                    tracking, conv_id, hostname,
                    result_type=result_type,
                    subject=subject,
                    sender=sender,
                )

            processed += 1

        _log(
            f"\n   [SCAN 1 COMPLETE] Processed: {processed} | "
            f"Drafted: {drafted} | Skipped: {skipped}"
        )

        # ═════════════════════════════════════════════════════════════════════
        # SCAN 2 — Recheck Pending Entries
        # ═════════════════════════════════════════════════════════════════════
        recheck_drafted = recheck_pending_entries(
            outlook, tracking, processed_conv_ids
        )

        # ── Summary ───────────────────────────────────────────────────────────
        _log(f"\n{'=' * 68}")
        _log("✅ Run complete.")
        _log(f"   New signoff drafts  : {drafted}")
        _log(f"   Recheck drafts      : {recheck_drafted}")
        _log(f"   Emails skipped      : {skipped}")
        _log(f"   Tracking entries    : {len(tracking)}")
        _log("   Review all drafts in Outlook Drafts folder before sending.")
        _log(f"{'=' * 68}\n")

    finally:
        # Always save tracking and release the lock — even on unexpected errors.
        save_tracking(tracking)
        release_lock()


if __name__ == '__main__':
    main()
