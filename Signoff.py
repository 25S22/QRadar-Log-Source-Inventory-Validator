import requests
import urllib3
import os
import win32com.client

from datetime import datetime, timedelta

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
QRADAR_HOST     = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL      = False

# Subject matching — email must contain SUBJECT_KEYWORD and SUBJECT_SEPARATOR
# and the hostname is extracted from the RIGHT side of the separator.
# Example subject: "Security Signoff | HOSTNAME-01"
SUBJECT_KEYWORD   = 'Security Signoff'   # word(s) that must appear left of the pipe
SUBJECT_SEPARATOR = '|'

# Only emails received within this window are considered
LOOKBACK_HOURS = 24

# Sender allowlist — only these addresses trigger a draft.
# Use exact addresses OR @domain entries for whole-domain matching.
# Example: 'analyst@org.com' or '@soc.org.com'
# Leave EMPTY to allow all senders — all other checks remain active.
ALLOWED_SENDERS = []

# Your own reply address — used to guard against processing your own sent items
# that may appear in the inbox (e.g. on shared mailboxes)
YOUR_EMAIL_ADDRESS = 'youremail@yourorg.com'

# DL string that must appear somewhere in the email body for the script to act.
# The body wording does not matter — only this string being present matters.
# Example: '@SOC-DL@yourorg.com' or just 'SOC-Team' — whatever your DL looks like.
# Case-insensitive match. Set to '' to disable this check entirely.
TRIGGER_DL = '@SOC-DL@yourorg.com'

# Outlook folder names
# Set SIGNOFF_FOLDER_NAME to the dedicated subfolder under Inbox where your
# Outlook rule routes signoff emails. If you have not created a rule yet and
# want to scan the full Inbox temporarily, set this to None.
# Example: 'SIEM Signoffs' → scans Inbox\SIEM Signoffs only
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'   # set to None to scan full Inbox

DRAFTS_FOLDER_NAME = 'Drafts'           # leave as-is unless using a custom drafts folder
SENT_FOLDER_NAME   = 'Sent Items'       # leave as-is unless using a custom sent folder

# Where the run log is written — one log file, appended on every run
RUN_LOG_PATH = r'C:\path\to\your\signoff_runner.log'

# Lockfile path — prevents two instances running simultaneously
LOCKFILE_PATH = r'C:\path\to\your\signoff.lock'

# Request timeout — mirrors main script
REQUEST_TIMEOUT = 30

# Expected Log Source Types to validate per hostname.
# Uses the same fuzzy keyword matching as the main inventory script —
# 'Microsoft Security' matches 'Microsoft Windows Security Event Log'.
# When a hostname is queried, ALL returned log sources are checked and
# each expected type is marked FOUND or MISSING in the reply.
#
# If ALL types are found   → green banner, confirmed full coverage
# If ANY type is missing   → amber banner, shows found vs missing per type
# If host not in QRadar    → red banner, no type check performed
#
# Leave EMPTY to skip type validation entirely — script behaves exactly
# as before, returning the single best source result only.
EXPECTED_LS_TYPES_CHECK = []
# Example:
# EXPECTED_LS_TYPES_CHECK = ['Microsoft Security', 'Linux OS']

# Companion rules — certain log source types require another type to also
# be present to be considered fully covered.
#
# Key   = the primary type (must match an entry in EXPECTED_LS_TYPES_CHECK)
# Value = the companion type that must also exist for that host
#
# Uses the same fuzzy keyword matching as EXPECTED_LS_TYPES_CHECK.
# Types NOT listed here are self-sufficient (e.g. Linux OS alone is fine).
#
# Example: Microsoft Security requires WinCollect alongside it.
# If Microsoft Security found but WinCollect missing → amber warning row.
# If Linux OS found alone → green, no companion needed.
LS_COMPANION_RULES = {}
# Example:
# LS_COMPANION_RULES = {
#     'Microsoft Security': 'WinCollect',
# }

# ─── REPLY TEMPLATES ───────────────────────────────────────────────────────────
# The reply wording is built in _build_reply_html() below the config block.
# Three scenarios are handled automatically:
#   Active      — green banner, confirmed reporting, log source details shown
#   Inactive    — amber banner, found but not reporting, details + warning shown
#   Not Found   — red banner, not in QRadar inventory
#
# To change wording, edit the summary_line strings inside _build_reply_html().
# Placeholders available: hostname, actual_name, ls_type, last_seen, days_since_last_event

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

# Activity threshold for Active vs Inactive determination (days)
# Mirrors the same logic as the main inventory script
ACTIVITY_THRESHOLD_DAYS = 7

# Valid timestamp boundaries — mirrors main script
_MIN_TS = 0
_MAX_TS = 2147483647

# Global Log Source Types cache — same pattern as main script
LOG_SOURCE_TYPES_CACHE = {}


# ─── LOGGING ───────────────────────────────────────────────────────────────────
def _log(message):
    """
    Appends a timestamped line to the run log file and prints to console.
    """
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    try:
        with open(RUN_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception as e:
        print(f"⚠️  Could not write to log file: {e}")


# ─── LOCKFILE ──────────────────────────────────────────────────────────────────
def acquire_lock():
    """
    Writes a lockfile to prevent two instances running simultaneously.
    Returns True if the lock was acquired, False if another instance is running.
    """
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
    """Removes the lockfile on clean exit."""
    try:
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
    except Exception as e:
        _log(f"⚠️  Could not remove lockfile: {e}")


# ─── QRADAR CONNECTION ─────────────────────────────────────────────────────────
def test_qradar_connection():
    """
    Tests QRadar connection before processing any emails.
    Mirrors the same pre-flight check as the main inventory script.
    If this fails, no drafts are created — emails stay untouched for the next run.
    """
    _log("🔗 Testing QRadar connection...")
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/help/versions"

    try:
        resp = requests.get(
            endpoint,
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )
        if resp.status_code == 200:
            _log("✅ QRadar connection successful.")
            return True
        elif resp.status_code == 401:
            _log("❌ Authentication failed. Check username/password.")
            return False
        else:
            _log(f"⚠️  Unexpected response: {resp.status_code}")
            return False
    except Exception as e:
        _log(f"❌ Connection failed: {e}")
        return False


def fetch_log_source_types():
    """
    Pre-fetches Log Source Type ID → Name dictionary into memory.
    Exact same pattern as the main inventory script.
    """
    _log("📥 Fetching Log Source Types into cache...")
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources/log_source_management/log_source_types"

    try:
        resp = requests.get(
            endpoint,
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
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
    except Exception as e:
        _log(f"❌ Error fetching Log Source Types: {e}")


# ─── QRADAR QUERY — READ ONLY ──────────────────────────────────────────────────
def _empty_result():
    return {
        'status':               'Not Found',
        'actual_name':          'N/A',
        'ls_type':              'N/A',
        'last_seen':            'N/A',
        'activity_status':      'Not Found',
        'days_since_last_event': None,
    }


def _safe_timestamp(timestamp_ms):
    """
    Converts QRadar epoch ms timestamp to a readable string and activity status.
    Mirrors safe_timestamp_conversion from the main script exactly.
    Read-only — no side effects.
    """
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None

    try:
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)

        epoch_s = timestamp_ms / 1000.0 if timestamp_ms > 4102444800 else timestamp_ms

        if epoch_s <= _MIN_TS or epoch_s > _MAX_TS:
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None

        last_event_dt = datetime.fromtimestamp(epoch_s)
        last_seen     = last_event_dt.strftime('%Y-%m-%d %H:%M:%S')
        days_ago      = (datetime.now() - last_event_dt).days
        threshold_dt  = datetime.now() - timedelta(days=ACTIVITY_THRESHOLD_DAYS)

        activity = 'Active' if last_event_dt > threshold_dt else 'Inactive'
        return last_seen, activity, days_ago

    except Exception:
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def query_log_source_readonly(hostname):
    """
    STRICTLY READ-ONLY QRadar query for a given hostname.

    Only HTTP GET requests are used — no POST, PUT, PATCH or DELETE anywhere
    in this function. Nothing in QRadar is modified, created or deleted.

    Logic mirrors get_log_source_details from the main inventory script with
    the IP fallback removed — signoff emails always use hostnames, never IPs.

    Returns a result dict with status, name, type, last_seen, activity_status,
    and days_since_last_event.
    """
    clean_hostname = str(hostname).replace('"', '').replace("'", "").strip()
    endpoint       = (
        f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
        f"/log_source_management/log_sources"
    )

    try:
        resp = requests.get(
            endpoint,
            params={'filter': f'name ilike "%{clean_hostname}%"'},
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )

        if resp.status_code != 200:
            return {**_empty_result(), 'status': f'API Error {resp.status_code}'}

        ls_data = resp.json()
        if not ls_data:
            return _empty_result()

        # Pick best source: enabled first, then most recent last_event_time
        enabled  = [s for s in ls_data if s.get('enabled') is True]
        disabled = [s for s in ls_data if s.get('enabled') is False]
        enabled.sort(key=lambda x: x.get('last_event_time') or 0, reverse=True)
        disabled.sort(key=lambda x: x.get('last_event_time') or 0, reverse=True)
        found = enabled[0] if enabled else (disabled[0] if disabled else None)

        if not found:
            return _empty_result()

        type_id      = found.get('type_id')
        ls_type_name = LOG_SOURCE_TYPES_CACHE.get(type_id, f"Unknown Type ID: {type_id}")
        last_seen, activity_status, days_ago = _safe_timestamp(found.get('last_event_time'))

        return {
            'status':                'Found',
            'actual_name':           found.get('name', hostname),
            'ls_type':               ls_type_name,
            'last_seen':             last_seen,
            'activity_status':       activity_status,
            'days_since_last_event': days_ago,
        }

    except Exception as e:
        return {**_empty_result(), 'status': f'Error: {str(e)[:60]}'}


def query_all_log_sources_readonly(hostname):
    """
    STRICTLY READ-ONLY — fetches ALL log sources matching the hostname.
    Used when EXPECTED_LS_TYPES_CHECK is defined so every returned source
    can be checked against each expected type.

    Only HTTP GET. Nothing in QRadar is modified, created or deleted.

    Returns a dict:
      {
        'status': 'Found' | 'Not Found' | 'API Error ...',
        'sources': [
          {
            'name':      str,
            'ls_type':   str,
            'enabled':   bool,
            'last_seen': str,
            'activity':  str,
            'days_ago':  int | None
          },
          ...
        ]
      }
    """
    clean_hostname = str(hostname).replace('"', '').replace("'", "").strip()
    endpoint = (
        f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
        f"/log_source_management/log_sources"
    )

    try:
        resp = requests.get(
            endpoint,
            params={'filter': f'name ilike "%{clean_hostname}%"'},
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )

        if resp.status_code != 200:
            return {'status': f'API Error {resp.status_code}', 'sources': []}

        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', 'sources': []}

        sources = []
        for src in ls_data:
            type_id      = src.get('type_id')
            ls_type_name = LOG_SOURCE_TYPES_CACHE.get(type_id, f"Unknown Type ID: {type_id}")
            last_seen, activity, days_ago = _safe_timestamp(src.get('last_event_time'))
            sources.append({
                'name':     src.get('name', hostname),
                'ls_type':  ls_type_name,
                'enabled':  src.get('enabled', False),
                'last_seen': last_seen,
                'activity': activity,
                'days_ago': days_ago,
            })

        return {'status': 'Found', 'sources': sources}

    except Exception as e:
        return {'status': f'Error: {str(e)[:60]}', 'sources': []}


def validate_expected_types(all_sources_result):
    """
    Checks each entry in EXPECTED_LS_TYPES_CHECK against the full list of
    returned log sources using the same fuzzy keyword matching as the main
    inventory script.

    Returns a list of dicts — one per expected type — in order:
      {
        'expected':  str,           # the keyword string from config
        'found':     bool,          # True if at least one source matched
        'ls_type':   str | None,    # resolved QRadar type name if found
        'ls_name':   str | None,    # log source name if found
        'last_seen': str | None,
        'activity':  str | None,
        'days_ago':  int | None
      }

    If a type matches multiple sources, the best one is chosen:
    enabled first, then most recent last_event_time.
    """
    results  = []
    sources  = all_sources_result.get('sources', [])

    for expected_kw in EXPECTED_LS_TYPES_CHECK:
        exp_words = str(expected_kw).lower().split()

        # Collect all sources that fuzzy-match this expected keyword
        matched = [
            s for s in sources
            if all(w in str(s.get('ls_type', '')).lower() for w in exp_words)
        ]

        if not matched:
            results.append({
                'expected': expected_kw,
                'found':    False,
                'ls_type':  None,
                'ls_name':  None,
                'last_seen': None,
                'activity': None,
                'days_ago': None,
            })
            continue

        # Best match: enabled first, then most recent
        matched_enabled  = [s for s in matched if s.get('enabled')]
        matched_disabled = [s for s in matched if not s.get('enabled')]
        matched_enabled.sort(key=lambda x: x.get('days_ago') or 99999)
        matched_disabled.sort(key=lambda x: x.get('days_ago') or 99999)
        best = matched_enabled[0] if matched_enabled else matched_disabled[0]

        results.append({
            'expected': expected_kw,
            'found':    True,
            'ls_type':  best.get('ls_type'),
            'ls_name':  best.get('name'),
            'last_seen': best.get('last_seen'),
            'activity': best.get('activity'),
            'days_ago': best.get('days_ago'),
        })

def validate_expected_types(all_sources_result):
    """
    Checks each entry in EXPECTED_LS_TYPES_CHECK against the full list of
    returned log sources using the same fuzzy keyword matching as the main
    inventory script.

    Also checks LS_COMPANION_RULES — if a type has a defined companion,
    verifies the companion type exists in the source list too.

    Returns a list of dicts — one per expected type — in order:
      {
        'expected':          str,         # keyword from config
        'found':             bool,        # primary type found
        'companion_needed':  str | None,  # companion keyword if rule defined
        'companion_found':   bool,        # True if companion present or not needed
        'ls_type':           str | None,  # resolved QRadar type name if found
        'ls_name':           str | None,
        'last_seen':         str | None,
        'days_ago':          int | None
      }
    """
    results = []
    sources = all_sources_result.get('sources', [])

    for expected_kw in EXPECTED_LS_TYPES_CHECK:
        exp_words = str(expected_kw).lower().split()

        matched = [
            s for s in sources
            if all(w in str(s.get('ls_type', '')).lower() for w in exp_words)
        ]

        if not matched:
            results.append({
                'expected':         expected_kw,
                'found':            False,
                'companion_needed': LS_COMPANION_RULES.get(expected_kw),
                'companion_found':  False,
                'ls_type':          None,
                'ls_name':          None,
                'last_seen':        None,
                'days_ago':         None,
            })
            continue

        # Best match: enabled first, then most recent
        matched_enabled  = [s for s in matched if s.get('enabled')]
        matched_disabled = [s for s in matched if not s.get('enabled')]
        matched_enabled.sort(key=lambda x: x.get('days_ago') or 99999)
        matched_disabled.sort(key=lambda x: x.get('days_ago') or 99999)
        best = matched_enabled[0] if matched_enabled else matched_disabled[0]

        # Check companion rule if one is defined for this type
        companion_kw     = LS_COMPANION_RULES.get(expected_kw)
        companion_found  = True   # default: no rule = satisfied
        if companion_kw:
            comp_words      = str(companion_kw).lower().split()
            companion_found = any(
                all(w in str(s.get('ls_type', '')).lower() for w in comp_words)
                for s in sources
            )

        results.append({
            'expected':         expected_kw,
            'found':            True,
            'companion_needed': companion_kw,
            'companion_found':  companion_found,
            'ls_type':          best.get('ls_type'),
            'ls_name':          best.get('name'),
            'last_seen':        best.get('last_seen'),
            'days_ago':         best.get('days_ago'),
        })

    return results


# ─── SENDER VALIDATION ─────────────────────────────────────────────────────────
    """
    Checks sender against ALLOWED_SENDERS.
    Supports exact address matching and @domain wildcard matching.
    Case-insensitive on both sides.

    If ALLOWED_SENDERS is empty, all senders are allowed — all other
    checks (subject guards, DL in body, conversation deduplication) remain active.
    """
    if not ALLOWED_SENDERS:
        return True   # empty list = allow all senders, no restriction

    if not sender_address:
        return False

    sender_clean = sender_address.strip().lower()

    for entry in ALLOWED_SENDERS:
        entry_clean = entry.strip().lower()

        # @domain wildcard — sender must end with this domain
        if entry_clean.startswith('@'):
            if sender_clean.endswith(entry_clean):
                return True
        # Exact address match
        else:
            if sender_clean == entry_clean:
                return True

    return False


# ─── SUBJECT GUARDS ────────────────────────────────────────────────────────────
def passes_subject_guards(subject):
    """
    All four guards must pass before an email is considered for processing.
    Any failure returns False with a reason string for logging.

    Guards:
      1. Must contain the SUBJECT_SEPARATOR
      2. Must contain SUBJECT_KEYWORD (case-insensitive)
      3. Must NOT start with Re: / RE: / Fw: / FW: / Fwd: / FWD:
      4. Must NOT contain [Processed] — tag added to outgoing reply subjects
    """
    if not subject:
        return False, "empty subject"

    subject_stripped = subject.strip()
    subject_lower    = subject_stripped.lower()

    # Guard 3 — reply/forward prefixes
    reply_prefixes = ('re:', 'fw:', 'fwd:')
    if any(subject_lower.startswith(p) for p in reply_prefixes):
        return False, f"reply/forward prefix detected: '{subject_stripped[:30]}'"

    # Guard 4 — already processed tag
    if '[processed]' in subject_lower:
        return False, "subject contains [Processed] tag"

    # Guard 1 — separator must be present
    if SUBJECT_SEPARATOR not in subject_stripped:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found in subject"

    # Guard 2 — keyword must appear left of separator
    left_side = subject_stripped.split(SUBJECT_SEPARATOR)[0].strip().lower()
    if SUBJECT_KEYWORD.lower() not in left_side:
        return False, f"keyword '{SUBJECT_KEYWORD}' not found left of separator"

    return True, "ok"


def extract_hostname(subject):
    """
    Extracts the hostname from the right side of the subject separator.
    Returns None if the result is empty after stripping.
    """
    parts    = subject.split(SUBJECT_SEPARATOR, 1)
    hostname = parts[1].strip() if len(parts) > 1 else ''
    return hostname if hostname else None


# ─── CONVERSATION CHECK ────────────────────────────────────────────────────────
def is_already_handled(mail_item, sent_folder, drafts_folder):
    """
    Checks whether this email thread has already been handled by looking for
    any item sharing the same ConversationID in Sent Items or Drafts.

    This is the primary deduplication mechanism — no local storage needed.
    If a reply was already sent OR a draft already exists, returns True.

    If you delete a draft manually, the next run will create a fresh one.
    """
    conv_id = mail_item.ConversationID

    try:
        for item in sent_folder.Items:
            try:
                if item.ConversationID == conv_id:
                    return True, "reply already in Sent Items"
            except Exception:
                continue
    except Exception as e:
        _log(f"⚠️  Could not scan Sent Items: {e}")

    try:
        for item in drafts_folder.Items:
            try:
                if item.ConversationID == conv_id:
                    return True, "draft already exists in Drafts folder"
            except Exception:
                continue
    except Exception as e:
        _log(f"⚠️  Could not scan Drafts folder: {e}")

    return False, "not handled"


def _build_reply_html(hostname, qradar_result, type_validation=None):
    """
    Builds a clean, professional HTML reply body.

    Inactive state removed entirely — signoff emails are for new devices
    which will always be recent. Days since last event is shown in the
    table and speaks for itself.

    States:
      Not Found                → red banner
      All types found, all
        companions satisfied   → green banner
      Any type missing OR
        any companion missing  → amber banner
      Single-source mode
        (no type check)        → green if found, red if not found
    """
    status   = qradar_result.get('status') if qradar_result else 'Not Found'
    run_time = datetime.now().strftime('%d %B %Y, %H:%M')

    # ── NOT FOUND / ERROR ──
    if status != 'Found':
        return f"""
<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;color:#222;
             font-size:13px;line-height:1.6;margin:0;padding:0;">
  <div style="max-width:620px;padding:20px 0;">
    <p style="margin:0 0 16px 0;">Hi,</p>
    <div style="background:#c0392b;color:#fff;padding:10px 16px;
                border-radius:6px;font-size:13px;font-weight:600;
                margin-bottom:16px;">
      ✖&nbsp; Not Found in QRadar
    </div>
    <p style="margin:0 0 4px 0;">
      <b>{hostname}</b> was <b>not found</b> in the QRadar log source
      inventory. Please ensure the asset is onboarded and configured
      correctly in SIEM.
    </p>
    <p style="margin:20px 0 4px 0;color:#555;font-size:12px;">
      This is an automated response from the SIEM monitoring system.<br>
      Checked against QRadar on {run_time}.
    </p>
    <p style="margin:16px 0 0 0;">Regards,<br>
    <span style="font-weight:600;">SOC — Automated SIEM Check</span></p>
  </div>
</body>
</html>"""

    # ── TYPE VALIDATION MODE ──
    if type_validation is not None:
        any_missing          = any(not r['found'] for r in type_validation)
        any_companion_missing = any(
            r['found'] and r['companion_needed'] and not r['companion_found']
            for r in type_validation
        )

        if not any_missing and not any_companion_missing:
            banner_color = '#1a7a4a'
            banner_label = '✔&nbsp; All Expected Log Sources Confirmed'
            summary_line = (
                f"<b>{hostname}</b> has all expected log source types "
                f"present in SIEM."
            )
        else:
            banner_color  = '#c87800'
            found_count   = sum(1 for r in type_validation if r['found'])
            total_count   = len(type_validation)
            if any_missing:
                banner_label = '⚠&nbsp; Partial Coverage — Types Missing'
                summary_line = (
                    f"<b>{hostname}</b> — {found_count} of {total_count} expected "
                    f"log source types found. Missing types are highlighted below."
                )
            else:
                banner_label = '⚠&nbsp; Companion Log Source Missing'
                summary_line = (
                    f"<b>{hostname}</b> — all primary types found but a required "
                    f"companion log source is missing. See details below."
                )

        type_rows = ''
        for r in type_validation:
            days_str = (
                f"<span style='color:#888;font-size:11px;'>"
                f"({'Today' if r['days_ago'] == 0 else str(r['days_ago']) + ' days ago'})"
                f"</span>"
                if r['days_ago'] is not None else
                "<span style='color:#aaa;font-size:11px;'>(no events yet)</span>"
            )

            if not r['found']:
                # Primary type missing — red row
                type_rows += f"""
                <tr style="background:#fff5f5;">
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c0392b;font-weight:700;
                             width:24px;text-align:center;">✖</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c0392b;font-weight:600;">
                    {r['expected']}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c0392b;">—</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c0392b;">—</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c0392b;font-style:italic;">
                    Not found — please onboard this log source type.</td>
                </tr>"""

            elif r['companion_needed'] and not r['companion_found']:
                # Primary found, companion missing — amber row
                type_rows += f"""
                <tr style="background:#fffbf0;">
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c87800;font-weight:700;
                             width:24px;text-align:center;">⚠</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#333;font-weight:600;">
                    {r['expected']}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#555;">{r.get('ls_name','N/A')}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#555;">
                    {r.get('last_seen','N/A')} {days_str}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c87800;font-style:italic;">
                    Found — but <b>{r['companion_needed']}</b> companion missing.</td>
                </tr>"""

            else:
                # Primary found, companion satisfied or not needed — green row
                type_rows += f"""
                <tr style="background:#f0faf4;">
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#1a7a4a;font-weight:700;
                             width:24px;text-align:center;">✔</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#333;font-weight:600;">
                    {r['expected']}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#555;">{r.get('ls_name','N/A')}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#555;">
                    {r.get('last_seen','N/A')} {days_str}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#1a7a4a;font-weight:600;">
                    Confirmed</td>
                </tr>"""

        detail_block = f"""
        <table style="width:100%;max-width:640px;border-collapse:collapse;
                      margin-top:16px;border:1px solid #e0e0e0;
                      border-radius:6px;overflow:hidden;">
          <tr style="background:#f5f5f5;">
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;
                       width:24px;"></th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Expected Type</th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Log Source Name</th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Last Event</th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Result</th>
          </tr>
          {type_rows}
        </table>"""

    # ── SINGLE-SOURCE MODE (EXPECTED_LS_TYPES_CHECK empty) ──
    else:
        banner_color = '#1a7a4a'
        banner_label = '✔&nbsp; Reporting in SIEM'
        summary_line = f"<b>{hostname}</b> is confirmed present in SIEM."
        days_val     = qradar_result.get('days_since_last_event')
        days_display = (
            'Today' if days_val == 0
            else f"{days_val} days ago" if days_val is not None
            else 'N/A'
        )
        detail_block = f"""
        <table style="width:100%;max-width:480px;border-collapse:collapse;
                      margin-top:16px;border:1px solid #e0e0e0;
                      border-radius:6px;overflow:hidden;">
          <tr>
            <td style="padding:7px 12px;color:#555;font-size:13px;
                       border-bottom:1px solid #eee;width:160px;">
              Log Source Name</td>
            <td style="padding:7px 12px;font-size:13px;
                       border-bottom:1px solid #eee;font-weight:600;color:#222;">
              {qradar_result.get('actual_name','N/A')}</td>
          </tr>
          <tr>
            <td style="padding:7px 12px;color:#555;font-size:13px;
                       border-bottom:1px solid #eee;">Log Source Type</td>
            <td style="padding:7px 12px;font-size:13px;
                       border-bottom:1px solid #eee;color:#333;">
              {qradar_result.get('ls_type','N/A')}</td>
          </tr>
          <tr>
            <td style="padding:7px 12px;color:#555;font-size:13px;">
              Last Event</td>
            <td style="padding:7px 12px;font-size:13px;color:#333;">
              {qradar_result.get('last_seen','N/A')}
              &nbsp;<span style="color:#888;font-size:12px;">
                ({days_display})
              </span>
            </td>
          </tr>
        </table>"""

    return f"""
<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;color:#222;
             font-size:13px;line-height:1.6;margin:0;padding:0;">
  <div style="max-width:660px;padding:20px 0;">
    <p style="margin:0 0 16px 0;">Hi,</p>
    <div style="background:{banner_color};color:#fff;padding:10px 16px;
                border-radius:6px;font-size:13px;font-weight:600;
                margin-bottom:16px;letter-spacing:0.2px;">
      {banner_label}
    </div>
    <p style="margin:0 0 4px 0;">{summary_line}</p>
    {detail_block}
    <p style="margin:20px 0 4px 0;color:#555;font-size:12px;">
      This is an automated response from the SIEM monitoring system.<br>
      Checked against QRadar on {run_time}.
    </p>
    <p style="margin:16px 0 0 0;">Regards,<br>
    <span style="font-weight:600;">SOC — Automated SIEM Check</span></p>
  </div>
</body>
</html>"""


# ─── DRAFT BUILDER ─────────────────────────────────────────────────────────────
def build_reply_body(hostname, qradar_result):
    """
    Routes to the correct HTML builder based on whether EXPECTED_LS_TYPES_CHECK
    is populated in config.

    If EXPECTED_LS_TYPES_CHECK is empty:
      Uses qradar_result (single best source) — original behaviour unchanged.

    If EXPECTED_LS_TYPES_CHECK is populated:
      qradar_result is the all-sources result from query_all_log_sources_readonly.
      Runs validate_expected_types and passes the breakdown to _build_reply_html.
    """
    if not EXPECTED_LS_TYPES_CHECK:
        return _build_reply_html(hostname, qradar_result, type_validation=None)

    # Type validation mode — qradar_result is the all-sources dict here
    if qradar_result.get('status') != 'Found':
        return _build_reply_html(hostname, qradar_result, type_validation=None)

    validation = validate_expected_types(qradar_result)
    return _build_reply_html(hostname, qradar_result, type_validation=validation)


def create_draft_reply(mail_item, html_body, hostname):
    """
    Creates a draft reply to the original email and saves it silently to Drafts.
    Uses ReplyAll so all original recipients are included — change to Reply() if needed.

    Sets HTMLBody so the reply renders as formatted HTML in Outlook.
    The draft subject gets [Processed] prepended so subject guards catch it
    if it ever appears as an incoming item.

    THIS IS DRAFT ONLY — mail.Save() is called, NOT mail.Send().
    No email is sent from this script under any circumstance.
    """
    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"[Processed] {mail_item.Subject}"
        reply.Save()   # ← saves to Drafts, does NOT send
        _log(f"      ✅ Draft saved to Drafts folder for: {hostname}")
        return True

    except Exception as e:
        _log(f"      ❌ Failed to create draft for {hostname}: {e}")
        return False


# ─── OUTLOOK SETUP ─────────────────────────────────────────────────────────────
def get_outlook_folders():
    """
    Connects to the running Outlook instance and returns the three required folders.
    Exits cleanly if Outlook is not open — emails are left completely untouched.

    If SIGNOFF_FOLDER_NAME is set, opens that subfolder under Inbox.
    If SIGNOFF_FOLDER_NAME is None, falls back to the full Inbox.
    """
    try:
        outlook    = win32com.client.Dispatch('Outlook.Application')
        ns         = outlook.GetNamespace('MAPI')
        main_inbox = ns.GetDefaultFolder(6)   # 6 = Inbox
        drafts     = ns.GetDefaultFolder(16)  # 16 = Drafts
        sent       = ns.GetDefaultFolder(5)   # 5  = Sent Items

        if SIGNOFF_FOLDER_NAME:
            try:
                inbox = main_inbox.Folders[SIGNOFF_FOLDER_NAME]
                _log(f"📁 Monitoring subfolder: Inbox\\{SIGNOFF_FOLDER_NAME}")
            except Exception:
                _log(f"⚠️  Subfolder '{SIGNOFF_FOLDER_NAME}' not found — "
                     f"falling back to full Inbox. Create the Outlook rule first.")
                inbox = main_inbox
        else:
            inbox = main_inbox
            _log("📁 Monitoring: Full Inbox (no subfolder configured)")

        return inbox, drafts, sent

    except Exception as e:
        _log(f"❌ Could not connect to Outlook: {e}. Is Outlook open and logged in?")
        return None, None, None


# ─── BODY DL CHECK ────────────────────────────────────────────────────────────
def body_contains_dl(mail_item):
    """
    Checks whether the email body contains the TRIGGER_DL string.
    The body wording is completely irrelevant — only the DL presence matters.
    This means wording like 'please check', 'can you validate', 'kindly confirm'
    can change freely without breaking the script.

    Checks plain text body first. Falls back to HTML body if plain text is empty
    so rich-text / HTML-only emails are not missed.

    If TRIGGER_DL is set to '' in config, this check is skipped entirely
    and all sender-validated emails with matching subjects are processed.

    Case-insensitive throughout.
    """
    if not TRIGGER_DL.strip():
        return True   # DL check disabled in config

    dl_lower = TRIGGER_DL.strip().lower()

    try:
        # Primary: plain text body
        body = mail_item.Body or ''
        if dl_lower in body.strip().lower():
            return True

        # Fallback: HTML body — covers cases where .Body is empty in HTML-only emails
        html_body = mail_item.HTMLBody or ''
        if dl_lower in html_body.strip().lower():
            return True

        return False

    except Exception as e:
        _log(f"⚠️  Could not read email body for DL check: {e}")
        return False


# ─── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _log("=" * 60)
    _log("🚀 QRadar Signoff Auto-Draft starting...")
    _log(f"   Lookback   : {LOOKBACK_HOURS}h  |  Keyword: '{SUBJECT_KEYWORD}'")
    _log(f"   Separator  : '{SUBJECT_SEPARATOR}'  |  "
         f"Allowed senders: {'ALL' if not ALLOWED_SENDERS else len(ALLOWED_SENDERS)}")
    _log(f"   Trigger DL : '{TRIGGER_DL}' (must appear in email body)")
    _log(f"   Folder     : '{SIGNOFF_FOLDER_NAME or 'Full Inbox'}'")
    _log(f"   MODE       : DRAFT ONLY — nothing is sent automatically")

    # ── Lockfile ──
    if not acquire_lock():
        return

    try:
        # ── Outlook ──
        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        # ── QRadar ──
        if not test_qradar_connection():
            _log("❌ QRadar unreachable — exiting. All emails left untouched.")
            return

        fetch_log_source_types()

        # ── Scan inbox for emails within the lookback window ──
        cutoff_time = datetime.now() - timedelta(hours=LOOKBACK_HOURS)
        _log(f"\n📬 Scanning Inbox for emails since {cutoff_time.strftime('%Y-%m-%d %H:%M:%S')}...")

        processed = 0
        skipped   = 0
        drafted   = 0

        # Restrict filters at COM level — only emails within the lookback window
        # are returned before anything loads into Python.
        # %I = 12-hour clock (required by Outlook's Restrict filter), %p = AM/PM
        cutoff_str  = cutoff_time.strftime('%m/%d/%Y %I:%M %p')
        inbox_items = list(inbox.Items.Restrict(
            f"[ReceivedTime] >= '{cutoff_str}'"
        ))
        _log(f"   {len(inbox_items)} email(s) found in window.")

        for mail_item in inbox_items:

            # Skip non-mail items (calendar invites, meeting requests etc.)
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

            # ── Subject guards ──
            passed, reason = passes_subject_guards(subject)
            if not passed:
                skipped += 1
                _log(f"   ⏭️  SKIP (subject guard — {reason}): '{subject[:60]}'")
                continue

            # ── Sender guard ──
            try:
                sender = mail_item.SenderEmailAddress or ''
            except Exception:
                sender = ''

            # Skip your own address — prevents processing replies you sent
            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                _log(f"   ⏭️  SKIP (own address): '{subject[:60]}'")
                continue

            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"   ⏭️  SKIP (sender not in allowlist — {sender}): '{subject[:60]}'")
                continue

            # ── Body DL check — wording does not matter, only DL presence does ──
            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"   ⏭️  SKIP ('{TRIGGER_DL}' not found in body): '{subject[:60]}'")
                continue

            # ── Hostname extraction ──
            hostname = extract_hostname(subject)
            if not hostname:
                skipped += 1
                _log(f"   ⏭️  SKIP (empty hostname after separator): '{subject[:60]}'")
                continue

            _log(f"\n🔹 Candidate: '{subject[:70]}'")
            _log(f"      Sender  : {sender}")
            _log(f"      Hostname: {hostname}")

            # ── Conversation deduplication ──
            handled, handle_reason = is_already_handled(mail_item, sent, drafts)
            if handled:
                skipped += 1
                _log(f"      ⏭️  SKIP ({handle_reason})")
                continue

            # ── QRadar query — read only ──
            _log(f"      🔍 Querying QRadar...")
            if EXPECTED_LS_TYPES_CHECK:
                qradar_result = query_all_log_sources_readonly(hostname)
                _log(f"      📊 Status: {qradar_result['status']} | "
                     f"Sources returned: {len(qradar_result.get('sources', []))} | "
                     f"Checking {len(EXPECTED_LS_TYPES_CHECK)} expected type(s)")
            else:
                qradar_result = query_log_source_readonly(hostname)
                _log(f"      📊 Result: {qradar_result['status']} | "
                     f"Activity: {qradar_result.get('activity_status', 'N/A')} | "
                     f"Last seen: {qradar_result.get('last_seen', 'N/A')}")

            # ── Build and save draft ──
            body = build_reply_body(hostname, qradar_result)
            success = create_draft_reply(mail_item, body, hostname)

            if success:
                drafted += 1
            processed += 1

        # ── Summary ──
        _log(f"\n{'='*60}")
        _log(f"✅ Run complete.")
        _log(f"   Emails processed : {processed}")
        _log(f"   Drafts created   : {drafted}")
        _log(f"   Skipped          : {skipped}")
        _log(f"   Drafts are in your Drafts folder — review and send manually.")
        _log(f"{'='*60}\n")

    finally:
        # Always release lock even if an exception occurs mid-run
        release_lock()


if __name__ == '__main__':
    main()
