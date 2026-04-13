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

# Conversation deduplication windows:
# - normal signoff replies use hour-based deduplication
# - partial/not-found replies can be revalidated after this day window
NORMAL_SIGNOFF_DEDUP_HOURS = LOOKBACK_HOURS
PARTIAL_REVALIDATION_DAYS  = 14

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

# Smart OS-based log source type validation.
# The script detects the device OS from QRadar sources and validates only
# that OS group's required types — no cross-group false flags.
#
# Detection: the FIRST entry in each group's 'required' list is its
# OS signature. First group whose signature fuzzy-matches wins.
# Key order = detection priority (Windows checked before Linux here).
#
# Each entry in 'required' gets its own row in the reply table.
# All must be present AND reporting for a green result.
#
# Leave OS_TYPE_GROUPS = {} to disable type validation entirely and fall
# back to a simple found/not-found reply.
OS_TYPE_GROUPS = {
    'Windows': {
        'required': ['Microsoft Security', 'WinCollect'],
    },
    'Linux': {
        'required': ['Linux OS'],
    },
}

# Onboarding escalation — when a required log source type is missing OR found
# but has sent no events yet, this person is added to CC and mentioned in the
# reply body.
# Set ONBOARD_REQUEST_CC to '' to disable CC entirely.
# ONBOARD_REQUEST_NAME is how they appear in the email body e.g. '@xyz'
ONBOARD_REQUEST_CC   = 'onboarding-owner@yourorg.com'
ONBOARD_REQUEST_NAME = '@xyz'

# Recipient override for Partial / Not Found outcomes.
# When set, these drafts are routed ONLY to these recipients (plus trigger/onboard
# recipients below) instead of replying to the full original thread.
# Keep [] to use only TRIGGER_DL (To) and ONBOARD_REQUEST_CC (CC).
PARTIAL_NOT_FOUND_TO_RECIPIENTS = []
PARTIAL_NOT_FOUND_CC_RECIPIENTS = []

# ─── REPLY TEMPLATES ───────────────────────────────────────────────────────────
# The reply wording is built in _build_reply_html() below the config block.
# Four scenarios are handled automatically:
#   Active (all confirmed)  — green banner, all log source details shown
#   Silent (found, no events) — amber banner, found but zero events, CC escalated
#   Partial (types missing) — amber banner, missing types highlighted, CC escalated
#   Not Found               — red banner, not in QRadar inventory
#
# To change wording, edit the summary_line strings inside _build_reply_html().

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

# Activity threshold for Active vs Inactive determination (days)
ACTIVITY_THRESHOLD_DAYS = 7

# Valid timestamp boundaries — mirrors main script
_MIN_TS = 0
_MAX_TS = 2147483647

# Global Log Source Types cache
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
    Only HTTP GET requests — nothing in QRadar is modified, created or deleted.

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
    Used when OS_TYPE_GROUPS is configured so every returned source
    can be checked against each required type for the detected OS group.

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
                'name':      src.get('name', hostname),
                'ls_type':   ls_type_name,
                'enabled':   src.get('enabled', False),
                'last_seen': last_seen,
                'activity':  activity,
                'days_ago':  days_ago,
            })

        return {'status': 'Found', 'sources': sources}

    except Exception as e:
        return {'status': f'Error: {str(e)[:60]}', 'sources': []}


def validate_expected_types(all_sources_result, required_types):
    """
    Checks each entry in required_types against all returned QRadar sources.
    Uses fuzzy keyword matching — each word in the required type keyword must
    appear in the log source type name.

    Each required type gets its own result row — shown as a separate table row
    in the reply email regardless of found/not-found/silent state.

    Returns a list of dicts, one per required type:
      {
        'expected':  str,         # keyword from config
        'found':     bool,        # True if type exists in QRadar
        'ls_type':   str | None,  # resolved QRadar type name if found
        'ls_name':   str | None,
        'last_seen': str | None,
        'days_ago':  int | None   # None means found but zero events recorded
      }
    """
    results = []
    sources = all_sources_result.get('sources', [])

    for expected_kw in required_types:
        exp_words = str(expected_kw).lower().split()

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
                'days_ago': None,
            })
            continue

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
            'days_ago': best.get('days_ago'),   # None = found but no events yet
        })

    return results


def detect_os_group(sources):
    """
    Scans returned QRadar sources to determine which OS_TYPE_GROUPS group
    this device belongs to.

    Detection: the first entry in each group's 'required' list is its
    signature type. Fuzzy-matches against all returned sources.
    First matching group (by key order in OS_TYPE_GROUPS) wins.

    Returns (group_name, group_rules) or (None, None) if no group matched.
    If OS_TYPE_GROUPS is empty, returns (None, None) — legacy mode active.
    """
    if not OS_TYPE_GROUPS:
        return None, None

    for group_name, rules in OS_TYPE_GROUPS.items():
        required = rules.get('required', [])
        if not required:
            continue

        signature_words = str(required[0]).lower().split()
        signature_found = any(
            all(w in str(s.get('ls_type', '')).lower() for w in signature_words)
            for s in sources
        )

        if signature_found:
            return group_name, rules

    return None, None


# ─── SENDER VALIDATION ─────────────────────────────────────────────────────────
def is_sender_allowed(sender_address):
    """
    Checks sender against ALLOWED_SENDERS.
    Supports exact address matching and @domain wildcard matching.
    Case-insensitive on both sides.

    If ALLOWED_SENDERS is empty, all senders are allowed.
    """
    if not ALLOWED_SENDERS:
        return True

    if not sender_address:
        return False

    sender_clean = sender_address.strip().lower()

    for entry in ALLOWED_SENDERS:
        entry_clean = entry.strip().lower()

        if entry_clean.startswith('@'):
            if sender_clean.endswith(entry_clean):
                return True
        else:
            if sender_clean == entry_clean:
                return True

    return False


# ─── SUBJECT GUARDS ────────────────────────────────────────────────────────────
def passes_subject_guards(subject):
    """
    All four guards must pass before an email is considered for processing.

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

    reply_prefixes = ('re:', 'fw:', 'fwd:')
    if any(subject_lower.startswith(p) for p in reply_prefixes):
        return False, f"reply/forward prefix detected: '{subject_stripped[:30]}'"

    if '[processed]' in subject_lower:
        return False, "subject contains [Processed] tag"

    if SUBJECT_SEPARATOR not in subject_stripped:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found in subject"

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


def _normalize_recipients(values):
    """
    Normalizes recipient inputs into a de-duplicated list while preserving order.
    Accepts a list/tuple/set or a semicolon/comma-separated string.
    Non-string entries are ignored.
    """
    if values is None:
        return []

    if isinstance(values, str):
        raw_values = values.replace(',', ';').split(';')
    else:
        raw_values = []
        for val in values:
            if isinstance(val, str):
                raw_values.extend(val.replace(',', ';').split(';'))

    deduped = []
    seen = set()
    for value in raw_values:
        clean = value.strip()
        if not clean:
            continue
        key = clean.lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(clean)
    return deduped


def _get_partial_not_found_recipients():
    """
    Builds To/CC recipient lists for Partial / Not Found outcomes.
    """
    to_candidates = list(PARTIAL_NOT_FOUND_TO_RECIPIENTS or [])
    cc_candidates = list(PARTIAL_NOT_FOUND_CC_RECIPIENTS or [])

    if TRIGGER_DL.strip():
        to_candidates.append(TRIGGER_DL.strip())
    if ONBOARD_REQUEST_CC.strip():
        cc_candidates.append(ONBOARD_REQUEST_CC.strip())

    to_list = _normalize_recipients(to_candidates)
    cc_list = _normalize_recipients(cc_candidates)

    return to_list, cc_list


# ─── CONVERSATION CHECK ────────────────────────────────────────────────────────
def is_already_handled(mail_item, sent_folder, drafts_folder, outcome_type='normal'):
    """
    Checks whether this email thread has already been handled by looking for
    any item sharing the same ConversationID in Sent Items or Drafts.

    Deduplication is outcome-aware and trigger-aware:
      - normal outcome: checks recent Sent/Draft items within NORMAL_SIGNOFF_DEDUP_HOURS
      - partial/not-found outcome: checks recent Sent/Draft items within
        PARTIAL_REVALIDATION_DAYS (so older unresolved cases can be revalidated)
      - previously handled items only block if they are newer than (or equal to)
        the current trigger email's ReceivedTime; older replies in the same
        conversation do not block revalidation when a newer trigger arrives.
    """
    conv_id = mail_item.ConversationID
    now = datetime.now()
    try:
        trigger_received_time = mail_item.ReceivedTime
    except Exception:
        try:
            trigger_received_time = mail_item.CreationTime
        except Exception:
            trigger_received_time = None

    if outcome_type == 'partial_or_not_found':
        cutoff_time = now - timedelta(days=PARTIAL_REVALIDATION_DAYS)
        window_desc = f"{PARTIAL_REVALIDATION_DAYS} day revalidation window"
    else:
        cutoff_time = now - timedelta(hours=NORMAL_SIGNOFF_DEDUP_HOURS)
        window_desc = f"{NORMAL_SIGNOFF_DEDUP_HOURS} hour normal window"

    if trigger_received_time is None:
        trigger_received_time = cutoff_time
        _log("⚠️  Trigger email has no ReceivedTime/CreationTime; "
             "using dedup cutoff as fallback trigger time.")

    try:
        for item in sent_folder.Items:
            try:
                if item.ConversationID != conv_id:
                    continue
                sent_on = item.SentOn
                if not sent_on:
                    continue
                if trigger_received_time and sent_on < trigger_received_time:
                    continue
                if sent_on >= cutoff_time:
                    return True, f"reply already in Sent Items ({window_desc})"
            except Exception:
                continue
    except Exception as e:
        _log(f"⚠️  Could not scan Sent Items: {e}")

    try:
        for item in drafts_folder.Items:
            try:
                if item.ConversationID != conv_id:
                    continue
                created_on = item.CreationTime
                if not created_on:
                    continue
                if trigger_received_time and created_on < trigger_received_time:
                    continue
                if created_on >= cutoff_time:
                    return True, f"draft already exists in Drafts folder ({window_desc})"
            except Exception:
                continue
    except Exception as e:
        _log(f"⚠️  Could not scan Drafts folder: {e}")

    return False, "not handled"


# ─── HTML REPLY BUILDER ────────────────────────────────────────────────────────
def _build_reply_html(hostname, qradar_result, type_validation=None, os_group=None):
    """
    Builds the HTML reply body.

    Four states:
      Not Found / no sources     → red banner
      All types found + reporting → green banner, all rows green
      Any type found but silent  → amber banner, silent rows amber with CC name
      Any type missing entirely  → amber banner, missing rows red with CC name
    """
    status   = qradar_result.get('status') if qradar_result else 'Not Found'
    sources  = qradar_result.get('sources', []) if qradar_result else []
    run_time = datetime.now().strftime('%d %B %Y, %H:%M')

    # ── NOT FOUND ──────────────────────────────────────────────────────────────
    if status != 'Found' or not sources:
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
    <span style="font-weight:600;">Cyberdefence</span></p>
  </div>
</body>
</html>"""

    # ── TYPE VALIDATION MODE ───────────────────────────────────────────────────
    if type_validation is not None:
        os_label = f' ({os_group} device)' if os_group else ''

        # FIX: any_problem covers BOTH missing types AND found-but-silent sources
        any_missing   = any(not r['found'] for r in type_validation)
        any_no_events = any(r['found'] and r['days_ago'] is None for r in type_validation)
        any_problem   = any_missing or any_no_events

        if not any_problem:
            # All types found AND all have sent events — full green
            banner_color = '#1a7a4a'
            banner_label = '✔&nbsp; Confirmed Reporting on SIEM'
            summary_line = (
                f"<b>{hostname}</b> is reporting on our SIEM{os_label}. "
                f"All required log sources are present and active."
            )

        else:
            # Amber banner — distinguish wording by problem type
            banner_color = '#c87800'
            found_count  = sum(1 for r in type_validation if r['found'])
            total_count  = len(type_validation)

            if any_missing:
                banner_label = '⚠&nbsp; Partial — Log Sources Missing'
                summary_line = (
                    f"<b>{hostname}</b>{os_label} — {found_count} of {total_count} "
                    f"required log sources found on SIEM. "
                    f"Missing sources are highlighted below."
                )
            else:
                # All onboarded but at least one is completely silent
                banner_label = '⚠&nbsp; Partial — Log Sources Not Reporting'
                summary_line = (
                    f"<b>{hostname}</b>{os_label} — all {total_count} required log "
                    f"sources are onboarded but one or more have not sent any events "
                    f"yet. Please investigate."
                )

        # ── Build per-type table rows ──────────────────────────────────────────
        type_rows = ''
        for r in type_validation:
            days_str = (
                f"<span style='color:#888;font-size:11px;'>"
                f"({'Today' if r['days_ago'] == 0 else str(r['days_ago']) + ' days ago'})"
                f"</span>"
                if r['days_ago'] is not None else ''
            )

            if not r['found']:
                # ── RED ROW — type missing entirely ───────────────────────────
                onboard_note = (
                    f"{ONBOARD_REQUEST_NAME}, request you to onboard this "
                    f"log source on SIEM."
                    if ONBOARD_REQUEST_NAME.strip() else
                    "Not found — please onboard this log source."
                )
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
                    {onboard_note}</td>
                </tr>"""

            elif r['days_ago'] is None:
                # ── AMBER ROW — found in QRadar but zero events received yet ──
                #
                # FIX: This is the new third row state. Previously these rows
                # fell through to the green block and showed "Confirmed" despite
                # having no events. Now they render amber with an explicit
                # "no events received" note and the CC person called out by name.
                no_event_note = (
                    f"{ONBOARD_REQUEST_NAME}, no events received from this log "
                    f"source yet — please investigate."
                    if ONBOARD_REQUEST_NAME.strip() else
                    "No events received yet — please investigate."
                )
                type_rows += f"""
                <tr style="background:#fffbf0;">
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c87800;font-weight:700;
                             width:24px;text-align:center;">⚠</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c87800;font-weight:600;">
                    {r['expected']}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#555;">{r.get('ls_name', 'N/A')}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c87800;font-style:italic;">
                    No events recorded</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#c87800;font-style:italic;">
                    {no_event_note}</td>
                </tr>"""

            else:
                # ── GREEN ROW — found and actively reporting ───────────────────
                type_rows += f"""
                <tr style="background:#f0faf4;">
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#1a7a4a;font-weight:700;
                             width:24px;text-align:center;">✔</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#333;font-weight:600;">
                    {r['expected']}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#555;">{r.get('ls_name', 'N/A')}</td>
                  <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                             font-size:12px;color:#555;">
                    {r.get('last_seen', 'N/A')} {days_str}</td>
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
                       text-align:left;border-bottom:2px solid #ddd;width:24px;"></th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Log Source Type</th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Log Source Name</th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Last Event</th>
            <th style="padding:7px 12px;font-size:11px;color:#777;font-weight:600;
                       text-align:left;border-bottom:2px solid #ddd;">
              Status</th>
          </tr>
          {type_rows}
        </table>"""

    # ── SIMPLE FOUND MODE (OS_TYPE_GROUPS empty or OS undetected) ─────────────
    else:
        banner_color = '#1a7a4a'
        banner_label = '✔&nbsp; Confirmed Reporting on SIEM'
        summary_line = f"<b>{hostname}</b> is reporting on our SIEM."

        best_src = None
        if sources:
            enabled  = [s for s in sources if s.get('enabled')]
            disabled = [s for s in sources if not s.get('enabled')]
            enabled.sort(key=lambda x: x.get('days_ago') or 99999)
            best_src = enabled[0] if enabled else (disabled[0] if disabled else None)

        if best_src:
            days_val     = best_src.get('days_ago')
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
                  {best_src.get('name', 'N/A')}</td>
              </tr>
              <tr>
                <td style="padding:7px 12px;color:#555;font-size:13px;
                           border-bottom:1px solid #eee;">Log Source Type</td>
                <td style="padding:7px 12px;font-size:13px;
                           border-bottom:1px solid #eee;color:#333;">
                  {best_src.get('ls_type', 'N/A')}</td>
              </tr>
              <tr>
                <td style="padding:7px 12px;color:#555;font-size:13px;">
                  Last Event</td>
                <td style="padding:7px 12px;font-size:13px;color:#333;">
                  {best_src.get('last_seen', 'N/A')}
                  &nbsp;<span style="color:#888;font-size:12px;">
                    ({days_display})
                  </span>
                </td>
              </tr>
            </table>"""
        else:
            detail_block = ''

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
    <span style="font-weight:600;">Cyberdefence</span></p>
  </div>
</body>
</html>"""


# ─── DRAFT BUILDER ─────────────────────────────────────────────────────────────
def build_reply_body(hostname, qradar_result):
    """
    Routes to the correct HTML builder.

    If qradar_result has no sources at all → Not Found reply.
    If OS_TYPE_GROUPS is populated → detect OS, validate required types only.
    If OS detection fails (neither signature matched but sources exist) →
      show what was found with an unrecognised device type warning.
    If OS_TYPE_GROUPS = {} → simple found reply, no type validation.

    Returns (html_body, needs_cc, outcome_type) tuple.
    outcome_type is:
      - 'normal'
      - 'partial_or_not_found'

    FIX: needs_cc is True when ANY required type is either:
      - missing entirely (not r['found']), OR
      - found but has sent zero events (r['found'] and r['days_ago'] is None)
    Previously only the missing case triggered needs_cc, so found-but-silent
    sources never added the CC recipient or escalation wording.
    """
    status  = qradar_result.get('status')
    sources = qradar_result.get('sources', [])

    # Genuinely not found or API error
    if status != 'Found' or not sources:
        return _build_reply_html(hostname, qradar_result,
                                 type_validation=None, os_group=None), False, 'partial_or_not_found'

    # No type validation configured — simple found reply
    if not OS_TYPE_GROUPS:
        return _build_reply_html(hostname, qradar_result,
                                 type_validation=None, os_group=None), False, 'normal'

    # Detect OS group
    group_name, group_rules = detect_os_group(sources)

    if group_name is None:
        # Sources exist but no OS signature matched — show raw sources,
        # no type validation, to avoid a false Not Found.
        _log(f"      ⚠️  OS group undetected for {hostname} — "
             f"showing raw sources, no type validation applied.")
        return _build_reply_html(hostname, qradar_result,
                                 type_validation=None, os_group=None), False, 'normal'

    validation = validate_expected_types(
        qradar_result,
        required_types=group_rules.get('required', []),
    )

    # FIX: CC is needed for missing types OR for found-but-silent types
    needs_cc = any(
        not r['found'] or (r['found'] and r['days_ago'] is None)
        for r in validation
    )

    html = _build_reply_html(
        hostname, qradar_result,
        type_validation=validation,
        os_group=group_name
    )
    outcome_type = 'partial_or_not_found' if needs_cc else 'normal'
    return html, needs_cc, outcome_type


def create_draft_reply(mail_item, html_body, hostname, needs_cc=False, outcome_type='normal'):
    """
    Creates a draft reply to the original email and saves it silently to Drafts.
    Uses ReplyAll so all original recipients are included.

    needs_cc: if True and ONBOARD_REQUEST_CC is set, adds that address to CC.
              Triggered when types are missing OR found but sending no events.
    outcome_type: if 'partial_or_not_found', recipients are restricted to the
                  configured escalation distribution lists.

    THIS IS DRAFT ONLY — mail.Save() is called, NOT mail.Send().
    No email is sent from this script under any circumstance.
    """
    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"[Processed] {mail_item.Subject}"

        if outcome_type == 'partial_or_not_found':
            to_list, cc_list = _get_partial_not_found_recipients()
            original_cc = reply.CC or ''

            if to_list:
                reply.To = '; '.join(to_list)
            else:
                _log("      ⚠️  Partial/Not Found recipient override has no TO entries; "
                     "keeping original ReplyAll recipients.")

            if cc_list:
                reply.CC = '; '.join(cc_list)
            elif to_list:
                reply.CC = ''
            else:
                reply.CC = original_cc

            _log(f"      📬 Recipient override applied for Partial/Not Found. "
                 f"To={reply.To} | CC={reply.CC}")

        elif needs_cc and ONBOARD_REQUEST_CC.strip():
            existing_cc  = reply.CC or ''
            reply.CC     = (
                f"{existing_cc}; {ONBOARD_REQUEST_CC}".strip('; ')
                if existing_cc else ONBOARD_REQUEST_CC
            )
            _log(f"      📧 CC added: {ONBOARD_REQUEST_CC}")

        reply.Save()
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

    Checks plain text body first. Falls back to HTML body if plain text is empty.
    Case-insensitive. If TRIGGER_DL = '' in config, this check is skipped.
    """
    if not TRIGGER_DL.strip():
        return True

    dl_lower = TRIGGER_DL.strip().lower()

    try:
        body = mail_item.Body or ''
        if dl_lower in body.strip().lower():
            return True

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
    _log(f"   Dedup      : normal={NORMAL_SIGNOFF_DEDUP_HOURS}h | "
         f"partial/not-found={PARTIAL_REVALIDATION_DAYS}d")
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
        _log(f"\n📬 Scanning Inbox for emails since "
             f"{cutoff_time.strftime('%Y-%m-%d %H:%M:%S')}...")

        processed = 0
        skipped   = 0
        drafted   = 0

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

            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                _log(f"   ⏭️  SKIP (own address): '{subject[:60]}'")
                continue

            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"   ⏭️  SKIP (sender not in allowlist — {sender}): "
                     f"'{subject[:60]}'")
                continue

            # ── Body DL check ──
            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"   ⏭️  SKIP ('{TRIGGER_DL}' not found in body): "
                     f"'{subject[:60]}'")
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

            # ── QRadar query — read only ──
            _log(f"      🔍 Querying QRadar...")
            if OS_TYPE_GROUPS:
                qradar_result = query_all_log_sources_readonly(hostname)
                _log(f"      📊 Status: {qradar_result['status']} | "
                     f"Sources returned: {len(qradar_result.get('sources', []))}")
            else:
                qradar_result = query_log_source_readonly(hostname)
                _log(f"      📊 Result: {qradar_result['status']} | "
                     f"Last seen: {qradar_result.get('last_seen', 'N/A')}")

            # ── Build and save draft ──
            body, needs_cc, outcome_type = build_reply_body(hostname, qradar_result)

            # ── Conversation deduplication (outcome-aware window) ──
            handled, handle_reason = is_already_handled(
                mail_item, sent, drafts, outcome_type=outcome_type
            )
            if handled:
                skipped += 1
                _log(f"      ⏭️  SKIP ({handle_reason})")
                continue

            success = create_draft_reply(mail_item, body, hostname,
                                         needs_cc=needs_cc,
                                         outcome_type=outcome_type)

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
        release_lock()


if __name__ == '__main__':
    main()
