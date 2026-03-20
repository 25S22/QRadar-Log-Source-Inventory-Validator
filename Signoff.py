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
ALLOWED_SENDERS = [
    'analyst1@yourorg.com',
    'analyst2@yourorg.com',
    'youremail@yourorg.com',       # add your own for testing
    '@soc.yourorg.com',            # whole domain example
]

# Your own reply address — used to guard against processing your own sent items
# that may appear in the inbox (e.g. on shared mailboxes)
YOUR_EMAIL_ADDRESS = 'youremail@yourorg.com'

# DL string that must appear somewhere in the email body for the script to act.
# The body wording does not matter — only this string being present matters.
# Example: '@SOC-DL@yourorg.com' or just 'SOC-Team' — whatever your DL looks like.
# Case-insensitive match. Set to '' to disable this check entirely.
TRIGGER_DL = '@SOC-DL@yourorg.com'

# Outlook folder names
INBOX_FOLDER_NAME  = 'Inbox'    # folder to monitor
DRAFTS_FOLDER_NAME = 'Drafts'   # where drafts are saved
SENT_FOLDER_NAME   = 'Sent Items'

# Where the run log is written — one log file, appended on every run
RUN_LOG_PATH = r'C:\path\to\your\signoff_runner.log'

# Lockfile path — prevents two instances running simultaneously
LOCKFILE_PATH = r'C:\path\to\your\signoff.lock'

# Request timeout — mirrors main script
REQUEST_TIMEOUT = 30

# ─── REPLY TEMPLATES ───────────────────────────────────────────────────────────
# Edit the wording to match your organisation's language exactly.
# Available placeholders for all three templates:
#   {hostname}   — the hostname extracted from the subject
#   {last_seen}  — formatted datetime of last event (Active/Inactive only)
#   {days_ago}   — number of days since last event (Active/Inactive only)
#   {ls_type}    — QRadar Log Source Type string (Active/Inactive only)
#   {ls_name}    — QRadar actual log source name as it appears in QRadar

REPLY_ACTIVE = (
    "Hi,\n\n"
    "{hostname} is confirmed reporting in SIEM.\n\n"
    "Log Source Name : {ls_name}\n"
    "Log Source Type : {ls_type}\n"
    "Last Event      : {last_seen} ({days_ago} days ago)\n\n"
    "Regards,\n"
    "Automated SOC Response"
)

REPLY_INACTIVE = (
    "Hi,\n\n"
    "{hostname} was found in SIEM but has not reported recently.\n\n"
    "Log Source Name : {ls_name}\n"
    "Log Source Type : {ls_type}\n"
    "Last Event      : {last_seen} ({days_ago} days ago)\n\n"
    "Please investigate why this source has gone silent.\n\n"
    "Regards,\n"
    "Automated SOC Response"
)

REPLY_NOT_FOUND = (
    "Hi,\n\n"
    "{hostname} was not found in the QRadar log source inventory.\n\n"
    "Please ensure the asset is onboarded and configured correctly in SIEM.\n\n"
    "Regards,\n"
    "Automated SOC Response"
)

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


# ─── SENDER VALIDATION ─────────────────────────────────────────────────────────
def is_sender_allowed(sender_address):
    """
    Checks sender against ALLOWED_SENDERS.
    Supports exact address matching and @domain wildcard matching.
    Case-insensitive on both sides.
    """
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


# ─── DRAFT BUILDER ─────────────────────────────────────────────────────────────
def build_reply_body(hostname, qradar_result):
    """
    Selects the correct template based on QRadar result and fills in placeholders.
    Returns the plain text body string.
    All three templates are defined in config — wording is fully yours to control.
    """
    status = qradar_result.get('status')

    if status != 'Found':
        return REPLY_NOT_FOUND.format(hostname=hostname)

    activity = qradar_result.get('activity_status', '')
    data     = {
        'hostname': hostname,
        'ls_name':  qradar_result.get('actual_name', 'N/A'),
        'ls_type':  qradar_result.get('ls_type', 'N/A'),
        'last_seen': qradar_result.get('last_seen', 'N/A'),
        'days_ago':  qradar_result.get('days_since_last_event', 'N/A'),
    }

    if activity == 'Active':
        return REPLY_ACTIVE.format(**data)
    else:
        return REPLY_INACTIVE.format(**data)


def create_draft_reply(mail_item, body_text, hostname):
    """
    Creates a draft reply to the original email and saves it silently to Drafts.
    Uses ReplyAll so all original recipients are included — change to Reply() if needed.

    The draft subject gets [Processed] prepended so the conversation check and
    subject guards both catch it if it ever appears as an incoming item.

    THIS IS DRAFT ONLY — mail.Save() is called, NOT mail.Send().
    No email is sent from this script under any circumstance.
    """
    try:
        reply         = mail_item.ReplyAll()
        reply.Body    = body_text
        reply.Subject = f"[Processed] {mail_item.Subject}"
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
    """
    try:
        outlook  = win32com.client.Dispatch('Outlook.Application')
        ns       = outlook.GetNamespace('MAPI')
        inbox    = ns.GetDefaultFolder(6)   # 6 = Inbox
        drafts   = ns.GetDefaultFolder(16)  # 16 = Drafts
        sent     = ns.GetDefaultFolder(5)   # 5  = Sent Items
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
    _log(f"   Separator  : '{SUBJECT_SEPARATOR}'  |  Allowed senders: {len(ALLOWED_SENDERS)}")
    _log(f"   Trigger DL : '{TRIGGER_DL}' (must appear in email body)")
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

        # Collect items first to avoid COM collection mutation during iteration
        inbox_items = list(inbox.Items)

        for mail_item in inbox_items:

            # Skip non-mail items (calendar invites etc.)
            try:
                if mail_item.Class != 43:   # 43 = olMail
                    continue
            except Exception:
                continue

            # ── Time filter ──
            try:
                received = mail_item.ReceivedTime
                # win32com returns a timezone-aware datetime; strip tz for comparison
                received_naive = received.replace(tzinfo=None)
                if received_naive < cutoff_time:
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
