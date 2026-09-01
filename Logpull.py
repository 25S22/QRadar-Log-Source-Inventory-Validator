"""
QRadar EKS Application Log Validator
=====================================
Answers a narrower, more common SOC question than "is anything logging
from EKS": for ONE specific application running on EKS, are the log
events we expect (logins, starts, errors, whatever the on-call runbook
cares about) actually landing in QRadar?

The scope is resolved in three narrowing stages, in this order:

  1. WHICH LOG SOURCES  → every ENABLED log source whose name matches the
     configured EKS pattern, resolved dynamically against Log Source
     Management (not hardcoded IDs, since nodes/log sources rotate).
  2. WHICH APPLICATION  → asked interactively at runtime rather than baked
     into the config. The EKS scope above is shared across many
     applications running on the same cluster — hardcoding one app name
     in CONFIG would mean editing the file every time you wanted to check
     a different service. The payload is filtered to only events
     containing this application name.
  3. WHICH EVENTS       → for each configured validation case (a name plus
     a list of keywords like "Login", "Start", "OutOfMemory"), payload is
     matched against those keywords, case-insensitively.

Each validation case becomes its own sheet in a single Excel workbook, so
you get one file per run that shows, side by side, whether every expected
event type showed up for that application in the configured time window.

Zero AQL fallback gaps: every search has a primary path (multi-keyword
ILIKE, UTF8-decoded) and a backup path (per-keyword decomposition without
the UTF8() wrapper) — see the note below.

    GET  /api/system/about
    GET  /api/config/event_sources/log_source_management/log_sources
    POST /api/ariel/searches
    GET  /api/ariel/searches/{id}
    GET  /api/ariel/searches/{id}/results

Output → one multi-sheet .xlsx workbook (xlsxwriter), one sheet per
         validation case, professionally formatted for readability.

A NOTE ON THE APPLICATION-NAME PROMPT
---------------------------------------
This is intentionally NOT a CONFIG value. The EKS log source scope in
CONFIG is meant to be resolved once and reused for many different
applications sharing that cluster. Prompting for the app name at run time
means the same script serves every "did app X's logs show up" question
without a file edit. The value is escaped exactly like any other keyword
before being spliced into AQL — see `_escape()`.

A NOTE ON THE FALLBACK PATH
------------------------------
Some QRadar builds / custom parsers reject large OR'd ILIKE expressions or
choke on UTF8() wrapping for certain payload encodings. Rather than let
that fail the entire validation case, each case retries with backoff and
then, if it's still failing, decomposes into one simplified search per
keyword (plain `payload` instead of `UTF8(payload)`) and merges the
results. A case only shows as failed in the report if BOTH paths fail.
"""

import os
import sys
import time
import logging

import requests
import urllib3
import xlsxwriter

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION  ←  edit this block only
# ══════════════════════════════════════════════════════════════════════════════

# ── Connection ────────────────────────────────────────────────────────────────
QRADAR_HOST     = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL      = False          # locally hosted / self-signed → safe to leave False

API_VERSION     = '20.0'         # match your QRadar Console's supported REST API version

# ── Dynamic scope resolution ───────────────────────────────────────────────────
# Stage 1 of scope narrowing — every enabled log source whose NAME contains
# this substring (case-insensitive). The application-name filter (stage 2)
# is asked interactively at runtime, not configured here — see the module
# docstring for why.
LOG_SOURCE_NAME_PATTERN = 'eks'

# Extra hostnames / IPs folded into scope alongside matched log sources.
ADDITIONAL_HOSTS = [
    # '10.20.30.40',
    # 'eks-worker-node-07',
]

# ── Search parameters ──────────────────────────────────────────────────────────
TIME_RANGE_DAYS          = 7
RECORD_LIMIT_PER_SEARCH  = 500
SEARCH_POLL_INTERVAL_SEC = 3
SEARCH_TIMEOUT_SEC       = 300

# ── API / networking / retry ───────────────────────────────────────────────────
REQUEST_TIMEOUT        = 30     # seconds per HTTP call
MAX_RETRIES            = 3      # attempts before giving up (exponential backoff)
RETRY_DELAY_BASE       = 1.5    # seconds — waits: 1.5s → 3s → 6s
MAX_RETRIES_PER_SEARCH = 2      # per-case retry count before falling back to per-keyword decomposition
LOG_SOURCE_PAGE_SIZE   = 100    # QRadar Range header page size for log source listing
MAX_PAGES              = 100    # safety cap on pagination loops

# ── Validation test cases: name → keywords (OR'd, case-insensitive) ──────────
# Applied AFTER the application-name filter — these are the specific event
# types you expect to see for that application (e.g. Login, Start, errors).
VALIDATION_CASES = {
    'Application Start':    ['Start', 'Started', 'Startup complete'],
    'Login Events':         ['Login', 'Logged in', 'Authentication succeeded'],
    'Login Failures':       ['Login failed', 'Unauthorized', '403'],
    'Error / Crash Events': ['Exception', 'FATAL', 'panic', 'OOMKilled'],
}

# ── Output ──────────────────────────────────────────────────────────────────
OUTPUT_FILE = 'qradar_eks_app_validation_report.xlsx'

# ══════════════════════════════════════════════════════════════════════════════
#  END CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)
logger = logging.getLogger(__name__)


# ─── AUTH HELPERS ───────────────────────────────────────────────────────────────

def _validate_auth_config():
    """Catches the most common copy-paste mistake — running the script
    against the placeholder credentials — with a clear message instead of
    a confusing 401 three steps later."""
    placeholders = {
        'QRADAR_HOST': 'https://your-qradar-host',
        'QRADAR_USERNAME': 'your-username',
        'QRADAR_PASSWORD': 'your-password',
    }
    still_default = [name for name, val in placeholders.items()
                      if globals()[name] == val]
    if still_default:
        raise RuntimeError(
            f"CONFIG still has placeholder value(s) for: {', '.join(still_default)}. "
            "Edit the CONFIGURATION block at the top of the script with your real "
            "QRadar host and credentials before running."
        )


# ─── SHARED HTTP HELPERS ────────────────────────────────────────────────────────

def _build_url(path):
    host = QRADAR_HOST if QRADAR_HOST.startswith('http') else f'https://{QRADAR_HOST}'
    return f"{host.rstrip('/')}{path}"


def _http_request(method, path, params=None, headers=None, label='request', timeout=REQUEST_TIMEOUT):
    """
    Generic HTTP call with exponential-backoff retry on timeout / connection
    errors, and clean propagation of auth / HTTP errors — so one flaky
    network blip mid-run doesn't take down the whole validation pass.
    """
    url = _build_url(path)
    hdrs = {'Accept': 'application/json', 'Version': API_VERSION}
    if headers:
        hdrs.update(headers)

    last_err = None
    for attempt in range(MAX_RETRIES):
        if attempt > 0:
            wait = RETRY_DELAY_BASE * (2 ** (attempt - 1))
            logger.warning("Retry %d/%d for '%s' — waiting %.1fs after %s",
                           attempt, MAX_RETRIES - 1, label, wait, type(last_err).__name__)
            time.sleep(wait)
        try:
            resp = requests.request(
                method, url, params=params, headers=hdrs,
                auth=(QRADAR_USERNAME, QRADAR_PASSWORD), verify=VERIFY_SSL, timeout=timeout,
            )
            if resp.status_code in (200, 201, 202, 206):
                return resp
            if resp.status_code == 401:
                raise RuntimeError("401 Unauthorized — invalid SEC token or credentials.")
            if resp.status_code == 403:
                raise RuntimeError("403 Forbidden — credentials valid but lack required role/capability.")
            last_err = RuntimeError(f"HTTP {resp.status_code}: {resp.text[:300]}")
            logger.warning("HTTP %d for '%s' (attempt %d/%d)", resp.status_code, label, attempt + 1, MAX_RETRIES)
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_err = exc
            logger.warning("%s on attempt %d for '%s'", type(exc).__name__, attempt + 1, label)
        except RuntimeError:
            raise
        except Exception:
            logger.error("Non-retriable error for '%s'", label, exc_info=True)
            raise

    raise RuntimeError(f"All {MAX_RETRIES} attempts failed for '{label}': {last_err}")


def _api_get_page(path, range_start, range_end, params=None, label='request'):
    """Fetches one page of a QRadar list endpoint via an explicit Range header."""
    resp = _http_request('GET', path, params=params,
                          headers={'Range': f'items={range_start}-{range_end}'}, label=label)
    total = None
    cr = resp.headers.get('Content-Range', '')
    if cr:
        try:
            total = int(cr.split('/')[-1].strip())
        except Exception:
            pass
    return resp.json(), total


def _api_get_all(path, params=None, label='request', page_size=LOG_SOURCE_PAGE_SIZE, max_pages=MAX_PAGES):
    """Loops _api_get_page until every item is fetched — avoids silently
    truncating a large log source list to a single page."""
    all_items, start = [], 0
    for _ in range(max_pages):
        end = start + page_size - 1
        items, total = _api_get_page(path, start, end, params=params,
                                      label=f'{label} (items {start}-{end})')
        if not items:
            break
        all_items.extend(items)
        if total is not None and len(all_items) >= total:
            break
        if len(items) < page_size:
            break
        start += page_size
    else:
        logger.warning("Hit max_pages=%d for '%s' — data may be incomplete.", max_pages, label)
    return all_items


# ─── CONNECTION TEST ────────────────────────────────────────────────────────────

def test_connection():
    print("🔗 Testing QRadar connection...")
    try:
        resp = _http_request('GET', '/api/system/about', label='connection test')
        info = resp.json()
        print(f"✅ Connected — release: {info.get('release_name', 'unknown')} "
              f"| build: {info.get('build_version', 'unknown')}")
        return True
    except RuntimeError as e:
        print(f"❌ {e}")
        return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


# ─── SCOPE RESOLUTION ───────────────────────────────────────────────────────────

def fetch_active_eks_log_sources():
    """Stage 1 of scope narrowing: every enabled log source matching the
    configured EKS name pattern, fully paginated."""
    print(f"📥 Resolving active log sources matching pattern '{LOG_SOURCE_NAME_PATTERN}'...")
    try:
        data = _api_get_all(
            '/api/config/event_sources/log_source_management/log_sources',
            params={'filter': 'enabled=true', 'fields': 'id,name,type_id'},
            label='log sources',
        )
    except Exception as e:
        print(f"   ❌ Failed to fetch log sources: {e}")
        return []
    matched = [s for s in data if LOG_SOURCE_NAME_PATTERN.lower() in s.get('name', '').lower()]
    print(f"   ✅ {len(matched)} active log source(s) matched (of {len(data)} enabled total).")
    return matched


def prompt_application_name():
    """
    Stage 2 of scope narrowing, asked interactively — see the module
    docstring's note on why this isn't a CONFIG value.
    """
    print("\n" + "─" * 62)
    app_name = input("🔎 Enter the application name to filter payloads by (e.g. 'checkout-service'): ").strip()
    while not app_name:
        app_name = input("   Application name cannot be empty — please enter a value: ").strip()
    print(f"   Using application filter: '{app_name}'")
    print("─" * 62 + "\n")
    return app_name


# ─── AQL CONSTRUCTION ───────────────────────────────────────────────────────────

def _escape(value):
    """Escape single quotes so keywords/hosts/app-names can't break out of
    the AQL string literal."""
    return str(value).replace("'", "''")


def build_scope_clause(log_source_ids, hosts):
    clauses = []
    if log_source_ids:
        ids = ",".join(str(i) for i in log_source_ids)
        clauses.append(f"logsourceid IN ({ids})")
    for h in hosts or []:
        h_esc = _escape(h)
        clauses.append(f"sourceip = '{h_esc}'")
        clauses.append(f"UTF8(LOGSOURCENAME(logsourceid)) ILIKE '%{h_esc}%'")
    if not clauses:
        raise ValueError("No log sources matched and no additional hosts configured — empty scope.")
    return "(" + " OR ".join(clauses) + ")"


def build_app_clause(app_name, payload_field='UTF8(payload)'):
    """Stage 2 filter — payload must contain the application name."""
    return f"{payload_field} ILIKE '%{_escape(app_name)}%'"


def build_keyword_clause(keywords, payload_field='UTF8(payload)'):
    """Stage 3 filter — payload must contain at least one of the case's keywords."""
    parts = [f"{payload_field} ILIKE '%{_escape(kw)}%'" for kw in keywords]
    return "(" + " OR ".join(parts) + ")"


def build_aql(scope_clause, app_clause, keyword_clause, days, limit):
    fields = (
        "DATEFORMAT(devicetime, 'yyyy-MM-dd HH:mm:ss') AS \"Time\", "
        "sourceip AS \"Source\", "
        "LOGSOURCENAME(logsourceid) AS \"Log Source\", "
        "UTF8(payload) AS \"Raw Payload\""
    )
    return (
        f"SELECT {fields} FROM events "
        f"WHERE {scope_clause} AND {app_clause} AND {keyword_clause} "
        f"LAST {days} DAYS LIMIT {limit}"
    )


# ─── ARIEL SEARCH EXECUTION ─────────────────────────────────────────────────────

def submit_search(aql):
    resp = _http_request('POST', '/api/ariel/searches', params={'query_expression': aql}, label='submit search')
    return resp.json()['search_id']


def wait_for_search(search_id):
    deadline = time.time() + SEARCH_TIMEOUT_SEC
    while time.time() < deadline:
        resp = _http_request('GET', f'/api/ariel/searches/{search_id}', label=f'poll search {search_id}')
        status = resp.json()
        state = status.get('status')
        if state == 'COMPLETED':
            return status
        if state in ('CANCELED', 'ERROR'):
            raise RuntimeError(f"Search {search_id} ended with status {state}: {status}")
        time.sleep(SEARCH_POLL_INTERVAL_SEC)
    raise RuntimeError(f"Search {search_id} timed out after {SEARCH_TIMEOUT_SEC}s")


def get_results(search_id):
    resp = _http_request('GET', f'/api/ariel/searches/{search_id}/results',
                          label=f'fetch results {search_id}', timeout=60)
    return resp.json().get('events', [])


def run_aql(aql):
    search_id = submit_search(aql)
    wait_for_search(search_id)
    return get_results(search_id)


def run_with_retry(aql, retries=MAX_RETRIES_PER_SEARCH):
    last_err = None
    for attempt in range(1, retries + 2):
        try:
            return run_aql(aql)
        except RuntimeError as e:
            last_err = e
            logger.warning("Search attempt %d/%d failed: %s", attempt, retries + 1, e)
            if attempt <= retries:
                time.sleep(2 * attempt)
    raise last_err


def fallback_per_keyword(scope_clause, app_clause, keywords):
    """
    Backup path when the combined multi-keyword AQL fails outright. See the
    module docstring's note on the fallback path for the rationale.
    Decomposes into one simplified search per keyword (plain `payload`
    instead of `UTF8(payload)`) and merges/dedupes the results.
    """
    merged, any_ok = {}, False
    for kw in keywords:
        simple_kw_clause = build_keyword_clause([kw], payload_field='payload')
        aql = build_aql(scope_clause, app_clause, simple_kw_clause, TIME_RANGE_DAYS, RECORD_LIMIT_PER_SEARCH)
        try:
            events = run_with_retry(aql, retries=0)
            any_ok = True
        except RuntimeError as e:
            logger.warning("Fallback search for keyword '%s' failed: %s", kw, e)
            continue
        for ev in events:
            key = (ev.get('Time'), ev.get('Source'), ev.get('Raw Payload'))
            merged[key] = ev
    if not any_ok:
        raise RuntimeError("All per-keyword fallback searches failed.")
    return list(merged.values())


# ─── EXCEL REPORT ───────────────────────────────────────────────────────────────

class ExcelReportBuilder:
    """One worksheet per validation case, formatted for immediate readability:
    bold colored header row, fixed column widths, top-aligned wrapped payload
    text, and a clearly labeled placeholder row for zero-result or failed
    searches instead of a blank sheet."""

    _INVALID_SHEET_CHARS = set(r"[]:*?/\\")

    def __init__(self, filepath):
        self.workbook = xlsxwriter.Workbook(filepath)
        self.header_fmt = self.workbook.add_format({
            'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1,
            'align': 'center', 'valign': 'vcenter',
        })
        self.wrap_fmt = self.workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1})
        self.cell_fmt = self.workbook.add_format({'valign': 'top', 'border': 1})
        self.empty_fmt = self.workbook.add_format({'italic': True, 'font_color': '#888888', 'align': 'left'})
        self._used_names = set()

    def _safe_sheet_name(self, name):
        cleaned = "".join(c for c in name if c not in self._INVALID_SHEET_CHARS).strip() or "Sheet"
        cleaned = cleaned[:31]
        base, i = cleaned, 1
        while cleaned.lower() in self._used_names:
            suffix = f"_{i}"
            cleaned = base[: 31 - len(suffix)] + suffix
            i += 1
        self._used_names.add(cleaned.lower())
        return cleaned

    def write_case(self, case_name, rows, error=None):
        ws = self.workbook.add_worksheet(self._safe_sheet_name(case_name))
        headers = ['Time', 'Source', 'Log Source', 'Raw Payload']
        widths = [20, 16, 32, 110]
        for col, (h, w) in enumerate(zip(headers, widths)):
            ws.set_column(col, col, w)
            ws.write(0, col, h, self.header_fmt)
        ws.freeze_panes(1, 0)

        if error:
            ws.merge_range(1, 0, 1, 3, f"Search failed: {error}", self.empty_fmt)
            return
        if not rows:
            ws.merge_range(1, 0, 1, 3, "No matching events found in the configured time window.", self.empty_fmt)
            return

        for r, row in enumerate(rows, start=1):
            ws.write(r, 0, row.get('Time', ''), self.cell_fmt)
            ws.write(r, 1, row.get('Source', ''), self.cell_fmt)
            ws.write(r, 2, row.get('Log Source', ''), self.cell_fmt)
            ws.write(r, 3, row.get('Raw Payload', ''), self.wrap_fmt)

    def close(self):
        self.workbook.close()


# ─── MAIN ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 62)
    print("  QRadar EKS Application Log Validator")
    print("=" * 62)
    print(f"  Host             : {QRADAR_HOST}")
    print(f"  Log source filter: '{LOG_SOURCE_NAME_PATTERN}'")
    print(f"  Time window      : last {TIME_RANGE_DAYS} day(s)")
    print(f"  Retry config     : {MAX_RETRIES} attempts, {RETRY_DELAY_BASE}s base backoff")
    print("=" * 62)

    try:
        _validate_auth_config()
    except RuntimeError as e:
        print(f"❌ {e}")
        return

    if not test_connection():
        return

    log_sources = fetch_active_eks_log_sources()
    log_source_ids = [s['id'] for s in log_sources]

    try:
        scope_clause = build_scope_clause(log_source_ids, ADDITIONAL_HOSTS)
    except ValueError as e:
        print(f"❌ {e}")
        return

    app_name = prompt_application_name()
    app_clause = build_app_clause(app_name)

    report = ExcelReportBuilder(OUTPUT_FILE)

    print("🔍 Running validation cases...\n")
    for case_name, keywords in VALIDATION_CASES.items():
        print(f"   ▶ {case_name} ({len(keywords)} keyword(s))")
        keyword_clause = build_keyword_clause(keywords)
        primary_aql = build_aql(scope_clause, app_clause, keyword_clause, TIME_RANGE_DAYS, RECORD_LIMIT_PER_SEARCH)

        rows, error = [], None
        try:
            rows = run_with_retry(primary_aql)
        except RuntimeError as primary_err:
            print(f"      ⚠️  Primary AQL failed — trying per-keyword fallback: {primary_err}")
            try:
                rows = fallback_per_keyword(scope_clause, app_clause, keywords)
            except RuntimeError as fallback_err:
                error = f"primary: {primary_err} | fallback: {fallback_err}"
                print(f"      ❌ Fallback also failed: {fallback_err}")

        report.write_case(case_name, rows, error=error)
        status_icon = "✅" if not error else "❌"
        print(f"      {status_icon} {len(rows)} matching event(s)\n")

    report.close()
    print(f"💾 Report saved → {os.path.abspath(OUTPUT_FILE)}")
    print("\n✅ Done!")


if __name__ == '__main__':
    main()
