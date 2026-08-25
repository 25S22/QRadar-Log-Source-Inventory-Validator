"""
QRadar Rule Effectiveness Auditor
==================================
Cross-references every ENABLED correlation rule against offense history
across three lookback windows — 1 / 3 / 6 months — to answer the question
that actually matters to a SOC: are our rules working, and will they fire
when the conditions they were built for actually happen?

Looking at a single 30-day window (the old approach) can't tell the
difference between a rule that's *always* been unused and a rule that fired
every week for a year and then suddenly went quiet last month. Those are
very different problems — one is stale content, the other could be a
silent detection gap (a log source stopped feeding it, a field mapping
changed, a reference set emptied, an upstream rule it depends on broke).
Comparing the three windows side by side is how you tell them apart.

Per rule, this produces:
  · Dead — Never Fired   Zero offenses across the FULL 6-month window.
                          Strongest signal a rule is obsolete or was never
                          wired up correctly. Candidate for review/removal.
  · Recently Silent       Fired at some point in the 6-month window, but
                          nothing at all in the last 90 days. The "used to
                          work, might be quietly broken" signal — investigate
                          before assuming it's just low frequency.
  · Highly Active         Firing well above the expected volume for its
                          window. NOT automatically bad — a prompt to
                          confirm the volume is genuine and thresholds /
                          suppression are tuned as intended.
  · Active                Contributing at a normal, expected rate.

Zero AQL. All data from REST API:
  GET /api/analytics/rules
  GET /api/siem/offenses

Output → five-sheet Excel workbook + an Outlook HTML email draft with a
         1/3/6-month comparison view up top and a full switchable section
         per lookback window underneath.

A NOTE ON RULE DESCRIPTIONS
----------------------------
QRadar's public REST API does not reliably expose a rule's full boolean
test logic (the AND/OR conditions built in the Rule Wizard) as a plain-text
field in most versions. This script pulls whatever descriptive metadata IS
commonly available (name, type, owner, creation/modification dates, and a
notes/description field if your environment populates one) and is written
defensively so it degrades to a clear placeholder rather than guessing. For
the full trigger logic behind a specific rule, cross-reference it in the
QRadar Console (Offenses → Rules → select rule → Actions → Export/View).

A NOTE ON "SWITCHING VIEWS" IN THE EMAIL
------------------------------------------
Outlook's HTML rendering engine (Word) strips JavaScript and doesn't
support the CSS tricks (":target", modern selectors) that make click-to-
hide-and-show tabs work on the web. Building the email as "hidden until
clicked" tabs would risk a recipient opening it in desktop Outlook and
seeing a blank card. Instead, all three windows are always fully visible,
stacked with clear section headers, with a side-by-side comparison table up
top and quick-jump links to each section — this works in every mail client,
Outlook included, while still giving you the three views to compare.
"""

import os
import time
import tempfile
import traceback
import logging
from datetime import datetime, timedelta

import requests
import urllib3
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import win32com.client

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# ─── LOGGING ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION  ←  edit this block only
# ══════════════════════════════════════════════════════════════════════════════

QRADAR_HOST     = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL      = False          # locally hosted → safe to leave False

# ── Lookback windows ──────────────────────────────────────────────────────────
# The auditor builds THREE views over the same rule set — short / medium /
# long — so a rule that looks merely "quiet" in isolation stands out as a
# real trend once you can see its history.
#
# high_activity_threshold scales with window length (held at a constant
# rate of ~1.7 offenses/day) so "Highly Active" means the same thing in
# every view, instead of the 6-month window flagging everything as active
# purely because it's had more time to accumulate offenses.
LOOKBACK_PERIODS = [
    {'key': '1_month', 'label': '1 Month',  'days': 30,  'high_activity_threshold': 50},
    {'key': '3_month', 'label': '3 Months', 'days': 90,  'high_activity_threshold': 150},
    {'key': '6_month', 'label': '6 Months', 'days': 180, 'high_activity_threshold': 300},
]

# A rule with zero offense contributions across the FULL (longest) lookback
# window is classified "Dead". Kept at 0 — any activity at all means it's
# not truly dead, just possibly rare.
RULE_DEAD_THRESHOLD = 0

# ── Output ────────────────────────────────────────────────────────────────────
OUTPUT_DIR = r'C:\path\to\your\output'   # folder where Excel file is saved

# ── API / networking ──────────────────────────────────────────────────────────
REQUEST_TIMEOUT  = 30    # seconds per HTTP call
MAX_RETRIES      = 3     # attempts before giving up (exponential backoff)
RETRY_DELAY_BASE = 1.5   # seconds — waits: 1.5s → 3s → 6s
API_PAGE_SIZE    = 9999  # max items per page (QRadar Range header, 0-indexed)
MAX_PAGES        = 100   # safety cap on pagination loops (100 x 10000 = 1M items)

# ══════════════════════════════════════════════════════════════════════════════
#  END CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

_MAPI_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

OUTPUT_EXCEL = os.path.join(OUTPUT_DIR, 'qradar_rule_effectiveness.xlsx')

# Status labels used throughout (kept as constants so every sheet, chart, and
# table agree on the exact string — a typo here would silently break a filter).
STATUS_DEAD             = 'Dead — Never Fired'
STATUS_RECENTLY_SILENT  = 'Recently Silent'
STATUS_HIGHLY_ACTIVE    = 'Highly Active'
STATUS_ACTIVE           = 'Active'

# Per-period statuses (used for the three switchable views) are the same
# idea but scoped to a single window, so "Dead" here just means "0 offenses
# in THIS window" rather than the stronger 6-month claim above.
PERIOD_STATUS_DEAD          = 'Dead'
PERIOD_STATUS_HIGHLY_ACTIVE = 'Highly Active'
PERIOD_STATUS_ACTIVE        = 'Active'

PERIOD_NOTE = {
    PERIOD_STATUS_DEAD:          'No offenses in this window — review whether the rule is still needed.',
    PERIOD_STATUS_HIGHLY_ACTIVE: 'Firing well above typical volume for this window — confirm thresholds are tuned as intended.',
    PERIOD_STATUS_ACTIVE:        'Contributing normally.',
}


# ─── SHARED HTTP HELPERS ───────────────────────────────────────────────────────

def _api_get_page(path, range_start, range_end, params=None, label='request'):
    """
    Fetches a single page of a QRadar list endpoint using an explicit Range
    header. Returns (items, total) where total is parsed from the
    Content-Range response header when QRadar provides one, else None.

    Handles exponential-backoff retry on Timeout / ConnectionError and clean
    propagation of 401 / other HTTP errors.
    """
    url     = f"{QRADAR_HOST.rstrip('/')}{path}"
    headers = {
        'Accept':  'application/json',
        'Version': '14.0',
        'Range':   f'items={range_start}-{range_end}',
    }
    last_err = None

    for attempt in range(MAX_RETRIES):
        if attempt > 0:
            wait = RETRY_DELAY_BASE * (2 ** (attempt - 1))
            logger.warning("Retry %d/%d for '%s' — waiting %.1fs after %s",
                           attempt, MAX_RETRIES - 1, label, wait,
                           type(last_err).__name__)
            time.sleep(wait)
        try:
            resp = requests.get(
                url,
                params=params,
                auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
                verify=VERIFY_SSL,
                timeout=REQUEST_TIMEOUT,
                headers=headers,
            )
            if resp.status_code in (200, 206):
                total = None
                cr = resp.headers.get('Content-Range', '')
                if cr:
                    try:
                        total = int(cr.split('/')[-1].strip())
                    except Exception:
                        pass
                return resp.json(), total

            elif resp.status_code == 401:
                raise RuntimeError("Authentication failed — check QRADAR_USERNAME / PASSWORD")
            else:
                last_err = RuntimeError(f"HTTP {resp.status_code}")
                logger.warning("HTTP %d for '%s' (attempt %d/%d)",
                               resp.status_code, label, attempt + 1, MAX_RETRIES)

        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_err = exc
            logger.warning("%s on attempt %d for '%s'",
                           type(exc).__name__, attempt + 1, label)
        except RuntimeError:
            raise
        except Exception:
            logger.error("Non-retriable error for '%s':\n%s", label, traceback.format_exc())
            raise

    raise RuntimeError(f"All {MAX_RETRIES} attempts failed for '{label}': {last_err}")


def _api_get(path, params=None, label='request'):
    """Single-page GET — fine for small endpoints. Warns (but doesn't
    auto-paginate) if the server reports more items than fit in one page."""
    items, total = _api_get_page(path, 0, API_PAGE_SIZE, params=params, label=label)
    if total is not None and total > API_PAGE_SIZE + 1:
        logger.warning(
            "Pagination cap hit for '%s': %d items on server, only %d fetched. "
            "Use _api_get_all for this endpoint.", label, total, API_PAGE_SIZE + 1
        )
    return items


def _api_get_all(path, params=None, label='request', page_size=None, max_pages=MAX_PAGES):
    """
    Loops _api_get_page until every item is fetched. This matters once the
    offense lookback stretches to 3-6 months — a busy SOC can generate more
    offenses in that span than fit in a single page, and silently
    truncating that data would skew the whole rule-effectiveness picture
    toward "everything looks dead."
    """
    page_size = page_size or (API_PAGE_SIZE + 1)
    all_items = []
    start = 0
    for page_num in range(max_pages):
        end = start + page_size - 1
        items, total = _api_get_page(
            path, start, end, params=params,
            label=f'{label} (items {start}-{end})'
        )
        if not items:
            break
        all_items.extend(items)
        if total is not None and len(all_items) >= total:
            break
        if len(items) < page_size:
            break
        start += page_size
    else:
        logger.warning(
            "Hit max_pages=%d for '%s' — data may be incomplete. Raise MAX_PAGES "
            "if your deployment has more history than that.", max_pages, label
        )
    return all_items


# ─── CONNECTION TEST ──────────────────────────────────────────────────────────

def test_connection():
    print("🔗 Testing QRadar connection...")
    try:
        result = _api_get('/api/help/versions', label='connection test')
        if result:
            print("✅ QRadar connection successful!")
            return True
        print("⚠️  Unexpected empty response from /api/help/versions")
        return False
    except RuntimeError as e:
        print(f"❌ {e}")
        return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


# ─── DATA FETCHERS ─────────────────────────────────────────────────────────────

def fetch_all_rules():
    """
    Returns list of enabled correlation rule objects. Building blocks are
    excluded — they never generate offenses and would pollute the dead-rule
    analysis with hundreds of false positives.
    """
    print("📥 Fetching correlation rules...")
    try:
        data = _api_get_all('/api/analytics/rules', label='analytics rules')
        rules = [
            r for r in data
            if r.get('enabled', False) is True
            and 'BB' not in str(r.get('type', '')).upper()
            and 'BUILDING_BLOCK' not in str(r.get('type', '')).upper()
        ]
        print(f"   ✅ {len(rules)} enabled correlation rules (of {len(data)} total).")
        return rules
    except Exception as e:
        print(f"   ❌ Failed to fetch rules: {e}")
        return []


def fetch_offenses_for_lookback(days):
    """
    Fetches every offense with start_time within the last `days` days,
    fully paginated. Called ONCE for the longest configured window — the
    shorter 1/3-month views are derived by slicing this same list in memory
    (see filter_offenses_by_days), so QRadar is only queried once no matter
    how many lookback windows are configured.
    """
    print(f"📥 Fetching offenses from the last {days} days (paginated)...")
    cutoff_ms  = int((datetime.now() - timedelta(days=days)).timestamp() * 1000)
    filter_str = f'start_time >= {cutoff_ms}'
    fields_str = 'id,rules,start_time,magnitude,status,severity'
    try:
        data = _api_get_all(
            '/api/siem/offenses',
            params={'filter': filter_str, 'fields': fields_str},
            label='offenses'
        )
        print(f"   ✅ {len(data)} offenses fetched.")
        return data
    except Exception as e:
        print(f"   ⚠️  Filtered offense fetch failed ({e}). Trying unfiltered...")
        try:
            data = _api_get_all('/api/siem/offenses', label='offenses (unfiltered)')
            data = [o for o in data if (o.get('start_time') or 0) >= cutoff_ms]
            print(f"   ✅ {len(data)} offenses after in-memory filter.")
            return data
        except Exception as e2:
            print(f"   ❌ Offense fetch failed entirely: {e2} — rule analysis will show every rule as having zero contributions.")
            return []


def filter_offenses_by_days(offenses, days):
    """Slices an already-fetched offense list down to the last `days` days."""
    cutoff_ms = int((datetime.now() - timedelta(days=days)).timestamp() * 1000)
    return [o for o in offenses if (o.get('start_time') or 0) >= cutoff_ms]


# ─── RULE METADATA HELPERS ─────────────────────────────────────────────────────

def _extract_rule_ids(offense):
    """
    Extracts contributing rule IDs from an offense object.
    Handles: list of dicts ({'id': X, 'type': Y}), list of ints, or missing field.
    """
    ids = []
    for r in offense.get('rules', []):
        if isinstance(r, dict):
            rid = r.get('id')
        elif isinstance(r, (int, str)):
            rid = r
        else:
            rid = None
        if rid is not None:
            try:
                ids.append(int(rid))
            except (TypeError, ValueError):
                pass
    return ids


def _format_epoch_ms(val):
    """QRadar timestamps are epoch milliseconds. Returns a readable date, or '—' if absent/invalid."""
    if not val:
        return '—'
    try:
        return datetime.fromtimestamp(int(val) / 1000).strftime('%Y-%m-%d')
    except (TypeError, ValueError, OSError, OverflowError):
        return '—'


def _get_rule_description(rule):
    """
    Best-effort human-readable description of what the rule does / how it
    can be triggered.

    IMPORTANT: QRadar's public REST API does not reliably expose the full
    boolean test-stack (the AND/OR conditions built in the Rule Wizard) as
    plain text in most versions. This checks the few field names that
    sometimes carry notes in different QRadar versions, and falls back to a
    placeholder pointing back to the console rather than guessing.
    """
    for field in ('notes', 'description', 'rule_description', 'comment'):
        val = rule.get(field)
        if val and isinstance(val, str) and val.strip():
            return val.strip()
    return 'Not exposed via API — review in QRadar Console (Offenses → Rules) for full trigger logic.'


def _classify_single_period(count, high_activity_threshold):
    """Independent Dead/Highly Active/Active classification using ONLY this window's count."""
    if count <= RULE_DEAD_THRESHOLD:
        return PERIOD_STATUS_DEAD
    if count >= high_activity_threshold:
        return PERIOD_STATUS_HIGHLY_ACTIVE
    return PERIOD_STATUS_ACTIVE


def _recommendation_for_status(status):
    return {
        STATUS_DEAD: (
            f"No offense contributions across the full {LOOKBACK_PERIODS[-1]['days']}-day "
            "lookback. Review the rule's logic and confirm it's still needed — "
            "disable it if it's obsolete."
        ),
        STATUS_RECENTLY_SILENT: (
            "Fired in the past but nothing recently. This can mean the condition "
            "genuinely hasn't occurred — or that something upstream broke (a log "
            "source stopped feeding it, a field mapping changed, a reference set "
            "emptied). Worth a quick check rather than assuming it's just quiet."
        ),
        STATUS_HIGHLY_ACTIVE: (
            "Firing well above typical volume for its window. Confirm this reflects "
            "genuine, expected activity and that thresholds/suppression are tuned "
            "as intended."
        ),
        STATUS_ACTIVE: 'Contributing normally across all lookback windows. No action required.',
    }.get(status, '')


# ─── RULE ANALYSIS ────────────────────────────────────────────────────────────

def build_master_rule_table(rules, offenses_all):
    """
    Builds one row per enabled rule with:
      - descriptive metadata (name, type, origin, owner, dates, description)
      - offense counts for each configured lookback period
      - an INDEPENDENT Dead/Highly Active/Active status per period
        (status_1_month, status_3_month, status_6_month — drives the three
        switchable email views)
      - one COMPOSITE 'overall_status' that reasons across periods together
        (drives the Excel workbook and the "needs investigation" callout) —
        this is what catches a rule that fired steadily for months and then
        went quiet, which looks merely "quiet" in any single-period view
        but is a much stronger signal once you can see the history.
    """
    period_offenses = {p['key']: filter_offenses_by_days(offenses_all, p['days'])
                        for p in LOOKBACK_PERIODS}

    period_counts = {}
    for key, offs in period_offenses.items():
        counts = {}
        for o in offs:
            for rid in _extract_rule_ids(o):
                counts[rid] = counts.get(rid, 0) + 1
        period_counts[key] = counts

    rows = []
    for rule in rules:
        rid = rule.get('id')
        row = {
            'rule_id':     rid,
            'rule_name':   rule.get('name', f'Rule {rid}'),
            'rule_type':   rule.get('type', 'UNKNOWN'),
            'origin':      rule.get('origin', 'UNKNOWN'),
            'owner':       rule.get('owner', 'Unknown'),
            'created':     _format_epoch_ms(rule.get('creation_date')),
            'modified':    _format_epoch_ms(rule.get('modification_date')),
            'description': _get_rule_description(rule),
        }
        for p in LOOKBACK_PERIODS:
            count = period_counts[p['key']].get(rid, 0)
            row[f"offenses_{p['key']}"] = count
            row[f"status_{p['key']}"]   = _classify_single_period(count, p['high_activity_threshold'])
        rows.append(row)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    longest_key  = LOOKBACK_PERIODS[-1]['key']    # e.g. 6_month — defines "ever fired"
    mid_key      = LOOKBACK_PERIODS[1]['key']     # e.g. 3_month — defines "still recently active"
    shortest_key = LOOKBACK_PERIODS[0]['key']     # e.g. 1_month — defines "currently highly active"

    def _overall(r):
        if r[f'offenses_{longest_key}'] <= RULE_DEAD_THRESHOLD:
            return STATUS_DEAD
        if r[f'offenses_{mid_key}'] <= RULE_DEAD_THRESHOLD:
            return STATUS_RECENTLY_SILENT
        if r[f'status_{shortest_key}'] == PERIOD_STATUS_HIGHLY_ACTIVE:
            return STATUS_HIGHLY_ACTIVE
        return STATUS_ACTIVE

    df['overall_status']  = df.apply(_overall, axis=1)
    df['recommendation']  = df['overall_status'].map(_recommendation_for_status)
    return df


def get_period_view(master_df, period_key):
    """
    Extracts one lookback window's self-contained view from the master
    table — this is what powers each of the three switchable email
    sections. Sorted with the most actionable rows first.
    """
    if master_df.empty:
        return master_df
    count_col  = f"offenses_{period_key}"
    status_col = f"status_{period_key}"
    view = master_df[['rule_id', 'rule_name', 'rule_type', 'origin', count_col, status_col]].copy()
    view = view.rename(columns={count_col: 'offense_count', status_col: 'status'})
    view['note'] = view['status'].map(PERIOD_NOTE).fillna('')
    order = {PERIOD_STATUS_DEAD: 0, PERIOD_STATUS_HIGHLY_ACTIVE: 1, PERIOD_STATUS_ACTIVE: 2}
    view['_sort'] = view['status'].map(order).fillna(3)
    view = (view.sort_values(['_sort', 'offense_count'], ascending=[True, False])
                .drop(columns=['_sort'])
                .reset_index(drop=True))
    return view


# ─── CHART ────────────────────────────────────────────────────────────────────

_BG = '#0a0618'   # dark background baked into the chart PNG


def generate_rule_trend_chart(master_df):
    """
    Grouped bar chart: Dead / Highly Active / Active rule counts, one group
    per lookback window. This is the single visual answer to "switching
    between views" — rather than hiding two windows at a time, it puts all
    three side by side so the trend (rules going dead vs. staying active as
    the window widens) is visible in one glance.
    """
    if master_df.empty:
        return None

    categories = [PERIOD_STATUS_DEAD, PERIOD_STATUS_HIGHLY_ACTIVE, PERIOD_STATUS_ACTIVE]
    colours    = {PERIOD_STATUS_DEAD: '#ef4444', PERIOD_STATUS_HIGHLY_ACTIVE: '#3b82f6', PERIOD_STATUS_ACTIVE: '#10b981'}

    period_labels = [p['label'] for p in LOOKBACK_PERIODS]
    counts_by_cat = {cat: [] for cat in categories}
    for p in LOOKBACK_PERIODS:
        vc = master_df[f"status_{p['key']}"].value_counts()
        for cat in categories:
            counts_by_cat[cat].append(int(vc.get(cat, 0)))

    fig, ax = plt.subplots(figsize=(6.5, 4), facecolor=_BG)
    ax.set_facecolor(_BG)

    x_positions = list(range(len(period_labels)))
    width   = 0.24
    offsets = [-width, 0, width]

    max_val = 1
    for i, cat in enumerate(categories):
        positions = [xi + offsets[i] for xi in x_positions]
        bars = ax.bar(positions, counts_by_cat[cat], width=width,
                       color=colours[cat], label=cat, edgecolor=_BG, linewidth=1)
        for bar, val in zip(bars, counts_by_cat[cat]):
            max_val = max(max_val, val)
            if val > 0:
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(),
                        str(val), ha='center', va='bottom', color='#c4b5fd',
                        fontsize=8, fontfamily='monospace')

    ax.set_ylim(0, max_val * 1.18)
    ax.set_xticks(x_positions)
    ax.set_xticklabels(period_labels, color='#c4b5fd', fontsize=9, fontfamily='monospace')
    ax.set_ylabel('Rule count', color='#7c6fa0', fontsize=8, fontfamily='monospace')
    ax.set_title('Rule Status Across Lookback Windows', color='#a78bfa',
                 fontsize=10, fontweight='700', fontfamily='monospace', pad=12)
    ax.tick_params(colors='#c4b5fd', labelsize=8)
    for spine in ax.spines.values():
        spine.set_edgecolor('#3a2d6a')
        spine.set_linewidth(0.5)
    ax.legend(fontsize=8, frameon=False, labelcolor='#c4b5fd', loc='upper right')

    plt.tight_layout(pad=0.8)
    tmp  = tempfile.NamedTemporaryFile(suffix='.png', delete=False, prefix='qr_trend_')
    path = tmp.name
    tmp.close()
    plt.savefig(path, bbox_inches='tight', dpi=110, facecolor=_BG, edgecolor='none')
    plt.close(fig)
    return path


# ─── EXCEL REPORT ─────────────────────────────────────────────────────────────

_FILLS = {
    'red':    PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid'),
    'orange': PatternFill(start_color='FFBF47', end_color='FFBF47', fill_type='solid'),
    'green':  PatternFill(start_color='A8E6CF', end_color='A8E6CF', fill_type='solid'),
    'blue':   PatternFill(start_color='74B9FF', end_color='74B9FF', fill_type='solid'),
    'header': PatternFill(start_color='2D2257', end_color='2D2257', fill_type='solid'),
}
_HDR_FONT = Font(bold=True, color='E8E0FF', size=10)
_BOLD     = Font(bold=True, size=10)
_CENTRE   = Alignment(horizontal='center', vertical='center')
_LEFT     = Alignment(horizontal='left',   vertical='center')
_WRAP     = Alignment(horizontal='left',   vertical='top', wrap_text=True)


def _status_fill_key(status):
    return {
        STATUS_DEAD:            'red',
        STATUS_RECENTLY_SILENT: 'orange',
        STATUS_HIGHLY_ACTIVE:   'blue',
        STATUS_ACTIVE:          'green',
    }.get(status, 'green')


def _write_sheet_header(ws, columns, col_widths):
    ws.append(columns)
    for col_idx, (cell, width) in enumerate(zip(ws[1], col_widths), start=1):
        cell.fill      = _FILLS['header']
        cell.font      = _HDR_FONT
        cell.alignment = _CENTRE
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 20


def _write_summary_sheet(wb, master_df):
    ws = wb.active
    ws.title = 'Executive Summary'

    ws['A1'] = 'QRadar Rule Effectiveness Report'
    ws['A1'].font = Font(bold=True, size=16, color='2D2257')
    ws['A2'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws['A2'].font = Font(italic=True, size=10, color='7C6FA0')
    ws['A3'] = f"Lookback windows: {', '.join(p['label'] for p in LOOKBACK_PERIODS)}"
    ws['A3'].font = Font(size=10, color='7C6FA0')

    ws['A5'] = 'Total enabled correlation rules analyzed:'
    ws['B5'] = int(len(master_df))
    ws['B5'].font = _BOLD

    row = 7
    ws.cell(row=row, column=1, value='Status').fill = _FILLS['header']
    ws.cell(row=row, column=1).font = _HDR_FONT
    for i, p in enumerate(LOOKBACK_PERIODS, start=2):
        c = ws.cell(row=row, column=i, value=p['label'])
        c.fill = _FILLS['header']
        c.font = _HDR_FONT
        c.alignment = _CENTRE
    row += 1

    for status, fill_key in [
        (PERIOD_STATUS_DEAD, 'red'),
        (PERIOD_STATUS_HIGHLY_ACTIVE, 'blue'),
        (PERIOD_STATUS_ACTIVE, 'green'),
    ]:
        ws.cell(row=row, column=1, value=status).font = _BOLD
        for i, p in enumerate(LOOKBACK_PERIODS, start=2):
            count = int((master_df[f"status_{p['key']}"] == status).sum())
            c = ws.cell(row=row, column=i, value=count)
            c.alignment = _CENTRE
            c.fill = _FILLS[fill_key]
        row += 1

    row += 1
    silent_count = int((master_df['overall_status'] == STATUS_RECENTLY_SILENT).sum())
    dead_count   = int((master_df['overall_status'] == STATUS_DEAD).sum())
    ws.cell(row=row, column=1, value='Dead — never fired in 6 months:').font = _BOLD
    ws.cell(row=row, column=2, value=dead_count).font = _BOLD
    row += 1
    ws.cell(row=row, column=1, value='Recently silent — fired historically, quiet now:').font = _BOLD
    ws.cell(row=row, column=2, value=silent_count).font = _BOLD
    row += 2

    note = (
        'Note: rule descriptions reflect whatever metadata QRadar exposes via REST '
        'API for each rule. Full boolean trigger logic (the AND/OR test stack built '
        'in the Rule Wizard) is not reliably exposed via this API in most QRadar '
        'versions — cross-reference the Rules console for full logic wherever a '
        'description is marked as unavailable.'
    )
    ws.cell(row=row, column=1, value=note)
    ws.cell(row=row, column=1).font = Font(italic=True, size=9, color='7C6FA0')
    ws.cell(row=row, column=1).alignment = _WRAP
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 48

    for i, w in enumerate([46, 14, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_master_sheet(wb, master_df):
    ws = wb.create_sheet('Rule Effectiveness (All Rules)')

    period_headers = [f"Offenses ({p['label']})" for p in LOOKBACK_PERIODS]
    cols = (['Rule ID', 'Rule Name', 'Type', 'Origin', 'Owner', 'Created', 'Modified']
            + period_headers
            + ['Overall Status', 'Recommendation', 'Description / Trigger Logic'])
    widths = ([9, 40, 14, 10, 14, 12, 12]
              + [16] * len(period_headers)
              + [20, 46, 60])
    _write_sheet_header(ws, cols, widths)

    if master_df.empty:
        return

    status_col_idx = cols.index('Overall Status') + 1
    desc_col_idx   = cols.index('Description / Trigger Logic') + 1

    for _, r in master_df.iterrows():
        values = [int(r['rule_id']), str(r['rule_name']), str(r['rule_type']),
                  str(r['origin']), str(r['owner']), str(r['created']), str(r['modified'])]
        values += [int(r[f"offenses_{p['key']}"]) for p in LOOKBACK_PERIODS]
        values += [str(r['overall_status']), str(r['recommendation']), str(r['description'])]
        ws.append(values)

        r_idx = ws.max_row
        status_cell = ws.cell(row=r_idx, column=status_col_idx)
        status_cell.font = _BOLD
        status_cell.fill = _FILLS[_status_fill_key(r['overall_status'])]

        ws.cell(row=r_idx, column=desc_col_idx).alignment = _WRAP

    ws.freeze_panes = 'A2'


def _write_filtered_sheet(wb, master_df, status_value, sheet_title, sort_col, ascending):
    ws = wb.create_sheet(sheet_title)

    cols = (['Rule ID', 'Rule Name', 'Type', 'Origin', 'Owner', 'Created']
            + [f"Offenses ({p['label']})" for p in LOOKBACK_PERIODS]
            + ['Recommendation', 'Description / Trigger Logic'])
    widths = [9, 42, 14, 10, 14, 12] + [16] * len(LOOKBACK_PERIODS) + [46, 60]
    _write_sheet_header(ws, cols, widths)

    if master_df.empty:
        ws.append(['No rule data available.'])
        return

    filtered = master_df[master_df['overall_status'] == status_value]
    if sort_col in filtered.columns:
        filtered = filtered.sort_values(sort_col, ascending=ascending)

    if filtered.empty:
        ws.append(['No rules currently fall into this category — nothing to review.'])
        return

    desc_col_idx = cols.index('Description / Trigger Logic') + 1

    for _, r in filtered.iterrows():
        values = [int(r['rule_id']), str(r['rule_name']), str(r['rule_type']),
                  str(r['origin']), str(r['owner']), str(r['created'])]
        values += [int(r[f"offenses_{p['key']}"]) for p in LOOKBACK_PERIODS]
        values += [str(r['recommendation']), str(r['description'])]
        ws.append(values)
        ws.cell(row=ws.max_row, column=desc_col_idx).alignment = _WRAP

    ws.freeze_panes = 'A2'


def save_excel_report(master_df, path):
    """
    Saves the five-sheet workbook:
      1. Executive Summary
      2. Rule Effectiveness (All Rules) — the sortable master table
      3. Dead Rules — never fired in 6 months, oldest first
      4. Recently Silent — fired historically, quiet in the last 90 days
      5. Highly Active Rules — firing well above typical volume
    """
    try:
        wb = openpyxl.Workbook()

        if master_df.empty:
            ws = wb.active
            ws.title = 'No Data'
            ws['A1'] = 'No enabled rules or offense data were available to analyze.'
            ws['A1'].font = _BOLD
            wb.save(path)
            print(f"⚠️  Excel saved with no data → {path}")
            return

        _write_summary_sheet(wb, master_df)
        _write_master_sheet(wb, master_df)
        _write_filtered_sheet(wb, master_df, STATUS_DEAD, 'Dead Rules',
                               sort_col='created', ascending=True)
        _write_filtered_sheet(wb, master_df, STATUS_RECENTLY_SILENT, 'Recently Silent',
                               sort_col=f"offenses_{LOOKBACK_PERIODS[-1]['key']}", ascending=False)
        _write_filtered_sheet(wb, master_df, STATUS_HIGHLY_ACTIVE, 'Highly Active Rules',
                               sort_col=f"offenses_{LOOKBACK_PERIODS[0]['key']}", ascending=False)

        wb.save(path)
        print(f"✅ Excel saved → {path}")
    except PermissionError:
        print(f"❌ Cannot save Excel — is '{path}' already open?")
    except Exception as e:
        logger.error("Excel save failed:\n%s", traceback.format_exc())
        print(f"❌ Excel save error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
#  EMAIL HTML
# ══════════════════════════════════════════════════════════════════════════════

_C = {
    'purple':   '#9b72f5',
    'violet':   '#c4b5fd',
    'lavender': '#a78bfa',
    'dim':      '#7c6fa0',
    'green':    '#10b981',
    'red':      '#ef4444',
    'orange':   '#f97316',
    'amber':    '#f59e0b',
    'gray':     '#8b9ab0',
    'cyan':     '#06b6d4',
    'blue':     '#3b82f6',
    'badge_red':   '#7f1d1d',
    'badge_amber': '#78350f',
    'badge_green': '#065f46',
    'badge_gray':  '#2d3748',
    'badge_blue':  '#1e3a5f',
}


def _chip(label, value, color):
    if value == 0:
        return ''
    return (
        f'<span style="display:inline-block;border-left:3px solid {color};'
        f'padding:1px 10px 1px 7px;margin:3px 8px 3px 0;font-size:10px;'
        f'font-family:monospace;color:{color};font-weight:700;letter-spacing:0.3px;">'
        f'{value}&nbsp;{label}</span>'
    )


def _rule_status_badge(status):
    meta = {
        PERIOD_STATUS_DEAD:          {'bg': _C['badge_red'],   'label': 'DEAD',          'icon': '●'},
        PERIOD_STATUS_HIGHLY_ACTIVE: {'bg': _C['badge_blue'],  'label': 'HIGHLY ACTIVE', 'icon': '▲'},
        PERIOD_STATUS_ACTIVE:        {'bg': _C['badge_green'], 'label': 'ACTIVE',        'icon': '✔'},
        STATUS_RECENTLY_SILENT:      {'bg': _C['badge_amber'], 'label': 'INVESTIGATE',   'icon': '?'},
    }.get(status, {'bg': _C['badge_gray'], 'label': str(status).upper()[:14], 'icon': '◌'})
    return (
        f'<span style="background:{meta["bg"]};color:#f0eaff;font-size:9px;'
        f'font-weight:700;padding:2px 8px;border-radius:3px;letter-spacing:0.5px;'
        f'white-space:nowrap;font-family:monospace;">{meta["icon"]}&nbsp;{meta["label"]}</span>'
    )


def _build_period_table_html(view_df):
    """Shows only Dead + Highly Active rows for this window — the actionable
    ones. Active rules are healthy; the full list is always in the Excel."""
    C = _C
    if view_df.empty:
        return f'<p style="color:{C["dim"]};font-size:11px;font-family:monospace;padding:8px 0;">No rule data available.</p>'

    actionable = view_df[view_df['status'].isin([PERIOD_STATUS_DEAD, PERIOD_STATUS_HIGHLY_ACTIVE])]
    if actionable.empty:
        return (f'<p style="color:{C["green"]};font-size:11px;font-weight:700;'
                f'font-family:monospace;padding:8px 0;">✔ No dead or highly active rules in this window.</p>')

    rows_html = ''
    _rb = f'border-bottom:1px solid {C["dim"]}30;'
    for _, row in actionable.iterrows():
        name = str(row['rule_name'])
        name_short = name[:48] + '…' if len(name) > 48 else name
        count_color = C['red'] if row['status'] == PERIOD_STATUS_DEAD else C['blue']
        badge = _rule_status_badge(row['status'])
        rows_html += f"""
        <tr>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{C['violet']};
                     font-family:monospace;" title="{name}">{name_short}</td>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{C['dim']};
                     font-family:monospace;text-align:center;">{row['rule_type']}</td>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{count_color};
                     font-family:monospace;text-align:center;font-weight:700;">{row['offense_count']}</td>
          <td style="padding:7px 10px;{_rb}text-align:center;">{badge}</td>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{C['dim']};max-width:220px;">{row['note']}</td>
        </tr>"""

    _hdr = f'border-top:2px solid {C["purple"]};border-bottom:1px solid {C["purple"]}50;'
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:8px;">
      <thead><tr>
        <th style="padding:6px 10px;text-align:left;font-size:9px;color:{C['purple']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Rule Name</th>
        <th style="padding:6px 10px;text-align:center;font-size:9px;color:{C['purple']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Type</th>
        <th style="padding:6px 10px;text-align:center;font-size:9px;color:{C['purple']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Offenses</th>
        <th style="padding:6px 10px;text-align:center;font-size:9px;color:{C['purple']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Status</th>
        <th style="padding:6px 10px;text-align:left;font-size:9px;color:{C['purple']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Note</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>"""


def _build_silent_table_html(master_df, long_key):
    C = _C
    silent = master_df[master_df['overall_status'] == STATUS_RECENTLY_SILENT]
    if silent.empty:
        return (f'<p style="color:{C["green"]};font-size:11px;font-weight:700;font-family:monospace;'
                f'padding:8px 0;">✔ No rules have gone quiet — everything that fired historically '
                f'is still firing recently.</p>')

    silent = silent.sort_values(f'offenses_{long_key}', ascending=False)
    rows_html = ''
    _rb = f'border-bottom:1px solid {C["dim"]}30;'
    for _, row in silent.iterrows():
        name = str(row['rule_name'])
        name_short = name[:48] + '…' if len(name) > 48 else name
        rows_html += f"""
        <tr>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{C['violet']};
                     font-family:monospace;" title="{name}">{name_short}</td>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{C['dim']};
                     font-family:monospace;text-align:center;">{row['rule_type']}</td>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{C['amber']};
                     font-family:monospace;text-align:center;font-weight:700;">{int(row[f'offenses_{long_key}'])}</td>
          <td style="padding:7px 10px;{_rb}font-size:10px;color:{C['dim']};
                     font-family:monospace;text-align:center;">0</td>
        </tr>"""

    _hdr = f'border-top:2px solid {C["orange"]};border-bottom:1px solid {C["orange"]}50;'
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:8px;">
      <thead><tr>
        <th style="padding:6px 10px;text-align:left;font-size:9px;color:{C['orange']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Rule Name</th>
        <th style="padding:6px 10px;text-align:center;font-size:9px;color:{C['orange']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Type</th>
        <th style="padding:6px 10px;text-align:center;font-size:9px;color:{C['orange']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Offenses (long window)</th>
        <th style="padding:6px 10px;text-align:center;font-size:9px;color:{C['orange']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.4px;font-family:monospace;{_hdr}">Offenses (recent)</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>"""


def _build_comparison_table_html(master_df):
    """The at-a-glance 'switch between views' table — every window's
    category counts side by side, no clicking required."""
    C = _C
    _hdr = f'border-top:2px solid {C["purple"]};border-bottom:1px solid {C["purple"]}50;'
    _rb  = f'border-bottom:1px solid {C["dim"]}30;'

    header_cells = ''.join(
        f'<th style="padding:6px 10px;text-align:center;font-size:9px;color:{C["purple"]};'
        f'font-weight:700;text-transform:uppercase;letter-spacing:1.2px;'
        f'font-family:monospace;{_hdr}">{p["label"]}</th>'
        for p in LOOKBACK_PERIODS
    )

    def _row(label, color):
        cells = ''
        for p in LOOKBACK_PERIODS:
            count = int((master_df[f"status_{p['key']}"] == label).sum())
            cells += (f'<td style="padding:7px 10px;{_rb}font-size:12px;color:{color};'
                      f'font-family:monospace;text-align:center;font-weight:700;">{count}</td>')
        return (f'<tr><td style="padding:7px 10px;{_rb}font-size:10px;color:{C["violet"]};'
                f'font-family:monospace;font-weight:700;">{label}</td>{cells}</tr>')

    rows_html = (
        _row(PERIOD_STATUS_DEAD, C['red'])
        + _row(PERIOD_STATUS_HIGHLY_ACTIVE, C['blue'])
        + _row(PERIOD_STATUS_ACTIVE, C['green'])
    )

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:8px;">
      <thead><tr>
        <th style="padding:6px 10px;text-align:left;font-size:9px;color:{C['purple']};font-weight:700;
                   text-transform:uppercase;letter-spacing:1.2px;font-family:monospace;{_hdr}">Status</th>
        {header_cells}
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>"""


def build_email_html(master_df, chart_cid):
    """Assembles the full HTML email body."""
    C        = _C
    run_time = datetime.now().strftime('%d %b %Y  ·  %H:%M:%S')

    if master_df.empty:
        return f"""<!DOCTYPE html><html><body style="font-family:monospace;padding:24px;">
        <h2 style="color:{C['purple']};">QRadar Rule Effectiveness Report</h2>
        <p style="color:{C['dim']};">No enabled rules or offense data were available at
        {run_time}. Check the console output for connection or permission errors.</p>
        </body></html>"""

    total_rules  = len(master_df)
    dead_total   = int((master_df['overall_status'] == STATUS_DEAD).sum())
    silent_total = int((master_df['overall_status'] == STATUS_RECENTLY_SILENT).sum())
    hi_1m_total  = int((master_df[f"status_{LOOKBACK_PERIODS[0]['key']}"] == PERIOD_STATUS_HIGHLY_ACTIVE).sum())

    if dead_total > 10 or silent_total > 5:
        hdr_bg, hdr_txt = C['badge_red'], f'⚠  {dead_total + silent_total} RULES NEED ATTENTION'
    elif dead_total > 0 or silent_total > 0:
        hdr_bg, hdr_txt = C['badge_amber'], f'⚠  {dead_total + silent_total} RULES NEED ATTENTION'
    else:
        hdr_bg, hdr_txt = C['badge_green'], '✔  RULES HEALTHY'

    def badge(bg, txt):
        return (f'<span style="background:{bg};color:#f0eaff;font-size:10px;font-weight:700;'
                f'padding:4px 12px;border-radius:3px;letter-spacing:0.8px;'
                f'font-family:monospace;white-space:nowrap;">{txt}</span>')

    def metric(label, value, color, note=''):
        note_html = (f'<div style="font-size:9px;color:{C["dim"]};margin-top:3px;'
                     f'font-family:monospace;">{note}</div>') if note else ''
        return (f'<td style="padding:0 22px 0 0;text-align:center;vertical-align:top;">'
                f'<div style="font-size:30px;font-weight:800;color:{color};line-height:1;'
                f'font-family:monospace;letter-spacing:-1px;">{value}</div>'
                f'<div style="font-size:9px;color:{C["dim"]};margin-top:4px;'
                f'text-transform:uppercase;letter-spacing:1.2px;">{label}</div>{note_html}</td>')

    headline_metrics = (
        metric('Total Rules', total_rules, C['purple'])
        + metric('Dead (6mo)', dead_total, C['red'] if dead_total > 0 else C['green'])
        + metric('Investigate', silent_total, C['orange'] if silent_total > 0 else C['green'], 'quiet after firing')
        + metric('Highly Active (1mo)', hi_1m_total, C['blue'] if hi_1m_total > 0 else C['green'])
    )

    nav_pill_style = (f'display:inline-block;padding:6px 16px;margin:0 6px 8px 0;'
                      f'border:1px solid {C["purple"]}60;border-radius:20px;'
                      f'color:{C["violet"]};font-size:10px;font-weight:700;'
                      f'font-family:monospace;text-decoration:none;letter-spacing:0.5px;')
    nav_html = ''.join(
        f'<a href="#view-{p["key"]}" style="{nav_pill_style}">{p["label"].upper()}</a>'
        for p in LOOKBACK_PERIODS
    )

    comparison_table = _build_comparison_table_html(master_df)
    silent_table      = _build_silent_table_html(master_df, LOOKBACK_PERIODS[-1]['key'])
    chart_html = (f'<img src="cid:{chart_cid}" alt="Rule status trend chart" '
                  f'style="display:block;max-width:100%;margin:14px auto 0;">') if chart_cid else ''

    period_sections = ''
    for p in LOOKBACK_PERIODS:
        view = get_period_view(master_df, p['key'])
        dead_n   = int((view['status'] == PERIOD_STATUS_DEAD).sum())          if not view.empty else 0
        hi_n     = int((view['status'] == PERIOD_STATUS_HIGHLY_ACTIVE).sum()) if not view.empty else 0
        active_n = int((view['status'] == PERIOD_STATUS_ACTIVE).sum())        if not view.empty else 0

        if dead_n > 10 or hi_n > 5:
            p_bg, p_txt = C['badge_red'], f'⚠ {dead_n + hi_n} flagged'
        elif dead_n > 0 or hi_n > 0:
            p_bg, p_txt = C['badge_amber'], f'⚠ {dead_n + hi_n} flagged'
        else:
            p_bg, p_txt = C['badge_green'], '✔ healthy'

        table_html = _build_period_table_html(view)

        period_sections += f"""
  <tr><td id="view-{p['key']}" style="padding:28px 0 4px;border-top:2px solid {C['purple']}50;">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td>
        <span style="font-size:13px;font-weight:700;color:{C['purple']};font-family:monospace;">
          {p['label']} View</span>
        <span style="font-size:10px;color:{C['dim']};margin-left:12px;font-family:monospace;">
          lookback: {p['days']} days</span>
      </td>
      <td align="right">{badge(p_bg, p_txt)}</td>
    </tr></table>
    <div style="margin-top:8px;">
      {_chip('Dead', dead_n, C['red'])}
      {_chip('Highly Active', hi_n, C['blue'])}
      {_chip('Active', active_n, C['green'])}
    </div>
    {table_html}
  </td></tr>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="color-scheme" content="light dark">
<meta name="supported-color-schemes" content="light dark">
</head>
<body style="margin:0;padding:0;font-family:'Segoe UI',Helvetica,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:24px 0;">
<tr><td align="center">
<table width="660" cellpadding="0" cellspacing="0" style="max-width:660px;width:100%;">

  <!-- ══ MASTHEAD ══ -->
  <tr><td style="padding:0 0 10px;border-bottom:3px solid {C['purple']};">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td>
        <div style="font-size:9px;color:{C['dim']};letter-spacing:3px;
                    text-transform:uppercase;font-family:monospace;margin-bottom:8px;">
          QRadar &nbsp;·&nbsp; Weekly Intelligence Report
        </div>
        <div style="font-size:23px;font-weight:800;color:{C['violet']};
                    letter-spacing:-0.5px;line-height:1.2;">
          Rule Effectiveness Auditor
        </div>
        <div style="margin-top:8px;font-size:11px;color:{C['dim']};font-family:monospace;">{run_time}</div>
      </td>
      <td align="right" valign="top">{badge(hdr_bg, hdr_txt)}</td>
    </tr></table>
  </td></tr>

  <!-- ══ HEADLINE METRICS ══ -->
  <tr><td style="padding:20px 0 8px;">
    <table cellpadding="0" cellspacing="0"><tr>{headline_metrics}</tr></table>
  </td></tr>

  <!-- ══ COMPARISON — the "all three views at once" table ══ -->
  <tr><td style="padding:16px 0 2px;">
    <span style="font-size:9px;color:{C['purple']};text-transform:uppercase;letter-spacing:2px;
                 font-family:monospace;font-weight:700;">1 / 3 / 6 Month Comparison</span>
    <span style="font-size:10px;color:{C['dim']};margin-left:10px;">
      Rule count by status, across every lookback window
    </span>
  </td></tr>
  <tr><td style="padding:0 0 8px;">{comparison_table}</td></tr>
  <tr><td style="padding:0 0 20px;text-align:center;">{chart_html}</td></tr>

  <!-- ══ NEEDS INVESTIGATION ══ -->
  <tr><td style="padding:20px 0 4px;border-top:2px solid {C['orange']}60;">
    <span style="font-size:13px;font-weight:700;color:{C['orange']};font-family:monospace;">
      ⚠ Needs Investigation — Recently Silent</span>
    <div style="font-size:10px;color:{C['dim']};margin-top:4px;">
      Fired at some point in the last {LOOKBACK_PERIODS[-1]['days']} days, but nothing in the
      last {LOOKBACK_PERIODS[1]['days']} — check these before assuming they're just low-frequency.
    </div>
  </td></tr>
  <tr><td style="padding:0 0 28px;">{silent_table}</td></tr>

  <!-- ══ QUICK NAV ══ -->
  <tr><td style="padding:4px 0 4px;border-top:1px solid {C['purple']}30;">
    <div style="font-size:9px;color:{C['dim']};text-transform:uppercase;letter-spacing:1.5px;
                font-family:monospace;padding:14px 0 6px;">Jump to a window's detail</div>
    {nav_html}
  </td></tr>

  {period_sections}

  <!-- ══ FOOTER ══ -->
  <tr><td style="padding:20px 0 20px;border-top:1px solid {C['purple']}30;">
    <div style="font-size:9px;color:{C['dim']};font-family:monospace;letter-spacing:0.5px;">
      QRadar Rule Effectiveness Auditor &nbsp;·&nbsp; Auto-generated {run_time} &nbsp;·&nbsp;
      Source: QRadar REST API (no AQL) &nbsp;·&nbsp; Full data for every rule in the Excel attachment
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>"""


# ─── OUTLOOK DRAFT ────────────────────────────────────────────────────────────

def create_outlook_draft(excel_path, subject, html_body, images):
    """
    Creates an Outlook draft with embedded PNG chart(s) and Excel attachment.
    images = {'cid_key': '/path/to/file.png', ...}
    """
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail    = outlook.CreateItem(0)
        mail.Subject = subject

        if excel_path and os.path.exists(excel_path):
            mail.Attachments.Add(excel_path)

        for cid, img_path in images.items():
            if img_path and os.path.exists(img_path):
                att = mail.Attachments.Add(img_path)
                att.PropertyAccessor.SetProperty(_MAPI_PR_ATTACH_CONTENT_ID, cid)

        mail.HTMLBody = html_body
        mail.Display()
        print("\n✉️  Outlook draft created.")

    except Exception as e:
        logger.error("Outlook draft failed:\n%s", traceback.format_exc())
        print(f"\n❌ Outlook draft error: {e}")

    finally:
        for cid, img_path in images.items():
            if img_path and os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("=" * 62)
    print("  QRadar Rule Effectiveness Auditor")
    print("=" * 62)
    print(f"  Host            : {QRADAR_HOST}")
    print(f"  Lookback windows: {', '.join(p['label'] for p in LOOKBACK_PERIODS)}")
    print(f"  Retry config    : {MAX_RETRIES} attempts, {RETRY_DELAY_BASE}s base backoff")
    print("=" * 62)

    if not test_connection():
        return

    rules = fetch_all_rules()
    if not rules:
        print("\n❌ No enabled rules returned — check credentials, host URL, and permissions.")
        return

    max_days = max(p['days'] for p in LOOKBACK_PERIODS)
    offenses_all = fetch_offenses_for_lookback(max_days)
    if not offenses_all:
        print("\n⚠️  WARNING: no offenses were retrieved for the lookback window. This could "
              "mean the environment genuinely has none, but it could also mean the offense "
              "fetch failed silently — check the warnings above before treating every rule "
              "below as 'Dead'.")

    print("\n🔍 Building rule effectiveness table across all lookback windows...")
    master_df = build_master_rule_table(rules, offenses_all)

    if master_df.empty:
        print("\n❌ No rule data to report.")
        return

    for p in LOOKBACK_PERIODS:
        vc = master_df[f"status_{p['key']}"].value_counts()
        print(f"   {p['label']:>9}: Dead={vc.get(PERIOD_STATUS_DEAD, 0):<4} "
              f"Highly Active={vc.get(PERIOD_STATUS_HIGHLY_ACTIVE, 0):<4} "
              f"Active={vc.get(PERIOD_STATUS_ACTIVE, 0)}")

    silent_count = int((master_df['overall_status'] == STATUS_RECENTLY_SILENT).sum())
    print(f"\n   ⚠️  {silent_count} rule(s) fired historically but have gone quiet recently.")

    print(f"\n💾 Saving Excel report → {OUTPUT_EXCEL}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    save_excel_report(master_df, OUTPUT_EXCEL)

    print("\n📊 Generating trend chart...")
    chart_path = generate_rule_trend_chart(master_df)
    if chart_path:
        print("   ✅ Trend chart generated.")
    else:
        print("   ⚠️  Trend chart skipped (no rule data).")

    print("\n✉️  Building email draft...")
    html_body = build_email_html(master_df, chart_cid='rule_trend' if chart_path else None)

    dead_total   = int((master_df['overall_status'] == STATUS_DEAD).sum())
    subject = (
        f"QRadar Rule Effectiveness — {dead_total} dead, "
        f"{silent_count} need investigation"
    )

    images = {'rule_trend': chart_path} if chart_path else {}
    create_outlook_draft(OUTPUT_EXCEL, subject, html_body, images)
    print("\n✅ Done!")


if __name__ == '__main__':
    main()
