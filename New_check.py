import requests
import urllib3
import os
import win32com.client

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from datetime import datetime, timedelta

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
QRADAR_HOST     = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL      = False

# How far back to look for newly created log sources (days)
CREATION_LOOKBACK_DAYS = 30

# Only report on ENABLED log sources (set False to include disabled ones too)
ENABLED_ONLY = True

# Where timestamped HTML reports and the temp chart are saved.
# A subfolder 'onboarding_reports' is created here automatically.
REPORT_OUTPUT_DIR = r'C:\path\to\your\output\folder'

# Timeout for each API request in seconds
REQUEST_TIMEOUT = 30

# ─── HOSTNAME EXTRACTION ───────────────────────────────────────────────────────
# QRadar typically names log sources as "HOSTNAME @ IP_ADDRESS".
# The separator below is used to split the log source name and extract
# the left side as the hostname key for grouping.
#
# Examples:
#   "SERVER01 @ 10.1.2.3"          → hostname = "SERVER01"
#   "DC-PROD-01 @ 192.168.0.5"     → hostname = "DC-PROD-01"
#
# Set to '' to use the full log source name as the hostname (no grouping).
HOSTNAME_EXTRACTION_SEPARATOR = ' @ '

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

_MAX_SOURCES = 500   # raised from 150 — with grouping you may pull more sources
_MIN_TS      = 0
_MAX_TS      = 2147483647
_NOT_AVAILABLE = 'NOT_AVAILABLE'


# ─── REPORT FOLDER ─────────────────────────────────────────────────────────────
def _ensure_report_dir():
    reports_dir = os.path.join(REPORT_OUTPUT_DIR, 'onboarding_reports')
    os.makedirs(reports_dir, exist_ok=True)
    return reports_dir


# ─── CONNECTION ────────────────────────────────────────────────────────────────
def test_qradar_connection():
    print("🔗 Testing QRadar connection...")
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
            print("✅ QRadar connection successful!")
            return True
        elif resp.status_code == 401:
            print("❌ Authentication failed! Check username/password.")
            return False
        else:
            print(f"⚠️  Unexpected response: {resp.status_code}")
            return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


# ─── DATA FETCHING ─────────────────────────────────────────────────────────────
def fetch_recent_log_sources():
    """
    Fetches log sources whose creation_date falls within CREATION_LOOKBACK_DAYS.
    """
    print(f"\n📥 Fetching log sources created in the last {CREATION_LOOKBACK_DAYS} days...")

    endpoint        = (f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
                       f"/log_source_management/log_sources")
    cutoff_epoch_ms = int((datetime.now() - timedelta(days=CREATION_LOOKBACK_DAYS))
                          .timestamp() * 1000)

    api_filter = f'creation_date >= {cutoff_epoch_ms}'
    if ENABLED_ONLY:
        api_filter += ' and enabled = true'

    try:
        resp = requests.get(
            endpoint,
            params={'filter': api_filter},
            auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={
                'Accept':  'application/json',
                'Version': '14.0',
                'Range':   f'items=0-{_MAX_SOURCES - 1}'
            }
        )
        if resp.status_code == 200:
            sources = resp.json()
            print(f"✅ Retrieved {len(sources)} log source(s).")
            return sources
        elif resp.status_code == 404:
            print("✅ No log sources found in the lookback window.")
            return []
        else:
            print(f"❌ API returned {resp.status_code}. Cannot continue.")
            return []
    except Exception as e:
        print(f"❌ Error fetching log sources: {e}")
        return []


# ─── TIMESTAMP HELPER ──────────────────────────────────────────────────────────
def _epoch_to_dt(epoch_ms):
    if not epoch_ms:
        return None
    try:
        if isinstance(epoch_ms, float):
            epoch_ms = int(epoch_ms)
        epoch_s = epoch_ms / 1000.0 if epoch_ms > 4102444800 else epoch_ms
        if epoch_s <= _MIN_TS or epoch_s > _MAX_TS:
            return None
        return datetime.fromtimestamp(epoch_s)
    except Exception:
        return None


def _fmt_dt(epoch_ms):
    dt = _epoch_to_dt(epoch_ms)
    return dt.strftime('%Y-%m-%d %H:%M') if dt else '—'


# ─── HOST GROUPING & CATEGORISATION ───────────────────────────────────────────
def _extract_hostname(log_source_name):
    """
    Extracts the hostname portion from a log source name.

    QRadar default naming: "HOSTNAME @ IP_ADDRESS"
    With HOSTNAME_EXTRACTION_SEPARATOR = ' @ ', this returns "HOSTNAME".

    If the separator is absent or HOSTNAME_EXTRACTION_SEPARATOR is '',
    the full log source name is returned unchanged as the hostname key.
    """
    name = log_source_name or 'Unknown'
    if HOSTNAME_EXTRACTION_SEPARATOR and HOSTNAME_EXTRACTION_SEPARATOR in name:
        return name.split(HOSTNAME_EXTRACTION_SEPARATOR)[0].strip()
    return name.strip()


def _source_status(src, cutoff_dt):
    """
    Returns the status for a single log source:

      'ACTIVE'        — last_event_time is within the lookback window
      'STARTED'       — has a last_event_time but it's older than the window
      'NOT_AVAILABLE' — QRadar status is NOT_AVAILABLE or no events ever received
    """
    raw_status    = src.get('status', '')
    qradar_status = (
        str(raw_status.get('status', '')
            if isinstance(raw_status, dict) else raw_status)
    ).upper()

    last_event_dt = _epoch_to_dt(src.get('last_event_time'))

    if qradar_status == _NOT_AVAILABLE or last_event_dt is None:
        return 'NOT_AVAILABLE'
    elif last_event_dt >= cutoff_dt:
        return 'ACTIVE'
    else:
        return 'STARTED'


def group_and_categorise(raw_sources):
    """
    Groups all log sources by their extracted hostname, then determines a
    per-host health status:

      FULLY_ACTIVE   — every log source for this host is actively sending logs
      PARTIAL        — at least one source is active AND at least one is not
      FULLY_INACTIVE — no log source for this host has ever sent any logs

    Returns a dict:
    {
      'FULLY_ACTIVE':   [ HostGroup, ... ],
      'PARTIAL':        [ HostGroup, ... ],
      'FULLY_INACTIVE': [ HostGroup, ... ],
      'total_hosts':    int,
      'total_sources':  int,
    }

    HostGroup structure:
    {
      'hostname': str,
      'status':   'FULLY_ACTIVE' | 'PARTIAL' | 'FULLY_INACTIVE',
      'sources':  [
          { 'src': <raw dict>, 'status': 'ACTIVE' | 'STARTED' | 'NOT_AVAILABLE' },
          ...
      ]
    }
    """
    cutoff_dt = datetime.now() - timedelta(days=CREATION_LOOKBACK_DAYS)

    # Group raw sources by hostname
    host_map = {}
    for src in raw_sources:
        hostname = _extract_hostname(src.get('name', 'Unknown'))
        if hostname not in host_map:
            host_map[hostname] = []
        host_map[hostname].append({
            'src':    src,
            'status': _source_status(src, cutoff_dt),
        })

    buckets = {'FULLY_ACTIVE': [], 'PARTIAL': [], 'FULLY_INACTIVE': []}

    for hostname, sources_ws in host_map.items():
        statuses  = {s['status'] for s in sources_ws}
        has_good  = 'ACTIVE' in statuses
        has_bad   = 'NOT_AVAILABLE' in statuses or 'STARTED' in statuses

        if has_good and has_bad:
            host_status = 'PARTIAL'
        elif has_good:
            host_status = 'FULLY_ACTIVE'
        else:
            host_status = 'FULLY_INACTIVE'

        # Sort individual sources: active first, then started, then not_available
        order = {'ACTIVE': 0, 'STARTED': 1, 'NOT_AVAILABLE': 2}
        sources_ws.sort(key=lambda x: order.get(x['status'], 9))

        buckets[host_status].append({
            'hostname': hostname,
            'status':   host_status,
            'sources':  sources_ws,
        })

    # Sort each bucket alphabetically by hostname
    for key in buckets:
        buckets[key].sort(key=lambda h: h['hostname'].lower())

    return {
        **buckets,
        'total_hosts':   len(host_map),
        'total_sources': len(raw_sources),
    }


# ─── CHART ─────────────────────────────────────────────────────────────────────
def generate_onboarding_pie_chart(host_counts, reports_dir, run_stamp):
    """
    Generates a host-level pie chart.
    Three slices: Fully Active, Partial (some not reporting), Fully Inactive.
    """
    color_map = {
        'Fully Active':                  '#28a745',
        'Partial (Some Not Reporting)':  '#fd7e14',
        'Fully Inactive (No Events)':    '#dc3545',
    }
    display = {
        'Fully Active':                  host_counts.get('FULLY_ACTIVE',   0),
        'Partial (Some Not Reporting)':  host_counts.get('PARTIAL',        0),
        'Fully Inactive (No Events)':    host_counts.get('FULLY_INACTIVE', 0),
    }
    filtered = {k: v for k, v in display.items() if v > 0}
    if not filtered:
        return None

    labels = list(filtered.keys())
    sizes  = list(filtered.values())
    colors = [color_map[l] for l in labels]

    fig, ax = plt.subplots(figsize=(5, 4))
    ax.pie(
        sizes, labels=labels, colors=colors, autopct='%1.1f%%',
        startangle=140, textprops={'fontsize': 9},
        wedgeprops={'edgecolor': 'white'}
    )
    ax.axis('equal')
    plt.title(
        f"New Host Onboarding Health — Last {CREATION_LOOKBACK_DAYS} Days",
        pad=15, fontsize=11, fontweight='bold'
    )

    filepath = os.path.join(reports_dir, f"onboarding_chart_{run_stamp}.png")
    plt.savefig(filepath, bbox_inches='tight', dpi=100)
    plt.close()
    return filepath


# ─── EMAIL HTML ────────────────────────────────────────────────────────────────
_SOURCE_STATUS_META = {
    'ACTIVE':        {'color': '#28a745', 'label': 'ACTIVE',          'name_color': '#a8e8b8'},
    'STARTED':       {'color': '#fd7e14', 'label': 'STARTED / QUIET', 'name_color': '#e8c090'},
    'NOT_AVAILABLE': {'color': '#dc3545', 'label': 'NOT AVAILABLE',   'name_color': '#e8a0a0'},
}

_HOST_STATUS_META = {
    'FULLY_ACTIVE':   {'accent': '#28a745', 'row_bg': '#0a160c',
                       'label': 'FULLY ACTIVE',   'label_color': '#28a745'},
    'PARTIAL':        {'accent': '#fd7e14', 'row_bg': '#181208',
                       'label': 'PARTIAL',         'label_color': '#fd7e14'},
    'FULLY_INACTIVE': {'accent': '#dc3545', 'row_bg': '#160808',
                       'label': 'NO EVENTS',       'label_color': '#dc3545'},
}


def _host_group_rows(host_group):
    """
    Returns the HTML for one host: a bold host header row followed by
    indented source rows for each log source belonging to that host.
    """
    hostname    = host_group['hostname']
    sources_ws  = host_group['sources']
    host_status = host_group['status']
    meta        = _HOST_STATUS_META[host_status]

    active_count  = sum(1 for s in sources_ws if s['status'] == 'ACTIVE')
    total_src     = len(sources_ws)
    accent        = meta['accent']

    if host_status == 'FULLY_ACTIVE':
        summary_txt = f"All {total_src} source{'s' if total_src != 1 else ''} active"
    elif host_status == 'PARTIAL':
        summary_txt = (f"{active_count} of {total_src} source"
                       f"{'s' if total_src != 1 else ''} active")
    else:
        summary_txt = f"0 of {total_src} sources sending events"

    # Host header row
    html = f"""
        <tr style="background:{meta['row_bg']};">
          <td colspan="4"
              style="padding:10px 14px;border-top:2px solid {accent}40;
                     border-bottom:1px solid #252f47;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td>
                  <span style="font-size:12px;font-weight:700;color:#d0d8f0;">
                    🖥&nbsp; {hostname}
                  </span>
                  <span style="font-size:10px;color:#6a80b0;margin-left:10px;">
                    {summary_txt}
                  </span>
                </td>
                <td align="right" style="white-space:nowrap;">
                  <span style="background:{meta['label_color']}22;
                               color:{meta['label_color']};font-size:9px;
                               font-weight:700;padding:3px 10px;
                               border-radius:8px;letter-spacing:0.5px;
                               border:1px solid {meta['label_color']}55;">
                    {meta['label']}
                  </span>
                  &nbsp;
                  <span style="background:#252f47;color:#6a80b0;font-size:9px;
                               font-weight:600;padding:3px 9px;border-radius:8px;">
                    {total_src} SOURCE{'S' if total_src != 1 else ''}
                  </span>
                </td>
              </tr>
            </table>
          </td>
        </tr>"""

    # Individual source rows
    for i, s in enumerate(sources_ws):
        src      = s['src']
        s_status = s['status']
        s_meta   = _SOURCE_STATUS_META[s_status]
        row_bg   = '#161d2e' if i % 2 == 0 else '#131929'

        html += f"""
        <tr style="background:{row_bg};">
          <td style="padding:7px 14px 7px 30px;border-bottom:1px solid #1e2840;
                     font-size:11px;color:{s_meta['name_color']};">
            <span style="color:#3a5a9a;margin-right:6px;">↳</span>
            {src.get('name', 'Unknown')}
          </td>
          <td style="padding:7px 14px;border-bottom:1px solid #1e2840;
                     font-size:11px;color:#8090b0;text-align:center;">
            {_fmt_dt(src.get('creation_date'))}
          </td>
          <td style="padding:7px 14px;border-bottom:1px solid #1e2840;
                     font-size:11px;color:#8090b0;text-align:center;">
            {_fmt_dt(src.get('last_event_time'))}
          </td>
          <td style="padding:7px 14px;border-bottom:1px solid #1e2840;
                     text-align:center;">
            <span style="background:{s_meta['color']};color:#fff;font-size:9px;
                         font-weight:700;padding:2px 8px;border-radius:8px;
                         letter-spacing:0.3px;">{s_meta['label']}</span>
          </td>
        </tr>"""

    return html


def _section_block(title, accent, host_groups, section_status):
    """
    Builds one collapsible-style section card containing all hosts
    belonging to a given health tier.
    """
    host_count   = len(host_groups)
    source_count = sum(len(hg['sources']) for hg in host_groups)
    count_label  = (f"{host_count} host{'s' if host_count != 1 else ''}"
                    f" · {source_count} log source{'s' if source_count != 1 else ''}")

    if not host_groups:
        return f"""
    <div style="background:#161d2e;border:1px solid #252f47;border-radius:10px;
                margin-bottom:20px;overflow:hidden;">
      <div style="background:{accent}18;border-left:4px solid {accent};
                  padding:12px 18px;">
        <span style="font-size:14px;font-weight:700;color:#d0d8f0;">{title}</span>
        <span style="font-size:11px;color:#3a4a6a;margin-left:10px;">0 hosts</span>
      </div>
      <div style="padding:16px;text-align:center;color:#3a4a6a;
                  font-size:12px;font-style:italic;">No hosts in this category.</div>
    </div>"""

    all_rows = ''.join(_host_group_rows(hg) for hg in host_groups)

    return f"""
    <div style="background:#161d2e;border:1px solid #252f47;border-radius:10px;
                margin-bottom:20px;overflow:hidden;">
      <div style="background:{accent}18;border-left:4px solid {accent};
                  padding:12px 18px;">
        <span style="font-size:14px;font-weight:700;color:#d0d8f0;">{title}</span>
        <span style="font-size:11px;color:#6a80b0;margin-left:10px;">{count_label}</span>
      </div>
      <table width="100%" cellpadding="0" cellspacing="0"
             style="border-collapse:collapse;">
        <tr style="background:#1a2240;">
          <th style="padding:8px 14px 8px 30px;text-align:left;font-size:10px;
                     color:#6a80b0;font-weight:600;text-transform:uppercase;
                     letter-spacing:0.8px;border-bottom:2px solid {accent};">
            Host / Log Source</th>
          <th style="padding:8px 14px;text-align:center;font-size:10px;
                     color:#6a80b0;font-weight:600;text-transform:uppercase;
                     letter-spacing:0.8px;border-bottom:2px solid {accent};">
            Created</th>
          <th style="padding:8px 14px;text-align:center;font-size:10px;
                     color:#6a80b0;font-weight:600;text-transform:uppercase;
                     letter-spacing:0.8px;border-bottom:2px solid {accent};">
            Last Event</th>
          <th style="padding:8px 14px;text-align:center;font-size:10px;
                     color:#6a80b0;font-weight:600;text-transform:uppercase;
                     letter-spacing:0.8px;border-bottom:2px solid {accent};">
            Status</th>
        </tr>
        {all_rows}
      </table>
    </div>"""


def build_onboarding_email_html(data, chart_cid, run_time):
    """
    Builds the complete HTML email body using host-grouped data.

    Summary cards now show HOST counts (not raw log source counts)
    so the numbers reflect unique new assets, not individual source entries.
    """
    total_hosts    = data['total_hosts']
    total_sources  = data['total_sources']
    active_hosts   = len(data['FULLY_ACTIVE'])
    partial_hosts  = len(data['PARTIAL'])
    inactive_hosts = len(data['FULLY_INACTIVE'])

    # Header badge — driven by host health, not raw source counts
    problem_hosts = partial_hosts + inactive_hosts
    if problem_hosts == 0:
        badge_bg, badge_label = '#1a7a4a', 'ALL HOSTS REPORTING'
    elif inactive_hosts == 0:
        badge_bg, badge_label = '#c87800', f'{partial_hosts} HOST{"S" if partial_hosts != 1 else ""} PARTIAL'
    else:
        badge_bg, badge_label = '#c0392b', f'{problem_hosts} HOST{"S" if problem_hosts != 1 else ""} NOT FULLY REPORTING'

    def card(label, value, accent, sub=''):
        sub_html = (f'<div style="font-size:9px;color:#4a5a7a;margin-top:3px;">'
                    f'{sub}</div>') if sub else ''
        return f"""
        <td style="padding:5px;">
          <div style="background:#1e2535;border-left:4px solid {accent};
                      border-radius:6px;padding:14px 16px;min-width:110px;
                      text-align:center;">
            <div style="font-size:28px;font-weight:700;color:{accent};
                        letter-spacing:-1px;line-height:1;">{value}</div>
            <div style="font-size:10px;color:#7a86a0;margin-top:5px;
                        text-transform:uppercase;letter-spacing:0.6px;">{label}</div>
            {sub_html}
          </div>
        </td>"""

    cards = (
        card('New Hosts',          total_hosts,    '#4a6fa5', f'Last {CREATION_LOOKBACK_DAYS} days') +
        card('Fully Active',       active_hosts,   '#28a745', 'All sources OK') +
        card('Partial',            partial_hosts,  '#fd7e14', 'Some sources silent') +
        card('Fully Inactive',     inactive_hosts, '#dc3545', 'No events received')
    )

    # Sub-note showing total log sources vs unique hosts
    sources_note = (
        f"{total_sources} log source{'s' if total_sources != 1 else ''} across "
        f"{total_hosts} unique host{'s' if total_hosts != 1 else ''}"
    )

    chart_html = (
        f'<img src="cid:{chart_cid}" '
        f'style="display:block;max-width:360px;margin:16px auto 4px;">'
    ) if chart_cid else ''

    detail_sections = (
        _section_block(
            '🔴 Fully Inactive — No Events Ever Received',
            '#dc3545', data['FULLY_INACTIVE'], 'FULLY_INACTIVE'
        ) +
        _section_block(
            '🟠 Partial — Some Log Sources Not Reporting',
            '#fd7e14', data['PARTIAL'], 'PARTIAL'
        ) +
        _section_block(
            '🟢 Fully Active — All Log Sources Reporting',
            '#28a745', data['FULLY_ACTIVE'], 'FULLY_ACTIVE'
        )
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0d1117;
             font-family:'Segoe UI',Helvetica,Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#0d1117;padding:24px 0;">
<tr><td align="center">
<table width="720" cellpadding="0" cellspacing="0" style="max-width:720px;width:100%;">

  <!-- HEADER -->
  <tr>
    <td style="background:linear-gradient(135deg,#0c1628 0%,#172040 60%,#1a2a50 100%);
               border-radius:12px 12px 0 0;padding:28px 32px 24px;
               border-bottom:2px solid #2a3d6b;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <div style="font-size:10px;color:#3a5a9a;letter-spacing:3px;
                        text-transform:uppercase;margin-bottom:8px;">
              Automated SOC Intelligence Report
            </div>
            <div style="font-size:24px;font-weight:700;color:#e8ecf4;
                        letter-spacing:-0.5px;line-height:1.25;">
              QRadar Log Source<br>Onboarding Tracker
            </div>
            <div style="margin-top:10px;font-size:11px;color:#4a6590;">
              New sources created in the last
              <strong style="color:#6a8ac0;">{CREATION_LOOKBACK_DAYS} days</strong>
              &nbsp;·&nbsp; {run_time}
            </div>
            <div style="margin-top:4px;font-size:10px;color:#3a5070;">
              {sources_note}
            </div>
          </td>
          <td align="right" valign="top">
            <div style="background:{badge_bg};color:#fff;font-size:11px;
                        font-weight:700;padding:9px 18px;border-radius:20px;
                        letter-spacing:0.8px;white-space:nowrap;
                        box-shadow:0 2px 8px rgba(0,0,0,0.4);">
              {badge_label}
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- METRICS — host-level counts -->
  <tr>
    <td style="background:#131929;padding:22px 24px 18px;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:14px;">
        Host Onboarding Summary
      </div>
      <table cellpadding="0" cellspacing="0"><tr>{cards}</tr></table>
    </td>
  </tr>

  <!-- CHART -->
  <tr>
    <td style="background:#131929;padding:4px 24px 22px;text-align:center;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:6px;">
        Host Onboarding Health Breakdown
      </div>
      {chart_html}
    </td>
  </tr>

  <!-- DIVIDER -->
  <tr>
    <td style="padding:0;">
      <div style="height:1px;background:linear-gradient(90deg,
                  #0d1117,#2a3d6b 30%,#2a3d6b 70%,#0d1117);"></div>
    </td>
  </tr>

  <!-- DETAIL — grouped by host -->
  <tr>
    <td style="background:#131929;padding:22px 24px;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:4px;">
        Host Detail by Category
      </div>
      <div style="font-size:10px;color:#3a4a6a;margin-bottom:16px;">
        Each host is expanded to show all its individual log sources.
        Partial hosts flag which specific sources are silent.
      </div>
      {detail_sections}
    </td>
  </tr>

  <!-- FOOTER -->
  <tr>
    <td style="background:#0a0f1a;border-radius:0 0 12px 12px;
               padding:16px 32px;text-align:center;
               border-top:1px solid #141c2e;">
      <div style="font-size:10px;color:#2a3a5a;letter-spacing:0.5px;">
        Automated Cyber Defense Reporting &nbsp;·&nbsp; Generated {run_time}
      </div>
    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""


# ─── OUTLOOK DRAFT ─────────────────────────────────────────────────────────────
def create_onboarding_outlook_draft(subject, html_body, chart_path):
    """
    Creates an Outlook draft with the onboarding report.
    Embeds the chart via Content-ID. Draft only — nothing is sent automatically.
    """
    chart_cid = "onboarding_chart"
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail    = outlook.CreateItem(0)
        mail.Subject = subject

        if chart_path and os.path.exists(chart_path):
            attachment = mail.Attachments.Add(chart_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", chart_cid
            )

        mail.HTMLBody = html_body
        mail.Display()
        print("\n✉️  Onboarding email draft created successfully.")

    except Exception as e:
        print(f"\n❌ Failed to create Outlook draft: {e}")


def save_html_report(html_body, reports_dir, run_stamp):
    filepath = os.path.join(reports_dir, f"onboarding_report_{run_stamp}.html")
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html_body)
        print(f"💾 HTML report saved: {filepath}")
    except Exception as e:
        print(f"⚠️  Could not save HTML report: {e}")


# ─── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    run_stamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    run_time  = datetime.now().strftime('%d %B %Y  •  %H:%M:%S')

    print("🚀 Starting QRadar Onboarding Tracker...")
    print(f"   Lookback window      : {CREATION_LOOKBACK_DAYS} days")
    print(f"   Enabled only         : {ENABLED_ONLY}")
    print(f"   Hostname separator   : '{HOSTNAME_EXTRACTION_SEPARATOR}'")

    if not test_qradar_connection():
        return

    raw_sources = fetch_recent_log_sources()
    if not raw_sources:
        print("\n✅ No log sources found in the lookback window. Nothing to report.")
        return

    print("\n🔍 Grouping log sources by host and categorising...")
    data = group_and_categorise(raw_sources)

    total_hosts    = data['total_hosts']
    total_sources  = data['total_sources']
    active_hosts   = len(data['FULLY_ACTIVE'])
    partial_hosts  = len(data['PARTIAL'])
    inactive_hosts = len(data['FULLY_INACTIVE'])

    print(f"\n📊 Results:")
    print(f"   Total log sources    : {total_sources}")
    print(f"   Unique new hosts     : {total_hosts}")
    print(f"   ✅ Fully Active      : {active_hosts}  (all sources sending logs)")
    print(f"   🟠 Partial           : {partial_hosts}  (some sources silent)")
    print(f"   🔴 Fully Inactive    : {inactive_hosts}  (no sources ever sent logs)")

    if partial_hosts:
        print(f"\n⚠️  Partial hosts:")
        for hg in data['PARTIAL']:
            active_count = sum(1 for s in hg['sources'] if s['status'] == 'ACTIVE')
            total_src    = len(hg['sources'])
            print(f"     {hg['hostname']} — {active_count}/{total_src} sources active")
            for s in hg['sources']:
                icon = '✔' if s['status'] == 'ACTIVE' else '✖'
                print(f"       {icon} {s['src'].get('name','?')} [{s['status']}]")

    reports_dir = _ensure_report_dir()

    chart_path = generate_onboarding_pie_chart(
        {
            'FULLY_ACTIVE':   active_hosts,
            'PARTIAL':        partial_hosts,
            'FULLY_INACTIVE': inactive_hosts,
        },
        reports_dir,
        run_stamp
    )

    chart_cid = "onboarding_chart" if chart_path else None
    html_body = build_onboarding_email_html(data, chart_cid, run_time)

    save_html_report(html_body, reports_dir, run_stamp)

    subject = (
        f"QRadar Onboarding Report — {total_hosts} New Host"
        f"{'s' if total_hosts != 1 else ''} · "
        f"{inactive_hosts + partial_hosts} Not Fully Reporting "
        f"(Last {CREATION_LOOKBACK_DAYS} Days)"
    )

    create_onboarding_outlook_draft(subject, html_body, chart_path)

    print("\n✅ Onboarding Tracker completed!")
    print(f"   Reports folder: {reports_dir}")


if __name__ == '__main__':
    main()
