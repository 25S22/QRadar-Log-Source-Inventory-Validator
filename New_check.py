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
# A subfolder 'onboarding_reports' is created here automatically if it does not exist.
REPORT_OUTPUT_DIR = r'C:\path\to\your\output\folder'

# Timeout for each API request in seconds — mirrors main script
REQUEST_TIMEOUT = 30

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

# Maximum sources to fetch — there will never be more than 150 new sources
# created within a 30-day window in this environment, so no pagination needed.
_MAX_SOURCES = 150

# Valid timestamp boundaries — mirrors main script
_MIN_TS = 0
_MAX_TS = 2147483647

# QRadar status string for sources that have never forwarded a single event
_NOT_AVAILABLE = 'NOT_AVAILABLE'


# ─── REPORT FOLDER ─────────────────────────────────────────────────────────────
def _ensure_report_dir():
    """
    Creates REPORT_OUTPUT_DIR/onboarding_reports/ if it does not already exist.
    Returns the full path to the reports subfolder.
    """
    reports_dir = os.path.join(REPORT_OUTPUT_DIR, 'onboarding_reports')
    os.makedirs(reports_dir, exist_ok=True)
    return reports_dir


# ─── CONNECTION ────────────────────────────────────────────────────────────────
def test_qradar_connection():
    """
    Test QRadar connection and validate credentials before processing.
    Uses the same endpoint and header pattern as the main inventory script.
    """
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
    Uses the same header/auth pattern as the main inventory script.
    Range is capped at 150 — the maximum realistic new-source count per month.
    """
    print(f"\n📥 Fetching log sources created in the last {CREATION_LOOKBACK_DAYS} days...")

    endpoint       = f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"
    cutoff_epoch_ms = int((datetime.now() - timedelta(days=CREATION_LOOKBACK_DAYS)).timestamp() * 1000)

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
    """
    Converts a QRadar epoch millisecond timestamp to a datetime.
    Returns None if the value is missing, zero, or out of the valid range.
    Mirrors safe_timestamp_conversion from the main script.
    """
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
    """Returns a formatted date string or '—' for display in tables."""
    dt = _epoch_to_dt(epoch_ms)
    return dt.strftime('%Y-%m-%d %H:%M') if dt else '—'


# ─── CATEGORISATION ────────────────────────────────────────────────────────────
def categorise_log_sources(raw_sources):
    """
    Splits sources into three mutually exclusive buckets:

      NOT_AVAILABLE — QRadar status is NOT_AVAILABLE OR last_event_time is null/zero.
                      These have never forwarded a single event.

      STARTED       — Has a valid last_event_time but it falls OUTSIDE the lookback
                      window. Sent logs once, then went quiet.

      ACTIVE        — last_event_time is valid AND within the lookback window.
                      Actively onboarding right now.
    """
    cutoff_dt = datetime.now() - timedelta(days=CREATION_LOOKBACK_DAYS)
    buckets   = {'NOT_AVAILABLE': [], 'STARTED': [], 'ACTIVE': [], 'all': raw_sources}

    for src in raw_sources:
        # QRadar can return status as a dict or a plain string depending on API version
        raw_status    = src.get('status', '')
        qradar_status = (
            str(raw_status.get('status', '') if isinstance(raw_status, dict) else raw_status)
        ).upper()

        last_event_dt = _epoch_to_dt(src.get('last_event_time'))

        if qradar_status == _NOT_AVAILABLE or last_event_dt is None:
            buckets['NOT_AVAILABLE'].append(src)
        elif last_event_dt >= cutoff_dt:
            buckets['ACTIVE'].append(src)
        else:
            buckets['STARTED'].append(src)

    return buckets


# ─── CHART ─────────────────────────────────────────────────────────────────────
def generate_onboarding_pie_chart(counts, reports_dir, run_stamp):
    """
    Generates the onboarding status pie chart.
    Saves a timestamped copy into reports_dir (permanent record)
    and returns the path so it can be embedded in the email.
    """
    color_map = {
        'Active (Sending Logs)':         '#28a745',
        'Started (Gone Quiet)':          '#fd7e14',
        'Not Available (Never Started)': '#dc3545',
    }
    display = {
        'Active (Sending Logs)':         counts.get('ACTIVE',        0),
        'Started (Gone Quiet)':          counts.get('STARTED',       0),
        'Not Available (Never Started)': counts.get('NOT_AVAILABLE', 0),
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
    plt.title(f"Onboarding Status — Last {CREATION_LOOKBACK_DAYS} Days",
              pad=15, fontsize=11, fontweight='bold')

    filepath = os.path.join(reports_dir, f"onboarding_chart_{run_stamp}.png")
    plt.savefig(filepath, bbox_inches='tight', dpi=100)
    plt.close()

    return filepath


# ─── EMAIL HTML ────────────────────────────────────────────────────────────────
def _detail_rows(sources, badge_color, badge_label):
    if not sources:
        return """
        <tr>
          <td colspan="4"
              style="padding:12px 14px;text-align:center;
                     color:#4a5a7a;font-size:12px;font-style:italic;">
            No sources in this category.
          </td>
        </tr>"""

    rows = ""
    for i, src in enumerate(sources):
        bg = '#161d2e' if i % 2 == 0 else '#131929'
        rows += f"""
        <tr style="background:{bg};">
          <td style="padding:8px 14px;border-bottom:1px solid #1e2840;
                     font-size:12px;color:#c8d0e8;">{src.get('name','Unknown')}</td>
          <td style="padding:8px 14px;border-bottom:1px solid #1e2840;
                     font-size:12px;color:#8090b0;text-align:center;">
            {_fmt_dt(src.get('creation_date'))}
          </td>
          <td style="padding:8px 14px;border-bottom:1px solid #1e2840;
                     font-size:12px;color:#8090b0;text-align:center;">
            {_fmt_dt(src.get('last_event_time'))}
          </td>
          <td style="padding:8px 14px;border-bottom:1px solid #1e2840;
                     text-align:center;">
            <span style="background:{badge_color};color:#fff;font-size:10px;
                         font-weight:700;padding:3px 9px;border-radius:10px;
                         letter-spacing:0.5px;">{badge_label}</span>
          </td>
        </tr>"""
    return rows


def _section_table(title, accent, sources, badge_color, badge_label):
    count_label = f"{len(sources)} source{'s' if len(sources) != 1 else ''}"
    header = f"""
    <tr style="background:#1a2240;">
      <th style="padding:9px 14px;text-align:left;font-size:11px;color:#6a80b0;
                 font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                 border-bottom:2px solid {accent};">Log Source Name</th>
      <th style="padding:9px 14px;text-align:center;font-size:11px;color:#6a80b0;
                 font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                 border-bottom:2px solid {accent};">Created</th>
      <th style="padding:9px 14px;text-align:center;font-size:11px;color:#6a80b0;
                 font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                 border-bottom:2px solid {accent};">Last Event</th>
      <th style="padding:9px 14px;text-align:center;font-size:11px;color:#6a80b0;
                 font-weight:600;text-transform:uppercase;letter-spacing:0.8px;
                 border-bottom:2px solid {accent};">Status</th>
    </tr>"""

    return f"""
    <div style="background:#161d2e;border:1px solid #252f47;border-radius:10px;
                margin-bottom:20px;overflow:hidden;">
      <div style="background:{accent}18;border-left:4px solid {accent};
                  padding:12px 18px;">
        <span style="font-size:14px;font-weight:700;color:#d0d8f0;">{title}</span>
        <span style="font-size:11px;color:#6a80b0;margin-left:10px;">{count_label}</span>
      </div>
      <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
        {header}
        {_detail_rows(sources, badge_color, badge_label)}
      </table>
    </div>"""


def build_onboarding_email_html(buckets, chart_cid, run_time):
    total         = len(buckets['all'])
    active_count  = len(buckets['ACTIVE'])
    started_count = len(buckets['STARTED'])
    na_count      = len(buckets['NOT_AVAILABLE'])

    na_pct = (na_count / total * 100) if total > 0 else 0
    if na_pct == 0:
        badge_bg, badge_label = '#1a7a4a', 'ALL ONBOARDED'
    elif na_pct <= 30:
        badge_bg, badge_label = '#c87800', f'{na_count} NOT AVAILABLE'
    else:
        badge_bg, badge_label = '#c0392b', f'{na_count} NOT AVAILABLE'

    def card(label, value, accent, sub=''):
        sub_html = f'<div style="font-size:9px;color:#4a5a7a;margin-top:3px;">{sub}</div>' if sub else ''
        return f"""
        <td style="padding:5px;">
          <div style="background:#1e2535;border-left:4px solid {accent};
                      border-radius:6px;padding:14px 16px;min-width:110px;text-align:center;">
            <div style="font-size:28px;font-weight:700;color:{accent};
                        letter-spacing:-1px;line-height:1;">{value}</div>
            <div style="font-size:10px;color:#7a86a0;margin-top:5px;
                        text-transform:uppercase;letter-spacing:0.6px;">{label}</div>
            {sub_html}
          </div>
        </td>"""

    cards = (
        card('Total Created',       total,         '#4a6fa5', f'Last {CREATION_LOOKBACK_DAYS} days') +
        card('Active (Sending)',     active_count,  '#28a745') +
        card('Started / Quiet',      started_count, '#fd7e14') +
        card('Not Available',        na_count,      '#dc3545', 'Never sent logs')
    )

    chart_html = (
        f'<img src="cid:{chart_cid}" '
        f'style="display:block;max-width:360px;margin:16px auto 4px;">'
    ) if chart_cid else ''

    return f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0d1117;
             font-family:'Segoe UI',Helvetica,Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#0d1117;padding:24px 0;">
<tr><td align="center">
<table width="700" cellpadding="0" cellspacing="0" style="max-width:700px;width:100%;">

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
          </td>
          <td align="right" valign="top">
            <div style="background:{badge_bg};color:#fff;font-size:12px;
                        font-weight:700;padding:9px 18px;border-radius:20px;
                        letter-spacing:0.8px;white-space:nowrap;
                        box-shadow:0 2px 8px rgba(0,0,0,0.4);">
              ⚠&nbsp; {badge_label}
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- METRICS -->
  <tr>
    <td style="background:#131929;padding:22px 24px 18px;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:14px;">Onboarding Summary</div>
      <table cellpadding="0" cellspacing="0"><tr>{cards}</tr></table>
    </td>
  </tr>

  <!-- CHART -->
  <tr>
    <td style="background:#131929;padding:4px 24px 22px;text-align:center;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:6px;">Onboarding Health Breakdown</div>
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

  <!-- DETAIL TABLES -->
  <tr>
    <td style="background:#131929;padding:22px 24px;">
      <div style="font-size:10px;color:#3a5a9a;text-transform:uppercase;
                  letter-spacing:2px;margin-bottom:16px;">Source Detail by Category</div>
      {_section_table('🔴 Not Available — Never Sent Logs',      '#dc3545', buckets['NOT_AVAILABLE'], '#dc3545', 'NOT AVAILABLE')}
      {_section_table('🟠 Started — Gone Quiet Since Creation',  '#fd7e14', buckets['STARTED'],       '#fd7e14', 'STARTED / QUIET')}
      {_section_table('🟢 Active — Currently Sending Logs',      '#28a745', buckets['ACTIVE'],        '#28a745', 'ACTIVE')}
    </td>
  </tr>

  <!-- FOOTER -->
  <tr>
    <td style="background:#0a0f1a;border-radius:0 0 12px 12px;
               padding:16px 32px;text-align:center;border-top:1px solid #141c2e;">
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
    Embeds the chart via Content-ID — same win32com pattern as the main script.
    The chart file is NOT deleted after — it is the permanent timestamped record.
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
    """
    Saves the full email HTML as a timestamped file so every run is archived.
    File name example: onboarding_report_2025-06-01_14-32-00.html
    """
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

    # Timestamp used for both the chart filename and the HTML report filename
    run_stamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    run_time  = datetime.now().strftime('%d %B %Y  •  %H:%M:%S')

    print("🚀 Starting QRadar Onboarding Tracker...")
    print(f"   Lookback window : {CREATION_LOOKBACK_DAYS} days")
    print(f"   Enabled only    : {ENABLED_ONLY}")

    if not test_qradar_connection():
        return

    raw_sources = fetch_recent_log_sources()
    if not raw_sources:
        print("\n✅ No log sources found in the lookback window. Nothing to report.")
        return

    print("\n🔍 Categorising log sources...")
    buckets = categorise_log_sources(raw_sources)

    active_count  = len(buckets['ACTIVE'])
    started_count = len(buckets['STARTED'])
    na_count      = len(buckets['NOT_AVAILABLE'])
    total         = len(buckets['all'])

    print(f"\n📊 Results:")
    print(f"   Total Created      : {total}")
    print(f"   ✅ Active          : {active_count}  (sending logs within lookback window)")
    print(f"   🟠 Started / Quiet : {started_count}  (sent logs before, gone quiet)")
    print(f"   🔴 Not Available   : {na_count}  (never sent a single log)")

    reports_dir = _ensure_report_dir()

    chart_path = generate_onboarding_pie_chart(
        {'ACTIVE': active_count, 'STARTED': started_count, 'NOT_AVAILABLE': na_count},
        reports_dir,
        run_stamp
    )

    chart_cid  = "onboarding_chart" if chart_path else None
    html_body  = build_onboarding_email_html(buckets, chart_cid, run_time)

    # Save full HTML report with timestamp — permanent archive
    save_html_report(html_body, reports_dir, run_stamp)

    subject = (
        f"QRadar Onboarding Report — {na_count} Source{'s' if na_count != 1 else ''} "
        f"Not Sending Logs ({total} Created in Last {CREATION_LOOKBACK_DAYS} Days)"
    )

    create_onboarding_outlook_draft(subject, html_body, chart_path)

    print("\n✅ Onboarding Tracker completed!")
    print(f"   Reports folder: {reports_dir}")


if __name__ == '__main__':
    main()
