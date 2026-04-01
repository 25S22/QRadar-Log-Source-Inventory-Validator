import logging
import math
import time
from datetime import datetime, timedelta

import pandas as pd
import requests
import urllib3

# ─── LOGGING SETUP ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)
logger = logging.getLogger(__name__)

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL = False

REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
RETRY_DELAY_BASE = 1.5
RANGE_MAX = 9999

EPS_LOOKBACK_DAYS = 7
TOP_N_SOURCES = 100

# Set this to your licensed cap to enable "days until cap" projection.
LICENSE_CAP_EPS = 5000.0

OUTPUT_XLSX = 'eps_burn_rate_report.xlsx'
# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

if not VERIFY_SSL:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def _request(method, url, **kwargs):
    for attempt in range(MAX_RETRIES):
        try:
            response = requests.request(
                method=method,
                url=url,
                auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
                verify=VERIFY_SSL,
                timeout=REQUEST_TIMEOUT,
                headers={
                    'Accept': 'application/json',
                    'Version': '14.0',
                    **kwargs.pop('headers', {}),
                },
                **kwargs,
            )
            return response
        except requests.RequestException as exc:
            if attempt == MAX_RETRIES - 1:
                raise
            delay = RETRY_DELAY_BASE * (2 ** attempt)
            logger.warning("Request failed (%s). Retrying in %.1fs...", exc, delay)
            time.sleep(delay)


def test_qradar_connection():
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/help/versions"
    response = _request('GET', endpoint)
    if response.status_code != 200:
        raise RuntimeError(f"Connection test failed: HTTP {response.status_code} - {response.text}")
    logger.info("QRadar connection successful.")


def submit_aql_search(query):
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/ariel/searches"
    response = _request('POST', endpoint, params={'query_expression': query})
    if response.status_code not in (200, 201):
        raise RuntimeError(f"Failed to submit AQL search: HTTP {response.status_code} - {response.text}")
    payload = response.json()
    search_id = payload.get('search_id')
    if not search_id:
        raise RuntimeError("AQL submission did not return search_id.")
    return search_id


def wait_for_search(search_id):
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/ariel/searches/{search_id}"
    while True:
        response = _request('GET', endpoint)
        if response.status_code != 200:
            raise RuntimeError(f"Failed while polling search {search_id}: HTTP {response.status_code} - {response.text}")
        payload = response.json()
        status = payload.get('status', '').upper()
        if status in {'COMPLETED', 'EXECUTE', 'DONE'}:
            return
        if status in {'CANCELED', 'ERROR', 'FAILED'}:
            raise RuntimeError(f"AQL search {search_id} failed with status {status}.")
        time.sleep(2)


def fetch_search_results(search_id):
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/ariel/searches/{search_id}/results"
    response = _request('GET', endpoint)
    if response.status_code != 200:
        raise RuntimeError(f"Failed to fetch search results: HTTP {response.status_code} - {response.text}")
    payload = response.json()
    rows = payload.get('events') or payload.get('flows') or payload.get('results') or []
    if not isinstance(rows, list):
        return []
    return rows


def _to_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _trend_arrow(current_value, previous_value):
    if current_value > previous_value * 1.05:
        return '↑'
    if current_value < previous_value * 0.95:
        return '↓'
    return '→'


def _project_days_until_cap(current_total_eps, previous_total_eps, period_days):
    if not LICENSE_CAP_EPS or LICENSE_CAP_EPS <= 0:
        return None
    if current_total_eps >= LICENSE_CAP_EPS:
        return 0.0
    # current_total_eps and previous_total_eps are period averages (EPS), so we
    # convert their delta into a daily slope using the actual period length.
    daily_growth = (current_total_eps - previous_total_eps) / max(float(period_days), 1.0)
    if daily_growth <= 0:
        return math.inf
    return (LICENSE_CAP_EPS - current_total_eps) / daily_growth


def _projection_text(projected_days):
    if projected_days == 0:
        return "Cap already reached"
    if projected_days == math.inf:
        return "No projected cap (stable/decreasing)"
    return f"{projected_days:.1f}"


def build_eps_query(days):
    return f"""
SELECT
  LOGSOURCENAME(logsourceid) AS log_source_name,
  logsourceid AS log_source_id,
  DATEFORMAT(starttime, 'yyyy-MM-dd') AS day_bucket,
  COUNT(*) / 86400.0 AS avg_eps
FROM events
WHERE logsourceid IS NOT NULL
GROUP BY log_source_id, log_source_name, day_bucket
ORDER BY avg_eps DESC
LAST {int(days)} DAYS
"""


def normalize_rows(rows):
    normalized = []
    for row in rows:
        normalized.append({
            'log_source_id': row.get('log_source_id') or row.get('logsourceid'),
            'log_source_name': row.get('log_source_name') or row.get('logsourcename(logsourceid)') or 'Unknown',
            'day_bucket': row.get('day_bucket') or row.get("dateformat(starttime, 'yyyy-MM-dd')"),
            'avg_eps': _to_float(row.get('avg_eps') or row.get('count(*) / 86400.0') or row.get('count') or 0),
        })
    return pd.DataFrame(normalized)


def build_report_dataframe(df_daily):
    if df_daily.empty:
        return pd.DataFrame(columns=[
            'log_source_id', 'log_source_name', 'current_week_avg_eps', 'previous_week_avg_eps',
            'trend', 'delta_eps', 'license_share_percent',
        ])

    df_daily['day_bucket'] = pd.to_datetime(df_daily['day_bucket'], errors='coerce')
    now_utc = datetime.utcnow()
    current_start = now_utc - timedelta(days=EPS_LOOKBACK_DAYS)
    previous_start = now_utc - timedelta(days=EPS_LOOKBACK_DAYS * 2)

    current_df = df_daily[(df_daily['day_bucket'] >= current_start) & (df_daily['day_bucket'] <= now_utc)]
    previous_df = df_daily[(df_daily['day_bucket'] >= previous_start) & (df_daily['day_bucket'] < current_start)]

    current_agg = current_df.groupby(['log_source_id', 'log_source_name'], as_index=False)['avg_eps'].mean()
    current_agg.rename(columns={'avg_eps': 'current_week_avg_eps'}, inplace=True)

    previous_agg = previous_df.groupby(['log_source_id', 'log_source_name'], as_index=False)['avg_eps'].mean()
    previous_agg.rename(columns={'avg_eps': 'previous_week_avg_eps'}, inplace=True)

    merged = current_agg.merge(
        previous_agg,
        on=['log_source_id', 'log_source_name'],
        how='left',
    )
    merged['previous_week_avg_eps'] = merged['previous_week_avg_eps'].fillna(0.0)
    merged['delta_eps'] = merged['current_week_avg_eps'] - merged['previous_week_avg_eps']
    merged['trend'] = merged.apply(
        lambda r: _trend_arrow(r['current_week_avg_eps'], r['previous_week_avg_eps']),
        axis=1,
    )

    total_current = merged['current_week_avg_eps'].sum()
    if total_current > 0:
        merged['license_share_percent'] = (merged['current_week_avg_eps'] / total_current) * 100.0
    else:
        merged['license_share_percent'] = 0.0

    merged.sort_values(by='current_week_avg_eps', ascending=False, inplace=True)
    return merged.head(TOP_N_SOURCES).reset_index(drop=True)


def main():
    logger.info("Starting EPS burn-rate monitor...")
    test_qradar_connection()

    search_id = submit_aql_search(build_eps_query(EPS_LOOKBACK_DAYS * 2))
    logger.info("Submitted AQL search: %s", search_id)
    wait_for_search(search_id)
    rows = fetch_search_results(search_id)

    df_daily = normalize_rows(rows)
    df_ranked = build_report_dataframe(df_daily)

    current_total_eps = _to_float(df_ranked['current_week_avg_eps'].sum()) if not df_ranked.empty else 0.0
    previous_total_eps = _to_float(df_ranked['previous_week_avg_eps'].sum()) if not df_ranked.empty else 0.0
    projected_days = _project_days_until_cap(
        current_total_eps,
        previous_total_eps,
        period_days=EPS_LOOKBACK_DAYS,
    )
    projection_text = _projection_text(projected_days)

    summary = pd.DataFrame([
        {'metric': 'report_generated_utc', 'value': datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')},
        {'metric': 'current_week_total_avg_eps', 'value': round(current_total_eps, 3)},
        {'metric': 'previous_week_total_avg_eps', 'value': round(previous_total_eps, 3)},
        {'metric': 'license_cap_eps', 'value': LICENSE_CAP_EPS},
        {'metric': 'days_until_cap_projection', 'value': projection_text},
    ])

    with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)
        df_ranked.to_excel(writer, sheet_name='Source_Ranking', index=False)

    logger.info("EPS burn-rate report saved: %s", OUTPUT_XLSX)
    print(df_ranked.head(20).to_string(index=False))


if __name__ == '__main__':
    main()
