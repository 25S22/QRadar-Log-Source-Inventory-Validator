import logging
import time
from datetime import datetime

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

RULE_LOOKBACK_DAYS = 30
DEAD_RULE_FIRE_THRESHOLD = 0
NOISE_RULE_FIRE_THRESHOLD = 500

OUTPUT_XLSX = 'rule_effectiveness_audit.xlsx'
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


def list_enabled_rules():
    endpoint = f"{QRADAR_HOST.rstrip('/')}/api/analytics/rules"
    response = _request(
        'GET',
        endpoint,
        headers={'Range': f'items=0-{RANGE_MAX}'},
        params={'filter': 'enabled=true'},
    )
    if response.status_code != 200:
        raise RuntimeError(f"Failed to list enabled rules: HTTP {response.status_code} - {response.text}")
    payload = response.json()
    if not isinstance(payload, list):
        return []
    return payload


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


def build_rule_fires_query(days):
    return f"""
SELECT
  rulename AS rule_name,
  COUNT(*) AS fire_count
FROM events
WHERE rulename IS NOT NULL
GROUP BY rule_name
ORDER BY fire_count DESC
LAST {int(days)} DAYS
"""


def normalize_rule_fire_rows(rows):
    normalized = []
    for row in rows:
        normalized.append({
            'rule_name': row.get('rule_name') or row.get('rulename') or 'Unknown',
            'fire_count': int(float(row.get('fire_count') or row.get('count(*)') or 0)),
        })
    return pd.DataFrame(normalized)


def build_audit_dataframe(enabled_rules, fire_counts_df):
    rules = []
    for rule in enabled_rules:
        rules.append({
            'rule_id': rule.get('id'),
            'rule_name': rule.get('name') or 'Unknown',
            'rule_enabled': bool(rule.get('enabled', False)),
            'owner': rule.get('owner') or '',
            'type': rule.get('type') or '',
        })
    rules_df = pd.DataFrame(rules)

    if rules_df.empty:
        rules_df = pd.DataFrame(columns=['rule_id', 'rule_name', 'rule_enabled', 'owner', 'type'])
    if fire_counts_df.empty:
        fire_counts_df = pd.DataFrame(columns=['rule_name', 'fire_count'])

    merged = rules_df.merge(fire_counts_df, on='rule_name', how='left')
    merged['fire_count'] = merged['fire_count'].fillna(0).astype(int)

    def classify(count):
        if count <= DEAD_RULE_FIRE_THRESHOLD:
            return 'dead_rule'
        if count >= NOISE_RULE_FIRE_THRESHOLD:
            return 'noise_generator'
        return 'healthy'

    merged['classification'] = merged['fire_count'].apply(classify)
    merged.sort_values(by='fire_count', ascending=False, inplace=True)
    return merged


def main():
    logger.info("Starting rule effectiveness auditor...")
    test_qradar_connection()

    enabled_rules = list_enabled_rules()
    logger.info("Enabled rules fetched: %d", len(enabled_rules))

    search_id = submit_aql_search(build_rule_fires_query(RULE_LOOKBACK_DAYS))
    logger.info("Submitted AQL search: %s", search_id)
    wait_for_search(search_id)
    fire_rows = fetch_search_results(search_id)
    fire_counts_df = normalize_rule_fire_rows(fire_rows)

    audit_df = build_audit_dataframe(enabled_rules, fire_counts_df)
    dead_df = audit_df[audit_df['classification'] == 'dead_rule'].copy()
    noise_df = audit_df[audit_df['classification'] == 'noise_generator'].copy()

    summary = pd.DataFrame([
        {'metric': 'report_generated_utc', 'value': datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')},
        {'metric': 'lookback_days', 'value': RULE_LOOKBACK_DAYS},
        {'metric': 'enabled_rules_total', 'value': len(enabled_rules)},
        {'metric': 'dead_rules_total', 'value': int(len(dead_df))},
        {'metric': 'noise_rules_total', 'value': int(len(noise_df))},
        {'metric': 'noise_threshold', 'value': NOISE_RULE_FIRE_THRESHOLD},
    ])

    with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)
        audit_df.to_excel(writer, sheet_name='All_Enabled_Rules', index=False)
        dead_df.to_excel(writer, sheet_name='Dead_Rules', index=False)
        noise_df.to_excel(writer, sheet_name='Noise_Generators', index=False)

    logger.info("Rule effectiveness audit saved: %s", OUTPUT_XLSX)
    print(audit_df.head(20).to_string(index=False))


if __name__ == '__main__':
    main()
