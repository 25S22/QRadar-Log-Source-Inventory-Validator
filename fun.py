def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Get log source details directly from the API.
    FIX: Uses 'contains' for Name searches to allow partial matches (e.g., 'ABC' finds 'Server @ ABC').
    """
    
    # 1. CLEAN THE IDENTIFIER
    # Remove quotes and whitespace to prevent syntax errors in the query
    clean_identifier = str(identifier).replace('"', '').replace("'", "").strip()
    
    # 2. CONSTRUCT THE CORRECT FILTER
    if is_ip:
        # Strict IP search: The value must exist inside the protocol_parameters array
        query_filter = f'protocol_parameters contains value="{clean_identifier}"'
    else:
        # Partial Name search: 'contains' allows us to find "Server_NYC_1" using just "NYC_1"
        query_filter = f'name contains "{clean_identifier}"'
    
    ls_endpoint = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"

    try:
        # Get log source details with all fields
        resp = requests.get(
            ls_endpoint,
            params={'filter': query_filter},
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )
        
        if resp.status_code != 200:
            print(f"   ❌ Log source API error: {resp.status_code} - {resp.text}")
            return {'status': f'API Error {resp.status_code}', **_empty_details()}

        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', **_empty_details()}

        # 3. SELECT THE BEST MATCH
        found_source = None
        
        if is_ip:
            # For IP, verify the specific IP exists in parameters to avoid false positives
            for source in ls_data:
                params = source.get('protocol_parameters', [])
                if any(p.get('value') == clean_identifier for p in params):
                    found_source = source
                    break
            # Fallback: if strict loop fails, take the first result (best effort)
            if not found_source: 
                found_source = ls_data[0]
        else:
            # For Name, 'contains' might return multiple results. 
            # We take the first one. If you need stricter logic (e.g., shortest name), modify here.
            found_source = ls_data[0]

        ls_id = found_source.get('id')
        # Use the actual name from QRadar, not the search term
        ls_name = found_source.get('name', identifier)
        
        print(f"   📋 Found log source: {ls_name} (ID: {ls_id})")

        # Get last_event_time directly from the API response (in milliseconds)
        last_event_time_ms = found_source.get('last_event_time')
        
        # Use safe timestamp conversion
        last_seen, activity_status, days_since_last_event = safe_timestamp_conversion(last_event_time_ms)
        
        # Get additional useful fields from the API
        enabled = found_source.get('enabled', False)
        enabled_str = 'Yes' if enabled else 'No'
            
        print(f"   📊 Last Event: {last_seen} | Status: {activity_status} | Enabled: {enabled_str} | Days Since: {days_since_last_event}")

        return {
            'status': 'Found',
            'qradar_id': str(ls_id) if ls_id is not None else '',
            'enabled': enabled_str,
            'last_seen': last_seen,
            'activity_status': activity_status,
            'days_since_last_event': days_since_last_event
        }

    except Exception as e:
        print(f"   ❌ Unexpected error: {e}")
        return {'status': f'Error: {str(e)[:50]}...', **_empty_details()}
