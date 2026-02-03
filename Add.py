def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column, in_qradar_col):
    """
    Optimized Sheet Processing:
    1. CLEANS HEADERS to remove hidden spaces.
    2. Filters DataFrame for 'Yes' rows first.
    3. Iterates ONLY relevant rows.
    4. Provides detailed console output.
    """
    print(f"\n{'='*60}")
    print(f"📋 Processing Sheet: {sheet_name}")
    print(f"{'='*60}")
    
    # ─── CRITICAL FIX: CLEAN COLUMN HEADERS ───
    # Removes hidden spaces (e.g., "In Qradar? " -> "In Qradar?")
    if not df.empty:
        df.columns = df.columns.str.strip()

    # Initialize columns if missing
    cols_to_init = {
        'status': 'object',
        'qradar_id': 'object', 
        'enabled': 'object',
        'last_seen': 'object',
        'activity_status': 'object',
        'days_since_last_event': 'float64',
        'remarks': 'object'
    }
    
    for col, dtype in cols_to_init.items():
        if col not in df.columns:
            df[col] = pd.Series(dtype=dtype)
    
    # ─── STEP 1: EFFICIENT FILTERING ───
    if in_qradar_col not in df.columns:
        print(f"❌ ERROR: Column '{in_qradar_col}' not found in '{sheet_name}'!")
        print(f"   Available Columns: {list(df.columns)}")
        print("   Skipping this sheet.")
        return df

    # Create a mask for rows where In Qradar contains "yes" (case-insensitive)
    process_mask = df[in_qradar_col].astype(str).str.lower().str.contains("yes", na=False)
    
    rows_to_process = df[process_mask]
    total_rows = len(df)
    target_count = len(rows_to_process)
    
    print(f"📊 Total Rows: {total_rows}")
    print(f"🎯 Rows marked 'Yes': {target_count} (Rows to scan)")
    
    # Mark skipped rows efficiently
    df.loc[~process_mask, 'remarks'] = "Skipped (In Qradar != Yes)"
    
    if target_count == 0:
        print("✅ No rows to process in this sheet.")
        return df

    # ─── STEP 2: ITERATE ONLY TARGET ROWS ───
    current_idx = 0
    
    for idx, row in rows_to_process.iterrows():
        current_idx += 1
        name_val = str(row[logsource_column]).strip()
        ip_val = str(row[ip_column]).strip()
        
        # Clean cleanup
        if name_val.lower() in ['nan', 'none', '', 'null']: name_val = None
        if ip_val.lower() in ['nan', 'none', '', 'null']: ip_val = None
        
        print(f"\n🔹 [{current_idx}/{target_count}] Processing Row {idx+1}")
        
        details = None
        search_method = "None"
        
        # 1. Search by Name
        if name_val:
            print(f"   🔍 Searching Name: '{name_val}' ... ", end="")
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)
            
            if details['status'] == 'Found':
                print("✅ Found!")
                search_method = "Name"
            else:
                print("⚠️ Not Found.")
        
        # 2. Fallback to IP (Only if Name failed/missing)
        if (not details or details['status'] == 'Not Found') and ip_val:
            print(f"   🔁 Fallback to IP: '{ip_val}' ... ", end="")
            details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)
            
            if details['status'] == 'Found':
                print("✅ Found!")
                search_method = "IP"
            else:
                print("❌ Not Found.")

        # Initialize if completely failed
        if not details:
            details = {'status': 'Empty/Invalid', **_empty_details()}
            
        # ─── STEP 3: UPDATE DATAFRAME & DISPLAY DETAILS ───
        
        if details['status'] == 'Found':
            # Update Data
            df.at[idx, 'status'] = 'Found'
            df.at[idx, 'remarks'] = f"Found by {search_method}"
            df.at[idx, 'qradar_id'] = details['qradar_id']
            df.at[idx, 'enabled'] = details['enabled']
            df.at[idx, 'last_seen'] = details['last_seen']
            df.at[idx, 'activity_status'] = details['activity_status']
            df.at[idx, 'days_since_last_event'] = details['days_since_last_event']
            
            # Print Details
            # Ensure actual_name key exists (was previously called ls_name in getter)
            display_name = details.get('actual_name', name_val) 
            print(f"      📌 Log Source: {display_name}")
            print(f"      📊 Activity:   {details['activity_status']}")
            print(f"      📅 Last Event: {details['last_seen']} ({details['days_since_last_event']} days ago)")
            
        else:
            # Handle Not Found Logic
            status_val = "Not Found"
            remark_val = "❌ Not found by Name or IP"
            act_val = "Not Found"
            
            # Smart Inference
            if name_val and "AP" in name_val:
                status_val = "Inferred"
                remark_val = "Under WLC (Inferred)"
                act_val = "Inferred"
                print("      ℹ️  Result: Inferred as WLC (AP in name)")
            elif name_val and "FW" in name_val:
                status_val = "Inferred"
                remark_val = "Under Forti (Inferred)"
                act_val = "Inferred"
                print("      ℹ️  Result: Inferred as FortiGate (FW in name)")
            else:
                print("      ❌ Result: ABSENT in QRadar.")

            df.at[idx, 'status'] = status_val
            df.at[idx, 'remarks'] = remark_val
            df.at[idx, 'activity_status'] = act_val
            
            # Clear other fields to be clean
            df.at[idx, 'last_seen'] = "N/A"
            df.at[idx, 'days_since_last_event'] = None

        time.sleep(0.2) # Slight delay to be nice to API
    
    return df
