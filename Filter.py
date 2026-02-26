# ─── EFFICIENT FILTERING MASKS ───
    in_qradar_series = df[in_qradar_col].astype(str).str.lower()
    process_mask = in_qradar_series.str.contains("yes", na=False)
    pending_mask = in_qradar_series.str.contains("pending-maintenance", na=False)
    
    rows_to_process = df[process_mask]
    total_rows = len(df)
    target_count = len(rows_to_process)
    pending_count = pending_mask.sum()
    
    print(f"📊 Total Rows: {total_rows} | 🎯 'Yes' (To Scan): {target_count} | ⏳ Pending: {pending_count}")
    
    # Mark skipped rows clearly (rows that are neither 'Yes' nor 'Pending-Maintenance')
    skipped_mask = ~(process_mask | pending_mask)
    df.loc[skipped_mask, 'remarks'] = "Skipped (Not Yes or Pending)"
    
    # Pre-fill Pending-Maintenance rows so they bypass the thread pool entirely
    if pending_count > 0:
        df.loc[pending_mask, 'status'] = 'Pending-Maintenance'
        df.loc[pending_mask, 'activity_status'] = 'Pending-Maintenance'
        df.loc[pending_mask, 'remarks'] = 'Pending Maintenance (Not Scanned)'
        df.loc[pending_mask, 'last_seen'] = 'N/A'
