def filter_and_email(df_dict, draft_path):
    """
    Filter inactive/error log sources and create report.
    UPDATED: Excludes 'Inferred' (WLC/Forti) items from the Excel attachment and Issue Counts.
    """
    frames = []
    
    # Global Counters for Email Body Context
    stats = {
        'total_scanned': 0,
        'found_active': 0,
        'found_inactive': 0,
        'inferred': 0,
        'not_found': 0,
        'errors': 0
    }

    for name, df in df_dict.items():
        if 'status' not in df.columns: continue
        
        # Only process rows that weren't skipped
        processed_df = df[df['status'].notna()]
        stats['total_scanned'] += len(processed_df)
        
        # --- Calculate Counts (For Email Body Text) ---
        stats['found_active'] += len(processed_df[(processed_df['status'] == 'Found') & (processed_df['activity_status'] == 'Active')])
        
        # Identify Inactive
        mask_inactive = (processed_df['status'] == 'Found') & ((processed_df['activity_status'] == 'Inactive') | (processed_df['activity_status'] == 'No Activity'))
        stats['found_inactive'] += mask_inactive.sum()
        
        # Identify Inferred (WLC/Forti)
        mask_inferred = (processed_df['status'] == 'Inferred')
        stats['inferred'] += mask_inferred.sum()
        
        # Identify Not Found & Errors
        mask_not_found = (processed_df['status'] == 'Not Found')
        stats['not_found'] += mask_not_found.sum()
        
        mask_error = processed_df['status'].str.startswith('API Error', na=False)
        stats['errors'] += mask_error.sum()

        # --- Filter for Excel Report (ACTIONABLE ISSUES ONLY) ---
        # Exclude 'is_inferred' from this mask so they don't appear in the attachment
        mask_report = mask_inactive | mask_not_found | mask_error
        
        if mask_report.any():
            sub = processed_df[mask_report].copy()
            # Update remark for inactive items specifically to be clear
            sub.loc[mask_inactive, 'remarks'] = f'Inactive - No events in last {ACTIVITY_THRESHOLD_DAYS} days'
            sub['sheet_name'] = name
            frames.append(sub)

    # If no "Bad" items found, skip email
    if not frames:
        print("✅ No Actionable Issues (Inactive/Not Found) detected; skipping email.")
        return

    # Check permissions before saving report
    try:
        result_df = pd.concat(frames, ignore_index=True)
        result_df.to_excel(draft_path, index=False)
        print(f"\n💾 Filtered report saved to: {draft_path}")

        # Total Issues = Rows in the attachment (Inactive + Not Found + Errors)
        total_issues = len(result_df)

        # Email Body
        subject = f"QRadar Action Report - {total_issues} Issues Require Attention"
        
        body = f"""Hello,

Attached is the QRadar log source status report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.

⚠️ ACTION REQUIRED: {total_issues} ISSUES
(Count includes Inactive sources and Missing assets only)

📊 SCAN SUMMARY
----------------------------------------
Total Assets Scanned:  {stats['total_scanned']} (Rows where 'In Qradar' = Yes)

🔴 ISSUES FOUND (Attached):
   - Inactive:          {stats['found_inactive']} (No events in {ACTIVITY_THRESHOLD_DAYS} days)
   - Not Found:         {stats['not_found']} (Absent in QRadar)
   - API Errors:        {stats['errors']}

🟢 HEALTHY / IGNORED (Not Attached):
   - Active:            {stats['found_active']}
   - Under WLC/Forti:   {stats['inferred']} (Found via inference logic)

Please review the attached Excel file for the list of items requiring attention.

Best regards,
QRadar Automation System
"""
        create_outlook_draft(draft_path, subject, body)
        
    except PermissionError:
            print(f"❌ ERROR: Could not save report to '{draft_path}'. Is the file open?")
