/************************************************************
 * Nova KPI Dashboard - Sheet Builder
 * Creates dynamic dashboard sheets with auto-refresh capabilities
 ************************************************************/

/**
 * Creates all dashboard sheets with proper structure and formulas
 * Run once to set up your dashboard
 */
function createDashboardSheets() {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // 1. Live_Snapshot Tab - Real-time active sessions & queue
    createLiveSnapshotSheet_(ss);
    
    // 2. Management_View Tab - KPI chips and summary tables
    createManagementViewSheet_(ss);
    
    // 3. Rep_QuickView Tab - Today's counts only
    createRepQuickViewSheet_(ss);
    
    // 4. Dashboard_Config Tab - Settings and refresh controls
    createDashboardConfigSheet_(ss);
    
    SpreadsheetApp.getActive().toast('Dashboard sheets created! Refresh data using "Rescue → Refresh Now (Live)"');
    
    // Auto-populate with current data
    refreshDashboardData();
    
  } catch (e) {
    SpreadsheetApp.getActive().toast('Error creating sheets: ' + e.toString().substring(0, 50));
    Logger.log('createDashboardSheets error: ' + e.toString());
  }
}

/**
 * Creates Live_Snapshot sheet - shows active sessions and waiting queue
 */
function createLiveSnapshotSheet_(ss) {
  let sh = ss.getSheetByName('Live_Snapshot');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Live_Snapshot');
  
  // Header
  sh.getRange(1, 1).setValue('Live Session Snapshot');
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1).setValue('Last Updated:');
  sh.getRange(2, 2).setFormula('=NOW()');
  sh.getRange(2, 2).setNumberFormat('mm/dd/yyyy hh:mm:ss AM/PM');
  
  // Refresh button placeholder
  sh.getRange(2, 4).setValue('→ Run "Rescue → Refresh Now (Live)" to update');
  
  // Active Sessions section
  sh.getRange(4, 1).setValue('Active Sessions');
  sh.getRange(4, 1).setFontSize(14).setFontWeight('bold');
  sh.getRange(5, 1, 1, 5).setValues([['Technician', 'Customer', 'Start (local)', 'Live Duration', 'Session ID']]);
  sh.getRange(5, 1, 1, 5).setFontWeight('bold').setBackground('#E5E7EB');
  
  // Data will be populated by refreshDashboardData()
  const activeStartRow = 6;
  sh.getRange(activeStartRow, 1).setValue('(Run refresh to see active sessions)');
  
  // Waiting Queue section
  const queueStartRow = activeStartRow + 15;
  sh.getRange(queueStartRow, 1).setValue('Waiting Queue');
  sh.getRange(queueStartRow, 1).setFontSize(14).setFontWeight('bold');
  sh.getRange(queueStartRow + 1, 1, 1, 4).setValues([['Channel', 'Customer', 'Waiting Since', 'Wait Duration']]);
  sh.getRange(queueStartRow + 1, 1, 1, 4).setFontWeight('bold').setBackground('#E5E7EB');
  
  // KPI Summary Boxes (top right)
  sh.getRange(4, 7).setValue('Rescue: Active Now');
  sh.getRange(4, 8).setFormula('=COUNTA(A6:A20)-1');
  sh.getRange(4, 8).setFontSize(18).setFontWeight('bold');
  
  sh.getRange(5, 7).setValue('Rescue: Waiting Now');
  sh.getRange(5, 8).setFormula(`=COUNTA(A${queueStartRow+2}:A${queueStartRow+20})-1`);
  sh.getRange(5, 8).setFontSize(18).setFontWeight('bold');
  
  // Formatting
  sh.setColumnWidth(1, 150); // Technician/Channel
  sh.setColumnWidth(2, 150); // Customer
  sh.setColumnWidth(3, 120); // Time
  sh.setColumnWidth(4, 100); // Duration
  sh.setColumnWidth(5, 120); // Session ID
  sh.setColumnWidth(7, 150); // KPI Labels
  sh.setColumnWidth(8, 80);  // KPI Values
  
  sh.getRange(1, 1, 1, 8).merge();
}

/**
 * Creates Management_View sheet - KPI summary and tables
 */
function createManagementViewSheet_(ss) {
  let sh = ss.getSheetByName('Management_View');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Management_View');
  
  // Header
  sh.getRange(1, 1).setValue('Management Dashboard');
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1).setValue('Last Updated:');
  sh.getRange(2, 2).setFormula('=NOW()');
  sh.getRange(2, 2).setNumberFormat('mm/dd/yyyy hh:mm:ss AM/PM');
  
  // KPI Chips Row
  const kpiRow = 4;
    const kpis = [
    ['Sessions (Today)', '=COUNTIFS(Performance!A:A, ">="&TODAY(), Performance!A:A, "<"&TODAY()+1)'],
    ['AHT', '=AVERAGE(Performance!G:G)'],
    ['SLA Hit %', '=COUNTIFS(Performance!H:H, "<=60")/COUNT(Performance!H:H)'],
    ['CSAT', '4.7/5'], // Placeholder - would come from separate source
    ['Utilization', '=SUM(Performance!F:F)/(8*3600*COUNTUNIQUE(Performance!C:C))'], // Work time / scheduled time
    ['SLA Target', '≤ 60s']
  ];
  
  sh.getRange(kpiRow, 1, kpis.length, 2).setValues(kpis.map(k => [k[0], k[1]]));
  
  // Format KPI boxes
  for (let i = 0; i < kpis.length; i++) {
    sh.getRange(kpiRow + i, 1, 1, 2).setBorder(true, true, true, true, true, true);
    sh.getRange(kpiRow + i, 1).setFontWeight('bold').setFontSize(10);
    sh.getRange(kpiRow + i, 2).setFontSize(14).setFontWeight('bold');
  }
  
  // Performance Table (will be populated by refreshDashboardData)
  const tableStartRow = kpiRow + kpis.length + 3;
  sh.getRange(tableStartRow, 1).setValue('Performance by Technician');
  sh.getRange(tableStartRow, 1).setFontSize(14).setFontWeight('bold');
  
  const headers = ['Technician', 'Sessions', 'AHT', 'SLA Hit %', 'Utilization', 'CSAT'];
  sh.getRange(tableStartRow + 1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(tableStartRow + 1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E5E7EB');
  
  // Formatting
  sh.setColumnWidths(1, headers.length, 120);
}

/**
 * Creates Rep_QuickView sheet - Today's quick stats
 */
function createRepQuickViewSheet_(ss) {
  let sh = ss.getSheetByName('Rep_QuickView');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Rep_QuickView');
  
  sh.getRange(1, 1).setValue('Today\'s Quick View');
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  sh.getRange(3, 1).setValue('Your Sessions Today:');
  sh.getRange(3, 2).setFormula('=COUNTIFS(Performance!C:C, "="&SUBSTITUTE(USEREMAIL(), "@nova.com", ""), Performance!A:A, ">="&TODAY())');
  
  sh.getRange(4, 1).setValue('Your AHT Today:');
  sh.getRange(4, 2).setFormula('=AVERAGEIFS(Performance!G:G, Performance!C:C, "="&SUBSTITUTE(USEREMAIL(), "@nova.com", ""), Performance!A:A, ">="&TODAY())');
  
  sh.getRange(5, 1).setValue('Your SLA Hit %:');
  sh.getRange(5, 2).setFormula('=COUNTIFS(Performance!C:C, "="&SUBSTITUTE(USEREMAIL(), "@nova.com", ""), Performance!H:H, "<=60", Performance!A:A, ">="&TODAY())/COUNTIFS(Performance!C:C, "="&SUBSTITUTE(USEREMAIL(), "@nova.com", ""), Performance!A:A, ">="&TODAY())');
  
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 120);
}

/**
 * Creates Dashboard_Config sheet - Settings
 */
function createDashboardConfigSheet_(ss) {
  let sh = ss.getSheetByName('Dashboard_Config');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Dashboard_Config');
  
  sh.getRange(1, 1).setValue('Dashboard Configuration');
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  const config = [
    ['Setting', 'Value', 'Description'],
    ['Auto-Refresh Interval', '15', 'Minutes between auto-refresh (0 = manual only)'],
    ['SLA Target (seconds)', '30', 'Target pickup time'],
    ['AHT Target (minutes)', '30', 'Target average handle time'],
    ['', '', ''],
    ['Refresh Status', '=IF(ISBLANK(Live_Snapshot!B2), "Never", "Last: "&TEXT(Live_Snapshot!B2, "hh:mm:ss AM/PM"))', 'Last refresh time']
  ];
  
  sh.getRange(3, 1, config.length, 3).setValues(config);
  sh.getRange(3, 1, 1, 3).setFontWeight('bold').setBackground('#E5E7EB');
  
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 300);
}

/**
 * Refreshes all dashboard data from BigQuery
 * Called automatically after data ingestion, or manually
 */
function refreshDashboardData() {
  try {
    const cfg = getCfg_();
    const ids = bqIds_(cfg);
    
    // Pull data from BigQuery for today and recent
    const today = new Date();
    const todayStr = today.toISOString().split('T')[0];
    const weekAgo = new Date(today);
    weekAgo.setDate(weekAgo.getDate() - 7);
    const weekAgoStr = weekAgo.toISOString().split('T')[0];
    
    // Get active sessions (status = 'Active' and today)
    const activeSQL = `
      SELECT 
        technician_name,
        COALESCE(customer_name, 'Anonymous') as customer,
        FORMAT_TIMESTAMP('%I:%M:%S %p', start_time, 'America/New_York') as start_local,
        FORMAT('%02d:%02d:%02d', 
          CAST(COALESCE(duration_active_seconds, 0) / 3600 AS INT64),
          CAST((COALESCE(duration_active_seconds, 0) % 3600) / 60 AS INT64),
          COALESCE(duration_active_seconds, 0) % 60
        ) as live_duration,
        session_id
      FROM \`${ids.project}.${ids.dataset}.rescue_sessions_latest\`
      WHERE session_status = 'Active'
        AND DATE(start_time) = CURRENT_DATE()
      ORDER BY start_time DESC
      LIMIT 20`;
    
    // Get waiting queue (status = 'Waiting' and today)
    const waitingSQL = `
      SELECT 
        channel_name,
        COALESCE(customer_name, '—') as customer,
        FORMAT_TIMESTAMP('%I:%M %p', start_time, 'America/New_York') as waiting_since,
        FORMAT('%02d:%02d', 
          CAST(COALESCE(pickup_seconds, 0) / 60 AS INT64),
          COALESCE(pickup_seconds, 0) % 60
        ) as wait_duration
      FROM \`${ids.project}.${ids.dataset}.rescue_sessions_latest\`
      WHERE session_status = 'Waiting'
        AND DATE(start_time) = CURRENT_DATE()
      ORDER BY start_time ASC
      LIMIT 20`;
    
    // Get performance data (last 7 days)
    const performanceSQL = `
      SELECT 
        DATE(start_time) as date,
        technician_name,
        technician_id,
        COUNT(DISTINCT session_id) as sessions,
        AVG(duration_total_seconds) as avg_aht_seconds,
        SUM(duration_work_seconds) as total_work_seconds,
    COUNTIF(pickup_seconds <= 60) / COUNT(*) as sla_hit_pct,
        AVG(pickup_seconds) as avg_pickup_seconds
      FROM \`${ids.project}.${ids.dataset}.rescue_sessions_latest\`
      WHERE DATE(start_time) >= @weekAgo
        AND technician_name IS NOT NULL
      GROUP BY date, technician_name, technician_id
      ORDER BY date DESC, sessions DESC`;
    
    const ss = SpreadsheetApp.getActive();
    
    // Populate Active Sessions
    try {
      const activeData = bqQuery_(cfg, activeSQL);
      const activeRows = bqResultsToArray_(activeData);
      populateSheetTable_(ss, 'Live_Snapshot', 6, activeRows, 5);
    } catch (e) {
      Logger.log('Error loading active sessions: ' + e.toString());
    }
    
    // Populate Waiting Queue
    try {
      const waitingData = bqQuery_(cfg, waitingSQL);
      const waitingRows = bqResultsToArray_(waitingData);
      populateSheetTable_(ss, 'Live_Snapshot', 21, waitingRows, 4);
    } catch (e) {
      Logger.log('Error loading waiting queue: ' + e.toString());
    }
    
    // Create/Update Performance sheet
    try {
      const perfData = bqQuery_(cfg, performanceSQL, { weekAgo: weekAgoStr });
      const perfRows = bqResultsToArray_(perfData);
      createPerformanceSheet_(ss, perfRows);
    } catch (e) {
      Logger.log('Error loading performance data: ' + e.toString());
    }
    
    // Update timestamp
    const liveSheet = ss.getSheetByName('Live_Snapshot');
    if (liveSheet) {
      liveSheet.getRange(2, 2).setValue(new Date());
    }
    
    SpreadsheetApp.getActive().toast('Dashboard refreshed');
    
  } catch (e) {
    Logger.log('refreshDashboardData error: ' + e.toString());
    SpreadsheetApp.getActive().toast('Refresh error: ' + e.toString().substring(0, 50));
  }
}

/**
 * Helper: Convert BigQuery results to 2D array
 */
function bqResultsToArray_(bqResult) {
  if (!bqResult.rows || !bqResult.rows.length) return [];
  
  return bqResult.rows.map(row => {
    return row.f.map(cell => cell.v || '');
  });
}

/**
 * Helper: Populate a table in a sheet
 */
function populateSheetTable_(ss, sheetName, startRow, data, numCols) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  
  // Clear old data
  sh.getRange(startRow, 1, 100, numCols).clearContent();
  
  if (!data || !data.length) {
    sh.getRange(startRow, 1).setValue('(No data available)');
    return;
  }
  
  // Write data
  sh.getRange(startRow, 1, data.length, numCols).setValues(data);
  
  // Formatting
  sh.getRange(startRow, 1, data.length, numCols).setBorder(true, true, true, true, true, true);
  sh.getRange(startRow, 1, data.length, numCols).setVerticalAlignment('middle');
}

/**
 * Creates/Updates Performance sheet with detailed data
 */
function createPerformanceSheet_(ss, data) {
  let sh = ss.getSheetByName('Performance');
  if (!sh) {
    sh = ss.insertSheet('Performance');
    sh.hideSheet(); // Hidden - used for formulas only
  }
  
  sh.clear();
  
  // Headers
  const headers = ['Date', 'Technician', 'Tech ID', 'Sessions', 'AHT (sec)', 'Work Time (sec)', 'AHT (formatted)', 'SLA (sec)', 'SLA Hit %', 'Utilization %'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  if (data && data.length > 0) {
    // Add formatted columns
    const formattedData = data.map(row => {
      const ahtSec = Number(row[4]) || 0;
      const ahtFormatted = `${Math.floor(ahtSec / 60)}m ${Math.floor(ahtSec % 60)}s`;
      const slaHitPct = Number(row[6]) || 0;
      const utilPct = 0; // Would need scheduled hours
      
      return [
        row[0], // Date
        row[1], // Technician
        row[2], // Tech ID
        row[3], // Sessions
        row[4], // AHT (sec)
        row[5], // Work Time (sec)
        ahtFormatted, // AHT formatted
        row[7], // SLA (sec)
        (slaHitPct * 100).toFixed(1) + '%', // SLA Hit %
        utilPct + '%' // Utilization
      ];
    });
    
    sh.getRange(2, 1, formattedData.length, headers.length).setValues(formattedData);
  }
}

