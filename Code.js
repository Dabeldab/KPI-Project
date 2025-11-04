/************************************************************
 * Nova Support KPIs – Rescue LISTALL → Google Sheets
 * Simplified Sheets-only version (no BigQuery)
 * - Pulls data from LogMeIn Rescue API
 * - Stores directly in Google Sheets
 * - Analytics dashboard with time frame filtering
 ************************************************************/

/* ===== Default nodes ===== */
const NODE_CANDIDATES_DEFAULT = [5648341, 300589800, 1367438801, 863388310, -2];

/* ===== Runtime settings ===== */
const FORCE_TEXT_OUTPUT = true;
const SHEETS_SESSIONS_TABLE = 'Sessions'; // Main data storage sheet

/* ===== Menu ===== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Rescue')
    .addItem('Configure Secrets', 'uiConfigureSecrets')
    .addSeparator()
    .addItem('🔍 API Smoke Test', 'apiSmokeTest')
    .addSeparator()
    .addItem('Pull Yesterday → Sheets', 'pullDateRangeYesterday')
    .addItem('Pull Today → Sheets', 'pullDateRangeToday')
    .addItem('Pull Last Week → Sheets', 'pullDateRangeLastWeek')
    .addItem('Pull This Week → Sheets', 'pullDateRangeThisWeek')
    .addItem('Pull Previous Month → Sheets', 'pullDateRangePreviousMonth')
    .addItem('Pull Current Month → Sheets', 'pullDateRangeCurrentMonth')
    .addItem('Pull Custom Range → Sheets', 'uiIngestRangeToSheets')
    .addSeparator()
    .addItem('🚀 Analytics Dashboard', 'createAnalyticsDashboard')
    .addItem('🔄 Refresh Dashboard (Pull from API)', 'refreshDashboardFromAPI')
    .addSeparator()
    .addItem('📈 Advanced Analytics Dashboard', 'createAdvancedAnalyticsDashboard')
    .addToUi();
}

/* ===== Secrets / Config ===== */
const PROP_KEYS = {
  RESCUE_BASE:  'RESCUE_BASE',
  RESCUE_USER:  'RESCUE_USER',
  RESCUE_PASS:  'RESCUE_PASS',
  NODE_JSON:    'NODE_CANDIDATES_JSON'
};

function getProps_() { return PropertiesService.getScriptProperties(); }
function setProp_(k, v) { getProps_().setProperty(k, v); }
function getProp_(k, d) { const v = getProps_().getProperty(k); return v != null ? v : d; }

function getCfg_() {
  const props = getProps_();
  return {
    rescueBase: getProp_(PROP_KEYS.RESCUE_BASE, 'https://secure.logmeinrescue.com/API'),
    user: getProp_(PROP_KEYS.RESCUE_USER, ''),
    pass: getProp_(PROP_KEYS.RESCUE_PASS, ''),
    nodes: (() => {
      const j = getProp_(PROP_KEYS.NODE_JSON, '[]');
      try { return JSON.parse(j); } catch(e) { return NODE_CANDIDATES_DEFAULT; }
    })()
  };
}

/* ===== UI Configuration ===== */
function uiConfigureSecrets() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; }
      input { width: 300px; margin: 5px 0; padding: 5px; }
      button { padding: 10px 20px; margin: 10px 5px 0 0; }
    </style>
    <h3>Configure Rescue API Secrets</h3>
    <p>Rescue Base URL:<br><input type="text" id="base" placeholder="https://secure.logmeinrescue.com/API"></p>
    <p>Email:<br><input type="email" id="user"></p>
    <p>Password:<br><input type="password" id="pass"></p>
    <p>Node IDs (JSON array):<br><input type="text" id="nodes" placeholder='[5648341, 300589800]'></p>
    <button onclick="save()">Save</button>
    <button onclick="google.script.host.close()">Cancel</button>
    <script>
      function save() {
        const base = document.getElementById('base').value;
        const user = document.getElementById('user').value;
        const pass = document.getElementById('pass').value;
        const nodes = document.getElementById('nodes').value;
        google.script.run.saveSecrets(base, user, pass, nodes);
        google.script.host.close();
      }
    </script>
  `).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Secrets');
}

function saveSecrets(base, user, pass, nodes) {
  const props = getProps_();
  if (base) setProp_(PROP_KEYS.RESCUE_BASE, base);
  if (user) setProp_(PROP_KEYS.RESCUE_USER, user);
  if (pass) setProp_(PROP_KEYS.RESCUE_PASS, pass);
  if (nodes) {
    try { JSON.parse(nodes); setProp_(PROP_KEYS.NODE_JSON, nodes); } catch(e) {}
  }
  SpreadsheetApp.getActive().toast('Secrets saved');
}

/* ===== HTTP Helpers ===== */
function apiGet_(base, endpoint, params, cookie, tries, mute) {
  const qs = Object.keys(params || {}).map(k => `${k}=${encodeURIComponent(params[k])}`).join('&');
  const url = `${base}/${endpoint}${qs ? '?' + qs : ''}`;
  const opts = {
    method: 'get',
    headers: Object.assign(
      {'User-Agent': 'Mozilla/5.0'},
      cookie ? {'Cookie': cookie} : {}
    ),
    muteHttpExceptions: !!mute
  };
  return UrlFetchApp.fetch(url, opts);
}

function extractCookie_(response) {
  const headers = response.getHeaders();
  const setCookie = headers['Set-Cookie'] || headers['set-cookie'];
  if (!setCookie) return null;
  const cookies = Array.isArray(setCookie) ? setCookie : [setCookie];
  return cookies.map(c => c.split(';')[0]).join('; ');
}

/* ===== API Functions (per LogMeIn documentation) ===== */
function login_(base, email, pwd) {
  const res = apiGet_(base, 'login.aspx', {email, pwd}, null, 3, true);
  const body = (res.getContentText()||'').trim();
  if (!/^OK/i.test(body)) {
    throw new Error(`Login failed: ${body.slice(0, 200)}`);
  }
  const cookie = extractCookie_(res);
  if (!cookie) throw new Error('Login: no cookie received');
  return cookie;
}

function setReportAreaSession_(base, cookie) {
  const r = apiGet_(base, 'setReportArea.aspx', {area: '0'}, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t)) throw new Error(`setReportArea: ${t}`);
}

function setReportTypeListAll_(base, cookie) {
  // Per API documentation: parameter name is 'type' (not 'reporttype')
  const r = apiGet_(base, 'setReportType.aspx', {type: 'LISTALL'}, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t) && !/NOTSUPPORTED/i.test(t)) {
    throw new Error(`setReportType: ${t}`);
  }
}

function setReportTypeSummary_(base, cookie) {
  // Per API documentation: parameter name is 'type' (not 'reporttype')
  const r = apiGet_(base, 'setReportType.aspx', {type: 'SUMMARY'}, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t) && !/NOTSUPPORTED/i.test(t)) {
    throw new Error(`setReportType SUMMARY: ${t}`);
  }
}

// Get current report type from API (per LogMeIn API documentation)
// Returns 'LISTALL' or 'SUMMARY' or null if error
// Response format: "OK REPORTTYPE:SUMMARY" or "OK REPORTTYPE:LISTALL"
function getReportType_(base, cookie) {
  try {
    const r = apiGet_(base, 'getReportType.aspx', {}, cookie, 2, true);
    const t = (r.getContentText()||'').trim();
    if (!/^OK/i.test(t)) {
      Logger.log(`getReportType error: ${t}`);
      return null;
    }
    // Response format: "OK REPORTTYPE:SUMMARY" or "OK REPORTTYPE:LISTALL"
    const match = t.match(/REPORTTYPE:\s*(LISTALL|SUMMARY)/i);
    if (match) {
      return match[1].toUpperCase();
    }
    // Also try format without "REPORTTYPE:" prefix
    const match2 = t.match(/OK\s+(LISTALL|SUMMARY)/i);
    if (match2) {
      return match2[1].toUpperCase();
    }
    Logger.log(`getReportType unexpected format: ${t}`);
    return null;
  } catch (e) {
    Logger.log(`getReportType exception: ${e.toString()}`);
    return null;
  }
}

// Set report area to Performance (area code varies, but typically area 1 or 2 for performance)
// Per LogMeIn API: area 0 = Session, other areas may be for performance/summary data
function setReportAreaPerformance_(base, cookie) {
  // Try area 1 first (common for performance data), fallback to area 0 if needed
  const r = apiGet_(base, 'setReportArea.aspx', {area: '1'}, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t)) {
    // Fallback to area 0 if area 1 doesn't work
    Logger.log(`setReportArea Performance: area 1 failed, trying area 0`);
    const r2 = apiGet_(base, 'setReportArea.aspx', {area: '0'}, cookie, 2, true);
    const t2 = (r2.getContentText()||'').trim();
    if (!/^OK/i.test(t2)) {
      Logger.log(`setReportArea Performance warning: ${t2} (continuing anyway)`);
    }
  }
}

function setReportTimeAllDay_(base, cookie) {
  const r = apiGet_(base, 'setReportTime.aspx', {btime: '00:00:00', etime: '23:59:59'}, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t)) throw new Error(`setReportTime: ${t}`);
}

// Set timezone to Eastern Time (EDT/EST) with automatic DST detection
// Per LogMeIn API documentation
// EDT (Daylight Saving) = UTC-4 = -240 minutes
// EST (Standard Time) = UTC-5 = -300 minutes
function setTimezoneEDT_(base, cookie) {
  // Detect if DST is currently active for Eastern Time
  // DST in US: 2nd Sunday in March to 1st Sunday in November
  const now = new Date();
  const year = now.getFullYear();
  
  // Find 2nd Sunday in March (DST starts)
  const march2ndSunday = new Date(year, 2, 1); // March 1st
  const marchDayOfWeek = march2ndSunday.getDay(); // 0=Sunday, 1=Monday, etc.
  const daysToAdd = (7 - marchDayOfWeek) % 7 + 7; // Days to 2nd Sunday
  const dstStart = new Date(year, 2, 1 + daysToAdd);
  dstStart.setHours(2, 0, 0, 0); // 2 AM
  
  // Find 1st Sunday in November (DST ends)
  const nov1stSunday = new Date(year, 10, 1); // November 1st
  const novDayOfWeek = nov1stSunday.getDay();
  const daysToAddNov = (7 - novDayOfWeek) % 7; // Days to 1st Sunday
  const dstEnd = new Date(year, 10, 1 + daysToAddNov);
  dstEnd.setHours(2, 0, 0, 0); // 2 AM
  
  // Check if DST is active
  const isDST = now >= dstStart && now < dstEnd;
  const timezoneOffsetMinutes = isDST ? -240 : -300; // EDT (UTC-4) or EST (UTC-5)
  const timezoneName = isDST ? 'EDT (UTC-4)' : 'EST (UTC-5)';
  
  Logger.log(`Setting timezone to ${timezoneName} (offset: ${timezoneOffsetMinutes} minutes)`);
  
  const r = apiGet_(base, 'setTimezone.aspx', {timezone: String(timezoneOffsetMinutes)}, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t)) {
    Logger.log(`setTimezone warning: ${t} (continuing anyway)`);
    // Non-fatal - continue even if timezone setting fails
  } else {
    Logger.log(`Timezone set to ${timezoneName}`);
  }
}

// Standard setReportDate per LogMeIn API documentation
function setReportDate_(base, cookie, fromIso, toIso) {
  const bdate = mdy_(fromIso);
  const edate = mdy_(toIso);
  if (!bdate || !edate || bdate.includes('undefined') || edate.includes('undefined')) {
    throw new Error(`Invalid date format: ${fromIso} → ${toIso}`);
  }
  const r = apiGet_(base, 'setReportDate.aspx', { bdate, edate }, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t)) {
    if (/INVALIDFORMAT/i.test(t)) throw new Error(`setReportDate: Invalid date format`);
    if (/INVALIDDATERANGE/i.test(t)) throw new Error(`setReportDate: End date is earlier than start date`);
    if (/NOTLOGGEDIN/i.test(t)) throw new Error(`setReportDate: Not logged in`);
    throw new Error(`setReportDate: ${t}`);
  }
}

function setDelimiter_(base, cookie, delimiter) {
  const r = apiGet_(base, 'setDelimiter.aspx', { delimiter: String(delimiter) }, cookie, 2, true);
  const t = (r.getContentText()||'').trim();
  if (!/^OK/i.test(t)) throw new Error(`setDelimiter: ${t}`);
}

function setOutputXMLOrFallback_(base, cookie) {
  if (FORCE_TEXT_OUTPUT) {
    const rt = apiGet_(base, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
    const tt = (rt.getContentText() || '').trim();
    if (!/^OK/i.test(tt)) throw new Error(`setOutput TEXT failed: ${tt}`);
    return 'TEXT';
  }
  try {
    const rx = apiGet_(base, 'setOutput.aspx', {output: 'XML'}, cookie, 2, true);
    const tx = (rx.getContentText()||'').trim();
    if (!/^OK/i.test(tx)) throw new Error(tx);
    return 'XML';
  } catch (e) {
    const rt = apiGet_(base, 'setOutput.aspx', {output: 'TEXT'}, cookie, 2, true);
    const tt = (rt.getContentText()||'').trim();
    if (!/^OK/i.test(tt)) throw new Error(`setOutput fallback TEXT failed: ${tt}`);
    return 'TEXT';
  }
}

function getReportTry_(base, cookie, nodeId, noderef) {
  try {
    const r = apiGet_(base, 'getReport.aspx', { node: String(nodeId), noderef: noderef }, cookie, 4, true);
    const t = (r.getContentText()||'').trim();
    if (/^OK/i.test(t)) return t;
    return null;
  } catch (e) {
    Logger.log(`getReportTry_ failed for node ${nodeId} (${noderef}): ${e.toString()}`);
    return null;
  }
}

/* ===== Parsing ===== */
function parsePipe_(okBody, delimiter) {
  if (!okBody || typeof okBody !== 'string') return { headers: [], rows: [] };
  const body = String(okBody).replace(/^OK\s*/i, '').trim();
  if (!body) return { headers: [], rows: [] };
  const lines = body.split(/\r?\n/).filter(Boolean);
  if (!lines.length) return { headers: [], rows: [] };
  const delim = String(delimiter || '|');
  const header = lines[0].split(delim).map(h => h.trim());
  if (!header.length) return { headers: [], rows: [] };
  const out = [];
  for (let i=1; i<lines.length; i++) {
    const cols = lines[i].split(delim);
    while (cols.length < header.length) cols.push('');
    const obj = {};
    for (let j=0; j<header.length; j++) {
      obj[header[j]] = (cols[j] || '').trim();
    }
    out.push(obj);
  }
  return { headers: header, rows: out };
}

/* ===== Mapping to Schema ===== */
function mapRow_(rec) {
  if (!rec || typeof rec !== 'object') return null;
  const g = (o, keys, d='') => { 
    for (const k of keys) {
      if (o[k] != null && String(o[k]).length) return String(o[k]).trim();
    }
    return d; 
  };
  const toSec = (val) => {
    const s = String(val||'').trim();
    if (!s) return 0;
    if (/^\d+$/.test(s)) return Number(s);
    const m = s.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
    if (m) return (Number(m[1])*3600 + Number(m[2])*60 + Number(m[3]))|0;
    return 0;
  };
  const toTs = (s) => {
    const v = String(s||'').trim();
    if (!v) return null;
    try {
      const d = new Date(v);
      if (isNaN(d.getTime())) return null;
      return d.toISOString();
    } catch (e) {
      return null;
    }
  };
  return {
    session_id: g(rec, ['Session ID'], ''),
    session_type: g(rec, ['Session Type'], ''),
    session_status: g(rec, ['Status'], ''),
    technician_id: g(rec, ['Technician ID'], ''),
    technician_name: g(rec, ['Technician Name'], ''),
    technician_email: g(rec, ['Technician Email'], ''),
    technician_group: g(rec, ['Technician Group'], ''),
    customer_name: g(rec, ['Your Name:'], ''),
    customer_email: g(rec, ['Your Email:', 'Customer Email'], ''),
    tracking_id: g(rec, ['Tracking ID'], ''),
    ip_address: g(rec, ['Customer IP'], ''),
    device_id: g(rec, ['Device ID'], ''),
    platform: g(rec, ['Platform'], ''),
    browser: g(rec, ['Browser Type'], ''),
    host: g(rec, ['Host Name'], ''),
    start_time: toTs(g(rec, ['Start Time'], '')),
    end_time: toTs(g(rec, ['End Time'], '')),
    last_action_time: toTs(g(rec, ['Last Action Time'], '')),
    duration_active_seconds: toSec(g(rec, ['Active Time'], '')),
    duration_work_seconds: toSec(g(rec, ['Work Time'], '')),
    duration_total_seconds: toSec(g(rec, ['Total Time'], '')),
    pickup_seconds: toSec(g(rec, ['Waiting Time'], '')),
    channel_id: g(rec, ['Channel ID'], ''),
    channel_name: g(rec, ['Channel Name'], ''),
    company_name: g(rec, ['Company name:'], ''),
    caller_name: g(rec, ['Your Name:'], ''),
    caller_phone: g(rec, ['Your Phone #:'], ''),
    resolved_unresolved: g(rec, ['Resolved Unresolved'], ''),
    calling_card: g(rec, ['Calling Card'], ''),
    browser_type: g(rec, ['Browser Type'], ''),
    connecting_time: toSec(g(rec, ['Connecting Time'], '')),
    waiting_time: toSec(g(rec, ['Waiting Time'], '')),
    total_time: toSec(g(rec, ['Total Time'], '')),
    active_time: toSec(g(rec, ['Active Time'], '')),
    work_time: toSec(g(rec, ['Work Time'], '')),
    hold_time: toSec(g(rec, ['Hold Time'], '')),
    time_in_transfer: toSec(g(rec, ['Time in Transfer'], '')),
    reconnecting_time: toSec(g(rec, ['Reconnecting Time'], '')),
    rebooting_time: toSec(g(rec, ['Rebooting Time'], '')),
    ingested_at: new Date().toISOString()
  };
}

/* ===== Date Helpers ===== */
function isoDate_(d) {
  const dt = d instanceof Date ? d : new Date(d);
  return dt.toISOString().split('T')[0];
}

function mdy_(iso) {
  const d = new Date(iso + 'T00:00:00');
  return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`;
}

/* ===== Sheets Storage ===== */
function getOrCreateSessionsSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
  if (!sh) {
    sh = ss.insertSheet(SHEETS_SESSIONS_TABLE);
  }
  
  // Always ensure headers exist and are correct (fixes issue where first column shows date)
  // Column remapping: session_id→technician_id, session_type→session_id, session_status→session_type,
  // technician_email→platform, technician_group→technician_email, customer_name→session_status,
  // device_id→ip_address, browser→technician_group, company_name→customer_phone_number,
  // caller_phone→customer_name, technician_id→applicationid
  const headers = [
    'technician_id', 'session_id', 'session_type', 'applicationid', 'technician_name',
    'platform', 'technician_email', 'session_status', 'customer_email',
    'tracking_id', 'ip_address', 'ip_address', 'platform', 'technician_group', 'host',
    'start_time', 'end_time', 'last_action_time',
    'duration_active_seconds', 'duration_work_seconds', 'duration_total_seconds',
    'pickup_seconds', 'channel_id', 'channel_name', 'customer_phone_number', 'session_status',
    'customer_name', 'resolved_unresolved', 'calling_card', 'browser_type',
    'connecting_time', 'waiting_time', 'total_time', 'active_time',
    'work_time', 'hold_time', 'time_in_transfer', 'reconnecting_time',
    'rebooting_time', 'ingested_at'
  ];
  
  // Check if header row exists and is correct
  const headerRange = sh.getRange(1, 1, 1, headers.length);
  const existingHeaders = headerRange.getValues()[0];
  const needsHeaderUpdate = !existingHeaders || existingHeaders.length !== headers.length || 
                           !existingHeaders[0] || existingHeaders[0] !== 'technician_id';
  
  if (needsHeaderUpdate) {
    // Clear and set headers
    headerRange.clear();
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    sh.setFrozenRows(1);
    Logger.log('Updated Sessions sheet headers');
  }
  
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(16, 180);
  sh.setColumnWidth(17, 180);
  sh.setColumnWidth(18, 180);
  
  return sh;
}

function writeRowsToSheets_(ss, rows, clearExisting = false) {
  if (!rows || !rows.length) return 0;
  const sh = getOrCreateSessionsSheet_(ss);
  
  // Clear existing data if requested (for range-specific pulls)
  // IMPORTANT: Keep header row (row 1) intact
  if (clearExisting) {
    const dataRange = sh.getDataRange();
    if (dataRange.getNumRows() > 1) {
      // Delete all rows except header (row 1)
      sh.deleteRows(2, dataRange.getNumRows() - 1);
      Logger.log('Cleared existing Sessions data (kept headers)');
    }
  }
  
  const existingIds = new Set();
  const dataRange = sh.getDataRange();
  if (dataRange.getNumRows() > 1) {
    // Column 1 is now technician_id (was session_id), but we still dedupe by session_id (now in column 2)
    const existing = sh.getRange(2, 2, dataRange.getNumRows() - 1, 1).getValues();
    existing.forEach(r => { if (r[0]) existingIds.add(String(r[0])); });
  }
  // Filter by session_id (now in column 2, but still use r.session_id from mapRow_)
  const toInsert = rows.filter(r => r && r.session_id && !existingIds.has(String(r.session_id)));
  if (!toInsert.length) return 0;
  
  // Convert ISO timestamp strings to Date objects for proper timezone handling
  const toDate = (isoStr) => {
    if (!isoStr) return null;
    try {
      return new Date(isoStr);
    } catch (e) {
      return null;
    }
  };
  
  // Map data to new column positions according to remapping:
  // Col 1: technician_id (was session_id position, contains technician_id data)
  // Col 2: session_id (was session_type position, contains session_id data)
  // Col 3: session_type (was session_status position, contains session_type data)
  // Col 4: applicationid (was technician_id position, contains technician_id data for now)
  // Col 5: technician_name, Col 6: platform (was technician_email, contains platform data)
  // Col 7: technician_email (was technician_group, contains technician_email data)
  // Col 8: session_status (was customer_name, contains session_status data)
  // Col 9: customer_email, Col 10: tracking_id, Col 11: ip_address,
  // Col 12: ip_address (was device_id, contains ip_address data again)
  // Col 13: platform (duplicate), Col 14: technician_group (was browser, contains technician_group data)
  // Col 15: host, Col 16-18: timestamps, Col 19-21: durations,
  // Col 22: pickup_seconds, Col 23-24: channel info,
  // Col 25: customer_phone_number (was company_name, contains caller_phone data)
  // Col 26: session_status (was caller_name, contains session_status data again)
  // Col 27: customer_name (was caller_phone, contains customer_name/caller_name data)
  // Col 28+: rest
  const values = toInsert.map(r => [
    r.technician_id, r.session_id, r.session_type, r.technician_id, r.technician_name,
    r.platform, r.technician_email, r.session_status, r.customer_email,
    r.tracking_id, r.ip_address, r.ip_address, r.platform, r.technician_group, r.host,
    toDate(r.start_time), toDate(r.end_time), toDate(r.last_action_time),
    r.duration_active_seconds, r.duration_work_seconds, r.duration_total_seconds,
    r.pickup_seconds, r.channel_id, r.channel_name, r.caller_phone, r.session_status,
    r.customer_name || r.caller_name, r.resolved_unresolved, r.calling_card, r.browser_type,
    r.connecting_time, r.waiting_time, r.total_time, r.active_time,
    r.work_time, r.hold_time, r.time_in_transfer, r.reconnecting_time,
    r.rebooting_time, toDate(r.ingested_at)
  ]);
  
  const newRowStart = sh.getLastRow() + 1;
  sh.getRange(newRowStart, 1, values.length, values[0].length).setValues(values);
  
  // Set date format for timestamp columns (columns 16, 17, 18, and 38)
  // start_time (col 16), end_time (col 17), last_action_time (col 18), ingested_at (col 38)
  const dateFormat = 'mm/dd/yyyy hh:mm:ss AM/PM';
  if (values.length > 0) {
    sh.getRange(newRowStart, 16, values.length, 1).setNumberFormat(dateFormat); // start_time
    sh.getRange(newRowStart, 17, values.length, 1).setNumberFormat(dateFormat); // end_time
    sh.getRange(newRowStart, 18, values.length, 1).setNumberFormat(dateFormat); // last_action_time
    sh.getRange(newRowStart, 38, values.length, 1).setNumberFormat(dateFormat); // ingested_at
  }
  
  return toInsert.length;
}

/* ===== Ingestion Functions ===== */
function pullDateRangeYesterday() {
  try {
    const range = getTimeFrameRange_('Yesterday');
    const cfg = getCfg_();
    
    // Ensure we're using local timezone dates (not UTC)
    // startDate should be yesterday 00:00:00 local, endDate should be yesterday 23:59:59.999 local
    const startISO = `${range.startDate.getFullYear()}-${String(range.startDate.getMonth() + 1).padStart(2, '0')}-${String(range.startDate.getDate()).padStart(2, '0')}T00:00:00.000Z`;
    const endISO = `${range.endDate.getFullYear()}-${String(range.endDate.getMonth() + 1).padStart(2, '0')}-${String(range.endDate.getDate()).padStart(2, '0')}T23:59:59.999Z`;
    
    Logger.log(`Pulling yesterday: ${startISO} to ${endISO} (local date: ${range.startDate.toLocaleDateString()})`);
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    SpreadsheetApp.getActive().toast(`✅ Ingested ${rowsIngested} rows from yesterday`, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
    Logger.log('pullDateRangeYesterday error: ' + e.toString());
  }
}

function pullDateRangeToday() {
  try {
    const range = getTimeFrameRange_('Today');
    const cfg = getCfg_();
    
    // Ensure we're using local timezone dates (not UTC)
    // startDate should be today 00:00:00 local, endDate should be today 23:59:59.999 local
    const startISO = `${range.startDate.getFullYear()}-${String(range.startDate.getMonth() + 1).padStart(2, '0')}-${String(range.startDate.getDate()).padStart(2, '0')}T00:00:00.000Z`;
    const endISO = `${range.endDate.getFullYear()}-${String(range.endDate.getMonth() + 1).padStart(2, '0')}-${String(range.endDate.getDate()).padStart(2, '0')}T23:59:59.999Z`;
    
    Logger.log(`Pulling today: ${startISO} to ${endISO} (local date: ${range.startDate.toLocaleDateString()})`);
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    SpreadsheetApp.getActive().toast(`✅ Ingested ${rowsIngested} rows from today`, 5);
    
    // Schedule auto-refresh for today (every 10 minutes)
    scheduleAutoRefreshToday_();
  } catch (e) {
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
    Logger.log('pullDateRangeToday error: ' + e.toString());
  }
}

function pullDateRangeLastWeek() {
  try {
    const range = getTimeFrameRange_('Previous Week');
    const cfg = getCfg_();
    // Ensure time is set to full day boundaries
    const startISO = `${range.startDate.getFullYear()}-${String(range.startDate.getMonth() + 1).padStart(2, '0')}-${String(range.startDate.getDate()).padStart(2, '0')}T00:00:00.000Z`;
    const endISO = `${range.endDate.getFullYear()}-${String(range.endDate.getMonth() + 1).padStart(2, '0')}-${String(range.endDate.getDate()).padStart(2, '0')}T23:59:59.999Z`;
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    SpreadsheetApp.getActive().toast(`✅ Ingested ${rowsIngested} rows from last week`, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
    Logger.log('pullDateRangeLastWeek error: ' + e.toString());
  }
}

function pullDateRangeThisWeek() {
  try {
    const range = getTimeFrameRange_('This Week');
    const cfg = getCfg_();
    const startISO = `${range.startDate.getFullYear()}-${String(range.startDate.getMonth() + 1).padStart(2, '0')}-${String(range.startDate.getDate()).padStart(2, '0')}T00:00:00.000Z`;
    const endISO = `${range.endDate.getFullYear()}-${String(range.endDate.getMonth() + 1).padStart(2, '0')}-${String(range.endDate.getDate()).padStart(2, '0')}T23:59:59.999Z`;
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    SpreadsheetApp.getActive().toast(`✅ Ingested ${rowsIngested} rows from this week`, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
    Logger.log('pullDateRangeThisWeek error: ' + e.toString());
  }
}

function pullDateRangePreviousMonth() {
  try {
    const range = getTimeFrameRange_('Last Month');
    const cfg = getCfg_();
    const startISO = `${range.startDate.getFullYear()}-${String(range.startDate.getMonth() + 1).padStart(2, '0')}-${String(range.startDate.getDate()).padStart(2, '0')}T00:00:00.000Z`;
    const endISO = `${range.endDate.getFullYear()}-${String(range.endDate.getMonth() + 1).padStart(2, '0')}-${String(range.endDate.getDate()).padStart(2, '0')}T23:59:59.999Z`;
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    SpreadsheetApp.getActive().toast(`✅ Ingested ${rowsIngested} rows from previous month`, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
    Logger.log('pullDateRangePreviousMonth error: ' + e.toString());
  }
}

function pullDateRangeCurrentMonth() {
  try {
    const range = getTimeFrameRange_('This Month');
    const cfg = getCfg_();
    const startISO = `${range.startDate.getFullYear()}-${String(range.startDate.getMonth() + 1).padStart(2, '0')}-${String(range.startDate.getDate()).padStart(2, '0')}T00:00:00.000Z`;
    const endISO = `${range.endDate.getFullYear()}-${String(range.endDate.getMonth() + 1).padStart(2, '0')}-${String(range.endDate.getDate()).padStart(2, '0')}T23:59:59.999Z`;
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    SpreadsheetApp.getActive().toast(`✅ Ingested ${rowsIngested} rows from current month`, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
    Logger.log('pullDateRangeCurrentMonth error: ' + e.toString());
  }
}

function scheduleAutoRefreshToday_() {
  // Remove existing auto-refresh triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoRefreshToday_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Schedule refresh every 10 minutes
  ScriptApp.newTrigger('autoRefreshToday_')
    .timeBased()
    .everyMinutes(10)
    .create();
}

function autoRefreshToday_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const configSheet = ss.getSheetByName('Dashboard_Config');
    if (!configSheet) return;
    
    const timeFrame = configSheet.getRange('B3').getValue();
    if (timeFrame !== 'Today') {
      // Stop auto-refresh if not set to Today
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === 'autoRefreshToday_') {
          ScriptApp.deleteTrigger(t);
        }
      });
      return;
    }
    
    const range = getTimeFrameRange_('Today');
    const cfg = getCfg_();
    const startISO = range.startDate.toISOString();
    const endISO = range.endDate.toISOString();
    ingestTimeRangeToSheets_(startISO, endISO, cfg, false); // Don't clear, just append new
  } catch (e) {
    Logger.log('autoRefreshToday_ error: ' + e.toString());
  }
}

function uiIngestRangeToSheets() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial; padding: 20px; }
      input { width: 200px; margin: 5px 0; padding: 5px; }
      button { padding: 10px 20px; margin: 10px 5px 0 0; }
    </style>
    <h3>Pull Custom Date Range</h3>
    <p>Start Date:<br><input type="date" id="startDate"></p>
    <p>End Date:<br><input type="date" id="endDate"></p>
    <button onclick="pull()">Pull Data</button>
    <button onclick="google.script.host.close()">Cancel</button>
    <script>
      function pull() {
        const start = document.getElementById('startDate').value;
        const end = document.getElementById('endDate').value;
        if (!start || !end) {
          alert('Please select both start and end dates');
          return;
        }
        google.script.run.withSuccessHandler(function(msg) {
          alert(msg);
          google.script.host.close();
        }).ingestCustomRangeToSheets(start, end);
      }
    </script>
  `).setWidth(350).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pull Custom Range');
}

function ingestCustomRangeToSheets(startDateStr, endDateStr) {
  try {
    // Parse dates and ensure they're treated as local dates (not UTC)
    // startDateStr and endDateStr are in format YYYY-MM-DD
    const startISO = `${startDateStr}T00:00:00.000Z`;
    const endISO = `${endDateStr}T23:59:59.999Z`;
    const cfg = getCfg_();
    Logger.log(`Pulling custom range: ${startISO} to ${endISO} (dates: ${startDateStr} to ${endDateStr})`);
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    return `✅ Ingested ${rowsIngested} rows from ${startDateStr} to ${endDateStr}`;
  } catch (e) {
    Logger.log('ingestCustomRangeToSheets error: ' + e.toString());
    return 'Error: ' + e.toString().substring(0, 100);
  }
}

function ingestTimeRangeToSheets_(startTimestamp, endTimestamp, cfg, clearExisting = true) {
  const ss = SpreadsheetApp.getActive();
  
  // Show loading indicator on dashboard
  showLoadingIndicator_(ss, true);
  
  try {
    const nodes = cfg.nodes.map(n => Number(n)).filter(Number.isFinite);
    if (!nodes.length) throw new Error('No valid nodes configured');
    const noderefs = ['NODE','CHANNEL'];
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    
    // Set timezone to EDT (Eastern Daylight Time) before setting dates/times
    setTimezoneEDT_(cfg.rescueBase, cookie);
    
    setReportAreaSession_(cfg.rescueBase, cookie);
    setReportTypeListAll_(cfg.rescueBase, cookie);
    setOutputXMLOrFallback_(cfg.rescueBase, cookie);
    setDelimiter_(cfg.rescueBase, cookie, '|');
    // Extract date strings (YYYY-MM-DD) for API date setting
    const startDateIso = startTimestamp.split('T')[0];
    const endDateIso = endTimestamp.split('T')[0];
    
    Logger.log(`Setting API date range: ${startDateIso} to ${endDateIso}`);
    setReportDate_(cfg.rescueBase, cookie, startDateIso, endDateIso);
    
    // For same-day pulls, set time range to 00:00:00 to 23:59:59
    if (startDateIso === endDateIso) {
      // Always use full day (00:00:00 to 23:59:59) for single-day pulls
      // This ensures we get all data from that specific day only
      Logger.log(`Setting time range for single day: 00:00:00 to 23:59:59`);
      setReportTimeAllDay_(cfg.rescueBase, cookie);
    } else {
      // For multi-day ranges, use full day for each day
      Logger.log(`Setting time range for multi-day: 00:00:00 to 23:59:59`);
      setReportTimeAllDay_(cfg.rescueBase, cookie);
    }
    let allMappedRows = [];
    for (const nr of noderefs) {
      for (const node of nodes) {
        try {
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          if (!t || !/^OK/i.test(t)) continue;
          const parseResult = parsePipe_(t, '|');
          const parsed = parseResult.rows || [];
          if (!parsed || !parsed.length) continue;
          // Extract date strings (YYYY-MM-DD) from timestamps for comparison
          // startTimestamp format: YYYY-MM-DDTHH:MM:SS.sssZ
          const startDateStr = startTimestamp.split('T')[0];
          const endDateStr = endTimestamp.split('T')[0];
          
          Logger.log(`Filtering sessions: date range=${startDateStr} to ${endDateStr} (strict: only dates within this range)`);
          
          const mapped = parsed.map(mapRow_).filter(r => {
            if (!r || !r.session_id) return false;
            
            // Use start_time for date filtering (LogMeIn returns sessions by start date)
            const rowStartTime = r.start_time ? new Date(r.start_time) : null;
            if (!rowStartTime || isNaN(rowStartTime.getTime())) {
              Logger.log(`Session ${r.session_id} has invalid start_time: ${r.start_time}`);
              return false;
            }
            
            // Convert row timestamp to date string (YYYY-MM-DD) for comparison
            // The API returns timestamps, we need to extract just the date part
            // Use UTC date to match the ISO date string format
            const rowDateStr = rowStartTime.toISOString().split('T')[0];
            
            // Strict date filtering: only include if session started on exact date within range
            // This ensures "Last Month" (October) only includes sessions from 10/01 to 10/31, not 11/01 or 11/02
            const isInRange = rowDateStr >= startDateStr && rowDateStr <= endDateStr;
            
            if (!isInRange && parsed.length < 100) {
              // Only log if we have a small number of rows (to avoid spam)
              Logger.log(`Filtering out session ${r.session_id}: rowDate=${rowDateStr} is outside range ${startDateStr} to ${endDateStr}`);
            }
            
            return isInRange;
          });
          
          Logger.log(`After filtering: ${mapped.length} sessions match date range ${startDateStr} to ${endDateStr} (from ${parsed.length} total parsed rows)`);
          if (!mapped.length) {
            Logger.log(`No sessions found in date range ${startDateStr} to ${endDateStr} for node ${node} (${nr})`);
            continue;
          }
          allMappedRows.push(...mapped);
          Utilities.sleep(200);
        } catch (e) {
          Logger.log(`Error processing node ${node} (${nr}): ${e.toString()}`);
        }
      }
    }
    if (allMappedRows.length > 0) {
      const written = writeRowsToSheets_(ss, allMappedRows, clearExisting);
      Logger.log(`Ingested ${written} new rows to Sheets`);
      
      // Get the date range used for this pull
      const pullStartDate = new Date(startTimestamp);
      const pullEndDate = new Date(endTimestamp);
      
      // Auto-refresh dashboard and create summaries with the pulled range
      refreshAnalyticsDashboard_(pullStartDate, pullEndDate);
      createDailySummarySheet_(ss, pullStartDate, pullEndDate);
      createSupportDataSheet_(ss, pullStartDate, pullEndDate);
      generateTechnicianTabs_(pullStartDate, pullEndDate);
      refreshAdvancedAnalyticsDashboard_(pullStartDate, pullEndDate);
      
      // Hide loading indicator
      showLoadingIndicator_(ss, false);
      
      return written;
    } else {
      showLoadingIndicator_(ss, false);
      return 0;
    }
  } catch (e) {
    showLoadingIndicator_(ss, false);
    throw e;
  }
}

function showLoadingIndicator_(ss, show) {
  const dashboardSheet = ss.getSheetByName('Analytics_Dashboard');
  if (!dashboardSheet) return;
  
  if (show) {
    dashboardSheet.getRange(2, 6).setValue('🔄 Loading...');
    dashboardSheet.getRange(2, 6).setFontColor('#EA8600').setFontWeight('bold');
    SpreadsheetApp.flush();
  } else {
    dashboardSheet.getRange(2, 6).clearContent();
  }
}

function fetchLiveActiveSessions_(cfg) {
  try {
    // Pull current active sessions from LogMeIn API
    const today = new Date();
    const todayStr = isoDate_(today);
    
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    setReportAreaSession_(cfg.rescueBase, cookie);
    setReportTypeListAll_(cfg.rescueBase, cookie);
    setOutputXMLOrFallback_(cfg.rescueBase, cookie);
    setDelimiter_(cfg.rescueBase, cookie, '|');
    setReportDate_(cfg.rescueBase, cookie, todayStr, todayStr);
    setReportTimeAllDay_(cfg.rescueBase, cookie);
    
    const nodes = cfg.nodes.map(n => Number(n)).filter(Number.isFinite);
    const noderefs = ['NODE','CHANNEL'];
    const allActiveRows = [];
    
    for (const nr of noderefs) {
      for (const node of nodes) {
        try {
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          if (!t || !/^OK/i.test(t)) continue;
          const parseResult = parsePipe_(t, '|');
          const parsed = parseResult.rows || [];
          if (!parsed || !parsed.length) continue;
          
          const activeSessions = parsed.map(mapRow_).filter(r => {
            if (!r || !r.session_id) return false;
            // Only include sessions with status "Active"
            return r.session_status === 'Active';
          });
          
          if (activeSessions.length > 0) {
            allActiveRows.push(...activeSessions);
          }
          Utilities.sleep(200);
        } catch (e) {
          Logger.log(`Error fetching active sessions from node ${node} (${nr}): ${e.toString()}`);
        }
      }
    }
    
    // Remove duplicates by session_id
    const uniqueActive = [];
    const seenIds = new Set();
    allActiveRows.forEach(row => {
      if (row.session_id && !seenIds.has(row.session_id)) {
        seenIds.add(row.session_id);
        uniqueActive.push(row);
      }
    });
    
    // Format for display
    return uniqueActive.slice(0, 25).map(row => {
      const startTime = row.start_time ? new Date(row.start_time) : new Date();
      const now = new Date();
      const liveDurationSec = Math.floor((now - startTime) / 1000);
      const hours = Math.floor(liveDurationSec / 3600);
      const minutes = Math.floor((liveDurationSec % 3600) / 60);
      const seconds = liveDurationSec % 60;
      const liveDuration = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
      const startTimeStr = row.start_time ? new Date(row.start_time).toLocaleString('en-US', { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true }) : '';
      return [
        row.technician_name || '', 
        row.customer_name || 'Anonymous', 
        startTimeStr, 
        liveDuration,
        row.channel_name || '—',
        row.session_id || ''
      ];
    });
  } catch (e) {
    Logger.log('fetchLiveActiveSessions_ error: ' + e.toString());
    return [];
  }
}

function fetchCurrentSessions_(cfg) {
  try {
    // Use getSession_v3 API to get current active sessions and waiting queue
    // Reference: https://support.logmein.com/rescue/help/rescue-api-reference-guide
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    
    // Get all channels - need to query by CHANNEL noderef for all channels
    // Per API documentation, we should query all channels, not just specific nodes
    const nodes = cfg.nodes.map(n => Number(n)).filter(Number.isFinite);
    const activeSessions = [];
    const waitingSessions = [];
    
    // Query all channels using CHANNEL noderef
    // Per API docs: getSession_v3 returns current sessions for a channel
    for (const node of nodes) {
      try {
        // Query by CHANNEL to get all sessions on that channel
        const r = apiGet_(cfg.rescueBase, 'getSession_v3.aspx', { 
          node: String(node), 
          noderef: 'CHANNEL' 
        }, cookie, 4, true);
        
        const responseText = (r.getContentText() || '').trim();
        if (!responseText || !/^OK/i.test(responseText)) continue;
        
        // Parse the response - getSession_v3 returns pipe-separated data
        // Format per documentation: SessionID|Status|...|TechnicianName|StartTime|Duration|Customer|...
        // Example: 1571322|Connecting|0||337366|John Doe|10/14/2011 10:12 AM|105|Customer1|||en|||yes||
        const lines = responseText.replace(/^OK\s*/i, '').split(/\r?\n/).filter(Boolean);
        if (lines.length < 2) continue; // Need at least header + 1 row
        
        // First line is headers, rest are data rows - both pipe-separated
        const headers = lines[0].split('|').map(h => h.trim());
        const sessionIdIdx = headers.findIndex(h => /session.*id/i.test(h));
        const statusIdx = headers.findIndex(h => /status/i.test(h));
        const techIdx = headers.findIndex(h => /technician.*name/i.test(h));
        const startTimeIdx = headers.findIndex(h => /start.*time/i.test(h));
        const durationIdx = headers.findIndex(h => /duration/i.test(h));
        const customerIdx = headers.findIndex(h => /customer|your.*name/i.test(h));
        const channelIdx = headers.findIndex(h => /channel/i.test(h));
        
        for (let i = 1; i < lines.length; i++) {
          const cols = lines[i].split('|');
          if (cols.length < headers.length) continue;
          
          const sessionId = (cols[sessionIdIdx] || '').trim();
          const status = (cols[statusIdx] || '').trim();
          const techName = (cols[techIdx] || '').trim();
          const startTime = (cols[startTimeIdx] || '').trim();
          const duration = (cols[durationIdx] || '').trim();
          const customer = (cols[customerIdx] || '').trim() || 'Anonymous';
          const channel = (cols[channelIdx] || '').trim() || '—';
          
          // Calculate wait duration
          let waitDuration = 0;
          if (startTime) {
            try {
              // Parse time in EDT format (MM/DD/YYYY HH:MM:SS AM/PM)
              const startDate = new Date(startTime);
              if (!isNaN(startDate.getTime())) {
                waitDuration = Math.floor((new Date() - startDate) / 1000);
              }
            } catch (e) {
              // Ignore date parsing errors
            }
          }
          
          const sessionData = {
            sessionId,
            status,
            technician: techName,
            startTime,
            duration,
            customer,
            channel,
            waitDuration
          };
          
          if (status === 'Waiting' || status === 'Connecting') {
            waitingSessions.push(sessionData);
          } else if (status === 'Active' || status === 'Connected' || status === 'In Session') {
            activeSessions.push(sessionData);
          }
        }
        
        Utilities.sleep(200);
      } catch (e) {
        Logger.log(`Error fetching sessions from channel node ${node}: ${e.toString()}`);
      }
    }
    
    // Remove duplicates by session ID
    const uniqueActive = [];
    const uniqueWaiting = [];
    const seenActive = new Set();
    const seenWaiting = new Set();
    
    activeSessions.forEach(s => {
      if (s.sessionId && !seenActive.has(s.sessionId)) {
        seenActive.add(s.sessionId);
        uniqueActive.push(s);
      }
    });
    
    waitingSessions.forEach(s => {
      if (s.sessionId && !seenWaiting.has(s.sessionId)) {
        seenWaiting.add(s.sessionId);
        uniqueWaiting.push(s);
      }
    });
    
    Logger.log(`fetchCurrentSessions_: Found ${uniqueActive.length} active, ${uniqueWaiting.length} waiting`);
    return {
      active: uniqueActive,
      waiting: uniqueWaiting
    };
  } catch (e) {
    Logger.log('fetchCurrentSessions_ error: ' + e.toString());
    return { active: [], waiting: [] };
  }
}

// Fetch performance summary data from API for accurate averages
// Uses SUMMARY report type and performance area to get aggregated metrics
function fetchPerformanceSummaryData_(cfg, startDate, endDate) {
  const performanceData = {};
  try {
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    
    // Set timezone to EDT/EST with DST detection
    setTimezoneEDT_(cfg.rescueBase, cookie);
    
    // Set up for performance/summary report
    setReportAreaPerformance_(cfg.rescueBase, cookie);
    setReportTypeSummary_(cfg.rescueBase, cookie);
    setOutputXMLOrFallback_(cfg.rescueBase, cookie);
    setDelimiter_(cfg.rescueBase, cookie, '|');
    
    // Set date range
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    setReportDate_(cfg.rescueBase, cookie, startStr, endStr);
    setReportTimeAllDay_(cfg.rescueBase, cookie);
    
    const nodes = cfg.nodes.map(n => Number(n)).filter(Number.isFinite);
    const noderefs = ['NODE', 'CHANNEL'];
    
    for (const nr of noderefs) {
      for (const node of nodes) {
        try {
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          if (!t || !/^OK/i.test(t)) continue;
          
          const parseResult = parsePipe_(t, '|');
          const parsed = parseResult.rows || [];
          if (!parsed || !parsed.length) continue;
          
          // Parse summary data - format varies, but typically includes technician name and averages
          // Log headers for debugging
          if (parseResult.headers && parseResult.headers.length > 0) {
            Logger.log(`Summary data headers: ${parseResult.headers.join('|')}`);
          }
          
          parsed.forEach(row => {
            // Try multiple variations of technician name column
            const techName = row['Technician Name'] || row['Technician'] || row['TechnicianName'] || 
                           row['Technician Name:'] || row['Tech Name'] || '';
            if (!techName) {
              // Log first row to see structure if no tech name found
              if (parsed.indexOf(row) === 0) {
                Logger.log(`First summary row (no tech name found): ${JSON.stringify(row).substring(0, 200)}`);
              }
              return;
            }
            
            // Extract average duration and pickup from summary data
            // Try multiple column name variations
            const avgDurCol = row['Avg Duration'] || row['Average Duration'] || row['Total Time'] || 
                            row['Avg Total Time'] || row['Average Total Time'] || row['Duration'] ||
                            row['Avg Duration:'] || row['Total Time:'] || '';
            const avgPickupCol = row['Avg Pickup'] || row['Average Pickup'] || row['Waiting Time'] || 
                              row['Avg Waiting Time'] || row['Average Waiting Time'] || row['Pickup'] ||
                              row['Avg Pickup:'] || row['Waiting Time:'] || '';
            
            if (!performanceData[techName]) {
              performanceData[techName] = {
                avgDuration: 0,
                avgPickup: 0,
                count: 0
              };
            }
            
            // Parse duration (could be in seconds or MM:SS format)
            let avgDur = 0;
            if (avgDurCol) {
              const durStr = String(avgDurCol).trim();
              if (/^\d+$/.test(durStr)) {
                avgDur = Number(durStr);
              } else {
                // Try parsing MM:SS or HH:MM:SS format
                const parts = durStr.split(':');
                if (parts.length === 2) {
                  avgDur = Number(parts[0]) * 60 + Number(parts[1]);
                } else if (parts.length === 3) {
                  avgDur = Number(parts[0]) * 3600 + Number(parts[1]) * 60 + Number(parts[2]);
                }
              }
            }
            
            // Parse pickup time (typically in seconds)
            let avgPickup = 0;
            if (avgPickupCol) {
              const pickupStr = String(avgPickupCol).trim();
              if (/^\d+$/.test(pickupStr)) {
                avgPickup = Number(pickupStr);
              } else {
                // Try parsing MM:SS format
                const parts = pickupStr.split(':');
                if (parts.length === 2) {
                  avgPickup = Number(parts[0]) * 60 + Number(parts[1]);
                }
              }
            }
            
            // Accumulate (in case multiple rows per technician)
            performanceData[techName].avgDuration += avgDur;
            performanceData[techName].avgPickup += avgPickup;
            performanceData[techName].count++;
          });
          
          Utilities.sleep(200);
        } catch (e) {
          Logger.log(`Error fetching performance data from node ${node} (${nr}): ${e.toString()}`);
        }
      }
    }
    
    // Calculate averages if we have multiple entries per technician
    Object.keys(performanceData).forEach(tech => {
      const perf = performanceData[tech];
      if (perf.count > 1) {
        perf.avgDuration = perf.avgDuration / perf.count;
        perf.avgPickup = perf.avgPickup / perf.count;
      }
    });
    
    Logger.log(`fetchPerformanceSummaryData_: Found performance data for ${Object.keys(performanceData).length} technicians`);
  } catch (e) {
    Logger.log('fetchPerformanceSummaryData_ error: ' + e.toString());
  }
  
  return performanceData;
}

// Fetch channel summary data from API for Support_Data sheet
// Uses SUMMARY report type with CHANNEL noderef to get all channel-level metrics
// Per API documentation: https://support.logmein.com/rescue/help/rescue-api-reference-guide
function fetchChannelSummaryData_(cfg, startDate, endDate) {
  const channelData = [];
  try {
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    
    // Set timezone to EDT/EST with DST detection
    setTimezoneEDT_(cfg.rescueBase, cookie);
    
    // Set up for summary report (can use Session area or Performance area)
    // Using Session area to get channel-level session summaries
    setReportAreaSession_(cfg.rescueBase, cookie);
    
    // Set output format and delimiter BEFORE setting report type
    setOutputXMLOrFallback_(cfg.rescueBase, cookie);
    setDelimiter_(cfg.rescueBase, cookie, '|');
    
    // Set date range BEFORE setting report type
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    setReportDate_(cfg.rescueBase, cookie, startStr, endStr);
    setReportTimeAllDay_(cfg.rescueBase, cookie);
    
    // IMPORTANT: Set report type to SUMMARY (must be set AFTER other settings)
    // Per API documentation: setReportType with type=SUMMARY
    setReportTypeSummary_(cfg.rescueBase, cookie);
    
    // Verify report type was set correctly using getReportType (non-blocking)
    try {
      const currentReportType = getReportType_(cfg.rescueBase, cookie);
      Logger.log(`Current report type after setting: ${currentReportType}`);
      if (!currentReportType || currentReportType !== 'SUMMARY') {
        Logger.log(`WARNING: Report type is ${currentReportType}, expected SUMMARY. Attempting to re-set...`);
        setReportTypeSummary_(cfg.rescueBase, cookie);
        const verifyType = getReportType_(cfg.rescueBase, cookie);
        Logger.log(`Report type after re-set: ${verifyType}`);
      } else {
        Logger.log(`✓ Report type confirmed as SUMMARY`);
      }
    } catch (e) {
      Logger.log(`getReportType verification failed (non-fatal, continuing): ${e.toString()}`);
      // Continue anyway - setReportType should have worked
    }
    
    const nodes = cfg.nodes.map(n => Number(n)).filter(Number.isFinite);
    
    // Query by CHANNEL noderef to get channel-level summaries
    for (const node of nodes) {
      try {
        // Verify and ensure report type is SUMMARY before each query (non-blocking)
        try {
          const typeBeforeQuery = getReportType_(cfg.rescueBase, cookie);
          if (!typeBeforeQuery || typeBeforeQuery !== 'SUMMARY') {
            Logger.log(`Report type before query is ${typeBeforeQuery}, re-setting to SUMMARY...`);
            setReportTypeSummary_(cfg.rescueBase, cookie);
            const verifyBeforeQuery = getReportType_(cfg.rescueBase, cookie);
            Logger.log(`Report type after re-set before query: ${verifyBeforeQuery}`);
          }
        } catch (e) {
          Logger.log(`getReportType before query failed (non-fatal, re-setting to be safe): ${e.toString()}`);
          // Re-set anyway to be safe
          setReportTypeSummary_(cfg.rescueBase, cookie);
        }
        
        const t = getReportTry_(cfg.rescueBase, cookie, node, 'CHANNEL');
        if (!t || !/^OK/i.test(t)) continue;
        
        const parseResult = parsePipe_(t, '|');
        const parsed = parseResult.rows || [];
        if (!parsed || !parsed.length) continue;
        
        // Log headers for debugging - SUMMARY format should have different headers than LISTALL
        if (parseResult.headers && parseResult.headers.length > 0) {
          Logger.log(`Channel summary headers (SUMMARY format): ${parseResult.headers.join('|')}`);
          Logger.log(`Number of summary rows: ${parsed.length}`);
          
          // Verify we're getting summary format (should have aggregated metrics, not individual sessions)
          // Summary format typically has columns like: Channel Name, Total Sessions, Avg Duration, etc.
          const hasSummaryColumns = parseResult.headers.some(h => 
            /total|average|avg|count|sum|sessions/i.test(h)
          );
          if (!hasSummaryColumns) {
            Logger.log(`WARNING: Headers don't look like SUMMARY format. Headers: ${parseResult.headers.join(', ')}`);
          }
        }
        
        // Each row represents a channel summary with all available columns
        // In SUMMARY format, each row is an aggregated channel summary, not an individual session
        parsed.forEach(row => {
          // Create a copy of the row object to preserve all columns
          const channelRow = {};
          Object.keys(row).forEach(key => {
            channelRow[key] = row[key];
          });
          
          // Ensure we have a channel identifier
          if (!channelRow['Channel Name'] && !channelRow['Channel'] && !channelRow['Channel ID']) {
            // Try to find channel name in other columns
            const channelName = row['Channel Name:'] || row['Channel:'] || row['Name'] || 
                              row['Channel Name'] || String(node);
            channelRow['Channel Name'] = channelName;
          }
          
          channelData.push(channelRow);
        });
        
        Utilities.sleep(200);
      } catch (e) {
        Logger.log(`Error fetching channel summary from node ${node}: ${e.toString()}`);
      }
    }
    
    Logger.log(`fetchChannelSummaryData_: Found ${channelData.length} channel summaries`);
  } catch (e) {
    Logger.log('fetchChannelSummaryData_ error: ' + e.toString());
  }
  
  return channelData;
}

function fetchLoggedInTechnicians_(cfg) {
  try {
    // Use isAnyTechAvailableOnChannel API to get currently logged in technicians
    // Reference: https://support.logmein.com/rescue/help/rescue-api-reference-guide
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    
    const nodes = cfg.nodes.map(n => Number(n)).filter(Number.isFinite);
    const loggedInTechs = new Set();
    
    // Query each channel to check for available technicians
    for (const node of nodes) {
      try {
        // Use isAnyTechAvailableOnChannel API per documentation
        const r = apiGet_(cfg.rescueBase, 'isAnyTechAvailableOnChannel.aspx', { 
          node: String(node) 
        }, cookie, 4, true);
        
        const responseText = (r.getContentText() || '').trim();
        if (!responseText || !/^OK/i.test(responseText)) continue;
        
        // Parse response - format may vary, but typically includes technician info
        // If response indicates technicians are available, we can also get their names
        // For now, we'll also check getSession_v3 to get actual technician names
        const sessionR = apiGet_(cfg.rescueBase, 'getSession_v3.aspx', { 
          node: String(node), 
          noderef: 'CHANNEL' 
        }, cookie, 4, true);
        
        const sessionText = (sessionR.getContentText() || '').trim();
        if (sessionText && /^OK/i.test(sessionText)) {
          const lines = sessionText.replace(/^OK\s*/i, '').split(/\r?\n/).filter(Boolean);
          if (lines.length >= 2) {
            const headers = lines[0].split('|').map(h => h.trim());
            const techIdx = headers.findIndex(h => /technician.*name/i.test(h));
            
            for (let i = 1; i < lines.length; i++) {
              const cols = lines[i].split('|');
              if (cols.length > techIdx && cols[techIdx]) {
                const techName = cols[techIdx].trim();
                if (techName) {
                  loggedInTechs.add(techName);
                }
              }
            }
          }
        }
        
        Utilities.sleep(200);
      } catch (e) {
        Logger.log(`Error checking technicians on channel node ${node}: ${e.toString()}`);
      }
    }
    
    // Format as [Technician Name, Status]
    const result = Array.from(loggedInTechs).sort().map(name => [name, 'Logged In']);
    Logger.log(`fetchLoggedInTechnicians_: Found ${result.length} logged in technicians`);
    return result;
  } catch (e) {
    Logger.log('fetchLoggedInTechnicians_ error: ' + e.toString());
    return [];
  }
}

function createDailySummarySheet_(ss, startDate, endDate) {
  try {
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    
    let summarySheet = ss.getSheetByName('Daily_Summary');
    if (!summarySheet) summarySheet = ss.insertSheet('Daily_Summary');
    
    summarySheet.clear();
    
    // Get all data from Sessions
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    
    const startIdx = headers.indexOf('start_time');
    const techIdx = headers.indexOf('technician_name');
    const statusIdx = headers.indexOf('session_status');
    const durationIdx = headers.indexOf('duration_total_seconds');
    const workIdx = headers.indexOf('duration_work_seconds');
    const pickupIdx = headers.indexOf('pickup_seconds');
    
    // Filter by date range
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
      return rowDate >= startStr && rowDate <= endStr;
    });
    
    // Group by date
    const dailyData = {};
    filtered.forEach(row => {
      if (!row[startIdx]) return;
      const dateStr = new Date(row[startIdx]).toISOString().split('T')[0];
      if (!dailyData[dateStr]) {
        dailyData[dateStr] = {
          sessions: [],
          techs: new Set(),
          totalWorkSeconds: 0
        };
      }
      dailyData[dateStr].sessions.push(row);
      if (row[techIdx]) dailyData[dateStr].techs.add(row[techIdx]);
      if (row[workIdx]) dailyData[dateStr].totalWorkSeconds += Number(row[workIdx]);
    });
    
    // Get all dates in range
    const dates = [];
    let currentDate = new Date(startDate);
    const endDateObj = new Date(endDate);
    while (currentDate <= endDateObj) {
      dates.push(currentDate.toISOString().split('T')[0]);
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    // Build summary table
    const summaryRows = [];
    
    // Header row with dates
    const headerRow = ['Metric'];
    dates.forEach(d => {
      const dObj = new Date(d + 'T00:00:00');
      headerRow.push(`${dObj.getMonth()+1}/${dObj.getDate()}/${dObj.getFullYear()}`);
    });
    headerRow.push('Totals/Averages');
    summaryRows.push(headerRow);
    
    // Total Node calculation summary
    summaryRows.push(['LMI:', ...Array(dates.length + 1).fill('')]);
    
    // Total sessions per day
    const totalSessionsRow = ['Total sessions'];
    let totalSessionsAll = 0;
    dates.forEach(d => {
      const count = dailyData[d] ? dailyData[d].sessions.length : 0;
      totalSessionsRow.push(count);
      totalSessionsAll += count;
    });
    totalSessionsRow.push(totalSessionsAll);
    summaryRows.push(totalSessionsRow);
    
    // Total duration/Work Time per day
    const totalDurationRow = ['Total duration/ Total Work Time'];
    let totalWorkSecondsAll = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.totalWorkSeconds > 0) {
        const hours = Math.floor(data.totalWorkSeconds / 3600);
        const minutes = Math.floor((data.totalWorkSeconds % 3600) / 60);
        const seconds = data.totalWorkSeconds % 60;
        totalDurationRow.push(`${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`);
        totalWorkSecondsAll += data.totalWorkSeconds;
      } else {
        totalDurationRow.push('0:00:00');
      }
    });
    const totalHours = Math.floor(totalWorkSecondsAll / 3600);
    const totalMins = Math.floor((totalWorkSecondsAll % 3600) / 60);
    const totalSecs = totalWorkSecondsAll % 60;
    totalDurationRow.push(`${totalHours}:${String(totalMins).padStart(2, '0')}:${String(totalSecs).padStart(2, '0')}`);
    summaryRows.push(totalDurationRow);
    
    // Avg Session duration per day
    const avgSessionRow = ['Avg Session'];
    let totalAvgSeconds = 0;
    let daysWithData = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.sessions.length > 0) {
        const durations = data.sessions.map(s => Number(s[durationIdx] || 0)).filter(Boolean);
        if (durations.length > 0) {
          const avg = durations.reduce((a, b) => a + b, 0) / durations.length;
          const mins = Math.floor(avg / 60);
          const secs = Math.floor(avg % 60);
          avgSessionRow.push(`0:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`);
          totalAvgSeconds += avg;
          daysWithData++;
        } else {
          avgSessionRow.push('0:00:00');
        }
      } else {
        avgSessionRow.push('0:00:00');
      }
    });
    if (daysWithData > 0) {
      const overallAvg = totalAvgSeconds / daysWithData;
      const mins = Math.floor(overallAvg / 60);
      const secs = Math.floor(overallAvg % 60);
      avgSessionRow.push(`0:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`);
    } else {
      avgSessionRow.push('0:00:00');
    }
    summaryRows.push(avgSessionRow);
    
    // Avg Pick-up Speed per day
    const avgPickupRow = ['Avg Pick-up Speed'];
    let totalPickupSeconds = 0;
    let totalPickupCount = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.sessions.length > 0) {
        const pickups = data.sessions.map(s => Number(s[pickupIdx] || 0)).filter(p => p > 0);
        if (pickups.length > 0) {
          const avg = pickups.reduce((a, b) => a + b, 0) / pickups.length;
          const mins = Math.floor(avg / 60);
          const secs = Math.floor(avg % 60);
          avgPickupRow.push(`0:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`);
          totalPickupSeconds += avg * pickups.length;
          totalPickupCount += pickups.length;
        } else {
          avgPickupRow.push('0:00:00');
        }
      } else {
        avgPickupRow.push('0:00:00');
      }
    });
    if (totalPickupCount > 0) {
      const overallAvg = totalPickupSeconds / totalPickupCount;
      const mins = Math.floor(overallAvg / 60);
      const secs = Math.floor(overallAvg % 60);
      avgPickupRow.push(`0:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`);
    } else {
      avgPickupRow.push('0:00:00');
    }
    summaryRows.push(avgPickupRow);
    
    // Percentage of daily calls/total sessions
    const pctRow = ['Percentage of daily calls/total sessions'];
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && totalSessionsAll > 0) {
        const pct = (data.sessions.length / totalSessionsAll).toFixed(2);
        pctRow.push(pct);
      } else {
        pctRow.push('0');
      }
    });
    pctRow.push('1');
    summaryRows.push(pctRow);
    
    // Average REAL (placeholder - need to understand what this means)
    const avgRealRow = ['Average REAL'];
    dates.forEach(() => avgRealRow.push('0:00:00'));
    avgRealRow.push('0:00:00');
    summaryRows.push(avgRealRow);
    
    // Empty row
    summaryRows.push(Array(dates.length + 2).fill(''));
    
    // Daily Calculations section
    summaryRows.push(['Daily Calculations', ...Array(dates.length + 1).fill('')]);
    summaryRows.push(['# of techs working will be a total of techs that picked up sessions that day', ...Array(dates.length + 1).fill('')]);
    
    // Total # of Techs working per day
    const techsWorkingRow = ['Total # of Techs working'];
    let totalTechsAll = 0;
    dates.forEach(d => {
      const count = dailyData[d] ? dailyData[d].techs.size : 0;
      techsWorkingRow.push(count);
      totalTechsAll += count;
    });
    techsWorkingRow.push(totalTechsAll);
    summaryRows.push(techsWorkingRow);
    
    // Total hours of technicians per day (assuming 8 hours per tech)
    const techHoursRow = ['Total hours of technicans'];
    dates.forEach(d => {
      const hours = dailyData[d] ? dailyData[d].techs.size * 8 : 0;
      techHoursRow.push(hours);
    });
    techHoursRow.push(totalTechsAll * 8);
    summaryRows.push(techHoursRow);
    
    // Average sessions per tech per day
    const avgSessionsPerTechRow = ['Average sessions per tech'];
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.techs.size > 0) {
        const avg = (data.sessions.length / data.techs.size).toFixed(1);
        avgSessionsPerTechRow.push(avg);
      } else {
        avgSessionsPerTechRow.push('0');
      }
    });
    avgSessionsPerTechRow.push((totalSessionsAll / Math.max(1, totalTechsAll)).toFixed(1));
    summaryRows.push(avgSessionsPerTechRow);
    
    // Average calls per tech (same as sessions)
    const avgCallsPerTechRow = ['Average calls per tech'];
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.techs.size > 0) {
        const avg = (data.sessions.length / data.techs.size).toFixed(1);
        avgCallsPerTechRow.push(avg);
      } else {
        avgCallsPerTechRow.push('0');
      }
    });
    avgCallsPerTechRow.push((totalSessionsAll / Math.max(1, totalTechsAll)).toFixed(1));
    summaryRows.push(avgCallsPerTechRow);
    
    // Empty rows
    summaryRows.push(Array(dates.length + 2).fill(''));
    summaryRows.push(Array(dates.length + 2).fill(''));
    summaryRows.push(Array(dates.length + 2).fill(''));
    
    // Summary section (Technician totals)
    summaryRows.push(['Summary', ...Array(dates.length + 1).fill('')]);
    const summaryHeaders = ['Technician Name', 'Total Sessions', '% Of Total sessions', 'Sessions per HR', 'Avg Pick-up Speed', 'Avg Duration', 'Average Work Time', ...Array(dates.length - 6).fill('')];
    summaryRows.push(summaryHeaders);
    
    // Calculate per-technician stats
    const techStats = {};
    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      if (!techStats[tech]) {
        techStats[tech] = {
          sessions: 0,
          durations: [],
          pickups: [],
          workSeconds: 0
        };
      }
      techStats[tech].sessions++;
      if (row[durationIdx]) techStats[tech].durations.push(Number(row[durationIdx]));
      if (row[pickupIdx]) techStats[tech].pickups.push(Number(row[pickupIdx]));
      if (row[workIdx]) techStats[tech].workSeconds += Number(row[workIdx]);
    });
    
    Object.keys(techStats).sort((a, b) => techStats[b].sessions - techStats[a].sessions).forEach(tech => {
      const stats = techStats[tech];
      const pct = totalSessionsAll > 0 ? ((stats.sessions / totalSessionsAll) * 100).toFixed(2) : '0';
      const sessionsPerHr = (stats.sessions / Math.max(1, (stats.workSeconds / 3600))).toFixed(2);
      const avgPickup = stats.pickups.length > 0 ? (stats.pickups.reduce((a, b) => a + b, 0) / stats.pickups.length) : 0;
      const avgPickupStr = avgPickup > 0 ? `0:${String(Math.floor(avgPickup / 60)).padStart(2, '0')}:${String(Math.floor(avgPickup % 60)).padStart(2, '0')}` : '0:00:00';
      const avgDur = stats.durations.length > 0 ? (stats.durations.reduce((a, b) => a + b, 0) / stats.durations.length) : 0;
      const avgDurStr = avgDur > 0 ? `0:${String(Math.floor(avgDur / 60)).padStart(2, '0')}:${String(Math.floor(avgDur % 60)).padStart(2, '0')}` : '0:00:00';
      const workTime = stats.workSeconds > 0 ? `0:${String(Math.floor(stats.workSeconds / 60)).padStart(2, '0')}:${String(Math.floor(stats.workSeconds % 60)).padStart(2, '0')}` : '0:00:00';
      
      const techRow = [tech, stats.sessions, pct, sessionsPerHr, avgPickupStr, avgDurStr, workTime, ...Array(dates.length - 6).fill('')];
      summaryRows.push(techRow);
    });
    
    // Write to sheet
    if (summaryRows.length > 0) {
      summarySheet.getRange(1, 1, summaryRows.length, summaryRows[0].length).setValues(summaryRows);
      summarySheet.getRange(1, 1, 1, summaryRows[0].length).setFontWeight('bold').setBackground('#E5E7EB');
      summarySheet.getRange(2, 1, 1, summaryRows[0].length).setFontWeight('bold');
      summarySheet.getRange(summaryRows.length - Object.keys(techStats).length, 1, 1, summaryRows[0].length).setFontWeight('bold').setBackground('#E5E7EB');
      summarySheet.setFrozenRows(1);
      summarySheet.setColumnWidth(1, 300);
      for (let i = 2; i <= summaryRows[0].length; i++) {
        summarySheet.setColumnWidth(i, 120);
      }
    }
    
    Logger.log('Daily summary sheet created');
  } catch (e) {
    Logger.log('createDailySummarySheet_ error: ' + e.toString());
  }
}

function createSupportDataSheet_(ss, startDate, endDate) {
  try {
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    
    let supportSheet = ss.getSheetByName('Support_Data');
    if (!supportSheet) supportSheet = ss.insertSheet('Support_Data');
    
    supportSheet.clear();
    
    // Get all data from Sessions
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    
    const startIdx = headers.indexOf('start_time');
    const techIdx = headers.indexOf('technician_name');
    const statusIdx = headers.indexOf('session_status');
    const durationIdx = headers.indexOf('duration_total_seconds');
    const workIdx = headers.indexOf('duration_work_seconds');
    const pickupIdx = headers.indexOf('pickup_seconds');
    const customerIdx = headers.indexOf('customer_name');
    const sessionIdIdx = headers.indexOf('session_id');
    const channelIdx = headers.indexOf('channel_name');
    const resolvedIdx = headers.indexOf('resolved_unresolved');
    const callingCardIdx = headers.indexOf('calling_card');
    
    // Filter by date range
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    
    // Header
    supportSheet.getRange(1, 1).setValue('📊 SUPPORT DATA SUMMARY');
    supportSheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#1A73E8');
    supportSheet.getRange(1, 1, 1, 8).merge();
    
    supportSheet.getRange(2, 1).setValue('Date Range:');
    supportSheet.getRange(2, 2).setValue(`${startStr} to ${endStr}`);
    supportSheet.getRange(2, 1).setFontWeight('bold');
    
    // Summary KPIs
    const kpiRow = 4;
    const totalSessions = filtered.length;
    const totalWorkSeconds = filtered.reduce((sum, row) => sum + (Number(row[workIdx]) || 0), 0);
    const totalWorkHours = (totalWorkSeconds / 3600).toFixed(1);
    const avgPickup = filtered.length > 0 ? 
      Math.round(filtered.reduce((sum, row) => sum + (Number(row[pickupIdx]) || 0), 0) / filtered.length) : 0;
    const resolvedCount = filtered.filter(row => row[resolvedIdx] === 'Resolved').length;
    const resolutionRate = totalSessions > 0 ? ((resolvedCount / totalSessions) * 100).toFixed(1) : '0';
    
    // Count Nova Wave sessions (calling card contains "Nova wave chat")
    const novaWaveCount = callingCardIdx >= 0 ? 
      filtered.filter(row => {
        const callingCard = String(row[callingCardIdx] || '').toLowerCase();
        return callingCard.includes('nova wave chat');
      }).length : 0;
    
    const kpis = [
      ['Total Sessions', totalSessions],
      ['Nova Wave Sessions', novaWaveCount],
      ['Total Work Hours', totalWorkHours + ' hrs'],
      ['Avg Pickup Time', avgPickup + ' sec'],
      ['Resolution Rate', resolutionRate + '%'],
      ['Resolved Sessions', resolvedCount],
      ['Unresolved Sessions', totalSessions - resolvedCount]
    ];
    
    for (let i = 0; i < kpis.length; i++) {
      const row = kpiRow + Math.floor(i / 3);
      const col = (i % 3) * 3 + 1;
      supportSheet.getRange(row, col).setValue(kpis[i][0]);
      supportSheet.getRange(row, col).setFontSize(11).setFontColor('#666666');
      supportSheet.getRange(row, col + 1).setValue(kpis[i][1]);
      supportSheet.getRange(row, col + 1).setFontSize(14).setFontWeight('bold').setFontColor('#1A73E8');
      supportSheet.getRange(row, col, 1, 2).setBorder(true, true, true, true, true, true);
    }
    
    // Technician Performance Table
    const tableRow = kpiRow + 3;
    supportSheet.getRange(tableRow, 1).setValue('Technician Performance');
    supportSheet.getRange(tableRow, 1).setFontSize(14).setFontWeight('bold');
    const tableHeaders = ['Technician', 'Sessions', 'Nova Wave Sessions', 'Avg Pickup (sec)', 'Avg Duration (min)', 'Work Hours', 'Resolution Rate'];
    supportSheet.getRange(tableRow + 1, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    supportSheet.getRange(tableRow + 1, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    // Calculate per-technician stats
    const techStats = {};
    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      if (!techStats[tech]) {
        techStats[tech] = {
          sessions: 0,
          durations: [],
          pickups: [],
          workSeconds: 0,
          resolved: 0,
          novaWave: 0
        };
      }
      techStats[tech].sessions++;
      if (row[durationIdx]) techStats[tech].durations.push(Number(row[durationIdx]));
      if (row[pickupIdx]) techStats[tech].pickups.push(Number(row[pickupIdx]));
      if (row[workIdx]) techStats[tech].workSeconds += Number(row[workIdx]);
      if (row[resolvedIdx] === 'Resolved') techStats[tech].resolved++;
      // Count Nova Wave sessions
      if (callingCardIdx >= 0 && row[callingCardIdx]) {
        const callingCard = String(row[callingCardIdx]).toLowerCase();
        if (callingCard.includes('nova wave chat')) {
          techStats[tech].novaWave++;
        }
      }
    });
    
    const techRows = Object.keys(techStats).sort((a, b) => techStats[b].sessions - techStats[a].sessions).map(tech => {
      const stats = techStats[tech];
      const avgPickup = stats.pickups.length > 0 ? Math.round(stats.pickups.reduce((a, b) => a + b, 0) / stats.pickups.length) : 0;
      const avgDur = stats.durations.length > 0 ? (stats.durations.reduce((a, b) => a + b, 0) / stats.durations.length / 60).toFixed(1) : '0';
      const workHours = (stats.workSeconds / 3600).toFixed(1);
      const resRate = stats.sessions > 0 ? ((stats.resolved / stats.sessions) * 100).toFixed(1) : '0';
      return [tech, stats.sessions, stats.novaWave, avgPickup, avgDur, workHours, resRate + '%'];
    });
    
    if (techRows.length > 0) {
      supportSheet.getRange(tableRow + 2, 1, techRows.length, tableHeaders.length).setValues(techRows);
    }
    
    // Channel Performance Summary (from API)
    const channelRow = tableRow + techRows.length + 4;
    supportSheet.getRange(channelRow, 1).setValue('Channel Performance Summary (from API)');
    supportSheet.getRange(channelRow, 1).setFontSize(14).setFontWeight('bold');
    
    // Fetch channel summary data from API using SUMMARY report type
    const cfg = getCfg_();
    const channelSummaryData = fetchChannelSummaryData_(cfg, startDate, endDate);
    
    if (channelSummaryData.length > 0) {
      // Get all unique column names from the channel summary data
      const allColumnNames = new Set();
      channelSummaryData.forEach(row => {
        Object.keys(row).forEach(key => allColumnNames.add(key));
      });
      
      // Sort columns for consistent display (put Channel Name first, then common metrics)
      const sortedColumns = Array.from(allColumnNames).sort((a, b) => {
        const priority = ['Channel Name', 'Channel', 'Channel ID', 'Sessions', 'Total Sessions', 
                          'Avg Duration', 'Average Duration', 'Avg Pickup', 'Average Pickup',
                          'Total Time', 'Active Time', 'Work Time', 'Waiting Time'];
        const aIdx = priority.indexOf(a);
        const bIdx = priority.indexOf(b);
        if (aIdx >= 0 && bIdx >= 0) return aIdx - bIdx;
        if (aIdx >= 0) return -1;
        if (bIdx >= 0) return 1;
        return a.localeCompare(b);
      });
      
      // Set headers
      supportSheet.getRange(channelRow + 1, 1, 1, sortedColumns.length).setValues([sortedColumns]);
      supportSheet.getRange(channelRow + 1, 1, 1, sortedColumns.length).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
      
      // Build rows with all available columns
      const channelRows = channelSummaryData.map(channelData => {
        return sortedColumns.map(col => {
          const value = channelData[col];
          if (value === null || value === undefined) return '';
          return String(value);
        });
      });
      
      if (channelRows.length > 0) {
        supportSheet.getRange(channelRow + 2, 1, channelRows.length, sortedColumns.length).setValues(channelRows);
      }
    } else {
      // Fallback to calculated data if API doesn't return channel summary
      const channelHeaders = ['Channel', 'Sessions', 'Avg Pickup (sec)', 'Avg Duration (min)'];
      supportSheet.getRange(channelRow + 1, 1, 1, channelHeaders.length).setValues([channelHeaders]);
      supportSheet.getRange(channelRow + 1, 1, 1, channelHeaders.length).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
      
      const channelStats = {};
      filtered.forEach(row => {
        const channel = row[channelIdx] || 'Unknown';
        if (!channelStats[channel]) {
          channelStats[channel] = {
            sessions: 0,
            durations: [],
            pickups: []
          };
        }
        channelStats[channel].sessions++;
        if (row[durationIdx]) channelStats[channel].durations.push(Number(row[durationIdx]));
        if (row[pickupIdx]) channelStats[channel].pickups.push(Number(row[pickupIdx]));
      });
      
      const channelRows = Object.keys(channelStats).sort((a, b) => channelStats[b].sessions - channelStats[a].sessions).map(channel => {
        const stats = channelStats[channel];
        const avgPickup = stats.pickups.length > 0 ? Math.round(stats.pickups.reduce((a, b) => a + b, 0) / stats.pickups.length) : 0;
        const avgDur = stats.durations.length > 0 ? (stats.durations.reduce((a, b) => a + b, 0) / stats.durations.length / 60).toFixed(1) : '0';
        return [channel, stats.sessions, avgPickup, avgDur];
      });
      
      if (channelRows.length > 0) {
        supportSheet.getRange(channelRow + 2, 1, channelRows.length, channelHeaders.length).setValues(channelRows);
      }
    }
    
    // Formatting
    supportSheet.setColumnWidth(1, 200);
    supportSheet.setColumnWidth(2, 120);
    supportSheet.setColumnWidth(3, 120);
    supportSheet.setColumnWidth(4, 120);
    supportSheet.setColumnWidth(5, 120);
    supportSheet.setColumnWidth(6, 120);
    supportSheet.setFrozenRows(1);
    
    Logger.log('Support Data sheet created');
  } catch (e) {
    Logger.log('createSupportDataSheet_ error: ' + e.toString());
  }
}

/* ===== Time Frame Selector ===== */
function getTimeFrameRange_(timeFrame) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let startDate, endDate;
  switch(timeFrame) {
    case 'Today':
      startDate = new Date(today);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(today);
      endDate.setHours(23, 59, 59, 999);
      break;
    case 'Yesterday':
      startDate = new Date(today);
      startDate.setDate(startDate.getDate() - 1);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(startDate);
      endDate.setHours(23, 59, 59, 999);
      break;
    case 'This Week':
      const dayOfWeek = today.getDay();
      const diff = today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
      startDate = new Date(today.getFullYear(), today.getMonth(), diff);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(today);
      endDate.setHours(23, 59, 59, 999);
      break;
    case 'Last Week':
    case 'Previous Week':
      const dayOfWeek2 = today.getDay();
      const diff2 = today.getDate() - dayOfWeek2 + (dayOfWeek2 === 0 ? -6 : 1);
      const thisMonday = new Date(today.getFullYear(), today.getMonth(), diff2);
      startDate = new Date(thisMonday);
      startDate.setDate(startDate.getDate() - 7);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(thisMonday);
      endDate.setDate(endDate.getDate() - 1);
      endDate.setHours(23, 59, 59, 999);
      break;
      
    case 'Last Month':
      // Get first day of last month (e.g., October 1st if today is November)
      const lastMonth = today.getMonth() - 1;
      const lastMonthYear = lastMonth < 0 ? today.getFullYear() - 1 : today.getFullYear();
      const lastMonthActual = lastMonth < 0 ? 11 : lastMonth;
      startDate = new Date(lastMonthYear, lastMonthActual, 1);
      startDate.setHours(0, 0, 0, 0);
      // Get last day of last month (e.g., October 31st if today is November)
      // Day 0 of current month gives us the last day of previous month
      endDate = new Date(today.getFullYear(), today.getMonth(), 0);
      endDate.setHours(23, 59, 59, 999);
      // Double-check: ensure endDate is actually in the previous month
      if (endDate.getMonth() !== lastMonthActual || endDate.getFullYear() !== lastMonthYear) {
        // Manually set to last day of last month
        endDate = new Date(lastMonthYear, lastMonthActual + 1, 0);
        endDate.setHours(23, 59, 59, 999);
      }
      break;
    case 'This Month':
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(today);
      endDate.setHours(23, 59, 59, 999);
      break;
    case 'Last Month':
      startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(today.getFullYear(), today.getMonth(), 0);
      endDate.setHours(23, 59, 59, 999);
      break;
    case 'Custom':
      const ss = SpreadsheetApp.getActive();
      const configSheet = ss.getSheetByName('Dashboard_Config');
      if (configSheet) {
        const customStart = configSheet.getRange('B4').getValue();
        const customEnd = configSheet.getRange('B5').getValue();
        if (customStart && customEnd) {
          startDate = new Date(customStart);
          startDate.setHours(0, 0, 0, 0);
          endDate = new Date(customEnd);
          endDate.setHours(23, 59, 59, 999);
          break;
        }
      }
      startDate = new Date(today);
      endDate = new Date(today);
      endDate.setHours(23, 59, 59, 999);
      break;
    default:
      startDate = new Date(today);
      endDate = new Date(today);
      endDate.setHours(23, 59, 59, 999);
  }
  return { startDate, endDate };
}

/* ===== Analytics Dashboard ===== */
function createAnalyticsDashboard() {
  try {
    const ss = SpreadsheetApp.getActive();
    createFilterControl_(ss);
    createMainAnalyticsPage_(ss);
    SpreadsheetApp.getActive().toast('Analytics dashboard created! Select time frame and click Refresh.');
  } catch (e) {
    Logger.log('createAnalyticsDashboard error: ' + e.toString());
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
  }
}

function createFilterControl_(ss) {
  let sh = ss.getSheetByName('Dashboard_Config');
  if (!sh) sh = ss.insertSheet('Dashboard_Config');
  sh.clear();
  sh.getRange(1, 1).setValue('📊 Dashboard Control Panel');
  sh.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#1A73E8');
  sh.getRange(1, 1, 1, 3).merge();
  sh.getRange(3, 1).setValue('Time Frame:');
  sh.getRange(3, 1).setFontWeight('bold');
  const timeFrames = ['Today', 'Yesterday', 'This Week', 'Last Week', 'This Month', 'Last Month', 'Custom'];
  const tfRange = sh.getRange(3, 2);
  tfRange.setValue('Today');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(timeFrames).setAllowInvalid(false).build();
  tfRange.setDataValidation(rule);
  sh.getRange(4, 1).setValue('Custom Start:');
  sh.getRange(4, 2).setValue('=TODAY()-7');
  sh.getRange(4, 2).setNumberFormat('mm/dd/yyyy');
  sh.getRange(5, 1).setValue('Custom End:');
  sh.getRange(5, 2).setValue('=TODAY()');
  sh.getRange(5, 2).setNumberFormat('mm/dd/yyyy');
  sh.getRange(7, 1).setValue('Selected Range:');
  sh.getRange(7, 1).setFontWeight('bold');
  sh.getRange(7, 2).setFormula('=IF(B3="Custom", B4&" to "&B5, B3)');
  sh.getRange(9, 1).setValue('Last Refreshed:');
  sh.getRange(9, 2).setValue('Never');
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 200);
}

function createMainAnalyticsPage_(ss) {
  let sh = ss.getSheetByName('Analytics_Dashboard');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Analytics_Dashboard');
  sh.getRange(1, 1).setValue('📊 RESCUE ANALYTICS DASHBOARD');
  sh.getRange(1, 1).setFontSize(20).setFontWeight('bold').setFontColor('#1A73E8');
  sh.getRange(1, 1, 1, 10).merge();
  sh.getRange(2, 1).setValue('Time Frame:');
  sh.getRange(2, 2).setFormula('=Dashboard_Config!B3');
  sh.getRange(2, 4).setValue('Last Updated:');
  sh.getRange(2, 5).setValue(new Date().toLocaleString());
  const kpiRow = 4;
  const kpiCards = [
    ['Total Sessions', '=COUNTA(Sessions!B2:B)'], // Column B is now session_id (was A)
    ['Active Sessions', '=COUNTIFS(Sessions!H2:H, "Active")'], // Column H is now session_status (was C)
    ['Nova Wave Sessions', ''], // Will be calculated dynamically in refreshAnalyticsDashboard_ based on time frame
    ['Avg Duration', '=IF(COUNT(Sessions!U2:U)>0, ROUND(AVERAGE(Sessions!U2:U)/60, 1)&" min", "0 min")'],
    ['Avg Pickup Time', '=IF(COUNT(Sessions!V2:V)>0, ROUND(AVERAGE(Sessions!V2:V))&" sec", "0 sec")'],
    ['Longest Session', '=IF(MAX(Sessions!U2:U)>0, ROUND(MAX(Sessions!U2:U)/60, 1)&" min", "0 min")'],
    ['SLA Hit %', '=IF(COUNT(Sessions!V2:V)>0, TEXT(COUNTIFS(Sessions!V2:V, "<=30")/COUNT(Sessions!V2:V), "0.0%"), "0%")'],
    ['Avg Sessions/Hour', '=IF(COUNTA(Sessions!B2:B)>0, ROUND(COUNTA(Sessions!B2:B)/8, 1), "0")'] // Column B is now session_id
  ];
  for (let i = 0; i < kpiCards.length; i++) {
    const row = kpiRow + Math.floor(i / 3);
    const col = (i % 3) * 3 + 1;
    sh.getRange(row, col).setValue(kpiCards[i][0]);
    sh.getRange(row, col).setFontSize(10).setFontColor('#666666');
    // Nova Wave Sessions (i=2) will be set dynamically, others use formulas
    if (i !== 2) {
      sh.getRange(row, col + 1).setFormula(kpiCards[i][1]);
    }
    sh.getRange(row, col + 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sh.getRange(row, col, 1, 2).setBorder(true, true, true, true, true, true);
  }
  
  // Add Live Active Sessions section (current sessions)
  const liveRow = kpiRow + 3;
  sh.getRange(liveRow, 1).setValue('🟢 LIVE ACTIVE SESSIONS');
  sh.getRange(liveRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#34A853');
  sh.getRange(liveRow, 1, 1, 6).merge();
  const liveHeaders = ['Technician', 'Customer', 'Start Time', 'Live Duration', 'Channel', 'Session ID'];
  sh.getRange(liveRow + 1, 1, 1, liveHeaders.length).setValues([liveHeaders]);
  sh.getRange(liveRow + 1, 1, 1, liveHeaders.length).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sh.getRange(liveRow + 1, 1, 1, liveHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
  const tableRow = liveRow + 18;
  sh.getRange(tableRow, 1).setValue('Team Performance');
  sh.getRange(tableRow, 1).setFontSize(14).setFontWeight('bold');
  const headers = ['Technician', 'Total Sessions', 'Avg Duration', 'Avg Pickup', 'SLA Hit %', 'Sessions/Hour', 'Total Work Time'];
  sh.getRange(tableRow + 1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(tableRow + 1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
  sh.getRange(tableRow + 1, 1, 1, headers.length).setBorder(false, false, false, false, false, false); // Remove borders
  const activeRow = tableRow + 15;
  sh.getRange(activeRow, 1).setValue('Active Sessions (Selected Time Frame)');
  sh.getRange(activeRow, 1).setFontSize(14).setFontWeight('bold');
  const activeHeaders = ['Technician', 'Customer Name', 'Start Time', 'Duration', 'Session ID', 'Calling Card'];
  sh.getRange(activeRow + 1, 1, 1, activeHeaders.length).setValues([activeHeaders]);
  sh.getRange(activeRow + 1, 1, 1, activeHeaders.length).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sh.getRange(activeRow + 1, 1, 1, activeHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
  const queueRow = 80; // Changed from activeRow + 18 to fixed row 80
  sh.getRange(queueRow, 1).setValue('Waiting Queue');
  sh.getRange(queueRow, 1).setFontSize(14).setFontWeight('bold');
  const queueHeaders = ['Channel', 'Customer', 'Waiting Since', 'Wait Duration'];
  sh.getRange(queueRow + 1, 1, 1, queueHeaders.length).setValues([queueHeaders]);
  sh.getRange(queueRow + 1, 1, 1, queueHeaders.length).setFontWeight('bold').setBackground('#EA8600').setFontColor('#FFFFFF');
  sh.getRange(queueRow + 1, 1, 1, queueHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 150);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 150);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 150);
  sh.setFrozenRows(1);
}

function refreshDashboardFromAPI() {
  try {
    const ss = SpreadsheetApp.getActive();
    const configSheet = ss.getSheetByName('Dashboard_Config');
    if (!configSheet) {
      SpreadsheetApp.getActive().toast('Please create dashboard first: Rescue → Analytics Dashboard');
      return;
    }
    const timeFrame = configSheet.getRange('B3').getValue() || 'Today';
    const range = getTimeFrameRange_(timeFrame);
    const cfg = getCfg_();
    const startISO = range.startDate.toISOString();
    const endISO = range.endDate.toISOString();
    Logger.log(`Pulling data for ${timeFrame}: ${startISO} to ${endISO}`);
    const rowsIngested = ingestTimeRangeToSheets_(startISO, endISO, cfg);
    refreshAnalyticsDashboard_(range.startDate, range.endDate);
    generateTechnicianTabs_(range.startDate, range.endDate);
    refreshAdvancedAnalyticsDashboard_(range.startDate, range.endDate);
    configSheet.getRange('B9').setValue(new Date().toLocaleString());
    SpreadsheetApp.getActive().toast(`✅ Dashboard refreshed! ${rowsIngested} rows ingested for ${timeFrame}`, 5);
  } catch (e) {
    Logger.log('refreshDashboardFromAPI error: ' + e.toString());
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
  }
}

function refreshAnalyticsDashboard_(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const cfg = getCfg_();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    const startIdx = headers.indexOf('start_time');
    const techIdx = headers.indexOf('technician_name');
    const statusIdx = headers.indexOf('session_status'); // Now in column 8 (was 3)
    const durationIdx = headers.indexOf('duration_total_seconds');
    const pickupIdx = headers.indexOf('pickup_seconds');
    const workIdx = headers.indexOf('duration_work_seconds');
    const customerIdx = headers.indexOf('customer_name'); // Now in column 27 (was 8)
    const channelIdx = headers.indexOf('channel_name');
    const sessionIdIdx = headers.indexOf('session_id'); // Now in column 2 (was 1)
    const callingCardIdx = headers.indexOf('calling_card'); // Define once at the top
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
      return rowDate >= startStr && rowDate <= endStr;
    });
    // Get performance/summary data from API for accurate averages
    const performanceData = fetchPerformanceSummaryData_(cfg, startDate, endDate);
    
    // Build team data from filtered sessions (for session counts, SLA, work hours, durations)
    const teamData = {};
    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      if (!teamData[tech]) {
        teamData[tech] = { sessions: 0, pickups: [], durations: [], workSeconds: 0, slaHits: 0 };
      }
      teamData[tech].sessions++;
      if (row[pickupIdx]) {
        teamData[tech].pickups.push(Number(row[pickupIdx]));
        if (Number(row[pickupIdx]) <= 30) teamData[tech].slaHits++;
      }
      if (row[durationIdx]) {
        teamData[tech].durations.push(Number(row[durationIdx]));
      }
      if (row[workIdx]) teamData[tech].workSeconds += Number(row[workIdx]);
    });
    
    // Merge performance data from API with session counts
    const teamRows = Object.keys(teamData).map(tech => {
      const td = teamData[tech];
      const perf = performanceData[tech] || {};
      
      // Calculate avg duration from actual session durations (in seconds, convert to minutes)
      const avgDur = td.durations.length > 0 ? 
        (td.durations.reduce((a, b) => a + b, 0) / td.durations.length / 60).toFixed(1) : '0';
      
      const avgPickup = perf.avgPickup ? String(Math.round(perf.avgPickup)) : 
                        (td.pickups.length > 0 ? Math.round(td.pickups.reduce((a,b) => a+b, 0) / td.pickups.length).toFixed(0) : '0');
      
      const slaPct = td.pickups.length > 0 ? ((td.slaHits / td.pickups.length) * 100).toFixed(1) : '0';
      const days = Math.max(1, (new Date(endDate) - new Date(startDate)) / (1000*60*60*24));
      const sessionsPerHour = (td.sessions / days / 8).toFixed(1);
      const workHours = (td.workSeconds / 3600).toFixed(1);
      return [tech, String(td.sessions), avgDur + ' min', avgPickup + ' sec', slaPct + '%', sessionsPerHour, workHours + ' hrs'];
    }).sort((a, b) => Number(b[1]) - Number(a[1]));
    
    // Always get current sessions from getSession_v3 API (for real-time queue and active sessions)
    // This refreshes regardless of time frame - always shows current live data
    // Reference: https://support.logmein.com/rescue/help/rescue-api-reference-guide
    const currentSessionData = fetchCurrentSessions_(cfg);
    
    // LIVE Active Sessions from getSession_v3 API
    const liveActiveRows = currentSessionData.active.slice(0, 25).map(session => {
      const startTime = session.startTime ? new Date(session.startTime).toLocaleString('en-US', { 
        timeZone: 'America/New_York',
        hour: 'numeric', 
        minute: '2-digit', 
        second: '2-digit', 
        hour12: true 
      }) : '';
      const now = new Date();
      const startDate = session.startTime ? new Date(session.startTime) : now;
      const liveDurationSec = Math.floor((now - startDate) / 1000);
      const hours = Math.floor(liveDurationSec / 3600);
      const minutes = Math.floor((liveDurationSec % 3600) / 60);
      const seconds = liveDurationSec % 60;
      const liveDuration = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
      return [
        session.technician || '', 
        session.customer || 'Anonymous', 
        startTime, 
        liveDuration,
        session.channel || '—',
        session.sessionId || ''
      ];
    });
    
    // Get active sessions from filtered data (for the selected time frame)
    // Shows all sessions (not just Active status) from the selected time frame
    // Format: Technician, Customer Name, Start Time, Duration, Session ID, Calling Card
    // callingCardIdx already defined above
    const activeRows = filtered.slice(0, 50).map(row => {
      const startTime = row[startIdx] ? new Date(row[startIdx]).toLocaleString('en-US', { 
        timeZone: 'America/New_York',
        hour: 'numeric', 
        minute: '2-digit', 
        second: '2-digit',
        hour12: true 
      }) : '';
      const duration = row[durationIdx] ? `${Math.floor(row[durationIdx]/60)}:${String(row[durationIdx]%60).padStart(2,'0')}` : '0:00';
      const customerName = row[customerIdx] || 'Anonymous';
      const callingCard = row[callingCardIdx] || '';
      return [row[techIdx] || '', customerName, startTime, duration, row[sessionIdIdx] || '', callingCard];
    });
    
    // Waiting Queue from getSession_v3 API
    const waitingRows = currentSessionData.waiting.map(session => {
      const waitSince = session.startTime ? new Date(session.startTime).toLocaleString('en-US', { 
        timeZone: 'America/New_York',
        hour: 'numeric', 
        minute: '2-digit', 
        hour12: true 
      }) : '';
      const waitDur = session.waitDuration ? `${Math.floor(session.waitDuration/60)}:${String(session.waitDuration%60).padStart(2,'0')}` : '0:00';
      return [session.channel || '—', session.customer || '—', waitSince, waitDur];
    });
    const dashboardSheet = ss.getSheetByName('Analytics_Dashboard');
    if (dashboardSheet) {
      // Update LIVE Active Sessions (headers at row 8, data starts at row 9)
      const liveHeaderRow = 8; // liveRow + 1 = 7 + 1 = 8
      const liveDataRow = 9; // liveHeaderRow + 1
      
      // Ensure headers are always present
      const liveHeaders = ['Technician', 'Customer', 'Start Time', 'Live Duration', 'Channel', 'Session ID'];
      dashboardSheet.getRange(liveHeaderRow, 1, 1, liveHeaders.length).setValues([liveHeaders]);
      dashboardSheet.getRange(liveHeaderRow, 1, 1, liveHeaders.length).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
      dashboardSheet.getRange(liveHeaderRow, 1, 1, liveHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
      
      // Clear data rows and populate
      dashboardSheet.getRange(liveDataRow, 1, 100, 6).clearContent();
      if (liveActiveRows.length > 0) {
        dashboardSheet.getRange(liveDataRow, 1, liveActiveRows.length, 6).setValues(liveActiveRows);
      } else {
        dashboardSheet.getRange(liveDataRow, 1).setValue('No active sessions');
        dashboardSheet.getRange(liveDataRow, 1).setFontStyle('italic').setFontColor('#999999');
      }
      
      // Update Team Performance section
      // Row 25: "Team Performance" title (keep it)
      // Row 26: Headers (keep/update them)
      // Row 27+: Data rows
      const teamTitleRow = 25;
      const teamHeaderRow = 26;
      const teamDataRow = 27;
      
      // Ensure title is present
      dashboardSheet.getRange(teamTitleRow, 1).setValue('Team Performance');
      dashboardSheet.getRange(teamTitleRow, 1).setFontSize(14).setFontWeight('bold');
      
      // Ensure headers are always present
      const teamHeaders = ['Technician', 'Total Sessions', 'Avg Duration', 'Avg Pickup', 'SLA Hit %', 'Sessions/Hour', 'Total Work Time'];
      dashboardSheet.getRange(teamHeaderRow, 1, 1, teamHeaders.length).setValues([teamHeaders]);
      dashboardSheet.getRange(teamHeaderRow, 1, 1, teamHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
      dashboardSheet.getRange(teamHeaderRow, 1, 1, teamHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
      
      // Clear data rows and populate
      dashboardSheet.getRange(teamDataRow, 1, 100, 7).clearContent();
      if (teamRows.length > 0) {
        dashboardSheet.getRange(teamDataRow, 1, teamRows.length, 7).setValues(teamRows);
      }
      
      // Always get currently logged in technicians from API (refreshes regardless of time frame)
      // Uses isAnyTechAvailableOnChannel API per documentation
      const loggedInTechs = fetchLoggedInTechnicians_(cfg);
      
      // Update Currently Logged In Technicians (row 33)
      const techStatusRow = 33;
      dashboardSheet.getRange(techStatusRow, 1).setValue('👥 Currently Logged In Technicians');
      dashboardSheet.getRange(techStatusRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#1A73E8');
      dashboardSheet.getRange(techStatusRow, 1, 1, 3).merge();
      dashboardSheet.getRange(techStatusRow + 1, 1, 100, 2).clearContent();
      if (loggedInTechs.length > 0) {
        dashboardSheet.getRange(techStatusRow + 1, 1, loggedInTechs.length, 2).setValues(loggedInTechs);
      } else {
        dashboardSheet.getRange(techStatusRow + 1, 1).setValue('No technicians currently logged in');
        dashboardSheet.getRange(techStatusRow + 1, 1).setFontStyle('italic').setFontColor('#999999');
      }
      
      // Update Active Sessions section (shows sessions from selected time frame)
      // Row 40: Title (updates based on time frame)
      // Row 41: Headers
      // Row 42+: Data rows
      const activeTitleRow = 40;
      const activeHeaderRow = 41;
      const activeDataRow = 42;
      
      // Update title to reflect selected time frame
      const configSheet = ss.getSheetByName('Dashboard_Config');
      const timeFrame = configSheet ? (configSheet.getRange('B3').getValue() || 'Selected Time Frame') : 'Selected Time Frame';
      dashboardSheet.getRange(activeTitleRow, 1).setValue(`Active Sessions (${timeFrame})`);
      dashboardSheet.getRange(activeTitleRow, 1).setFontSize(14).setFontWeight('bold');
      
      // Ensure headers are always present (includes Customer Name and Calling Card)
      const activeHeaders = ['Technician', 'Customer Name', 'Start Time', 'Duration', 'Session ID', 'Calling Card'];
      dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setValues([activeHeaders]);
      dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
      dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
      
      // Clear data rows and populate with sessions from selected time frame
      dashboardSheet.getRange(activeDataRow, 1, 100, 6).clearContent();
      if (activeRows.length > 0) {
        dashboardSheet.getRange(activeDataRow, 1, activeRows.length, 6).setValues(activeRows);
      } else {
        dashboardSheet.getRange(activeDataRow, 1).setValue('No sessions found for selected time frame');
        dashboardSheet.getRange(activeDataRow, 1).setFontStyle('italic').setFontColor('#999999');
      }
      
      // Update Waiting Queue (row 80) - Keep columns visible even if empty
      const queueHeaderRow = 80;
      dashboardSheet.getRange(queueHeaderRow, 1).setValue('⏳ Waiting Queue');
      dashboardSheet.getRange(queueHeaderRow, 1).setFontSize(14).setFontWeight('bold');
      const queueHeaders = ['Channel', 'Customer', 'Waiting Since', 'Wait Duration'];
      dashboardSheet.getRange(queueHeaderRow + 1, 1, 1, 4).setValues([queueHeaders]);
      dashboardSheet.getRange(queueHeaderRow + 1, 1, 1, 4).setFontWeight('bold').setBackground('#EA8600').setFontColor('#FFFFFF');
      dashboardSheet.getRange(queueHeaderRow + 1, 1, 1, 4).setBorder(false, false, false, false, false, false); // Remove borders
      dashboardSheet.getRange(queueHeaderRow + 2, 1, 100, 4).clearContent();
      if (waitingRows.length > 0) {
        dashboardSheet.getRange(queueHeaderRow + 2, 1, waitingRows.length, 4).setValues(waitingRows);
      } else {
        dashboardSheet.getRange(queueHeaderRow + 2, 1).setValue('No sessions waiting');
        dashboardSheet.getRange(queueHeaderRow + 2, 1).setFontStyle('italic').setFontColor('#999999');
      }
      
      // Update Nova Wave Sessions count for selected time frame
      const novaWaveCount = filtered.filter(row => {
        const callingCard = String(row[callingCardIdx] || '').toLowerCase();
        return callingCard.includes('nova wave chat');
      }).length;
      
      // Find and update Nova Wave Sessions KPI card (3rd card in first row: row 4, column 8)
      // Card layout: Card 0 (cols 1-2), Card 1 (cols 4-5), Card 2 (cols 7-8)
      const novaWaveRow = 4; // kpiRow (first row)
      const novaWaveCol = 8; // 3rd card value column (col 7 is label, col 8 is value)
      dashboardSheet.getRange(novaWaveRow, novaWaveCol).setValue(novaWaveCount);
      dashboardSheet.getRange(novaWaveRow, novaWaveCol).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
      
      dashboardSheet.getRange(2, 5).setValue(new Date().toLocaleString());
    }
  } catch (e) {
    Logger.log('refreshAnalyticsDashboard_ error: ' + e.toString());
  }
}

function generateTechnicianTabs_(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    
    // Get all existing technician tabs and clear them first to prevent stale data
    const allSheets = ss.getSheets();
    const reservedSheets = ['Sessions', 'Analytics_Dashboard', 'Dashboard_Config', 'Daily_Summary', 'Support_Data', 'Progress'];
    const existingTechSheets = allSheets.filter(sheet => {
      const sheetName = sheet.getName();
      return !reservedSheets.some(reserved => sheetName === reserved);
    });
    
    // Clear all existing technician tabs before regenerating
    for (const techSheet of existingTechSheets) {
      try {
        techSheet.clear();
        Logger.log(`Cleared existing technician tab: ${techSheet.getName()}`);
      } catch (e) {
        Logger.log(`Error clearing sheet ${techSheet.getName()}: ${e.toString()}`);
      }
    }
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    const startIdx = headers.indexOf('start_time');
    const techIdx = headers.indexOf('technician_name');
    const statusIdx = headers.indexOf('session_status');
    const durationIdx = headers.indexOf('duration_total_seconds');
    const pickupIdx = headers.indexOf('pickup_seconds');
    const workIdx = headers.indexOf('duration_work_seconds');
    const customerIdx = headers.indexOf('customer_name'); // Now in column 27 (was 8)
    const sessionIdIdx = headers.indexOf('session_id'); // Now in column 2 (was 1)
    const phoneIdx = headers.indexOf('customer_phone_number'); // Now customer_phone_number (was company_name/column 24)
    const companyIdx = headers.indexOf('customer_phone_number'); // Use phoneIdx for this
    const callingCardIdx = headers.indexOf('calling_card'); // Define once at the top of function
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx] || !row[techIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    const techs = [...new Set(filtered.map(row => row[techIdx]).filter(Boolean))];
    for (const techName of techs) {
      const safeName = techName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
      let techSheet = ss.getSheetByName(safeName);
      if (!techSheet) techSheet = ss.insertSheet(safeName);
      const techRows = filtered.filter(row => row[techIdx] === techName);
      // Clear again to ensure clean slate (in case sheet existed)
      techSheet.clear();
      techSheet.getRange(1, 1).setValue(`👤 ${techName} - Personal Dashboard`);
      techSheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#9C27B0');
      techSheet.getRange(1, 1, 1, 5).merge();
      techSheet.getRange(2, 1).setValue('Time Frame:');
      techSheet.getRange(2, 2).setFormula('=Dashboard_Config!B3');
      techSheet.getRange(2, 4).setValue('Last Updated:');
      techSheet.getRange(2, 5).setValue(new Date().toLocaleString());
      const durations = techRows.map(r => r[durationIdx]).filter(Boolean).map(Number);
      const pickups = techRows.map(r => r[pickupIdx]).filter(Boolean).map(Number);
      const workSeconds = techRows.map(r => r[workIdx]).filter(Boolean).reduce((a,b) => a + Number(b), 0);
      const completed = techRows.filter(r => r[statusIdx] === 'Ended').length;
      const active = techRows.filter(r => r[statusIdx] === 'Active').length;
      const avgDur = durations.length > 0 ? (durations.reduce((a,b) => a+b, 0) / durations.length / 60).toFixed(1) : '0';
      const avgPickup = pickups.length > 0 ? (pickups.reduce((a,b) => a+b, 0) / pickups.length / 60).toFixed(1) : '0';
      const slaHits = pickups.filter(p => p <= 30).length;
      const slaPct = pickups.length > 0 ? ((slaHits / pickups.length) * 100).toFixed(1) : '0';
      const days = Math.max(1, (new Date(endDate) - new Date(startDate)) / (1000*60*60*24));
      const sessionsPerHour = (techRows.length / days / 8).toFixed(1);
      const workHours = (workSeconds / 3600).toFixed(1);
      // Count Nova Wave sessions for this technician
      // callingCardIdx already defined above
      const novaWaveCount = callingCardIdx >= 0 ? 
        techRows.filter(row => {
          const callingCard = String(row[callingCardIdx] || '').toLowerCase();
          return callingCard.includes('nova wave chat');
        }).length : 0;
      
      const kpiRow = 4;
      const kpis = [
        ['Total Sessions', String(techRows.length)],
        ['Nova Wave Sessions', String(novaWaveCount)],
        ['Avg Duration', avgDur + ' min'],
        ['Avg Pickup Time', avgPickup + ' min'],
        ['SLA Hit %', slaPct + '%'],
        ['Total Work Hours', workHours + ' hrs'],
        ['Sessions/Hour', sessionsPerHour],
        ['Completed', String(completed)],
        ['Active', String(active)]
      ];
      for (let i = 0; i < kpis.length; i++) {
        const row = kpiRow + Math.floor(i / 2);
        const col = (i % 2) * 3 + 1;
        techSheet.getRange(row, col).setValue(kpis[i][0]);
        techSheet.getRange(row, col).setFontSize(11).setFontColor('#666666');
        techSheet.getRange(row, col + 1).setValue(kpis[i][1]);
        techSheet.getRange(row, col + 1).setFontSize(16).setFontWeight('bold').setFontColor('#9C27B0');
        techSheet.getRange(row, col, 1, 2).setBorder(true, true, true, true, true, true);
      }
      const detailRow = kpiRow + Math.ceil(kpis.length / 2) + 2;
      techSheet.getRange(detailRow, 1).setValue('Session Details');
      techSheet.getRange(detailRow, 1).setFontSize(14).setFontWeight('bold');
      // Updated column order: Date, Session ID, Status, Customer Name, Phone Number, Duration, Pickup
      // Column 3: Status (was Customer Name position)
      // Column 4: Customer Name (was Phone Number position)
      // Column 5: Phone Number (was Company Name position)
      // Removed old Status column (was column 8)
      const detailHeaders = ['Date', 'Session ID', 'Status', 'Customer Name', 'Phone Number', 'Duration (hh:mm)', 'Pickup (sec)'];
      techSheet.getRange(detailRow + 1, 1, 1, detailHeaders.length).setValues([detailHeaders]);
      techSheet.getRange(detailRow + 1, 1, 1, detailHeaders.length).setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF');
      
      // Get resolved_unresolved index for status information
      const resolvedIdx = headers.indexOf('resolved_unresolved');
      const callerNameIdx = headers.indexOf('customer_name'); // customer_name is now in column 27 (was caller_phone)
      // callingCardIdx already defined above
      
      // Format duration from seconds to hh:mm
      const formatDuration = (seconds) => {
        if (!seconds || seconds === 0) return '0:00';
        const totalSeconds = Math.floor(Number(seconds));
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        return `${hours}:${String(minutes).padStart(2, '0')}`;
      };
      
      const detailRows = techRows.slice(0, 50).map(row => {
        const date = row[startIdx] ? new Date(row[startIdx]).toISOString().split('T')[0] : '';
        
        // Status: combine session_status and resolved_unresolved if available
        // Remove calling card info (rc, ps, sp, etc.)
        let status = row[statusIdx] || '';
        if (resolvedIdx >= 0 && row[resolvedIdx]) {
          const resolved = String(row[resolvedIdx]).trim();
          if (resolved && resolved !== '') {
            // If resolved_unresolved has value like "Closed by technician" or "Closed by active customer", use it
            if (resolved.toLowerCase().includes('closed by')) {
              status = resolved;
            } else if (status && resolved) {
              status = `${status} - ${resolved}`;
            } else if (resolved) {
              status = resolved;
            }
          }
        }
        
        // Remove calling card info from status (e.g., "calling card or applet - rc ps sp")
        // Remove patterns like "- rc", "- ps", "- sp", "rc ps sp", "applet", "calling card"
        if (status) {
          status = String(status)
            .replace(/calling card or applet\s*-?\s*/gi, '')
            .replace(/\s*-?\s*(rc|ps|sp|applet|calling card)\s*/gi, '')
            .replace(/\s*-\s*(rc|ps|sp)\s*/gi, '')
            .replace(/\s+(rc|ps|sp)\s+/gi, ' ')
            .trim();
        }
        
        // Customer name is now in column 27 (was caller_phone position)
        const customerName = row[customerIdx] || 'Anonymous';
        
        // Phone number from customer_phone_number (was company_name, now in column 25)
        const phoneNumber = row[phoneIdx] || '';
        
        // Duration in hh:mm format
        const dur = row[durationIdx] ? formatDuration(row[durationIdx]) : '0:00';
        
        // Pickup time in seconds (not minutes)
        const pickup = row[pickupIdx] ? String(Math.round(Number(row[pickupIdx]))) : '0';
        
        // New column order: Date, Session ID, Status, Customer Name, Phone Number, Duration, Pickup
        return [
          date, 
          row[sessionIdIdx] || '', 
          status,
          customerName, 
          phoneNumber, 
          dur, 
          pickup
        ];
      });
      if (detailRows.length > 0) {
        techSheet.getRange(detailRow + 2, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
      }
      techSheet.setColumnWidth(1, 100);
      techSheet.setColumnWidth(2, 150);
      techSheet.setColumnWidth(3, 150);
      techSheet.setColumnWidth(4, 120);
      techSheet.setColumnWidth(5, 150);
      techSheet.setColumnWidth(6, 120);
      techSheet.setColumnWidth(7, 120);
      techSheet.setColumnWidth(8, 100);
    }
    Logger.log(`Generated ${techs.length} technician tabs`);
  } catch (e) {
    Logger.log('generateTechnicianTabs_ error: ' + e.toString());
  }
}

/* ===== API Smoke Test ===== */
function apiSmokeTest() {
  try {
    const cfg = getCfg_();
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    setReportAreaSession_(cfg.rescueBase, cookie);
    setReportTypeListAll_(cfg.rescueBase, cookie);
    setOutputXMLOrFallback_(cfg.rescueBase, cookie);
    setDelimiter_(cfg.rescueBase, cookie, '|');
    const today = new Date();
    const todayET = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const todayETStr = isoDate_(todayET);
    setReportDate_(cfg.rescueBase, cookie, todayETStr, todayETStr);
    setReportTimeAllDay_(cfg.rescueBase, cookie);
    const nodes = cfg.nodes.map(n => Number(n)).filter(Number.isFinite);
    const noderefs = ['NODE', 'CHANNEL'];
    let allRows = [];
    let nodeStats = [];
    let totalParsed = 0;
    let totalMapped = 0;
    for (let i = 0; i < noderefs.length; i++) {
      const nr = noderefs[i];
      for (let j = 0; j < nodes.length; j++) {
        const node = nodes[j];
        try {
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          if (!t || !/^OK/i.test(t)) {
            nodeStats.push({ node, noderef: nr, parsed: 0, mapped: 0, status: 'No data' });
            continue;
          }
          const parseResult = parsePipe_(t, '|');
          const parsed = parseResult.rows || [];
          totalParsed += parsed.length;
          if (parsed.length > 0) {
            const mapped = parsed.map(mapRow_).filter(r => r && r.session_id);
            totalMapped += mapped.length;
            allRows.push(...mapped.slice(0, 20));
            nodeStats.push({ node, noderef: nr, parsed: parsed.length, mapped: mapped.length, status: 'OK' });
          } else {
            nodeStats.push({ node, noderef: nr, parsed: 0, mapped: 0, status: 'Empty' });
          }
          Utilities.sleep(200);
        } catch (e) {
          nodeStats.push({ node, noderef: nr, parsed: 0, mapped: 0, status: 'Error: ' + e.toString().substring(0, 30) });
        }
      }
    }
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('API_Smoke_Test');
    if (sh) ss.deleteSheet(sh);
    sh = ss.insertSheet('API_Smoke_Test');
    sh.getRange(1, 1).setValue('🔍 API Smoke Test Results');
    sh.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sh.getRange(3, 1, 1, 2).setValues([['Total Parsed Rows', totalParsed]]);
    sh.getRange(4, 1, 1, 2).setValues([['Total Mapped Rows', totalMapped]]);
    sh.getRange(5, 1, 1, 2).setValues([['Date Tested', todayETStr]]);
    sh.getRange(7, 1).setValue('Node Coverage');
    sh.getRange(7, 1).setFontWeight('bold');
    sh.getRange(8, 1, 1, 5).setValues([['Node', 'Type', 'Parsed', 'Mapped', 'Status']]);
    sh.getRange(8, 1, 1, 5).setFontWeight('bold').setBackground('#E5E7EB');
    const nodeData = nodeStats.map(s => [s.node, s.noderef, s.parsed, s.mapped, s.status]);
    if (nodeData.length > 0) {
      sh.getRange(9, 1, nodeData.length, 5).setValues(nodeData);
    }
    if (allRows.length > 0) {
      sh.getRange(8 + nodeData.length + 3, 1).setValue('Sample Data (First 20 rows)');
      sh.getRange(8 + nodeData.length + 3, 1).setFontWeight('bold');
      const sampleHeaders = ['session_id', 'start_time', 'technician_name', 'session_status', 'duration_total_seconds', 'pickup_seconds', 'channel_name'];
      sh.getRange(8 + nodeData.length + 4, 1, 1, sampleHeaders.length).setValues([sampleHeaders]);
      sh.getRange(8 + nodeData.length + 4, 1, 1, sampleHeaders.length).setFontWeight('bold').setBackground('#E5E7EB');
      const sampleData = allRows.slice(0, 20).map(r => [
        r.session_id || '', r.start_time || '', r.technician_name || '', r.session_status || '',
        r.duration_total_seconds || 0, r.pickup_seconds || 0, r.channel_name || ''
      ]);
      if (sampleData.length > 0) {
        sh.getRange(8 + nodeData.length + 5, 1, sampleData.length, sampleHeaders.length).setValues(sampleData);
      }
    }
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 150);
    sh.setColumnWidth(4, 120);
    sh.setColumnWidth(5, 120);
    sh.setColumnWidth(6, 100);
    sh.setColumnWidth(7, 150);
    SpreadsheetApp.getActive().toast(`Smoke test complete: ${totalParsed} parsed, ${totalMapped} mapped`, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Smoke test error: ' + e.toString().substring(0, 50));
    Logger.log('apiSmokeTest error: ' + e.toString());
  }
}

/* ===== Advanced Analytics Dashboard ===== */
function createAdvancedAnalyticsDashboard() {
  try {
    const ss = SpreadsheetApp.getActive();
    let analyticsSheet = ss.getSheetByName('Advanced_Analytics');
    if (analyticsSheet) ss.deleteSheet(analyticsSheet);
    analyticsSheet = ss.insertSheet('Advanced_Analytics');
    
    // Header
    analyticsSheet.getRange(1, 1).setValue('📈 ADVANCED ANALYTICS DASHBOARD');
    analyticsSheet.getRange(1, 1).setFontSize(20).setFontWeight('bold').setFontColor('#1A73E8');
    analyticsSheet.getRange(1, 1, 1, 10).merge();
    analyticsSheet.getRange(2, 1).setValue('Time Frame:');
    analyticsSheet.getRange(2, 2).setFormula('=Dashboard_Config!B3');
    analyticsSheet.getRange(2, 4).setValue('Last Updated:');
    analyticsSheet.getRange(2, 5).setValue(new Date().toLocaleString());
    
    // Get time frame from config
    const configSheet = ss.getSheetByName('Dashboard_Config');
    const timeFrame = configSheet ? (configSheet.getRange('B3').getValue() || 'Today') : 'Today';
    const range = getTimeFrameRange_(timeFrame);
    
    // Generate all analytics sections
    refreshAdvancedAnalyticsDashboard_(range.startDate, range.endDate);
    
    SpreadsheetApp.getActive().toast('Advanced Analytics dashboard created!', 3);
  } catch (e) {
    Logger.log('createAdvancedAnalyticsDashboard error: ' + e.toString());
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
  }
}

// Refresh Advanced Analytics Dashboard based on time frame
function refreshAdvancedAnalyticsDashboard_(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    let analyticsSheet = ss.getSheetByName('Advanced_Analytics');
    
    // Create sheet if it doesn't exist
    if (!analyticsSheet) {
      analyticsSheet = ss.insertSheet('Advanced_Analytics');
      analyticsSheet.getRange(1, 1).setValue('📈 ADVANCED ANALYTICS DASHBOARD');
      analyticsSheet.getRange(1, 1).setFontSize(20).setFontWeight('bold').setFontColor('#1A73E8');
      analyticsSheet.getRange(1, 1, 1, 10).merge();
      analyticsSheet.getRange(2, 1).setValue('Time Frame:');
      analyticsSheet.getRange(2, 2).setFormula('=Dashboard_Config!B3');
      analyticsSheet.getRange(2, 4).setValue('Last Updated:');
    }
    
    // Update last updated timestamp
    analyticsSheet.getRange(2, 5).setValue(new Date().toLocaleString());
    
    // Clear existing content (except header rows 1-2)
    const lastRow = analyticsSheet.getLastRow();
    if (lastRow > 2) {
      analyticsSheet.getRange(3, 1, lastRow - 2, analyticsSheet.getLastColumn()).clearContent();
    }
    
    // Generate all analytics sections
    let currentRow = 4;
    
    // 1. Peak Hours & Day of Week Analysis
    currentRow = createPeakHoursAnalysis_(analyticsSheet, currentRow, startDate, endDate);
    currentRow += 5;
    
    // 2. Technician Effectiveness Comparison
    currentRow = createTechnicianEffectiveness_(analyticsSheet, currentRow, startDate, endDate);
    currentRow += 5;
    
    // 3. Repeat Customer Analysis
    currentRow = createRepeatCustomerAnalysis_(analyticsSheet, currentRow, startDate, endDate);
    currentRow += 5;
    
    // 4. Trend Analysis
    currentRow = createTrendAnalysis_(analyticsSheet, currentRow);
    currentRow += 5;
    
    // 5. Technician Utilization Rate
    currentRow = createUtilizationRate_(analyticsSheet, currentRow, startDate, endDate);
    currentRow += 5;
    
    // 6. Time to Resolution Distribution
    currentRow = createResolutionDistribution_(analyticsSheet, currentRow, startDate, endDate);
    currentRow += 5;
    
    // 7. Real-time Capacity Indicators
    currentRow = createCapacityIndicators_(analyticsSheet, currentRow);
    currentRow += 5;
    
    // 8. Predictive Analytics
    currentRow = createPredictiveAnalytics_(analyticsSheet, currentRow);
    
  } catch (e) {
    Logger.log('refreshAdvancedAnalyticsDashboard_ error: ' + e.toString());
  }
}

// 1. Peak Hours & Day of Week Analysis
function createPeakHoursAnalysis_(sheet, startRow, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startIdx = headers.indexOf('start_time');
    
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    
    // Analyze by hour (0-23)
    const hourCounts = Array(24).fill(0);
    const dayCounts = [0, 0, 0, 0, 0, 0, 0]; // Sun-Sat
    
    filtered.forEach(row => {
      if (row[startIdx]) {
        const date = new Date(row[startIdx]);
        const hour = date.getHours();
        const day = date.getDay();
        hourCounts[hour]++;
        dayCounts[day]++;
      }
    });
    
    // Title
    sheet.getRange(startRow, 1).setValue('📊 Peak Hours & Day of Week Analysis');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    // Peak Hours Table
    const hourRow = startRow + 2;
    sheet.getRange(hourRow, 1).setValue('Peak Hours (24-hour format)');
    sheet.getRange(hourRow, 1).setFontSize(14).setFontWeight('bold');
    sheet.getRange(hourRow + 1, 1).setValue('Hour');
    sheet.getRange(hourRow + 1, 2).setValue('Sessions');
    sheet.getRange(hourRow + 1, 3).setValue('Percentage');
    sheet.getRange(hourRow + 1, 1, 1, 3).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const totalSessions = filtered.length;
    const hourData = [];
    for (let i = 0; i < 24; i++) {
      const pct = totalSessions > 0 ? ((hourCounts[i] / totalSessions) * 100).toFixed(1) : '0';
      hourData.push([i + ':00', hourCounts[i], pct + '%']);
    }
    
    // Sort by sessions descending
    hourData.sort((a, b) => b[1] - a[1]);
    sheet.getRange(hourRow + 2, 1, Math.min(24, hourData.length), 3).setValues(hourData.slice(0, 24));
    
    // Day of Week Table
    const dayRow = hourRow + 27;
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    sheet.getRange(dayRow, 1).setValue('Day of Week Analysis');
    sheet.getRange(dayRow, 1).setFontSize(14).setFontWeight('bold');
    sheet.getRange(dayRow + 1, 1).setValue('Day');
    sheet.getRange(dayRow + 1, 2).setValue('Sessions');
    sheet.getRange(dayRow + 1, 3).setValue('Percentage');
    sheet.getRange(dayRow + 1, 1, 1, 3).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const dayData = dayNames.map((name, idx) => {
      const pct = totalSessions > 0 ? ((dayCounts[idx] / totalSessions) * 100).toFixed(1) : '0';
      return [name, dayCounts[idx], pct + '%'];
    });
    
    // Sort by sessions descending
    dayData.sort((a, b) => b[1] - a[1]);
    sheet.getRange(dayRow + 2, 1, dayData.length, 3).setValues(dayData);
    
    // Peak hour summary
    const peakHour = hourCounts.indexOf(Math.max(...hourCounts));
    const peakDay = dayCounts.indexOf(Math.max(...dayCounts));
    const summaryRow = dayRow + dayData.length + 3;
    sheet.getRange(summaryRow, 1).setValue('Summary:');
    sheet.getRange(summaryRow, 1).setFontWeight('bold');
    sheet.getRange(summaryRow, 2).setValue(`Peak Hour: ${peakHour}:00 (${hourCounts[peakHour]} sessions)`);
    sheet.getRange(summaryRow, 3).setValue(`Peak Day: ${dayNames[peakDay]} (${dayCounts[peakDay]} sessions)`);
    
    return summaryRow + 2;
  } catch (e) {
    Logger.log('createPeakHoursAnalysis_ error: ' + e.toString());
    return startRow + 50;
  }
}

// 2. Technician Effectiveness Comparison (Composite Score)
function createTechnicianEffectiveness_(sheet, startRow, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startIdx = headers.indexOf('start_time');
    const techIdx = headers.indexOf('technician_name');
    const durationIdx = headers.indexOf('duration_total_seconds');
    const pickupIdx = headers.indexOf('pickup_seconds');
    const workIdx = headers.indexOf('duration_work_seconds');
    const resolvedIdx = headers.indexOf('resolved_unresolved');
    
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    
    // Calculate metrics per technician
    const techMetrics = {};
    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      if (!techMetrics[tech]) {
        techMetrics[tech] = {
          sessions: 0,
          durations: [],
          pickups: [],
          workSeconds: 0,
          resolved: 0,
          slaHits: 0
        };
      }
      techMetrics[tech].sessions++;
      if (row[durationIdx]) techMetrics[tech].durations.push(Number(row[durationIdx]));
      if (row[pickupIdx]) {
        techMetrics[tech].pickups.push(Number(row[pickupIdx]));
        if (Number(row[pickupIdx]) <= 30) techMetrics[tech].slaHits++;
      }
      if (row[workIdx]) techMetrics[tech].workSeconds += Number(row[workIdx]);
      if (row[resolvedIdx] === 'Resolved') techMetrics[tech].resolved++;
    });
    
    // Calculate composite effectiveness score (0-100)
    // Components: Sessions/Hour (30%), SLA% (30%), Resolution Rate (20%), Avg Duration efficiency (20%)
    const techScores = Object.keys(techMetrics).map(tech => {
      const m = techMetrics[tech];
      const days = Math.max(1, (new Date(endDate) - new Date(startDate)) / (1000*60*60*24));
      const sessionsPerHour = (m.sessions / days / 8);
      const avgPickup = m.pickups.length > 0 ? m.pickups.reduce((a, b) => a + b, 0) / m.pickups.length : 0;
      const slaPct = m.pickups.length > 0 ? (m.slaHits / m.pickups.length) * 100 : 0;
      const resolutionRate = m.sessions > 0 ? (m.resolved / m.sessions) * 100 : 0;
      const avgDuration = m.durations.length > 0 ? m.durations.reduce((a, b) => a + b, 0) / m.durations.length : 0;
      
      // Normalize scores (higher is better for all)
      // Sessions/Hour: normalize to 0-100 (assuming 5 sessions/hour is max)
      const sessionsScore = Math.min(100, (sessionsPerHour / 5) * 100);
      
      // SLA%: already 0-100
      const slaScore = slaPct;
      
      // Resolution Rate: already 0-100
      const resolutionScore = resolutionRate;
      
      // Duration: inverse (lower duration = better), normalize assuming 30min avg
      const durationScore = Math.max(0, 100 - ((avgDuration / 60 - 30) / 30) * 100);
      
      // Weighted composite score
      const compositeScore = (sessionsScore * 0.30) + (slaScore * 0.30) + (resolutionScore * 0.20) + (durationScore * 0.20);
      
      return {
        tech,
        sessions: m.sessions,
        sessionsPerHour: sessionsPerHour.toFixed(1),
        slaPct: slaPct.toFixed(1),
        resolutionRate: resolutionRate.toFixed(1),
        avgDuration: (avgDuration / 60).toFixed(1),
        compositeScore: compositeScore.toFixed(1)
      };
    });
    
    // Sort by composite score descending
    techScores.sort((a, b) => parseFloat(b.compositeScore) - parseFloat(a.compositeScore));
    
    // Title
    sheet.getRange(startRow, 1).setValue('⭐ Technician Effectiveness Comparison');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    const tableRow = startRow + 2;
    const tableHeaders = ['Technician', 'Sessions', 'Sessions/Hr', 'SLA %', 'Resolution %', 'Avg Duration (min)', 'Composite Score'];
    sheet.getRange(tableRow, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(tableRow, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const tableData = techScores.map(t => [
      t.tech,
      t.sessions,
      t.sessionsPerHour,
      t.slaPct + '%',
      t.resolutionRate + '%',
      t.avgDuration,
      t.compositeScore
    ]);
    
    if (tableData.length > 0) {
      sheet.getRange(tableRow + 1, 1, tableData.length, tableHeaders.length).setValues(tableData);
      
      // Color code composite scores
      const scoreCol = tableHeaders.length;
      for (let i = 0; i < tableData.length; i++) {
        const score = parseFloat(tableData[i][scoreCol - 1]);
        const range = sheet.getRange(tableRow + 1 + i, scoreCol);
        if (score >= 80) {
          range.setBackground('#34A853').setFontColor('#FFFFFF');
        } else if (score >= 60) {
          range.setBackground('#EA8600').setFontColor('#FFFFFF');
        } else {
          range.setBackground('#EA4335').setFontColor('#FFFFFF');
        }
      }
    }
    
    return tableRow + tableData.length + 2;
  } catch (e) {
    Logger.log('createTechnicianEffectiveness_ error: ' + e.toString());
    return startRow + 50;
  }
}

// 3. Repeat Customer Analysis
function createRepeatCustomerAnalysis_(sheet, startRow, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startIdx = headers.indexOf('start_time');
    const customerIdx = headers.indexOf('customer_name');
    const customerEmailIdx = headers.indexOf('customer_email');
    const trackingIdIdx = headers.indexOf('tracking_id');
    const sessionIdIdx = headers.indexOf('session_id');
    
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    
    // Track customers by multiple identifiers
    const customerSessions = {};
    filtered.forEach(row => {
      // Use customer name, email, or tracking ID as identifier
      const identifier = row[customerEmailIdx] || row[customerIdx] || row[trackingIdIdx] || 'Unknown';
      const sessionId = row[sessionIdIdx];
      
      if (!customerSessions[identifier]) {
        customerSessions[identifier] = {
          name: row[customerIdx] || 'Unknown',
          email: row[customerEmailIdx] || '',
          sessions: [],
          count: 0
        };
      }
      customerSessions[identifier].sessions.push(sessionId);
      customerSessions[identifier].count++;
    });
    
    // Find repeat customers (2+ sessions)
    const repeatCustomers = Object.keys(customerSessions)
      .filter(id => customerSessions[id].count >= 2)
      .map(id => ({
        identifier: id,
        name: customerSessions[id].name,
        email: customerSessions[id].email,
        sessionCount: customerSessions[id].count,
        sessions: customerSessions[id].sessions
      }))
      .sort((a, b) => b.sessionCount - a.sessionCount);
    
    // Title
    sheet.getRange(startRow, 1).setValue('🔄 Repeat Customer Analysis');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    const summaryRow = startRow + 2;
    sheet.getRange(summaryRow, 1).setValue('Summary:');
    sheet.getRange(summaryRow, 1).setFontWeight('bold');
    sheet.getRange(summaryRow, 2).setValue(`Total Customers: ${Object.keys(customerSessions).length}`);
    sheet.getRange(summaryRow, 3).setValue(`Repeat Customers (2+): ${repeatCustomers.length}`);
    const repeatPct = Object.keys(customerSessions).length > 0 ? 
      ((repeatCustomers.length / Object.keys(customerSessions).length) * 100).toFixed(1) : '0';
    sheet.getRange(summaryRow, 4).setValue(`Repeat Rate: ${repeatPct}%`);
    
    // Repeat customers table
    const tableRow = summaryRow + 2;
    sheet.getRange(tableRow, 1).setValue('Top Repeat Customers');
    sheet.getRange(tableRow, 1).setFontSize(14).setFontWeight('bold');
    const tableHeaders = ['Customer Name', 'Email', 'Session Count', 'Session IDs'];
    sheet.getRange(tableRow + 1, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(tableRow + 1, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const tableData = repeatCustomers.slice(0, 20).map(c => [
      c.name,
      c.email || '—',
      c.sessionCount,
      c.sessions.slice(0, 5).join(', ') + (c.sessions.length > 5 ? '...' : '')
    ]);
    
    if (tableData.length > 0) {
      sheet.getRange(tableRow + 2, 1, tableData.length, tableHeaders.length).setValues(tableData);
    } else {
      sheet.getRange(tableRow + 2, 1).setValue('No repeat customers found');
      sheet.getRange(tableRow + 2, 1).setFontStyle('italic').setFontColor('#999999');
    }
    
    return tableRow + Math.max(tableData.length, 1) + 3;
  } catch (e) {
    Logger.log('createRepeatCustomerAnalysis_ error: ' + e.toString());
    return startRow + 50;
  }
}

// 4. Trend Analysis (Week-over-Week, Month-over-Month)
function createTrendAnalysis_(sheet, startRow) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const configSheet = ss.getSheetByName('Dashboard_Config');
    const timeFrame = configSheet ? (configSheet.getRange('B3').getValue() || 'Today') : 'Today';
    
    // Title
    sheet.getRange(startRow, 1).setValue('📈 Trend Analysis');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    const infoRow = startRow + 2;
    sheet.getRange(infoRow, 1).setValue('Trend comparisons require historical data.');
    sheet.getRange(infoRow, 2).setValue('Pull data for previous periods to enable WoW/MoM analysis.');
    sheet.getRange(infoRow, 1, 1, 2).setFontStyle('italic').setFontColor('#666666');
    
    // Calculate current period metrics
    const currentRange = getTimeFrameRange_(timeFrame);
    const currentMetrics = calculatePeriodMetrics_(currentRange.startDate, currentRange.endDate);
    
    // Calculate previous period metrics
    let previousRange;
    let previousMetrics = null;
    
    if (timeFrame === 'Today') {
      previousRange = getTimeFrameRange_('Yesterday');
      previousMetrics = calculatePeriodMetrics_(previousRange.startDate, previousRange.endDate);
    } else if (timeFrame === 'This Week') {
      previousRange = getTimeFrameRange_('Last Week');
      previousMetrics = calculatePeriodMetrics_(previousRange.startDate, previousRange.endDate);
    } else if (timeFrame === 'This Month') {
      previousRange = getTimeFrameRange_('Last Month');
      previousMetrics = calculatePeriodMetrics_(previousRange.startDate, previousRange.endDate);
    }
    
    const tableRow = infoRow + 3;
    sheet.getRange(tableRow, 1).setValue('Period Comparison');
    sheet.getRange(tableRow, 1).setFontSize(14).setFontWeight('bold');
    
    const headers = ['Metric', 'Current Period', 'Previous Period', 'Change', 'Change %'];
    sheet.getRange(tableRow + 1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(tableRow + 1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    if (previousMetrics) {
      const metrics = [
        ['Total Sessions', currentMetrics.sessions, previousMetrics.sessions],
        ['Avg Duration (min)', currentMetrics.avgDuration, previousMetrics.avgDuration],
        ['Avg Pickup (sec)', currentMetrics.avgPickup, previousMetrics.avgPickup],
        ['SLA Hit %', currentMetrics.slaPct, previousMetrics.slaPct],
        ['Resolution Rate', currentMetrics.resolutionRate, previousMetrics.resolutionRate]
      ];
      
      const tableData = metrics.map(([metric, current, previous]) => {
        const change = current - previous;
        const changePct = previous !== 0 ? ((change / previous) * 100).toFixed(1) : '0';
        return [metric, current, previous, change, changePct + '%'];
      });
      
      sheet.getRange(tableRow + 2, 1, tableData.length, headers.length).setValues(tableData);
      
      // Color code changes
      for (let i = 0; i < tableData.length; i++) {
        const change = tableData[i][3];
        const changeRange = sheet.getRange(tableRow + 2 + i, 4);
        if (change > 0 && (i === 0 || i === 4)) { // Sessions or Resolution Rate - positive is good
          changeRange.setFontColor('#34A853').setFontWeight('bold');
        } else if (change < 0 && (i === 0 || i === 4)) {
          changeRange.setFontColor('#EA4335').setFontWeight('bold');
        } else if (change < 0 && (i === 1 || i === 2)) { // Duration or Pickup - negative is good
          changeRange.setFontColor('#34A853').setFontWeight('bold');
        } else if (change > 0 && (i === 1 || i === 2)) {
          changeRange.setFontColor('#EA4335').setFontWeight('bold');
        }
      }
      
      return tableRow + tableData.length + 2;
    } else {
      sheet.getRange(tableRow + 2, 1).setValue('No previous period data available for comparison');
      return tableRow + 3;
    }
  } catch (e) {
    Logger.log('createTrendAnalysis_ error: ' + e.toString());
    return startRow + 50;
  }
}

// Helper function to calculate period metrics
function calculatePeriodMetrics_(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return { sessions: 0, avgDuration: 0, avgPickup: 0, slaPct: 0, resolutionRate: 0 };
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return { sessions: 0, avgDuration: 0, avgPickup: 0, slaPct: 0, resolutionRate: 0 };
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startIdx = headers.indexOf('start_time');
    const durationIdx = headers.indexOf('duration_total_seconds');
    const pickupIdx = headers.indexOf('pickup_seconds');
    const resolvedIdx = headers.indexOf('resolved_unresolved');
    
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    
    const durations = filtered.map(r => Number(r[durationIdx] || 0)).filter(d => d > 0);
    const pickups = filtered.map(r => Number(r[pickupIdx] || 0)).filter(p => p > 0);
    const slaHits = pickups.filter(p => p <= 30).length;
    const resolved = filtered.filter(r => r[resolvedIdx] === 'Resolved').length;
    
    return {
      sessions: filtered.length,
      avgDuration: durations.length > 0 ? (durations.reduce((a, b) => a + b, 0) / durations.length / 60).toFixed(1) : '0',
      avgPickup: pickups.length > 0 ? Math.round(pickups.reduce((a, b) => a + b, 0) / pickups.length) : '0',
      slaPct: pickups.length > 0 ? ((slaHits / pickups.length) * 100).toFixed(1) : '0',
      resolutionRate: filtered.length > 0 ? ((resolved / filtered.length) * 100).toFixed(1) : '0'
    };
  } catch (e) {
    Logger.log('calculatePeriodMetrics_ error: ' + e.toString());
    return { sessions: 0, avgDuration: 0, avgPickup: 0, slaPct: 0, resolutionRate: 0 };
  }
}

// 5. Technician Utilization Rate
function createUtilizationRate_(sheet, startRow, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startIdx = headers.indexOf('start_time');
    const techIdx = headers.indexOf('technician_name');
    const workIdx = headers.indexOf('duration_work_seconds');
    
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    
    // Calculate utilization per technician
    const techUtilization = {};
    const days = Math.max(1, (new Date(endDate) - new Date(startDate)) / (1000*60*60*24));
    const hoursPerDay = 8; // Assuming 8-hour workday
    const totalAvailableHours = days * hoursPerDay;
    
    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      if (!techUtilization[tech]) {
        techUtilization[tech] = {
          workSeconds: 0,
          sessions: 0
        };
      }
      techUtilization[tech].workSeconds += Number(row[workIdx] || 0);
      techUtilization[tech].sessions++;
    });
    
    const utilizationData = Object.keys(techUtilization).map(tech => {
      const util = techUtilization[tech];
      const workHours = util.workSeconds / 3600;
      const utilizationRate = totalAvailableHours > 0 ? (workHours / totalAvailableHours * 100) : 0;
      
      return {
        tech,
        workHours: workHours.toFixed(1),
        availableHours: totalAvailableHours.toFixed(1),
        utilizationRate: utilizationRate.toFixed(1),
        sessions: util.sessions
      };
    }).sort((a, b) => parseFloat(b.utilizationRate) - parseFloat(a.utilizationRate));
    
    // Title
    sheet.getRange(startRow, 1).setValue('⏱️ Technician Utilization Rate');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    const infoRow = startRow + 2;
    sheet.getRange(infoRow, 1).setValue(`Period: ${days.toFixed(1)} days × ${hoursPerDay} hours/day = ${totalAvailableHours.toFixed(1)} available hours per technician`);
    sheet.getRange(infoRow, 1).setFontStyle('italic').setFontColor('#666666');
    
    const tableRow = infoRow + 2;
    const tableHeaders = ['Technician', 'Work Hours', 'Available Hours', 'Utilization %', 'Sessions'];
    sheet.getRange(tableRow, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(tableRow, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const tableData = utilizationData.map(t => [
      t.tech,
      t.workHours,
      t.availableHours,
      t.utilizationRate + '%',
      t.sessions
    ]);
    
    if (tableData.length > 0) {
      sheet.getRange(tableRow + 1, 1, tableData.length, tableHeaders.length).setValues(tableData);
      
      // Color code utilization rates
      const utilCol = tableHeaders.indexOf('Utilization %') + 1;
      for (let i = 0; i < tableData.length; i++) {
        const rate = parseFloat(tableData[i][utilCol - 1]);
        const range = sheet.getRange(tableRow + 1 + i, utilCol);
        if (rate >= 80) {
          range.setBackground('#34A853').setFontColor('#FFFFFF');
        } else if (rate >= 60) {
          range.setBackground('#EA8600').setFontColor('#FFFFFF');
        } else {
          range.setBackground('#EA4335').setFontColor('#FFFFFF');
        }
      }
    }
    
    return tableRow + tableData.length + 2;
  } catch (e) {
    Logger.log('createUtilizationRate_ error: ' + e.toString());
    return startRow + 50;
  }
}

// 6. Time to Resolution Distribution (Percentiles)
function createResolutionDistribution_(sheet, startRow, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startIdx = headers.indexOf('start_time');
    const durationIdx = headers.indexOf('duration_total_seconds');
    
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
    
    const durations = filtered
      .map(r => Number(r[durationIdx] || 0))
      .filter(d => d > 0)
      .sort((a, b) => a - b);
    
    if (durations.length === 0) {
      sheet.getRange(startRow, 1).setValue('⏱️ Time to Resolution Distribution');
      sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold');
      sheet.getRange(startRow + 2, 1).setValue('No duration data available');
      return startRow + 5;
    }
    
    // Calculate percentiles
    const percentile = (arr, p) => {
      if (arr.length === 0) return 0;
      const index = Math.ceil((p / 100) * arr.length) - 1;
      return arr[Math.max(0, Math.min(index, arr.length - 1))];
    };
    
    const p25 = percentile(durations, 25);
    const p50 = percentile(durations, 50);
    const p75 = percentile(durations, 75);
    const p90 = percentile(durations, 90);
    const p95 = percentile(durations, 95);
    const p99 = percentile(durations, 99);
    const min = durations[0];
    const max = durations[durations.length - 1];
    const avg = durations.reduce((a, b) => a + b, 0) / durations.length;
    
    const formatTime = (seconds) => {
      const hours = Math.floor(seconds / 3600);
      const minutes = Math.floor((seconds % 3600) / 60);
      const secs = Math.floor(seconds % 60);
      if (hours > 0) return `${hours}h ${minutes}m`;
      if (minutes > 0) return `${minutes}m ${secs}s`;
      return `${secs}s`;
    };
    
    // Title
    sheet.getRange(startRow, 1).setValue('⏱️ Time to Resolution Distribution');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    const tableRow = startRow + 2;
    const tableHeaders = ['Metric', 'Time', 'Minutes'];
    sheet.getRange(tableRow, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    sheet.getRange(tableRow, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const tableData = [
      ['Minimum', formatTime(min), (min / 60).toFixed(1)],
      ['25th Percentile (P25)', formatTime(p25), (p25 / 60).toFixed(1)],
      ['50th Percentile (Median)', formatTime(p50), (p50 / 60).toFixed(1)],
      ['75th Percentile (P75)', formatTime(p75), (p75 / 60).toFixed(1)],
      ['90th Percentile (P90)', formatTime(p90), (p90 / 60).toFixed(1)],
      ['95th Percentile (P95)', formatTime(p95), (p95 / 60).toFixed(1)],
      ['99th Percentile (P99)', formatTime(p99), (p99 / 60).toFixed(1)],
      ['Maximum', formatTime(max), (max / 60).toFixed(1)],
      ['Average', formatTime(avg), (avg / 60).toFixed(1)]
    ];
    
    sheet.getRange(tableRow + 1, 1, tableData.length, tableHeaders.length).setValues(tableData);
    
    return tableRow + tableData.length + 2;
  } catch (e) {
    Logger.log('createResolutionDistribution_ error: ' + e.toString());
    return startRow + 50;
  }
}

// 7. Real-time Capacity Indicators
function createCapacityIndicators_(sheet, startRow) {
  try {
    const ss = SpreadsheetApp.getActive();
    const cfg = getCfg_();
    
    // Title
    sheet.getRange(startRow, 1).setValue('📊 Real-time Capacity Indicators');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    // Get current sessions from API
    const currentSessionData = fetchCurrentSessions_(cfg);
    const activeCount = currentSessionData.active.length;
    const waitingCount = currentSessionData.waiting.length;
    
    // Calculate average resolution time from recent data
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    let avgResolutionMinutes = 0;
    if (sessionsSheet) {
      const dataRange = sessionsSheet.getDataRange();
      if (dataRange.getNumRows() > 1) {
        const allData = sessionsSheet.getRange(2, 1, Math.min(100, dataRange.getNumRows() - 1), dataRange.getNumColumns()).getValues();
        const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
        const durationIdx = headers.indexOf('duration_total_seconds');
        const durations = allData
          .map(r => Number(r[durationIdx] || 0))
          .filter(d => d > 0);
        if (durations.length > 0) {
          avgResolutionMinutes = (durations.reduce((a, b) => a + b, 0) / durations.length) / 60;
        }
      }
    }
    
    // Estimate capacity
    const loggedInTechs = fetchLoggedInTechnicians_(cfg);
    const availableTechs = loggedInTechs.length;
    
    // Calculate capacity metrics
    const tableRow = startRow + 2;
    sheet.getRange(tableRow, 1).setValue('Current Status');
    sheet.getRange(tableRow, 1).setFontSize(14).setFontWeight('bold');
    
    const metrics = [
      ['Active Sessions', activeCount],
      ['Waiting in Queue', waitingCount],
      ['Available Technicians', availableTechs],
      ['Avg Resolution Time', avgResolutionMinutes > 0 ? avgResolutionMinutes.toFixed(1) + ' min' : 'N/A'],
      ['Est. Capacity (sessions/hour)', availableTechs > 0 && avgResolutionMinutes > 0 ? 
        Math.floor((availableTechs * 60) / avgResolutionMinutes).toString() : 'N/A'],
      ['Queue Wait Time (est.)', waitingCount > 0 && availableTechs > 0 && avgResolutionMinutes > 0 ?
        ((waitingCount * avgResolutionMinutes) / availableTechs).toFixed(1) + ' min' : '0 min']
    ];
    
    const headers = ['Metric', 'Value'];
    sheet.getRange(tableRow + 1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(tableRow + 1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const tableData = metrics.map(([metric, value]) => [metric, value]);
    sheet.getRange(tableRow + 2, 1, tableData.length, headers.length).setValues(tableData);
    
    // Color code queue wait time
    if (waitingCount > 0) {
      const waitTime = parseFloat(metrics[5][1]) || 0;
      const waitRange = sheet.getRange(tableRow + 2 + 5, 2);
      if (waitTime > 15) {
        waitRange.setBackground('#EA4335').setFontColor('#FFFFFF').setFontWeight('bold');
      } else if (waitTime > 5) {
        waitRange.setBackground('#EA8600').setFontColor('#FFFFFF');
      } else {
        waitRange.setBackground('#34A853').setFontColor('#FFFFFF');
      }
    }
    
    return tableRow + tableData.length + 2;
  } catch (e) {
    Logger.log('createCapacityIndicators_ error: ' + e.toString());
    return startRow + 50;
  }
}

// 8. Predictive Analytics
function createPredictiveAnalytics_(sheet, startRow) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allData = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const startIdx = headers.indexOf('start_time');
    
    // Analyze historical patterns (last 30 days of data)
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    
    const historicalData = allData.filter(row => {
      if (!row[startIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]);
        return rowDate >= thirtyDaysAgo;
      } catch (e) {
        return false;
      }
    });
    
    if (historicalData.length < 10) {
      sheet.getRange(startRow, 1).setValue('🔮 Predictive Analytics');
      sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold');
      sheet.getRange(startRow + 2, 1).setValue('Insufficient historical data (need 30+ days) for predictions');
      return startRow + 5;
    }
    
    // Analyze patterns by hour and day
    const hourPatterns = Array(24).fill(0).map(() => []);
    const dayPatterns = Array(7).fill(0).map(() => []);
    
    historicalData.forEach(row => {
      if (row[startIdx]) {
        const date = new Date(row[startIdx]);
        const hour = date.getHours();
        const day = date.getDay();
        hourPatterns[hour].push(date);
        dayPatterns[day].push(date);
      }
    });
    
    // Calculate average sessions per hour/day
    const avgSessionsPerHour = hourPatterns.map((sessions, hour) => ({
      hour,
      avg: sessions.length / 30 // Average over 30 days
    }));
    
    const avgSessionsPerDay = dayPatterns.map((sessions, day) => ({
      day,
      avg: sessions.length / (30 / 7) // Average per day type
    }));
    
    // Find peak hours
    avgSessionsPerHour.sort((a, b) => b.avg - a.avg);
    const topHours = avgSessionsPerHour.slice(0, 3);
    
    // Find peak days
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    avgSessionsPerDay.sort((a, b) => b.avg - a.avg);
    const topDays = avgSessionsPerDay.slice(0, 3);
    
    // Title
    sheet.getRange(startRow, 1).setValue('🔮 Predictive Analytics');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    sheet.getRange(startRow, 1, 1, 3).merge();
    
    const infoRow = startRow + 2;
    sheet.getRange(infoRow, 1).setValue('Based on last 30 days of historical data');
    sheet.getRange(infoRow, 1).setFontStyle('italic').setFontColor('#666666');
    
    const forecastRow = infoRow + 2;
    sheet.getRange(forecastRow, 1).setValue('Forecasted Peak Periods');
    sheet.getRange(forecastRow, 1).setFontSize(14).setFontWeight('bold');
    
    const forecastHeaders = ['Type', 'Period', 'Expected Sessions/Day'];
    sheet.getRange(forecastRow + 1, 1, 1, forecastHeaders.length).setValues([forecastHeaders]);
    sheet.getRange(forecastRow + 1, 1, 1, forecastHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
    
    const forecastData = [
      ...topHours.map(h => ['Peak Hour', h.hour + ':00', h.avg.toFixed(1)]),
      ...topDays.map(d => ['Peak Day', dayNames[d.day], d.avg.toFixed(1)])
    ];
    
    if (forecastData.length > 0) {
      sheet.getRange(forecastRow + 2, 1, forecastData.length, forecastHeaders.length).setValues(forecastData);
    }
    
    // Recommendations
    const recRow = forecastRow + forecastData.length + 3;
    sheet.getRange(recRow, 1).setValue('Recommendations:');
    sheet.getRange(recRow, 1).setFontWeight('bold');
    const recommendations = [
      `Schedule extra staff during peak hours: ${topHours.map(h => h.hour + ':00').join(', ')}`,
      `Prepare for higher volume on: ${topDays.map(d => dayNames[d.day]).join(', ')}`,
      `Average daily session volume: ${(historicalData.length / 30).toFixed(1)} sessions/day`
    ];
    
    recommendations.forEach((rec, idx) => {
      sheet.getRange(recRow + 1 + idx, 1).setValue('• ' + rec);
    });
    
    return recRow + recommendations.length + 2;
  } catch (e) {
    Logger.log('createPredictiveAnalytics_ error: ' + e.toString());
    return startRow + 50;
  }
}
