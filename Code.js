// Cache Performance SUMMARY results for a short period to avoid redundant API calls
function getPerfSummaryCached_(cfg, startDate, endDate) {
  try {
    const cache = CacheService.getScriptCache();
    const key = `perf:${isoDate_(startDate)}:${isoDate_(endDate)}`;
    const cached = cache.get(key);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }
    const map = fetchPerformanceSummaryData_(cfg, startDate, endDate) || {};
    try { cache.put(key, JSON.stringify(map), 300); } catch (e) {}
    return map;
  } catch (e) {
    Logger.log('getPerfSummaryCached_ error: ' + e.toString());
    return fetchPerformanceSummaryData_(cfg, startDate, endDate) || {};
  }
}
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
// When false, we request XML from Rescue where supported and parse it to avoid column misalignment.
// We still fall back to TEXT automatically if XML is not available.
const FORCE_TEXT_OUTPUT = false;
const SHEETS_SESSIONS_TABLE = 'Sessions'; // Main data storage sheet
// If true, write raw API values into the Sessions sheet (preserve empty strings
// and original formatting) instead of converting timestamps/durations to Dates/numbers.
const STORE_RAW_SESSIONS = true;

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
    .addItem('🧭 Migrate Sessions Headers (Official)', 'migrateSessionsHeaders')
  .addItem('🧾 Dump Performance Raw', 'dumpPerformanceRawMenu')
  .addItem('🗺️ Dump Digium Map (raw)', 'dumpDigiumMapMenu')
  .addItem('🗺️ Dump Rescue LISTALL (raw)', 'dumpRescueListAllRawMenu')
  .addItem('⬅️ Shift Rescue_Map data (Session ID → Technician ID)', 'shiftRescueMapDataAlignMenu')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('SUMMARY Mode')
        .addItem('Use CHANNEL only (faster)', 'setSummaryModeChannelOnly')
        .addItem('Use BOTH: NODE + CHANNEL', 'setSummaryModeBoth')
    )
    .addSeparator()
    .addItem('📈 Advanced Analytics Dashboard', 'createAdvancedAnalyticsDashboard')
    .addToUi();
}

/* ===== Secrets / Config ===== */
const PROP_KEYS = {
  RESCUE_BASE:  'RESCUE_BASE',
  RESCUE_USER:  'RESCUE_USER',
  RESCUE_PASS:  'RESCUE_PASS',
  DIGIUM_HOST:   'DIGIUM_HOST',
  DIGIUM_USER:   'DIGIUM_USER',
  DIGIUM_PASS:   'DIGIUM_PASS',
  NODE_JSON:    'NODE_CANDIDATES_JSON',
  SUMMARY_MODE: 'SUMMARY_NODEREF_MODE' // 'CHANNEL_ONLY' (default) or 'BOTH'
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
    digiumHost: getProp_(PROP_KEYS.DIGIUM_HOST, 'https://nova.digiumcloud.net/xml'),
    digiumUser: getProp_(PROP_KEYS.DIGIUM_USER, ''),
    digiumPass: getProp_(PROP_KEYS.DIGIUM_PASS, ''),
    nodes: (() => {
      const j = getProp_(PROP_KEYS.NODE_JSON, '[]');
      try { return JSON.parse(j); } catch(e) { return NODE_CANDIDATES_DEFAULT; }
    })()
  };
}

// Determine which noderefs to use for SUMMARY calls.
// Returns ['CHANNEL'] by default for efficiency; if Script Property SUMMARY_NODEREF_MODE is 'BOTH', returns ['NODE','CHANNEL'].
function getSummaryNoderefs_() {
  try {
    const mode = getProp_(PROP_KEYS.SUMMARY_MODE, 'CHANNEL_ONLY');
    if (String(mode).toUpperCase() === 'BOTH') return ['NODE','CHANNEL'];
    return ['CHANNEL'];
  } catch (e) {
    return ['CHANNEL'];
  }
}

// Quick UI toggles for SUMMARY noderef behavior
function setSummaryModeChannelOnly() {
  try {
    setProp_(PROP_KEYS.SUMMARY_MODE, 'CHANNEL_ONLY');
    SpreadsheetApp.getActive().toast('SUMMARY mode set to CHANNEL_ONLY');
  } catch (e) {
    Logger.log('setSummaryModeChannelOnly error: ' + e.toString());
  }
}

function setSummaryModeBoth() {
  try {
    setProp_(PROP_KEYS.SUMMARY_MODE, 'BOTH');
    SpreadsheetApp.getActive().toast('SUMMARY mode set to BOTH (NODE + CHANNEL)');
  } catch (e) {
    Logger.log('setSummaryModeBoth error: ' + e.toString());
  }
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
    <hr>
    <h3>Digium / Switchvox (optional)</h3>
    <p>Digium Host (XML-RPC endpoint):<br><input type="text" id="dig_host" placeholder="https://nova.digiumcloud.net/xml"></p>
    <p>Digium Username:<br><input type="text" id="dig_user"></p>
    <p>Digium Password:<br><input type="password" id="dig_pass"></p>
    <button onclick="save()">Save</button>
    <button onclick="google.script.host.close()">Cancel</button>
    <script>
      function save() {
        const base = document.getElementById('base').value;
        const user = document.getElementById('user').value;
        const pass = document.getElementById('pass').value;
        const nodes = document.getElementById('nodes').value;
        const digHost = document.getElementById('dig_host').value;
        const digUser = document.getElementById('dig_user').value;
        const digPass = document.getElementById('dig_pass').value;
        google.script.run.saveSecrets(base, user, pass, nodes);
        if (digHost || digUser || digPass) google.script.run.saveDigiumSecrets(digHost, digUser, digPass);
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

function saveDigiumSecrets(host, user, pass) {
  try {
    if (host) setProp_('DIGIUM_HOST', host);
    if (user) setProp_('DIGIUM_USER', user);
    if (pass) setProp_('DIGIUM_PASS', pass);
    SpreadsheetApp.getActive().toast('Digium secrets saved');
  } catch (e) { Logger.log('saveDigiumSecrets error: ' + e.toString()); }
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

/* ===== Digium / Switchvox XML-RPC helpers ===== */
// Helper: escape text for XML content
function xmlEscape_(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Generic XML-RPC caller. Uses HTTP Basic Auth against the provided host.
function xmlRpcCall_(host, methodName, params, user, pass) {
  // Build a minimal XML-RPC request (params should be an array of simple values or objects already serialized)
  let body = '<?xml version="1.0"?>\n<methodCall><methodName>' + methodName + '</methodName><params>';
  (params || []).forEach(p => {
    body += '<param><value>';
    if (typeof p === 'number') body += '<int>' + String(p) + '</int>';
    else if (typeof p === 'boolean') body += '<boolean>' + (p ? '1' : '0') + '</boolean>';
  else body += '<string>' + xmlEscape_(String(p)) + '</string>';
    body += '</value></param>';
  });
  body += '</params></methodCall>';

  // Use application/xml per Digium error response which explicitly requires application/xml or application/json
  const headers = { 'Content-Type': 'application/xml; charset=UTF-8', 'Accept': 'application/xml' };
  if (user && pass) headers['Authorization'] = 'Basic ' + Utilities.base64Encode(user + ':' + pass);
  const opts = { method: 'post', contentType: 'application/xml; charset=UTF-8', payload: body, headers: headers, muteHttpExceptions: true };
  try {
    const resp = UrlFetchApp.fetch(host, opts);
    const code = resp.getResponseCode();
    const txt = resp.getContentText();
    if (code >= 200 && code < 300) {
      try {
        const xml = XmlService.parse(txt);
        return { ok: true, xml: xml, raw: txt };
      } catch (e) {
        return { ok: false, error: 'XML parse error: ' + e.toString(), raw: txt };
      }
    }
    return { ok: false, error: 'HTTP ' + code, raw: txt };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// Fetch Digium call summary for a date range. This is a small wrapper and uses placeholder method names.
// The Switchvox API uses XML-RPC; exact method names and params must be taken from the Digium docs.
function fetchDigiumCallSummary_(startDate, endDate) {
  const cfg = getCfg_();
  const host = cfg.digiumHost;
  const user = cfg.digiumUser;
  const pass = cfg.digiumPass;
  if (!host || !user || !pass) {
    Logger.log('fetchDigiumCallSummary_: Digium credentials not configured');
    return { ok: false, reason: 'missing_credentials' };
  }

  // NOTE: The method name and param structure below are placeholders. Refer to the Switchvox API docs
  // (http://developers.digium.com/switchvox) for the correct method to request call logs or CDRs.
  const method = 'switchvox.calls.getList'; // placeholder — replace with real method
  const params = [ { start: startDate.toISOString(), end: endDate.toISOString() } ];
  const r = xmlRpcCall_(host, method, params, user, pass);
  if (!r.ok) {
    Logger.log('fetchDigiumCallSummary_ failed: ' + (r.error || r.raw));
    return { ok: false, error: r.error || r.raw };
  }

  // TODO: parse r.xml to extract totals per day and raw call rows
  // For now, return the raw XML so we can iterate on parsing once we confirm method/response shape.
  return { ok: true, xml: r.xml, raw: r.raw };
}

// Send a Digium-style <request method="..."> XML envelope. Returns {ok, raw, xml (XmlService.Document)}
function digiumApiCall_(methodName, parametersXmlString, user, pass, host) {
  try {
    const body = '<?xml version="1.0"?>\n' +
      '<request method="' + xmlEscape_(methodName) + '">\n' +
      '  <parameters>' + (parametersXmlString || '') + '</parameters>\n' +
      '</request>';

    const headers = { 'Content-Type': 'application/xml; charset=UTF-8', 'Accept': 'application/xml' };
    if (user && pass) headers['Authorization'] = 'Basic ' + Utilities.base64Encode(user + ':' + pass);
    const opts = { method: 'post', contentType: 'application/xml; charset=UTF-8', payload: body, headers: headers, muteHttpExceptions: true };
    const resp = UrlFetchApp.fetch(host, opts);
    const code = resp.getResponseCode();
    const txt = resp.getContentText();
    if (code >= 200 && code < 300) {
      try {
        const xml = XmlService.parse(txt);
        return { ok: true, raw: txt, xml: xml };
      } catch (e) {
        return { ok: false, error: 'XML parse error: ' + e.toString(), raw: txt };
      }
    }
    return { ok: false, error: 'HTTP ' + code, raw: txt };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// Fetch callReports.search for a date range and return parsed wide-format data when possible.
function fetchDigiumCallReports_(startDate, endDate, options) {
  const cfg = getCfg_();
  const host = cfg.digiumHost;
  const user = cfg.digiumUser;
  const pass = cfg.digiumPass;
  if (!host || !user || !pass) return { ok: false, reason: 'missing_credentials' };

  const fmt = (d) => {
    const dt = (d instanceof Date) ? d : new Date(d);
    const Y = dt.getFullYear();
    const M = String(dt.getMonth()+1).padStart(2,'0');
    const D = String(dt.getDate()).padStart(2,'0');
    return `${Y}-${M}-${D} 00:00:00`;
  };

  const startStr = fmt(startDate);
  const endStr = (() => { const dt = (endDate instanceof Date) ? endDate : new Date(endDate); return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')} 23:59:59`; })();

  const fields = (options && options.report_fields) || [
    'total_calls','total_incoming_calls','total_outgoing_calls','talking_duration','call_duration','avg_talking_duration','avg_call_duration'
  ];

  // Human-friendly labels for metrics (used in wide table rows)
  const human = {
    total_calls: 'Total Calls', total_incoming_calls: 'Total Incoming Calls', total_outgoing_calls: 'Total Outgoing Calls',
    talking_duration: 'Talking Duration (s)', call_duration: 'Call Duration (s)', avg_talking_duration: 'Avg Talking Duration (s)', avg_call_duration: 'Avg Call Duration (s)'
  };

  // prefer per-day breakdown for wide output; use Switchvox token 'by_day' per API docs
  const breakdown = (options && options.breakdown) || 'by_day';

  // Build report_fields XML
  let reportFieldsXml = '<report_fields>'; 
  fields.forEach(f => { reportFieldsXml += '<report_field>' + xmlEscape_(f) + '</report_field>'; });
  reportFieldsXml += '</report_fields>';

  // ignore_weekends: 0 => do not ignore weekends (user requested we do NOT ignore weekends)
  // Optionally include account_ids (extensions) to filter results per account
  let accountIdsXml = '';
  if (options && options.account_ids) {
    // accept comma-separated string or array
    const ids = Array.isArray(options.account_ids) ? options.account_ids : String(options.account_ids).split(',').map(s=>s.trim()).filter(Boolean);
    if (ids && ids.length) {
      accountIdsXml = '\n    <account_ids>' + ids.map(id => '<account_id>' + xmlEscape_(id) + '</account_id>').join('') + '</account_ids>';
    }
  }

  const paramsXml = `\n    <start_date>${xmlEscape_(startStr)}</start_date>\n    <end_date>${xmlEscape_(endStr)}</end_date>\n    <ignore_weekends>0</ignore_weekends>\n    <breakdown>${xmlEscape_(breakdown)}</breakdown>${accountIdsXml}\n    ${reportFieldsXml}\n    <format>xml</format>\n  `;

  const r = digiumApiCall_('switchvox.callReports.search', paramsXml, user, pass, host);
  // Save raw response for inspection — append rather than clear so history is preserved (batched)
  try { appendDigiumRaw_(paramsXml, 'Raw XML Response', r.raw || (r.error || '')); } catch (e) { Logger.log('Failed to append Digium_Raw: ' + e.toString()); }

  // If Digium returned an error (for example "Invalid breakdown"), try a fallback without the breakdown
  try {
    if (r.ok && r.xml) {
      const rootChk = r.xml.getRootElement ? r.xml.getRootElement() : null;
      const errorsEl = rootChk ? rootChk.getChild('errors') : null;
      if (errorsEl) {
        const errEl = errorsEl.getChild('error');
        const msg = errEl && errEl.getAttribute ? (errEl.getAttribute('message') ? errEl.getAttribute('message').getValue() : '') : '';
        Logger.log('Digium returned errors: ' + msg);
        // If breakdown is invalid, retry without breakdown param (aggregate totals)
        if (/invalid breakdown/i.test(String(msg || ''))) {
          try {
            // Remove the <breakdown>...</breakdown> element from the paramsXml
            const paramsNoBreakdown = paramsXml.replace(/\s*<breakdown>.*?<\/breakdown>\s*/i, ' ');
            try { appendDigiumRaw_('Server reported invalid breakdown, retrying without <breakdown>\n' + paramsNoBreakdown, 'Retry Raw XML Response', ''); } catch (e) {}
            const r2 = digiumApiCall_('switchvox.callReports.search', paramsNoBreakdown, user, pass, host);
            try { appendDigiumRaw_(paramsNoBreakdown, 'Retry Raw XML Response', r2.raw || (r2.error || '')); } catch (e) {}
            if (!r2.ok) return { ok: false, error: r2.error || r2.raw, raw: r.raw };
            // Replace r with successful fallback response
            r = r2;
          } catch (e) {
            Logger.log('Retry without breakdown failed: ' + e.toString());
            return { ok: false, error: e.toString(), raw: r.raw };
          }
        } else {
          return { ok: false, error: msg || (r.error || r.raw), raw: r.raw };
        }
      }
    }
  } catch (e) {
    Logger.log('Error checking Digium errors element: ' + e.toString());
  }

  if (!r.ok) return { ok: false, error: r.error || r.raw };

  // Helper: parse Digium numeric fields. Accept plain numbers or HH:MM:SS; return seconds (raw).
  const parseDigiumNum = (txt) => {
    if (txt == null) return 0;
    const s = String(txt).trim();
    if (!s) return 0;
    // HH:MM:SS
    if (/^\d{1,2}:\d{2}:\d{2}$/.test(s)) return parseDurationSeconds_(s);
    const n = Number(s);
    return isNaN(n) ? 0 : n;
  };

  // Try to parse per-day results. Best-effort: look for <day date="YYYY-MM-DD">...</day>
  try {
    const doc = r.xml;
    const root = doc.getRootElement();

    // Build dates array from start to end
    const dates = [];
    const d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    const ed = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
    while (d <= ed) {
      dates.push(d.toISOString().split('T')[0]);
      d.setDate(d.getDate() + 1);
    }

    const fieldMap = {};
    // Initialize map for each date
    dates.forEach(dt => { fieldMap[dt] = {}; fields.forEach(f => fieldMap[dt][f] = 0); });

    // Helper: find all <day> elements anywhere
    const findDays = (el) => {
      const out = [];
      const children = el.getChildren();
      children.forEach(c => {
        if (String(c.getName()).toLowerCase() === 'day' && c.getAttribute('date')) out.push(c);
        out.push(...findDays(c));
      });
      return out;
    };

    const days = findDays(root);
    if (days && days.length) {
      days.forEach(dayEl => {
        const dateAttr = dayEl.getAttribute('date') ? dayEl.getAttribute('date').getValue() : null;
        const dateKey = dateAttr ? dateAttr.split(' ')[0] : null;
        if (!dateKey) return;
        fields.forEach(f => {
          const child = dayEl.getChild(f);
          if (child) {
            const txt = (child.getText() || '').trim();
            const num = parseDigiumNum(txt);
            fieldMap[dateKey][f] = num;
          }
        });
      });
    } else {
      // Try <rows><row .../></rows> where each row has date and metric attributes
      const resultEl = root.getChild('result') || root;
      const rowsEl = resultEl.getChild('rows');
      const rowChildren = rowsEl ? rowsEl.getChildren('row') : null;
      if (rowChildren && rowChildren.length) {
        rowChildren.forEach(rowEl => {
          let dateAttr = rowEl.getAttribute('date') ? rowEl.getAttribute('date').getValue() : null;
          let dateKey = null;
          if (dateAttr) {
            // Convert MM/DD/YYYY to YYYY-MM-DD
            const m = String(dateAttr).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
            if (m) dateKey = `${m[3]}-${('0'+m[1]).slice(-2)}-${('0'+m[2]).slice(-2)}`; else dateKey = dateAttr;
          }
          if (!dateKey) return;
          if (!fieldMap[dateKey]) { fieldMap[dateKey] = {}; fields.forEach(f => fieldMap[dateKey][f] = 0); }
          fields.forEach(f => {
            const a = rowEl.getAttribute(f);
            if (a) {
              const num = parseDigiumNum(a.getValue());
              fieldMap[dateKey][f] = num;
            }
          });
        });
      } else {
        // Fallback: Digium returned aggregated attributes on a <rows ... /> element (no per-day breakdown)
        // Example: <rows total_calls="69" total_incoming_calls="43" ... />
        if (rowsEl && rowsEl.getAttributes && rowsEl.getAttributes().length > 0) {
          // Build a single-date column representing the requested range
          const rangeLabel = `${startDate.toISOString().split('T')[0]} to ${endDate.toISOString().split('T')[0]}`;
          const single = {};
          fields.forEach(f => { single[f] = 0; });
          rowsEl.getAttributes().forEach(attr => {
            const name = attr.getName();
            const val = parseDigiumNum(attr.getValue());
            if (fields.indexOf(name) >= 0) single[name] = val;
          });
          // Return single-column wide data labeled with the range
          const wideRows = fields.map(f => [ (human && human[f]) || f, single[f] ]);
          const singleDates = [rangeLabel];
          return { ok: true, dates: singleDates, rows: wideRows, raw: r.raw };
        }
      }
    }

    // Build wide rows
    const wideRows = [];
    fields.forEach(f => {
      const row = [ human[f] || f ];
      dates.forEach(dKey => row.push(fieldMap[dKey] ? fieldMap[dKey][f] || 0 : 0));
      wideRows.push(row);
    });

    return { ok: true, dates: dates, rows: wideRows, raw: r.raw };
  } catch (e) {
    return { ok: false, error: 'parse_failed: ' + e.toString(), raw: r.raw };
  }
}

// Create a wide-format Digium calls sheet and write provided data (dates + metrics)
function createDigiumCallsSheet_(wideData) {
  // wideData: { dates: ['2025-11-01', ...], rows: [ ['Total Calls', 5, 3, ...], ['Total Duration', 3600, ...] ] }
  const ss = SpreadsheetApp.getActive();
  const name = 'Digium_Calls';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  const header = ['Metric'].concat(wideData.dates || []);
  sh.getRange(1,1,1,header.length).setValues([header]);
  for (let i=0;i<(wideData.rows||[]).length;i++) {
    sh.getRange(2+i,1,1,wideData.rows[i].length).setValues([wideData.rows[i]]);
  }
  // Basic styling
  try { sh.getRange(1,1,1,header.length).setFontWeight('bold').setBackground('#1976D2').setFontColor('#FFFFFF'); } catch(e){}
  return sh;
}

// Append Digium request/response in one batched write to reduce calls
function appendDigiumRaw_(requestXml, label, raw) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Digium_Raw');
    if (!sh) sh = ss.insertSheet('Digium_Raw');
    const startRow = Math.max(1, sh.getLastRow()) + 2;
    const rows = [
      ['Request (XML)'],
      [requestXml || ''],
      [''],
      [label || 'Raw XML Response'],
      [raw || '']
    ];
    sh.getRange(startRow, 1, rows.length, 1).setValues(rows);
    return true;
  } catch (e) {
    Logger.log('appendDigiumRaw_ failed: ' + e.toString());
    return false;
  }
}

// Convert numeric seconds in a wide table (with a Metric label column) to spreadsheet day-fractions
// and apply hh:mm:ss formatting to the value cells. Looks for a header cell with text 'Metric'.
function convertSecondsToDayFractionForTableAt_(sheet, headerRow, headerCol) {
  try {
    if (!sheet) return { ok: false, reason: 'no_sheet' };
    const hdr = String(sheet.getRange(headerRow, headerCol).getValue() || '').toLowerCase();
    if (!hdr.includes('metric')) return { ok: false, reason: 'no_metric_header' };

    // Detect width
    const maxScan = Math.max(10, Math.min(100, sheet.getLastColumn() - headerCol + 1));
    const headerVals = sheet.getRange(headerRow, headerCol, 1, maxScan).getValues()[0];
    let width = 1;
    for (let i = 1; i < maxScan; i++) {
      if (headerVals[i] === '' || headerVals[i] == null) break;
      width = i + 1;
    }
    if (width <= 1) return { ok: false, reason: 'no_value_columns' };

    // Detect height
    let height = 0;
    for (let r = headerRow + 1; r <= sheet.getMaxRows(); r++) {
      const v = sheet.getRange(r, headerCol).getValue();
      if (v === '' || v == null) break;
      height++;
      if (height > 2000) break;
    }
    if (height <= 0) return { ok: true, converted: 0 };

    // Bulk read/write
    const blockRange = sheet.getRange(headerRow + 1, headerCol, height, width);
    const data = blockRange.getValues();
    const newData = new Array(height);
    const durationRows = [];
    const countRows = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const label = String(row[0] || '').toLowerCase();
      const isDuration = /duration|time/.test(label);
      const out = row.slice();
      for (let c = 1; c < width; c++) {
        const v = out[c];
        const n = (v === '' || v == null) ? '' : Number(v);
        if (n === '' || isNaN(n)) continue;
        out[c] = isDuration ? (n / 86400) : n;
      }
      newData[i] = out;
      if (isDuration) durationRows.push(i); else countRows.push(i);
    }

    blockRange.setValues(newData);

    function applyRuns(rows, fmt) {
      if (!rows.length || width <= 1) return;
      let runStart = rows[0], prev = rows[0];
      for (let k = 1; k <= rows.length; k++) {
        const curr = rows[k];
        if (k === rows.length || curr !== prev + 1) {
          sheet.getRange(headerRow + 1 + runStart, headerCol + 1, prev - runStart + 1, width - 1)
               .setNumberFormat(fmt);
          if (k < rows.length) runStart = curr;
        }
        prev = curr;
      }
    }
    applyRuns(durationRows, 'hh:mm:ss');
    applyRuns(countRows, '0');

    return { ok: true, converted: height };
  } catch (e) {
    Logger.log('convertSecondsToDayFractionForTableAt_ error: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

// No-arg wrapper to apply duration formatting across common sheets.
function formatAllDurationColumns() {
  try {
    const ss = SpreadsheetApp.getActive();
    // Sessions table: try to apply hh:mm:ss to known duration columns if Sessions exists
    try {
      const sessions = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
      if (sessions) {
        // Known duration-like columns used elsewhere in code: 19,20,21,22,31,32,33,34,35,36,37
        const durCols = [19,20,21,22,31,32,33,34,35,36,37];
        const lastRow = Math.max(2, sessions.getLastRow());
        durCols.forEach(c => {
          try { sessions.getRange(2, c, lastRow - 1, 1).setNumberFormat('hh:mm:ss'); } catch (e) {}
        });
      }
    } catch (e) { Logger.log('formatAllDurationColumns sessions formatting failed: ' + e.toString()); }

    // Per request: keep Digium data raw (seconds/counts). Do NOT auto-convert Digium tables here.
    // If time-formatting is desired in the future, do it in a read-only presentation layer referencing raw cells.

    SpreadsheetApp.getActive().toast('Duration formatting applied across sheets');
  } catch (e) {
    Logger.log('formatAllDurationColumns error: ' + e.toString());
    SpreadsheetApp.getActive().toast('Formatting failed: ' + e.toString().slice(0,200));
  }
}

// Read extension_map sheet and return { technician_name_lower: [ext1, ext2, ...] }
function getExtensionMap_() {
  const out = {};
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('extension_map');
    if (!sh) return out;
    const rng = sh.getDataRange();
    const vals = rng.getValues();
    if (!vals || vals.length < 2) return out;
    const headers = vals[0].map(h => String(h || '').trim().toLowerCase());
    const nameIdx = headers.indexOf('technician_name');
    const extIdx = headers.indexOf('extension');
    const activeIdx = headers.indexOf('is_active');
    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
      const extCell = extIdx >= 0 ? String(row[extIdx] || '').trim() : '';
      const isActive = activeIdx >= 0 ? String(row[activeIdx] || '').toLowerCase() !== 'false' : true;
      if (!name || !extCell || !isActive) continue;
      const exts = extCell.split(',').map(s => s.trim()).filter(Boolean);
      const key = name.toLowerCase();
      if (!out[key]) out[key] = [];
      out[key].push(...exts);
    }
  } catch (e) { Logger.log('getExtensionMap_ failed: ' + e.toString()); }
  return out;
}

// Return a roster of technician names from the extension_map sheet (preserves casing)
// Only includes rows where is_active != 'false' (defaults to active when column missing)
function getRosterTechnicianNames_() {
  const names = [];
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('extension_map');
    if (!sh) return names;
    const vals = sh.getDataRange().getValues();
    if (!vals || vals.length < 2) return names;
    const headers = vals[0].map(h => String(h || '').trim().toLowerCase());
    const nameIdx = headers.indexOf('technician_name');
    const activeIdx = headers.indexOf('is_active');
    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const active = activeIdx >= 0 ? String(row[activeIdx] || '').toLowerCase() !== 'false' : true;
      const name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
      if (active && name) names.push(name);
    }
  } catch (e) { Logger.log('getRosterTechnicianNames_ failed: ' + e.toString()); }
  // de-dup, keep first casing
  const seen = new Set();
  return names.filter(n => { const k = n.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; });
}

// Test helper: writes a small Digium-style table to a test sheet and runs the converter so you can validate formatting.
function writeDurationTestRow() {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Format_Test');
    if (!sh) sh = ss.insertSheet('Format_Test');
    sh.clear();
    const header = ['Metric', '2025-11-05'];
    sh.getRange(1,1,1,header.length).setValues([header]);
    sh.getRange(2,1,1,2).setValues([['Test Duration', 3661]]);
    // Run conversion
    const res = convertSecondsToDayFractionForTableAt_(sh, 1, 1);
    SpreadsheetApp.getActive().toast('Format_Test conversion: ' + JSON.stringify(res));
    return res;
  } catch (e) {
    Logger.log('writeDurationTestRow failed: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

// High-level helper: ingest Digium calls for an ISO date range (YYYY-MM-DD) and write wide-format summary.
// This is a convenience entrypoint you can run from the Apps Script editor.
function ingestDigiumRangeToSheets_(startISO, endISO) {
  try {
    const startDate = new Date(startISO + 'T00:00:00Z');
    const endDate = new Date(endISO + 'T23:59:59Z');
    // Use the new callReports.search flow (will save raw XML and attempt to parse per-day metrics)
    const res = fetchDigiumCallReports_(startDate, endDate, {});
    if (!res.ok) {
      SpreadsheetApp.getActive().toast('Digium fetch failed: ' + (res.error || res.reason || 'unknown'));
      return { ok: false, error: res.error || res.reason };
    }

    // If the helper returned parsed wide-format data, write it; otherwise fall back to raw saving
    const ss = SpreadsheetApp.getActive();
    if (res.dates && res.rows) {
      createDigiumCallsSheet_({ dates: res.dates, rows: res.rows });
      SpreadsheetApp.getActive().toast('Digium response parsed and saved to Digium_Calls');
      return { ok: true };
    } else {
      try { appendDigiumRaw_('', 'Raw XML Response', res.raw || ''); } catch (e) {}
      SpreadsheetApp.getActive().toast('Digium response appended to Digium_Raw (no parsed metrics)');
      return { ok: true };
    }
  } catch (e) {
    Logger.log('ingestDigiumRangeToSheets_ error: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

// No-arg wrapper so the function shows in the Apps Script Run list and can be executed interactively.
function ingestDigiumRangeToSheets() {
  const ui = SpreadsheetApp.getUi();
  const now = new Date();
  const y = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
  const defaultStart = y.toISOString().split('T')[0];
  const defaultEnd = defaultStart;
  const res = ui.prompt('Digium Ingest', `Enter start and end dates (YYYY-MM-DD,YYYY-MM-DD) or single date.`, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const txt = res.getResponseText().trim();
  if (!txt) {
    ui.alert('No dates entered');
    return;
  }
  const parts = txt.split(',').map(s => s.trim()).filter(Boolean);
  const start = parts[0] || defaultStart;
  const end = parts[1] || start;
  ui.showModalDialog(HtmlService.createHtmlOutput(`<p>Starting ingest for ${start} → ${end}. This may take a moment.</p>`).setWidth(300).setHeight(80), 'Digium Ingest');
  const out = ingestDigiumRangeToSheets_(start, end);
  if (out && out.ok) ui.alert('Digium ingest finished (raw saved to Digium_Raw).'); else ui.alert('Digium ingest failed: ' + (out && out.error ? out.error : 'unknown'));
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

// Per Rescue API docs (see API.asmx?op=setOutput), the allowed HTTP GET endpoint is /API/setOutput.aspx
// with parameter 'output=XML' or 'output=TEXT'. We use GET with the authenticated session cookie.
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
    // Accept both TEXT (starts with OK) and XML (starts with '<') responses
    if (/^OK/i.test(t)) return t;
    if (t && t[0] === '<') return t;
    return null;
  } catch (e) {
    Logger.log(`getReportTry_ failed for node ${nodeId} (${noderef}): ${e.toString()}`);
    return null;
  }
}

/* ===== Parsing ===== */
function parsePipe_(okBody, delimiter) {
  if (!okBody || typeof okBody !== 'string') return { headers: [], rows: [] };
  const raw = String(okBody).trim();
  // If the response is XML (due to setOutput XML), parse XML instead
  if (raw && raw[0] === '<') return parseRescueXmlBody_(raw);
  const body = raw.replace(/^OK\s*/i, '').trim();
  if (!body) return { headers: [], rows: [] };
  if (body[0] === '<') return parseRescueXmlBody_(body);
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

// Parse Rescue XML report content into { headers, rows }
function parseRescueXmlBody_(xmlText) {
  try {
    const doc = XmlService.parse(String(xmlText));
    const root = doc.getRootElement();
    // Collect all elements
    const all = [];
    (function walk(el){
      all.push(el);
      const kids = el.getChildren();
      for (let i=0;i<kids.length;i++) walk(kids[i]);
    })(root);

    // Choose the most frequent element name as row tag, prefer common names
    const counts = {};
    const prefer = ['row','record','item','entry'];
    all.forEach(el => {
      const n = String(el.getName()||'').toLowerCase();
      counts[n] = (counts[n]||0) + 1;
    });
    let rowName = null;
    for (const p of prefer) { if (counts[p] && (!rowName || counts[p] > counts[rowName])) rowName = p; }
    if (!rowName) {
      // fallback to the most frequent non-root element
      rowName = Object.keys(counts).sort((a,b)=>counts[b]-counts[a])[0];
    }
    let rows = all.filter(el => String(el.getName()||'').toLowerCase() === String(rowName||'').toLowerCase());
    if (!rows.length) {
      // fallback: any element that has attributes (likely row-like)
      rows = all.filter(el => (el.getAttributes()||[]).length);
    }
    if (!rows.length) return { headers: [], rows: [] };

    // Build header union from attributes and immediate child elements
    const headerSet = new Set();
    rows.forEach(el => {
      (el.getAttributes()||[]).forEach(a => headerSet.add(a.getName()));
      (el.getChildren()||[]).forEach(c => headerSet.add(c.getName()));
    });
    const headers = Array.from(headerSet);
    if (!headers.length) return { headers: [], rows: [] };

    const out = rows.map(el => {
      const obj = {};
      (el.getAttributes()||[]).forEach(a => { obj[a.getName()] = a.getValue(); });
      (el.getChildren()||[]).forEach(c => { obj[c.getName()] = c.getText ? c.getText().trim() : ''; });
      return obj;
    });
    return { headers: headers, rows: out };
  } catch (e) {
    Logger.log('parseRescueXmlBody_ error: ' + e.toString());
    return { headers: [], rows: [] };
  }
}

// Get current output format from Rescue ('XML' or 'TEXT')
function getOutput_(base, cookie) {
  try {
    // Per Rescue API docs, allowed HTTP GET endpoint is /API/getOutput.aspx
    const r = apiGet_(base, 'getOutput.aspx', {}, cookie, 2, true);
    const t = (r.getContentText()||'').trim();
    // Accept common plain text formats
    let m = t.match(/OUTPUT\s*:\s*(XML|TEXT)/i) || t.match(/OK\s+(XML|TEXT)/i);
    if (m) return m[1].toUpperCase();
    // Some environments may return minimal XML; try to parse
    if (t && t[0] === '<') {
      try {
        const doc = XmlService.parse(t);
        const root = doc.getRootElement();
        // Search for any element named 'output' (case-insensitive) and read its text/value
        const stack = [root];
        while (stack.length) {
          const el = stack.pop();
          if (String(el.getName()||'').toLowerCase() === 'output') {
            const val = (el.getText ? el.getText() : '') || (el.getAttribute && el.getAttribute('value') ? el.getAttribute('value').getValue() : '');
            if (/^xml$/i.test(val)) return 'XML';
            if (/^text$/i.test(val)) return 'TEXT';
          }
          const kids = el.getChildren();
          for (let i=0;i<kids.length;i++) stack.push(kids[i]);
        }
      } catch (e) { /* ignore XML parse failure */ }
    }
    return null;
  } catch (e) {
    Logger.log('getOutput_ error: ' + e.toString());
    return null;
  }
}

// Log unmapped header names (helpful when API changes header labels)
function logUnmappedHeader_(header, sample) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheetName = 'Debug_Unmapped_Headers';
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.getRange(1,1,1,5).setValues([['Header','Sample','FirstSeen','LastSeen','Count']]);
      sh.setColumnWidth(1, 300);
      sh.setColumnWidth(2, 300);
      sh.setColumnWidth(3, 180);
      sh.setColumnWidth(4, 180);
      sh.setColumnWidth(5, 80);
    }
    const headerLower = String(header || '').toLowerCase();
    const dataRange = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1),1);
    const existing = dataRange.getValues().map(r => String(r[0] || '').toLowerCase());
    const foundIdx = existing.indexOf(headerLower);
    const now = new Date().toISOString();
    if (foundIdx >= 0) {
      const rowNum = foundIdx + 2;
      // preserve original sample if present, update last seen and increment count
      const sampleCell = sh.getRange(rowNum, 2);
      if (!sampleCell.getValue()) sampleCell.setValue(sample || '');
      sh.getRange(rowNum, 4).setValue(now);
      const countCell = sh.getRange(rowNum, 5);
      const prev = Number(countCell.getValue()) || 1;
      countCell.setValue(prev + 1);
    } else {
      sh.appendRow([String(header || ''), String(sample || ''), now, now, 1]);
    }
  } catch (e) {
    Logger.log('logUnmappedHeader_ failed: ' + e.toString());
  }
}

/* ===== Mapping to Schema ===== */
function mapRow_(rec) {
  if (!rec || typeof rec !== 'object') return null;
  // Normalize incoming keys from XML or TEXT: lowercase and strip non-alphanumerics
  const nk = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const normMap = (() => {
    const m = new Map();
    try {
      Object.keys(rec || {}).forEach(k => {
        const n = nk(k);
        if (n && !m.has(n)) m.set(n, k);
      });
    } catch (e) {}
    return m;
  })();
  const g = (o, keys, d='') => { 
    for (const k of keys) {
      // Exact key match
      if (o[k] != null && String(o[k]).length) return String(o[k]).trim();
      // Normalized key match (handles SessionID, session_id, Session Id, etc.)
      try {
        const actual = normMap.get(nk(k));
        if (actual != null && o[actual] != null && String(o[actual]).length) return String(o[actual]).trim();
      } catch (e) {}
    }
    return d; 
  };
  const toSec = (val) => {
    const s = String(val||'').trim();
    // Preserve empty/blank values as empty string so Sessions sheet reflects the API
    if (!s) return '';
    // If it's a plain integer string, return a Number so numeric formulas can consume it
    if (/^\d+$/.test(s)) return Number(s);
    // Parse HH:MM:SS style strings
    const m = s.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
    if (m) return (Number(m[1])*3600 + Number(m[2])*60 + Number(m[3]))|0;
    // Otherwise, preserve the original string value
    return s;
  };
  const toTs = (s) => {
    // Preserve original timestamp string from API; do not coerce here so Sessions
    // sheet shows raw API values. Caller may parse/convert as needed elsewhere.
    const v = String(s||'').trim();
    if (!v) return '';
    return v;
  };
  // Detect and log any header keys present in the incoming record that our mapping
  // doesn't explicitly handle. This helps capture new API header names for later
  // mapping updates.
  try {
    const knownVariants = [
      'session id','session type','status','technician id','technician name','technician email','technician group',
      'your name:','your name','customer name','customer','caller name','caller_name','customer_name','customername','contact name','client name','yourname',
      'your email:','customer email','customer_email','email','tracking id','customer ip','device id','platform','browser type','host name',
      'start time','end time','last action time','active time','work time','total time','waiting time','channel id','channel name',
      'company name:','company name','company','company_name','companyname','caller phone','caller_phone','phone','phone number','customer_phone','telephone','tel','mobile','mobile phone',
      'resolved unresolved','calling card','browser type','connecting time','time in transfer','reconnecting time','rebooting time','ingested_at'
    ];
    const knownSet = new Set(knownVariants.map(k => String(k).toLowerCase().trim()));
    Object.keys(rec || {}).forEach(k => {
      try {
        if (!k) return;
        const kl = String(k).toLowerCase().trim();
        if (!knownSet.has(kl)) {
          const v = rec[k];
          if (v != null && String(v).trim().length) {
            // Log unmapped header for later inspection (non-blocking)
            try { logUnmappedHeader_(k, String(v).slice(0, 200)); } catch (e) {}
          }
        }
      } catch (e) {}
    });
  } catch (e) {}
  return {
    session_id: g(rec, ['Session ID','SessionID','session_id','sessionid'], ''),
    session_type: g(rec, ['Session Type','SessionType','session_type','sessiontype'], ''),
    session_status: g(rec, ['Status','Session Status','session_status','sessionstatus'], ''),
    technician_id: g(rec, ['Technician ID','TechnicianID','technician_id','technicianid'], ''),
    technician_name: g(rec, ['Technician Name','Technician','TechnicianName','technician_name','technician'], ''),
    technician_email: g(rec, ['Technician Email','TechnicianEmail','technician_email','email'], ''),
    technician_group: g(rec, ['Technician Group','Group','Group Name','technician_group','group_name'], ''),
  customer_name: g(rec, ['Your Name:', 'Your Name', 'Customer Name', 'Customer', 'Caller Name', 'caller_name', 'customer_name', 'CustomerName', 'Contact Name', 'Client Name', 'YourName'], ''),
  customer_email: g(rec, ['Your Email:', 'Customer Email', 'customer_email', 'email', 'Email'], ''),
    tracking_id: g(rec, ['Tracking ID','TrackingID','tracking_id'], ''),
    ip_address: g(rec, ['Customer IP','IP Address','IPAddress','ip_address'], ''),
    device_id: g(rec, ['Device ID','DeviceID','device_id'], ''),
  platform: g(rec, ['Platform','platform'], ''),
    browser: g(rec, ['Browser Type','Browser','browser_type','browser'], ''),
    host: g(rec, ['Host Name','Host','host','host_name'], ''),
  // Additional fields present in LISTALL headers
  location_name: g(rec, ['Location Name:','Location Name','location_name'], ''),
  custom_field_4: g(rec, ['Custom field 4','Custom Field 4','custom_field_4'], ''),
  custom_field_5: g(rec, ['Custom field 5','Custom Field 5','custom_field_5'], ''),
  incident_tools_used: g(rec, ['Incident Tools Used','incident_tools_used'], ''),
    start_time: toTs(g(rec, ['Start Time','StartTime','start_time','starttime'], '')),
    end_time: toTs(g(rec, ['End Time','EndTime','end_time','endtime'], '')),
    last_action_time: toTs(g(rec, ['Last Action Time','LastActionTime','last_action_time','lastactiontime'], '')),
    duration_active_seconds: toSec(g(rec, ['Active Time','ActiveTime','active_time','activetime'], '')),
    duration_work_seconds: toSec(g(rec, ['Work Time','WorkTime','work_time','worktime'], '')),
    duration_total_seconds: toSec(g(rec, ['Total Time','TotalTime','total_time','totaltime'], '')),
    pickup_seconds: toSec(g(rec, ['Waiting Time','WaitingTime','waiting_time','pick_up','pickup','pickup_seconds'], '')),
    channel_id: g(rec, ['Channel ID','ChannelID','channel_id'], ''),
    channel_name: g(rec, ['Channel Name','Channel','ChannelName','channel_name'], ''),
  company_name: g(rec, ['Company name:', 'Company Name', 'Company', 'company_name', 'CompanyName'], ''),
  caller_name: g(rec, ['Caller Name', 'Caller', 'caller_name', 'Your Name:', 'Your Name', 'Contact Name'], ''),
  caller_phone: g(rec, ['Your Phone #:', 'Caller Phone', 'Phone', 'Phone Number', 'caller_phone', 'customer_phone', 'Telephone', 'Tel', 'Mobile', 'Mobile Phone'], ''),
    resolved_unresolved: g(rec, ['Resolved Unresolved','resolved_unresolved'], ''),
    calling_card: g(rec, ['Calling Card','calling_card'], ''),
    browser_type: g(rec, ['Browser Type','Browser','browser_type'], ''),
    connecting_time: toSec(g(rec, ['Connecting Time','ConnectingTime','connecting_time'], '')),
    waiting_time: toSec(g(rec, ['Waiting Time','WaitingTime','waiting_time'], '')),
    total_time: toSec(g(rec, ['Total Time','TotalTime','total_time'], '')),
    active_time: toSec(g(rec, ['Active Time','ActiveTime','active_time'], '')),
    work_time: toSec(g(rec, ['Work Time','WorkTime','work_time'], '')),
    hold_time: toSec(g(rec, ['Hold Time','HoldTime','hold_time'], '')),
    time_in_transfer: toSec(g(rec, ['Time in Transfer','TimeInTransfer','time_in_transfer'], '')),
    reconnecting_time: toSec(g(rec, ['Reconnecting Time','ReconnectingTime','reconnecting_time'], '')),
    rebooting_time: toSec(g(rec, ['Rebooting Time','RebootingTime','rebooting_time'], '')),
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

// Normalize a duration-like value into seconds.
// Accepts: numeric seconds, numeric time-fraction (days), or hh:mm:ss string.
function parseDurationSeconds_(val) {
  if (val == null || val === '') return 0;
  // If already a number
  if (typeof val === 'number') {
    // If small (< 2) assume it's a time-fraction (days) and convert to seconds
    if (Math.abs(val) < 2) return Math.round(val * 86400);
    // Otherwise assume it's raw seconds
    return Math.round(val);
  }
  // If string, try to parse numeric first
  const s = String(val).trim();
  if (!s) return 0;
  if (/^[0-9]+(\.[0-9]+)?$/.test(s)) {
    const n = Number(s);
    if (Math.abs(n) < 2) return Math.round(n * 86400);
    return Math.round(n);
  }
  // Try HH:MM:SS or MM:SS
  const parts = s.split(':').map(p => Number(p));
  if (parts.length === 3 && parts.every(p => !isNaN(p))) {
    return parts[0]*3600 + parts[1]*60 + parts[2];
  }
  if (parts.length === 2 && parts.every(p => !isNaN(p))) {
    return parts[0]*60 + parts[1];
  }
  // Fallback: try to parse as numeric with non-digits removed
  const n2 = Number(s.replace(/[^0-9.-]/g,''));
  return isNaN(n2) ? 0 : Math.round(n2);
}

/* ===== Sheets Storage ===== */
function getOrCreateSessionsSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
  if (!sh) {
    sh = ss.insertSheet(SHEETS_SESSIONS_TABLE);
  }
  
  // Use the exact column titles from LogMeIn Rescue API (Report Area: Session, Type: LISTALL)
  // This ensures we store raw data exactly as returned to avoid mixed-column issues.
  const headers = [
    'Start Time','End Time','Last Action Time',
    'Technician Name','Technician ID','Technician Email','Technician Group',
    'Session ID','Session Type','Status',
    'Your Name:','Your Phone #:','Company name:','Location Name:',
    'Custom field 4','Custom field 5','Tracking ID','Customer IP','Device ID','Incident Tools Used',
    'Resolved Unresolved','Channel ID','Channel Name','Calling Card',
    'Connecting Time','Waiting Time','Total Time','Active Time','Work Time','Hold Time','Time in Transfer','Rebooting Time','Reconnecting Time',
    'Platform','Browser Type','Host Name',
    // Additional internal field for ingestion timestamp
    'ingested_at'
  ];
  
  // Check if header row exists and is correct
  const headerRange = sh.getRange(1, 1, 1, headers.length);
  const existingHeaders = headerRange.getValues()[0];
  const needsHeaderUpdate = !existingHeaders || existingHeaders.length !== headers.length || 
                           !existingHeaders[0] || existingHeaders[0] !== 'Start Time';
  
  if (needsHeaderUpdate) {
    // Clear and set headers with improved styling
    headerRange.clear();
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold')
      .setBackground('#07123B')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    Logger.log('Updated Sessions sheet headers');
  }
  
  // Set column widths for better readability
  // Set key column widths for readability
  sh.setColumnWidth(1, 160);  // Start Time
  sh.setColumnWidth(2, 160);  // End Time
  sh.setColumnWidth(3, 160);  // Last Action Time
  sh.setColumnWidth(4, 180);  // Technician Name
  sh.setColumnWidth(6, 200);  // Technician Email
  sh.setColumnWidth(9, 160);  // Session Type
  sh.setColumnWidth(11, 160); // Your Phone #:
  sh.setColumnWidth(13, 180); // Company name:
  sh.setColumnWidth(23, 160); // Channel Name
  sh.setColumnWidth(headers.length, 180); // ingested_at
  
  // Apply consistent styling
  try { applyProfessionalTableStyling_(sh, headers.length); } catch (e) { Logger.log('Styling sessions sheet failed: ' + e.toString()); }

  // No duplicate columns in the official header set

  return sh;
}

// Migrate existing Sessions sheet to official LISTALL headers with correct column mapping.
// Keeps raw values; attempts to map from prior normalized headers and synonyms.
function migrateSessionsSheetToOfficial_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
  if (!sh) { SpreadsheetApp.getActive().toast('Sessions sheet not found'); return false; }

  // Official header order (same as getOrCreateSessionsSheet_)
  const official = [
    'Start Time','End Time','Last Action Time','Technician Name','Technician ID','Technician Email','Technician Group',
    'Session ID','Session Type','Status','Your Name:','Your Phone #:','Company name:','Location Name:',
    'Custom field 4','Custom field 5','Tracking ID','Customer IP','Device ID','Incident Tools Used',
    'Resolved Unresolved','Channel ID','Channel Name','Calling Card',
    'Connecting Time','Waiting Time','Total Time','Active Time','Work Time','Hold Time','Time in Transfer','Rebooting Time','Reconnecting Time',
    'Platform','Browser Type','Host Name','ingested_at'
  ];

  // Build alias -> canonical map (canonical is close to snake_case keys)
  const A = (k)=>k.toLowerCase().trim();
  const alias = new Map([
    ['Start Time',['Start Time','start_time']],
    ['End Time',['End Time','end_time']],
    ['Last Action Time',['Last Action Time','last_action_time']],
    ['Technician Name',['Technician Name','technician_name','Technician']],
    ['Technician ID',['Technician ID','technician_id']],
    ['Technician Email',['Technician Email','technician_email']],
    ['Technician Group',['Technician Group','technician_group','group_name']],
    ['Session ID',['Session ID','session_id']],
    ['Session Type',['Session Type','session_type']],
    ['Status',['Status','session_status']],
    ['Your Name:',['Your Name:','Your Name','Customer Name','customer_name']],
    ['Your Phone #:',['Your Phone #:','Phone','Phone Number','Caller Phone','caller_phone','customer_phone']],
    ['Company name:',['Company name:','Company Name','Company','company_name']],
    ['Location Name:',['Location Name:','Location Name','location_name']],
    ['Custom field 4',['Custom field 4','Custom Field 4','custom_field_4']],
    ['Custom field 5',['Custom field 5','Custom Field 5','custom_field_5']],
    ['Tracking ID',['Tracking ID','tracking_id']],
    ['Customer IP',['Customer IP','IP Address','ip_address']],
    ['Device ID',['Device ID','device_id']],
    ['Incident Tools Used',['Incident Tools Used','incident_tools_used']],
    ['Resolved Unresolved',['Resolved Unresolved','resolved_unresolved']],
    ['Channel ID',['Channel ID','channel_id']],
    ['Channel Name',['Channel Name','channel_name']],
    ['Calling Card',['Calling Card','calling_card']],
    ['Connecting Time',['Connecting Time','connecting_time']],
    ['Waiting Time',['Waiting Time','waiting_time','pickup_seconds']],
    ['Total Time',['Total Time','total_time','duration_total_seconds']],
    ['Active Time',['Active Time','active_time','duration_active_seconds']],
    ['Work Time',['Work Time','work_time','duration_work_seconds']],
    ['Hold Time',['Hold Time','hold_time']],
    ['Time in Transfer',['Time in Transfer','time_in_transfer']],
    ['Rebooting Time',['Rebooting Time','rebooting_time']],
    ['Reconnecting Time',['Reconnecting Time','reconnecting_time']],
    ['Platform',['Platform','platform']],
    ['Browser Type',['Browser Type','browser_type','browser']],
    ['Host Name',['Host Name','host','host_name']],
    ['ingested_at',['ingested_at','Ingested At']]
  ]);

  // Current header row
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol === 0 || lastRow < 1) { SpreadsheetApp.getActive().toast('Sessions sheet empty'); return false; }
  const currentHeaders = sh.getRange(1,1,1,lastCol).getValues()[0];
  const currentMap = new Map();
  currentHeaders.forEach((h, i) => { if (h != null && String(h).trim()) currentMap.set(A(String(h)), i); });

  // Resolve mapping: for each official header, find source column index
  const resolveIndex = (header) => {
    const alts = alias.get(header) || [header];
    for (const alt of alts) {
      const idx = currentMap.get(A(alt));
      if (typeof idx === 'number') return idx;
    }
    return -1;
  };

  // Read all data rows
  const dataRows = lastRow > 1 ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];

  // Build new table with official headers
  const newHeaders = official.slice();
  const newValues = dataRows.map(row => {
    return newHeaders.map(h => {
      const srcIdx = resolveIndex(h);
      return (srcIdx >= 0 && srcIdx < row.length) ? row[srcIdx] : '';
    });
  });

  // Replace sheet contents
  sh.clear();
  sh.getRange(1,1,1,newHeaders.length).setValues([newHeaders]);
  if (newValues.length) sh.getRange(2,1,newValues.length,newHeaders.length).setValues(newValues);
  try { applyProfessionalTableStyling_(sh, newHeaders.length); } catch (e) {}
  SpreadsheetApp.getActive().toast(`Sessions migrated to official headers (${newHeaders.length} columns)`);
  return true;
}

// Public wrapper for menu/run list
function migrateSessionsHeaders() {
  try { migrateSessionsSheetToOfficial_(); } catch (e) { SpreadsheetApp.getActive().toast('Migration failed: ' + e.toString()); }
}

// Apply consistent, professional styling to a sheet's table area.
// headerCols: number of header columns (optional). If omitted, uses sheet.getLastColumn().
function applyProfessionalTableStyling_(sheet, headerCols) {
  try {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = headerCols || sheet.getLastColumn();
    // Freeze header
    sheet.setFrozenRows(1);

    // Header styling
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight('bold')
  .setBackground('#07123B')
      .setFontColor('#FFFFFF')
      .setFontSize(11)
      .setHorizontalAlignment('left')
      .setFontFamily('Arial');

    // Apply subtle banding to data rows
    if (lastRow > 1) {
      try { sheet.getRange(2, 1, Math.max(1, lastRow - 1), lastCol).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) {}
    }

    // Create a filter for easy sorting/searching (ignore errors if already present)
    try { sheet.getRange(1, 1, Math.max(1, lastRow), lastCol).createFilter(); } catch (e) {}

    // Default font + vertical alignment for table
    try { sheet.getRange(1, 1, Math.max(1, lastRow), lastCol).setFontFamily('Arial').setVerticalAlignment('middle'); } catch (e) {}

    // Align numeric-like columns to the right based on header keywords
    try {
      const hdrs = headerRange.getValues()[0].map(h => String(h || '').toLowerCase());
      const numericKeywords = ['session', 'sessions', 'duration', 'time', 'total', 'pickup', 'id', 'count', 'seconds', 'hours', 'work'];
      hdrs.forEach((h, i) => {
        try {
          if (numericKeywords.some(k => h.indexOf(k) !== -1)) {
            sheet.getRange(2, i + 1, Math.max(1, lastRow - 1), 1).setHorizontalAlignment('right');
          } else {
            sheet.getRange(2, i + 1, Math.max(1, lastRow - 1), 1).setHorizontalAlignment('left');
          }
        } catch (e) {}
      });
    } catch (e) {}

  } catch (e) {
    Logger.log('applyProfessionalTableStyling_ failed: ' + e.toString());
  }
}

function writeRowsToSheets_(ss, rows, clearExisting = false) {
  if (!rows || !rows.length) return 0;
  const sh = getOrCreateSessionsSheet_(ss);
  
  // Clear existing data if requested (for range-specific pulls)
  // IMPORTANT: Keep header row (row 1) intact
  if (clearExisting) {
    const dataRange = sh.getDataRange();
    if (dataRange.getNumRows() > 1) {
      // Clear contents of all data rows (keep header and sheet structure)
      const numRows = dataRange.getNumRows() - 1;
      const numCols = dataRange.getNumColumns();
      sh.getRange(2, 1, numRows, numCols).clearContent();
      Logger.log('Cleared existing Sessions data (kept headers)');
    }
  }
  
  const existingIds = new Set();
  const dataRange = sh.getDataRange();
  if (dataRange.getNumRows() > 1) {
    // Find the 'Session ID' column in the Sessions table header for proper de-duplication
    const headerVals = sh.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    let idColIdx = headerVals.findIndex(h => String(h || '').toLowerCase().trim() === 'session id');
    if (idColIdx < 0) idColIdx = headerVals.findIndex(h => /session\s*id/i.test(String(h || '')));
    if (idColIdx < 0) idColIdx = 7; // Fallback to 0-based index 7 (8th col) per official header order
    const existing = sh.getRange(2, idColIdx + 1, dataRange.getNumRows() - 1, 1).getValues();
    existing.forEach(r => { const v = r[0]; if (v != null && String(v).trim() !== '') existingIds.add(String(v)); });
  }
  // Filter by session_id
  const toInsert = rows.filter(r => r && r.session_id && !existingIds.has(String(r.session_id)));
  if (!toInsert.length) return 0;

  // Heuristic normalization to fix common header/value shifts seen in some API responses.
  // If the API returns different header labels the parsed fields can shift; apply
  // content-based fixes so columns like caller_name / caller_phone / company_name
  // contain sensible values.
  const isPhone = (s) => { if (!s) return false; const t = String(s).replace(/[^0-9]/g,''); return t.length >= 6; };
  const isName = (s) => { if (!s) return false; const t = String(s).trim(); if (t.length < 2) return false; if (/\d/.test(t)) return false; if (/\b(rc|sp|ps)\b/i.test(t)) return false; return /^[A-Za-z .,'-]+$/.test(t); };
  const isStatus = (s) => { if (!s) return false; return /closed|waiting|active|connected|resolved|unresolved|connecting|in session|closed by/i.test(String(s)); };

  toInsert.forEach(r => {
    try {
      // Pattern observed: company_name contains phone, caller_name contains status text,
      // caller_phone contains the actual caller name. Shift into correct fields.
      if (isStatus(r.caller_name) && isPhone(r.company_name) && isName(r.caller_phone)) {
        // Move status into session_status
        r.session_status = r.caller_name;
        // Move caller name from caller_phone into caller_name
        r.caller_name = r.caller_phone;
        // Move phone from company_name into caller_phone
        r.caller_phone = r.company_name;
        // Clear company_name (unknown)
        r.company_name = '';
      }

      // If company_name looks like a phone but caller_phone is empty, move it
      if (!r.caller_phone && isPhone(r.company_name)) {
        r.caller_phone = r.company_name;
        r.company_name = '';
      }

      // If caller_phone appears numeric and company_name looks like a name, swap
      if (isPhone(r.caller_phone) && isName(r.company_name)) {
        const tmp = r.caller_phone;
        r.caller_phone = tmp;
        // company_name likely already correct; leave as-is
      }
    } catch (e) { /* ignore per-row normalization errors */ }
  });

  // Additional normalization: some API rows arrive shifted so session_id contains a timestamp
  // Detect session_id values that look like dates/times and shift them into start_time
  const looksLikeDateTime = (s) => {
    if (!s) return false;
    const t = String(s).trim();
    // common patterns: MM/DD/YYYY or YYYY-MM-DD or include ':' for time
    if (/\d{1,2}\/\d{1,2}\/\d{2,4}/.test(t)) return true;
    if (/\d{4}-\d{2}-\d{2}/.test(t)) return true;
    if (/\d{1,2}:\d{2}:\d{2}/.test(t)) return true;
    return false;
  };
  toInsert.forEach(r => {
    try {
      if (looksLikeDateTime(r.session_id) && !r.start_time) {
        // move timestamp into start_time and attempt to get a better session id
        const ts = r.session_id;
        r.start_time = ts;
        // prefer tracking_id as fallback session id
        if (r.tracking_id && String(r.tracking_id).length) {
          r.session_id = String(r.tracking_id);
        } else {
          r.session_id = '';
        }
      }
    } catch(e) {}
  });
  
  // Convert ISO timestamp strings to Date objects for proper timezone handling
  const toDate = (isoStr) => {
    if (!isoStr) return null;
    try {
      return new Date(isoStr);
    } catch (e) {
      return null;
    }
  };
  
  // Map data to the official LISTALL column order defined in getOrCreateSessionsSheet_ headers
  const values = toInsert.map(r => [
    // Start/End/Last Action Time (keep raw API strings when STORE_RAW_SESSIONS is true)
    (STORE_RAW_SESSIONS ? (r.start_time || '') : toDate(r.start_time)),
    (STORE_RAW_SESSIONS ? (r.end_time || '') : toDate(r.end_time)),
    (STORE_RAW_SESSIONS ? (r.last_action_time || '') : toDate(r.last_action_time)),
    // Technician details
    r.technician_name, r.technician_id, r.technician_email, r.technician_group,
    // Session IDs and status
    r.session_id, r.session_type, r.session_status,
    // Customer fields
    (r.customer_name || r.caller_name), r.caller_phone, r.company_name, r.location_name || '',
    // Custom fields
    r.custom_field_4 || '', r.custom_field_5 || '', r.tracking_id, r.ip_address, r.device_id, r.incident_tools_used || '',
    // Resolution / Channel / Calling Card
    r.resolved_unresolved, r.channel_id, r.channel_name, r.calling_card,
    // Timings (keep original seconds or hh:mm:ss as stored by mapRow_)
    r.connecting_time, r.waiting_time, r.total_time, r.active_time, r.work_time, r.hold_time, r.time_in_transfer, r.rebooting_time, r.reconnecting_time,
    // Platform / Browser / Host
    r.platform, r.browser_type || r.browser, r.host,
    // Ingested timestamp
    (STORE_RAW_SESSIONS ? (r.ingested_at || '') : toDate(r.ingested_at))
  ]);
  
  const newRowStart = sh.getLastRow() + 1;
  // Per your direction: keep raw values for timings (seconds or HH:MM:SS). Do not auto-convert to day-fractions here.
  sh.getRange(newRowStart, 1, values.length, values[0].length).setValues(values);
  
  // Optional: if not storing raw, apply date formats to the first three columns and ingested_at
  const dateFormat = 'mm/dd/yyyy hh:mm:ss AM/PM';
  if (!STORE_RAW_SESSIONS && values.length > 0) {
    sh.getRange(newRowStart, 1, values.length, 3).setNumberFormat(dateFormat); // Start/End/Last Action
    sh.getRange(newRowStart, headers.length, values.length, 1).setNumberFormat(dateFormat); // ingested_at
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
    // Ensure LISTALL report type is set and verified
    setReportTypeListAll_(cfg.rescueBase, cookie);
    try {
      const rt = getReportType_(cfg.rescueBase, cookie);
      if (rt !== 'LISTALL') {
        Logger.log(`Report type after first set was ${rt}, re-applying LISTALL`);
        setReportTypeListAll_(cfg.rescueBase, cookie);
      }
    } catch (e) {
      Logger.log('getReportType check failed (non-fatal): ' + e.toString());
    }
    setOutputXMLOrFallback_(cfg.rescueBase, cookie);
    setDelimiter_(cfg.rescueBase, cookie, '|');
    // Extract date strings (YYYY-MM-DD) for API date setting
    const startDateIso = startTimestamp.split('T')[0];
    const endDateIso = endTimestamp.split('T')[0];
    
    Logger.log(`Setting API date range: ${startDateIso} to ${endDateIso}`);
    setReportDate_(cfg.rescueBase, cookie, startDateIso, endDateIso);
    // Some environments reset report type after date/time changes: re-assert LISTALL
    setReportTypeListAll_(cfg.rescueBase, cookie);
    
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
    // Verify report type again before querying
    try {
      const rt2 = getReportType_(cfg.rescueBase, cookie);
      if (rt2 !== 'LISTALL') {
        Logger.log(`Report type before query was ${rt2}, re-applying LISTALL`);
        setReportTypeListAll_(cfg.rescueBase, cookie);
      }
    } catch (e) {
      Logger.log('getReportType pre-query check failed (non-fatal): ' + e.toString());
    }
    let allMappedRows = [];
    for (const nr of noderefs) {
      for (const node of nodes) {
        try {
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          if (!t) continue;
          const parseResult = parsePipe_(t, '|');
          const parsed = parseResult.rows || [];
          if (!parsed || !parsed.length) continue;
          // Extract date strings (YYYY-MM-DD) from timestamps for comparison
          // startTimestamp format: YYYY-MM-DDTHH:MM:SS.sssZ
          const startDateStr = startTimestamp.split('T')[0];
          const endDateStr = endTimestamp.split('T')[0];
          
          Logger.log(`Filtering sessions: date range=${startDateStr} to ${endDateStr} (strict: only dates within this range)`);
          
          // Trust the API date/time scope (we set date+time range above) to avoid timezone edge cases
          const mapped = parsed.map(mapRow_).filter(r => r && r.session_id);
          
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
  const perfMap = getPerfSummaryCached_(cfg, pullStartDate, pullEndDate);
  refreshAnalyticsDashboard_(pullStartDate, pullEndDate, perfMap);
      createDailySummarySheet_(ss, pullStartDate, pullEndDate);
      createSupportDataSheet_(ss, pullStartDate, pullEndDate);
  generateTechnicianTabs_(pullStartDate, pullEndDate, perfMap);
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
          // Accept XML or TEXT; getReportTry_ already validated acceptable shapes
          if (!t) continue;
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

  /* ===== Performance Raw Dump ===== */
  function dumpPerformanceRawMenu() {
    try {
      const ss = SpreadsheetApp.getActive();
      const configSheet = ss.getSheetByName('Dashboard_Config');
      const timeFrame = configSheet ? (configSheet.getRange('B3').getValue() || 'Today') : 'Today';
      const range = getTimeFrameRange_(timeFrame);
      dumpPerformanceRaw_(range.startDate, range.endDate);
      SpreadsheetApp.getActive().toast('Performance_Raw updated for ' + timeFrame, 4);
    } catch (e) {
      SpreadsheetApp.getActive().toast('Raw dump error: ' + e.toString().substring(0, 50));
      Logger.log('dumpPerformanceRawMenu error: ' + e.toString());
    }
  }

  function dumpPerformanceRaw_(startDate, endDate) {
    try {
      const ss = SpreadsheetApp.getActive();
      const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
      if (!sessionsSheet) return;
      const dataRange = sessionsSheet.getDataRange();
      if (dataRange.getNumRows() <= 1) return;

      // Read data
      const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
      const values = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();

      // Header index resolver with robust matching
      const idx = (variants) => {
        for (const v of variants) {
          const i = headers.findIndex(h => String(h || '').toLowerCase().trim() === String(v).toLowerCase().trim());
          if (i >= 0) return i;
        }
        return -1;
      };
      const startIdx = idx(['Start Time','start_time']);
      const techIdx = idx(['Technician Name','technician_name']);
      const sesIdx = idx(['Session ID','session_id']);
      const statusIdx = idx(['Status','session_status']);
      const chanIdx = idx(['Channel Name','channel_name']);
      const totalIdx = idx(['Total Time','duration_total_seconds','total_time']);
      const activeIdx = idx(['Active Time','duration_active_seconds','active_time']);
      const workIdx = idx(['Work Time','duration_work_seconds','work_time']);
      const waitingIdx = idx(['Waiting Time','pickup_seconds','waiting_time']);
      const connectingIdx = idx(['Connecting Time','connecting_time']);
      const holdIdx = idx(['Hold Time','hold_time']);
      const transferIdx = idx(['Time in Transfer','time_in_transfer']);
      const rebootIdx = idx(['Rebooting Time','rebooting_time']);
      const reconnectIdx = idx(['Reconnecting Time','reconnecting_time']);

      // Filter by date range
      const startStr = startDate.toISOString().split('T')[0];
      const endStr = endDate.toISOString().split('T')[0];
      const filtered = values.filter(r => {
        const d = r[startIdx];
        if (!d) return false;
        try {
          const ds = new Date(d).toISOString().split('T')[0];
          return ds >= startStr && ds <= endStr;
        } catch (e) { return false; }
      });

      // Create/clear output sheet
      let out = ss.getSheetByName('Performance_Raw');
      if (!out) out = ss.insertSheet('Performance_Raw');
      out.clear();

      const outHeaders = [
        'Date','Technician Name','Session ID','Status','Channel Name',
        'Total Time (sec)','Active Time (sec)','Work Time (sec)','Waiting Time (sec)',
        'Connecting (sec)','Hold (sec)','Transfer (sec)','Reboot (sec)','Reconnect (sec)'
      ];
      out.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]).setFontWeight('bold').setBackground('#E5E7EB');

      const rows = filtered.map(r => {
        const ds = new Date(r[startIdx]).toISOString().split('T')[0];
        const tName = r[techIdx] || '';
        const sid = r[sesIdx] || '';
        const status = r[statusIdx] || '';
        const channel = r[chanIdx] || '';
        const tot = parseDurationSeconds_(r[totalIdx] || 0);
        const act = parseDurationSeconds_(r[activeIdx] || 0);
        const work = parseDurationSeconds_(r[workIdx] || 0);
        const wait = parseDurationSeconds_(r[waitingIdx] || 0);
        const conn = parseDurationSeconds_(r[connectingIdx] || 0);
        const hold = parseDurationSeconds_(r[holdIdx] || 0);
        const trans = parseDurationSeconds_(r[transferIdx] || 0);
        const reb = parseDurationSeconds_(r[rebootIdx] || 0);
        const rec = parseDurationSeconds_(r[reconnectIdx] || 0);
        return [ds, tName, sid, status, channel, tot, act, work, wait, conn, hold, trans, reb, rec];
      });
      if (rows.length) {
        out.getRange(2, 1, rows.length, outHeaders.length).setValues(rows);
      }
      // Autosize a bit
      try { out.autoResizeColumns(1, outHeaders.length); } catch (e) {}
    } catch (e) {
      Logger.log('dumpPerformanceRaw_ error: ' + e.toString());
    }
  }
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
  const noderefs = getSummaryNoderefs_();
    
    for (const nr of noderefs) {
      for (const node of nodes) {
        try {
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          // Accept XML or TEXT; getReportTry_ already validated acceptable shapes
          if (!t) continue;
          
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

/* ===== Digium Map (raw, no formatting) ===== */
function dumpDigiumMapMenu() {
  try {
    const ss = SpreadsheetApp.getActive();
    const configSheet = ss.getSheetByName('Dashboard_Config');
    const timeFrame = configSheet ? (configSheet.getRange('B3').getValue() || 'Today') : 'Today';
    const range = getTimeFrameRange_(timeFrame);
    dumpDigiumMap_(range.startDate, range.endDate);
    SpreadsheetApp.getActive().toast('Digium map written for ' + timeFrame, 4);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Digium map error: ' + e.toString().substring(0, 50));
    Logger.log('dumpDigiumMapMenu error: ' + e.toString());
  }
}

function dumpDigiumMap_(startDate, endDate) {
  try {
    const cfg = getCfg_();
    const host = cfg.digiumHost;
    const user = cfg.digiumUser;
    const pass = cfg.digiumPass;
    if (!host || !user || !pass) {
      SpreadsheetApp.getActive().toast('Digium credentials missing');
      return;
    }

    // Build minimal parameters; omit breakdown for raw rows; include a broad set of report_fields if required
    const fmt = (d) => {
      const dt = (d instanceof Date) ? d : new Date(d);
      const Y = dt.getFullYear();
      const M = String(dt.getMonth()+1).padStart(2,'0');
      const D = String(dt.getDate()).padStart(2,'0');
      return `${Y}-${M}-${D}`;
    };
    const startStr = `${fmt(startDate)} 00:00:00`;
    const endStr = `${fmt(endDate)} 23:59:59`;

    // Try first without <breakdown> to get raw row listing
    let paramsXml = `\n      <start_date>${xmlEscape_(startStr)}</start_date>\n      <end_date>${xmlEscape_(endStr)}</end_date>\n      <ignore_weekends>0</ignore_weekends>\n      <format>xml</format>\n    `;
    let r = digiumApiCall_('switchvox.callReports.search', paramsXml, user, pass, host);
    if (!r.ok || !r.xml) {
      // Fallback: include a safe breakdown and common fields so we get structured rows
      const fields = ['total_calls','total_incoming_calls','total_outgoing_calls','talking_duration','call_duration','avg_talking_duration','avg_call_duration'];
      let reportFieldsXml = '<report_fields>' + fields.map(f=>`<report_field>${xmlEscape_(f)}</report_field>`).join('') + '</report_fields>';
      paramsXml = `\n      <start_date>${xmlEscape_(startStr)}</start_date>\n      <end_date>${xmlEscape_(endStr)}</end_date>\n      <ignore_weekends>0</ignore_weekends>\n      <breakdown>by_day</breakdown>\n      ${reportFieldsXml}\n      <format>xml</format>\n    `;
      r = digiumApiCall_('switchvox.callReports.search', paramsXml, user, pass, host);
    }

    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('map');
    if (!sh) sh = ss.insertSheet('map');
    sh.clear();

    if (!r.ok || !r.xml) {
      // Write raw XML so you can inspect it
      sh.getRange(1,1).setValue('Raw XML (unparsed)');
      sh.getRange(2,1).setValue(r.raw || r.error || 'no response');
      return;
    }

    // Parse generically: look for <rows><row .../> structure first
    const doc = r.xml;
    const root = doc.getRootElement();
    let rowsEl = null;
    // find first element named 'rows' in tree
    const findRows = (el) => {
      if (!el) return null;
      if (String(el.getName()).toLowerCase() === 'rows') return el;
      const kids = el.getChildren();
      for (let i=0;i<kids.length;i++) { const res = findRows(kids[i]); if (res) return res; }
      return null;
    };
    rowsEl = findRows(root);

    if (!rowsEl) {
      sh.getRange(1,1).setValue('No <rows> element found; writing raw XML');
      sh.getRange(2,1).setValue(r.raw || '');
      return;
    }

    const rowEls = rowsEl.getChildren('row');
    if (!rowEls || !rowEls.length) {
      // Maybe attributes exist directly on <rows>
      const attrs = rowsEl.getAttributes ? rowsEl.getAttributes().map(a => a.getName()) : [];
      if (attrs && attrs.length) {
        const header = attrs;
        sh.getRange(1,1,1,header.length).setValues([header]);
        const vals = [rowsEl.getAttributes().map(a => a.getValue())];
        sh.getRange(2,1,1,header.length).setValues(vals);
        return;
      }
      sh.getRange(1,1).setValue('Empty <rows>');
      return;
    }

    // Build header from union of all attribute names and first-level child element names
    const attrSet = new Set();
    const childSet = new Set();
    rowEls.forEach(re => {
      (re.getAttributes() || []).forEach(a => attrSet.add(a.getName()));
      (re.getChildren() || []).forEach(c => childSet.add(c.getName()));
    });
    const header = Array.from(attrSet).concat(Array.from(childSet));
    if (!header.length) {
      sh.getRange(1,1).setValue('Rows found but no attributes/children; writing raw XML');
      sh.getRange(2,1).setValue(r.raw || '');
      return;
    }
    sh.getRange(1,1,1,header.length).setValues([header]);

    const outRows = rowEls.map(re => {
      const rowVals = [];
      const attrMap = {};
      (re.getAttributes() || []).forEach(a => { attrMap[a.getName()] = a.getValue(); });
      const childMap = {};
      (re.getChildren() || []).forEach(c => { childMap[c.getName()] = c.getText ? c.getText().trim() : ''; });
      header.forEach(h => {
        if (attrMap.hasOwnProperty(h)) rowVals.push(attrMap[h]);
        else if (childMap.hasOwnProperty(h)) rowVals.push(childMap[h]);
        else rowVals.push('');
      });
      return rowVals;
    });
    if (outRows.length) sh.getRange(2,1,outRows.length,header.length).setValues(outRows);
    // No styling or links per request
  } catch (e) {
    Logger.log('dumpDigiumMap_ error: ' + e.toString());
  }
}

/* ===== Rescue LISTALL Map (raw, no formatting) ===== */
function dumpRescueListAllRawMenu() {
  try {
    const ss = SpreadsheetApp.getActive();
    const configSheet = ss.getSheetByName('Dashboard_Config');
    const timeFrame = configSheet ? (configSheet.getRange('B3').getValue() || 'Today') : 'Today';
    const range = getTimeFrameRange_(timeFrame);
    dumpRescueListAllRaw_(range.startDate, range.endDate);
    SpreadsheetApp.getActive().toast('Rescue LISTALL map written for ' + timeFrame, 4);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Rescue map error: ' + e.toString().substring(0, 50));
    Logger.log('dumpRescueListAllRawMenu error: ' + e.toString());
  }
}

function dumpRescueListAllRaw_(startDate, endDate) {
  try {
    const cfg = getCfg_();
    const base = cfg.rescueBase;
    const cookie = login_(base, cfg.user, cfg.pass);
    // Prepare LISTALL text pipe output
    setReportAreaSession_(base, cookie);
    setReportTypeListAll_(base, cookie);
    setOutputXMLOrFallback_(base, cookie); // will set TEXT if FORCE_TEXT_OUTPUT
    setDelimiter_(base, cookie, '|');
    setReportDate_(base, cookie, isoDate_(startDate), isoDate_(endDate));
    setReportTimeAllDay_(base, cookie);

    const nodes = (cfg.nodes || []).map(n => Number(n)).filter(n => Number.isFinite(n));
    const noderefs = ['NODE','CHANNEL'];

    // Collect rows and union headers across nodes/noderefs
    let headerUnion = [];
    const rows = [];
    const addHeaderUnion = (hs) => {
      hs.forEach(h => { if (headerUnion.indexOf(h) === -1) headerUnion.push(h); });
    };
    for (let i = 0; i < noderefs.length; i++) {
      const nr = noderefs[i];
      for (let j = 0; j < nodes.length; j++) {
        const node = nodes[j];
        try {
          const t = getReportTry_(base, cookie, node, nr);
          // Accept XML or TEXT; getReportTry_ already validated acceptable shapes
          if (!t) continue;
          const parsed = parsePipe_(t, '|');
          if (!parsed || !parsed.headers || !parsed.rows) continue;
          if (parsed.headers && parsed.headers.length) addHeaderUnion(parsed.headers);
          parsed.rows.forEach(r => rows.push(r));
          Utilities.sleep(150);
        } catch (e) { Logger.log('dumpRescueListAllRaw_ fetch error: ' + e.toString()); }
      }
    }

    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Rescue_Map');
    if (!sh) sh = ss.insertSheet('Rescue_Map');
    sh.clear();

    if (!headerUnion.length) {
      sh.getRange(1,1).setValue('No LISTALL data returned for selected range');
      return;
    }

    // Write headers
    sh.getRange(1,1,1,headerUnion.length).setValues([headerUnion]);

    // Write rows preserving header order; avoid formatting or formulas
    const out = rows.map(obj => headerUnion.map(h => (obj && obj.hasOwnProperty(h)) ? obj[h] : ''));
    if (out.length) {
      // Write in chunks to avoid size limits
      const chunk = 1000;
      for (let i = 0; i < out.length; i += chunk) {
        const part = out.slice(i, i + chunk);
        sh.getRange(2 + i, 1, part.length, headerUnion.length).setValues(part);
      }
    }
    // No styling per request
  } catch (e) {
    Logger.log('dumpRescueListAllRaw_ error: ' + e.toString());
  }
}

/* ===== One-time header fix for Rescue_Map ===== */
// removed: header-fix utilities, strip formatting, and re-ingest options per request

/* ===== Shift Rescue_Map data left so that values from 'Session ID' align under 'Technician ID' (first 4 columns unchanged) ===== */
function shiftRescueMapDataAlignMenu() {
  try {
    const ok = shiftRescueMapDataAlign_();
    SpreadsheetApp.getActive().toast(ok ? 'Rescue_Map data shifted left (Session ID → Technician ID)' : 'Rescue_Map not found or nothing to shift');
  } catch (e) {
    SpreadsheetApp.getActive().toast('Shift error: ' + e.toString().substring(0, 60));
    Logger.log('shiftRescueMapDataAlignMenu error: ' + e.toString());
  }
}

function shiftRescueMapDataAlign_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('sessions');
  if (!sh) return false;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1 || lastCol < 6) return false; // need headers + enough cols

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const findIdx = (vals, target) => vals.findIndex(h => String(h || '').trim().toLowerCase() === String(target).trim().toLowerCase());
  const techIdIdx0 = findIdx(headers, 'Technician ID'); // 0-based
  const sessionIdIdx0 = findIdx(headers, 'Session ID'); // 0-based
  if (techIdIdx0 < 0 || sessionIdIdx0 < 0) return false;
  if (sessionIdIdx0 <= techIdIdx0) return false; // nothing to shift

  const delta = sessionIdIdx0 - techIdIdx0; // number of positions to shift left

  // Read data block starting at Technician ID through end
  const startCol = techIdIdx0 + 1; // 1-based
  const width = lastCol - techIdIdx0;
  const height = lastRow - 1;
  if (width <= 0 || height <= 0) return false;
  const range = sh.getRange(2, startCol, height, width);
  const data = range.getValues();

  // Shift each row left by 'delta' positions within this block
  const shifted = data.map(row => {
    const out = new Array(width);
    for (let c = 0; c < width; c++) {
      const src = c + delta;
      out[c] = (src < width) ? row[src] : '';
    }
    return out;
  });

  range.setValues(shifted);
  return true;
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
    
  // Set up for summary report (prefer Performance area for channel metrics)
  // Performance area tends to return more robust channel-level metrics when used with SUMMARY.
  // If Performance area is not supported, the helper will fallback to Session area internally.
  setReportAreaPerformance_(cfg.rescueBase, cookie);
    
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
  // Small delay to give the Rescue API session state a moment to settle in case the server
  // returns previous state briefly (helps avoid intermittent LISTALL responses).
  try { Utilities.sleep(250); } catch (e) { /* ignore in non-runtime tests */ }
    
    // Verify report type was set correctly using getReportType (non-blocking)
    try {
      const currentReportType = getReportType_(cfg.rescueBase, cookie);
      Logger.log(`Current report type after setting (channel summary): ${currentReportType}`);
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
  // Accept XML or TEXT; getReportTry_ already validated acceptable shapes
  if (!t) continue;
        
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
    // Header resolver that supports both normalized keys and official API titles
    const getHeaderIndex = (variants) => {
      for (const v of variants) {
        const i = headers.findIndex(h => String(h || '').toLowerCase().trim() === String(v).toLowerCase().trim());
        if (i >= 0) return i;
      }
      return -1;
    };
    
    const startIdx = getHeaderIndex(['start_time','Start Time']);
    const techIdx = getHeaderIndex(['technician_name','Technician Name']);
    const statusIdx = getHeaderIndex(['session_status','Status']);
    const durationIdx = getHeaderIndex(['duration_total_seconds','Total Time']);
    const workIdx = getHeaderIndex(['duration_work_seconds','Work Time']);
    const pickupIdx = getHeaderIndex(['pickup_seconds','Waiting Time']);
    
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
      if (row[workIdx]) dailyData[dateStr].totalWorkSeconds += parseDurationSeconds_(row[workIdx]);
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
    
    // Total duration/Work Time per day (store as numeric time fractions)
    const totalDurationRow = ['Total duration/ Total Work Time'];
    let totalWorkSecondsAll = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.totalWorkSeconds > 0) {
        totalDurationRow.push(data.totalWorkSeconds / 86400);
        totalWorkSecondsAll += data.totalWorkSeconds;
      } else {
        totalDurationRow.push(0);
      }
    });
    totalDurationRow.push(totalWorkSecondsAll / 86400);
    summaryRows.push(totalDurationRow);
    
    // Avg Session duration per day (numeric time fractions)
    const avgSessionRow = ['Avg Session'];
    let totalAvgSeconds = 0;
    let daysWithData = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.sessions.length > 0) {
        const durations = data.sessions.map(s => parseDurationSeconds_(s[durationIdx] || 0)).filter(Boolean);
        if (durations.length > 0) {
          const avg = durations.reduce((a, b) => a + b, 0) / durations.length;
          avgSessionRow.push(avg / 86400);
          totalAvgSeconds += avg;
          daysWithData++;
        } else {
          avgSessionRow.push(0);
        }
      } else {
        avgSessionRow.push(0);
      }
    });
    if (daysWithData > 0) {
      const overallAvg = totalAvgSeconds / daysWithData;
      avgSessionRow.push(overallAvg / 86400);
    } else {
      avgSessionRow.push(0);
    }
    summaryRows.push(avgSessionRow);
    
    // Avg Pick-up Speed per day (numeric time fractions)
    const avgPickupRow = ['Avg Pick-up Speed'];
    let totalPickupSeconds = 0;
    let totalPickupCount = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.sessions.length > 0) {
        const pickups = data.sessions.map(s => parseDurationSeconds_(s[pickupIdx] || 0)).filter(p => p > 0);
        if (pickups.length > 0) {
          const avg = pickups.reduce((a, b) => a + b, 0) / pickups.length;
          avgPickupRow.push(avg / 86400);
          totalPickupSeconds += avg * pickups.length;
          totalPickupCount += pickups.length;
        } else {
          avgPickupRow.push(0);
        }
      } else {
        avgPickupRow.push(0);
      }
    });
    if (totalPickupCount > 0) {
      const overallAvg = totalPickupSeconds / totalPickupCount;
      avgPickupRow.push(overallAvg / 86400);
    } else {
      avgPickupRow.push(0);
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
  dates.forEach(() => avgRealRow.push(0));
  avgRealRow.push(0);
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
  if (row[durationIdx]) techStats[tech].durations.push(parseDurationSeconds_(row[durationIdx]));
  if (row[pickupIdx]) techStats[tech].pickups.push(parseDurationSeconds_(row[pickupIdx]));
  if (row[workIdx]) techStats[tech].workSeconds += parseDurationSeconds_(row[workIdx]);
    });
    
    Object.keys(techStats).sort((a, b) => techStats[b].sessions - techStats[a].sessions).forEach(tech => {
      const stats = techStats[tech];
      // pct as numeric fraction (0..1)
      const pct = totalSessionsAll > 0 ? (stats.sessions / totalSessionsAll) : 0;
      // sessions per hour numeric
      const sessionsPerHr = stats.workSeconds > 0 ? (stats.sessions / (stats.workSeconds / 3600)) : 0;
      const avgPickup = stats.pickups.length > 0 ? (stats.pickups.reduce((a, b) => a + b, 0) / stats.pickups.length) : 0;
      const avgDur = stats.durations.length > 0 ? (stats.durations.reduce((a, b) => a + b, 0) / stats.durations.length) : 0;
      const workTime = stats.workSeconds || 0;
      // store pickups/durations/workTime as numeric time fractions (seconds/86400)
      const techRow = [tech, stats.sessions, pct, sessionsPerHr, avgPickup / 86400, avgDur / 86400, workTime / 86400, ...Array(dates.length - 6).fill('')];
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
      try { applyProfessionalTableStyling_(summarySheet, summaryRows[0].length); } catch (e) { Logger.log('Styling daily summary failed: ' + e.toString()); }

      // Post-formatting: ensure duration/pickup cells are numeric time fractions and formatted as hh:mm:ss
      try {
        const lastCol = summaryRows[0].length;
        // load column A to find metric rows
        const colA = summarySheet.getRange(1,1,summaryRows.length,1).getValues().map(r => String((r[0]||'')).trim().toLowerCase());
        const findRow = (label) => colA.findIndex(v => v === label.toLowerCase());

        const totalSessionsRowIdx = findRow('total sessions');
        if (totalSessionsRowIdx >= 0) {
          summarySheet.getRange(totalSessionsRowIdx + 1, 2, 1, lastCol - 1).setNumberFormat('0');
        }

        const totalDurationRowIdx = findRow('total duration/ total work time');
        if (totalDurationRowIdx >= 0) {
          summarySheet.getRange(totalDurationRowIdx + 1, 2, 1, lastCol - 1).setNumberFormat('hh:mm:ss');
        }

        const avgSessionRowIdx = findRow('avg session');
        if (avgSessionRowIdx >= 0) {
          summarySheet.getRange(avgSessionRowIdx + 1, 2, 1, lastCol - 1).setNumberFormat('hh:mm:ss');
        }

        const avgPickupRowIdx = findRow('avg pick-up speed');
        if (avgPickupRowIdx >= 0) {
          summarySheet.getRange(avgPickupRowIdx + 1, 2, 1, lastCol - 1).setNumberFormat('hh:mm:ss');
        }

        const avgRealRowIdx = findRow('average real');
        if (avgRealRowIdx >= 0) {
          summarySheet.getRange(avgRealRowIdx + 1, 2, 1, lastCol - 1).setNumberFormat('hh:mm:ss');
        }

        // Technician summary formatting: find the header 'Technician Name' and format columns for tech rows
        const techHeaderIdx = colA.findIndex(v => v.indexOf('technician name') === 0 || v.indexOf('technician') === 0);
        if (techHeaderIdx >= 0) {
          const techDataStartRow = techHeaderIdx + 2; // data starts after header row
          const techDataRows = Object.keys(techStats).length;
          if (techDataRows > 0) {
            const techStartCol = 2; // Total Sessions column
            // Total Sessions integer
            summarySheet.getRange(techDataStartRow, techStartCol, techDataRows, 1).setNumberFormat('0');
            // % Of Total sessions -> percent
            summarySheet.getRange(techDataStartRow, techStartCol + 1, techDataRows, 1).setNumberFormat('0.0%');
            // Sessions per HR -> one decimal
            summarySheet.getRange(techDataStartRow, techStartCol + 2, techDataRows, 1).setNumberFormat('0.0');
            // Avg Pick-up Speed, Avg Duration, Average Work Time -> hh:mm:ss (cols 5,6,7)
            summarySheet.getRange(techDataStartRow, techStartCol + 3, techDataRows, 3).setNumberFormat('hh:mm:ss');
          }
        }
      } catch (e) { Logger.log('Daily summary post-formatting failed: ' + e.toString()); }
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
    const getHeaderIndex = (variants) => {
      for (const v of variants) {
        const i = headers.findIndex(h => String(h || '').toLowerCase().trim() === String(v).toLowerCase().trim());
        if (i >= 0) return i;
      }
      return -1;
    };
    
    const startIdx = getHeaderIndex(['start_time','Start Time']);
    const techIdx = getHeaderIndex(['technician_name','Technician Name']);
    const statusIdx = getHeaderIndex(['session_status','Status']);
    const durationIdx = getHeaderIndex(['duration_total_seconds','Total Time']);
    const workIdx = getHeaderIndex(['duration_work_seconds','Work Time']);
    const pickupIdx = getHeaderIndex(['pickup_seconds','Waiting Time']);
    const customerIdx = getHeaderIndex(['customer_name','Your Name:','Customer Name']);
    const sessionIdIdx = getHeaderIndex(['session_id','Session ID']);
    const channelIdx = getHeaderIndex(['channel_name','Channel Name']);
    const resolvedIdx = getHeaderIndex(['resolved_unresolved','Resolved Unresolved']);
    const callingCardIdx = getHeaderIndex(['calling_card','Calling Card']);
    
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
  supportSheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#E61C37');
    supportSheet.getRange(1, 1, 1, 8).merge();
    
    supportSheet.getRange(2, 1).setValue('Date Range:');
    supportSheet.getRange(2, 2).setValue(`${startStr} to ${endStr}`);
    supportSheet.getRange(2, 1).setFontWeight('bold');
    
    // Summary KPIs
    const kpiRow = 4;
    const totalSessions = filtered.length;
    const totalWorkSeconds = filtered.reduce((sum, row) => sum + parseDurationSeconds_(row[workIdx] || 0), 0);
    const totalWorkHours = (totalWorkSeconds / 3600).toFixed(1);
    const avgPickup = filtered.length > 0 ? 
      Math.round(filtered.reduce((sum, row) => sum + parseDurationSeconds_(row[pickupIdx] || 0), 0) / filtered.length) : 0;
  // resolved/unresolved columns exist in Sessions but Support_Data will not report resolution metrics
  const resolvedCount = filtered.filter(row => row[resolvedIdx] === 'Resolved').length;
    
    // Count Nova Wave sessions (calling card contains "Nova wave chat")
    const novaWaveCount = callingCardIdx >= 0 ? 
      filtered.filter(row => {
        const callingCard = String(row[callingCardIdx] || '').toLowerCase();
        return callingCard.includes('nova wave chat');
      }).length : 0;
    
    // Support_Data top KPIs: keep totals and timing metrics only (no resolution metrics)
    // Use numeric time fraction for Avg Pickup so it can be formatted as hh:mm:ss
    const kpis = [
      ['Total Sessions', totalSessions],
      ['Nova Wave Sessions', novaWaveCount],
      ['Total Work Hours', Number(totalWorkHours)],
      ['Avg Pickup Time', avgPickup / 86400]
    ];
    
    for (let i = 0; i < kpis.length; i++) {
      const row = kpiRow + Math.floor(i / 3);
      const col = (i % 3) * 3 + 1;
      supportSheet.getRange(row, col).setValue(kpis[i][0]);
      supportSheet.getRange(row, col).setFontSize(11).setFontColor('#666666');
      supportSheet.getRange(row, col + 1).setValue(kpis[i][1]);
  supportSheet.getRange(row, col + 1).setFontSize(14).setFontWeight('bold').setFontColor('#E61C37');
      supportSheet.getRange(row, col, 1, 2).setBorder(true, true, true, true, true, true);
    }
    
    // Technician Performance Table
    const tableRow = kpiRow + 3;
    supportSheet.getRange(tableRow, 1).setValue('Technician Performance');
  supportSheet.getRange(tableRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#E61C37');
  // Technician performance table (exclude resolution rate column)
  const tableHeaders = ['Technician', 'Sessions', 'Nova Wave Sessions', 'Avg Pickup', 'Avg Duration', 'Work Hours'];
  supportSheet.getRange(tableRow + 1, 1, 1, tableHeaders.length).setValues([tableHeaders]);
  supportSheet.getRange(tableRow + 1, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#07123B').setFontColor('#FFFFFF');
    
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
      if (row[durationIdx]) techStats[tech].durations.push(parseDurationSeconds_(row[durationIdx]));
      if (row[pickupIdx]) techStats[tech].pickups.push(parseDurationSeconds_(row[pickupIdx]));
      if (row[workIdx]) techStats[tech].workSeconds += parseDurationSeconds_(row[workIdx]);
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
      const avgPickup = stats.pickups.length > 0 ? (stats.pickups.reduce((a, b) => a + b, 0) / stats.pickups.length) : 0;
      const avgDur = stats.durations.length > 0 ? (stats.durations.reduce((a, b) => a + b, 0) / stats.durations.length) : 0;
      const workHoursNum = stats.workSeconds / 3600;
      // store average pickup and duration as numeric time fractions, work hours as numeric hours
      return [tech, stats.sessions, stats.novaWave, avgPickup / 86400, avgDur / 86400, workHoursNum];
    });
    
    if (techRows.length > 0) {
      supportSheet.getRange(tableRow + 2, 1, techRows.length, tableHeaders.length).setValues(techRows);
      // format avg pickup/duration as times and work hours as one-decimal number
      try {
        supportSheet.getRange(tableRow + 2, 4, techRows.length, 1).setNumberFormat('hh:mm:ss');
        supportSheet.getRange(tableRow + 2, 5, techRows.length, 1).setNumberFormat('hh:mm:ss');
        supportSheet.getRange(tableRow + 2, 6, techRows.length, 1).setNumberFormat('0.0');
      } catch (e) { /* ignore */ }
    }
    
    // Channel Performance Summary (daily wide format)
    // Build a wide table with rows for the requested KPIs and columns for each day in the range.
    const channelRow = tableRow + techRows.length + 4;
    supportSheet.getRange(channelRow, 1).setValue('Channel Performance Summary (daily)');
    supportSheet.getRange(channelRow, 1).setFontSize(14).setFontWeight('bold');

    // Group sessions by day from the already-filtered session data
    const dates = [];
    let currentDate = new Date(startDate);
    const endDateObj = new Date(endDate);
    while (currentDate <= endDateObj) {
      dates.push(currentDate.toISOString().split('T')[0]);
      currentDate.setDate(currentDate.getDate() + 1);
    }

    // Build a map of daily sessions
    const dailyData = {};
    filtered.forEach(row => {
      try {
        const d = new Date(row[startIdx]).toISOString().split('T')[0];
        if (!dailyData[d]) dailyData[d] = { sessions: [], totalWorkSeconds: 0 };
        dailyData[d].sessions.push(row);
        // Use robust parser for duration/work fields to avoid negative/invalid values
        dailyData[d].totalWorkSeconds += parseDurationSeconds_(row[workIdx] || 0);
      } catch (e) { /* ignore malformed rows */ }
    });

    // Header row: Metric | date... | Totals/Averages
    const headerRow = ['Metric'];
    dates.forEach(d => {
      const dObj = new Date(d + 'T00:00:00');
      headerRow.push(`${dObj.getMonth()+1}/${dObj.getDate()}/${dObj.getFullYear()}`);
    });
    headerRow.push('Totals/Averages');

    // Total sessions per day
    const totalSessionsRow = ['Total sessions'];
    let totalSessionsAll = 0;
    dates.forEach(d => {
      const count = dailyData[d] ? dailyData[d].sessions.length : 0;
      totalSessionsRow.push(count);
      totalSessionsAll += count;
    });
    totalSessionsRow.push(totalSessionsAll);

    // Total Work Time per day (as time fractions - numeric values so sheet formats stay correct)
    const totalWorkRow = ['Total Work Time'];
    let totalWorkSecondsAll = 0;
    dates.forEach(d => {
      const secs = dailyData[d] ? dailyData[d].totalWorkSeconds : 0;
      totalWorkSecondsAll += secs;
      totalWorkRow.push((secs || 0) / 86400); // store as fraction of day
    });
    totalWorkRow.push(totalWorkSecondsAll / 86400);

    // Avg Session per day (average of duration_total_seconds)
    const avgSessionRow = ['Avg Session'];
    let totalAvgSeconds = 0;
    let daysWithData = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.sessions.length > 0) {
        // Parse session durations defensively (supports seconds, HH:MM:SS, or day-fractions)
        const durations = data.sessions.map(s => parseDurationSeconds_(s[durationIdx] || 0)).filter(Boolean);
          if (durations.length > 0) {
            const avg = durations.reduce((a,b) => a+b, 0) / durations.length;
            avgSessionRow.push(avg / 86400);
            totalAvgSeconds += avg;
            daysWithData++;
          } else {
            avgSessionRow.push(0);
          }
      } else {
        avgSessionRow.push(0);
      }
    });
    if (daysWithData > 0) {
      const overallAvg = totalAvgSeconds / daysWithData;
      avgSessionRow.push(overallAvg / 86400);
    } else {
      avgSessionRow.push(0);
    }

    // Avg Pick-up Speed per day
    const avgPickupRow = ['Avg Pick-up Speed'];
    let totalPickupSeconds = 0;
    let totalPickupCount = 0;
    dates.forEach(d => {
      const data = dailyData[d];
      if (data && data.sessions.length > 0) {
        // Parse pickup times defensively
        const pickups = data.sessions.map(s => parseDurationSeconds_(s[pickupIdx] || 0)).filter(p => p > 0);
        if (pickups.length > 0) {
          const avg = pickups.reduce((a,b) => a+b, 0) / pickups.length;
          avgPickupRow.push(avg / 86400);
          totalPickupSeconds += avg * pickups.length;
          totalPickupCount += pickups.length;
        } else {
          avgPickupRow.push(0);
        }
      } else {
        avgPickupRow.push(0);
      }
    });
    if (totalPickupCount > 0) {
      const overallAvg = totalPickupSeconds / totalPickupCount;
      avgPickupRow.push(overallAvg / 86400);
    } else {
      avgPickupRow.push(0);
    }

    // Write table to sheet
    const wideRows = [headerRow, totalSessionsRow, totalWorkRow, avgSessionRow, avgPickupRow];
    supportSheet.getRange(channelRow + 1, 1, wideRows.length, wideRows[0].length).setValues(wideRows);
    // Make sure header is visible: use light background and dark font so titles aren't hidden
    supportSheet.getRange(channelRow + 1, 1, 1, wideRows[0].length).setFontWeight('bold').setBackground('#E5E7EB').setFontColor('#000000');
    // Explicit number formatting per KPI row so integers remain integers and times remain times
    const numCols = wideRows[0].length;
    try {
      // Total sessions row (row index channelRow+2): integers
      supportSheet.getRange(channelRow + 2, 2, 1, numCols - 1).setNumberFormat('0');
      // Total work row (time fractions)
      supportSheet.getRange(channelRow + 3, 2, 1, numCols - 1).setNumberFormat('hh:mm:ss');
      // Avg session row (time fractions)
      supportSheet.getRange(channelRow + 4, 2, 1, numCols - 1).setNumberFormat('hh:mm:ss');
      // Avg pickup row (time fractions)
      supportSheet.getRange(channelRow + 5, 2, 1, numCols - 1).setNumberFormat('hh:mm:ss');
    } catch (e) { Logger.log('Formatting support data wide table failed: ' + e.toString()); }
    
    // Formatting
    supportSheet.setColumnWidth(1, 200);
    supportSheet.setColumnWidth(2, 120);
    supportSheet.setColumnWidth(3, 120);
    supportSheet.setColumnWidth(4, 120);
    supportSheet.setColumnWidth(5, 120);
    supportSheet.setColumnWidth(6, 120);
    supportSheet.setFrozenRows(1);
  try { applyProfessionalTableStyling_(supportSheet, wideRows[0].length); } catch (e) { Logger.log('Styling support sheet failed: ' + e.toString()); }
    
    Logger.log('Support Data sheet created');
    // --- Insert Digium call report (wide-format) into Support_Data under Channel Performance Summary
    // Place starting at row 25, column A as requested
    try {
      const digRes = fetchDigiumCallReports_(startDate, endDate, {});
      if (digRes && digRes.ok && digRes.dates && digRes.rows) {
        const startCol = 1; // Column A
        const startRow = 25; // Under the Channel Performance Summary (daily)
        // Build header (Metric + dates)
        const header = ['Metric'].concat(digRes.dates || []);
        supportSheet.getRange(startRow, startCol, 1, header.length).setValues([header]);
        // Write metric rows
        if (digRes.rows && digRes.rows.length) {
          supportSheet.getRange(startRow + 1, startCol, digRes.rows.length, digRes.rows[0].length).setValues(digRes.rows);
        }
        // Styling
        try { supportSheet.getRange(startRow, startCol, 1, header.length).setFontWeight('bold').setBackground('#E5E7EB').setFontColor('#000000'); } catch (e) {}
        // Leave numbers raw (seconds/counts). No conversion here.
        try { supportSheet.setColumnWidth(startCol, 140); } catch (e) {}
        SpreadsheetApp.flush();
      } else {
        Logger.log('No Digium parsed data available for Support_Data: ' + (digRes && (digRes.error || digRes.raw) ? (digRes.error || digRes.raw).toString().substring(0,200) : 'no response'));
      }
    } catch (e) { Logger.log('Failed to write Digium data into Support_Data: ' + e.toString()); }
    // Apply duration formatting across sheets now that Support_Data and Digium data are populated.
    try {
      if (typeof formatAllDurationColumns === 'function') formatAllDurationColumns();
    } catch (e) { Logger.log('formatAllDurationColumns failed after support data generation: ' + e.toString()); }
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
  sh.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#07123B');
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
    ['Total Sessions', '=COUNTA(Sessions!A2:A)'], // Column A is session_id
    ['Active Sessions', '=COUNTIFS(Sessions!C2:C, "Active")'], // Column C is session_status
    ['Nova Wave Sessions', ''], // Will be calculated dynamically in refreshAnalyticsDashboard_ based on time frame
    // Use numeric formulas returning time fractions (seconds/86400) instead of TEXT so cells remain numeric
    ['Avg Duration', '=IF(COUNT(Sessions!U2:U)>0, AVERAGE(Sessions!U2:U)/86400, 0)'], // Column U is duration_total_seconds
    ['Avg Pickup Time', '=IF(COUNT(Sessions!V2:V)>0, AVERAGE(Sessions!V2:V)/86400, 0)'], // Column V is pickup_seconds
    ['Longest Session', '=IF(MAX(Sessions!U2:U)>0, MAX(Sessions!U2:U)/86400, 0)'],
    // SLA as numeric fraction (0..1) so we can apply a percentage format
    ['SLA Hit %', '=IF(COUNT(Sessions!V2:V)>0, COUNTIFS(Sessions!V2:V, "<=60")/COUNT(Sessions!V2:V), 0)'],
    ['Avg Sessions/Hour', '=IF(COUNTA(Sessions!A2:A)>0, ROUND(COUNTA(Sessions!A2:A)/8, 1), 0)'] // Column A is session_id
  ];
  // Build polished KPI cards (3 cards per row). Each card is 3 columns wide × 3 rows tall.
  const cardWidth = 3;
  const cardHeight = 3;
  for (let i = 0; i < kpiCards.length; i++) {
    const cardTop = kpiRow + Math.floor(i / 3) * cardHeight;
    const cardLeft = (i % 3) * cardWidth + 1;
    const title = kpiCards[i][0];
    const formula = kpiCards[i][1];

  // Card header strip
  const hdrRange = sh.getRange(cardTop, cardLeft, 1, cardWidth);
  hdrRange.merge();
  hdrRange.setValue(title).setFontSize(10).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1A73E8').setHorizontalAlignment('left');

    // Card value area (prominent)
    const valRange = sh.getRange(cardTop + 1, cardLeft, 1, cardWidth);
    valRange.merge();
    if (i !== 2 && formula) {
      // place formula for dynamic KPIs
      valRange.setFormula(formula);
      // Apply appropriate number formats so numeric results display correctly
      try {
        if (/avg duration/i.test(title) || /avg pickup/i.test(title) || /longest session/i.test(title)) {
          valRange.setNumberFormat('hh:mm:ss');
        } else if (/sla hit/i.test(title)) {
          valRange.setNumberFormat('0.0%');
        } else if (/total sessions/i.test(title)) {
          valRange.setNumberFormat('0');
        } else if (/avg sessions\/hour/i.test(title)) {
          valRange.setNumberFormat('0.0');
        }
      } catch (e) { /* ignore formatting errors */ }
    } else {
      // placeholder value for Nova Wave Sessions (filled by refreshAnalyticsDashboard_)
      valRange.setValue('—');
    }
  valRange.setFontSize(20).setFontWeight('bold').setFontColor('#0F172A').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#FFFFFF');

    // Card footer / note (small muted text)
    const noteRange = sh.getRange(cardTop + 2, cardLeft, 1, cardWidth);
  noteRange.merge();
  noteRange.setValue(' ').setFontSize(9).setFontColor('#666666').setBackground('#F3F4F6');

    // Card border
    try { sh.getRange(cardTop, cardLeft, cardHeight, cardWidth).setBorder(true, true, true, true, true, true); } catch (e) {}
  }
  // Highlight specific KPI value rows for visual separation
  try {
    const cardCols = cardWidth * 3; // total columns spanned by KPI grid
    // Row 7 (kpiRow + 3) - give a blue background across KPI width
    const visualRow1 = kpiRow + 3;
    sh.getRange(visualRow1, 1, 1, cardCols).setBackground('#1A73E8').setFontColor('#FFFFFF');
    // Find SLA Hit % position (index 6) and color the value row below its header
    const slaIndex = kpiCards.findIndex(k => /SLA Hit %/i.test(k[0]));
    if (slaIndex >= 0) {
      const slaGroupTop = kpiRow + Math.floor(slaIndex / 3) * cardHeight;
      const slaValueRow = slaGroupTop + 1; // row below header
      sh.getRange(slaValueRow, 1, 1, cardCols).setBackground('#1A73E8').setFontColor('#FFFFFF');
    }
  } catch (e) { Logger.log('Failed to apply KPI separator rows: ' + e.toString()); }
  
  // Add Live Active Sessions section (current sessions)
  // Place live area below KPI cards
  const cardRows = Math.ceil(kpiCards.length / 3) * cardHeight;
  const liveRow = kpiRow + cardRows + 2;
  sh.getRange(liveRow, 1).setValue('🟢 LIVE ACTIVE SESSIONS');
  sh.getRange(liveRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#0F172A');
  sh.getRange(liveRow, 1, 1, 6).merge();
  const liveHeaders = ['Technician', 'Customer', 'Start Time', 'Live Duration', 'Channel', 'Session ID'];
  sh.getRange(liveRow + 1, 1, 1, liveHeaders.length).setValues([liveHeaders]);
  // clearly visible header - use green for live sessions
  sh.getRange(liveRow + 1, 1, 1, liveHeaders.length).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sh.getRange(liveRow + 1, 1, 1, liveHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
  // reserve space for live rows (approx)
  const liveSectionHeight = 12;
  const tableRow = liveRow + liveSectionHeight + 2;
  sh.getRange(tableRow, 1).setValue('Team Performance');
  sh.getRange(tableRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#0F172A');
  const headers = ['Technician', 'Total Sessions', 'Avg Duration', 'Avg Pickup', 'SLA Hit %', 'Sessions/Hour', 'Total Work Time'];
  sh.getRange(tableRow + 1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(tableRow + 1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
  sh.getRange(tableRow + 1, 1, 1, headers.length).setBorder(false, false, false, false, false, false); // Remove borders
  const activeRow = tableRow + 15;
  sh.getRange(activeRow, 1).setValue('Active Sessions (Selected Time Frame)');
  sh.getRange(activeRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#0F172A');
  const activeHeaders = ['Technician', 'Customer Name', 'Start Time', 'Duration', 'Session ID', 'Calling Card'];
  sh.getRange(activeRow + 1, 1, 1, activeHeaders.length).setValues([activeHeaders]);
  // Use blue header for clarity (white text on blue background)
  sh.getRange(activeRow + 1, 1, 1, activeHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
  sh.getRange(activeRow + 1, 1, 1, activeHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
  // Place Waiting Queue to the right of Live Active Sessions (row 16, starting at column J)
  const queueRow = 16;
  const queueCol = 9; // column I (moved one left)
  sh.getRange(queueRow, queueCol).setValue('⏳ Waiting Queue');
  sh.getRange(queueRow, queueCol).setFontSize(14).setFontWeight('bold').setFontColor('#0F172A');
  const queueHeaders = ['Channel', 'Customer', 'Waiting Since', 'Wait Duration'];
  sh.getRange(queueRow + 1, queueCol, 1, queueHeaders.length).setValues([queueHeaders]);
  sh.getRange(queueRow + 1, queueCol, 1, queueHeaders.length).setFontWeight('bold').setBackground('#EA8600').setFontColor('#FFFFFF');
  sh.getRange(queueRow + 1, queueCol, 1, queueHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 150);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 150);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 150);
  // Ensure waiting queue columns (J-M) are visible and wide enough
  try {
    sh.setColumnWidth(10, 150);
    sh.setColumnWidth(11, 150);
    sh.setColumnWidth(12, 160);
    sh.setColumnWidth(13, 120);
  } catch (e) {}
  sh.setFrozenRows(1);
  try { applyProfessionalTableStyling_(sh, sh.getLastColumn()); } catch (e) { Logger.log('Styling analytics dashboard failed: ' + e.toString()); }
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
    // Cache performance summary once and reuse across dashboard + tech tabs
    const perfMap = getPerfSummaryCached_(cfg, range.startDate, range.endDate);
    refreshAnalyticsDashboard_(range.startDate, range.endDate, perfMap);
    generateTechnicianTabs_(range.startDate, range.endDate, perfMap);
    refreshAdvancedAnalyticsDashboard_(range.startDate, range.endDate);
    configSheet.getRange('B9').setValue(new Date().toLocaleString());
    SpreadsheetApp.getActive().toast(`✅ Dashboard refreshed! ${rowsIngested} rows ingested for ${timeFrame}`, 5);
  } catch (e) {
    Logger.log('refreshDashboardFromAPI error: ' + e.toString());
    SpreadsheetApp.getActive().toast('Error: ' + e.toString().substring(0, 50));
  }
}

function refreshAnalyticsDashboard_(startDate, endDate, perfMapOpt) {
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
    const getHeaderIndex = (variants) => {
      for (const v of variants) {
        const i = headers.findIndex(h => String(h || '').toLowerCase().trim() === String(v).toLowerCase().trim());
        if (i >= 0) return i;
      }
      return -1;
    };
    const startIdx = getHeaderIndex(['Start Time','start_time','start time','start_time_local']);
    const techIdx = getHeaderIndex(['Technician Name','technician_name','technician name','technician','tech']);
    const statusIdx = getHeaderIndex(['Status','session_status','status']);
    const durationIdx = getHeaderIndex(['Total Time','total_time','duration_total_seconds','duration_seconds','duration_total']);
    const pickupIdx = getHeaderIndex(['Waiting Time','waiting_time','pickup_seconds','pickup_seconds_total','pickup']);
    const workIdx = getHeaderIndex(['Work Time','work_time','duration_work_seconds','work_seconds']);
    const customerIdx = getHeaderIndex(['Your Name:','Customer Name','customer_name','customer','customer_name:']);
    const channelIdx = getHeaderIndex(['Channel Name','channel_name']);
    const sessionIdIdx = getHeaderIndex(['Session ID','session_id','session id','id']);
    const callingCardIdx = getHeaderIndex(['Calling Card','calling_card','calling card']);
    const filtered = allData.filter(row => {
      if (!row[startIdx]) return false;
      const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
      return rowDate >= startStr && rowDate <= endStr;
    });
    // Get performance/summary data (reuse cache when provided)
    const performanceData = perfMapOpt || getPerfSummaryCached_(cfg, startDate, endDate);
    
    // Build team data from filtered sessions (for session counts, SLA, work hours, durations)
    const teamData = {};
    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      if (!teamData[tech]) {
        teamData[tech] = { sessions: 0, pickups: [], durations: [], workSeconds: 0, slaHits: 0 };
      }
      teamData[tech].sessions++;
      if (pickupIdx >= 0 && row[pickupIdx]) {
        const ps = parseDurationSeconds_((row[pickupIdx]));
        teamData[tech].pickups.push(ps);
        if (ps <= 60) teamData[tech].slaHits++;
      }
      if (durationIdx >= 0 && row[durationIdx]) {
        teamData[tech].durations.push(parseDurationSeconds_(row[durationIdx]));
      }
      if (workIdx >= 0 && row[workIdx]) teamData[tech].workSeconds += parseDurationSeconds_(row[workIdx]);
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
      const durSec = (durationIdx >= 0 && row[durationIdx]) ? parseDurationSeconds_(row[durationIdx]) : 0;
      const duration = `${Math.floor(durSec/60)}:${String(durSec%60).padStart(2,'0')}`;
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
  // Compute positions so refresh uses the same layout as creation
  const kpiRow = 4;
  const kpiCardsCount = 8; // keep in sync with createMainAnalyticsPage_
  const cardHeight = 3;
  const cardRows = Math.ceil(kpiCardsCount / 3) * cardHeight;
  const liveRow = kpiRow + cardRows + 2;
  const liveSectionHeight = 12; // reservation used during creation
  const liveHeaderRow = liveRow + 1;
  const liveDataRow = liveHeaderRow + 1;
  // Column where the Waiting Queue is placed (matches createMainAnalyticsPage_ layout)
  const queueCol = 9; // column I (to the right side of the main table)
      
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
      
      // Update Team Performance section (positioned after live area)
      const teamTitleRow = liveRow + liveSectionHeight + 1; // title row just after live reservation
      const teamHeaderRow = teamTitleRow + 1;
      const teamDataRow = teamHeaderRow + 1;

      // Ensure title is present
  dashboardSheet.getRange(teamTitleRow, 1).setValue('Team Performance');
  dashboardSheet.getRange(teamTitleRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#0F172A');

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

      // Place Currently Logged In technicians to the right of Team Performance (same column as Waiting Queue)
      // Move to row 28 as requested and align to the queue column
      const techStatusRow = 28;
      dashboardSheet.getRange(techStatusRow, queueCol).setValue('👥 Currently Logged In Technicians');
      dashboardSheet.getRange(techStatusRow, queueCol).setFontSize(14).setFontWeight('bold').setFontColor('#07123B');
      // Merge the title across two columns (name + status)
      try { dashboardSheet.getRange(techStatusRow, queueCol, 1, 2).merge(); } catch (e) {}
      // Clear the area where we'll place logged-in rows (two columns: name, status)
      dashboardSheet.getRange(techStatusRow + 1, queueCol, 100, 2).clearContent();
      if (loggedInTechs.length > 0) {
        dashboardSheet.getRange(techStatusRow + 1, queueCol, loggedInTechs.length, 2).setValues(loggedInTechs);
      } else {
        dashboardSheet.getRange(techStatusRow + 1, queueCol).setValue('No technicians currently logged in');
        dashboardSheet.getRange(techStatusRow + 1, queueCol).setFontStyle('italic').setFontColor('#999999');
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
  dashboardSheet.getRange(activeTitleRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#0F172A');
      
      // Ensure headers are always present (includes Customer Name and Calling Card)
      const activeHeaders = ['Technician', 'Customer Name', 'Start Time', 'Duration', 'Session ID', 'Calling Card'];
      dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setValues([activeHeaders]);
  dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setFontWeight('bold').setBackground('#1A73E8').setFontColor('#FFFFFF');
      dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setBorder(false, false, false, false, false, false); // Remove borders
      
      // Clear data rows and populate with sessions from selected time frame
      dashboardSheet.getRange(activeDataRow, 1, 100, 6).clearContent();
      if (activeRows.length > 0) {
        dashboardSheet.getRange(activeDataRow, 1, activeRows.length, 6).setValues(activeRows);
      } else {
        dashboardSheet.getRange(activeDataRow, 1).setValue('No sessions found for selected time frame');
        dashboardSheet.getRange(activeDataRow, 1).setFontStyle('italic').setFontColor('#999999');
      }
      
    // Update Waiting Queue placed to the right of Live Active Sessions (row 16, column J)
    const queueRow = 16;
    /* queueCol already defined above to keep placement consistent */
      dashboardSheet.getRange(queueRow, queueCol).setValue('⏳ Waiting Queue');
      dashboardSheet.getRange(queueRow, queueCol).setFontSize(14).setFontWeight('bold');
      const queueHeaders = ['Channel', 'Customer', 'Waiting Since', 'Wait Duration'];
      dashboardSheet.getRange(queueRow + 1, queueCol, 1, 4).setValues([queueHeaders]);
      dashboardSheet.getRange(queueRow + 1, queueCol, 1, 4).setFontWeight('bold').setBackground('#EA8600').setFontColor('#FFFFFF');
      dashboardSheet.getRange(queueRow + 1, queueCol, 1, 4).setBorder(false, false, false, false, false, false); // Remove borders
      dashboardSheet.getRange(queueRow + 2, queueCol, 100, 4).clearContent();
      if (waitingRows.length > 0) {
        dashboardSheet.getRange(queueRow + 2, queueCol, waitingRows.length, 4).setValues(waitingRows);
      } else {
        dashboardSheet.getRange(queueRow + 2, queueCol).setValue('No sessions waiting');
        dashboardSheet.getRange(queueRow + 2, queueCol).setFontStyle('italic').setFontColor('#999999');
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
      dashboardSheet.getRange(novaWaveRow, novaWaveCol).setFontSize(16).setFontWeight('bold').setFontColor('#E61C37');

      // Re-apply consistent header styling after all data writes so refresh doesn't overwrite colors
      try {
        // KPI card headers (row 4) and KPI values (row 5)
        const kpiRow = 4;
        const kpiCols = 9; // covers the 3x3 card grid
        dashboardSheet.getRange(kpiRow, 1, 1, kpiCols).setBackground('#1A73E8').setFontColor('#FFFFFF').setFontWeight('bold');
        dashboardSheet.getRange(kpiRow + 1, 1, 1, kpiCols).setBackground('#FFFFFF').setFontColor('#0F172A').setFontWeight('bold');

        // Recreate the visual separator row (same as creation)
        const visualRow1 = kpiRow + 3;
        dashboardSheet.getRange(visualRow1, 1, 1, kpiCols).setBackground('#1A73E8').setFontColor('#FFFFFF');

        // Add special styling for SLA value row (keeps that card highlighted)
        const slaIndex = 6; // SLA Hit % is the 7th card in the kpiCards array (0-based)
        const cardHeight = 3;
        const slaGroupTop = kpiRow + Math.floor(slaIndex / 3) * cardHeight;
        const slaValueRow = slaGroupTop + 1; // value row under SLA header
        try { dashboardSheet.getRange(slaValueRow, 1, 1, kpiCols).setBackground('#1A73E8').setFontColor('#FFFFFF').setFontWeight('bold'); } catch (e) {}

        // Live header (green)
        if (typeof liveHeaderRow !== 'undefined') {
          dashboardSheet.getRange(liveHeaderRow, 1, 1, liveHeaders.length).setBackground('#34A853').setFontColor('#FFFFFF').setFontWeight('bold');
        }

        // Team header (blue)
        if (typeof teamHeaderRow !== 'undefined') {
          dashboardSheet.getRange(teamHeaderRow, 1, 1, teamHeaders.length).setBackground('#1A73E8').setFontColor('#FFFFFF').setFontWeight('bold');
        }

        // Active Sessions header (blue)
        if (typeof activeHeaderRow !== 'undefined') {
          dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setBackground('#1A73E8').setFontColor('#FFFFFF').setFontWeight('bold');
        }

        // Waiting Queue header (orange) — uses queueRow/queueCol
        if (typeof queueRow !== 'undefined' && typeof queueCol !== 'undefined') {
          dashboardSheet.getRange(queueRow + 1, queueCol, 1, 4).setBackground('#EA8600').setFontColor('#FFFFFF').setFontWeight('bold');
        }

        // Currently Logged In header (we merge 2 columns) - give it the same orange header style so it visually matches the waiting queue
        if (typeof techStatusRow !== 'undefined' && typeof queueCol !== 'undefined') {
          try { dashboardSheet.getRange(techStatusRow, queueCol, 1, 2).setBackground('#EA8600').setFontColor('#FFFFFF').setFontWeight('bold'); } catch (e) {}
        }

        // Ensure team data rows have dark font color (so values are visible)
        if (typeof teamDataRow !== 'undefined' && teamRows && teamRows.length > 0) {
          dashboardSheet.getRange(teamDataRow, 1, teamRows.length, teamHeaders.length).setFontColor('#0F172A');
        }

        // Re-run the professional table styling helper to restore banding/borders where necessary
        try { applyProfessionalTableStyling_(dashboardSheet, dashboardSheet.getLastColumn()); } catch (e) { /* non-fatal */ }
      } catch (e) { Logger.log('Re-applying header styles failed: ' + e.toString()); }

      dashboardSheet.getRange(2, 5).setValue(new Date().toLocaleString());
    }
  } catch (e) {
    Logger.log('refreshAnalyticsDashboard_ error: ' + e.toString());
  }
}

function generateTechnicianTabs_(startDate, endDate, perfMapOpt) {
  try {
    const ss = SpreadsheetApp.getActive();
    const cfg = getCfg_();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    // Load extension -> technician mapping for per-account Digium pulls (sheet: extension_map)
    const extMap = getExtensionMap_();
    
    // Clear existing personal dashboard tabs only (avoid wiping other sheets)
    const allSheets = ss.getSheets();
    const reservedSheets = ['Sessions', 'Analytics_Dashboard', 'Dashboard_Config', 'Daily_Summary', 'Support_Data', 'Progress', 'Advanced_Analytics', 'Digium_Raw', 'Digium_Calls', 'API_Smoke_Test'];
    const existingTechSheets = allSheets.filter(sheet => {
      const sheetName = sheet.getName();
      if (reservedSheets.indexOf(sheetName) !== -1) return false;
      try {
        const a1 = sheet.getRange(1, 1).getValue();
        return typeof a1 === 'string' && a1.indexOf('Personal Dashboard') !== -1;
      } catch (e) {
        return false;
      }
    });
    // Clear detected personal dashboards before regenerating
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
    const getHeaderIndex = (variants) => {
      for (const v of variants) {
        const i = headers.findIndex(h => String(h || '').toLowerCase().trim() === v.toLowerCase().trim());
        if (i >= 0) return i;
      }
      return -1;
    };
    const startIdx = getHeaderIndex(['Start Time','start_time', 'start time', 'start_time_local']);
    const techIdx = getHeaderIndex(['Technician Name','technician_name', 'technician name', 'technician', 'tech']);
    const statusIdx = getHeaderIndex(['Status','session_status', 'status']);
    const durationIdx = getHeaderIndex(['Total Time','total_time','duration_total_seconds', 'duration_seconds', 'duration_total']);
    const pickupIdx = getHeaderIndex(['Waiting Time','waiting_time','pickup_seconds', 'pickup_seconds_total', 'pickup']);
    const workIdx = getHeaderIndex(['Work Time','work_time','duration_work_seconds', 'work_seconds']);
    const activeIdx = getHeaderIndex(['Active Time','active_time', 'duration_active_seconds', 'active seconds', 'active']);
    const customerIdx = getHeaderIndex(['Your Name:','Customer Name','customer_name', 'customer', 'customer_name:']); // Column 8
    const sessionIdIdx = getHeaderIndex(['Session ID','session_id', 'session id', 'id']); // Column 1
    const phoneIdx = getHeaderIndex(['Your Phone #:','caller_phone', 'caller phone', 'phone']); // Column 26
    const companyIdx = getHeaderIndex(['Company name:','Company Name','company_name', 'company']); // Column 24
    const callingCardIdx = getHeaderIndex(['Calling Card','calling_card', 'calling card']); // Column 28
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx] || !row[techIdx]) return false;
      try {
        const rowDate = new Date(row[startIdx]).toISOString().split('T')[0];
        return rowDate >= startStr && rowDate <= endStr;
      } catch (e) {
        return false;
      }
    });
  // Build tech list from roster (extension_map) plus any techs present in the filtered data
  const rosterNames = getRosterTechnicianNames_();
  const techSet = new Set((rosterNames || []).filter(Boolean));
  filtered.forEach(row => { const t = row[techIdx]; if (t) techSet.add(String(t)); });
  const techs = Array.from(techSet);
  // Pull performance summary once for the whole range; used for Avg Session (API)
  const perfByTech = perfMapOpt || getPerfSummaryCached_(cfg, startDate, endDate) || {};
    for (const techName of techs) {
      const safeName = techName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
      let techSheet = ss.getSheetByName(safeName);
      if (!techSheet) techSheet = ss.insertSheet(safeName);
  const techRows = filtered.filter(row => row[techIdx] === techName);
      // Clear again to ensure clean slate (in case sheet existed)
      techSheet.clear();
      // Important: clearing values does not remove row/column groups; fully reset any prior groups
      try {
        const maxRows = techSheet.getMaxRows();
        const maxCols = techSheet.getMaxColumns();
        // Shift depths negatively to zero out any existing groups (safe no-op if none)
        techSheet.getRange(1, 1, Math.max(1, maxRows), 1).shiftRowGroupDepth(-8);
        techSheet.getRange(1, 1, 1, Math.max(1, maxCols)).shiftColumnGroupDepth(-8);
      } catch (e) { /* non-fatal */ }
      techSheet.getRange(1, 1).setValue(`👤 ${techName} - Personal Dashboard`);
      techSheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#9C27B0');
      techSheet.getRange(1, 1, 1, 5).merge();
      techSheet.getRange(2, 1).setValue('Time Frame:');
      techSheet.getRange(2, 2).setFormula('=Dashboard_Config!B3');
      techSheet.getRange(2, 4).setValue('Last Updated:');
      techSheet.getRange(2, 5).setValue(new Date().toLocaleString());
  // Normalize durations/pickups/work seconds to seconds for robust calculations
  const durations = techRows.map(r => parseDurationSeconds_(r[durationIdx] || 0)).filter(Boolean);
  const pickups = techRows.map(r => parseDurationSeconds_(r[pickupIdx] || 0)).filter(Boolean);
  const workSeconds = techRows.map(r => parseDurationSeconds_(r[workIdx] || 0)).filter(Boolean).reduce((a,b) => a + b, 0);
  const activeSecondsArr = techRows.map(r => parseDurationSeconds_(r[activeIdx] || 0)).filter(Boolean);
      const completed = techRows.filter(r => r[statusIdx] === 'Ended').length;
      const active = techRows.filter(r => r[statusIdx] === 'Active').length;
    const avgDur = durations.length > 0 ? (durations.reduce((a,b) => a+b, 0) / durations.length / 60).toFixed(1) : '0';
    const avgPickup = pickups.length > 0 ? (pickups.reduce((a,b) => a+b, 0) / pickups.length / 60).toFixed(1) : '0';
  const slaHits = pickups.filter(p => p <= 60).length;
      const slaPct = pickups.length > 0 ? ((slaHits / pickups.length) * 100).toFixed(1) : '0';
  const days = Math.max(1, (new Date(endDate) - new Date(startDate)) / (1000*60*60*24));
  // Sessions per hour should be a numeric value (not text) and formatted as a decimal later
  const sessionsPerHour = (techRows.length / days / 8);
  // Total Active Hours should reflect active time from sessions (not work time)
  const totalActiveSeconds = activeSecondsArr.reduce((a,b)=>a+b, 0);
  const activeHoursTotal = (totalActiveSeconds / 3600).toFixed(1);
      // Count Nova Wave sessions for this technician
      // callingCardIdx already defined above
      const novaWaveCount = callingCardIdx >= 0 ? 
        techRows.filter(row => {
          const callingCard = String(row[callingCardIdx] || '').toLowerCase();
          return callingCard.includes('nova wave chat');
        }).length : 0;
      
  const kpiRow = 4;
      // Build KPI values. Durations/pickups are stored here as spreadsheet time fractions (seconds/86400)
      const avgDurationSeconds = durations.length > 0 ? Math.round(durations.reduce((a,b)=>a+b,0) / durations.length) : 0;
      const avgPickupSeconds = pickups.length > 0 ? Math.round(pickups.reduce((a,b)=>a+b,0) / pickups.length) : 0;
      const kpis = [
        ['Total Sessions', techRows.length],
        ['Nova Wave Sessions', novaWaveCount],
        // Replace Avg Duration with Active Time (average) stored as numeric time fraction for hh:mm:ss formatting
        ['Active Time', (activeSecondsArr.length > 0 ? Math.round(activeSecondsArr.reduce((a,b)=>a+b,0) / activeSecondsArr.length) : 0) / 86400],
        ['Avg Pickup Time', avgPickupSeconds / 86400],
        ['SLA Hit %', slaPct + '%'],
        ['Total Active Hours', activeHoursTotal + ' hrs'],
        // Total Login Time is the sum of active seconds across all sessions, shown as time
        ['Total Login Time', totalActiveSeconds / 86400],
        ['Sessions/Hour', sessionsPerHour]
      ];
      for (let i = 0; i < kpis.length; i++) {
        const row = kpiRow + Math.floor(i / 2);
        const col = (i % 2) * 3 + 1;
        techSheet.getRange(row, col).setValue(kpis[i][0]);
        techSheet.getRange(row, col).setFontSize(11).setFontColor('#666666');
        techSheet.getRange(row, col + 1).setValue(kpis[i][1]);
        techSheet.getRange(row, col + 1).setFontSize(16).setFontWeight('bold').setFontColor('#9C27B0');
        // If this KPI is a duration/pickup numeric fraction, format it as hh:mm:ss
        try {
          const label = String(kpis[i][0] || '');
          if (/avg duration/i.test(label) || /avg pickup/i.test(label) || /active time/i.test(label) || /login time/i.test(label)) {
            techSheet.getRange(row, col + 1).setNumberFormat('hh:mm:ss');
          } else if (/sessions\/hour/i.test(label)) {
            // Ensure Sessions/Hour is displayed as a single-decimal number
            techSheet.getRange(row, col + 1).setNumberFormat('0.0');
          }
        } catch (e) { /* ignore formatting errors */ }
        techSheet.getRange(row, col, 1, 2).setBorder(true, true, true, true, true, true);
      }
      // --- Build wide per-day KPI summary for this technician ---
      // Place the per-day summary directly under the KPI cards so it appears first
      const summaryStart = kpiRow + Math.ceil(kpis.length / 2) + 2;
      techSheet.getRange(summaryStart, 1).setValue('Per-day Performance Summary');
      techSheet.getRange(summaryStart, 1).setFontSize(14).setFontWeight('bold');
      // build dates array
      const dates = [];
      let currentDate = new Date(startDate);
      const endDateObj = new Date(endDate);
      while (currentDate <= endDateObj) {
        dates.push(currentDate.toISOString().split('T')[0]);
        currentDate.setDate(currentDate.getDate() + 1);
      }
    const headerRow = ['Metric', ...dates, 'Totals/Averages'];
    const totalSessionsRow = ['Total sessions'];
    const totalActiveRow = ['Total Active Time'];
    const totalWorkRow = ['Total Work Time'];
    const totalLoginRow = ['Total Login Time'];
    const avgSessionRow = ['Avg Session (API)'];
    const avgPickupRow = ['Avg Pick-up Speed'];
    let totalSessionsAll = 0;
    let totalActiveAll = 0;
    let totalWorkAll = 0;
    let totalLoginAll = 0;
      let totalAvgSeconds = 0;
      let daysWithData = 0;
      let totalPickupSeconds = 0;
      let totalPickupCount = 0;
      const dailyMap = {};
      dates.forEach(d => { dailyMap[d] = { sessions: [], totalActiveSeconds: 0, totalWorkSeconds: 0 }; });
      techRows.forEach(r => {
        try {
          const d = new Date(r[startIdx]).toISOString().split('T')[0];
          if (!dailyMap[d]) dailyMap[d] = { sessions: [], totalActiveSeconds: 0, totalWorkSeconds: 0 };
          dailyMap[d].sessions.push(r);
          // Sum ACTIVE time for daily totals
          dailyMap[d].totalActiveSeconds += parseDurationSeconds_(r[activeIdx] || 0);
          // Sum WORK time for daily totals (if present)
          try { if (workIdx >= 0) dailyMap[d].totalWorkSeconds += parseDurationSeconds_(r[workIdx] || 0); } catch (e) {}
        } catch (e) { }
      });
      dates.forEach(d => {
        const data = dailyMap[d];
        const count = data && data.sessions ? data.sessions.length : 0;
        totalSessionsRow.push(count);
        totalSessionsAll += count;
        const activeSecs = data && data.totalActiveSeconds ? data.totalActiveSeconds : 0;
        totalActiveRow.push(activeSecs / 86400);
        totalActiveAll += activeSecs;
  const workSecs = data && data.totalWorkSeconds ? data.totalWorkSeconds : 0;
  totalWorkRow.push(workSecs / 86400);
  totalWorkAll += workSecs;
  // For Login Time, use the same source as Active (sum of active seconds for the day)
  totalLoginRow.push(activeSecs / 86400);
  totalLoginAll += activeSecs;
        // For Avg Session, prefer API performance average for this technician
        const perf = perfByTech[techName] || perfByTech[String(techName).trim()] || null;
        if (perf && perf.avgDuration) {
          const avg = Number(perf.avgDuration) || 0;
          avgSessionRow.push(avg / 86400);
          totalAvgSeconds += avg;
          daysWithData++;
        } else if (data && data.sessions.length > 0) {
          const durations = data.sessions.map(s => parseDurationSeconds_(s[durationIdx] || 0)).filter(Boolean);
          if (durations.length > 0) {
            const avg = durations.reduce((a,b)=>a+b,0) / durations.length;
            avgSessionRow.push(avg / 86400);
            totalAvgSeconds += avg;
            daysWithData++;
          } else {
            avgSessionRow.push(0);
          }
        } else {
          avgSessionRow.push(0);
        }

        if (data && data.sessions.length > 0) {
          const pickups = data.sessions.map(s => parseDurationSeconds_(s[pickupIdx] || 0)).filter(p=>p>0);
          if (pickups.length > 0) {
            const avgp = pickups.reduce((a,b)=>a+b,0) / pickups.length;
            avgPickupRow.push(avgp / 86400);
            totalPickupSeconds += avgp * pickups.length;
            totalPickupCount += pickups.length;
          } else {
            avgPickupRow.push(0);
          }
        } else {
          avgPickupRow.push(0);
        }
      });
  totalSessionsRow.push(totalSessionsAll);
  totalActiveRow.push(totalActiveAll / 86400);
  totalWorkRow.push(totalWorkAll / 86400);
  totalLoginRow.push(totalLoginAll / 86400);
      avgSessionRow.push(daysWithData>0 ? (totalAvgSeconds / daysWithData) / 86400 : 0);
      avgPickupRow.push(totalPickupCount>0 ? (totalPickupSeconds / totalPickupCount) / 86400 : 0);
  const wideRows = [headerRow, totalSessionsRow, totalActiveRow, totalLoginRow, totalWorkRow, avgSessionRow, avgPickupRow];
      techSheet.getRange(summaryStart + 1, 1, wideRows.length, wideRows[0].length).setValues(wideRows);
      // style header
      try { techSheet.getRange(summaryStart + 1, 1, 1, wideRows[0].length).setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF'); } catch(e){}
      // Format per-row: Total sessions should be an integer, the other rows are time fractions (hh:mm:ss)
      try {
        const valueCols = wideRows[0].length - 1; // excluding 'Metric' column
        if (valueCols > 0) {
          // Total sessions row -> integer
          techSheet.getRange(summaryStart + 2, 2, 1, valueCols).setNumberFormat('0');
          // Time rows -> hh:mm:ss (Active, Login, Work, Avg Session, Avg Pick-up)
          techSheet.getRange(summaryStart + 3, 2, 5, valueCols).setNumberFormat('hh:mm:ss');
        }
      } catch (e) { /* ignore formatting errors */ }

  // --- Session Details (placed after the per-day summary) ---
      const detailRow = summaryStart + 1 + wideRows.length + 2;
      techSheet.getRange(detailRow, 1).setValue('Session Details');
      techSheet.getRange(detailRow, 1).setFontSize(14).setFontWeight('bold');
      // Column order: Date, Session ID, Customer Name, Phone Number, Duration, Pickup
      const detailHeaders = ['Date', 'Session ID', 'Customer Name', 'Phone Number', 'Duration', 'Pickup'];
      techSheet.getRange(detailRow + 1, 1, 1, detailHeaders.length).setValues([detailHeaders]);
      techSheet.getRange(detailRow + 1, 1, 1, detailHeaders.length).setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF');

      // Resolve indices that might exist for fallback lookups
      const resolvedIdx = headers.indexOf('resolved_unresolved');
  const callerNameIdx = getHeaderIndex(['Your Name:','caller_name', 'caller name', 'caller', 'Caller Name', 'Customer Name']);

      // Helpers: content heuristics and formatting
      const isStatus = (s) => { if (!s) return false; return /closed|waiting|active|connected|resolved|unresolved|connecting|in session|closed by/i.test(String(s)); };
      const looksLikeName = (s) => { if (!s) return false; return /^[A-Za-z ,.'-]{3,}$/.test(String(s).trim()); };
      const looksLikePhone = (s) => { if (!s) return false; return /\d{6,}|\(\d{3}\)\s*\d{3}/.test(String(s)); };
      const secToTimeValue = (seconds) => {
        const s = Math.max(0, Math.floor(Number(seconds) || 0));
        return s / 86400; // spreadsheet time fraction
      };

      const pickCustomerFromRow = (row) => {
        // Candidates in order of preference
        const candIdx = [customerIdx, callerNameIdx, companyIdx, phoneIdx];
        for (const ci of candIdx) {
          if (ci >= 0) {
            const v = (row[ci] || '').toString().trim();
            if (v && looksLikeName(v) && !isStatus(v)) return v;
          }
        }
        // Try email/company fallback
        if (customerIdx >= 0 && row[customerIdx]) {
          const v = String(row[customerIdx]).trim();
          if (v && !isStatus(v)) return v;
        }
        if (companyIdx >= 0 && row[companyIdx]) {
          const v = String(row[companyIdx]).trim();
          if (v && looksLikeName(v)) return v;
        }
        return 'Anonymous';
      };

      const detailRows = techRows.slice(0, 50).map(row => {
        const date = row[startIdx] ? new Date(row[startIdx]).toISOString().split('T')[0] : '';
        const customerName = pickCustomerFromRow(row);
        // phone: prefer explicit phone column, otherwise try company or caller columns
        let phoneNumber = '';
        if (phoneIdx >= 0) phoneNumber = row[phoneIdx] || '';
        if ((!phoneNumber || !looksLikePhone(phoneNumber)) && companyIdx >= 0) phoneNumber = row[companyIdx] || phoneNumber || '';
        if ((!phoneNumber || !looksLikePhone(phoneNumber)) && callerNameIdx >= 0) phoneNumber = row[callerNameIdx] || phoneNumber || '';

        // Duration and pickup as true time values (fraction of day)
        const durTime = row[durationIdx] ? secToTimeValue(row[durationIdx]) : 0;
        const pickupTime = row[pickupIdx] ? secToTimeValue(row[pickupIdx]) : 0;
        // session id fallback: if session id looks like a timestamp, try tracking_id or other headers
        let sessionId = row[sessionIdIdx] || '';
        const looksLikeDateTime = (s) => { if (!s) return false; const t=String(s); return /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(t) || /\d{4}-\d{2}-\d{2}/.test(t) || /\d{1,2}:\d{2}:\d{2}/.test(t); };
        if (looksLikeDateTime(sessionId)) {
          // try tracking id
          const trackIdx = headers.findIndex(h => /tracking[_ ]?id/i.test(String(h||'')));
          if (trackIdx >= 0 && row[trackIdx]) sessionId = String(row[trackIdx]);
          else sessionId = '';
        }

        return [ date, sessionId, customerName, phoneNumber, durTime, pickupTime ];
      });
      if (detailRows.length > 0) {
        techSheet.getRange(detailRow + 2, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
        // Apply time format for duration and pickup columns
        const durCol = detailHeaders.indexOf('Duration') + 1;
        const pickupCol = detailHeaders.indexOf('Pickup') + 1;
        try {
          techSheet.getRange(detailRow + 2, durCol, detailRows.length, 1).setNumberFormat('hh:mm:ss');
          techSheet.getRange(detailRow + 2, pickupCol, detailRows.length, 1).setNumberFormat('hh:mm:ss');
        } catch (e) { /* ignore formatting errors */ }
        // Group only the data rows so title + header remain visible when collapsed
        try {
          // Ensure any old grouping around this region is cleared so header/title do not collapse
          try {
            const clearStart = Math.max(1, detailRow);
            const clearCount = Math.max(1, detailRows.length + 3); // title + header + data
            techSheet.getRange(clearStart, 1, clearCount, 1).shiftRowGroupDepth(-8);
          } catch (e) { /* ignore */ }
          // Create a fresh group starting at the first data row only
          const groupRange = techSheet.getRange(detailRow + 2, 1, detailRows.length, 1);
          groupRange.shiftRowGroupDepth(1);
          try { const grp = techSheet.getRowGroup(detailRow + 2, 1); if (grp) grp.collapse(); } catch (e) {}
        } catch (e) { Logger.log('Collapsible session details grouping failed: ' + e.toString()); }
      }
      techSheet.setColumnWidth(1, 100);
      techSheet.setColumnWidth(2, 150);
      techSheet.setColumnWidth(3, 150);
      techSheet.setColumnWidth(4, 120);
      techSheet.setColumnWidth(5, 120);
      techSheet.setColumnWidth(6, 120);

      // Insert Digium wide-format call summary for this tech at row 18, column H (8)
      try {
        // Pull Digium per technican if extension(s) exist; fallback to none if missing
        let digiumWide = null;
        const norm = (s) => String(s || '').trim().toLowerCase();
        const extList = extMap[norm(techName)] || [];
        if (extList && extList.length) {
          try {
            const dr = fetchDigiumCallReports_(startDate, endDate, { breakdown: 'by_day', account_ids: extList });
            if (dr && dr.ok && dr.dates && dr.rows) digiumWide = dr;
          } catch (e) { Logger.log('Digium per-tech fetch failed for '+techName+': ' + e.toString()); }
        }
        if (digiumWide) {
          const startCol = 8; // H
          const startRow = 18;
          const header = ['Metric'].concat(digiumWide.dates || []);
          techSheet.getRange(startRow, startCol, 1, header.length).setValues([header]);
          if (digiumWide.rows && digiumWide.rows.length) {
            techSheet.getRange(startRow + 1, startCol, digiumWide.rows.length, digiumWide.rows[0].length).setValues(digiumWide.rows);
          }
          try { techSheet.getRange(startRow, startCol, 1, header.length).setFontWeight('bold').setBackground('#E5E7EB').setFontColor('#000000'); } catch (e) {}
          techSheet.setColumnWidth(startCol, 140);
          // Convert duration rows to hh:mm:ss on this sheet only
          try { convertSecondsToDayFractionForTableAt_(techSheet, startRow, startCol); } catch (e) {}
        }
      } catch (e) { Logger.log('Failed to write Digium data to tech sheet ' + safeName + ': ' + e.toString()); }

      
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
          // Accept XML or TEXT; getReportTry_ already validated acceptable shapes
          if (!t) {
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
    // Re-apply visual polishing for readability after all sections are written
    try { applyAdvancedAnalyticsStyling_(analyticsSheet); } catch (e) { Logger.log('applyAdvancedAnalyticsStyling_ failed: ' + e.toString()); }
    
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
  sheet.getRange(hourRow + 1, 1, 1, 3).setFontWeight('bold').setBackground('#07123B').setFontColor('#FFFFFF');
    
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
  sheet.getRange(dayRow + 1, 1, 1, 3).setFontWeight('bold').setBackground('#07123B').setFontColor('#FFFFFF');
    
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

// Styling helper for Advanced_Analytics to improve readability without changing data
function applyAdvancedAnalyticsStyling_(sheet) {
  try {
    if (!sheet) return;
    const lastRow = sheet.getLastRow() || 1;
    const lastCol = sheet.getLastColumn() || 1;

    // Global font and sizing
    sheet.getDataRange().setFontFamily('Arial').setFontSize(10).setVerticalAlignment('top');

    // Freeze top area (title + timeframe rows)
    try { sheet.setFrozenRows(3); } catch (e) {}

    // Set widths for first several columns to improve readability
    try {
      const widths = [220, 140, 120, 120, 140, 120, 120, 120];
      for (let c = 1; c <= Math.min(widths.length, lastCol); c++) sheet.setColumnWidth(c, widths[c-1]);
      // For remaining columns, set a moderate width
      for (let c = widths.length + 1; c <= Math.min(lastCol, 30); c++) {
        if (sheet.getColumnWidth(c) < 100) sheet.setColumnWidth(c, 100);
      }
    } catch (e) { /* non-fatal */ }

    // Apply subtle banding to the main body area (rows 4..lastRow)
    if (lastRow > 4) {
      try {
        const bodyRange = sheet.getRange(4, 1, Math.max(1, lastRow - 3), lastCol);
        // remove existing banding then apply light grey banding
        try { bodyRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) {}
        // add thin borders for readability
        bodyRange.setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
      } catch (e) { /* ignore */ }
    }

    // Make top title row prominent
    try {
      sheet.getRange(1, 1, 1, lastCol).setFontSize(16).setFontWeight('bold').setFontColor('#1A73E8');
    } catch (e) {}

    // Ensure table header rows (bold & dark background set by creators) keep contrast
    try { sheet.getRange(1,1,3,lastCol).setWrap(true); } catch (e) {}

    // Auto-resize first N columns where useful (safe operation)
    try {
      const autoCols = Math.min(lastCol, 8);
      for (let c = 1; c <= autoCols; c++) sheet.autoResizeColumn(c);
    } catch (e) {}

    // Highlight header/title rows in the body (rows 4..lastRow): give them a light background and black text
    try {
      const bodyStart = 4;
      if (lastRow >= bodyStart) {
        const rowCount = lastRow - bodyStart + 1;
        const vals = sheet.getRange(bodyStart, 1, rowCount, lastCol).getValues();
        const weights = sheet.getRange(bodyStart, 1, rowCount, lastCol).getFontWeights();
        for (let i = 0; i < vals.length; i++) {
          try {
            const rowVals = vals[i];
            const rowWeights = weights[i];
            const firstWeight = rowWeights && rowWeights.length ? rowWeights[0] : 'normal';
            const nonEmpty = rowVals.reduce((acc, v) => acc + (v != null && String(v).trim() !== '' ? 1 : 0), 0);
            // If the first cell is bold (commonly a section header) OR the row has multiple non-empty cells and at least one bold cell,
            // treat it as a header/column title row and apply light background + black text for readability.
            const anyBold = rowWeights.some(w => String(w).toLowerCase() === 'bold');
            if (firstWeight === 'bold' || (nonEmpty > 1 && anyBold)) {
              const rr = bodyStart + i;
              try { sheet.getRange(rr, 1, 1, lastCol).setBackground('#E5E7EB').setFontColor('#000000'); } catch (e) {}
            }
          } catch (e) { /* per-row non-fatal */ }
        }
      }
    } catch (e) { /* non-fatal overall */ }

  // Do not clear background colors globally — this can remove section header backgrounds set earlier.
  // Instead ensure header rows (1-3) keep their formatting and let section code control header colors.

  } catch (e) {
    Logger.log('applyAdvancedAnalyticsStyling_ error: ' + e.toString());
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
  if (Number(row[pickupIdx]) <= 60) techMetrics[tech].slaHits++;
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
  sheet.getRange(tableRow, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#07123B').setFontColor('#FFFFFF');
    
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
  sheet.getRange(tableRow + 1, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#07123B').setFontColor('#FFFFFF');
    
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
  const slaHits = pickups.filter(p => p <= 60).length;
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
