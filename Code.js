// Cache Performance SUMMARY results for a short period to avoid redundant API calls
function getPerfSummaryCached_(cfg, startDate, endDate) {
  try {
    const cache = CacheService.getScriptCache();
    const key = `perf:${isoDate_(startDate)}:${isoDate_(endDate)}`;
    const cached = cache.get(key);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }
    const empty = {};
    try { cache.put(key, JSON.stringify(empty), 300); } catch (e) {}
    return empty;
  } catch (e) {
    Logger.log('getPerfSummaryCached_ error: ' + e.toString());
    return {};
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
const NODE_CANDIDATES_DEFAULT = [5648341];

/* ===== Runtime settings ===== */
// When false, we request XML from Rescue where supported and parse it to avoid column misalignment.
// We still fall back to TEXT automatically if XML is not available.
const FORCE_TEXT_OUTPUT = false;
const SHEETS_SESSIONS_TABLE = 'Sessions'; // Main data storage sheet
// If true, write raw API values into the Sessions sheet (preserve empty strings
// and original formatting) instead of converting timestamps/durations to Dates/numbers.
const STORE_RAW_SESSIONS = true;

// Optional verbose logging toggle for development. Set to true to enable noisy logs.
const DEBUG = false;
function dlog() { if (DEBUG) try { Logger.log([].map.call(arguments, String).join(' ')); } catch (e) {} }

// Thoroughly reset a personal dashboard sheet so no stale UI artifacts linger
// - Clears all values and direct formatting
// - Removes filters, charts, and conditional formatting rules scoped to this sheet
// - Removes named ranges that point to this sheet
// - Attempts to clear any existing row/column groups depth
function resetSheetCompletely_(sheet) {
  if (!sheet) return;
  const ss = sheet.getParent();
  try { sheet.clear(); } catch (e) { Logger.log('resetSheetCompletely_: clear failed: ' + e.toString()); }
  try { const f = sheet.getFilter(); if (f) f.remove(); } catch (e) { Logger.log('resetSheetCompletely_: remove filter failed: ' + e.toString()); }
  try { (sheet.getCharts() || []).forEach(c => { try { sheet.removeChart(c); } catch (e2) {} }); } catch (e) { Logger.log('resetSheetCompletely_: remove charts failed: ' + e.toString()); }
  // Remove conditional formatting rules that reference this sheet
  try {
    let getRulesFn = null;
    let setRulesFn = null;
    if (sheet && typeof sheet.getConditionalFormatRules === 'function' && typeof sheet.setConditionalFormatRules === 'function') {
      getRulesFn = () => sheet.getConditionalFormatRules();
      setRulesFn = (rules) => sheet.setConditionalFormatRules(rules);
    } else if (ss && typeof ss.getConditionalFormatRules === 'function' && typeof ss.setConditionalFormatRules === 'function') {
      getRulesFn = () => ss.getConditionalFormatRules();
      setRulesFn = (rules) => ss.setConditionalFormatRules(rules);
    }
    if (getRulesFn && setRulesFn) {
      const rules = getRulesFn();
    if (rules && rules.length) {
      const kept = [];
      for (let i = 0; i < rules.length; i++) {
        const rule = rules[i];
        const ranges = rule.getRanges && rule.getRanges();
        if (!ranges || !ranges.length) { kept.push(rule); continue; }
        const touchesSheet = ranges.some(r => r.getSheet && r.getSheet().getSheetId() === sheet.getSheetId());
        if (!touchesSheet) kept.push(rule);
      }
        if (kept.length !== rules.length) setRulesFn(kept);
      }
    }
  } catch (e) { Logger.log('resetSheetCompletely_: conditional formatting cleanup failed: ' + e.toString()); }
  // Remove named ranges tied to this sheet
  try {
    const named = ss.getNamedRanges();
    (named || []).forEach(nr => {
      try {
        const r = nr.getRange && nr.getRange();
        if (r && r.getSheet && r.getSheet().getSheetId() === sheet.getSheetId()) {
          nr.remove();
        }
      } catch (e2) { /* ignore */ }
    });
  } catch (e) { Logger.log('resetSheetCompletely_: named range cleanup failed: ' + e.toString()); }
  // Clear any row/column group depths that might have persisted
  try {
    const lastDataRow = Math.max(sheet.getLastRow(), 100);
    const lastDataCol = Math.max(sheet.getLastColumn(), 10);
    sheet.getRange(1, 1, lastDataRow, 1).shiftRowGroupDepth(-8);
    sheet.getRange(1, 1, 1, lastDataCol).shiftColumnGroupDepth(-8);
  } catch (e) { /* non-fatal */ }
}

/* ===== Digium report cache (in-memory for current execution) ===== */
const DIGIUM_REPORT_CACHE_ = {};
const DIGIUM_HOURLY_QUEUE_ACCOUNT_IDS = ['300'];

/* ===== Menu ===== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Rescue')
    //.addItem('Configure Secrets', 'uiConfigureSecrets')
    //.addSeparator()
    //.addItem('🔍 API Smoke Test', 'apiSmokeTest')
    //.addSeparator()
    .addItem('Pull Yesterday → Sheets', 'pullDateRangeYesterday')
    .addItem('Pull Today → Sheets', 'pullDateRangeToday')
    .addItem('Pull Last Week → Sheets', 'pullDateRangeLastWeek')
    .addItem('Pull This Week → Sheets', 'pullDateRangeThisWeek')
    .addItem('Pull Previous Month → Sheets', 'pullDateRangePreviousMonth')
    .addItem('Pull Current Month → Sheets', 'pullDateRangeCurrentMonth')
    .addItem('Pull Custom Range → Sheets', 'uiIngestRangeToSheets')
    //.addSeparator()
    //.addItem('🚀 Analytics Dashboard', 'createAnalyticsDashboard')
    //.addItem('🔄 Refresh Dashboard (Pull from API)', 'refreshDashboardFromAPI')

    .addSeparator()
    /*.addSubMenu(
      SpreadsheetApp.getUi().createMenu('SUMMARY Mode')
        .addItem('Use CHANNEL only (faster)', 'setSummaryModeChannelOnly')
        .addItem('Use BOTH: NODE + CHANNEL', 'setSummaryModeBoth')
    )
        */
    .addSeparator()
    //.addItem('📈 Advanced Analytics Dashboard', 'createAdvancedAnalyticsDashboard')
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
var DIGIUM_REPORT_CACHE = (typeof DIGIUM_REPORT_CACHE !== 'undefined' && DIGIUM_REPORT_CACHE) || {
  by_day: {},
  by_account: {}
};

const DIGIUM_EXTENSION_ACCOUNT_OVERRIDES = {
  '252': ['1381'],
  '304': ['1442'],
  '305': ['1430'],
  '306': ['1436'],
  '308': ['1421'],
  '322': ['1423'],
  '355': ['1439'],
  '356': ['1351']
};

function isExcludedTechnician_(name) {
  const norm = String(name || '').trim().toLowerCase();
  if (!norm) return false;
  switch (norm) {
    case 'leonardo duba':
    case 'duba, leonardo':
    case 'ulises pereyra':
    case 'pereyra, ulises':
    case 'other department':
      return true;
    default:
      return false;
  }
}

function filterOutExcludedTechnicians_(headers, rows) {
  if (!headers || !rows || !rows.length) return rows;
  let techIdx = -1;
  const variants = ['technician name', 'technician_name', 'technician'];
  for (let i = 0; i < headers.length; i++) {
    const headerNorm = String(headers[i] || '').trim().toLowerCase();
    if (variants.includes(headerNorm)) {
      techIdx = i;
      break;
    }
  }
  if (techIdx < 0) return rows;
  return rows.filter(row => !isExcludedTechnician_(row[techIdx]));
}

function normalizeTechnicianNameFull_(name) {
  return String(name || '').trim().toLowerCase();
}

function canonicalTechnicianName_(name) {
  if (!name) return '';
  const normalized = normalizeTechnicianNameFull_(name);
  if (!normalized) return '';
  switch (normalized) {
    case 'eddie talal':
      return 'ahmed talal';
    default:
  return normalized;
  }
}

function canonicalTechnicianKey_(name) {
  const canonical = canonicalTechnicianName_(name);
  if (canonical) return canonical;
  const normalized = normalizeTechnicianNameFull_(name);
  if (normalized) return normalized;
  return String(name || '').trim().toLowerCase();
}

function technicianFirstNameKey_(name) {
  const normalized = normalizeTechnicianNameFull_(name);
  if (!normalized) return '';
  return normalized.split(/\s+/)[0];
}

function areTechnicianNameVariations_(name1, name2) {
  const n1 = normalizeTechnicianNameFull_(name1);
  const n2 = normalizeTechnicianNameFull_(name2);
  if (!n1 || !n2) return false;
  if (n1 === n2) return true;
  if (n1.includes(n2) || n2.includes(n1)) {
    return technicianFirstNameKey_(name1) === technicianFirstNameKey_(name2);
  }
  return false;
}
function collectSalesforceTicketMetrics_(startDate, endDate) {
  const result = {
    overall: { totalCreated: 0, totalClosed: 0, openCurrent: 0, topIssues: {} },
    perCanonical: {}
  };
  try {
    const ss = SpreadsheetApp.getActive();
    const sfSheet = ss.getSheetByName('Salesforce_Raw');
    if (!sfSheet) return result;

    const range = sfSheet.getDataRange();
    if (range.getNumRows() <= 1) return result;

    const headers = sfSheet.getRange(1, 1, 1, range.getNumColumns()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, idx) => {
      const key = String(h || '').trim().toLowerCase();
      if (key) headerMap[key] = idx;
    });

    const getIndex = (variants) => {
      for (const v of variants) {
        const key = String(v || '').trim().toLowerCase();
        if (key && key in headerMap) return headerMap[key];
      }
      return -1;
    };

    const ticketIdx = getIndex(['novapos ticket: novapos ticket #', 'novapos ticket #', 'ticket #']);
    const createdByIdx = getIndex(['ticket created by']);
    const closedByIdx = getIndex(['ticket closed by rep.', 'ticket closed by rep']);
    const createdDateIdx = getIndex(['ticket created date', 'created date']);
    const closedDateIdx = getIndex(['ticket closed date', 'closed date']);
    const recordTypeIdx = getIndex(['novapos ticket: record type', 'record type']);
    const daysToCloseIdx = getIndex(['days to close ticket / ticket age', 'days to close ticket', 'ticket age']);
    const mainIssueIdx = getIndex(['tech main issue']);
    const techIssueIdx = getIndex(['technical issue']);
    const novaIssueIdx = getIndex(['nova wave issue', 'nova issue']);
    const novaTechIssueIdx = getIndex(['nova wave technical issue', 'nova technical issue']);

    if (ticketIdx < 0 || createdDateIdx < 0) {
      Logger.log('collectSalesforceTicketMetrics_: required columns missing (ticket or created date)');
      return result;
    }

    const parseDateCell = (value) => {
      if (value instanceof Date && !isNaN(value)) return value;
      if (value == null || value === '') return null;
      const parsed = new Date(value);
      return (parsed instanceof Date && !isNaN(parsed)) ? parsed : null;
    };
    const parseDaysToClose = (value) => {
      if (value == null || value === '') return null;
      if (typeof value === 'number') return isNaN(value) ? null : value;
      const cleaned = String(value).replace(/[^0-9.\-]/g, '').trim();
      if (!cleaned) return null;
      const num = parseFloat(cleaned);
      return isNaN(num) ? null : num;
    };

    const rangeStart = startDate.getTime();
    const rangeEnd = endDate.getTime();

    const accountNameIdx = getIndex(['account name', 'account']);
    const isClosedIdx = getIndex(['is this issue closed?', 'is this issue closed', 'is this issue close']);

    const data = sfSheet.getRange(2, 1, range.getNumRows() - 1, range.getNumColumns()).getValues();
    data.forEach(row => {
      const ticketVal = row[ticketIdx];
      if (ticketVal == null || ticketVal === '') return;

      const createdDate = parseDateCell(createdDateIdx >= 0 ? row[createdDateIdx] : null);
      const closedDate = parseDateCell(closedDateIdx >= 0 ? row[closedDateIdx] : null);
      const createdMillis = createdDate ? createdDate.getTime() : null;
      const closedMillis = closedDate ? closedDate.getTime() : null;

      const createdInRange = createdMillis != null && createdMillis >= rangeStart && createdMillis <= rangeEnd;
      const closedInRange = closedMillis != null && closedMillis >= rangeStart && closedMillis <= rangeEnd;

      const daysToCloseVal = parseDaysToClose(daysToCloseIdx >= 0 ? row[daysToCloseIdx] : null);
      const isClosedRaw = isClosedIdx >= 0 ? row[isClosedIdx] : '';
      const isClosedNorm = String(isClosedRaw || '').trim().toLowerCase();
      let isOpenByStatus = true;
      if (isClosedIdx >= 0) {
        if (['1', 'yes', 'true'].includes(isClosedNorm)) {
          isOpenByStatus = false;
        } else if (['0', 'no', 'false', ''].includes(isClosedNorm)) {
          isOpenByStatus = true;
        }
      } else {
        isOpenByStatus = !closedDate;
      }
      if (closedDate) isOpenByStatus = false;

      const createdByRaw = createdByIdx >= 0 ? row[createdByIdx] : '';
      const closedByRaw = closedByIdx >= 0 ? row[closedByIdx] : '';
      const createdCanonical = canonicalTechnicianName_(createdByRaw);
      const closedCanonical = canonicalTechnicianName_(closedByRaw);

      const ensureEntry = (canonical, rawName) => {
        if (!canonical) return null;
        if (!result.perCanonical[canonical]) {
          result.perCanonical[canonical] = { created: 0, closed: 0, open: 0, issues: {}, rawNames: new Set(), daysToClose: [] };
        }
        if (rawName) {
          const trimmed = String(rawName || '').trim();
          if (trimmed) result.perCanonical[canonical].rawNames.add(trimmed);
        }
        return result.perCanonical[canonical];
      };

      const addIssueLabel = (targetMap, label) => {
        if (!label) return;
        targetMap[label] = (targetMap[label] || 0) + 1;
      };

      const issues = [];
      const mainIssue = mainIssueIdx >= 0 ? row[mainIssueIdx] : '';
      const techIssue = techIssueIdx >= 0 ? row[techIssueIdx] : '';
      const novaIssue = novaIssueIdx >= 0 ? row[novaIssueIdx] : '';
      const novaTechIssue = novaTechIssueIdx >= 0 ? row[novaTechIssueIdx] : '';

      const buildIssueLabel = (primary, secondary) => {
        const primaryStr = String(primary || '').trim();
        const secondaryStr = String(secondary || '').trim();
        if (!primaryStr && !secondaryStr) return '';
        return primaryStr && secondaryStr ? `${primaryStr} – ${secondaryStr}` : (primaryStr || secondaryStr);
      };

      const creatorEntry = ensureEntry(createdCanonical, createdByRaw);
      const closerEntry = ensureEntry(closedCanonical, closedByRaw);
      const accountNameNorm = accountNameIdx >= 0 ? String(row[accountNameIdx] || '').trim().toLowerCase() : '';
      const recordTypeStr = recordTypeIdx >= 0 ? String(row[recordTypeIdx] || '').trim().toLowerCase() : '';
      const isTestAccount = accountNameNorm ? /\b(test|demo)\b/.test(accountNameNorm) : false;
      const isNovaRecordType = recordTypeStr.includes('nova');
      const isCurrentlyOpen = isOpenByStatus && isNovaRecordType && !isTestAccount;

      if (isCurrentlyOpen) {
        result.overall.openCurrent += 1;
        if (creatorEntry) creatorEntry.open += 1;
      }

      if (createdInRange) {
        result.overall.totalCreated += 1;
        if (creatorEntry) creatorEntry.created += 1;

        if (isNovaRecordType && !isTestAccount) {
        const issue1 = buildIssueLabel(mainIssue, techIssue);
        const issue2 = buildIssueLabel(novaIssue, novaTechIssue);
        if (issue1) issues.push(issue1);
        if (issue2) issues.push(issue2);
        issues.forEach(label => {
          addIssueLabel(result.overall.topIssues, label);
          if (creatorEntry) addIssueLabel(creatorEntry.issues, label);
        });
        }
      }

      if (closedInRange && isNovaRecordType && !isTestAccount) {
        result.overall.totalClosed += 1;
        if (closerEntry) {
          closerEntry.closed += 1;
        }
        if (creatorEntry && daysToCloseVal != null && daysToCloseVal > 0) {
          creatorEntry.daysToClose.push(daysToCloseVal);
        }
      }
    });

    Object.keys(result.perCanonical).forEach(key => {
      const entry = result.perCanonical[key];
      entry.rawNames = Array.from(entry.rawNames || []);
      entry.daysToClose = entry.daysToClose || [];
    });
  } catch (e) {
    Logger.log('collectSalesforceTicketMetrics_ error: ' + e.toString());
  }
  return result;
}
function createResolutionDistribution_(sheet, startRow, startDate, endDate, styleRegistry) {
  try {
    styleRegistry = styleRegistry || [];
    sheet.getRange(startRow, 1).setValue('🧩 Resolution Distribution');
    styleAnalyticsSectionHeader_(sheet, startRow, 4);
    const infoRow = startRow + 2;
    sheet.getRange(infoRow, 1, 1, 4).setValues([[
      'Resolution distribution visualizations will appear here once the Salesforce dataset includes the required fields.',
      '',
      '',
      ''
    ]]);
    sheet.getRange(infoRow, 1, 1, 4).merge().setFontColor('#475569').setFontStyle('italic');
    registerAnalyticsTable_(styleRegistry, infoRow, 1, 4, 1);
    return infoRow + 3;
  } catch (e) {
    Logger.log('createResolutionDistribution_ error: ' + e.toString());
    return startRow + 5;
  }
}
function createCapacityIndicators_(sheet, startRow, styleRegistry) {
  try {
    styleRegistry = styleRegistry || [];
    sheet.getRange(startRow, 1).setValue('🚦 Capacity Indicators');
    styleAnalyticsSectionHeader_(sheet, startRow, 4);
    const infoRow = startRow + 2;
    sheet.getRange(infoRow, 1, 1, 4).setValues([[
      'Capacity indicator widgets will appear here once capacity metrics are available.',
      '',
      '',
      ''
    ]]);
    sheet.getRange(infoRow, 1, 1, 4).merge().setFontColor('#475569').setFontStyle('italic');
    registerAnalyticsTable_(styleRegistry, infoRow, 1, 4, 1);
    return infoRow + 3;
  } catch (e) {
    Logger.log('createCapacityIndicators_ error: ' + e.toString());
    return startRow + 5;
  }
}

function createUtilizationRate_(sheet, startRow, startDate, endDate, styleRegistry) {
  try {
    styleRegistry = styleRegistry || [];
    sheet.getRange(startRow, 1).setValue('⚙️ Technician Utilization');
    styleAnalyticsSectionHeader_(sheet, startRow, 4);
    const messageRow = startRow + 2;
    const message = 'Utilization insights will appear here once sufficient performance data is available.';
    sheet.getRange(messageRow, 1, 1, 4).setValues([[message, '', '', '']]);
    sheet.getRange(messageRow, 1, 1, 4).merge().setFontColor('#334155').setFontStyle('italic');
    registerAnalyticsTable_(styleRegistry, messageRow, 1, 4, 1);
    return messageRow + 3;
  } catch (e) {
    Logger.log('createUtilizationRate_ error: ' + e.toString());
    return startRow + 5;
  }
}
function fetchDigiumCallQueueReportsByHour_(startDate, endDate, queueAccountIds) {
  try {
    const cfg = getCfg_();
    const host = cfg.digiumHost;
    const user = cfg.digiumUser;
    const pass = cfg.digiumPass;
    if (!host || !user || !pass) return { ok: false, reason: 'missing_credentials' };
    const ids = Array.isArray(queueAccountIds) ? queueAccountIds.map(id => String(id || '').trim()).filter(Boolean) : [];
    if (!ids.length) return { ok: false, reason: 'no_queue_ids' };

    const startIso = isoDate_(startDate);
    const endIso = isoDate_(endDate);
    const xmlIds = ids.map(id => `<queue_account_id>${xmlEscape_(id)}</queue_account_id>`).join('');
    const paramsXml =
      `\n    <start_date>${xmlEscape_(startIso + ' 00:00:00')}</start_date>\n` +
      `    <end_date>${xmlEscape_(endIso + ' 23:59:59')}</end_date>\n` +
      `    <ignore_weekends>0</ignore_weekends>\n` +
      `    <breakdown>by_hour_of_day</breakdown>\n` +
      `    <queue_account_ids>${xmlIds}</queue_account_ids>\n` +
      `    <report_fields><report_field>total_calls</report_field></report_fields>\n` +
      `    <format>xml</format>\n  `;

    const r = digiumApiCall_('switchvox.callQueueReports.search', paramsXml, user, pass, host);
    if (!r.ok) return { ok: false, error: r.error || r.raw };

    const doc = r.xml;
    const root = doc.getRootElement();
    const resultEl = root.getChild('result') || root;
    let targets = [];
    const hoursEl = resultEl.getChild('hours_of_day');
    if (hoursEl && hoursEl.getChildren) {
      targets = hoursEl.getChildren('hour') || [];
    } else {
      const rowsEl = resultEl.getChild('rows');
      if (rowsEl && rowsEl.getChildren) targets = rowsEl.getChildren('row') || [];
    }

    if (!targets.length) {
      return {
        ok: true,
        categories: [],
        rows: [['Total Calls']],
        totalsAll: { total_calls: 0 },
        fields: ['total_calls'],
        humanLabels: { total_calls: 'Total Calls' },
        raw: r.raw
      };
    }

    const categories = [];
    const values = [];
    targets.forEach(hourEl => {
      let hourAttr = hourEl.getAttribute('hour') || hourEl.getAttribute('name');
      const hourVal = hourAttr ? hourAttr.getValue() : '';
      const hourKey = hourVal !== '' ? hourVal : String(categories.length);
      categories.push(hourKey);
      const totalCallsAttr = hourEl.getAttribute('total_calls');
      const totalCallsVal = totalCallsAttr ? Number(totalCallsAttr.getValue()) : 0;
      values.push(isNaN(totalCallsVal) ? 0 : totalCallsVal);
    });

    const totalsAll = {
      total_calls: values.reduce((sum, v) => sum + (isNaN(v) ? 0 : v), 0)
    };

    return {
      ok: true,
      categories,
      rows: [['Total Calls', ...values]],
      totalsAll,
      fields: ['total_calls'],
      humanLabels: { total_calls: 'Total Calls' },
      raw: r.raw
    };
  } catch (e) {
    Logger.log('fetchDigiumCallQueueReportsByHour_ error: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}
function fetchDigiumCallReports_(startDate, endDate, options) {
  const cfg = getCfg_();
  const host = cfg.digiumHost;
  const user = cfg.digiumUser;
  const pass = cfg.digiumPass;
  if (!host || !user || !pass) return { ok: false, reason: 'missing_credentials' };

  const startIso = isoDate_(startDate);
  const endIso = isoDate_(endDate);
  const startStr = `${startIso} 00:00:00`;
  const endStr = `${endIso} 23:59:59`;

  const fields = (options && options.report_fields) || [
    'total_calls','total_incoming_calls','total_outgoing_calls','talking_duration','call_duration','avg_talking_duration','avg_call_duration'
  ];

  // Human-friendly labels for metrics (used in wide table rows)
  const human = {
    total_calls: 'Total Calls', total_incoming_calls: 'Total Incoming Calls', total_outgoing_calls: 'Total Outgoing Calls',
    talking_duration: 'Talking Duration (s)', call_duration: 'Call Duration (s)', avg_talking_duration: 'Avg Talking Duration (s)', avg_call_duration: 'Avg Call Duration (s)'
  };

  const targetExtensions = options && options.target_extensions ?
    (Array.isArray(options.target_extensions) ? options.target_extensions : String(options.target_extensions).split(','))
      .map(s => String(s).trim()).filter(Boolean) : null;

  let accountIdsForRequest = options && options.account_ids ?
    (Array.isArray(options.account_ids) ? options.account_ids : String(options.account_ids).split(','))
      .map(s => String(s).trim()).filter(Boolean) : null;

  const buildAccountResponseFromCache = (entry, extList) => {
    if (!entry) return { ok: false, reason: 'no_cache' };
    const fieldsForEntry = entry.fields || fields;
    const humanLabels = entry.humanLabels || human;
    const totalsAll = entry.totalsAll || {};
    const perExtension = entry.perExtension || {};
    const resultTotals = {};
    fieldsForEntry.forEach(f => { resultTotals[f] = totalsAll[f] || 0; });

    let matchedExtensions = 0;
    let subsetWeightedTalk = 0;
    let subsetWeightedCall = 0;
    let subsetWeightedCalls = 0;

    if (extList && extList.length) {
      fieldsForEntry.forEach(f => { resultTotals[f] = 0; });
      extList.forEach(ext => {
        const meta = perExtension[ext];
        if (!meta || !meta.metrics) return;
        matchedExtensions++;
        fieldsForEntry.forEach(f => {
          if (f === 'avg_talking_duration' || f === 'avg_call_duration') return;
          resultTotals[f] += meta.metrics[f] || 0;
        });
        const calls = meta.metrics.total_calls || 0;
        subsetWeightedCalls += calls;
        subsetWeightedTalk += (meta.metrics.avg_talking_duration || 0) * calls;
        subsetWeightedCall += (meta.metrics.avg_call_duration || 0) * calls;
      });
      if (fieldsForEntry.indexOf('avg_talking_duration') >= 0) {
        resultTotals.avg_talking_duration = subsetWeightedCalls > 0 ? subsetWeightedTalk / subsetWeightedCalls : 0;
      }
      if (fieldsForEntry.indexOf('avg_call_duration') >= 0) {
        resultTotals.avg_call_duration = subsetWeightedCalls > 0 ? subsetWeightedCall / subsetWeightedCalls : 0;
      }
      if (fieldsForEntry.indexOf('total_calls') >= 0 && subsetWeightedCalls > 0) {
        resultTotals.total_calls = subsetWeightedCalls;
      }
    }

    if (extList && extList.length && matchedExtensions === 0) {
      const wideRowsZero = fieldsForEntry.map(f => [humanLabels[f] || f, 0]);
      return {
        ok: true,
        dates: [],
        rows: wideRowsZero,
        perExtension,
        totalsAll,
        fields: fieldsForEntry.slice(),
        humanLabels,
        raw: entry.raw
      };
    }

    const wideRows = fieldsForEntry.map(f => [humanLabels[f] || f, resultTotals[f] || 0]);
    return {
      ok: true,
      dates: [],
      rows: wideRows,
      perExtension,
      totalsAll,
      fields: fieldsForEntry.slice(),
      humanLabels,
      raw: entry.raw
    };
  };

  // Breakdown options: 'by_date' (maps to 'by_day' in API), 'by_account', 'by_hour_of_day', 'by_day_of_week', 'cumulative'
  // Map 'by_date' to 'by_day' for API compatibility
  let breakdown = (options && options.breakdown) || 'by_day';
  if (breakdown === 'by_date') {
    breakdown = 'by_day'; // API uses 'by_day' for date breakdown
  }

  const startKey = startIso;
  const endKey = endIso;
  const fieldKey = fields.join(',');
  const accountKey = accountIdsForRequest && accountIdsForRequest.length ? accountIdsForRequest.slice().sort().join(',') : 'ALL';
  const cacheKeyBreakdown = breakdown || 'by_day';
  const baseCacheKey = `${cacheKeyBreakdown}|${startKey}|${endKey}|${fieldKey}|${accountKey}`;

  if (breakdown === 'by_account') {
    const cached = DIGIUM_REPORT_CACHE.by_account[baseCacheKey];
    if (cached) {
      Logger.log(`Digium cache hit (by_account) for ${baseCacheKey}`);
      return buildAccountResponseFromCache(cached, targetExtensions);
    }
  } else if (breakdown === 'by_day') {
    const cached = DIGIUM_REPORT_CACHE.by_day[baseCacheKey];
    if (cached) {
      Logger.log(`Digium cache hit (by_day) for ${baseCacheKey}`);
      return {
        ok: true,
        dates: cached.dates.slice(),
        rows: cached.rows.map(row => row.slice()),
        raw: cached.raw,
        fields: cached.fields.slice(),
        humanLabels: cached.humanLabels
      };
    }
  }

  // Build report_fields XML
  let reportFieldsXml = '<report_fields>'; 
  fields.forEach(f => { reportFieldsXml += '<report_field>' + xmlEscape_(f) + '</report_field>'; });
  reportFieldsXml += '</report_fields>';

  // ignore_weekends: 0 => do not ignore weekends (user requested we do NOT ignore weekends)
  // Only include account_ids when explicitly provided (e.g., by_day requests requiring mapped IDs)
  let accountIdsXml = '';
  const shouldIncludeAccountIds = accountIdsForRequest && accountIdsForRequest.length && breakdown !== 'by_account';
  if (shouldIncludeAccountIds) {
    const ids = Array.isArray(accountIdsForRequest)
      ? accountIdsForRequest
      : String(accountIdsForRequest).split(',').map(s => s.trim()).filter(Boolean);
    if (ids && ids.length) {
      accountIdsXml = '\n    <account_ids>' + ids.map(id => '<account_id>' + xmlEscape_(id) + '</account_id>').join('') + '</account_ids>';
    }
  }

  let sortFieldDefault = null;
  switch (breakdown) {
    case 'by_account':
      sortFieldDefault = 'extension';
      break;
    case 'by_day':
      sortFieldDefault = 'date';
      break;
    case 'by_day_of_week':
      sortFieldDefault = 'day_of_week';
      break;
    case 'by_hour_of_day':
      sortFieldDefault = 'hour';
      break;
    default:
      sortFieldDefault = null;
  }
  const sortField = (options && options.sort_field) || sortFieldDefault;
  const sortOrder = (options && options.sort_order) || (sortField ? 'ASC' : null);
  const itemsPerPage = (options && options.items_per_page) || (breakdown === 'by_account' ? 1000 : null);
  const pageNumber = (options && options.page_number) || (itemsPerPage ? 1 : null);

  let paramsXml = `\n    <start_date>${xmlEscape_(startStr)}</start_date>\n    <end_date>${xmlEscape_(endStr)}</end_date>\n    <ignore_weekends>0</ignore_weekends>\n    <breakdown>${xmlEscape_(breakdown)}</breakdown>`;
  if (accountIdsXml) paramsXml += accountIdsXml;
  paramsXml += `\n    ${reportFieldsXml}\n    <format>xml</format>`;
  if (sortField) paramsXml += `\n    <sort_field>${xmlEscape_(sortField)}</sort_field>`;
  if (sortOrder) paramsXml += `\n    <sort_order>${xmlEscape_(sortOrder)}</sort_order>`;
  if (itemsPerPage) paramsXml += `\n    <items_per_page>${xmlEscape_(itemsPerPage)}</items_per_page>`;
  if (pageNumber) paramsXml += `\n    <page_number>${xmlEscape_(pageNumber)}</page_number>`;
  paramsXml += `\n  `;

  // Log the request for debugging
  Logger.log(`Digium API Request - breakdown: ${breakdown}, account_ids: ${accountIdsForRequest && accountIdsForRequest.length ? accountIdsForRequest.join(',') : 'none'}`);
  Logger.log(`Digium API Request XML (first 500 chars): ${paramsXml.substring(0, 500)}`);

  let method = 'switchvox.callReports.search';
  if (breakdown === 'by_day' || breakdown === 'by_day_of_week' || breakdown === 'by_hour_of_day') {
    method = 'switchvox.callReports.phones.search';
  }

  let r = digiumApiCall_(method, paramsXml, user, pass, host);
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
            const fallbackMethod = method === 'switchvox.callReports.phones.search'
              ? 'switchvox.callReports.search'
              : method;
            const r2 = digiumApiCall_(fallbackMethod, paramsNoBreakdown, user, pass, host);
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
  // Try to parse results based on breakdown type
  try {
    const doc = r.xml;
    const root = doc.getRootElement();

    // Handle by_account breakdown - returns account-level totals for all accounts
    // Actual API structure: <result><rows><row extension="..." total_calls="..." .../></rows></result>
    if (breakdown === 'by_account') {
      const resultEl = root.getChild('result') || root;
      const rowsEl = resultEl.getChild('rows');

      const totalsAll = {};
      fields.forEach(f => { totalsAll[f] = 0; });
      const perExtension = {};

      if (!rowsEl) {
        Logger.log('Digium by_account: No <rows> element found - returning zeros');
        const cacheEntry = {
          perExtension: {},
          totalsAll,
          fields: fields.slice(),
          humanLabels: human,
          raw: r.raw
        };
        DIGIUM_REPORT_CACHE.by_account[baseCacheKey] = cacheEntry;
        const wideRows = fields.map(f => [human[f] || f, 0]);
        return { ok: true, dates: [], rows: wideRows, perExtension: cacheEntry.perExtension, totalsAll, fields: cacheEntry.fields, humanLabels: cacheEntry.humanLabels, raw: r.raw };
      }

      const rowElements = rowsEl.getChildren('row');
      if (rowElements && rowElements.length > 0) {
        let totalCallsWeighted = 0;
        let weightedAvgTalkTime = 0;
        let weightedAvgCallTime = 0;

        rowElements.forEach(rowEl => {
          const extAttr = rowEl.getAttribute('extension');
          const extension = extAttr ? String(extAttr.getValue()).trim() : '';
          if (!extension) return;

          const accountNameAttr = rowEl.getAttribute('account_name') || rowEl.getAttribute('name') || rowEl.getAttribute('account') || rowEl.getAttribute('user_name');
          const accountName = accountNameAttr ? String(accountNameAttr.getValue()).trim() : '';
          const label = accountName ? `${extension} - ${accountName}` : extension;

          const metrics = {};
          fields.forEach(f => {
            const attr = rowEl.getAttribute(f);
            metrics[f] = attr ? parseDigiumNum(attr.getValue()) : 0;
            if (f !== 'avg_talking_duration' && f !== 'avg_call_duration') {
              totalsAll[f] += metrics[f];
            }
          });

          const totalCalls = metrics.total_calls || 0;
          totalCallsWeighted += totalCalls;
          weightedAvgTalkTime += (metrics.avg_talking_duration || 0) * totalCalls;
          weightedAvgCallTime += (metrics.avg_call_duration || 0) * totalCalls;

          perExtension[extension] = { label, metrics };
        });

        if (fields.indexOf('avg_talking_duration') >= 0) {
          totalsAll.avg_talking_duration = totalCallsWeighted > 0 ? weightedAvgTalkTime / totalCallsWeighted : 0;
        }
        if (fields.indexOf('avg_call_duration') >= 0) {
          totalsAll.avg_call_duration = totalCallsWeighted > 0 ? weightedAvgCallTime / totalCallsWeighted : 0;
        }

        const cacheEntry = {
          perExtension,
          totalsAll,
          fields: fields.slice(),
          humanLabels: human,
          raw: r.raw
        };
        DIGIUM_REPORT_CACHE.by_account[baseCacheKey] = cacheEntry;
        return buildAccountResponseFromCache(cacheEntry, targetExtensions);
      } else {
        const rowsAttrs = rowsEl.getAttributes();
        if (rowsAttrs && rowsAttrs.length > 0) {
          rowsAttrs.forEach(attr => {
            const name = attr.getName();
            if (fields.indexOf(name) >= 0) {
              totalsAll[name] = parseDigiumNum(attr.getValue());
            }
          });
          const cacheEntry = {
            perExtension: {},
            totalsAll,
            fields: fields.slice(),
            humanLabels: human,
            raw: r.raw
          };
          DIGIUM_REPORT_CACHE.by_account[baseCacheKey] = cacheEntry;
          return buildAccountResponseFromCache(cacheEntry, targetExtensions);
        }

        Logger.log('Digium by_account: No <row> elements or attributes found - returning zeros');
        const cacheEntry = {
          perExtension: {},
          totalsAll,
          fields: fields.slice(),
          humanLabels: human,
          raw: r.raw
        };
        DIGIUM_REPORT_CACHE.by_account[baseCacheKey] = cacheEntry;
        const wideRows = fields.map(f => [human[f] || f, 0]);
        return { ok: true, dates: [], rows: wideRows, perExtension: cacheEntry.perExtension, totalsAll, fields: cacheEntry.fields, humanLabels: cacheEntry.humanLabels, raw: r.raw };
      }
    }

    if (breakdown === 'by_hour_of_day' || breakdown === 'by_day_of_week') {
      const resultEl = root.getChild('result') || root;
      const rowsEl = resultEl ? resultEl.getChild('rows') : null;
      const rowChildren = rowsEl ? rowsEl.getChildren('row') : null;
      if (rowChildren && rowChildren.length) {
        const categoryKey = breakdown === 'by_hour_of_day' ? 'hour' : 'day';
        const categoryMap = {};
        rowChildren.forEach(rowEl => {
          let categoryVal = null;
          const attr = rowEl.getAttribute(categoryKey);
          if (attr && attr.getValue) categoryVal = attr.getValue();
          if (!categoryVal && rowEl.getAttribute('name')) {
            categoryVal = rowEl.getAttribute('name').getValue();
          }
          if (categoryVal == null) return;
          const category = String(categoryVal).trim();
          if (!category) return;
          if (!categoryMap[category]) categoryMap[category] = {};
          fields.forEach(f => {
            const metricAttr = rowEl.getAttribute(f);
            categoryMap[category][f] = metricAttr ? parseDigiumNum(metricAttr.getValue()) : 0;
          });
        });

        let categories = Object.keys(categoryMap);
        if (breakdown === 'by_hour_of_day') {
          categories.sort((a, b) => Number(a) - Number(b));
        } else {
          const order = ['sunday','monday','tuesday','wednesday','thursday','friday','saturday'];
          categories.sort((a, b) => {
            const aIdx = order.indexOf(String(a).toLowerCase());
            const bIdx = order.indexOf(String(b).toLowerCase());
            if (aIdx >= 0 && bIdx >= 0) return aIdx - bIdx;
            return String(a).localeCompare(String(b));
          });
        }

        const wideRows = fields.map(f => {
          const row = [human[f] || f];
          categories.forEach(cat => row.push(categoryMap[cat][f] || 0));
          return row;
        });
        const totalsAll = {};
        fields.forEach((f, idx) => {
          totalsAll[f] = wideRows[idx].slice(1).reduce((sum, val) => {
            const num = Number(val);
            return sum + (isNaN(num) ? 0 : num);
          }, 0);
        });

        return {
          ok: true,
          categories,
          rows: wideRows,
          totalsAll,
          fields: fields.slice(),
          humanLabels: human,
          raw: r.raw
        };
      }
      return {
        ok: true,
        categories: [],
        rows: fields.map(f => [human[f] || f]),
        totalsAll: {},
        fields: fields.slice(),
        humanLabels: human,
        raw: r.raw
      };
    }

    // Handle by_day breakdown (original logic)
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

    DIGIUM_REPORT_CACHE.by_day[baseCacheKey] = {
      dates: dates.slice(),
      rows: wideRows.map(row => row.slice()),
      raw: r.raw,
      fields: fields.slice(),
      humanLabels: human
    };
    return { ok: true, dates: dates, rows: wideRows, raw: r.raw, fields: fields.slice(), humanLabels: human };
  } catch (e) {
    return { ok: false, error: 'parse_failed: ' + e.toString(), raw: r.raw };
  }
}
// Resolve Switchvox account_ids from a list of extension numbers
// Returns array of account_id strings. Falls back to returning the input list if lookup fails.
var EXTENSION_ACCOUNT_ID_MAP_CACHE = (typeof EXTENSION_ACCOUNT_ID_MAP_CACHE !== 'undefined' && EXTENSION_ACCOUNT_ID_MAP_CACHE) || {};

function resolveDigiumAccountIdsDetailed_(extensions) {
  const details = { list: [], map: {} };
  const finalizeDetails = () => {
    const accountSetFinal = new Set();
    Object.keys(details.map).forEach(ext => {
      const extKey = String(ext || '').trim();
      const existing = Array.isArray(details.map[ext]) ? details.map[ext] : [];
      const overrides = (DIGIUM_EXTENSION_ACCOUNT_OVERRIDES[extKey] || []).map(id => String(id || '').trim()).filter(Boolean);
      let normalized = [];
      if (overrides.length) {
        const overrideSet = new Set(overrides);
        normalized = Array.from(overrideSet);
      } else {
        const existingSet = new Set();
        existing.forEach(id => {
          const idStr = String(id || '').trim();
          if (idStr) existingSet.add(idStr);
        });
        normalized = Array.from(existingSet);
      }
      details.map[ext] = normalized.slice();
      EXTENSION_ACCOUNT_ID_MAP_CACHE[extKey] = normalized.slice();
      normalized.forEach(id => {
        const idStr = String(id || '').trim();
        if (idStr) accountSetFinal.add(idStr);
      });
    });
    if (!details.list || !details.list.length) {
      details.list = Array.from(accountSetFinal);
    } else {
      const merged = new Set();
      details.list.forEach(id => {
        const idStr = String(id || '').trim();
        if (idStr) merged.add(idStr);
      });
      accountSetFinal.forEach(id => merged.add(id));
      details.list = Array.from(merged);
    }
    if (!details.list.length) {
      const fallback = Array.isArray(extensions) ? extensions : [String(extensions || '')];
      details.list = fallback.map(ext => String(ext || '').trim()).filter(Boolean);
    }
    return details;
  };
  try {
    const cfg = getCfg_();
    const host = cfg.digiumHost;
    const user = cfg.digiumUser;
    const pass = cfg.digiumPass;
    if (!host || !user || !pass) {
      const fallbackList = Array.isArray(extensions) ? extensions : [String(extensions)];
      details.list = fallbackList.map(ext => String(ext || '').trim()).filter(Boolean);
      details.list.forEach(ext => { if (ext) details.map[ext] = [ext]; });
      return finalizeDetails();
    }

    const ids = Array.isArray(extensions)
      ? extensions.map(s => String(s || '').trim()).filter(Boolean)
      : String(extensions || '').split(',').map(s => s.trim()).filter(Boolean);
    if (!ids.length) {
      return finalizeDetails();
    }

    const extensionMeta = getActiveExtensionMetadata_();
    const extToAccountIds = extensionMeta && extensionMeta.extToAccountIds ? extensionMeta.extToAccountIds : {};
    const manualAccountSet = new Set();
    const remaining = [];

    ids.forEach(ext => {
      const key = String(ext || '').trim();
      if (!key) return;
      const manualIdsSource = extToAccountIds[key] && extToAccountIds[key].length ? extToAccountIds[key] : (DIGIUM_EXTENSION_ACCOUNT_OVERRIDES[key] || []);
      const manualIds = manualIdsSource.filter(id => String(id || '').trim());
      if (manualIds && manualIds.length) {
        details.map[key] = manualIds.slice();
        manualIds.forEach(id => {
          const idStr = String(id || '').trim();
          if (idStr) {
            manualAccountSet.add(idStr);
            EXTENSION_ACCOUNT_ID_MAP_CACHE[key] = EXTENSION_ACCOUNT_ID_MAP_CACHE[key] || [];
            if (!EXTENSION_ACCOUNT_ID_MAP_CACHE[key].includes(idStr)) {
              EXTENSION_ACCOUNT_ID_MAP_CACHE[key].push(idStr);
            }
          }
        });
      } else {
        remaining.push(key);
      }
    });

    if (remaining.length === 0) {
      if (manualAccountSet.size === 0) {
        ids.forEach(ext => {
          const key = String(ext || '').trim();
          if (!key) return;
          const overrideList = (DIGIUM_EXTENSION_ACCOUNT_OVERRIDES[key] || []).map(id => String(id || '').trim()).filter(Boolean);
          if (overrideList.length) {
            details.map[key] = overrideList.slice();
            overrideList.forEach(id => manualAccountSet.add(id));
          } else {
            details.map[key] = [key];
            manualAccountSet.add(key);
          }
        });
      }
      details.list = Array.from(manualAccountSet);
      return finalizeDetails();
    }

    const idsForLookup = remaining;

    const paramsXml =
      '\n    <page_number>1</page_number>\n    <items_per_page>250</items_per_page>\n    <filters>' +
      idsForLookup.map(ext => '<filter field="extension" value="' + xmlEscape_(String(ext)) + '" operator="eq" />').join('') +
      '</filters>\n    <format>xml</format>\n  ';

    let r = digiumApiCall_('switchvox.extensions.search', paramsXml, user, pass, host);
    if (!r.ok) {
      const paramsXmlAlt =
        '\n    <extensions>' +
        idsForLookup.map(ext => '<extension>' + xmlEscape_(String(ext)) + '</extension>').join('') +
        '</extensions>\n    <format>xml</format>\n  ';
      r = digiumApiCall_('switchvox.extensions.getInfo', paramsXmlAlt, user, pass, host);
    }

    if (!r.ok || !r.xml) {
      ids.forEach(ext => { if (ext) details.map[ext] = [ext]; });
      details.list = ids.slice();
      return finalizeDetails();
    }

    const accountSet = new Set();
    const mapOut = {};

    const addMapping = (extension, accountId) => {
      const extKey = String(extension || '').trim();
      const accountKey = String(accountId || '').trim();
      if (!extKey || !accountKey) return;
      accountSet.add(accountKey);
      if (!mapOut[extKey]) mapOut[extKey] = new Set();
      mapOut[extKey].add(accountKey);
    };

      const root = r.xml.getRootElement();
      const result = root.getChild('result') || root;
      const collect = (el) => {
      const children = el.getChildren ? el.getChildren() : [];
      children.forEach(child => {
        const name = String(child.getName() || '').toLowerCase();
          if (name === 'row' || name === 'extension' || name === 'account') {
          const accountAttr =
            (child.getAttribute && child.getAttribute('account_id') && child.getAttribute('account_id').getValue()) ||
            (child.getAttribute && child.getAttribute('user_id') && child.getAttribute('user_id').getValue()) ||
            (child.getAttribute && child.getAttribute('id') && child.getAttribute('id').getValue()) ||
            '';
          const extAttr =
            (child.getAttribute && child.getAttribute('extension') && child.getAttribute('extension').getValue()) ||
            '';
          addMapping(extAttr || '', accountAttr || '');
        }
        collect(child);
        });
      };
      collect(result);

    idsForLookup.forEach(ext => {
      const key = String(ext || '').trim();
      if (!key) return;
      if (!mapOut[key] || mapOut[key].size === 0) {
        // Fallback: map extension to itself so downstream filters still run, even if Digium lookup failed
        mapOut[key] = new Set([key]);
        accountSet.add(key);
      }
    });

    Object.keys(mapOut).forEach(ext => {
      const arr = Array.from(mapOut[ext]);
      details.map[ext] = arr;
      EXTENSION_ACCOUNT_ID_MAP_CACHE[ext] = arr.slice();
      arr.forEach(id => manualAccountSet.add(id));
    });
    details.list = Array.from(manualAccountSet.size ? manualAccountSet : accountSet);
    return finalizeDetails();
    } catch (e) {
    Logger.log('resolveDigiumAccountIdsDetailed_ error: ' + e.toString());
    const fallbackList = Array.isArray(extensions) ? extensions : [String(extensions)];
    details.list = fallbackList.map(ext => String(ext || '').trim()).filter(Boolean);
    details.list.forEach(ext => { if (ext) details.map[ext] = [ext]; });
    return finalizeDetails();
  }
}

function resolveDigiumAccountIdsFromExtensions_(extensions) {
  const details = resolveDigiumAccountIdsDetailed_(extensions);
  return details.list;
}

function resolveDigiumExtensionAccountMap_(extensions) {
  if (!extensions) return {};
  const extList = Array.isArray(extensions)
    ? extensions.map(ext => String(ext || '').trim()).filter(Boolean)
    : [String(extensions || '').trim()].filter(Boolean);

  const details = resolveDigiumAccountIdsDetailed_(extList);
  return details.map || {};
}

function aggregateDigiumCallsForExtensions_(startDate, endDate, breakdown, extensions) {
  const extList = Array.isArray(extensions)
    ? extensions.map(ext => String(ext || '').trim()).filter(Boolean)
    : [String(extensions || '').trim()].filter(Boolean);
  if (!extList.length) {
    return { ok: false, reason: 'no_extensions' };
  }

  const breakdownKey = breakdown || 'by_day';
  const params = { breakdown: breakdownKey };
  const accountDetails = resolveDigiumAccountIdsDetailed_(extList);
  const accountMap = accountDetails && accountDetails.map ? accountDetails.map : {};
  const accountIdList = accountDetails && Array.isArray(accountDetails.list)
    ? accountDetails.list.map(id => String(id || '').trim()).filter(Boolean)
    : [];

  if (['by_day', 'by_day_of_week', 'by_hour_of_day'].includes(breakdownKey)) {
    if (!accountIdList.length) {
      return { ok: false, reason: 'no_account_ids' };
    }
    params.account_ids = accountIdList.slice();
  }
  if (extList && extList.length) {
    params.target_extensions = extList.slice();
  }

  const res = fetchDigiumCallReports_(startDate, endDate, params);
  if (!res || !res.ok) {
    return res;
  }

  const cloned = {
    ok: true,
    categories: Array.isArray(res.categories) ? res.categories.slice() : [],
    rows: Array.isArray(res.rows) ? res.rows.map(r => r.slice()) : [],
    fields: res.fields ? res.fields.slice() : undefined,
    humanLabels: res.humanLabels,
    raw: res.raw,
    breakdown: breakdownKey,
    accountMap: accountMap,
    accountIds: accountIdList.slice()
  };
  cloned.totals = computeDigiumTotalsFromRows_(cloned.rows);
  if (breakdownKey === 'by_day' && res.dates) {
    cloned.dates = res.dates.slice();
  }
  return cloned;
}
function computeDigiumTotalsFromRows_(rows) {
  const totals = {
    total_calls: 0,
    total_incoming_calls: 0,
    total_outgoing_calls: 0,
    talking_duration: 0,
    call_duration: 0
  };
  if (!Array.isArray(rows)) return totals;

  const parseValue = (value) => {
    if (value == null || value === '') return 0;
    if (typeof value === 'number') return isNaN(value) ? 0 : value;
    const str = String(value).trim();
    if (!str) return 0;
    if (/^\d{1,2}:\d{2}:\d{2}$/.test(str)) return parseDurationSeconds_(str);
    const num = Number(str);
    return isNaN(num) ? 0 : num;
  };

  rows.forEach(row => {
    if (!row || row.length < 2) return;
    const label = String(row[0] || '').toLowerCase();
    let key = null;
    if (label.includes('total calls') && !label.includes('incoming') && !label.includes('outgoing')) {
      key = 'total_calls';
    } else if (label.includes('total incoming')) {
      key = 'total_incoming_calls';
    } else if (label.includes('total outgoing')) {
      key = 'total_outgoing_calls';
    } else if (label.includes('talking duration')) {
      key = 'talking_duration';
    } else if (label.includes('call duration')) {
      key = 'call_duration';
    }
    if (!key) return;
    for (let i = 1; i < row.length; i++) {
      totals[key] += parseValue(row[i]);
    }
  });

  return totals;
}

function aggregateDigiumCallsByDateForExtensions_(startDate, endDate, extensions) {
  return aggregateDigiumCallsForExtensions_(startDate, endDate, 'by_day', extensions);
}

function aggregateDigiumCallsByHourForExtensions_(startDate, endDate, extensions) {
  const extList = Array.isArray(extensions)
    ? extensions.map(ext => String(ext || '').trim()).filter(Boolean)
    : [String(extensions || '').trim()].filter(Boolean);
  return aggregateDigiumCallsForExtensions_(startDate, endDate, 'by_hour_of_day', extList);
}

function aggregateDigiumCallsByDayOfWeekForExtensions_(startDate, endDate, extensions) {
  const extList = Array.isArray(extensions)
    ? extensions.map(ext => String(ext || '').trim()).filter(Boolean)
    : [String(extensions || '').trim()].filter(Boolean);
  return aggregateDigiumCallsForExtensions_(startDate, endDate, 'by_day_of_week', extList);
}

var DIGIUM_DATASET_CACHE = (typeof DIGIUM_DATASET_CACHE !== 'undefined' && DIGIUM_DATASET_CACHE) || {};

function getDigiumDataset_(startDate, endDate, extensionMetaOpt) {
  const key = `${isoDate_(startDate)}|${isoDate_(endDate)}`;
  if (DIGIUM_DATASET_CACHE[key]) return DIGIUM_DATASET_CACHE[key];
  const dataset = fetchDigiumDataset_(startDate, endDate, extensionMetaOpt);
  DIGIUM_DATASET_CACHE[key] = dataset;
  return dataset;
}

function fetchDigiumDataset_(startDate, endDate, extensionMetaOpt) {
  const extensionMeta = extensionMetaOpt || getActiveExtensionMetadata_();
  const activeExtensions = (extensionMeta && Array.isArray(extensionMeta.list))
    ? extensionMeta.list.map(ext => String(ext || '').trim()).filter(Boolean)
    : [];

  const dataset = {
    startDate: new Date(startDate),
    endDate: new Date(endDate),
    extensions: activeExtensions.slice(),
    extensionMeta
  };

  try {
    dataset.byAccount = fetchDigiumCallReports_(startDate, endDate, {
      breakdown: 'by_account'
    });
  } catch (e) {
    Logger.log('fetchDigiumDataset_: by_account fetch failed: ' + e.toString());
    dataset.byAccount = { ok: false, error: e.toString() };
  }

  dataset.callMetricsByCanonical = buildCallMetricsByCanonical_(dataset.byAccount, extensionMeta);
  dataset.callDailyPerCanonical = buildCallDailyPerCanonical_(startDate, endDate, extensionMeta);
  dataset.byDay = aggregateDigiumCallsByDateForExtensions_(startDate, endDate, activeExtensions);
  dataset.byHour = aggregateDigiumCallsByHourForExtensions_(startDate, endDate, activeExtensions);
  dataset.byDow = aggregateDigiumCallsByDayOfWeekForExtensions_(startDate, endDate, activeExtensions);
  dataset.generatedAt = new Date();
  return dataset;
}
function buildCallMetricsByCanonical_(digTotals, extensionMeta) {
  const perExtension = digTotals && digTotals.perExtension && typeof digTotals.perExtension === 'object'
    ? digTotals.perExtension
    : {};
  if (!perExtension || !Object.keys(perExtension).length) return {};

  const allowedSet = new Set(
    ((extensionMeta && Array.isArray(extensionMeta.list)) ? extensionMeta.list : [])
      .map(ext => String(ext || '').trim())
      .filter(Boolean)
  );
  const extToName = (extensionMeta && extensionMeta.extToName) || {};
  const extMapData = getExtensionMap_() || {};

  const extToCanonical = {};
  const extPriority = {};
  const assignCanonical = (ext, rawName, priority) => {
    const extKey = String(ext || '').trim();
    if (!extKey) return;
    const canonical = canonicalTechnicianName_(rawName);
    if (!canonical) return;
    const currentPriority = extPriority[extKey];
    if (currentPriority == null || priority >= currentPriority) {
      extToCanonical[extKey] = canonical;
      extPriority[extKey] = priority;
    }
  };

  Object.keys(extToName).forEach(ext => assignCanonical(ext, extToName[ext], 3));

  Object.keys(extMapData).forEach(nameKey => {
    const exts = extMapData[nameKey];
    if (!Array.isArray(exts) || !exts.length) return;
    exts.forEach(ext => assignCanonical(ext, nameKey, 2));
  });

  Object.keys(perExtension).forEach(ext => {
    const meta = perExtension[ext];
    if (!meta || !meta.label) return;
    const label = String(meta.label || '');
    const parts = label.split('-').map(s => s.trim()).filter(Boolean);
    if (parts.length >= 2) {
      assignCanonical(ext, parts.slice(1).join(' - '), 1);
    } else if (parts.length === 1) {
      assignCanonical(ext, parts[0], 1);
    }
  });

  const parseMetric = (value) => {
    if (value == null || value === '') return 0;
    if (typeof value === 'number') return isNaN(value) ? 0 : value;
    const str = String(value).trim();
    if (!str) return 0;
    if (/^\d{1,2}:\d{2}:\d{2}$/.test(str)) return parseDurationSeconds_(str);
    const num = Number(str);
    return isNaN(num) ? 0 : num;
  };

  const result = {};
  Object.keys(perExtension).forEach(ext => {
    const extKey = String(ext || '').trim();
    if (!extKey) return;
    if (allowedSet.size && !allowedSet.has(extKey)) return;
    const canonical = extToCanonical[extKey];
    if (!canonical) return;
    const metrics = perExtension[ext].metrics || {};
    const bucket = result[canonical] || {
      totalCalls: 0,
      inboundCalls: 0,
      outboundCalls: 0,
      talkSeconds: 0,
      callSeconds: 0,
      extensions: new Set()
    };
    bucket.totalCalls += parseMetric(metrics.total_calls);
    bucket.inboundCalls += parseMetric(metrics.total_incoming_calls);
    bucket.outboundCalls += parseMetric(metrics.total_outgoing_calls);
    bucket.talkSeconds += parseMetric(metrics.talking_duration);
    bucket.callSeconds += parseMetric(metrics.call_duration);
    bucket.extensions.add(extKey);
    result[canonical] = bucket;
  });

  Object.keys(result).forEach(key => {
    const bucket = result[key];
    bucket.extensions = Array.from(bucket.extensions || []);
  });

  return result;
}

function buildCallDailyPerCanonical_(startDate, endDate, extensionMeta) {
  const result = {};
  const extToName = extensionMeta && extensionMeta.extToName ? extensionMeta.extToName : {};
  const extensions = Object.keys(extToName || {});
  if (!extensions.length) return result;

  const normalizeDateKey = (value) => {
    const str = String(value || '').trim();
    if (!str) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    const mmdd = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (mmdd) {
      return `${mmdd[3]}-${('0' + mmdd[1]).slice(-2)}-${('0' + mmdd[2]).slice(-2)}`;
    }
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString().split('T')[0];
    }
    return '';
  };

  extensions.forEach(ext => {
    const rawName = extToName[ext];
    if (!rawName) return;
    const canonical = canonicalTechnicianName_(rawName);
    if (!canonical) return;

    const res = aggregateDigiumCallsForExtensions_(startDate, endDate, 'by_day', [ext]);
    if (!res || !res.ok || !res.rows || !res.rows.length) return;

    const rawDates = (res.dates && res.dates.length ? res.dates : (res.categories || [])) || [];
    if (!rawDates.length) return;

    const normalizedDates = rawDates.map(normalizeDateKey).filter(Boolean);
    if (!normalizedDates.length) return;

    const perDay = {};
    normalizedDates.forEach(dateKey => {
      perDay[dateKey] = {
        totalCalls: 0,
        inbound: 0,
        outbound: 0,
        talkSeconds: 0,
        callSeconds: 0
      };
    });

    res.rows.forEach(row => {
      if (!row || row.length < 2) return;
      const label = String(row[0] || '').toLowerCase();
      for (let i = 1; i < row.length && i <= normalizedDates.length; i++) {
        const dateKey = normalizedDates[i - 1];
        if (!dateKey || !perDay[dateKey]) continue;
        const rawVal = row[i] != null ? Number(row[i]) : 0;
        const value = isNaN(rawVal) ? 0 : rawVal;
        if (label.includes('total calls') && !label.includes('incoming') && !label.includes('outgoing')) {
          perDay[dateKey].totalCalls += value;
        } else if (label.includes('incoming')) {
          perDay[dateKey].inbound += value;
        } else if (label.includes('outgoing')) {
          perDay[dateKey].outbound += value;
        } else if (label.includes('talk')) {
          perDay[dateKey].talkSeconds += value;
        } else if (label.includes('call duration')) {
          perDay[dateKey].callSeconds += value;
        }
      }
    });

    const accumulator = result[canonical] || {
      perDay: {},
      totals: {
        totalCalls: 0,
        inbound: 0,
        outbound: 0,
        talkSeconds: 0,
        callSeconds: 0
      },
      dates: []
    };

    normalizedDates.forEach(dateKey => {
      const dayMetrics = perDay[dateKey] || {
        totalCalls: 0,
        inbound: 0,
        outbound: 0,
        talkSeconds: 0,
        callSeconds: 0
      };
      if (!accumulator.perDay[dateKey]) {
        accumulator.perDay[dateKey] = {
          totalCalls: 0,
          inbound: 0,
          outbound: 0,
          talkSeconds: 0,
          callSeconds: 0
        };
      }
      const accDay = accumulator.perDay[dateKey];
      accDay.totalCalls += dayMetrics.totalCalls;
      accDay.inbound += dayMetrics.inbound;
      accDay.outbound += dayMetrics.outbound;
      accDay.talkSeconds += dayMetrics.talkSeconds;
      accDay.callSeconds += dayMetrics.callSeconds;

      accumulator.totals.totalCalls += dayMetrics.totalCalls;
      accumulator.totals.inbound += dayMetrics.inbound;
      accumulator.totals.outbound += dayMetrics.outbound;
      accumulator.totals.talkSeconds += dayMetrics.talkSeconds;
      accumulator.totals.callSeconds += dayMetrics.callSeconds;
    });

    accumulator.dates = Array.from(new Set([...(accumulator.dates || []), ...normalizedDates])).sort();
    result[canonical] = accumulator;
  });

  return result;
}
// Create a Digium calls sheet with both by-date and by-account summaries
function createDigiumCallsSheet_(datasetOrByDate, byAccountDataOpt, extensionMetaOpt) {
  const ss = SpreadsheetApp.getActive();
  const name = 'Digium_Calls';
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    sh.clear();
  }

  let currentRow = 1;
  let byDateData = null;
  let byAccountData = null;
  let dataset = null;
  let extensionMeta = extensionMetaOpt;

  if (datasetOrByDate && datasetOrByDate.byAccount) {
    dataset = datasetOrByDate;
    byDateData = dataset.byDay || dataset.byDate || null;
    byAccountData = dataset.byAccount || null;
    extensionMeta = extensionMetaOpt || dataset.extensionMeta || getActiveExtensionMetadata_();
  } else {
    byDateData = datasetOrByDate;
    byAccountData = byAccountDataOpt;
    extensionMeta = extensionMetaOpt || getActiveExtensionMetadata_();
  }

  const allowedList = (extensionMeta && extensionMeta.list) || [];
  const allowedSet = new Set(allowedList.map(ext => String(ext || '').trim()));
  const extToName = (extensionMeta && extensionMeta.extToName) || {};

  const parseDurationSecondsSafe = (value) => {
    if (value == null || value === '') return 0;
    if (typeof value === 'number') {
      return isNaN(value) ? 0 : value;
    }
    const str = String(value).trim();
    if (!str) return 0;
    if (str.includes(':')) return parseDurationSeconds_(str);
    const num = Number(str);
    return isNaN(num) ? 0 : num;
  };

  const writeSectionHeader = (title) => {
    sh.getRange(currentRow, 1).setValue(title);
    sh.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#1E3A8A');
    currentRow += 1;
  };

  const writeTable = (header, rows, durationMetricMatcher) => {
    if (!header || header.length === 0) return;
    sh.getRange(currentRow, 1, 1, header.length).setValues([header]);
    sh.getRange(currentRow, 1, 1, header.length).setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    currentRow += 1;
    if (!rows || !rows.length) return;

    const processedRows = rows.map(row => {
      const label = String(row[0] || '');
      const isDuration = durationMetricMatcher(label);
      const out = [label];
      for (let c = 1; c < header.length; c++) {
        const value = row[c] != null ? row[c] : '';
        if (isDuration) {
          const seconds = parseDurationSecondsSafe(value);
          out.push(seconds / 86400);
        } else if (value !== '') {
          const num = Number(value);
          out.push(!isNaN(num) ? num : value);
        } else {
          out.push(value);
        }
      }
      return out;
    });

    sh.getRange(currentRow, 1, processedRows.length, header.length).setValues(processedRows);

    processedRows.forEach((row, idx) => {
      const metricName = String(row[0] || '').toLowerCase();
      const isDuration = durationMetricMatcher(metricName);
      const range = sh.getRange(currentRow + idx, 2, 1, header.length - 1);
      try {
        if (isDuration) {
          range.setNumberFormat('hh:mm:ss');
        } else if (/total|calls/.test(metricName)) {
          range.setNumberFormat('0');
        } else if (metricName.startsWith('avg')) {
          range.setNumberFormat('0.0');
        }
      } catch (e) { /* ignore */ }
    });

    currentRow += processedRows.length + 1;
  };

  const isDurationMetric = (metric) => {
    const lower = String(metric || '').toLowerCase();
    return /duration|talking|wait|time/.test(lower);
  };

  // Section 1: By-date data (wide format: Metric | date1 | date2 | ...)
  if (byDateData && byDateData.rows && byDateData.rows.length) {
    writeSectionHeader('Digium Calls by Date');
    const header = ['Metric'].concat(byDateData.dates || []);
    const rawRows = byDateData.rows.map(r => r.slice());

    const filteredRows = [];
    let callDurationRow = null;
    let totalCallsRow = null;
    let avgCallDurationRow = null;

    rawRows.forEach(row => {
      const label = String(row[0] || '');
      const labelLower = label.toLowerCase();
      if (/talking duration/i.test(label)) {
        return; // remove talking duration rows
      }
      if (/call duration/i.test(label)) {
        callDurationRow = row;
        for (let c = 1; c < row.length; c++) {
          row[c] = parseDurationSecondsSafe(row[c]);
        }
      }
      if (/total calls/i.test(label)) {
        totalCallsRow = row;
        for (let c = 1; c < row.length; c++) {
          row[c] = Number(row[c]) || 0;
        }
      }
      if (/avg call duration/i.test(labelLower)) {
        avgCallDurationRow = row;
        for (let c = 1; c < row.length; c++) {
          row[c] = parseDurationSecondsSafe(row[c]);
        }
      }
      filteredRows.push(row);
    });

    if (callDurationRow && totalCallsRow) {
      const avgValues = ['Avg Call Time per Call'];
      for (let c = 1; c < header.length; c++) {
        const durationSeconds = parseDurationSecondsSafe(callDurationRow[c]);
        const totalCalls = Number(totalCallsRow[c]) || 0;
        const avgSeconds = totalCalls > 0 ? durationSeconds / totalCalls : 0;
        avgValues.push(avgSeconds);
      }
      if (avgCallDurationRow) {
        avgCallDurationRow[0] = 'Avg Call Time per Call';
        for (let c = 1; c < header.length; c++) {
          avgCallDurationRow[c] = avgValues[c] || 0;
        }
      } else {
        const baseIndex = filteredRows.indexOf(callDurationRow);
        const insertIndex = baseIndex >= 0 ? baseIndex + 1 : filteredRows.length;
        filteredRows.splice(insertIndex, 0, avgValues);
      }
    }

    writeTable(header, filteredRows, isDurationMetric);
  }

  // Section 2: By-account data (Metric | Totals | ext1 | ext2 | ...)
  if (byAccountData && byAccountData.perExtension) {
    writeSectionHeader('Digium Calls by Account');
    const perExtension = byAccountData.perExtension || {};
    const allExtensions = Object.keys(perExtension);
    const filteredExtensions = allowedSet.size > 0
      ? allExtensions.filter(ext => allowedSet.has(String(ext || '').trim()))
      : allExtensions;
    filteredExtensions.sort();
    const header = ['Metric', 'Totals'].concat(filteredExtensions.map(ext => {
      const meta = perExtension[ext];
      const label = extToName[ext] || (meta && meta.label) || ext;
      return `${label} (${ext})`;
    }));

    const fields = byAccountData.fields || Object.keys(byAccountData.totalsAll || {});
    const human = byAccountData.humanLabels || {};

    const totalsAll = byAccountData.totalsAll || {};
    const totalCallsOverall = Number(totalsAll.total_calls) || 0;
    const totalCallDurationSecondsOverall = parseDurationSecondsSafe(totalsAll.call_duration);
    const totalTalkingSecondsOverall = parseDurationSecondsSafe(totalsAll.talking_duration);

    const filteredTotals = {};
    const hasFilter = filteredExtensions.length > 0;
    if (hasFilter) {
      filteredExtensions.forEach(ext => {
        const meta = perExtension[ext];
        const metrics = meta && meta.metrics ? meta.metrics : {};
        fields.forEach(field => {
          const fieldKey = String(field || '').toLowerCase();
          if (fieldKey.startsWith('avg_')) return;
          let value = metrics[field];
          if (typeof value === 'string' && value.includes(':')) {
            value = parseDurationSecondsSafe(value);
          } else {
            value = Number(value) || 0;
          }
          filteredTotals[fieldKey] = (filteredTotals[fieldKey] || 0) + value;
        });
      });
    }

    const rows = fields.map(field => {
      const fieldKey = String(field || '').toLowerCase();
      let label = human[field] || field;
      if (fieldKey.includes('avg_call_duration')) label = 'Avg Call Time per Call';
      const row = [label, 0];

      filteredExtensions.forEach(ext => {
        const meta = perExtension[ext];
        const metrics = meta && meta.metrics ? meta.metrics : {};
        let value = metrics[field];
        if (typeof value === 'string' && value.includes(':')) {
          value = parseDurationSecondsSafe(value);
        } else {
          value = Number(value) || 0;
        }
        row.push(value);
      });

      if (hasFilter) {
        if (fieldKey === 'avg_call_duration' || fieldKey === 'avg_call_duration (s)') {
          const calls = filteredTotals['total_calls'] || 0;
          const seconds = filteredTotals['call_duration'] || filteredTotals['call_duration (s)'] || 0;
          row[1] = calls > 0 ? seconds / calls : 0;
        } else if (fieldKey === 'avg_talking_duration' || fieldKey === 'avg_talking_duration (s)') {
          const calls = filteredTotals['total_calls'] || 0;
          const seconds = filteredTotals['talking_duration'] || filteredTotals['talking_duration (s)'] || 0;
          row[1] = calls > 0 ? seconds / calls : 0;
        } else if (fieldKey === 'call_duration' || fieldKey === 'call_duration (s)' || fieldKey === 'talking_duration' || fieldKey === 'talking_duration (s)' || fieldKey === 'total_calls' || fieldKey === 'total_incoming_calls' || fieldKey === 'total_outgoing_calls') {
          row[1] = filteredTotals[fieldKey] || 0;
        } else {
          row[1] = filteredTotals[fieldKey] != null ? filteredTotals[fieldKey] : row.slice(1).reduce((sum, val, idx) => idx >= 0 ? sum + (Number(val) || 0) : sum, 0);
        }
      } else {
        const totalRaw = totalsAll[field];
        if (fieldKey === 'avg_call_duration' || fieldKey === 'avg_call_duration (s)') {
          row[1] = totalCallsOverall > 0 ? totalCallDurationSecondsOverall / totalCallsOverall : 0;
        } else if (fieldKey === 'avg_talking_duration' || fieldKey === 'avg_talking_duration (s)') {
          row[1] = totalCallsOverall > 0 ? totalTalkingSecondsOverall / totalCallsOverall : 0;
        } else if (totalRaw != null && totalRaw !== '') {
          if (isDurationMetric(label)) {
            row[1] = parseDurationSecondsSafe(totalRaw);
          } else {
            row[1] = Number(totalRaw) || 0;
          }
        } else {
          row[1] = row.slice(1).reduce((sum, val, idx) => idx >= 0 ? sum + (Number(val) || 0) : sum, 0);
        }
      }
      return row;
    });

    writeTable(header, rows, isDurationMetric);
  }
  try { sh.setFrozenRows(2); } catch (e) {}
  try { sh.autoResizeColumns(1, Math.max(1, sh.getLastColumn())); } catch (e) {}
  return sh;
}

// Append Digium request/response in one batched write to reduce calls
// Note: This function is called multiple times per pull, so we clear the sheet only once at the start of each pull
function appendDigiumRaw_(requestXml, label, raw) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Digium_Raw');
    if (!sh) {
      sh = ss.insertSheet('Digium_Raw');
      // New sheet is empty, no need to clear
    }
    // Check if this is the first append in this pull (sheet is empty or only has old data)
    // We'll clear it at the start of each pull in ingestTimeRangeToSheets_
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
// Initialize extension_map sheet with default data if it doesn't exist
function initializeExtensionMapSheet_() {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('extension_map');
    
    if (!sh) {
      // Create the sheet
      sh = ss.insertSheet('extension_map');
      Logger.log('Created extension_map sheet');
    } else {
      // Check if sheet has data - if it does, don't overwrite
      const dataRange = sh.getDataRange();
      if (dataRange.getNumRows() > 1) {
        Logger.log('extension_map sheet already has data, skipping initialization');
        return;
      }
    }
    
    // Set headers
    const headers = ['technician_name', 'technician_email', 'extension', 'account_id'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E5E7EB');
    
    // Default data provided by user
    const defaultData = [
      ['Mordi Turgeman', 'mordi@novapointofsale.com', '252', '1381'],
      ['Oscar Ocampo', 'oscaro@novapointofsale.com', '304', '1442'],
      ['Eduardo Brambila', 'eduardo@novapointofsale.com', '305', '1430'],
      ['Eddie Talal', 'eddie@novapointofsale.com', '306', '1436'],
      ['Oscar Umana', 'oscar@novapointofsale.com', '308', '1421'],
      ['Tomer', 'tomer@novapointofsale.com', '322', '1423'],
      ['Darius Parlor', 'darius@novapointofsale.com', '355,356', '1439,1351']
    ];
    
    // Write default data
    if (defaultData.length > 0) {
      sh.getRange(2, 1, defaultData.length, headers.length).setValues(defaultData);
    }
    
    // Set column widths
    sh.setColumnWidth(1, 200);
    sh.setColumnWidth(2, 250);
    sh.setColumnWidth(3, 100);
    sh.setColumnWidth(4, 120);
    
    Logger.log(`Initialized extension_map sheet with ${defaultData.length} technicians`);
  } catch (e) {
    Logger.log('initializeExtensionMapSheet_ failed: ' + e.toString());
  }
}
// Read extension_map sheet and return { technician_name_lower: [ext1, ext2, ...] }
function getExtensionMap_() {
  const out = {};
  const accountIdsByKey = {};
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('extension_map');
    if (!sh) {
      Logger.log('getExtensionMap_: extension_map sheet not found, initializing...');
      initializeExtensionMapSheet_();
      sh = ss.getSheetByName('extension_map');
      if (!sh) {
        Logger.log('getExtensionMap_: Failed to create extension_map sheet');
        return out;
      }
    }
    const rng = sh.getDataRange();
    const vals = rng.getValues();
    if (!vals || vals.length < 2) {
      Logger.log('getExtensionMap_: extension_map sheet has no data rows');
      return out;
    }
    const headers = vals[0].map(h => String(h || '').trim().toLowerCase());
    const nameIdx = headers.indexOf('technician_name');
    const extIdx = headers.indexOf('extension');
    const accountIdx = headers.indexOf('account_id');
    const activeIdx = headers.indexOf('is_active');
    
    Logger.log(`getExtensionMap_: Found headers - nameIdx=${nameIdx}, extIdx=${extIdx}, accountIdx=${accountIdx}, activeIdx=${activeIdx}`);
    Logger.log(`getExtensionMap_: All headers: ${headers.join(', ')}`);
    
    if (nameIdx < 0 || extIdx < 0) {
      Logger.log('getExtensionMap_: Required columns (technician_name, extension) not found');
      return out;
    }
    
    // Helper to normalize names for matching (handles variations like "Tomer" vs "Tomer Reiter")
    const normalizeName = (name) => {
      const normalized = String(name || '').trim().toLowerCase();
      // Extract first name for partial matching
      const firstName = normalized.split(/\s+/)[0];
      return { full: normalized, first: firstName };
    };
    
    const addMappingForKey = (key, ext, accountList) => {
      if (!key) return;
      if (!out[key]) out[key] = [];
      if (!out[key].includes(ext)) out[key].push(ext);
      if (accountList && accountList.length) {
        if (!accountIdsByKey[key]) accountIdsByKey[key] = [];
        accountList.forEach(id => {
          const idStr = String(id || '').trim();
          if (!idStr) return;
          if (!accountIdsByKey[key].includes(idStr)) accountIdsByKey[key].push(idStr);
        });
      }
    };
    
    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const name = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
      const extCell = extIdx >= 0 ? String(row[extIdx] || '').trim() : '';
      const accountCell = accountIdx >= 0 ? String(row[accountIdx] || '').trim() : '';
      const isActive = activeIdx >= 0 ? String(row[activeIdx] || '').toLowerCase() !== 'false' : true;
      if (!name || !extCell) continue;
      if (!isActive) {
        Logger.log(`getExtensionMap_: Skipping inactive technician: ${name}`);
        continue;
      }
      const exts = extCell.split(',').map(s => s.trim()).filter(Boolean);
      const accountPartsRaw = accountCell ? accountCell.split(',').map(s => s.trim()).filter(Boolean) : [];
      const nameNorm = normalizeName(name);
      exts.forEach((ext, idx) => {
        const acctValue = accountPartsRaw.length > 1 ? (accountPartsRaw[idx] || accountPartsRaw[0] || '') : (accountPartsRaw[0] || '');
        let acctList = acctValue ? acctValue.split(/[|;]/).map(s => s.trim()).filter(Boolean) : [];
        if (!acctList.length) {
          const overrides = DIGIUM_EXTENSION_ACCOUNT_OVERRIDES[ext] || [];
          acctList = overrides.map(id => String(id || '').trim()).filter(Boolean);
        }
        addMappingForKey(nameNorm.full, ext, acctList);
            if (nameNorm.first && nameNorm.first !== nameNorm.full) {
          addMappingForKey(nameNorm.first, ext, acctList);
            }
            if (nameNorm.full === 'eddie talal' || nameNorm.full === 'ahmed talal') {
              const otherName = nameNorm.full === 'eddie talal' ? 'ahmed talal' : 'eddie talal';
          addMappingForKey(otherName, ext, acctList);
              Logger.log(`getExtensionMap_: Mapped ${name} -> also mapped to ${otherName} (same person)`);
            }
        Logger.log(`getExtensionMap_: Mapped ${name} -> extension ${ext}${acctList.length ? ' (account_ids: ' + acctList.join(', ') + ')' : ''}`);
      });
    }
    const ensureManualMapping = (techName, extensions, accountIdsOpt) => {
      const nameNorm = normalizeName(techName);
      const keys = [nameNorm.full];
      if (nameNorm.first && nameNorm.first !== nameNorm.full) keys.push(nameNorm.first);
      const accountIds = Array.isArray(accountIdsOpt) ? accountIdsOpt : [];
      extensions.map(ext => String(ext || '').trim()).filter(Boolean).forEach((ext, idx) => {
        const acctIdsForExt = accountIds.length > 1 ? (accountIds[idx] ? [accountIds[idx]] : accountIds.filter(Boolean)) : accountIds.filter(Boolean);
        keys.forEach(key => addMappingForKey(key, ext, acctIdsForExt));
      });
    };
    ensureManualMapping('Mordi Turgeman', ['252'], ['1381']);
    ensureManualMapping('Oscar Ocampo', ['304'], ['1442']);
    ensureManualMapping('Eduardo Brambila', ['305'], ['1430']);
    ensureManualMapping('Eddie Talal', ['306'], ['1436']);
    ensureManualMapping('Ahmed Talal', ['306'], ['1436']);
    ensureManualMapping('Oscar Umana', ['308'], ['1421']);
    ensureManualMapping('Tomer', ['322'], ['1423']);
    ensureManualMapping('Darius Parlor', ['355','356'], ['1439','1351']);
    Logger.log(`getExtensionMap_: Total technicians mapped: ${Object.keys(out).length}`);
  } catch (e) { 
    Logger.log('getExtensionMap_ failed: ' + e.toString());
    Logger.log('getExtensionMap_ error stack: ' + (e.stack || 'no stack trace'));
  }
  try {
    Object.defineProperty(out, '__accountIds', { value: accountIdsByKey, enumerable: false, configurable: true });
  } catch (e) {
    out.__accountIds = accountIdsByKey;
  }
  return out;
}
function getActiveExtensionMetadata_() {
  const metadata = { list: [], extToName: {}, extToAccountIds: {}, accountIdList: [] };
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('extension_map');
    if (!sh) return metadata;
    const vals = sh.getDataRange().getValues();
    if (!vals || vals.length < 2) return metadata;
    const headers = vals[0].map(h => String(h || '').trim().toLowerCase());
    const nameIdx = headers.indexOf('technician_name');
    const extIdx = headers.indexOf('extension');
    const accountIdx = headers.indexOf('account_id');
    const activeIdx = headers.indexOf('is_active');
    if (nameIdx < 0 || extIdx < 0) return metadata;
    const seen = new Set();
    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const name = String(row[nameIdx] || '').trim();
      if (!name) continue;
      const isActive = activeIdx >= 0 ? String(row[activeIdx] || '').toLowerCase() !== 'false' : true;
      if (!isActive) continue;
      const extCell = String(row[extIdx] || '').trim();
      if (!extCell) continue;
      const exts = extCell.split(',').map(s => s.trim()).filter(Boolean);
      const accountCell = accountIdx >= 0 ? String(row[accountIdx] || '').trim() : '';
      const accountParts = accountCell ? accountCell.split(',').map(s => s.trim()).filter(Boolean) : [];
      exts.forEach((ext, idx) => {
        if (!ext) return;
        if (!seen.has(ext)) {
          metadata.list.push(ext);
          seen.add(ext);
        }
        if (!metadata.extToName[ext]) {
          metadata.extToName[ext] = name;
        }
        const acctId = accountParts.length > 1 ? (accountParts[idx] || '') : (accountParts[0] || '');
        if (acctId) {
          const acctArr = acctId.split(/[|;]/).map(s => s.trim()).filter(Boolean);
          if (!metadata.extToAccountIds[ext]) metadata.extToAccountIds[ext] = [];
          acctArr.forEach(id => {
            if (!metadata.extToAccountIds[ext].includes(id)) metadata.extToAccountIds[ext].push(id);
            if (!metadata.accountIdList.includes(id)) metadata.accountIdList.push(id);
          });
        }
      });
    }
    const ensureManualExtension = (ext, name) => {
      const extStr = String(ext || '').trim();
      if (!extStr) return;
      if (!metadata.list.includes(extStr)) metadata.list.push(extStr);
      if (!metadata.extToName[extStr]) metadata.extToName[extStr] = name;
    };
    ensureManualExtension('355', 'Darius Parlor');
    ensureManualExtension('356', 'Darius Parlor');

    Object.keys(DIGIUM_EXTENSION_ACCOUNT_OVERRIDES).forEach(ext => {
      ensureManualExtension(ext, metadata.extToName[ext] || ext);
      const overrideIds = DIGIUM_EXTENSION_ACCOUNT_OVERRIDES[ext] || [];
      if (!metadata.extToAccountIds[ext]) metadata.extToAccountIds[ext] = [];
      overrideIds.forEach(id => {
        const idStr = String(id || '').trim();
        if (!idStr) return;
        if (!metadata.extToAccountIds[ext].includes(idStr)) metadata.extToAccountIds[ext].push(idStr);
        if (!metadata.accountIdList.includes(idStr)) metadata.accountIdList.push(idStr);
      });
    });

    // Deduplicate account ids per extension and global list
    Object.keys(metadata.extToAccountIds).forEach(ext => {
      const uniq = [];
      const seenIds = new Set();
      metadata.extToAccountIds[ext].forEach(id => {
        const idStr = String(id || '').trim();
        if (!idStr || seenIds.has(idStr)) return;
        seenIds.add(idStr);
        uniq.push(idStr);
      });
      metadata.extToAccountIds[ext] = uniq;
    });
    const globalSeen = new Set();
    metadata.accountIdList = metadata.accountIdList.filter(id => {
      const idStr = String(id || '').trim();
      if (!idStr || globalSeen.has(idStr)) return false;
      globalSeen.add(idStr);
      return true;
    });
  } catch (e) {
    Logger.log('getActiveExtensionMetadata_ failed: ' + e.toString());
  }
  return metadata;
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
    
    const extensionMeta = getActiveExtensionMetadata_();
    const digiumDataset = getDigiumDataset_(startDate, endDate, extensionMeta);
    createDigiumCallsSheet_(digiumDataset, extensionMeta);
      SpreadsheetApp.getActive().toast('Digium response parsed and saved to Digium_Calls');
      return { ok: true };
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
  const r = apiGet_(base, 'setReportArea.aspx', {area: '4'}, cookie, 2, true);
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
// Response format: "OK" or "OK OUTPUT:XML" or "OK OUTPUT:TEXT"
function setOutputXMLOrFallback_(base, cookie) {
  if (FORCE_TEXT_OUTPUT) {
    const rt = apiGet_(base, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
    const tt = (rt.getContentText() || '').trim();
    if (!/^OK/i.test(tt)) throw new Error(`setOutput TEXT failed: ${tt}`);
    // Verify it was set
    const verified = getOutput_(base, cookie);
    if (verified !== 'TEXT') {
      Logger.log(`Warning: setOutput TEXT returned OK but getOutput returned ${verified}`);
    }
    return 'TEXT';
  }
  try {
    const rx = apiGet_(base, 'setOutput.aspx', {output: 'XML'}, cookie, 2, true);
    const tx = (rx.getContentText()||'').trim();
    if (!/^OK/i.test(tx)) {
      Logger.log(`setOutput XML failed: ${tx}, falling back to TEXT`);
      throw new Error(tx);
    }
    // Verify XML was actually set using getOutput
    const verified = getOutput_(base, cookie);
    if (verified !== 'XML') {
      Logger.log(`Warning: setOutput XML returned OK but getOutput returned ${verified}, retrying...`);
      // Retry once
      const rx2 = apiGet_(base, 'setOutput.aspx', {output: 'XML'}, cookie, 2, true);
      const tx2 = (rx2.getContentText()||'').trim();
      if (!/^OK/i.test(tx2)) throw new Error(`setOutput XML retry failed: ${tx2}`);
      const verified2 = getOutput_(base, cookie);
      if (verified2 !== 'XML') {
        Logger.log(`setOutput XML still not verified after retry (got ${verified2}), falling back to TEXT`);
        throw new Error(`XML output not verified: ${verified2}`);
      }
      Logger.log('XML output verified after retry');
      return 'XML';
    }
    Logger.log('XML output verified');
    return 'XML';
  } catch (e) {
    Logger.log(`setOutput XML failed: ${e.toString()}, falling back to TEXT`);
    const rt = apiGet_(base, 'setOutput.aspx', {output: 'TEXT'}, cookie, 2, true);
    const tt = (rt.getContentText()||'').trim();
    if (!/^OK/i.test(tt)) throw new Error(`setOutput fallback TEXT failed: ${tt}`);
    return 'TEXT';
  }
}

function getReportTry_(base, cookie, nodeId, noderef) {
  const maxAttempts = 4;
  let delayMs = 500;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
  try {
    const r = apiGet_(base, 'getReport.aspx', { 
      node: String(nodeId), 
      noderef: noderef
    }, cookie, 4, true);
    const t = (r.getContentText()||'').trim();
      if (/POLLRATEEXCEEDED/i.test(t)) {
        Logger.log(`getReportTry_ node ${nodeId} (${noderef}): POLLRATEEXCEEDED (attempt ${attempt}/${maxAttempts})`);
        if (attempt < maxAttempts) {
          Utilities.sleep(delayMs);
          delayMs = Math.min(delayMs * 2, 4000);
          continue;
        }
        return null;
      }
    // Accept both TEXT (starts with OK) and XML (starts with '<') responses
    if (/^OK/i.test(t)) {
      Logger.log(`getReportTry_ node ${nodeId} (${noderef}): TEXT format response (${t.length} chars)`);
      return t;
    }
    if (t && t[0] === '<') {
      Logger.log(`getReportTry_ node ${nodeId} (${noderef}): XML format response (${t.length} chars)`);
      return t;
    }
    Logger.log(`getReportTry_ node ${nodeId} (${noderef}): Unexpected response format (first 100 chars: ${t.substring(0, 100)})`);
    return null;
  } catch (e) {
      Logger.log(`getReportTry_ failed for node ${nodeId} (${noderef}) attempt ${attempt}/${maxAttempts}: ${e.toString()}`);
      if (attempt >= maxAttempts) return null;
      Utilities.sleep(delayMs);
      delayMs = Math.min(delayMs * 2, 4000);
    }
  }
  return null;
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
  
  // Optimize TEXT parsing for large responses
  const lines = body.split(/\r?\n/).filter(Boolean);
  if (!lines.length) return { headers: [], rows: [] };
  const delim = String(delimiter || '|');
  const header = lines[0].split(delim); // No trim - preserve raw header names from API
  if (!header.length) return { headers: [], rows: [] };
  
  // Limit processing to prevent timeout on extremely large datasets
  const maxRows = 10000; // Process max 10k rows to prevent timeout
  const rowsToProcess = Math.min(lines.length - 1, maxRows);
  if (lines.length - 1 > maxRows) {
    Logger.log(`Warning: Response has ${lines.length - 1} rows, processing first ${maxRows} to prevent timeout`);
  }
  
  const out = [];
  for (let i=1; i<=rowsToProcess; i++) {
    const cols = lines[i].split(delim);
    while (cols.length < header.length) cols.push('');
    const obj = {};
    for (let j=0; j<header.length; j++) {
      obj[header[j]] = cols[j] || ''; // No trim - preserve raw values from API
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
// Per API docs: https://secure.logmeinrescue.com/welcome/webhelp/en/rescueapi/API/API_Rescue_getOutput.html
// Response formats: "OK OUTPUT:XML" or "OK OUTPUT:TEXT" or XML response
function getOutput_(base, cookie) {
  try {
    // Per Rescue API docs, allowed HTTP GET endpoint is /API/getOutput.aspx
    const r = apiGet_(base, 'getOutput.aspx', {}, cookie, 2, true);
    const t = (r.getContentText()||'').trim();
    
    // Response format 1: "OK OUTPUT:XML" or "OK OUTPUT:TEXT"
    let m = t.match(/OUTPUT\s*:\s*(XML|TEXT)/i);
    if (m) {
      const fmt = m[1].toUpperCase();
      Logger.log(`getOutput returned: ${fmt}`);
      return fmt;
    }
    
    // Response format 2: "OK XML" or "OK TEXT"
    m = t.match(/^OK\s+(XML|TEXT)/i);
    if (m) {
      const fmt = m[1].toUpperCase();
      Logger.log(`getOutput returned (format 2): ${fmt}`);
      return fmt;
    }
    
    // Response format 3: XML response (parse XML structure)
    if (t && t[0] === '<') {
      try {
        const doc = XmlService.parse(t);
        const root = doc.getRootElement();
        // Search for any element named 'output' (case-insensitive) and read its text/value
        const stack = [root];
        while (stack.length) {
          const el = stack.pop();
          const name = String(el.getName()||'').toLowerCase();
          if (name === 'output' || name === 'setoutputresult') {
            const val = (el.getText ? el.getText() : '') || 
                       (el.getAttribute && el.getAttribute('value') ? el.getAttribute('value').getValue() : '') ||
                       (el.getAttribute && el.getAttribute('output') ? el.getAttribute('output').getValue() : '');
            if (/^xml$/i.test(val)) {
              Logger.log('getOutput returned XML (from XML response)');
              return 'XML';
            }
            if (/^text$/i.test(val)) {
              Logger.log('getOutput returned TEXT (from XML response)');
              return 'TEXT';
            }
          }
          const kids = el.getChildren();
          for (let i=0;i<kids.length;i++) stack.push(kids[i]);
        }
      } catch (e) { 
        Logger.log('getOutput XML parse failed: ' + e.toString());
      }
    }
    
    Logger.log(`getOutput unexpected format: ${t.substring(0, 100)}`);
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
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) {
      // Sheet only has header, no data rows yet
      sh.appendRow([String(header || ''), String(sample || ''), new Date().toISOString(), new Date().toISOString(), 1]);
      return;
    }
    const dataRange = sh.getRange(2, 1, lastRow - 1, 1);
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
  if (!(dt instanceof Date) || isNaN(dt)) return '';
  const tz = (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function')
    ? Session.getScriptTimeZone()
    : 'Etc/UTC';
  return Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
}

function mdy_(iso) {
  const d = new Date(iso + 'T00:00:00');
  return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`;
}
// Normalize a duration-like value into seconds.
// Accepts: numeric seconds, numeric time-fraction (days), or hh:mm:ss string.
function parseDurationSeconds_(val) {
  if (val == null || val === '') return 0;
  
  // Handle Date objects (Google Sheets converts time fractions to Date objects)
  if (val instanceof Date) {
    // Date objects from Sheets time values - convert to seconds
    // Date object represents days since 1899-12-30, so get the time value
    const timeValue = val.getTime();
    // If it's a date in 1899, it's likely a time fraction (0.00185 days = 2:41)
    // Convert to seconds: (date.getTime() - baseDate.getTime()) / 1000
    const baseDate = new Date(1899, 11, 30); // Dec 30, 1899
    const diffMs = timeValue - baseDate.getTime();
    // If the difference is less than 2 days in milliseconds, it's a time fraction
    if (Math.abs(diffMs) < 2 * 24 * 60 * 60 * 1000) {
      return Math.round(diffMs / 1000); // Convert milliseconds to seconds
    }
    // Otherwise it's a real date, return 0 (not a duration)
    return 0;
  }
  
  // If already a number
  if (typeof val === 'number') {
    // Heuristic: treat values < 1 as day-fractions (Google Sheets time); >=1 assume seconds.
    // This avoids misinterpreting genuine short second durations (e.g. 2 seconds) as multi-minute.
    if (val >= 0 && val < 1) return Math.round(val * 86400);
    return Math.round(val);
  }
  
  // If string, try to parse numeric first
  const s = String(val).trim();
  if (!s) return 0;
  if (/^[0-9]+(\.[0-9]+)?$/.test(s)) {
    const n = Number(s);
    if (n >= 0 && n < 1) return Math.round(n * 86400); // day fraction
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

function createEmptyDayStats_() {
  return {
    sessions: 0,
    durationSum: 0,
    durationCount: 0,
    pickupSum: 0,
    pickupCount: 0,
    workSeconds: 0,
    activeSeconds: 0,
    loginSeconds: 0,
    longestSeconds: 0,
    novaWave: 0,
    techSet: new Set()
  };
}

function buildTechnicianSessionContext_(rows, indexes, tz) {
  const {
    startIdx,
    techIdx,
    durationIdx,
    workIdx,
    pickupIdx,
    activeIdx,
    channelIdx,
    callingCardIdx
  } = indexes || {};

  const perCanonical = {};
  const teamDaily = {};
  const dateSet = new Set();

  const formatDateKey = (dateObj) => {
    try {
      return Utilities.formatDate(dateObj, tz || 'Etc/GMT', 'yyyy-MM-dd');
    } catch (e) {
      try {
        return dateObj.toISOString().split('T')[0];
      } catch (err) {
        return '';
      }
    }
  };

  const ensureDay = (container, dateKey) => {
    if (!container[dateKey]) container[dateKey] = createEmptyDayStats_();
    return container[dateKey];
  };

  const titleize = (txt) => {
    if (!txt) return '';
    return String(txt).split(/\s+/).map(part => part ? part.charAt(0).toUpperCase() + part.slice(1) : '').join(' ').trim();
  };

  const ensureCanonicalEntry = (canonical, rawName) => {
    if (!perCanonical[canonical]) {
      perCanonical[canonical] = {
        displayName: '',
        totals: createEmptyDayStats_(),
        daily: {}
      };
    }
    if (rawName) {
      const candidate = String(rawName).trim();
      if (candidate) {
        const existing = perCanonical[canonical].displayName;
        if (!existing || candidate.length > existing.length) {
          perCanonical[canonical].displayName = candidate;
        }
      }
    }
    return perCanonical[canonical];
  };

  const ensureTeamDay = (dateKey) => ensureDay(teamDaily, dateKey);

  const isNovaWaveRow = (row) => {
    if (!row) return false;
    if (channelIdx != null && channelIdx >= 0 && row[channelIdx]) {
      const ch = String(row[channelIdx]).toLowerCase();
      if (ch.includes('nova wave')) return true;
    }
    if (callingCardIdx != null && callingCardIdx >= 0 && row[callingCardIdx]) {
      const cc = String(row[callingCardIdx]).toLowerCase();
      if (cc.includes('nova wave')) return true;
    }
    return false;
  };

  const toCanonical = (name) => {
    if (!name) return 'unknown';
    return canonicalTechnicianName_(name) || normalizeTechnicianNameFull_(name) || 'unknown';
  };

  (rows || []).forEach(row => {
    if (!row) return;
    if (startIdx == null || startIdx < 0) return;
    const startValue = row[startIdx];
    if (!startValue) return;
    const dateObj = startValue instanceof Date ? startValue : new Date(startValue);
    if (!(dateObj instanceof Date) || isNaN(dateObj)) return;
    const dateKey = formatDateKey(dateObj);
    if (!dateKey) return;

    dateSet.add(dateKey);

    const rawTech = techIdx != null && techIdx >= 0 && row[techIdx]
      ? String(row[techIdx]).trim()
      : 'Unknown';
    const canonical = toCanonical(rawTech);
    const entry = ensureCanonicalEntry(canonical, rawTech);
    const dayStats = ensureDay(entry.daily, dateKey);
    const teamDay = ensureTeamDay(dateKey);

    const techIdentifier = canonical || (rawTech ? rawTech.toLowerCase() : '');
    if (techIdentifier) {
      if (!dayStats.techSet) dayStats.techSet = new Set();
      dayStats.techSet.add(techIdentifier);
      if (!teamDay.techSet) teamDay.techSet = new Set();
      teamDay.techSet.add(techIdentifier);
    }

    entry.totals.sessions++;
    dayStats.sessions++;
    teamDay.sessions++;

    const addDuration = (seconds) => {
      if (!seconds) return;
      entry.totals.durationSum += seconds;
      entry.totals.durationCount++;
      entry.totals.loginSeconds += seconds;
      if (seconds > entry.totals.longestSeconds) entry.totals.longestSeconds = seconds;
      dayStats.durationSum += seconds;
      dayStats.durationCount++;
      dayStats.loginSeconds += seconds;
      if (seconds > dayStats.longestSeconds) dayStats.longestSeconds = seconds;
      teamDay.durationSum += seconds;
      teamDay.durationCount++;
      teamDay.loginSeconds += seconds;
      if (seconds > teamDay.longestSeconds) teamDay.longestSeconds = seconds;
    };

    const addPickup = (seconds) => {
      if (!seconds) return;
      entry.totals.pickupSum += seconds;
      entry.totals.pickupCount++;
      dayStats.pickupSum += seconds;
      dayStats.pickupCount++;
      teamDay.pickupSum += seconds;
      teamDay.pickupCount++;
    };

    if (durationIdx != null && durationIdx >= 0 && row[durationIdx]) {
      const sec = parseDurationSeconds_(row[durationIdx]);
      if (sec > 0) addDuration(sec);
    }

    if (pickupIdx != null && pickupIdx >= 0 && row[pickupIdx]) {
      const pickSec = parseDurationSeconds_(row[pickupIdx]);
      if (pickSec > 0) addPickup(pickSec);
    }

    if (workIdx != null && workIdx >= 0 && row[workIdx]) {
      const workSec = parseDurationSeconds_(row[workIdx]);
      if (workSec > 0) {
        entry.totals.workSeconds += workSec;
        dayStats.workSeconds += workSec;
        teamDay.workSeconds += workSec;
      }
    }

    if (activeIdx != null && activeIdx >= 0 && row[activeIdx]) {
      const activeSec = parseDurationSeconds_(row[activeIdx]);
      if (activeSec > 0) {
        entry.totals.activeSeconds += activeSec;
        dayStats.activeSeconds += activeSec;
        teamDay.activeSeconds += activeSec;
      }
    }

    if (isNovaWaveRow(row)) {
      entry.totals.novaWave++;
      dayStats.novaWave++;
      teamDay.novaWave++;
    }
  });

  const orderedDates = Array.from(dateSet).sort();
  orderedDates.forEach(dateKey => {
    ensureDay(teamDaily, dateKey);
  });

  Object.keys(perCanonical).forEach(canonical => {
    const entry = perCanonical[canonical];
    if (!entry.displayName) {
      entry.displayName = canonical && canonical !== 'unknown'
        ? titleize(canonical)
        : 'Unknown';
    }
  });

  return {
    perCanonical,
    teamDaily,
    orderedDates
  };
}

/* ===== Sheets Storage ===== */
function getOrCreateSessionsSheet_(ss) {
  // Just get or create the sheet - don't touch headers, order, or formatting
  // Headers will be set by writeRowsToSheets_ from actual API response
  let sh = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
  if (!sh) {
    sh = ss.insertSheet(SHEETS_SESSIONS_TABLE);
    Logger.log('Created Sessions sheet (headers will be set from API data)');
  }
  return sh;
}

// Helper: Get column letter(s) by header name in Sessions sheet (for formulas)
function getSessionsColumnByHeader_(headerVariants) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return 'A'; // Fallback
    
    const headers = sessionsSheet.getRange(1, 1, 1, sessionsSheet.getLastColumn()).getValues()[0];
    for (const variant of headerVariants) {
      const idx = headers.findIndex(h => String(h || '').toLowerCase().trim() === String(variant).toLowerCase().trim());
      if (idx >= 0) {
        // Convert 0-based index to column letter (A=1, B=2, etc.)
        const colNum = idx + 1;
        let colLetter = '';
        let temp = colNum;
        while (temp > 0) {
          const remainder = (temp - 1) % 26;
          colLetter = String.fromCharCode(65 + remainder) + colLetter;
          temp = Math.floor((temp - 1) / 26);
        }
        return colLetter;
      }
    }
    return 'A'; // Fallback to column A if header not found
  } catch (e) {
    Logger.log('getSessionsColumnByHeader_ error: ' + e.toString());
    return 'A';
  }
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
function writeRowsToSheets_(ss, rawDataChunks, clearExisting = false) {
  if (!rawDataChunks || !rawDataChunks.length) return 0;
  const sh = getOrCreateSessionsSheet_(ss);
  
  // IMPORTANT: Clear any row groups, filters, and unhide all rows in Sessions sheet
  // This ensures no rows are hidden/collapsed in the Sessions sheet
  try {
    const lastRow = Math.max(sh.getLastRow(), 1);
    const lastCol = Math.max(sh.getLastColumn(), 1);
    
    if (lastRow > 1) {
      // Remove any row groups that might exist (shift depth to -8 to remove all groups)
      sh.getRange(1, 1, lastRow, 1).shiftRowGroupDepth(-8);
      Logger.log(`Cleared any row groups from Sessions sheet (${lastRow} rows)`);
      
      // Unhide all rows (in case any are hidden)
      try {
        sh.showRows(1, lastRow);
        Logger.log(`Unhid all rows in Sessions sheet`);
      } catch (e) {
        Logger.log(`Warning: Could not unhide rows: ${e.toString()}`);
      }
      
      // Remove any filters that might be hiding rows
      try {
        const filter = sh.getFilter();
        if (filter) {
          // Remove filter criteria that might hide rows
          // First, try to remove the filter entirely
          filter.remove();
          Logger.log(`Removed filter from Sessions sheet`);
        }
      } catch (e) {
        // Filter might not exist or might be in use - try to recreate it cleanly
        try {
          const filter = sh.getFilter();
          if (filter) {
            // Clear all filter criteria
            const range = filter.getRange();
            filter.remove();
            // Recreate filter without any criteria
            range.createFilter();
            Logger.log(`Recreated filter on Sessions sheet without criteria`);
          }
        } catch (e2) {
          Logger.log(`Warning: Could not manage filter: ${e2.toString()}`);
        }
      }
    }
  } catch (e) {
    Logger.log(`Warning: Failed to clear row groups/filters from Sessions sheet: ${e.toString()}`);
  }
  
  // Clear existing data if requested
  if (clearExisting) {
    const dataRange = sh.getDataRange();
    if (dataRange.getNumRows() > 1) {
      const numRows = dataRange.getNumRows() - 1;
      const numCols = dataRange.getNumColumns();
      sh.getRange(2, 1, numRows, numCols).clearContent();
      Logger.log('Cleared existing Sessions data (kept headers)');
    }
  }
  
  // Use first chunk's headers as master order (preserves API header order)
  // Then collect any additional headers from other chunks that might not be in first chunk
  let allHeaders = [];
  const allHeadersSet = new Set();
  
  // Get headers from first chunk to preserve API order
  if (rawDataChunks[0] && rawDataChunks[0].headers && Array.isArray(rawDataChunks[0].headers)) {
    allHeaders = rawDataChunks[0].headers; // Use exact order from API
    rawDataChunks[0].headers.forEach(h => allHeadersSet.add(h));
  }
  
  // Collect any additional headers from other chunks (append to end)
  rawDataChunks.slice(1).forEach(chunk => {
    if (chunk.headers && Array.isArray(chunk.headers)) {
      chunk.headers.forEach(h => {
        if (!allHeadersSet.has(h)) {
          allHeaders.push(h); // Append new headers to end
          allHeadersSet.add(h);
        }
      });
    }
  });
  
  if (!allHeaders.length) {
    Logger.log('No headers found in raw data chunks');
    return 0;
  }
  
  // Identify which column header contains "Nova Point of Sale" values (company name)
  // We'll keep the header but replace "Nova Point of Sale" values with empty strings
  // Exclude ID columns (Technician ID, Session ID, etc.) from company name detection
  const idColumnPatterns = ['id', 'technician id', 'session id', 'tracking id', 'device id', 'channel id'];
  let headerWithCompanyName = null;
  for (const chunk of rawDataChunks) {
    if (chunk.rows && Array.isArray(chunk.rows) && chunk.rows.length > 0) {
      // Check first row to find which column has "Nova Point of Sale"
      const firstRow = chunk.rows[0];
      for (const header of allHeaders) {
        // Skip ID columns
        const headerLower = String(header || '').toLowerCase().trim();
        const isIdColumn = idColumnPatterns.some(pattern => headerLower.includes(pattern));
        if (isIdColumn) continue;
        
        const value = String(firstRow[header] || '').trim();
        if (value && value.toLowerCase().includes('nova point of sale')) {
          headerWithCompanyName = header;
          break;
        }
      }
      if (headerWithCompanyName) break;
    }
  }
  
  if (headerWithCompanyName) {
    Logger.log(`Found column "${headerWithCompanyName}" containing company name - will replace "Nova Point of Sale" with empty values`);
  }
  if (!allHeaders.length) {
    Logger.log('No headers found in raw data chunks');
    return 0;
  }
  
  // Write headers if sheet is empty (no headers exist yet)
  const hasData = sh.getLastRow() > 0;
  if (!hasData) {
    sh.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
    Logger.log(`Wrote headers from API (${allHeaders.length} columns) in API order`);
  }
  
  // Flatten all rows from all chunks - preserve original types (numbers, strings, nulls)
  // Replace "Nova Point of Sale" values with empty string in the identified column
  // NO DEDUPLICATION - write all rows as-is since we're only calling the correct channel once
  const allRows = [];
  
  rawDataChunks.forEach(chunk => {
    if (chunk.rows && Array.isArray(chunk.rows)) {
      chunk.rows.forEach(row => {
        // Create row array matching header order
        const rowArray = allHeaders.map(header => {
          let value = row[header];
          // If this is the column with company name and value is "Nova Point of Sale", replace with empty string
          if (headerWithCompanyName && header === headerWithCompanyName) {
            const valueStr = String(value || '').trim();
            if (valueStr.toLowerCase().includes('nova point of sale')) {
              value = '';
            }
          }
          // Return value as-is: numbers stay numbers, strings stay strings, null stays null
          return value;
        });
        
        // Add all rows - no deduplication
        allRows.push(rowArray);
      });
    }
  });
  
  if (!allRows.length) {
    Logger.log('No rows to write after flattening chunks');
    return 0;
  }
  
  // If we cleared existing data, start writing at row 2 (after headers)
  // Otherwise, append after the last row
  const newRowStart = clearExisting ? 2 : (sh.getLastRow() + 1);
  sh.getRange(newRowStart, 1, allRows.length, allHeaders.length).setValues(allRows);
  Logger.log(`Wrote ${allRows.length} rows to Sessions sheet starting at row ${newRowStart} with ${allHeaders.length} columns (no deduplication)`);
  
  // Final cleanup: Ensure all rows are visible and no groups/filters are hiding rows
  try {
    const finalLastRow = Math.max(sh.getLastRow(), 1);
    if (finalLastRow > 1) {
      // Remove any row groups that might have been created
      sh.getRange(1, 1, finalLastRow, 1).shiftRowGroupDepth(-8);
      
      // Unhide all rows (critical - this ensures no rows are hidden)
      sh.showRows(1, finalLastRow);
      
      // Remove and recreate filter without any criteria (filters can hide rows)
      try {
        const filter = sh.getFilter();
        if (filter) {
          const filterRange = filter.getRange();
          filter.remove();
          // Recreate filter without any criteria (allows sorting but doesn't hide rows)
          filterRange.createFilter();
        }
      } catch (e) {
        // Filter might not exist - that's okay
      }
      
      Logger.log(`Final cleanup: Ensured all ${finalLastRow} rows are visible in Sessions sheet`);
    }
  } catch (e) {
    Logger.log(`Warning: Final cleanup failed: ${e.toString()}`);
  }
  
  return allRows.length;
}
/* ===== Ingestion Functions ===== */
function pullDateRangeYesterday() {
  try {
    const range = getTimeFrameRange_('Yesterday');
    const cfg = getCfg_();
    
    // Ensure we're using local timezone dates (not UTC)
    // startDate should be yesterday 00:00:00 local, endDate should be yesterday 23:59:59.999 local
    // Fix month/day bug: endISO must use endDate month/day
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
    // Fix month/day bug: endISO must use endDate month/day
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
    // Standardize on 'Last Week'
    const range = getTimeFrameRange_('Last Week');
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

function getTimeFrameRange_(label) {
  const makeDate = (y, m, d, h = 0, min = 0, s = 0, ms = 0) => new Date(y, m, d, h, min, s, ms);
  const now = new Date();
  const todayStart = makeDate(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
  const todayEnd = makeDate(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);

  const startOfWeek = (date) => {
    const d = makeDate(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
    const day = d.getDay(); // Sunday = 0
    const diff = day === 0 ? -6 : (1 - day); // Monday as start
    d.setDate(d.getDate() + diff);
    return d;
  };

  const normalizeLabel = String(label || '').trim().toLowerCase();

  const buildRange = (start, end) => ({
    startDate: new Date(start.getTime()),
    endDate: new Date(end.getTime())
  });

  switch (normalizeLabel) {
    case 'today':
      return buildRange(todayStart, todayEnd);
    case 'yesterday': {
      const start = new Date(todayStart.getTime());
      start.setDate(start.getDate() - 1);
      const end = new Date(start.getTime());
      end.setHours(23, 59, 59, 999);
      return buildRange(start, end);
    }
    case 'last week': {
      const thisWeekStart = startOfWeek(now);
      const start = new Date(thisWeekStart.getTime());
      start.setDate(start.getDate() - 7);
      const end = new Date(thisWeekStart.getTime() - 1);
      return buildRange(start, end);
    }
    case 'this week': {
      const weekStart = startOfWeek(now);
      return buildRange(weekStart, todayEnd);
    }
    case 'last month': {
      const start = makeDate(now.getFullYear(), now.getMonth() - 1, 1, 0, 0, 0, 0);
      const end = makeDate(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999);
      return buildRange(start, end);
    }
    case 'this month': {
      const start = makeDate(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
      return buildRange(start, todayEnd);
    }
    default:
      return buildRange(todayStart, todayEnd);
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
    // Use configured nodes; if none valid, fall back to default candidate
    let nodes = cfg.nodes.map(n => Number(n)).filter(n => Number.isFinite(n) && n > 0);
    if (!nodes.length) {
      Logger.log('No valid nodes configured; falling back to default nodes');
      nodes = NODE_CANDIDATES_DEFAULT.slice();
    }
    const noderefs = ['CHANNEL']; // Only use CHANNEL to avoid duplicate calls
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
    // Use TEXT output format (XML not working reliably)
    try {
      const rt = apiGet_(cfg.rescueBase, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
      const tt = (rt.getContentText() || '').trim();
      if (!/^OK/i.test(tt)) Logger.log(`setOutput TEXT warning: ${tt}`);
    } catch (e) {
      Logger.log('setOutput TEXT failed (non-fatal): ' + e.toString());
    }
    setDelimiter_(cfg.rescueBase, cookie, '|');
  // Extract date strings (YYYY-MM-DD) for API date setting (accept both date-only and datetime strings)
  const startDateIso = String(startTimestamp).split('T')[0];
  const endDateIso = String(endTimestamp).split('T')[0];
    
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
    const startTime = new Date().getTime();
    const maxProcessingTime = 240000; // 4 minutes max processing time
    for (const nr of noderefs) {
      for (const node of nodes) {
        // Check if we're approaching timeout
        const elapsed = new Date().getTime() - startTime;
        if (elapsed > maxProcessingTime) {
          Logger.log(`Processing timeout protection: ${elapsed}ms elapsed, stopping to prevent script timeout`);
          break;
        }
        try {
          dlog(`Processing node ${node} (${nr})...`);
          // Add delay before each request to prevent rate limiting
          Utilities.sleep(500);
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          if (!t) continue;
          const parseStart = new Date().getTime();
          const parseResult = parsePipe_(t, '|');
          dlog(`Parsed ${parseResult.rows.length} rows in ${new Date().getTime() - parseStart}ms`);
          const parsed = parseResult.rows || [];
          if (!parsed || !parsed.length) continue;
          
          // Debug: log first row keys to see what headers we're getting
          if (DEBUG && parsed.length > 0) {
            const firstRow = parsed[0];
            const firstRowKeys = Object.keys(firstRow).slice(0, 10);
            dlog(`Sample row keys: ${firstRowKeys.join(', ')}`);
          }
          
          // Store raw parsed data with headers for direct dump
          allMappedRows.push({
            headers: parseResult.headers || [],
            rows: parsed,
            node: node,
            noderef: nr
          });
          // Increased delay to prevent POLLRATEEXCEEDED - wait 1 second between requests
          Utilities.sleep(1000);
        } catch (e) {
          Logger.log(`Error processing node ${node} (${nr}): ${e.toString()}`);
        }
      }
      // Break outer loop if timeout approaching
      if (new Date().getTime() - startTime > maxProcessingTime) break;
    }
    if (allMappedRows.length > 0) {
      const written = writeRowsToSheets_(ss, allMappedRows, clearExisting);
      Logger.log(`Ingested ${written} new rows to Sheets`);
      
  // Get the date range used for this pull (construct with local midnight to avoid UTC shifts)
  const pullStartDate = new Date(startDateIso + 'T00:00:00');
  const pullEndDate = new Date(endDateIso + 'T23:59:59');
      
      // Clear Digium_Raw at the start of each pull to prevent old data from persisting
      try {
        const rawSh = ss.getSheetByName('Digium_Raw');
        if (rawSh) {
          rawSh.clear();
          Logger.log('Cleared Digium_Raw sheet at start of pull');
        }
      } catch (e) {
        Logger.log('Failed to clear Digium_Raw at start of pull: ' + e.toString());
      }
      
      // Auto-refresh dashboard and create summaries with the pulled range
        const extensionMeta = getActiveExtensionMetadata_();
      const digiumDataset = getDigiumDataset_(pullStartDate, pullEndDate, extensionMeta);
      try {
        const dailySheet = ss.getSheetByName('Daily_Summary');
        if (dailySheet) {
          ss.deleteSheet(dailySheet);
          Logger.log('Removed Daily_Summary sheet (no longer needed).');
        }
      } catch (e) {
        Logger.log('Failed to remove Daily_Summary sheet: ' + e.toString());
      }
      createSupportDataSheet_(ss, pullStartDate, pullEndDate, digiumDataset, extensionMeta);
      refreshAdvancedAnalyticsDashboard_(pullStartDate, pullEndDate, digiumDataset, extensionMeta);
      
      // Generate/refresh personal dashboards (reuse Digium dataset)
      generateTechnicianTabs_(pullStartDate, pullEndDate, digiumDataset, extensionMeta);
      
      // Update Digium_Calls sheet from cached dataset
      createDigiumCallsSheet_(digiumDataset, extensionMeta);
      
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
    // Set TEXT output (XML not working reliably)
    try {
      const rt = apiGet_(cfg.rescueBase, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
      const tt = (rt.getContentText() || '').trim();
      if (!/^OK/i.test(tt)) Logger.log(`setOutput TEXT warning: ${tt}`);
    } catch (e) { Logger.log('setOutput TEXT failed (non-fatal): ' + e.toString()); }
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
  const allRawRows = []; // Store ALL raw rows for dumping
  const allHeaders = []; // Store headers from API response
  try {
    let cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    
    // Set timezone to EDT/EST with DST detection
    setTimezoneEDT_(cfg.rescueBase, cookie);
    
    // Set up for performance/summary report
    setReportAreaPerformance_(cfg.rescueBase, cookie);
    setReportTypeSummary_(cfg.rescueBase, cookie);
    // Set TEXT output (XML not working reliably)
    try {
      const rt = apiGet_(cfg.rescueBase, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
      const tt = (rt.getContentText() || '').trim();
      if (!/^OK/i.test(tt)) Logger.log(`setOutput TEXT warning: ${tt}`);
    } catch (e) { Logger.log('setOutput TEXT failed (non-fatal): ' + e.toString()); }
    setDelimiter_(cfg.rescueBase, cookie, '|');
    
    // Set date range
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    Logger.log(`fetchPerformanceSummaryData_: Setting date range ${startStr} to ${endStr}`);
    setReportDate_(cfg.rescueBase, cookie, startStr, endStr);
    setReportTimeAllDay_(cfg.rescueBase, cookie);
    
    // Re-assert report area and type after setting dates (some environments reset after date changes)
    setReportAreaPerformance_(cfg.rescueBase, cookie);
    setReportTypeSummary_(cfg.rescueBase, cookie);
    
    // Verify report type
    try {
      const rt = getReportType_(cfg.rescueBase, cookie);
      if (rt !== 'SUMMARY') {
        Logger.log(`fetchPerformanceSummaryData_: Report type after date set was ${rt}, re-applying SUMMARY`);
        setReportTypeSummary_(cfg.rescueBase, cookie);
      }
    } catch (e) {
      Logger.log('fetchPerformanceSummaryData_: getReportType check failed (non-fatal): ' + e.toString());
    }
    
  // Only use node 5648341 - filter out failed nodes
  let nodes = [];
  if (Array.isArray(cfg.nodes) && cfg.nodes.length) {
    nodes = cfg.nodes
      .map(n => Number(n))
      .filter(n => Number.isFinite(n) && n >= 0);
  }
  if (!nodes.length) {
    nodes = NODE_CANDIDATES_DEFAULT.filter(n => Number.isFinite(n) && n >= 0);
  }
  if (!nodes.length) {
    Logger.log('fetchPerformanceSummaryData_: No valid nodes configured; defaulting to 5648341');
    nodes = [5648341];
  }
  const noderefSet = new Set(
    getSummaryNoderefs_()
      .map(n => String(n || '').toUpperCase())
      .filter(Boolean)
  );
  const noderefs = Array.from(noderefSet.size ? noderefSet : ['CHANNEL']);
  Logger.log(`fetchPerformanceSummaryData_: Using nodes ${nodes.join(', ')}, noderefs: ${noderefs.join(', ')}`);
    
    for (const nr of noderefs) {
      for (const node of nodes) {
        try {
          const t = getReportTry_(cfg.rescueBase, cookie, node, nr);
          // Accept XML or TEXT; getReportTry_ already validated acceptable shapes
          if (!t) continue;
          
          const parseResult = parsePipe_(t, '|');
          const parsed = parseResult.rows || [];
          if (!parsed || !parsed.length) {
            Logger.log(`fetchPerformanceSummaryData_: No rows returned from node ${node} (${nr}). Raw preview: ${String(t).substring(0, 200)}`);
            if (parseResult.headers && parseResult.headers.length && allHeaders.length === 0) {
              allHeaders.push(...parseResult.headers);
              Logger.log(`Summary data headers (empty rows): ${parseResult.headers.join('|')}`);
            }
            // Store raw text so we can inspect it in the sheet later
            allRawRows.push({ '__RAW_TEXT__': String(t).substring(0, 32760), '__NODE__': String(node), '__NODEREF__': String(nr) });
            continue;
          }
          
          Logger.log(`fetchPerformanceSummaryData_: Parsed ${parsed.length} rows from node ${node} (${nr})`);
          
          // Store headers from first successful response
          if (parseResult.headers && parseResult.headers.length > 0 && allHeaders.length === 0) {
            allHeaders.push(...parseResult.headers);
            Logger.log(`Summary data headers: ${parseResult.headers.join('|')}`);
          } else if (!parseResult.headers || parseResult.headers.length === 0) {
            Logger.log(`fetchPerformanceSummaryData_: WARNING - No headers found in response from node ${node} (${nr})`);
          }
          
          // Store ALL raw rows for dumping (don't filter by technician name)
          allRawRows.push(...parsed);
          
          // Log first row structure for debugging
          if (parsed.length > 0) {
            const firstRow = parsed[0];
            const firstRowKeys = Object.keys(firstRow).slice(0, 10);
            Logger.log(`fetchPerformanceSummaryData_: First row keys: ${firstRowKeys.join(', ')}`);
          }
          
          parsed.forEach(row => {
            // Try multiple variations of technician name column
            const techName = row['Technician Name'] || row['Technician'] || row['TechnicianName'] || 
                           row['Technician Name:'] || row['Tech Name'] || '';
            
            // Only parse if we have a technician name (for dashboard use)
            // But we've already stored the raw row above, so it will be dumped regardless
            if (!techName) {
              // Log first row to see structure if no tech name found
              if (parsed.indexOf(row) === 0) {
                Logger.log(`First summary row (no tech name found): ${JSON.stringify(row).substring(0, 200)}`);
              }
              // Skip parsing for dashboard use, but row is already in allRawRows for dumping
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
                    const totalSessionsCol = row['Total Sessions'] || row['Sessions'] || row['Total Sessions:'] || 
                                           row['Count'] || row['Session Count'] || row['Total'] || '';
                    const totalLoginTimeCol = row['Total Login Time'] || row['Login Time'] || 
                                            row['Total Login Time:'] || row['Login Time:'] ||
                                            row['Logged In Time'] || row['Logged In Time:'] ||
                                            row['Total Logged Time'] || row['Total Logged Time:'] || '';
            const totalActiveTimeCol = row['Total Active Time'] || row['Active Time'] ||
                                       row['Total Active Time:'] || row['Active Time:'] ||
                                       row['ActiveTime'] || row['ActiveTime:'] || '';
            const totalWorkTimeCol = row['Total Work Time'] || row['Work Time'] ||
                                     row['Total Work Time:'] || row['Work Time:'] ||
                                     row['Total Worktime'] || row['Total Worktime:'] || '';
            const sessionsPerHourCol = row['Number of Sessions per Hour'] || row['Sessions per Hour'] ||
                                       row['Sessions Per Hour'] || row['SessionsPerHour'] || '';
            
            if (!performanceData[techName]) {
              performanceData[techName] = {
                avgDuration: 0,
                avgPickup: 0,
                        totalSessions: 0,
                        totalLoginTime: 0,
                totalActiveTime: 0,
                totalWorkTime: 0,
                sessionsPerHour: 0,
                        count: 0,
                rawRow: row
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
            
            // Parse total sessions (should be a number)
            let totalSessions = 0;
            if (totalSessionsCol) {
              const sessionsStr = String(totalSessionsCol).trim();
              if (/^\d+$/.test(sessionsStr)) {
                totalSessions = Number(sessionsStr);
              }
            }
            
            // Parse total login time (could be in seconds or MM:SS or HH:MM:SS format)
            let totalLoginTime = 0;
            if (totalLoginTimeCol) {
              const loginStr = String(totalLoginTimeCol).trim();
              if (/^\d+$/.test(loginStr)) {
                totalLoginTime = Number(loginStr);
              } else {
                const parts = loginStr.split(':').map(p => Number(p));
                if (parts.length === 2 && parts.every(p => !isNaN(p))) {
                  // Treat 2-part totals as HH:MM (Rescue summary omits seconds)
                  totalLoginTime = parts[0] * 3600 + parts[1] * 60;
                } else if (parts.length === 3 && parts.every(p => !isNaN(p))) {
                  totalLoginTime = parts[0] * 3600 + parts[1] * 60 + parts[2];
                }
              }
            }
            
            // Parse total active time
            let totalActiveTime = 0;
            if (totalActiveTimeCol) {
              const activeStr = String(totalActiveTimeCol).trim();
              if (/^\d+$/.test(activeStr)) {
                totalActiveTime = Number(activeStr);
              } else {
                const parts = activeStr.split(':').map(p => Number(p));
                if (parts.length === 2 && parts.every(p => !isNaN(p))) {
                  totalActiveTime = parts[0] * 3600 + parts[1] * 60;
                } else if (parts.length === 3 && parts.every(p => !isNaN(p))) {
                  totalActiveTime = parts[0] * 3600 + parts[1] * 60 + parts[2];
                }
              }
            }
            
            // Parse total work time
            let totalWorkTime = 0;
            if (totalWorkTimeCol) {
              const workStr = String(totalWorkTimeCol).trim();
              if (/^\d+$/.test(workStr)) {
                totalWorkTime = Number(workStr);
              } else {
                const parts = workStr.split(':').map(p => Number(p));
                if (parts.length === 2 && parts.every(p => !isNaN(p))) {
                  totalWorkTime = parts[0] * 3600 + parts[1] * 60;
                } else if (parts.length === 3 && parts.every(p => !isNaN(p))) {
                  totalWorkTime = parts[0] * 3600 + parts[1] * 60 + parts[2];
                }
              }
            }
            // Parse sessions per hour
            let sessionsPerHour = 0;
            if (sessionsPerHourCol != null && sessionsPerHourCol !== '') {
              const sphStr = String(sessionsPerHourCol).trim();
              const num = Number(sphStr);
              if (!isNaN(num)) sessionsPerHour = num;
            }
            
            // Accumulate (in case multiple rows per technician)
            performanceData[techName].avgDuration += avgDur;
            performanceData[techName].avgPickup += avgPickup;
            performanceData[techName].totalSessions += totalSessions;
            performanceData[techName].totalLoginTime += totalLoginTime;
            performanceData[techName].totalActiveTime += totalActiveTime;
            performanceData[techName].totalWorkTime += totalWorkTime;
            if (sessionsPerHour) performanceData[techName].sessionsPerHour += sessionsPerHour;
            performanceData[techName].count++;
            // Store raw row data for dumping (keep the most recent one)
            performanceData[techName].rawRow = row;
          });
          
          Utilities.sleep(200);
        } catch (e) {
          Logger.log(`Error fetching performance data from node ${node} (${nr}): ${e.toString()}`);
        }
      }
    }
    // Calculate averages if we have multiple entries per technician
    // Note: totalSessions and totalLoginTime should be summed, not averaged
    Object.keys(performanceData).forEach(tech => {
      const perf = performanceData[tech];
      const divisor = perf.count > 0 ? perf.count : 1;
      perf.avgDuration = perf.avgDuration / divisor;
      perf.avgPickup = perf.avgPickup / divisor;
      perf.sessionsPerHour = perf.sessionsPerHour / divisor;
    });
    
    // Dump ALL raw rows to sheet (not just parsed ones)
    try {
      dumpPerformanceSummaryRaw_(startDate, endDate, allRawRows, allHeaders);
    } catch (e) {
      Logger.log('Warning: Failed to dump performance summary raw data: ' + e.toString());
    }
    
    Logger.log(`fetchPerformanceSummaryData_: Found performance data for ${Object.keys(performanceData).length} technicians, ${allRawRows.length} total raw rows`);
  } catch (e) {
    Logger.log('fetchPerformanceSummaryData_ error: ' + e.toString());
  }
  
  return performanceData;
}
/**
 * Helper to quickly preview the raw Rescue performance summary payload for a given date range / node.
 * Usage (Apps Script console):
 *    previewRescueSummary_('2025-11-07', '2025-11-07', 5648341, 'CHANNEL');
 * Dates default to today if omitted.
 */
function previewRescueSummary_(startISOOpt, endISOOpt, nodeOpt, noderefOpt) {
  try {
    const cfg = getCfg_();
    if (!cfg || !cfg.rescueBase) {
      Logger.log('previewRescueSummary_: Rescue credentials not configured');
      return '';
    }
    const startISO = startISOOpt || isoDate_(new Date());
    const endISO = endISOOpt || startISO;
    const startDate = new Date(startISO + 'T00:00:00');
    const endDate = new Date(endISO + 'T23:59:59');
    Logger.log(`previewRescueSummary_: Using date range ${startISO} to ${endISO}`);

    const cookie = login_(cfg.rescueBase, cfg.user, cfg.pass);
    setTimezoneEDT_(cfg.rescueBase, cookie);
    setReportAreaPerformance_(cfg.rescueBase, cookie);
    setReportTypeSummary_(cfg.rescueBase, cookie);
    try { apiGet_(cfg.rescueBase, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true); } catch (e) { Logger.log('setOutput TEXT failed (non-fatal): ' + e.toString()); }
    setDelimiter_(cfg.rescueBase, cookie, '|');
    setReportDate_(cfg.rescueBase, cookie, startISO, endISO);
    setReportTimeAllDay_(cfg.rescueBase, cookie);

    // Choose node / noderef
    let node = Number(nodeOpt);
    if (!Number.isFinite(node)) {
      const configured = Array.isArray(cfg.nodes) && cfg.nodes.length ? cfg.nodes.map(n => Number(n)).filter(Number.isFinite) : [];
      node = configured.length ? configured[0] : (NODE_CANDIDATES_DEFAULT[0] || 5648341);
    }
    const noderefCandidates = getSummaryNoderefs_();
    const noderef = noderefOpt ? String(noderefOpt).toUpperCase() : (noderefCandidates[0] ? String(noderefCandidates[0]).toUpperCase() : 'CHANNEL');

    Logger.log(`previewRescueSummary_: Requesting node ${node} with noderef ${noderef}`);
    const raw = getReportTry_(cfg.rescueBase, cookie, node, noderef);
    if (!raw) {
      Logger.log('previewRescueSummary_: No response received (null/empty)');
      return '';
    }

    const preview = raw.length > 1000 ? raw.substring(0, 1000) + '…' : raw;
    Logger.log(`previewRescueSummary_: Raw response preview (${raw.length} chars):\n${preview}`);

    try {
      const parsed = parsePipe_(raw, '|');
      Logger.log(`previewRescueSummary_: Parsed ${parsed.rows ? parsed.rows.length : 0} rows. Headers: ${(parsed.headers || []).join('|')}`);
    } catch (e) {
      Logger.log('previewRescueSummary_: parsePipe_ failed: ' + e.toString());
    }
    return raw;
  } catch (err) {
    Logger.log('previewRescueSummary_ error: ' + err.toString());
    return '';
  }
}

// Register as a custom menu option so it shows up in the Run dropdown / App Script UI.
function previewRescueSummary() {
  previewRescueSummary_();
}
// Dump raw performance summary data to a sheet for inspection
function dumpPerformanceSummaryRaw_(startDate, endDate, allRawRows, allHeaders) {
  try {
    const ss = SpreadsheetApp.getActive();
    let out = ss.getSheetByName('Performance_Summary_Raw');
    if (!out) {
      out = ss.insertSheet('Performance_Summary_Raw');
      out.clear();
    } else {
      // Only clear if this is a new date range (check existing date range in sheet)
      try {
        const existingDateRange = out.getRange(2, 1).getValue();
        const newDateRange = `Date Range: ${startDate.toISOString().split('T')[0]} to ${endDate.toISOString().split('T')[0]}`;
        if (existingDateRange && String(existingDateRange).trim() !== String(newDateRange).trim()) {
          // Date range changed - clear the sheet
          Logger.log(`Performance_Summary_Raw: Date range changed, clearing sheet`);
          out.clear();
        } else {
          // Date range is the same - don't clear, just update data rows
          Logger.log(`Performance_Summary_Raw: Same date range, updating data without clearing`);
        }
      } catch (e) {
        // If we can't read the date range, clear to be safe
        Logger.log(`Performance_Summary_Raw: Could not read existing date range, clearing: ${e.toString()}`);
        out.clear();
      }
    }
    
    // Check if we need to write headers (only if sheet was cleared or is new)
    const needsHeaders = out.getLastRow() === 0 || !out.getRange(1, 1).getValue();
    const dataStartRow = 5;
    
    if (needsHeaders) {
      // Write header with date range info
      out.getRange(1, 1).setValue('Performance Summary Raw Data');
      out.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    }
    
    // Always update date range and timestamp
    out.getRange(2, 1).setValue(`Date Range: ${startDate.toISOString().split('T')[0]} to ${endDate.toISOString().split('T')[0]}`);
    out.getRange(2, 1).setFontSize(12);
    out.getRange(3, 1).setValue(`Last Updated: ${new Date().toLocaleString()}`);
    out.getRange(3, 1).setFontSize(10).setFontColor('#666666');
    
    // Dump ALL raw rows with ALL columns
    if (allRawRows && allRawRows.length > 0) {
      // Get all unique column names from all rows
      const columnSet = new Set();
      allRawRows.forEach(row => {
        Object.keys(row).forEach(key => columnSet.add(key));
      });
      const columnNames = Array.from(columnSet).sort();
      
      Logger.log(`Performance_Summary_Raw: Found ${allRawRows.length} raw rows with ${columnNames.length} unique columns: ${columnNames.join(', ')}`);
      
      // Write headers
      if (needsHeaders && columnNames.length > 0) {
        out.getRange(dataStartRow, 1, 1, columnNames.length).setValues([columnNames]);
        out.getRange(dataStartRow, 1, 1, columnNames.length).setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF');
      }
      
      // Clear existing data rows and write new ones
      const lastDataRow = out.getLastRow();
      if (lastDataRow >= dataStartRow + 1) {
        out.getRange(dataStartRow + 1, 1, lastDataRow - dataStartRow, out.getLastColumn()).clearContent();
      }
      
      // Write all rows with all columns
      const rows = allRawRows.map(row => {
        return columnNames.map(col => row[col] || '');
      });
      
      if (rows.length > 0) {
        out.getRange(dataStartRow + 1, 1, rows.length, columnNames.length).setValues(rows);
        Logger.log(`Performance_Summary_Raw: Dumped ${rows.length} raw rows with ${columnNames.length} columns`);
      }
      
      // Auto-resize columns
      try { out.autoResizeColumns(1, columnNames.length); } catch (e) {}
    } else {
      Logger.log(`Performance_Summary_Raw: No raw rows to dump (allRawRows is empty or null)`);
    }
    
  } catch (e) {
    Logger.log('dumpPerformanceSummaryRaw_ error: ' + e.toString());
  }
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
    // Set TEXT output (XML not working reliably)
    try {
      const rt = apiGet_(base, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
      const tt = (rt.getContentText() || '').trim();
      if (!/^OK/i.test(tt)) Logger.log(`setOutput TEXT warning: ${tt}`);
    } catch (e) { Logger.log('setOutput TEXT failed (non-fatal): ' + e.toString()); }
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
  // Fix: sheet name should match main Sessions sheet constant
  const sh = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
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
    // Set TEXT output (XML not working reliably)
    try {
      const rt = apiGet_(cfg.rescueBase, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
      const tt = (rt.getContentText() || '').trim();
      if (!/^OK/i.test(tt)) Logger.log(`setOutput TEXT warning: ${tt}`);
    } catch (e) { Logger.log('setOutput TEXT failed (non-fatal): ' + e.toString()); }
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
function createDailySummarySheet_(ss, startDate, endDate, digiumDatasetOpt, extensionMetaOpt) {
  try {
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    
    let summarySheet = ss.getSheetByName('Daily_Summary');
    if (!summarySheet) summarySheet = ss.insertSheet('Daily_Summary');
    
    summarySheet.clear();
    
    const extensionMeta = extensionMetaOpt || getActiveExtensionMetadata_();
    const digiumDataset = digiumDatasetOpt || getDigiumDataset_(startDate, endDate, extensionMeta);
    const callMetricsByCanonicalGlobal = (digiumDataset && digiumDataset.callMetricsByCanonical) ? digiumDataset.callMetricsByCanonical : {};
    const callMetricsByCanonical = {};
    if (callMetricsByCanonicalGlobal && typeof callMetricsByCanonicalGlobal === 'object') {
      Object.keys(callMetricsByCanonicalGlobal).forEach(key => {
        const normalizedKey = canonicalTechnicianName_(key);
        const targetKey = normalizedKey || key;
        callMetricsByCanonical[targetKey] = callMetricsByCanonicalGlobal[key];
      });
    }
    
    // Get all data from Sessions
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    
    const allDataRaw = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const allData = filterOutExcludedTechnicians_(headers, allDataRaw);
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
    const startMillis = startDate.getTime();
    const endMillis = endDate.getTime();
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx]) return false;
      try {
        const rowObj = row[startIdx] instanceof Date ? row[startIdx] : new Date(row[startIdx]);
        if (!(rowObj instanceof Date) || isNaN(rowObj)) return false;
        const rowMillis = rowObj.getTime();
        if (rowMillis < startMillis || rowMillis > endMillis) return false;
        if (techIdx >= 0 && isExcludedTechnician_(row[techIdx])) return false;
        return true;
      } catch (e) {
        return false;
      }
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
    const formatDateKey = (d) => {
      try {
        return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      } catch (e) {
        return isoDate_(d);
      }
    };
    const formatDisplayDate = (key) => {
      try {
        const d = new Date(key + 'T00:00:00');
        return Utilities.formatDate(d, tz, 'M/d/yyyy');
      } catch (e) {
        return key;
      }
    };

    const dates = [];
    let currentDate = new Date(startDate);
    const endDateObj = new Date(endDate);
    while (currentDate <= endDateObj) {
      dates.push(formatDateKey(currentDate));
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
    
    // Summary section (Technician totals) - uses 7 fixed columns
    const summaryHeaderCols = 7; // Fixed 7 columns for technician summary
    summaryRows.push(['Summary', ...Array(summaryHeaderCols - 1).fill('')]);
    // Summary headers: 7 fixed columns
    const summaryHeaders = ['Technician Name', 'Total Sessions', '% Of Total sessions', 'Sessions per HR', 'Avg Pick-up Speed', 'Avg Duration', 'Average Work Time'];
    summaryRows.push(summaryHeaders);
    
    const perTechDaily = {};
    const dailyDisplayNames = {};
    const ensureDailyStats = (canonical, dateKey) => {
      if (!canonical) return null;
      if (!perTechDaily[canonical]) perTechDaily[canonical] = {};
      const techDays = perTechDaily[canonical];
      if (!techDays[dateKey]) {
        techDays[dateKey] = {
          sessions: 0,
          durationSum: 0,
          durationCount: 0,
          pickupSum: 0,
          pickupCount: 0,
          workSeconds: 0,
          activeSeconds: 0,
          longestSeconds: 0,
          loginSeconds: 0
        };
      }
      return techDays[dateKey];
    };
    
    // Calculate per-technician stats
    const techStats = {};
    const sfMetrics = collectSalesforceTicketMetrics_(startDate, endDate);
    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      if (!techStats[tech]) {
        techStats[tech] = {
          sessions: 0,
          durations: [],
          pickups: [],
          workSeconds: 0,
          talkSeconds: 0,
          totalCalls: 0,
          inboundCalls: 0,
          outboundCalls: 0,
          novaWave: 0,
          tickets: 0,
          daysToCloseAvg: 0
        };
      }
      techStats[tech].sessions++;
  if (row[durationIdx]) techStats[tech].durations.push(parseDurationSeconds_(row[durationIdx]));
  if (row[pickupIdx]) techStats[tech].pickups.push(parseDurationSeconds_(row[pickupIdx]));
  if (row[workIdx]) techStats[tech].workSeconds += parseDurationSeconds_(row[workIdx]);
    });
    
    const ticketsByCanonical = {};
    const daysToCloseByCanonical = {};
    if (sfMetrics && sfMetrics.perCanonical) {
      Object.keys(sfMetrics.perCanonical).forEach(canonical => {
        const entry = sfMetrics.perCanonical[canonical];
        if (!entry) return;
        ticketsByCanonical[canonical] = (ticketsByCanonical[canonical] || 0) + (entry.created || 0);
        if (entry.daysToClose && entry.daysToClose.length) {
          daysToCloseByCanonical[canonical] = (daysToCloseByCanonical[canonical] || []).concat(entry.daysToClose);
        }
        (entry.rawNames || []).forEach(rawName => {
          const normalized = canonicalTechnicianName_(rawName);
          if (!normalized || normalized === canonical) return;
          ticketsByCanonical[normalized] = (ticketsByCanonical[normalized] || 0) + (entry.created || 0);
          if (entry.daysToClose && entry.daysToClose.length) {
            daysToCloseByCanonical[normalized] = (daysToCloseByCanonical[normalized] || []).concat(entry.daysToClose);
          }
        });
      });
      Object.keys(techStats).forEach(tech => {
        const canonical = canonicalTechnicianName_(tech);
        const ticketCount = ticketsByCanonical[canonical] || ticketsByCanonical[tech] || 0;
        if (ticketCount) techStats[tech].tickets += ticketCount;
        const daysArr = daysToCloseByCanonical[canonical] || daysToCloseByCanonical[tech] || [];
        if (daysArr.length) {
          techStats[tech].daysToCloseAvg = daysArr.reduce((a,b)=>a+b,0) / daysArr.length;
        } else {
          techStats[tech].daysToCloseAvg = 0;
        }
      });
    }
    Object.keys(techStats).forEach(tech => {
      const stats = techStats[tech];
      const canonicalKey = canonicalTechnicianName_(tech);
      const normalizedFull = normalizeTechnicianNameFull_(tech);
      const metrics =
        callMetricsByCanonical[canonicalKey] ||
        callMetricsByCanonical[normalizedFull] ||
        null;
      if (metrics) {
        stats.totalCalls = metrics.totalCalls || 0;
        stats.inboundCalls = metrics.inboundCalls || 0;
        stats.outboundCalls = metrics.outboundCalls || 0;
        stats.talkSeconds = metrics.talkSeconds || 0;
      }
    });
    const techRows = Object.keys(techStats).sort((a, b) => techStats[b].sessions - techStats[a].sessions).map(tech => {
      const stats = techStats[tech];
      const avgPickup = stats.pickups.length > 0 ? (stats.pickups.reduce((a, b) => a + b, 0) / stats.pickups.length) : 0;
      const avgDur = stats.durations.length > 0 ? (stats.durations.reduce((a, b) => a + b, 0) / stats.durations.length) : 0;
      const workTimeFraction = stats.workSeconds / 86400;
      const talkTime = stats.talkSeconds / 86400;
      return [
        tech,
        stats.sessions,
        stats.novaWave,
        avgPickup / 86400,
        avgDur / 86400,
        workTimeFraction,
        stats.totalCalls,
        stats.inboundCalls,
        stats.outboundCalls,
        talkTime,
        stats.tickets,
        stats.daysToCloseAvg || 0
      ];
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
      // Tech row must match header structure with exactly headerCols columns
      // Tech row must have exactly 7 columns to match summaryHeaders
      const techRow = [tech, stats.sessions, pct, sessionsPerHr, avgPickup / 86400, avgDur / 86400, workTime / 86400];
      // Ensure exactly 7 columns
      while (techRow.length < summaryHeaderCols) {
        techRow.push('');
      }
      techRow.length = summaryHeaderCols;
      summaryRows.push(techRow);
    });
    
    // Write to sheet - ensure all rows have the same number of columns
    if (summaryRows.length > 0) {
      // Find maximum column count across all rows
      const maxCols = Math.max(...summaryRows.map(row => row.length));
      // Pad all rows to maxCols to ensure consistent column count
      summaryRows.forEach(row => {
        while (row.length < maxCols) {
          row.push('');
        }
        row.length = maxCols; // Ensure exact length
      });
      summarySheet.getRange(1, 1, summaryRows.length, maxCols).setValues(summaryRows);
      summarySheet.getRange(1, 1, 1, maxCols).setFontWeight('bold').setBackground('#E5E7EB');
      summarySheet.getRange(2, 1, 1, maxCols).setFontWeight('bold');
      summarySheet.getRange(summaryRows.length - Object.keys(techStats).length, 1, 1, maxCols).setFontWeight('bold').setBackground('#E5E7EB');
      summarySheet.setFrozenRows(1);
      summarySheet.setColumnWidth(1, 300);
      for (let i = 2; i <= maxCols; i++) {
        summarySheet.setColumnWidth(i, 120);
      }
      try { applyProfessionalTableStyling_(summarySheet, maxCols); } catch (e) { Logger.log('Styling daily summary failed: ' + e.toString()); }

      // Post-formatting: ensure duration/pickup cells are numeric time fractions and formatted as hh:mm:ss
      try {
        const lastCol = maxCols;
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
function createSupportDataSheet_(ss, startDate, endDate, digiumDatasetOpt, extensionMetaOpt) {
  try {
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    
    let supportSheet = ss.getSheetByName('Support_Data');
    if (!supportSheet) supportSheet = ss.insertSheet('Support_Data');
    
    supportSheet.clear();
    try { supportSheet.getBandings().forEach(b => b.remove()); } catch (e) { Logger.log('Support_Data: failed to remove existing banding: ' + e.toString()); }
    try { supportSheet.setConditionalFormatRules([]); } catch (e) { Logger.log('Support_Data: failed to clear conditional formats: ' + e.toString()); }
    
    const extensionMeta = extensionMetaOpt || getActiveExtensionMetadata_();
    const digiumDataset = digiumDatasetOpt || getDigiumDataset_(startDate, endDate, extensionMeta);
    const digByAccount = digiumDataset && digiumDataset.byAccount && digiumDataset.byAccount.ok ? digiumDataset.byAccount : null;
    const digByDay = digiumDataset && digiumDataset.byDay && digiumDataset.byDay.ok ? digiumDataset.byDay : null;
    const callMetricsByCanonicalGlobal = (digiumDataset && digiumDataset.callMetricsByCanonical) ? digiumDataset.callMetricsByCanonical : {};
    const callMetricsByCanonical = {};
    if (callMetricsByCanonicalGlobal && typeof callMetricsByCanonicalGlobal === 'object') {
      Object.keys(callMetricsByCanonicalGlobal).forEach(key => {
        const normalizedKey = canonicalTechnicianName_(key);
        const targetKey = normalizedKey || key;
        callMetricsByCanonical[targetKey] = callMetricsByCanonicalGlobal[key];
      });
    }
    const perTechDaily = {};
    const dailyDisplayNames = {};

    // Get all data from Sessions
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    
    const allDataRaw = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const allData = filterOutExcludedTechnicians_(headers, allDataRaw);
    const getHeaderIndex = (variants) => {
      for (const v of variants) {
        const i = headers.findIndex(h => String(h || '').toLowerCase().trim() === String(v).toLowerCase().trim());
        if (i >= 0) return i;
      }
      return -1;
    };
    
    // Use exact header names from API as provided by user
    const startIdx = getHeaderIndex(['Start Time', 'start_time', 'start time']);
    const techIdx = getHeaderIndex(['Technician Name', 'technician_name', 'technician name']);
    const statusIdx = getHeaderIndex(['Status', 'session_status', 'status']);
    const durationIdx = getHeaderIndex(['Total Time', 'total_time', 'duration_total_seconds']);
    const workIdx = getHeaderIndex(['Work Time', 'work_time', 'duration_work_seconds']);
    const pickupIdx = getHeaderIndex(['Waiting Time', 'waiting_time', 'pickup_seconds']);
    const activeIdx = getHeaderIndex(['Active Time', 'active_time', 'duration_active_seconds', 'active seconds', 'duration_active_seconds']);
    const customerIdx = getHeaderIndex(['Your Name:', 'customer_name', 'Customer Name']);
    const sessionIdIdx = getHeaderIndex(['Session ID', 'session_id', 'session id']);
    const channelIdx = getHeaderIndex(['Channel Name', 'channel_name']);
    const resolvedIdx = getHeaderIndex(['Resolved Unresolved', 'resolved_unresolved', 'Resolved, Unresolved']);
    const callingCardIdx = getHeaderIndex(['Calling Card', 'calling_card', 'calling card']);
    
    // Log header indices for debugging
    Logger.log(`Support Data header indices - Start: ${startIdx}, Tech: ${techIdx}, Duration: ${durationIdx}, Work: ${workIdx}, Pickup: ${pickupIdx}`);
    
    // Filter by date range
    const startMillis = startDate.getTime();
    const endMillis = endDate.getTime();
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx]) return false;
      try {
        const rowObj = row[startIdx] instanceof Date ? row[startIdx] : new Date(row[startIdx]);
        if (!(rowObj instanceof Date) || isNaN(rowObj)) return false;
        const rowMillis = rowObj.getTime();
        if (rowMillis < startMillis || rowMillis > endMillis) return false;
        if (techIdx >= 0 && isExcludedTechnician_(row[techIdx])) return false;
        return true;
      } catch (e) {
        return false;
      }
    });
    
    const tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Etc/GMT';
    const indexInfo = {
      startIdx,
      techIdx,
      durationIdx,
      workIdx,
      pickupIdx,
      activeIdx,
      channelIdx,
      callingCardIdx
    };
    const sessionContext = buildTechnicianSessionContext_(filtered, indexInfo, tz);
    const perCanonicalSessions = sessionContext.perCanonical || {};
    const teamDaily = sessionContext.teamDaily || {};
    let dates = Array.isArray(sessionContext.orderedDates) ? sessionContext.orderedDates.slice() : [];
    Object.keys(perCanonicalSessions).forEach(canonical => {
      perTechDaily[canonical] = perCanonicalSessions[canonical].daily || {};
      if (!dailyDisplayNames[canonical]) {
        dailyDisplayNames[canonical] = perCanonicalSessions[canonical].displayName || canonical;
      }
    });

    if (!dates.length) {
      const tempDates = [];
      let cursor = new Date(startDate);
      const endDateObjAlt = new Date(endDate);
      while (cursor <= endDateObjAlt) {
        tempDates.push(Utilities.formatDate(cursor, tz, 'yyyy-MM-dd'));
        cursor.setDate(cursor.getDate() + 1);
      }
      dates = tempDates;
    }

    const formatDateKey = (dateObj) => {
      try {
        return Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
      } catch (e) {
        return isoDate_(dateObj);
      }
    };
    const formatDisplayDate = (key) => {
      try {
        const d = new Date(key + 'T00:00:00');
        return Utilities.formatDate(d, tz, 'M/d/yyyy');
      } catch (e) {
        return key;
      }
    };

    const callDailyPerCanonical = (digiumDataset && digiumDataset.callDailyPerCanonical)
      ? digiumDataset.callDailyPerCanonical
      : {};

    // Header
    supportSheet.getRange(1, 1).setValue('📊 SUPPORT DATA SUMMARY');
    supportSheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#0B1F50');
    supportSheet.getRange(1, 1, 1, 8).merge();

    supportSheet.getRange(2, 1).setValue('Date Range:');
    const displayStart = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
    const displayEnd = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');
    supportSheet.getRange(2, 2).setValue(`${displayStart} to ${displayEnd}`);
    supportSheet.getRange(2, 1).setFontWeight('bold');
    // Add Last Updated timestamp for uniformity with other dashboards
    supportSheet.getRange(2, 4).setValue('Last Updated:');
    supportSheet.getRange(2, 5).setValue(new Date().toLocaleString());
    supportSheet.getRange(2, 4).setFontWeight('bold');
    
    // Summary KPIs
    const kpiRow = 4;
    const totalSessions = filtered.length;
    const totalWorkSeconds = workIdx >= 0 ? filtered.reduce((sum, row) => sum + parseDurationSeconds_(row[workIdx] || 0), 0) : 0;
    const totalWorkTimeFraction = totalWorkSeconds / 86400;
    const avgPickup = (filtered.length > 0 && pickupIdx >= 0) ? 
      Math.round(filtered.reduce((sum, row) => sum + parseDurationSeconds_(row[pickupIdx] || 0), 0) / filtered.length) : 0;
    if (workIdx < 0) Logger.log('WARNING: Work Time column not found in Support Data');
    if (pickupIdx < 0) Logger.log('WARNING: Waiting Time column not found in Support Data');
    
    // Count Nova Wave sessions (calling card contains "Nova wave chat")
    const novaWaveCount = filtered.filter(row => {
      if (channelIdx >= 0 && row[channelIdx]) {
        const channelName = String(row[channelIdx] || '').toLowerCase();
        if (channelName.includes('nova wave')) return true;
      }
      if (callingCardIdx >= 0 && row[callingCardIdx]) {
        const callingCard = String(row[callingCardIdx] || '').toLowerCase();
        if (callingCard.includes('nova wave')) return true;
      }
      return false;
    }).length;
    
    // Aggregate call totals from Digium data (if available)
    let totalCalls = 0;
    let totalIncomingCalls = 0;
    let totalOutgoingCalls = 0;
    let totalTalkingSeconds = 0;
    let totalCallSeconds = 0;
    Object.values(callMetricsByCanonical).forEach(metrics => {
      if (!metrics) return;
      totalCalls += Number(metrics.totalCalls || 0);
      totalIncomingCalls += Number(metrics.inboundCalls || 0);
      totalOutgoingCalls += Number(metrics.outboundCalls || 0);
      totalTalkingSeconds += Number(metrics.talkSeconds || 0);
      totalCallSeconds += Number(metrics.callSeconds || 0);
    });
    if ((!totalCalls || !totalTalkingSeconds) && digByAccount && digByAccount.totalsAll) {
      const totalsAll = digByAccount.totalsAll || {};
      totalCalls = totalCalls || Number(totalsAll.total_calls || 0);
      totalIncomingCalls = totalIncomingCalls || Number(totalsAll.total_incoming_calls || 0);
      totalOutgoingCalls = totalOutgoingCalls || Number(totalsAll.total_outgoing_calls || 0);
      totalTalkingSeconds = totalTalkingSeconds || Number(totalsAll.talking_duration || 0);
      totalCallSeconds = totalCallSeconds || Number(totalsAll.call_duration || 0);
    }
    if (!totalTalkingSeconds && digByDay && digByDay.rows && digByDay.rows.length) {
      const talkRow = digByDay.rows.find(row => /talk/i.test(String(row[0] || '')));
      if (talkRow) {
        totalTalkingSeconds = talkRow.slice(1).reduce((sum, val) => sum + (Number(val) || 0), 0);
      }
    }

    // Aggregate Salesforce ticket data
    const sfMetrics = collectSalesforceTicketMetrics_(startDate, endDate);
    const totalTickets = sfMetrics.overall.totalCreated;
    const openTickets = sfMetrics.overall.openCurrent || 0;
    const closedTickets = sfMetrics.overall.totalClosed;
    const topIssuesOverall = Object.entries(sfMetrics.overall.topIssues || {})
      .sort((a, b) => Number(b[1]) - Number(a[1]))
      .slice(0, 5);
    
    const kpis = [
      { label: 'Total Sessions', value: totalSessions, format: 'int' },
      { label: 'Nova Wave Sessions', value: novaWaveCount, format: 'int' },
      { label: 'Total Work Hours', value: totalWorkTimeFraction, format: 'time' },
      { label: 'Avg Pickup Time', value: avgPickup / 86400, format: 'time' },
      { label: 'Total Calls', value: totalCalls, format: 'int' },
      { label: 'Total Incoming Calls', value: totalIncomingCalls, format: 'int' },
      { label: 'Total Outgoing Calls', value: totalOutgoingCalls, format: 'int' },
      { label: 'Total Talking Time', value: totalTalkingSeconds / 86400, format: 'time' },
      { label: 'Total Tickets', value: totalTickets, format: 'int' },
      { label: 'Unresolved Tickets', value: openTickets, format: 'int' },
      { label: 'Tickets Closed', value: closedTickets, format: 'int' }
    ];
    const kpisPerRow = 4;
    kpis.forEach((kpi, idx) => {
      const row = kpiRow + Math.floor(idx / kpisPerRow);
      const col = (idx % kpisPerRow) * 3 + 1;
      supportSheet.getRange(row, col).setValue(kpi.label);
      supportSheet.getRange(row, col).setFontSize(12).setFontColor('#1F2937');
      const valueCell = supportSheet.getRange(row, col + 1);
      valueCell.setValue(kpi.value);
      valueCell.setFontSize(16).setFontWeight('bold').setFontColor('#0F172A');
      switch (kpi.format) {
        case 'time':
          valueCell.setNumberFormat('hh:mm:ss');
          break;
        case 'hours':
          valueCell.setNumberFormat('0.0');
          break;
        default:
          valueCell.setNumberFormat('0');
      }
      supportSheet.getRange(row, col, 1, 2).setBorder(true, true, true, true, true, true).setBackground('#E8F1FF');
    });
    
    let currentRow = kpiRow + Math.ceil(kpis.length / kpisPerRow) + 2;
    
    const issueBlockCol = kpisPerRow * 3 + 1;
    const issueBlockWidth = 2;
    const maxIssueRows = 5;
    const issueBlockHeight = maxIssueRows + 2;
    try { supportSheet.getRange(kpiRow, issueBlockCol, issueBlockHeight, issueBlockWidth).clear(); } catch (e) {}
    supportSheet.getRange(kpiRow, issueBlockCol, 1, issueBlockWidth).merge()
      .setValue('Top Ticket Issues (Salesforce)')
      .setFontSize(12).setFontWeight('bold').setFontColor('#0B1F50')
      .setBackground('#E8F1FF').setHorizontalAlignment('center');
    supportSheet.getRange(kpiRow + 1, issueBlockCol, 1, issueBlockWidth)
      .setValues([['Issue', 'Count']])
      .setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    const issueRows = topIssuesOverall.length ? topIssuesOverall.map(([label, count]) => [label, count]) : [['No ticket issues found', ' ']];
    while (issueRows.length < maxIssueRows) issueRows.push(['—', ' ']);
    supportSheet.getRange(kpiRow + 2, issueBlockCol, issueRows.length, issueBlockWidth).setValues(issueRows);
    supportSheet.getRange(kpiRow + 2, issueBlockCol + 1, issueRows.length, 1).setNumberFormat('0');
    if (!topIssuesOverall.length) {
      supportSheet.getRange(kpiRow + 2, issueBlockCol, 1, issueBlockWidth)
        .setFontStyle('italic').setFontColor('#6B7280');
    }
    try {
      supportSheet.getRange(kpiRow + 1, issueBlockCol, issueBlockHeight - 1, issueBlockWidth)
        .setBorder(true, true, true, true, true, true);
    } catch (e) { /* ignore */ }
    currentRow = Math.max(currentRow, kpiRow + issueBlockHeight + 2);
    
    // Technician Performance Table
    supportSheet.getRange(currentRow, 1).setValue('Technician Performance');
    supportSheet.getRange(currentRow, 1).setFontSize(13).setFontWeight('bold').setFontColor('#0B1F50');
    const tableHeaders = ['Technician', 'Sessions', 'Nova Wave Sessions', 'Avg Pickup', 'Avg Duration', 'Session Work Time', 'Total Calls', 'Incoming Calls', 'Outgoing Calls', 'Total Talk Time', 'Total Tickets', 'Avg Days to Close', 'Ticket Open Rate'];
    supportSheet.getRange(currentRow + 1, 1, 1, tableHeaders.length).setValues([tableHeaders]);
    const techHeaderRow = currentRow + 1;
    supportSheet.getRange(techHeaderRow, 1, 1, tableHeaders.length)
      .setFontWeight('bold').setFontSize(12).setBackground('#1E3A8A').setFontColor('#FFFFFF');
    
    const canonicalDisplayOverrides = {
      'ahmed talal': 'Ahmed Talal'
    };
    const canonicalToDisplay = {};
    const ensureDisplayName = (canonical) => {
      if (!canonical) return 'Unknown';
      if (canonicalDisplayOverrides[canonical]) return canonicalDisplayOverrides[canonical];
      if (dailyDisplayNames[canonical]) return dailyDisplayNames[canonical];
      if (perCanonicalSessions[canonical] && perCanonicalSessions[canonical].displayName) {
        dailyDisplayNames[canonical] = perCanonicalSessions[canonical].displayName;
        return perCanonicalSessions[canonical].displayName;
      }
      const computed = canonical.split(/\s+/).map(part => part ? part.charAt(0).toUpperCase() + part.slice(1) : '').join(' ').trim() || 'Unknown';
      dailyDisplayNames[canonical] = computed;
      return computed;
    };

    const ticketAggregates = {};
    const daysToCloseAggregates = {};
    const openAggregates = {};
    if (sfMetrics && sfMetrics.perCanonical) {
      Object.keys(sfMetrics.perCanonical).forEach(key => {
        const entry = sfMetrics.perCanonical[key];
        if (!entry) return;
        const canonicalSet = new Set();
        const primary = canonicalTechnicianName_(key);
        if (primary) canonicalSet.add(primary);
        (entry.rawNames || []).forEach(rawName => {
          const c = canonicalTechnicianName_(rawName);
          if (c) canonicalSet.add(c);
        });
        canonicalSet.forEach(canonical => {
          ticketAggregates[canonical] = (ticketAggregates[canonical] || 0) + (entry.created || 0);
          if (entry.daysToClose && entry.daysToClose.length) {
            daysToCloseAggregates[canonical] = (daysToCloseAggregates[canonical] || []).concat(entry.daysToClose);
          }
          if (entry.open) {
            openAggregates[canonical] = (openAggregates[canonical] || 0) + entry.open;
          }
        });
      });
    }

    const allCanonicalKeys = Array.from(new Set([
      ...Object.keys(perCanonicalSessions),
      ...Object.keys(callMetricsByCanonical),
      ...Object.keys(ticketAggregates)
    ]));

    const techRows = allCanonicalKeys
      .map(canonical => {
        if (!canonical) return null;
        const sessionStats = perCanonicalSessions[canonical];
        const totals = sessionStats && sessionStats.totals ? sessionStats.totals : createEmptyDayStats_();
        const displayName = ensureDisplayName(canonical);
        if (!displayName || isExcludedTechnician_(displayName)) return null;
        const callTotals = callMetricsByCanonical[canonical] || {};
        const totalCallsForTech = Number(callTotals.totalCalls || 0);
        const inboundCalls = Number(callTotals.inboundCalls || 0);
        const outboundCalls = Number(callTotals.outboundCalls || 0);
        const talkSeconds = Number(callTotals.talkSeconds || 0);
        const ticketCount = ticketAggregates[canonical] || 0;
        const daysArr = daysToCloseAggregates[canonical] || [];
        const avgDaysToClose = daysArr.length ? (daysArr.reduce((a, b) => a + b, 0) / daysArr.length) : 0;
        const openTickets = openAggregates[canonical] || 0;
        const avgPickupSeconds = totals.pickupCount > 0 ? (totals.pickupSum / totals.pickupCount) : 0;
        const avgDurationSeconds = totals.durationCount > 0 ? (totals.durationSum / totals.durationCount) : 0;
        const ticketOpenRate = ticketCount > 0 ? (totals.sessions + totalCallsForTech) / ticketCount : 0;
        return [
          displayName,
          totals.sessions || 0,
          totals.novaWave || 0,
          avgPickupSeconds / 86400,
          avgDurationSeconds / 86400,
          (totals.workSeconds || 0) / 86400,
          totalCallsForTech,
          inboundCalls,
          outboundCalls,
          talkSeconds / 86400,
          ticketCount,
          avgDaysToClose || 0,
          ticketOpenRate
        ];
      })
      .filter(Boolean)
      .sort((a, b) => b[1] - a[1]);

    const techDataStartRow = currentRow + 2;
    if (techRows.length > 0) {
      supportSheet.getRange(techDataStartRow, 1, techRows.length, tableHeaders.length).setValues(techRows);
      try {
        supportSheet.getRange(techDataStartRow, 2, techRows.length, 2).setNumberFormat('0');
        supportSheet.getRange(techDataStartRow, 4, techRows.length, 3).setNumberFormat('hh:mm:ss');
        supportSheet.getRange(techDataStartRow, 7, techRows.length, 3).setNumberFormat('0');
        supportSheet.getRange(techDataStartRow, 10, techRows.length, 1).setNumberFormat('hh:mm:ss');
        supportSheet.getRange(techDataStartRow, 11, techRows.length, 1).setNumberFormat('0');
        supportSheet.getRange(techDataStartRow, 12, techRows.length, 1).setNumberFormat('0.0');
        supportSheet.getRange(techDataStartRow, 13, techRows.length, 1).setNumberFormat('0.0%');
      } catch (e) { /* ignore formatting errors */ }
    }
    try {
      const techBandingRange = supportSheet.getRange(techHeaderRow, 1, Math.max(1, techRows.length + 1), tableHeaders.length);
      const techBanding = techBandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      techBanding.setHeaderRowColor('#1E3A8A')
                 .setFirstRowColor('#F5F7FB')
                 .setSecondRowColor('#FFFFFF')
                 .setFooterRowColor(null);
    } catch (e) { Logger.log('Support_Data: tech table banding failed: ' + e.toString()); }
    currentRow = techDataStartRow + techRows.length + 2;
    
    // Channel Performance Summary (daily wide format)
    supportSheet.getRange(currentRow, 1).setValue('Logmein Performance Summary (daily)');
    supportSheet.getRange(currentRow, 1).setFontSize(13).setFontWeight('bold').setFontColor('#0B1F50');
    const channelRow = currentRow + 1;

    const getTeamDayStats = (dateKey) => {
      const stats = teamDaily[dateKey];
      if (stats) return stats;
      const empty = createEmptyDayStats_();
      empty.techSet = new Set();
      return empty;
    };

    const headerRow = ['Metric'];
    dates.forEach(dateKey => headerRow.push(formatDisplayDate(dateKey)));
    headerRow.push('Totals/Averages');

    const totalSessionsRow = ['Total sessions'];
    let totalSessionsAll = 0;
    dates.forEach(dateKey => {
      const dayStats = getTeamDayStats(dateKey);
      const count = dayStats.sessions || 0;
      totalSessionsRow.push(count);
      totalSessionsAll += count;
    });
    totalSessionsRow.push(totalSessionsAll);

    const totalWorkRow = ['Total Work Time'];
    let totalWorkSecondsAll = 0;
    dates.forEach(dateKey => {
      const dayStats = getTeamDayStats(dateKey);
      const secs = dayStats.workSeconds || 0;
      totalWorkRow.push(secs / 86400);
      totalWorkSecondsAll += secs;
    });
    totalWorkRow.push(totalWorkSecondsAll / 86400);

    const avgSessionRow = ['Avg Session'];
    let sessionDurationSum = 0;
    let sessionDurationCount = 0;
    dates.forEach(dateKey => {
      const dayStats = getTeamDayStats(dateKey);
      if (dayStats.durationCount > 0) {
        const avg = dayStats.durationSum / dayStats.durationCount;
        avgSessionRow.push(avg / 86400);
        sessionDurationSum += avg;
        sessionDurationCount++;
      } else {
        avgSessionRow.push(0);
      }
    });
    avgSessionRow.push(sessionDurationCount > 0 ? (sessionDurationSum / sessionDurationCount) / 86400 : 0);

    const avgPickupRow = ['Avg Pick-up Speed'];
    let pickupSumTotal = 0;
    let pickupCountTotal = 0;
    dates.forEach(dateKey => {
      const dayStats = getTeamDayStats(dateKey);
      if (dayStats.pickupCount > 0) {
        const avg = dayStats.pickupSum / dayStats.pickupCount;
        avgPickupRow.push(avg / 86400);
        pickupSumTotal += avg * dayStats.pickupCount;
        pickupCountTotal += dayStats.pickupCount;
      } else {
        avgPickupRow.push(0);
      }
    });
    avgPickupRow.push(pickupCountTotal > 0 ? (pickupSumTotal / pickupCountTotal) / 86400 : 0);

    const pctRow = ['Percentage of daily calls/total sessions'];
    dates.forEach(dateKey => {
      const dayStats = getTeamDayStats(dateKey);
      const count = dayStats.sessions || 0;
      const pct = totalSessionsAll > 0 ? (count / totalSessionsAll) : 0;
      pctRow.push(pct);
    });
    pctRow.push(1);

    const avgRealRow = ['Average REAL', ...Array(dates.length).fill(0), 0];

    const techsWorkingRow = ['Total # of Techs working'];
    let totalTechsAll = 0;
    dates.forEach(dateKey => {
      const dayStats = getTeamDayStats(dateKey);
      const techCount = dayStats.techSet ? dayStats.techSet.size : 0;
      techsWorkingRow.push(techCount);
      totalTechsAll += techCount;
    });
    techsWorkingRow.push(totalTechsAll);

    const techHoursRow = ['Total hours of technicans'];
    dates.forEach((_, idx) => {
      const techCount = techsWorkingRow[idx + 1] || 0;
      techHoursRow.push(techCount * 8);
    });
    techHoursRow.push(totalTechsAll * 8);

    const avgSessionsPerTechRow = ['Average sessions per tech'];
    dates.forEach((_, idx) => {
      const techCount = techsWorkingRow[idx + 1] || 0;
      const sessionCount = totalSessionsRow[idx + 1] || 0;
      const avg = techCount > 0 ? sessionCount / techCount : 0;
      avgSessionsPerTechRow.push(Number(avg.toFixed(1)));
    });
    avgSessionsPerTechRow.push(Number((totalSessionsAll / Math.max(1, totalTechsAll)).toFixed(1)));

    const avgCallsPerTechRow = ['Average calls per tech'];
    dates.forEach((_, idx) => {
      const techCount = techsWorkingRow[idx + 1] || 0;
      const sessionCount = totalSessionsRow[idx + 1] || 0;
      const avg = techCount > 0 ? sessionCount / techCount : 0;
      avgCallsPerTechRow.push(Number(avg.toFixed(1)));
    });
    avgCallsPerTechRow.push(Number((totalSessionsAll / Math.max(1, totalTechsAll)).toFixed(1)));

    const summaryRows = [];
    summaryRows.push(headerRow);
    summaryRows.push(['LMI:', ...Array(dates.length + 1).fill('')]);
    summaryRows.push(totalSessionsRow);
    summaryRows.push(totalWorkRow);
    summaryRows.push(avgSessionRow);
    summaryRows.push(avgPickupRow);
    summaryRows.push(pctRow);
    summaryRows.push(avgRealRow);
    summaryRows.push(Array(dates.length + 2).fill(''));
    summaryRows.push(['Daily Calculations', ...Array(dates.length + 1).fill('')]);
    summaryRows.push(['# of techs working will be a total of techs that picked up sessions that day', ...Array(dates.length + 1).fill('')]);
    summaryRows.push(techsWorkingRow);
    summaryRows.push(techHoursRow);
    summaryRows.push(avgSessionsPerTechRow);
    summaryRows.push(avgCallsPerTechRow);

    const summaryHeight = summaryRows.length;
    supportSheet.getRange(channelRow, 1, summaryHeight, headerRow.length).setValues(summaryRows);
    supportSheet.getRange(channelRow, 1, 1, headerRow.length)
      .setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF').setFontSize(12);

    for (let i = 1; i < summaryHeight; i++) {
      const metricName = String(summaryRows[i][0] || '').toLowerCase();
      const range = supportSheet.getRange(channelRow + i, 2, 1, headerRow.length - 1);
      try {
        if (metricName.includes('time') || metricName.includes('duration')) {
          range.setNumberFormat('hh:mm:ss');
        } else if (metricName.includes('percentage')) {
          range.setNumberFormat('0.00');
        } else if (metricName.includes('average sessions') || metricName.includes('average calls')) {
          range.setNumberFormat('0.0');
        } else {
          range.setNumberFormat('0');
        }
      } catch (e) { /* ignore */ }
    }
    try {
      const channelBandingRange = supportSheet.getRange(channelRow, 1, summaryHeight, headerRow.length);
      const channelBanding = channelBandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      channelBanding.setHeaderRowColor('#1E3A8A')
                    .setFirstRowColor('#F5F7FB')
                    .setSecondRowColor('#FFFFFF')
                    .setFooterRowColor(null);
    } catch (e) { Logger.log('Support_Data: channel table banding failed: ' + e.toString()); }
    currentRow = channelRow + summaryHeight + 2;
    // Section header for Call data
    supportSheet.getRange(currentRow, 1).setValue('Call Data (Digium)');
    supportSheet.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold').setFontColor('#0B1F50');
    currentRow += 1;
    let callHeaderRowIndex = -1;
    let callRowCount = 0;
    let callHeaderLength = 0;
    if (digByDay && digByDay.ok && digByDay.rows && digByDay.rows.length) {
      const categories = Array.isArray(digByDay.categories) ? digByDay.categories : [];
      if (categories.length > 0) {
        const callHeader = ['Metric'].concat(categories.map(dateStr => {
          const dt = new Date(dateStr + 'T00:00:00');
          return isNaN(dt.getTime()) ? dateStr : `${dt.getMonth() + 1}/${dt.getDate()}/${dt.getFullYear()}`;
        }));
      supportSheet.getRange(currentRow, 1, 1, callHeader.length).setValues([callHeader]);
      supportSheet.getRange(currentRow, 1, 1, callHeader.length)
        .setFontWeight('bold').setFontSize(12).setBackground('#1E3A8A').setFontColor('#FFFFFF');
      callHeaderRowIndex = currentRow;
      callHeaderLength = callHeader.length;
      
        const processedRows = digByDay.rows.map(row => {
        const metricName = String(row[0] || '');
        const isDuration = metricName.toLowerCase().includes('duration') || metricName.toLowerCase().includes('talking');
        const out = [metricName];
        for (let i = 1; i < callHeader.length; i++) {
            const rawVal = row[i] != null ? Number(row[i]) : 0;
            const normalizedVal = !isNaN(rawVal) ? rawVal : 0;
          if (isDuration) {
              out.push(normalizedVal >= 0 ? normalizedVal / 86400 : 0);
          } else {
              out.push(normalizedVal);
          }
        }
        return out;
      });
      
        supportSheet.getRange(currentRow + 1, 1, processedRows.length, callHeader.length).setValues(processedRows);
        processedRows.forEach((row, idx) => {
          const metricName = String(row[0] || '').toLowerCase();
          const range = supportSheet.getRange(currentRow + 1 + idx, 2, 1, callHeader.length - 1);
          try {
            if (metricName.includes('duration') || metricName.includes('talking')) {
              range.setNumberFormat('hh:mm:ss');
            } else {
              range.setNumberFormat('0');
            }
          } catch (e) { /* ignore */ }
        });
        callRowCount = processedRows.length + 1;
        currentRow += processedRows.length + 2;
      } else {
        const callHeader = ['Metric', 'Total'];
        supportSheet.getRange(currentRow, 1, 1, callHeader.length).setValues([callHeader]);
        supportSheet.getRange(currentRow, 1, 1, callHeader.length)
          .setFontWeight('bold').setFontSize(12).setBackground('#1E3A8A').setFontColor('#FFFFFF');
        callHeaderRowIndex = currentRow;
        callHeaderLength = callHeader.length;
        
        const processedRows = digByDay.rows.map(row => {
          const metricName = String(row[0] || '');
          const metricLower = metricName.toLowerCase();
          const isDuration = metricLower.includes('duration') || metricLower.includes('talking');
          let aggregatedValue = 0;
          if (metricLower.includes('total calls') && !metricLower.includes('incoming') && !metricLower.includes('outgoing')) {
            aggregatedValue = totalCalls;
          } else if (metricLower.includes('incoming')) {
            aggregatedValue = totalIncomingCalls;
          } else if (metricLower.includes('outgoing')) {
            aggregatedValue = totalOutgoingCalls;
          } else if (metricLower.includes('talking duration')) {
            aggregatedValue = totalTalkingSeconds;
          } else if (metricLower.includes('call duration')) {
            aggregatedValue = totalCallSeconds;
          } else if (metricLower.includes('avg talking')) {
            aggregatedValue = totalCalls > 0 ? totalTalkingSeconds / totalCalls : 0;
          } else if (metricLower.includes('avg call')) {
            aggregatedValue = totalCalls > 0 ? totalCallSeconds / totalCalls : 0;
          }
          const converted = isDuration ? (aggregatedValue / 86400) : aggregatedValue;
          return [metricName, converted];
        });
        
        supportSheet.getRange(currentRow + 1, 1, processedRows.length, callHeader.length).setValues(processedRows);
        processedRows.forEach((row, idx) => {
          const metricName = String(row[0] || '').toLowerCase();
          const range = supportSheet.getRange(currentRow + 1 + idx, 2, 1, 1);
          try {
            if (metricName.includes('duration') || metricName.includes('talking')) {
              range.setNumberFormat('hh:mm:ss');
            } else {
              range.setNumberFormat('0');
            }
          } catch (e) { /* ignore */ }
        });
        callRowCount = processedRows.length + 1;
        currentRow += processedRows.length + 2;
      }
    } else if (digByAccount && digByAccount.ok && digByAccount.rows) {
      const callHeader = ['Metric', 'Total'];
      supportSheet.getRange(currentRow, 1, 1, callHeader.length).setValues([callHeader]);
      supportSheet.getRange(currentRow, 1, 1, callHeader.length)
        .setFontWeight('bold').setFontSize(12).setBackground('#1E3A8A').setFontColor('#FFFFFF');
      callHeaderRowIndex = currentRow;
      callHeaderLength = callHeader.length;

      const processedRows = digByAccount.rows.map(row => {
        const metricName = String(row[0] || '');
        const isDuration = metricName.toLowerCase().includes('duration') || metricName.toLowerCase().includes('talking');
        const val = row[1] != null ? Number(row[1]) : 0;
        const converted = isDuration ? (val && !isNaN(val) && val >= 0 ? val / 86400 : 0) : (isNaN(val) ? 0 : val);
        return [metricName, converted];
      });

      if (processedRows.length > 0) {
        supportSheet.getRange(currentRow + 1, 1, processedRows.length, callHeader.length).setValues(processedRows);
        processedRows.forEach((row, idx) => {
          const metricName = String(row[0] || '').toLowerCase();
          const range = supportSheet.getRange(currentRow + 1 + idx, 2, 1, 1);
          try {
            if (metricName.includes('duration') || metricName.includes('talking')) {
              range.setNumberFormat('hh:mm:ss');
            } else {
              range.setNumberFormat('0');
            }
          } catch (e) { /* ignore */ }
        });
        callRowCount = processedRows.length + 1;
        currentRow += processedRows.length + 2;
      } else {
        callRowCount = 1;
        currentRow += 1;
      }
    } else {
      supportSheet.getRange(currentRow, 1).setValue('No call data available for the selected range.');
      currentRow += 2;
    }
    if (callRowCount > 0 && callHeaderRowIndex >= 0 && callHeaderLength > 0) {
      try {
        const callBandingRange = supportSheet.getRange(callHeaderRowIndex, 1, callRowCount, callHeaderLength);
        const callBanding = callBandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
        callBanding.setHeaderRowColor('#1E3A8A')
                   .setFirstRowColor('#F5F7FB')
                   .setSecondRowColor('#FFFFFF')
                   .setFooterRowColor(null);
      } catch (e) { Logger.log('Support_Data: call table banding failed: ' + e.toString()); }
    }

    const getCallMetricsForCanonical = (canonical) => {
      if (!canonical) return null;
      if (callMetricsByCanonical[canonical]) return callMetricsByCanonical[canonical];
      if (canonical === 'ahmed talal' && callMetricsByCanonical['eddie talal']) return callMetricsByCanonical['eddie talal'];
      return null;
    };
    const toTitleCaseCanonical = (canonical) => {
      if (!canonical) return 'Unknown';
      return canonical.split(/\s+/).map(part => part ? part.charAt(0).toUpperCase() + part.slice(1) : '').join(' ').trim() || 'Unknown';
    };

    const perTechKeysSorted = Object.keys(perTechDaily || {})
      .map(canonical => {
        const name = ensureDisplayName ? ensureDisplayName(canonical) : toTitleCaseCanonical(canonical);
        return { canonical, name };
      })
      .filter(item => item.name && !isExcludedTechnician_(item.name))
      .sort((a, b) => a.name.localeCompare(b.name));

    perTechKeysSorted.forEach(({ canonical, name }) => {
      const dayStatsMap = perTechDaily[canonical] || {};
      const hasSessionData = dates.some(dateKey => {
        const dayStats = dayStatsMap[dateKey];
        return dayStats && dayStats.sessions > 0;
      });
      const callMetrics = getCallMetricsForCanonical(canonical);
      const hasCallData = callMetrics && (
        Number(callMetrics.totalCalls || 0) > 0 ||
        Number(callMetrics.inboundCalls || 0) > 0 ||
        Number(callMetrics.outboundCalls || 0) > 0 ||
        Number(callMetrics.talkSeconds || 0) > 0 ||
        Number(callMetrics.callSeconds || 0) > 0
      );
      if (!hasSessionData && !hasCallData) {
        return;
      }

      supportSheet.getRange(currentRow, 1).setValue(`👤 ${name} - Daily Performance`);
      supportSheet.getRange(currentRow, 1).setFontSize(12).setFontWeight('bold').setFontColor('#1E3A8A');
      currentRow += 1;

      const sessionHeader = ['Metric'];
      dates.forEach(d => sessionHeader.push(formatDisplayDate(d)));
      sessionHeader.push('Totals/Averages');

      const sessionRows = [];

      const totalSessionsRow = ['Total sessions'];
      let totalSessionsAll = 0;
      dates.forEach(dateKey => {
        const dayStats = dayStatsMap[dateKey] || {};
        const count = dayStats.sessions || 0;
        totalSessionsRow.push(count);
        totalSessionsAll += count;
      });
      totalSessionsRow.push(totalSessionsAll);
      sessionRows.push(totalSessionsRow);

      const totalActiveRow = ['Total Active Time'];
      let totalActiveAll = 0;
      dates.forEach(dateKey => {
        const secs = (dayStatsMap[dateKey] && dayStatsMap[dateKey].activeSeconds) || 0;
        totalActiveRow.push(secs / 86400);
        totalActiveAll += secs;
      });
      totalActiveRow.push(totalActiveAll / 86400);
      sessionRows.push(totalActiveRow);

      const totalWorkRow = ['Total Work Time'];
      let totalWorkAll = 0;
      dates.forEach(dateKey => {
        const secs = (dayStatsMap[dateKey] && dayStatsMap[dateKey].workSeconds) || 0;
        totalWorkRow.push(secs / 86400);
        totalWorkAll += secs;
      });
      totalWorkRow.push(totalWorkAll / 86400);
      sessionRows.push(totalWorkRow);

      const totalLoginRow = ['Total Login Time'];
      let totalLoginAll = 0;
      dates.forEach(dateKey => {
        const secs = (dayStatsMap[dateKey] && dayStatsMap[dateKey].loginSeconds) || 0;
        totalLoginRow.push(secs / 86400);
        totalLoginAll += secs;
      });
      totalLoginRow.push(totalLoginAll / 86400);
      sessionRows.push(totalLoginRow);

      const avgSessionRow = ['Avg Session'];
      let totalAvgSeconds = 0;
      let daysWithAvg = 0;
      dates.forEach(dateKey => {
        const dayStats = dayStatsMap[dateKey] || {};
        if (dayStats.durationCount > 0) {
          const avg = dayStats.durationSum / dayStats.durationCount;
          avgSessionRow.push(avg / 86400);
          totalAvgSeconds += avg;
          daysWithAvg++;
        } else {
          avgSessionRow.push(0);
        }
      });
      avgSessionRow.push(daysWithAvg > 0 ? (totalAvgSeconds / daysWithAvg) / 86400 : 0);
      sessionRows.push(avgSessionRow);

      const avgPickupRow = ['Avg Pick-up Speed'];
      let totalPickupSeconds = 0;
      let totalPickupCount = 0;
      dates.forEach(dateKey => {
        const dayStats = dayStatsMap[dateKey] || {};
        if (dayStats.pickupCount > 0) {
          const avg = dayStats.pickupSum / dayStats.pickupCount;
          avgPickupRow.push(avg / 86400);
          totalPickupSeconds += avg * dayStats.pickupCount;
          totalPickupCount += dayStats.pickupCount;
        } else {
          avgPickupRow.push(0);
        }
      });
      avgPickupRow.push(totalPickupCount > 0 ? (totalPickupSeconds / totalPickupCount) / 86400 : 0);
      sessionRows.push(avgPickupRow);

      const longestSessionRow = ['Longest Session Time'];
      let longestOverall = 0;
      dates.forEach(dateKey => {
        const secs = (dayStatsMap[dateKey] && dayStatsMap[dateKey].longestSeconds) || 0;
        if (secs > longestOverall) longestOverall = secs;
        longestSessionRow.push(secs / 86400);
      });
      longestSessionRow.push(longestOverall / 86400);
      sessionRows.push(longestSessionRow);

      const sessionsPerHourRow = ['Sessions Per Hour'];
      let totalSessionsPerHourBase = 0;
      dates.forEach(dateKey => {
        const count = (dayStatsMap[dateKey] && dayStatsMap[dateKey].sessions) || 0;
        const value = count > 0 ? Number((count / 8).toFixed(1)) : 0;
        sessionsPerHourRow.push(value);
        totalSessionsPerHourBase += count;
      });
      const totalSessionsPerHour = totalSessionsPerHourBase > 0 ? Number(((totalSessionsPerHourBase / dates.length) / 8).toFixed(1)) : 0;
      sessionsPerHourRow.push(totalSessionsPerHour);
      sessionRows.push(sessionsPerHourRow);

      supportSheet.getRange(currentRow, 1, 1, sessionHeader.length).setValues([sessionHeader]);
      supportSheet.getRange(currentRow, 1, 1, sessionHeader.length)
        .setFontWeight('bold').setFontSize(10).setBackground('#E8F1FF').setFontColor('#1E3A8A');
      supportSheet.getRange(currentRow + 1, 1, sessionRows.length, sessionHeader.length).setValues(sessionRows);
      sessionRows.forEach((rowVals, idx) => {
        const metricName = String(rowVals[0] || '').toLowerCase();
        const range = supportSheet.getRange(currentRow + 1 + idx, 2, 1, sessionHeader.length - 1);
        try {
          if (metricName.includes('time') || metricName.includes('duration') || metricName.includes('pickup')) {
            range.setNumberFormat('hh:mm:ss');
          } else if (metricName.includes('sessions per hour')) {
            range.setNumberFormat('0.0');
          } else {
            range.setNumberFormat('0');
          }
        } catch (e) { /* ignore */ }
      });
      currentRow += sessionRows.length + 2;

      const callDailyEntry = callDailyPerCanonical[canonical] || null;
      const perDayCallStats = callDailyEntry && callDailyEntry.perDay ? callDailyEntry.perDay : {};
      const callTotals = callDailyEntry && callDailyEntry.totals ? callDailyEntry.totals : {};
      const totalCallsOverall = (callTotals.totalCalls != null ? callTotals.totalCalls : Number(callMetrics.totalCalls || 0));
      const inboundOverall = (callTotals.inbound != null ? callTotals.inbound : Number(callMetrics.inboundCalls || 0));
      const outboundOverall = (callTotals.outbound != null ? callTotals.outbound : Number(callMetrics.outboundCalls || 0));
      const talkSecondsOverall = (callTotals.talkSeconds != null ? callTotals.talkSeconds : Number(callMetrics.talkSeconds || 0));
      const callSecondsOverall = (callTotals.callSeconds != null ? callTotals.callSeconds : Number(callMetrics.callSeconds || 0));

      if (hasCallData || Object.keys(perDayCallStats).length) {
        supportSheet.getRange(currentRow, 1).setValue(`📞 ${name} - Daily Call Data`);
        supportSheet.getRange(currentRow, 1).setFontSize(11).setFontWeight('bold').setFontColor('#1E3A8A');
        currentRow += 1;

        const callHeader = ['Metric'];
        dates.forEach(d => callHeader.push(formatDisplayDate(d)));
        callHeader.push('Totals/Averages');

        const callRows = [];
        const buildCallRow = (label, valueExtractor, totalValue, asDuration) => {
          const row = [label];
          dates.forEach(dateKey => {
            const entry = perDayCallStats[dateKey] || {};
            const val = valueExtractor(entry) || 0;
            row.push(asDuration ? val / 86400 : val);
          });
          row.push(asDuration ? totalValue / 86400 : totalValue);
          return row;
        };

        callRows.push(buildCallRow('Total Calls', entry => entry.totalCalls, totalCallsOverall, false));
        callRows.push(buildCallRow('Incoming Calls', entry => entry.inbound, inboundOverall, false));
        callRows.push(buildCallRow('Outgoing Calls', entry => entry.outbound, outboundOverall, false));
        callRows.push(buildCallRow('Total Talk Time', entry => entry.talkSeconds, talkSecondsOverall, true));
        callRows.push(buildCallRow('Total Call Duration', entry => entry.callSeconds, callSecondsOverall, true));

        const avgTalkRow = ['Avg Talk Time per Call'];
        dates.forEach(dateKey => {
          const entry = perDayCallStats[dateKey] || {};
          const tc = entry.totalCalls || 0;
          avgTalkRow.push(tc > 0 ? (entry.talkSeconds || 0) / tc / 86400 : 0);
        });
        avgTalkRow.push(totalCallsOverall > 0 ? (talkSecondsOverall / totalCallsOverall) / 86400 : 0);
        callRows.push(avgTalkRow);

        const avgCallDurationRow = ['Avg Call Duration'];
        dates.forEach(dateKey => {
          const entry = perDayCallStats[dateKey] || {};
          const tc = entry.totalCalls || 0;
          avgCallDurationRow.push(tc > 0 ? (entry.callSeconds || 0) / tc / 86400 : 0);
        });
        avgCallDurationRow.push(totalCallsOverall > 0 ? (callSecondsOverall / totalCallsOverall) / 86400 : 0);
        callRows.push(avgCallDurationRow);

        supportSheet.getRange(currentRow, 1, 1, callHeader.length).setValues([callHeader]);
        supportSheet.getRange(currentRow, 1, 1, callHeader.length)
          .setFontWeight('bold').setFontSize(10).setBackground('#E8F1FF').setFontColor('#1E3A8A');
        supportSheet.getRange(currentRow + 1, 1, callRows.length, callHeader.length).setValues(callRows);
        callRows.forEach((rowVals, idx) => {
          const metricName = String(rowVals[0] || '').toLowerCase();
          const range = supportSheet.getRange(currentRow + 1 + idx, 2, 1, callHeader.length - 1);
          try {
            if (metricName.includes('time') || metricName.includes('duration')) {
              range.setNumberFormat('hh:mm:ss');
            } else {
              range.setNumberFormat('0');
            }
          } catch (e) { /* ignore */ }
        });
        currentRow += callRows.length + 2;
      }

      currentRow += 1;
    });
    
    // Formatting
    supportSheet.setColumnWidth(1, 200);
    supportSheet.setColumnWidth(2, 120);
    supportSheet.setColumnWidth(3, 120);
    supportSheet.setColumnWidth(4, 120);
    supportSheet.setColumnWidth(5, 120);
    supportSheet.setColumnWidth(6, 120);
    supportSheet.setFrozenRows(1);
    try {
      if (typeof formatAllDurationColumns === 'function') formatAllDurationColumns();
    } catch (e) { Logger.log('formatAllDurationColumns failed after support data generation: ' + e.toString()); }
    Logger.log('Support Data sheet created');
  } catch (e) {
    Logger.log('createSupportDataSheet_ error: ' + e.toString());
  }
}
function refreshAnalyticsDashboard_(startDate, endDate, perfMapOpt) {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Get or create Analytics_Dashboard sheet (preserve layout if exists)
    let dashboardSheet = ss.getSheetByName('Analytics_Dashboard');
    if (!dashboardSheet) {
      createMainAnalyticsPage_(ss); // Create structure if doesn't exist
      dashboardSheet = ss.getSheetByName('Analytics_Dashboard');
    } else {
      // Update formulas with current column references (in case headers changed)
      const kpiRow = 4;
      const sessionIdCol = getSessionsColumnByHeader_(['Session ID', 'session_id', 'Session ID']);
      const statusCol = getSessionsColumnByHeader_(['Status', 'session_status', 'Status']);
      const totalTimeCol = getSessionsColumnByHeader_(['Total Time', 'duration_total_seconds', 'Total Time']);
      const waitingTimeCol = getSessionsColumnByHeader_(['Waiting Time', 'pickup_seconds', 'Waiting Time']);
      
      // Update KPI card formulas (cards are at positions: 0,1,3,4,5,6,7 - skip 2 which is Nova Wave)
      const kpiCards = [
        {row: 0, formula: `=COUNTA(Sessions!${sessionIdCol}2:${sessionIdCol})`},
        {row: 1, formula: `=COUNTIFS(Sessions!${statusCol}2:${statusCol}, "Active")`},
        {row: 3, formula: `=IF(COUNT(Sessions!${totalTimeCol}2:${totalTimeCol})>0, AVERAGE(Sessions!${totalTimeCol}2:${totalTimeCol})/86400, 0)`},
        {row: 4, formula: `=IF(COUNT(Sessions!${waitingTimeCol}2:${waitingTimeCol})>0, AVERAGE(Sessions!${waitingTimeCol}2:${waitingTimeCol})/86400, 0)`},
        {row: 5, formula: `=IF(MAX(Sessions!${totalTimeCol}2:${totalTimeCol})>0, MAX(Sessions!${totalTimeCol}2:${totalTimeCol})/86400, 0)`},
        {row: 6, formula: `=IF(COUNT(Sessions!${waitingTimeCol}2:${waitingTimeCol})>0, COUNTIFS(Sessions!${waitingTimeCol}2:${waitingTimeCol}, "<=60")/COUNT(Sessions!${waitingTimeCol}2:${waitingTimeCol}), 0)`},
        {row: 7, formula: `=IF(COUNTA(Sessions!${sessionIdCol}2:${sessionIdCol})>0, ROUND(COUNTA(Sessions!${sessionIdCol}2:${sessionIdCol})/8, 1), 0)`}
      ];
      
      const cardWidth = 3;
      const cardHeight = 3;
      kpiCards.forEach(kpi => {
        const cardTop = kpiRow + Math.floor(kpi.row / 3) * cardHeight;
        const cardLeft = (kpi.row % 3) * cardWidth + 1;
        const valRange = dashboardSheet.getRange(cardTop + 1, cardLeft, 1, cardWidth);
        valRange.merge();
        valRange.setFormula(kpi.formula);
        // Re-apply number formats
        try {
          if (kpi.row === 3 || kpi.row === 4 || kpi.row === 5) {
            valRange.setNumberFormat('hh:mm:ss');
          } else if (kpi.row === 6) {
            valRange.setNumberFormat('0.0%');
          } else if (kpi.row === 0) {
            valRange.setNumberFormat('0');
          } else if (kpi.row === 7) {
            valRange.setNumberFormat('0.0');
          }
        } catch (e) {}
      });
      
      // Clear existing data but preserve layout/headers
      const lastRow = dashboardSheet.getLastRow();
      const lastCol = dashboardSheet.getLastColumn();
      if (lastRow > 2 && lastCol > 0) {
        // Keep header rows (1-2) and clear data rows
        dashboardSheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
      }
    }
    const cfg = getCfg_();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    const allDataRaw = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const allData = filterOutExcludedTechnicians_(headers, allDataRaw);
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
    const startMillis = startDate.getTime();
    const endMillis = endDate.getTime();
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx]) return false;
      try {
        const rowObj = row[startIdx] instanceof Date ? row[startIdx] : new Date(row[startIdx]);
        if (!(rowObj instanceof Date) || isNaN(rowObj)) return false;
        const rowMillis = rowObj.getTime();
        if (rowMillis < startMillis || rowMillis > endMillis) return false;
        if (techIdx >= 0 && isExcludedTechnician_(row[techIdx])) return false;
        return true;
      } catch (e) {
        return false;
      }
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
    
    // Use the dashboard sheet we got/created at the start of this function
    if (dashboardSheet) {
  // Compute positions so refresh uses the same layout as creation
  const kpiRow = 4;
  const kpiCardsCount = 8; // keep in sync with createMainAnalyticsPage_
  const cardHeight = 3;
  const cardRows = Math.ceil(kpiCardsCount / 3) * cardHeight;
  const tableRow = kpiRow + cardRows + 2;
      
      // Update Team Performance section (positioned after KPI cards)
      const teamTitleRow = tableRow;
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

        // Team header (blue)
        if (typeof teamHeaderRow !== 'undefined') {
          dashboardSheet.getRange(teamHeaderRow, 1, 1, teamHeaders.length).setBackground('#1A73E8').setFontColor('#FFFFFF').setFontWeight('bold');
        }

        // Active Sessions header (blue)
        if (typeof activeHeaderRow !== 'undefined') {
          dashboardSheet.getRange(activeHeaderRow, 1, 1, activeHeaders.length).setBackground('#1A73E8').setFontColor('#FFFFFF').setFontWeight('bold');
        }

        // Ensure team data rows have dark font color (so values are visible)
        if (typeof teamDataRow !== 'undefined' && teamRows && teamRows.length > 0) {
          dashboardSheet.getRange(teamDataRow, 1, teamRows.length, teamHeaders.length).setFontColor('#0F172A');
        }

                  // Don't apply professional table styling - it overwrites our custom header colors
                  // try { applyProfessionalTableStyling_(dashboardSheet, dashboardSheet.getLastColumn()); } catch (e) { /* non-fatal */ }
      } catch (e) { Logger.log('Re-applying header styles failed: ' + e.toString()); }
      
      dashboardSheet.getRange(2, 5).setValue(new Date().toLocaleString());
    }
  } catch (e) {
    Logger.log('refreshAnalyticsDashboard_ error: ' + e.toString());
  }
}
function generateTechnicianTabs_(startDate, endDate, digiumDatasetOpt, extensionMetaOpt) {
  try {
    const ss = SpreadsheetApp.getActive();
    const cfg = getCfg_();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return;
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return;
    // Load extension -> technician mapping for per-account Digium pulls (sheet: extension_map)
    const extMap = getExtensionMap_();
    const normalizeSheetName = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    
    const allDataRaw = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const allData = filterOutExcludedTechnicians_(headers, allDataRaw);
    const startStr = startDate.toISOString().split('T')[0];
    const endStr = endDate.toISOString().split('T')[0];
    const getHeaderIndex = (variants) => {
      for (const v of variants) {
        const i = headers.findIndex(h => String(h || '').toLowerCase().trim() === v.toLowerCase().trim());
        if (i >= 0) return i;
      }
      return -1;
    };
    // Use exact header names from API as provided by user
    const startIdx = getHeaderIndex(['Start Time', 'start_time', 'start time']);
    const techIdx = getHeaderIndex(['Technician Name', 'technician_name', 'technician name']);
    const techIdIdx = getHeaderIndex(['Technician ID', 'technician_id', 'technician id']);
    const statusIdx = getHeaderIndex(['Status', 'session_status', 'status']);
    const durationIdx = getHeaderIndex(['Total Time', 'total_time', 'duration_total_seconds', 'Total Time']);
    const pickupIdx = getHeaderIndex(['Waiting Time', 'waiting_time', 'pickup_seconds', 'Waiting Time']);
    const workIdx = getHeaderIndex(['Work Time', 'work_time', 'duration_work_seconds', 'Work Time']);
    const activeIdx = getHeaderIndex(['Active Time', 'active_time', 'duration_active_seconds', 'Active Time']);
    const customerIdx = getHeaderIndex(['Your Name:', 'Customer Name', 'customer_name', 'Your Name:']);
    const sessionIdIdx = getHeaderIndex(['Session ID', 'session_id', 'session id', 'Session ID']);
    const phoneIdx = getHeaderIndex(['Your Phone #:', 'caller_phone', 'phone', 'Your Phone #:']);
    const companyIdx = getHeaderIndex(['Company name:', 'Company Name', 'company_name', 'Company name:']);
    const callingCardIdx = getHeaderIndex(['Calling Card', 'calling_card', 'calling card', 'Calling Card']);
    const channelNameIdx = getHeaderIndex(['Channel Name', 'channel_name', 'channel name', 'Channel Name']);
    
    // Log header indices for debugging
    Logger.log(`Header indices - Start: ${startIdx}, Tech: ${techIdx}, Status: ${statusIdx}, Duration: ${durationIdx}, Pickup: ${pickupIdx}, Work: ${workIdx}, Active: ${activeIdx}, SessionID: ${sessionIdIdx}, Channel: ${channelNameIdx}`);
    Logger.log(`generateTechnicianTabs_: Filtering Sessions data for date range ${startStr} to ${endStr} (total rows in sheet: ${allData.length})`);
    
    const startMillis = startDate.getTime();
    const endMillis = endDate.getTime();
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx]) return false;
      try {
        const rowObj = row[startIdx] instanceof Date ? row[startIdx] : new Date(row[startIdx]);
        if (!(rowObj instanceof Date) || isNaN(rowObj)) return false;
        const rowMillis = rowObj.getTime();
        if (rowMillis < startMillis || rowMillis > endMillis) return false;
        if (techIdx >= 0 && isExcludedTechnician_(row[techIdx])) return false;
        return true;
      } catch (e) {
        return false;
      }
    });
    
    Logger.log(`generateTechnicianTabs_: Filtered ${filtered.length} rows matching date range ${startStr} to ${endStr}`);
    const sfMetrics = collectSalesforceTicketMetrics_(startDate, endDate);
  const extensionMeta = extensionMetaOpt || getActiveExtensionMetadata_();
    const activeExtensions = (extensionMeta && Array.isArray(extensionMeta.list))
      ? extensionMeta.list.map(ext => String(ext || '').trim()).filter(Boolean)
      : [];
  const digiumDataset = digiumDatasetOpt || getDigiumDataset_(startDate, endDate, extensionMeta);
  const callMetricsByCanonical = (digiumDataset && digiumDataset.callMetricsByCanonical) ? digiumDataset.callMetricsByCanonical : {};
  // Build tech list from roster (extension_map) and from filtered data
  // Normalize names to handle variations like "Tomer" vs "Tomer Reiter" (same person)
  // BUT keep distinct people with same first name separate (e.g., "Oscar Ocampo" vs "Oscar Umana")
  const normalizeTechNameForGrouping = technicianFirstNameKey_;
  const areNameVariations = areTechnicianNameVariations_;
  
  const rosterNames = (getRosterTechnicianNames_() || []).filter(name => !isExcludedTechnician_(name));
  const canonicalIsMapped = (name) => {
    const canonical = canonicalTechnicianName_(name);
    const normFull = normalizeTechnicianNameFull_(name);
    const firstKey = technicianFirstNameKey_(name);
    return (canonical && extMap[canonical]) ||
           (normFull && extMap[normFull]) ||
           (firstKey && extMap[firstKey]);
  };
  const techSet = new Set();
  rosterNames.forEach(name => {
    if (canonicalIsMapped(name)) techSet.add(String(name));
  });
  filtered.forEach(row => {
    const t = row[techIdx];
    if (t && canonicalIsMapped(t)) techSet.add(String(t));
  });
  if (sfMetrics && sfMetrics.perCanonical) {
    Object.values(sfMetrics.perCanonical).forEach(entry => {
      (entry.rawNames || []).forEach(rawName => {
        if (rawName && canonicalIsMapped(rawName)) techSet.add(String(rawName));
      });
    });
  }
  
  // Group technicians only if they are variations of the same person (e.g., "Tomer" and "Tomer Reiter")
  // Do NOT group distinct people with same first name (e.g., "Oscar Ocampo" and "Oscar Umana")
  const techGroups = {};
  Array.from(techSet).forEach(techName => {
    // Check if this name is a variation of an existing group
    let foundGroup = false;
    for (const [groupKey, groupName] of Object.entries(techGroups)) {
      if (areNameVariations(techName, groupName)) {
        // This is a variation, use the existing group
        foundGroup = true;
        break;
      }
    }
    
    if (!foundGroup) {
      // This is a new person (or variation not yet grouped)
      // Use full normalized name as key to keep distinct people separate
      const normalizedFull = String(techName).trim().toLowerCase();
      // Prefer roster name if available
      const rosterMatch = rosterNames.find(r => areNameVariations(r, techName));
      techGroups[normalizedFull] = rosterMatch || techName;
    }
  });
  // Use grouped names (this ensures "Tomer" and "Tomer Reiter" map to one dashboard, but keeps Oscars separate)
  const techs = Object.values(techGroups);
  // Track which sheets we've processed to avoid duplicates (use first name for grouping)
  // This ensures "Tomer" and "Tomer Reiter" map to the same sheet
  const processedSheets = new Set();
  const allowedSafeNames = new Set();
  // Process all techs (both roster and those with data)
    // Track processed sheet names to prevent duplicates
    const processedSheetNames = new Set();
    
    // Map "Eddie Talal" to "Ahmed Talal" for matching (same person) - defined once for reuse
    const normalizeForMatching = canonicalTechnicianName_;

    // Build lookup tables so Salesforce ticket owners map cleanly to technician dashboards
    const exactTechMap = {};
    techs.forEach(tech => {
      const key = normalizeForMatching(tech);
      if (key) exactTechMap[key] = tech;
    });

    const firstNameMap = {};
    const duplicateFirstNames = new Set();
    techs.forEach(tech => {
      const firstKey = normalizeTechNameForGrouping(tech);
      if (!firstKey) return;
      if (firstNameMap[firstKey] && firstNameMap[firstKey] !== tech) {
        duplicateFirstNames.add(firstKey);
      } else {
        firstNameMap[firstKey] = tech;
      }
    });

    const matchTechForSalesforceName = (rawName) => {
      if (!rawName) return null;
      const normalizedFull = normalizeForMatching(rawName);
      if (normalizedFull && exactTechMap[normalizedFull]) {
        return exactTechMap[normalizedFull];
      }
      const firstKey = normalizeTechNameForGrouping(rawName);
      if (firstKey && firstNameMap[firstKey] && !duplicateFirstNames.has(firstKey)) {
        return firstNameMap[firstKey];
      }
      for (const tech of techs) {
        if (areNameVariations(rawName, tech)) {
          return tech;
        }
      }
      return null;
    };

    const ticketStatsByTech = {};
    if (sfMetrics && sfMetrics.perCanonical) {
      Object.keys(sfMetrics.perCanonical).forEach(canonical => {
        const entry = sfMetrics.perCanonical[canonical];
        if (!entry) return;
        const matchedTechs = new Set();
        const rawNames = entry.rawNames || [];
        rawNames.forEach(rawName => {
            const match = matchTechForSalesforceName(rawName);
            if (match) matchedTechs.add(match);
          });
        const canonicalMatch = matchTechForSalesforceName(canonical);
        if (canonicalMatch) matchedTechs.add(canonicalMatch);
        if (!matchedTechs.size) return;
        matchedTechs.forEach(techName => {
          if (!ticketStatsByTech[techName]) {
            ticketStatsByTech[techName] = { created: 0, closed: 0, open: 0, issues: {} };
          }
          const stats = ticketStatsByTech[techName];
          stats.created += entry.created || 0;
          stats.closed += entry.closed || 0;
          stats.open += entry.open || 0;
          if (entry.issues) {
            Object.keys(entry.issues).forEach(label => {
              const count = entry.issues[label] || 0;
              stats.issues[label] = (stats.issues[label] || 0) + count;
            });
          }
    const canonicalKey = canonicalTechnicianName_(techName);
    if (canonicalKey && !ticketStatsByTech[canonicalKey]) {
      ticketStatsByTech[canonicalKey] = stats;
          }
        });
      });
    }
    for (const techName of techs) {
      if (isExcludedTechnician_(techName)) {
        Logger.log(`Skipping excluded technician ${techName}`);
        continue;
      }
      // Standardize sheet name to "FirstName_LastName" format
      // First normalize name (handles Eddie/Ahmed mapping)
      const normalizedTechName = normalizeForMatching(techName);
      const nameParts = String(normalizedTechName).trim().split(/\s+/).filter(Boolean);
      let safeName;
      if (nameParts.length >= 2) {
        // Use first name + last name
        const firstName = nameParts[0].replace(/[^a-zA-Z0-9]/g, '');
        const lastName = nameParts[nameParts.length - 1].replace(/[^a-zA-Z0-9]/g, '');
        safeName = `${firstName}_${lastName}`.substring(0, 30);
      } else if (nameParts.length === 1) {
        // Only one name part - use it as-is
        safeName = nameParts[0].replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
      } else {
        // Fallback: use first name grouping
        const techNameFirstForSheet = normalizeTechNameForGrouping(techName);
        safeName = techNameFirstForSheet.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
      }
      
      // Check for duplicate sheet names and skip if already processed
      if (processedSheetNames.has(safeName)) {
        Logger.log(`Skipping duplicate sheet name "${safeName}" for technician "${techName}"`);
        continue;
      }
      processedSheetNames.add(safeName);
    // Use full normalized name as key to keep distinct people separate (not just first name)
    const normKey = String(techName).trim().toLowerCase();
    
    // Check if we've already processed this (avoid duplicates within a single run)
    if (processedSheets.has(normKey)) {
      Logger.log(`Skipping duplicate processing for ${techName} (already processed)`);
      continue;
    }
    processedSheets.add(normKey);
    allowedSafeNames.add(normalizeSheetName(safeName));
    
    // Get existing sheet by exact name or by normalized match (to avoid multiple variants)
      let techSheet = ss.getSheetByName(safeName);
    if (!techSheet) {
      const reservedSheets = new Set(['Sessions','Analytics_Dashboard','Dashboard_Config','Daily_Summary','Support_Data','Progress','Advanced_Analytics','Digium_Raw','Digium_Calls','API_Smoke_Test']);
      // Try to find an existing sheet whose normalized name matches the technician name
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length && !techSheet; i++) {
        const nm = sheets[i].getName();
        if (reservedSheets.has(nm)) continue;
        if (normalizeSheetName(nm) === normKey) {
          techSheet = sheets[i];
        }
      }
      // As a fallback, try to find by A1 title containing the technician name
      if (!techSheet) {
        const target = techName.toLowerCase();
        for (let i = 0; i < sheets.length && !techSheet; i++) {
          const sh = sheets[i];
          const nm = sh.getName();
          if (reservedSheets.has(nm)) continue;
          try {
            const a1 = String(sh.getRange(1,1).getValue() || '').toLowerCase();
            if (a1.indexOf(target) !== -1) {
              techSheet = sh;
            }
          } catch (e) { /* ignore */ }
        }
      }
      // If still not found, create a new one using the canonical safe name
      if (!techSheet) {
        techSheet = ss.insertSheet(safeName);
        Logger.log(`Created new personal dashboard sheet: ${safeName}`);
      } else {
        // Clear existing sheet to repopulate (don't delete, just clear)
        resetSheetCompletely_(techSheet);
        Logger.log(`Reusing existing personal dashboard sheet: ${techSheet.getName()} (normalized match)`);
      }
    } else {
      // Clear existing sheet to repopulate (don't delete, just clear)
      resetSheetCompletely_(techSheet);
      Logger.log(`Cleared existing personal dashboard sheet: ${safeName}`);
    }
    
    // Match technician names more flexibly (case-insensitive, trim whitespace)
    // Also handle variations like "Tomer" vs "Tomer Reiter" and "Eddie Talal" vs "Ahmed Talal"
    const normalizeTechName = (name) => String(name || '').trim().toLowerCase();
    const normalizeTechNameFirst = (name) => {
      const normalized = normalizeTechName(name);
      return normalized.split(/\s+/)[0]; // Get first name only
    };
    
    // normalizeForMatching is already defined above, reuse it
    
    // Match technician names - prioritize exact matches, then known variations
    const techNameForMatching = normalizeForMatching(techName);
    const techNameNormalizedFull = normalizeTechName(techName);
    const techNameNormalizedFirst = normalizeTechNameFirst(techName);
    const techRows = filtered.filter(row => {
      const rowTech = row[techIdx];
      if (!rowTech) return false;
      const rowTechNormalized = normalizeForMatching(rowTech);
      const rowTechNormalizedFull = normalizeTechName(rowTech);
      const rowTechNormalizedFirst = normalizeTechNameFirst(rowTech);
      
      // First: Match by normalized name (handles "Eddie Talal" vs "Ahmed Talal")
      if (rowTechNormalized === techNameForMatching) return true;
      
      // Second: Match by full name (case-insensitive) - exact match
      if (rowTechNormalizedFull === techNameNormalizedFull) return true;
      
      // Third: Match by name variations (e.g., "Tomer" matches "Tomer Reiter")
      // Only do this if they share the same first name AND one is a substring of the other
      if (techNameNormalizedFirst === rowTechNormalizedFirst) {
        // Check if one name contains the other (variation match)
        if (techNameNormalizedFull.includes(rowTechNormalizedFull) || 
            rowTechNormalizedFull.includes(techNameNormalizedFull)) {
          return true;
        }
      }
      
      return false;
    });
    
    Logger.log(`${techName} - Found ${techRows.length} rows after filtering (filtered total: ${filtered.length}, allData total: ${allData.length})`);
    
    // Debug: Log sample technician names from filtered data to help diagnose matching issues
    if (filtered.length > 0 && techRows.length === 0) {
      const sampleTechNames = [...new Set(filtered.slice(0, 10).map(r => r[techIdx]).filter(Boolean))];
      Logger.log(`${techName} - WARNING: No matching rows found. Sample technician names in filtered data: ${sampleTechNames.join(', ')}`);
      Logger.log(`${techName} - Searching for: "${techName}" (normalized: "${normalizeTechName(techName)}")`);
    }
    
    // Check if there's Digium call data for this technician's extension
    const extMap = getExtensionMap_();
    const norm = (s) => String(s || '').trim().toLowerCase();
    const techNameNorm = normalizeForMatching(techName); // Use normalized name for lookup (handles Eddie/Ahmed mapping)
    const firstName = techNameNorm.split(/\s+/)[0];
    const extList = extMap[techNameNorm] || extMap[firstName] || [];
    const hasExtension = extList && extList.length > 0;
    
    Logger.log(`${techName} - Extension map lookup: found ${extList.length} extensions: ${extList.join(', ')} (searched: "${techNameNorm}", "${firstName}")`);
    
    const canonicalName = canonicalTechnicianName_(techName);
    const ticketStats = ticketStatsByTech[techName]
      || ticketStatsByTech[canonicalName]
      || ticketStatsByTech[normalizeTechnicianNameFull_(techName)]
      || { created: 0, closed: 0, open: 0, issues: {} };
    const callMetrics = { totalCalls: 0, inboundCalls: 0, outboundCalls: 0, talkSeconds: 0, callSeconds: 0 };
    const canonicalMetrics = callMetricsByCanonical[canonicalName] || callMetricsByCanonical[normalizeTechnicianNameFull_(techName)];
    if (canonicalMetrics) {
      callMetrics.totalCalls = canonicalMetrics.totalCalls || 0;
      callMetrics.inboundCalls = canonicalMetrics.inboundCalls || 0;
      callMetrics.outboundCalls = canonicalMetrics.outboundCalls || 0;
      callMetrics.talkSeconds = canonicalMetrics.talkSeconds || 0;
      callMetrics.callSeconds = canonicalMetrics.callSeconds || 0;
    }
    const applyCallMetrics = (rows) => {
      if (!Array.isArray(rows)) return;
      rows.forEach(row => {
        const metricName = String(row && row[0] ? row[0] : '').toLowerCase();
        const raw = row && row.length > 1 ? row[1] : 0;
        const num = typeof raw === 'number' ? raw : Number(raw);
        const value = isNaN(num) ? 0 : num;
        if (metricName.includes('total calls') && !metricName.includes('incoming') && !metricName.includes('outgoing')) {
          callMetrics.totalCalls = value;
        } else if (metricName.includes('total incoming')) {
          callMetrics.inboundCalls = value;
        } else if (metricName.includes('total outgoing')) {
          callMetrics.outboundCalls = value;
        } else if (metricName.includes('talking duration')) {
          callMetrics.talkSeconds = value;
        }
      });
    };
    
    // Create dashboard if there's session data OR if there's an extension (Digium data will be fetched)
    const hasSessionData = techRows && techRows.length > 0;
    const hasTicketData = (ticketStats.created || ticketStats.closed || ticketStats.open) > 0;
    const shouldCreateDashboard = hasSessionData || hasExtension || hasTicketData;
    
    // If no data and sheet exists, delete it
    if (!shouldCreateDashboard) {
      if (techSheet) {
        try {
          ss.deleteSheet(techSheet);
          Logger.log(`Deleted dashboard sheet for ${techName} (no session data and no extension)`);
        } catch (e) {
          Logger.log(`Warning: Could not delete dashboard sheet for ${techName}: ${e.toString()}`);
        }
      } else {
        Logger.log(`No data for ${techName} (no sessions and no extension), no sheet to delete`);
      }
      continue;
    }
    
    if (!hasSessionData) {
      Logger.log(`${techName} - No session data, but has extension ${extList.join(', ')}, creating dashboard with Digium data only`);
    }
      // Important: clearing values does not remove row/column groups; fully reset any prior groups
      // Only clear groups on THIS personal dashboard sheet, not any other sheets
      try {
        // Get the actual last row with data to avoid affecting entire sheet unnecessarily
        const lastDataRow = Math.max(techSheet.getLastRow(), 100); // Use at least 100 rows to clear any groups
        const lastDataCol = Math.max(techSheet.getLastColumn(), 10); // Use at least 10 cols
        // Shift depths negatively to zero out any existing groups (safe no-op if none)
        // Only affect this specific sheet's range
        techSheet.getRange(1, 1, lastDataRow, 1).shiftRowGroupDepth(-8);
        techSheet.getRange(1, 1, 1, lastDataCol).shiftColumnGroupDepth(-8);
      } catch (e) { 
        Logger.log(`Warning: Failed to clear row groups for ${techName}: ${e.toString()}`);
      }
      techSheet.getRange(1, 1).setValue(`👤 ${techName} - Personal Dashboard`);
      techSheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#9C27B0');
      techSheet.getRange(1, 1, 1, 5).merge();
      techSheet.getRange(2, 1).setValue('Time Frame:');
      techSheet.getRange(2, 2).setFormula('=Dashboard_Config!B3');
      techSheet.getRange(2, 4).setValue('Last Updated:');
      techSheet.getRange(2, 5).setValue(new Date().toLocaleString());
      
      // Normalize durations/pickups/work seconds to seconds for robust calculations
      // Add defensive checks for missing columns
      // Use empty arrays if no session data (will result in zeros)
      const durations = hasSessionData && durationIdx >= 0 ? techRows.map(r => parseDurationSeconds_(r[durationIdx] || 0)).filter(v => v > 0) : [];
      const pickups = hasSessionData && pickupIdx >= 0 ? techRows.map(r => parseDurationSeconds_(r[pickupIdx] || 0)).filter(v => v > 0) : [];
      const workSeconds = hasSessionData && workIdx >= 0 ? techRows.map(r => parseDurationSeconds_(r[workIdx] || 0)).filter(v => v > 0).reduce((a,b) => a + b, 0) : 0;
      const activeSecondsArr = hasSessionData && activeIdx >= 0 ? techRows.map(r => parseDurationSeconds_(r[activeIdx] || 0)).filter(v => v > 0) : [];
      const daysToCloseArr = (sfMetrics && sfMetrics.perCanonical && sfMetrics.perCanonical[canonicalTechnicianName_(techName)])
        ? (sfMetrics.perCanonical[canonicalTechnicianName_(techName)].daysToClose || [])
        : [];
      const avgDaysToClose = daysToCloseArr.length ? (daysToCloseArr.reduce((a,b)=>a+b,0) / daysToCloseArr.length) : 0;
      const totalCallsForTech = callMetrics.totalCalls || 0;
      const ticketsCreated = ticketStats.created || 0;
      const ticketsOpen = ticketStats.open || 0;
      const inboundCalls = callMetrics.inboundCalls || 0;
      const outboundCalls = callMetrics.outboundCalls || 0;
      const talkSeconds = callMetrics.talkSeconds || 0;
      
      // Log sample values to debug why values are 0
      if (hasSessionData && techRows.length > 0) {
        const sampleRow = techRows[0];
        Logger.log(`${techName} - Sample row values: Duration[${durationIdx}]=${sampleRow[durationIdx]}, Pickup[${pickupIdx}]=${sampleRow[pickupIdx]}, Work[${workIdx}]=${sampleRow[workIdx]}, Active[${activeIdx}]=${sampleRow[activeIdx]}`);
        Logger.log(`${techName} - Parsed values: durations=${durations.length}, pickups=${pickups.length}, workSeconds=${workSeconds}, activeSeconds=${activeSecondsArr.length}`);
      }
      
      if (durationIdx < 0) Logger.log(`WARNING: Duration column not found for ${techName}`);
      if (pickupIdx < 0) Logger.log(`WARNING: Pickup column not found for ${techName}`);
      if (workIdx < 0) Logger.log(`WARNING: Work Time column not found for ${techName}`);
      if (activeIdx < 0) Logger.log(`WARNING: Active Time column not found for ${techName}`);
      
      const completed = hasSessionData ? techRows.filter(r => r[statusIdx] === 'Ended').length : 0;
      const active = hasSessionData ? techRows.filter(r => r[statusIdx] === 'Active').length : 0;
      const avgDur = durations.length > 0 ? (durations.reduce((a,b) => a+b, 0) / durations.length / 60).toFixed(1) : '0';
      const avgPickup = pickups.length > 0 ? (pickups.reduce((a,b) => a+b, 0) / pickups.length / 60).toFixed(1) : '0';
      const slaHits = pickups.filter(p => p <= 60).length;
      const slaPct = pickups.length > 0 ? ((slaHits / pickups.length) * 100).toFixed(1) : '0';
      const days = Math.max(1, (new Date(endDate) - new Date(startDate)) / (1000*60*60*24));
      // Total Active Hours should reflect active time from sessions (not work time)
      const totalActiveSecondsFromSessions = activeSecondsArr.reduce((a,b)=>a+b, 0);
      const totalActiveSecondsPerf = 0;
      const totalActiveSeconds = totalActiveSecondsPerf > 0 ? totalActiveSecondsPerf : totalActiveSecondsFromSessions;
      const activeHoursTotal = (totalActiveSeconds / 3600).toFixed(1);
      // Count Nova Wave sessions for this technician - use Channel Name instead of Calling Card
      const novaWaveCount = hasSessionData ? techRows.filter(row => {
        if (channelNameIdx >= 0 && row[channelNameIdx]) {
          const channelName = String(row[channelNameIdx] || '').toLowerCase();
          if (channelName.includes('nova wave')) return true;
        }
        if (callingCardIdx >= 0 && row[callingCardIdx]) {
          const callingCard = String(row[callingCardIdx] || '').toLowerCase();
          if (callingCard.includes('nova wave')) return true;
        }
        return false;
      }).length : 0;
      
      // Total Login Time should come from performance summary data (Rescue API)
      // Fall back to sum of session durations if performance data not available
      let totalLoginTimeSeconds = 0;
      if (totalLoginTimeSeconds === 0 && hasSessionData && durationIdx >= 0) {
        // Fallback: calculate from session durations if performance data not available
        totalLoginTimeSeconds = techRows.map(r => parseDurationSeconds_(r[durationIdx] || 0)).reduce((a,b)=>a+b, 0);
      }
      
      const kpiRow = 4;
      const topIssues = Object.entries(ticketStats.issues || {})
        .sort((a, b) => Number(b[1]) - Number(a[1]))
        .slice(0, 4);
      // Build KPI values. Durations/pickups are stored here as spreadsheet time fractions (seconds/86400)
      const avgDurationSeconds = durations.length > 0 ? Math.round(durations.reduce((a,b)=>a+b,0) / durations.length) : 0;
      const avgPickupSeconds = pickups.length > 0 ? Math.round(pickups.reduce((a,b)=>a+b,0) / pickups.length) : 0;
      // Calculate values with proper defensive checks
      const avgActiveTimeSeconds = (totalActiveSeconds > 0 && ((0 || 0) || techRows.length))
        ? Math.round(totalActiveSeconds / Math.max(1, (0 || 0) || techRows.length || 1))
        : (activeSecondsArr.length > 0 ? Math.round(activeSecondsArr.reduce((a,b)=>a+b,0) / activeSecondsArr.length) : 0);
      const avgPickupTimeSeconds = avgPickupSeconds; // Already calculated above
      const slaHitPercent = pickups.length > 0 ? (slaHits / pickups.length) : 0; // Numeric 0-1 for percentage format
      const totalActiveHoursNum = totalActiveSeconds / 3600; // Numeric hours for decimal format
      
      // Log calculated values for debugging
      Logger.log(`${techName} KPIs - Sessions: ${techRows.length}, Pickups: ${pickups.length}, Active Seconds: ${totalActiveSeconds}, Avg Pickup: ${avgPickupTimeSeconds}s`);
      
      // Get total sessions from performance data if available, otherwise fall back to techRows.length
      const totalSessionsFromPerfKPI = 0;
      const totalSessions = totalSessionsFromPerfKPI > 0 ? totalSessionsFromPerfKPI : techRows.length;
      const sessionsPerHourFromPerf = null;
      const sessionsPerHour = sessionsPerHourFromPerf != null && sessionsPerHourFromPerf > 0
        ? sessionsPerHourFromPerf
        : (hasSessionData ? (techRows.length / days / 8) : 0);
      const ticketOpenRatio = ticketsCreated > 0 ? ((totalCallsForTech || 0) + totalSessions) / ticketsCreated : 0;
      
      // Get Digium data once per technician (by_account breakdown)
      const kpis = [
        ['Total Sessions', totalSessions],
        ['Nova Wave Sessions', novaWaveCount],
        ['Total Tickets', ticketsCreated],
        ['Unresolved Tickets', ticketStats.open],
        ['Tickets Closed', ticketStats.closed],
        ['Total Calls', totalCallsForTech],
        ['Incoming Calls', inboundCalls],
        ['Outgoing Calls', outboundCalls],
        ['Total Talk Time', talkSeconds > 0 ? talkSeconds / 86400 : 0],
        ['Days to Close (Avg)', avgDaysToClose],
        ['Active Time', totalActiveSeconds > 0 ? totalActiveSeconds / 86400 : 0],
        ['Avg Pickup Time', avgPickupTimeSeconds > 0 ? avgPickupTimeSeconds / 86400 : 0],
        ['SLA Hit %', slaHitPercent],
        ['Total Active Hours', totalActiveHoursNum],
        ['Total Login Time', totalLoginTimeSeconds > 0 ? totalLoginTimeSeconds / 86400 : 0],
        ['Sessions/Hour', sessionsPerHour],
        ['Ticket Open Rate', ticketOpenRatio]
      ];
      const issuesCol = 7;
      techSheet.getRange(kpiRow, issuesCol)
        .setValue('Top Technical Issues')
        .setFontSize(12)
        .setFontWeight('bold')
        .setFontColor('#FFFFFF')
        .setBackground('#6A1B9A')
        .setHorizontalAlignment('center');
      for (let i = 0; i < kpis.length; i++) {
        const row = kpiRow + Math.floor(i / 2);
        const col = (i % 2) * 3 + 1;
        techSheet.getRange(row, col).setValue(kpis[i][0]);
        techSheet.getRange(row, col).setFontSize(11).setFontColor('#666666');
        techSheet.getRange(row, col + 1).setValue(kpis[i][1]);
        techSheet.getRange(row, col + 1).setFontSize(16).setFontWeight('bold').setFontColor('#9C27B0');
        // Apply appropriate number formats
        try {
          const label = String(kpis[i][0] || '');
          if (/avg pickup/i.test(label) || /active time/i.test(label) || /login time/i.test(label) || /total talk time/i.test(label)) {
            techSheet.getRange(row, col + 1).setNumberFormat('hh:mm:ss');
          } else if (/sla hit/i.test(label)) {
            // SLA Hit % as percentage format (0.0%)
            techSheet.getRange(row, col + 1).setNumberFormat('0.0%');
          } else if (/total active hours/i.test(label)) {
            // Total Active Hours as decimal hours (0.0 hrs)
            techSheet.getRange(row, col + 1).setNumberFormat('0.0');
          } else if (/sessions\/hour/i.test(label)) {
            // Sessions/Hour as single decimal
            techSheet.getRange(row, col + 1).setNumberFormat('0.0');
          } else if (/ticket open rate/i.test(label)) {
            techSheet.getRange(row, col + 1).setNumberFormat('0.0');
          } else if (/incoming calls/i.test(label) || /outgoing calls/i.test(label) || /total calls/i.test(label) || /total sessions/i.test(label) || /nova wave/i.test(label) || /tickets/i.test(label)) {
            // Total Calls, Total Sessions, Nova Wave Sessions as integers
            techSheet.getRange(row, col + 1).setNumberFormat('0');
          }
        } catch (e) { /* ignore formatting errors */ }
        techSheet.getRange(row, col, 1, 2).setBorder(true, true, true, true, true, true);
      }

      // Top technical issues block (columns 7-8)
      const issuesWidth = 2;
      const issuesRows = 6; // header + column headings + up to 4 entries
      try { techSheet.getRange(kpiRow, issuesCol, issuesRows, issuesWidth).clear(); } catch (e) {}
      techSheet.getRange(kpiRow, issuesCol, 1, issuesWidth).merge();
      techSheet.getRange(kpiRow + 1, issuesCol, 1, issuesWidth)
        .setValues([['Issue', 'Count']])
        .setFontWeight('bold')
        .setBackground('#1E3A8A')
        .setFontColor('#FFFFFF');
      if (topIssues.length) {
        const issueRows = topIssues.map(([label, count]) => [label, count]);
        techSheet.getRange(kpiRow + 2, issuesCol, issueRows.length, issuesWidth).setValues(issueRows);
        techSheet.getRange(kpiRow + 2, issuesCol + 1, issueRows.length, 1).setNumberFormat('0');
      } else {
        techSheet.getRange(kpiRow + 2, issuesCol, 1, issuesWidth)
          .setValue('No ticket issues found')
          .setFontStyle('italic')
          .setFontColor('#6B7280');
      }
      try {
        techSheet.getRange(kpiRow + 1, issuesCol, issuesRows - 1, issuesWidth)
          .setBorder(true, true, true, true, true, true);
      } catch (e) { /* ignore */ }

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
    // Always include dates in header, even for single day (so date appears above data)
    const headerRow = ['Metric'];
    dates.forEach(d => {
      const dObj = new Date(d + 'T00:00:00');
      headerRow.push(`${dObj.getMonth()+1}/${dObj.getDate()}/${dObj.getFullYear()}`);
    });
    headerRow.push('Totals/Averages');
    const totalSessionsRow = ['Total sessions'];
    const totalActiveRow = ['Total Active Time'];
    const totalWorkRow = ['Total Work Time'];
    const totalLoginRow = ['Total Login Time'];
    const avgSessionRow = ['Avg Session (API)'];
    const avgPickupRow = ['Avg Pick-up Speed'];
    const longestSessionRow = ['Longest Session Time'];
    const sessionsPerHourRow = ['Sessions Per Hour'];
    // Get total sessions from performance data for this technician
    const perfDataForSummary = {};
    const totalSessionsFromPerfSummary = perfDataForSummary.totalSessions || 0;
    const totalActiveFromPerfSummary = perfDataForSummary.totalActiveTime || 0;
    const totalWorkFromPerfSummary = perfDataForSummary.totalWorkTime || 0;
    const totalLoginFromPerfSummary = perfDataForSummary.totalLoginTime || 0;
    const sessionsPerHourFromPerfSummary = perfDataForSummary.sessionsPerHour || null;
    let totalSessionsAll = 0;
    let totalActiveAll = 0;
    let totalWorkAll = 0;
    let totalLoginAll = 0;
      let totalAvgSeconds = 0;
      let daysWithData = 0;
      let totalPickupSeconds = 0;
      let totalPickupCount = 0;
      let longestSessionAll = 0;
      let totalSessionsPerHour = 0;
      const dailyMap = {};
              dates.forEach(d => { dailyMap[d] = { sessions: [], totalActiveSeconds: 0, totalWorkSeconds: 0, longestSessionSeconds: 0 }; });
              if (hasSessionData) {
                techRows.forEach(r => {
                  try {
                    const d = new Date(r[startIdx]).toISOString().split('T')[0];
                    if (!dailyMap[d]) dailyMap[d] = { sessions: [], totalActiveSeconds: 0, totalWorkSeconds: 0, longestSessionSeconds: 0 };
                    dailyMap[d].sessions.push(r);
                    // Sum ACTIVE time for daily totals (with defensive check)
                    if (activeIdx >= 0) dailyMap[d].totalActiveSeconds += parseDurationSeconds_(r[activeIdx] || 0);
                    // Sum WORK time for daily totals (if present)
                    if (workIdx >= 0) dailyMap[d].totalWorkSeconds += parseDurationSeconds_(r[workIdx] || 0);
                    // Track longest session per day
                    if (durationIdx >= 0) {
                      const sessionDur = parseDurationSeconds_(r[durationIdx] || 0);
                      if (sessionDur > dailyMap[d].longestSessionSeconds) {
                        dailyMap[d].longestSessionSeconds = sessionDur;
                      }
                    }
                  } catch (e) { }
                });
              }
      dates.forEach(d => {
        const data = dailyMap[d];
        // For each day, use the count from sessions data
        // The total at the end will use performance data if available
        const count = data && data.sessions ? data.sessions.length : 0;
        totalSessionsRow.push(count);
        totalSessionsAll += count;
        const activeSecs = data && data.totalActiveSeconds ? data.totalActiveSeconds : 0;
        totalActiveRow.push(activeSecs / 86400);
        totalActiveAll += activeSecs;
  const workSecs = data && data.totalWorkSeconds ? data.totalWorkSeconds : 0;
  totalWorkRow.push(workSecs / 86400);
  totalWorkAll += workSecs;
  // For Login Time, use performance summary data if available, otherwise use session durations
  const perfDataForDayLogin = {};
  const loginTimeForDay = perfDataForDayLogin.totalLoginTime || 0;
  if (loginTimeForDay > 0) {
    // Use performance summary login time (divide by number of days if it's a total)
    const daysInRange = dates.length;
    const dailyLoginTime = loginTimeForDay / daysInRange;
    totalLoginRow.push(dailyLoginTime / 86400);
    totalLoginAll += dailyLoginTime;
  } else if (data && data.sessions.length > 0 && durationIdx >= 0) {
    // Fallback: use sum of session durations for this day
    const dayDurations = data.sessions.map(s => parseDurationSeconds_(s[durationIdx] || 0)).reduce((a,b)=>a+b, 0);
    totalLoginRow.push(dayDurations / 86400);
    totalLoginAll += dayDurations;
  } else {
    totalLoginRow.push(0);
  }
        // For Avg Session per day, use API performance data if available, otherwise calculate from sessions
        const perfDataForDay = {}; // legacy summary removed; calculate from sessions only
        const apiAvgDuration = perfDataForDay.avgDuration || 0;
        
        if (data && data.sessions.length > 0 && durationIdx >= 0) {
          const durations = data.sessions.map(s => parseDurationSeconds_(s[durationIdx] || 0)).filter(Boolean);
          if (durations.length > 0) {
            // Use API average if available, otherwise calculate from session data
            const avg = apiAvgDuration > 0 ? apiAvgDuration : (durations.reduce((a,b)=>a+b,0) / durations.length);
            avgSessionRow.push(avg / 86400);
            totalAvgSeconds += avg;
            daysWithData++;
          } else if (apiAvgDuration > 0) {
            // No sessions for this day, but we have API average - use it
            avgSessionRow.push(apiAvgDuration / 86400);
            totalAvgSeconds += apiAvgDuration;
            daysWithData++;
          } else {
            avgSessionRow.push(0);
          }
        } else if (apiAvgDuration > 0) {
          // No sessions for this day, but we have API average - use it
          avgSessionRow.push(apiAvgDuration / 86400);
          totalAvgSeconds += apiAvgDuration;
          daysWithData++;
        } else {
          avgSessionRow.push(0);
        }

        // For Avg Pick-up Speed per day, calculate from sessions data for that day
        if (data && data.sessions.length > 0 && pickupIdx >= 0) {
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

        // Longest Session Time per day
        if (data && data.longestSessionSeconds) {
          longestSessionRow.push(data.longestSessionSeconds / 86400);
          if (data.longestSessionSeconds > longestSessionAll) {
            longestSessionAll = data.longestSessionSeconds;
          }
        } else {
          longestSessionRow.push(0);
        }

        // Sessions Per Hour per day (assuming 8 hour work day)
        const sessionsPerHour = count > 0 ? (count / 8).toFixed(1) : 0;
        sessionsPerHourRow.push(Number(sessionsPerHour));
        totalSessionsPerHour += Number(sessionsPerHour);
      });
  totalSessionsRow.push(totalSessionsAll);
  totalActiveRow.push(totalActiveAll / 86400);
  totalWorkRow.push(totalWorkAll / 86400);
    totalLoginRow.push(totalLoginAll / 86400);
      // For "Totals/Averages" column, use API performance data if available, otherwise calculate from session data
      const apiAvgDuration = 0; // legacy summary removed; calculate from sessions only
      const apiAvgPickup = 0; // legacy summary removed; calculate from sessions only
      
      if (apiAvgDuration > 0) {
        // Use API average duration if available
        avgSessionRow.push(apiAvgDuration / 86400);
      } else {
        // Fall back to calculated average from session data
        avgSessionRow.push(daysWithData>0 ? (totalAvgSeconds / daysWithData) / 86400 : 0);
      }
      
      if (apiAvgPickup > 0) {
        // Use API average pickup if available
        avgPickupRow.push(apiAvgPickup / 86400);
      } else {
        // Fall back to calculated average from session data
        avgPickupRow.push(totalPickupCount>0 ? (totalPickupSeconds / totalPickupCount) / 86400 : 0);
      }
  longestSessionRow.push(longestSessionAll / 86400);
  const sessionsPerHourTotal = totalSessionsAll > 0 ? (totalSessionsAll / (dates.length * 8)) : 0;
  sessionsPerHourRow.push(Number(sessionsPerHourTotal.toFixed(1)));
  const wideRows = [headerRow, totalSessionsRow, totalActiveRow, totalLoginRow, totalWorkRow, avgSessionRow, avgPickupRow, longestSessionRow, sessionsPerHourRow];
      techSheet.getRange(summaryStart + 1, 1, wideRows.length, wideRows[0].length).setValues(wideRows);
      // style header
      try { techSheet.getRange(summaryStart + 1, 1, 1, wideRows[0].length).setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF'); } catch(e){}
      // Format per-row: Total sessions should be an integer, Sessions Per Hour is decimal, time rows are hh:mm:ss
      try {
        const valueCols = wideRows[0].length - 1; // excluding 'Metric' column
        if (valueCols > 0) {
          // Total sessions row -> integer (row index 1 after header)
          techSheet.getRange(summaryStart + 2, 2, 1, valueCols).setNumberFormat('0');
          // Time rows -> hh:mm:ss (Active, Login, Work, Avg Session, Avg Pickup, Longest Session) - rows 2-7
          techSheet.getRange(summaryStart + 3, 2, 6, valueCols).setNumberFormat('hh:mm:ss');
          // Sessions Per Hour row -> one decimal (row index 8 after header)
          techSheet.getRange(summaryStart + 9, 2, 1, valueCols).setNumberFormat('0.0');
        }
      } catch (e) { /* ignore formatting errors */ }
  // --- Calls Summary (Digium data) ---
      // Use aggregated call metrics that were computed earlier for this technician
      let callsSummaryRow = summaryStart + 1 + wideRows.length + 2;
      try {
        const totalCallsValue = Number(totalCallsForTech || 0);
        const inboundValue = Number(inboundCalls || 0);
        const outboundValue = Number(outboundCalls || 0);
        const talkSecondsValue = Number(callMetrics.talkSeconds || 0);
        const callSecondsValue = Number(callMetrics.callSeconds || 0);
        const avgTalkSecondsValue = totalCallsValue > 0 ? talkSecondsValue / totalCallsValue : 0;
        const avgCallSecondsValue = totalCallsValue > 0 ? callSecondsValue / totalCallsValue : 0;
        
            techSheet.getRange(callsSummaryRow, 1).setValue('Calls Summary');
            techSheet.getRange(callsSummaryRow, 1).setFontSize(14).setFontWeight('bold');
            
            const callsHeaderRow = ['Metric', 'Total'];
            techSheet.getRange(callsSummaryRow + 1, 1, 1, callsHeaderRow.length).setValues([callsHeaderRow]);
            techSheet.getRange(callsSummaryRow + 1, 1, 1, callsHeaderRow.length).setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF');
            
        const callRows = [
          ['Total Calls', totalCallsValue],
          ['Total Incoming Calls', inboundValue],
          ['Total Outgoing Calls', outboundValue],
          ['Talking Duration', talkSecondsValue / 86400],
          ['Call Duration', callSecondsValue / 86400],
          ['Avg Talking Duration', avgTalkSecondsValue / 86400],
          ['Avg Call Duration', avgCallSecondsValue / 86400]
        ];
        
        techSheet.getRange(callsSummaryRow + 2, 1, callRows.length, 2).setValues(callRows);
        for (let i = 0; i < callRows.length; i++) {
          const label = String(callRows[i][0] || '').toLowerCase();
          const cell = techSheet.getRange(callsSummaryRow + 2 + i, 2);
          if (label.includes('duration') || label.includes('avg')) {
            cell.setNumberFormat('hh:mm:ss');
                } else {
            cell.setNumberFormat('0');
          }
        }
        callsSummaryRow += 2 + callRows.length + 2;
                  } catch (e) { 
        Logger.log(`${techName} - Failed to build Calls Summary: ${e.toString()}`);
          callsSummaryRow += 1;
    }
      
  // --- Session Details (placed after Calls Summary) ---
      const detailRow = callsSummaryRow;
      techSheet.getRange(detailRow, 1).setValue('Session Details');
      techSheet.getRange(detailRow, 1).setFontSize(14).setFontWeight('bold');
      // Column order: Date, Session ID, Customer Name, Phone Number, Location Name, Duration, Pickup
      const detailHeaders = ['Date', 'Session ID', 'Customer Name', 'Phone Number', 'Location Name', 'Duration', 'Pickup'];
      techSheet.getRange(detailRow + 1, 1, 1, detailHeaders.length).setValues([detailHeaders]);
      techSheet.getRange(detailRow + 1, 1, 1, detailHeaders.length).setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF');
      
      // Resolve indices that might exist for fallback lookups
      const resolvedIdx = headers.indexOf('resolved_unresolved');
  const callerNameIdx = getHeaderIndex(['Your Name:','caller_name', 'caller name', 'caller', 'Caller Name', 'Customer Name']);
  const locationNameIdx = getHeaderIndex(['Location Name', 'location_name', 'location name', 'Location', 'location', 'Company Name', 'company_name']);

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

      const detailRows = hasSessionData ? techRows.slice(0, 50).map(row => {
        const date = row[startIdx] ? new Date(row[startIdx]).toISOString().split('T')[0] : '';
        const customerName = pickCustomerFromRow(row);
        // phone: prefer explicit phone column, otherwise try company or caller columns
        let phoneNumber = '';
        if (phoneIdx >= 0) phoneNumber = row[phoneIdx] || '';
        if ((!phoneNumber || !looksLikePhone(phoneNumber)) && companyIdx >= 0) phoneNumber = row[companyIdx] || phoneNumber || '';
        if ((!phoneNumber || !looksLikePhone(phoneNumber)) && callerNameIdx >= 0) phoneNumber = row[callerNameIdx] || phoneNumber || '';
        
        // Location Name from session data
        const locationName = (locationNameIdx >= 0 && row[locationNameIdx]) ? String(row[locationNameIdx]).trim() : '';

        // Duration and pickup as true time values (fraction of day)
        const durTime = (durationIdx >= 0 && row[durationIdx]) ? secToTimeValue(parseDurationSeconds_(row[durationIdx])) : 0;
        const pickupTime = (pickupIdx >= 0 && row[pickupIdx]) ? secToTimeValue(parseDurationSeconds_(row[pickupIdx])) : 0;
        // session id fallback: if session id looks like a timestamp, try tracking_id or other headers
        let sessionId = row[sessionIdIdx] || '';
        const looksLikeDateTime = (s) => { if (!s) return false; const t=String(s); return /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(t) || /\d{4}-\d{2}-\d{2}/.test(t) || /\d{1,2}:\d{2}:\d{2}/.test(t); };
        if (looksLikeDateTime(sessionId)) {
          // try tracking id
          const trackIdx = headers.findIndex(h => /tracking[_ ]?id/i.test(String(h||'')));
          if (trackIdx >= 0 && row[trackIdx]) sessionId = String(row[trackIdx]);
          else sessionId = '';
        }

        return [ date, sessionId, customerName, phoneNumber, locationName, durTime, pickupTime ];
      }) : [];
      // If no details, still draw headers (no rows)
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
        // IMPORTANT: Only apply grouping to THIS personal dashboard sheet's Session Details section
        // Never apply grouping to Sessions sheet or any other sheet
        try {
          // Ensure any old grouping around this region is cleared so header/title do not collapse
          // Use a very specific, limited range to avoid affecting other parts
          const clearStart = Math.max(1, detailRow);
          const clearCount = Math.min(detailRows.length + 3, 100); // Limit to reasonable size (max 100 rows)
          // Only clear groups in this specific range on THIS sheet
          techSheet.getRange(clearStart, 1, clearCount, 1).shiftRowGroupDepth(-8);
          
          // Create a fresh group starting at the first data row only (detailRow + 2)
          // This is the Session Details section - ONLY group these specific rows on THIS sheet
          if (detailRows.length > 0) {
            const groupStartRow = detailRow + 2;
            const groupEndRow = groupStartRow + detailRows.length - 1;
            // Verify we're only grouping the Session Details data rows
            const groupRange = techSheet.getRange(groupStartRow, 1, detailRows.length, 1);
            groupRange.shiftRowGroupDepth(1);
            try { 
              const grp = techSheet.getRowGroup(groupStartRow, 1); 
              if (grp) {
                grp.collapse();
                Logger.log(`${techName} - Collapsed Session Details rows ${groupStartRow}-${groupEndRow} on personal dashboard only`);
              }
            } catch (e) {
              Logger.log(`Warning: Could not collapse row group for ${techName} Session Details: ${e.toString()}`);
            }
          }
        } catch (e) { 
          Logger.log(`Collapsible session details grouping failed for ${techName}: ${e.toString()}`);
        }
      }
      techSheet.setColumnWidth(1, 100);
      techSheet.setColumnWidth(2, 150);
      techSheet.setColumnWidth(3, 150);
      techSheet.setColumnWidth(4, 120);
      techSheet.setColumnWidth(5, 150); // Location Name
      techSheet.setColumnWidth(6, 120); // Duration
      techSheet.setColumnWidth(7, 120); // Pickup

      // Note: Digium calls data is now placed in the Calls Summary section above (after per-day performance summary)

      
    }
    try {
      const allowedNormalized = new Set(Array.from(allowedSafeNames));
      const sheetsToRemove = ss.getSheets().filter(sh => {
        const nameNorm = normalizeSheetName(sh.getName());
        if (allowedNormalized.has(nameNorm)) return false;
        try {
          const title = String(sh.getRange(1, 1).getValue() || '');
          return /personal dashboard/i.test(title);
        } catch (e) {
          return false;
        }
      });
      sheetsToRemove.forEach(sh => {
        Logger.log(`Removing personal dashboard sheet not mapped to extension: ${sh.getName()}`);
        ss.deleteSheet(sh);
      });
    } catch (e) {
      Logger.log('generateTechnicianTabs_: cleanup of disallowed dashboards failed: ' + e.toString());
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
    // Set TEXT output (XML not working reliably)
    try {
      const rt = apiGet_(cfg.rescueBase, 'setOutput.aspx', { output: 'TEXT' }, cookie, 2, true);
      const tt = (rt.getContentText() || '').trim();
      if (!/^OK/i.test(tt)) Logger.log(`setOutput TEXT warning: ${tt}`);
    } catch (e) { Logger.log('setOutput TEXT failed (non-fatal): ' + e.toString()); }
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

var ANALYTICS_THEME = {
  background: '#F8FAFC',
  header: '#E0F2FE',
  subheader: '#F1F5F9',
  tableHeader: '#D7E3FC',
  rowEven: '#FFFFFF',
  rowOdd: '#EDF2FB',
  accent: '#1D4ED8',
  text: '#1F2937',
  muted: '#64748B',
  border: '#CBD5F5'
};

var ANALYTICS_GOOD_COLOR = '#16A34A';
var ANALYTICS_WARN_COLOR = '#F59E0B';
var ANALYTICS_BAD_COLOR = '#EF4444';

function styleAnalyticsSectionHeader_(sheet, row, colSpan) {
  try {
    const span = Math.max(colSpan || 6, 1);
    sheet.getRange(row, 1, 1, span)
      .setBackground(ANALYTICS_THEME.subheader)
      .setFontColor(ANALYTICS_THEME.accent)
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, ANALYTICS_THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) { Logger.log('styleAnalyticsSectionHeader_ error: ' + e.toString()); }
}

function registerAnalyticsTable_(registry, headerRow, startCol, numCols, dataRows) {
  if (registry) {
    registry.push({
      type: 'table',
      headerRow: headerRow,
      startCol: startCol || 1,
      numCols: numCols || 1,
      dataRows: Math.max(0, dataRows || 0)
    });
  }
}

function registerAnalyticsHighlight_(registry, entry) {
  if (registry && entry) registry.push(entry);
}

function applyAnalyticsTableTheme_(sheet, entry) {
  try {
    const totalRows = Math.max(1, (entry.dataRows || 0) + 1);
    const range = sheet.getRange(entry.headerRow, entry.startCol, totalRows, entry.numCols);
    const banding = range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    banding.setHeaderRowColor(ANALYTICS_THEME.tableHeader)
      .setFirstRowColor(ANALYTICS_THEME.rowEven)
      .setSecondRowColor(ANALYTICS_THEME.rowOdd)
      .setFooterRowColor(null);
    sheet.getRange(entry.headerRow, entry.startCol, 1, entry.numCols)
      .setFontColor(ANALYTICS_THEME.accent)
      .setFontWeight('bold')
      .setFontSize(11);
    if (entry.dataRows > 0) {
      sheet.getRange(entry.headerRow + 1, entry.startCol, entry.dataRows, entry.numCols)
        .setFontColor(ANALYTICS_THEME.text);
    }
    sheet.getRange(entry.headerRow, entry.startCol, totalRows, entry.numCols)
      .setFontFamily('Roboto')
      .setBorder(true, true, true, true, false, false, ANALYTICS_THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) {
    Logger.log('applyAnalyticsTableTheme_ error: ' + e.toString());
  }
}
function applyAnalyticsHighlight_(sheet, entry) {
  try {
    if (entry.type === 'thresholdColor') {
      if (!entry.rows || entry.rows <= 0) return;
      const range = sheet.getRange(entry.startRow, entry.column, entry.rows, 1);
      const values = range.getValues();
      for (let i = 0; i < values.length; i++) {
        let raw = values[i][0];
        if (raw == null || raw === '') continue;
        raw = String(raw).replace('%', '');
        const numeric = parseFloat(raw);
        if (isNaN(numeric)) continue;
        let color = ANALYTICS_BAD_COLOR;
        if (numeric >= entry.thresholds.high) {
          color = ANALYTICS_GOOD_COLOR;
        } else if (numeric >= entry.thresholds.mid) {
          color = ANALYTICS_WARN_COLOR;
        }
        const cell = range.getCell(i + 1, 1);
        cell.setBackground(color).setFontColor('#0B1120').setFontWeight('bold');
      }
    } else if (entry.type === 'trendChange') {
      if (!entry.rows || entry.rows <= 0) return;
      const metricRange = sheet.getRange(entry.startRow, entry.metricCol, entry.rows, 1).getValues();
      const changeRange = sheet.getRange(entry.startRow, entry.changeCol, entry.rows, 1);
      const changeValues = changeRange.getValues();
      for (let i = 0; i < changeValues.length; i++) {
        const metricName = String(metricRange[i][0] || '').toLowerCase();
        let raw = changeValues[i][0];
        if (raw == null || raw === '') continue;
        raw = String(raw).replace('%', '');
        const change = parseFloat(raw);
        if (isNaN(change)) continue;
        let goodWhenPositive = true;
        if (metricName.includes('avg duration') || metricName.includes('avg pickup')) {
          goodWhenPositive = false;
        }
        let color = ANALYTICS_WARN_COLOR;
        if (change === 0) {
          color = ANALYTICS_WARN_COLOR;
        } else if ((change > 0 && goodWhenPositive) || (change < 0 && !goodWhenPositive)) {
          color = ANALYTICS_GOOD_COLOR;
        } else {
          color = ANALYTICS_BAD_COLOR;
        }
        const cell = changeRange.getCell(i + 1, 1);
        cell.setBackground(color).setFontColor('#0B1120').setFontWeight('bold');
      }
    }
  } catch (e) {
    Logger.log('applyAnalyticsHighlight_ error: ' + e.toString());
  }
}
function analyticsGetHeaderIndex_(headers, variants) {
  if (!headers || !headers.length) return -1;
  const normalized = headers.map(h => String(h || '').trim().toLowerCase());
  for (const variant of variants || []) {
    const target = String(variant || '').trim().toLowerCase();
    if (!target) continue;
    const idx = normalized.findIndex(h => h === target);
    if (idx >= 0) return idx;
  }
  // fallback: allow partial match
  for (const variant of variants || []) {
    const target = String(variant || '').trim().toLowerCase();
    if (!target) continue;
    const idx = normalized.findIndex(h => h.includes(target));
    if (idx >= 0) return idx;
  }
  return -1;
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
function refreshAdvancedAnalyticsDashboard_(startDate, endDate, digiumDatasetOpt, extensionMetaOpt) {
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
    
    try { analyticsSheet.getBandings().forEach(b => b.remove()); } catch (e) { Logger.log('Advanced analytics banding clear failed: ' + e.toString()); }
    try { analyticsSheet.setConditionalFormatRules([]); } catch (e) { Logger.log('Advanced analytics conditional formats clear failed: ' + e.toString()); }
    
    // Update last updated timestamp
    analyticsSheet.getRange(2, 5).setValue(new Date().toLocaleString());
    
    // Clear existing content (except header rows 1-2) to ensure fresh data
    const lastRow = analyticsSheet.getLastRow();
    const lastCol = analyticsSheet.getLastColumn();
    if (lastRow > 2 && lastCol > 0) {
      analyticsSheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
    }
    
    const extensionMeta = extensionMetaOpt || getActiveExtensionMetadata_();
    const digiumDataset = digiumDatasetOpt || getDigiumDataset_(startDate, endDate, extensionMeta);
    const callAnalytics = {
      extensionMeta,
      byHour: digiumDataset && digiumDataset.byHour ? digiumDataset.byHour : null,
      byDow: digiumDataset && digiumDataset.byDow ? digiumDataset.byDow : null,
      byDate: digiumDataset && digiumDataset.byDay ? digiumDataset.byDay : null,
      byAccount: digiumDataset && digiumDataset.byAccount ? digiumDataset.byAccount : null
    };
    
    const styleRegistry = [];
    
    // Generate all analytics sections
    let currentRow = 4;
    
    // 1. Peak Hours & Day of Week Analysis
    currentRow = createPeakHoursAnalysis_(analyticsSheet, currentRow, startDate, endDate, styleRegistry, callAnalytics);
    currentRow += 5;
    
    // 2. Technician Effectiveness Comparison
    currentRow = createTechnicianEffectiveness_(analyticsSheet, currentRow, startDate, endDate, styleRegistry, callAnalytics);
    currentRow += 5;
    
    // 3. Repeat Customer Analysis
    currentRow = createRepeatCustomerAnalysis_(analyticsSheet, currentRow, startDate, endDate, styleRegistry);
    currentRow += 5;
    
    // 4. Trend Analysis
    currentRow = createTrendAnalysis_(analyticsSheet, currentRow, startDate, endDate, styleRegistry, callAnalytics);
    currentRow += 5;
    
    // 5. Technician Utilization Rate
    currentRow = createUtilizationRate_(analyticsSheet, currentRow, startDate, endDate, styleRegistry);
    currentRow += 5;
    
    // 6. Time to Resolution Distribution
    currentRow = createResolutionDistribution_(analyticsSheet, currentRow, startDate, endDate, styleRegistry);
    currentRow += 5;
    
    // 7. Real-time Capacity Indicators
    currentRow = createCapacityIndicators_(analyticsSheet, currentRow, styleRegistry);
    currentRow += 5;
    
    // 8. Predictive Analytics
    currentRow = createPredictiveAnalytics_(analyticsSheet, currentRow, styleRegistry, startDate, endDate, callAnalytics);
    // Re-apply visual polishing for readability after all sections are written
    try { applyAdvancedAnalyticsStyling_(analyticsSheet, styleRegistry); } catch (e) { Logger.log('applyAdvancedAnalyticsStyling_ failed: ' + e.toString()); }
    
  } catch (e) {
    Logger.log('refreshAdvancedAnalyticsDashboard_ error: ' + e.toString());
  }
}
function parseHourLabelToNumber_(label) {
  if (label == null) return null;
  const raw = String(label).trim();
  if (!raw) return null;
  if (/^\d+$/.test(raw)) {
    const num = Number(raw);
    return (num >= 0 && num < 24) ? num : null;
  }
  const amPmMatch = raw.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/i);
  if (amPmMatch) {
    let hour = Number(amPmMatch[1]) % 12;
    if (/PM/i.test(amPmMatch[3])) hour += 12;
    return (hour >= 0 && hour < 24) ? hour : null;
  }
  const hhmmMatch = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (hhmmMatch) {
    const hour = Number(hhmmMatch[1]);
    return (hour >= 0 && hour < 24) ? hour : null;
  }
  return null;
}
// 1. Peak Hours & Day of Week Analysis
function createPeakHoursAnalysis_(sheet, startRow, startDate, endDate, styleRegistry, callAnalytics) {
  try {
    styleRegistry = styleRegistry || [];
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allDataRaw = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const allData = filterOutExcludedTechnicians_(headers, allDataRaw);
    const startIdx = analyticsGetHeaderIndex_(headers, ['start_time', 'start time', 'Start Time']);
    const techIdx = analyticsGetHeaderIndex_(headers, ['technician name', 'technician_name', 'technician', 'tech']);
    if (startIdx < 0) {
      Logger.log('createPeakHoursAnalysis_: start_time header not found');
      return startRow;
    }
    
    const startMillis = startDate.getTime();
    const endMillis = endDate.getTime();
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx]) return false;
      try {
        const rowObj = row[startIdx] instanceof Date ? row[startIdx] : new Date(row[startIdx]);
        if (!(rowObj instanceof Date) || isNaN(rowObj)) return false;
        const rowMillis = rowObj.getTime();
        if (rowMillis < startMillis || rowMillis > endMillis) return false;
        if (techIdx >= 0 && isExcludedTechnician_(row[techIdx])) return false;
        return true;
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
    
    // Call analytics
    const callHourCounts = Array(24).fill(0);
    const callDayCounts = [0, 0, 0, 0, 0, 0, 0];
    if (callAnalytics && callAnalytics.byHour && callAnalytics.byHour.ok && Array.isArray(callAnalytics.byHour.rows)) {
      const categories = callAnalytics.byHour.categories || [];
      const callRow = callAnalytics.byHour.rows.find(r => /total calls/i.test(String(r[0] || '')));
      if (callRow) {
        categories.forEach((cat, idx) => {
          const hour = parseHourLabelToNumber_(cat);
          if (hour != null && hour >= 0 && hour < 24) {
            const value = callRow.length > idx + 1 ? Number(callRow[idx + 1]) || 0 : 0;
            callHourCounts[hour] = value;
          }
        });
      }
    }
    if (callAnalytics && callAnalytics.byDow && callAnalytics.byDow.ok && Array.isArray(callAnalytics.byDow.rows)) {
      const categories = callAnalytics.byDow.categories || [];
      const callRow = callAnalytics.byDow.rows.find(r => /total calls/i.test(String(r[0] || '')));
      if (callRow) {
        categories.forEach((cat, idx) => {
          const normalized = String(cat || '').trim().toLowerCase();
          const dayMap = { 'sunday':0,'monday':1,'tuesday':2,'wednesday':3,'thursday':4,'friday':5,'saturday':6 };
          let dayIndex = dayMap.hasOwnProperty(normalized) ? dayMap[normalized] : Number(normalized);
          if (!isNaN(dayIndex) && dayIndex >= 0 && dayIndex < 7) {
            callDayCounts[dayIndex] = Number(callRow[idx + 1]) || 0;
          }
        });
      }
    }
    
    const totalSessions = filtered.length;
    const totalCalls = callHourCounts.reduce((a, b) => a + b, 0);
    
    const sessionHourData = hourCounts.map((count, hour) => {
      const pct = totalSessions > 0 ? ((count / totalSessions) * 100).toFixed(1) : '0';
      return [hour + ':00', count, pct + '%'];
    }).sort((a, b) => b[1] - a[1]).slice(0, 10);
    
    const callHourData = callHourCounts.map((count, hour) => {
      const pct = totalCalls > 0 ? ((count / totalCalls) * 100).toFixed(1) : '0';
      return [hour + ':00', count, pct + '%'];
    }).sort((a, b) => b[1] - a[1]).slice(0, 10);
    
    const combinedHourData = hourCounts.map((sessionCount, hour) => {
      const callCount = callHourCounts[hour] || 0;
      const combined = sessionCount + callCount;
      return {
        hour,
        sessions: sessionCount,
        calls: callCount,
        sessionsPct: totalSessions > 0 ? (sessionCount / totalSessions) * 100 : 0,
        callsPct: totalCalls > 0 ? (callCount / totalCalls) * 100 : 0,
        combined
      };
    }).sort((a, b) => b.combined - a.combined).slice(0, 10);
    
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const sessionDayData = dayNames.map((name, idx) => {
      const pct = totalSessions > 0 ? ((dayCounts[idx] / totalSessions) * 100).toFixed(1) : '0';
      return [name, dayCounts[idx], pct + '%'];
    }).sort((a, b) => b[1] - a[1]);
    
    const callDayData = dayNames.map((name, idx) => {
      const pct = totalCalls > 0 ? ((callDayCounts[idx] / totalCalls) * 100).toFixed(1) : '0';
      return [name, callDayCounts[idx], pct + '%'];
    }).sort((a, b) => b[1] - a[1]);
    
    const combinedDayData = dayNames.map((name, idx) => {
      const sessions = dayCounts[idx];
      const calls = callDayCounts[idx] || 0;
      const combined = sessions + calls;
      return {
        name,
        sessions,
        calls,
        sessionsPct: totalSessions > 0 ? (sessions / totalSessions) * 100 : 0,
        callsPct: totalCalls > 0 ? (calls / totalCalls) * 100 : 0,
        combined
      };
    }).sort((a, b) => b.combined - a.combined);
    // Title
    sheet.getRange(startRow, 1).setValue('📊 Peak Hours & Day of Week Analysis');
    sheet.getRange(startRow, 1, 1, 10).merge();
    styleAnalyticsSectionHeader_(sheet, startRow, 10);
    
    // Session vs Call tables (hour)
    const hourSectionRow = startRow + 2;
    const sessionCol = 1;
    const callCol = 5;
    
    sheet.getRange(hourSectionRow, sessionCol).setValue('Session Volume by Hour');
    sheet.getRange(hourSectionRow, sessionCol).setFontSize(13).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    sheet.getRange(hourSectionRow + 1, sessionCol, 1, 3).setValues([['Hour', 'Sessions', 'Percentage']]);
    sheet.getRange(hourSectionRow + 2, sessionCol, sessionHourData.length, 3).setValues(sessionHourData);
    registerAnalyticsTable_(styleRegistry, hourSectionRow + 1, sessionCol, 3, sessionHourData.length);
    
    sheet.getRange(hourSectionRow, callCol).setValue('Call Volume by Hour');
    sheet.getRange(hourSectionRow, callCol).setFontSize(13).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    sheet.getRange(hourSectionRow + 1, callCol, 1, 3).setValues([['Hour', 'Calls', 'Percentage']]);
    if (callHourData.length) {
      sheet.getRange(hourSectionRow + 2, callCol, callHourData.length, 3).setValues(callHourData);
    }
    registerAnalyticsTable_(styleRegistry, hourSectionRow + 1, callCol, 3, Math.max(callHourData.length, 1));
    
    const combinedHourRow = hourSectionRow + 2 + Math.max(sessionHourData.length, callHourData.length) + 3;
    const combinedHourHeaders = ['Hour', 'Sessions', 'Calls', 'Sessions %', 'Calls %'];
    const combinedHourTable = combinedHourData.map(entry => [
      entry.hour + ':00',
      entry.sessions,
      entry.calls,
      entry.sessionsPct.toFixed(1) + '%',
      entry.callsPct.toFixed(1) + '%'
    ]);
    sheet.getRange(combinedHourRow, sessionCol).setValue('Combined Peaks by Hour');
    sheet.getRange(combinedHourRow, sessionCol).setFontSize(13).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    sheet.getRange(combinedHourRow + 1, sessionCol, 1, combinedHourHeaders.length).setValues([combinedHourHeaders]);
    if (combinedHourTable.length) {
      sheet.getRange(combinedHourRow + 2, sessionCol, combinedHourTable.length, combinedHourHeaders.length).setValues(combinedHourTable);
    }
    registerAnalyticsTable_(styleRegistry, combinedHourRow + 1, sessionCol, combinedHourHeaders.length, Math.max(combinedHourTable.length, 1));
    
    // Day-of-week section
    const daySectionRow = combinedHourRow + 2 + Math.max(combinedHourTable.length, 1) + 4;
    sheet.getRange(daySectionRow, sessionCol).setValue('Session Volume by Day');
    sheet.getRange(daySectionRow, sessionCol).setFontSize(13).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    sheet.getRange(daySectionRow + 1, sessionCol, 1, 3).setValues([['Day', 'Sessions', 'Percentage']]);
    sheet.getRange(daySectionRow + 2, sessionCol, sessionDayData.length, 3).setValues(sessionDayData);
    registerAnalyticsTable_(styleRegistry, daySectionRow + 1, sessionCol, 3, sessionDayData.length);
    
    sheet.getRange(daySectionRow, callCol).setValue('Call Volume by Day');
    sheet.getRange(daySectionRow, callCol).setFontSize(13).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    sheet.getRange(daySectionRow + 1, callCol, 1, 3).setValues([['Day', 'Calls', 'Percentage']]);
    if (callDayData.length) {
      sheet.getRange(daySectionRow + 2, callCol, callDayData.length, 3).setValues(callDayData);
    }
    registerAnalyticsTable_(styleRegistry, daySectionRow + 1, callCol, 3, Math.max(callDayData.length, 1));
    
    const combinedDayRow = daySectionRow + 2 + Math.max(sessionDayData.length, callDayData.length) + 3;
    const combinedDayHeaders = ['Day', 'Sessions', 'Calls', 'Sessions %', 'Calls %'];
    const combinedDayTable = combinedDayData.map(entry => [
      entry.name,
      entry.sessions,
      entry.calls,
      entry.sessionsPct.toFixed(1) + '%',
      entry.callsPct.toFixed(1) + '%'
    ]);
    sheet.getRange(combinedDayRow, sessionCol).setValue('Combined Peaks by Day');
    sheet.getRange(combinedDayRow, sessionCol).setFontSize(13).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    sheet.getRange(combinedDayRow + 1, sessionCol, 1, combinedDayHeaders.length).setValues([combinedDayHeaders]);
    if (combinedDayTable.length) {
      sheet.getRange(combinedDayRow + 2, sessionCol, combinedDayTable.length, combinedDayHeaders.length).setValues(combinedDayTable);
    }
    registerAnalyticsTable_(styleRegistry, combinedDayRow + 1, sessionCol, combinedDayHeaders.length, Math.max(combinedDayTable.length, 1));
    
    const summaryRow = combinedDayRow + 2 + Math.max(combinedDayTable.length, 1) + 2;
    const peakHourEntry = combinedHourData[0] || { hour: 0, sessions: 0, calls: 0 };
    const peakDayEntry = combinedDayData[0] || { name: 'Sunday', sessions: 0, calls: 0 };
    sheet.getRange(summaryRow, sessionCol, 1, 5)
      .setValues([[
        'Highlights',
        `Peak Hour: ${peakHourEntry.hour}:00 (Sessions: ${peakHourEntry.sessions}, Calls: ${peakHourEntry.calls})`,
        `Peak Day: ${peakDayEntry.name} (Sessions: ${peakDayEntry.sessions}, Calls: ${peakDayEntry.calls})`,
        `Total Sessions: ${totalSessions}`,
        `Total Calls: ${totalCalls}`
      ]])
      .setFontColor(ANALYTICS_THEME.accent)
      .setFontWeight('bold');
    
    return summaryRow + 2;
  } catch (e) {
    Logger.log('createPeakHoursAnalysis_ error: ' + e.toString());
    return startRow + 50;
  }
}
// 2. Technician Effectiveness Comparison
function createTechnicianEffectiveness_(sheet, startRow, startDate, endDate, styleRegistry, callAnalytics) {
  try {
    styleRegistry = styleRegistry || [];
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;

    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;

    const allDataRaw = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const allData = filterOutExcludedTechnicians_(headers, allDataRaw);

    const startIdx = analyticsGetHeaderIndex_(headers, ['start_time', 'start time', 'Start Time']);
    const techIdx = analyticsGetHeaderIndex_(headers, ['technician name', 'technician_name', 'technician', 'tech']);
    const durationIdx = analyticsGetHeaderIndex_(headers, ['total time', 'total_time', 'duration_total_seconds', 'duration seconds']);
    const pickupIdx = analyticsGetHeaderIndex_(headers, ['waiting time', 'waiting_time', 'pickup_seconds', 'pickup seconds']);
    const workIdx = analyticsGetHeaderIndex_(headers, ['work time', 'work_time', 'duration_work_seconds', 'work seconds']);
    const activeIdx = analyticsGetHeaderIndex_(headers, ['active time', 'active_time', 'duration_active_seconds', 'active seconds']);

    if (startIdx < 0 || techIdx < 0) {
      Logger.log('createTechnicianEffectiveness_: required headers missing (start or technician)');
      return startRow;
    }

    const startMillis = startDate.getTime();
    const endMillis = endDate.getTime();
    const filtered = allData.filter(row => {
      if (!row || !row[startIdx]) return false;
      try {
        const rowObj = row[startIdx] instanceof Date ? row[startIdx] : new Date(row[startIdx]);
        if (!(rowObj instanceof Date) || isNaN(rowObj)) return false;
        const rowMillis = rowObj.getTime();
        if (rowMillis < startMillis || rowMillis > endMillis) return false;
        if (techIdx >= 0 && isExcludedTechnician_(row[techIdx])) return false;
        return true;
      } catch (e) {
        return false;
      }
    });

    if (!filtered.length) return startRow;

    const techStats = {};
    const ensureEntry = (name) => {
      if (!techStats[name]) {
        techStats[name] = {
          sessions: 0,
          totalDuration: 0,
          durationCount: 0,
          totalPickup: 0,
          pickupCount: 0,
          slaHits: 0,
          workSeconds: 0,
          activeSeconds: 0
        };
      }
      return techStats[name];
    };

    filtered.forEach(row => {
      const tech = row[techIdx] || 'Unknown';
      const stats = ensureEntry(tech);
      stats.sessions++;

      if (durationIdx >= 0 && row[durationIdx]) {
        const dur = parseDurationSeconds_(row[durationIdx]);
        if (!isNaN(dur) && dur > 0) {
          stats.totalDuration += dur;
          stats.durationCount++;
        }
      }
      if (pickupIdx >= 0 && row[pickupIdx]) {
        const pickup = parseDurationSeconds_(row[pickupIdx]);
        if (!isNaN(pickup) && pickup > 0) {
          stats.totalPickup += pickup;
          stats.pickupCount++;
          if (pickup <= 60) stats.slaHits++;
        }
      }
      if (workIdx >= 0 && row[workIdx]) {
        const work = parseDurationSeconds_(row[workIdx]);
        if (!isNaN(work) && work > 0) {
          stats.workSeconds += work;
        }
      }
      if (activeIdx >= 0 && row[activeIdx]) {
        const active = parseDurationSeconds_(row[activeIdx]);
        if (!isNaN(active) && active > 0) {
          stats.activeSeconds += active;
        }
      }
    });

    const extensionMeta = (callAnalytics && callAnalytics.extensionMeta) || getActiveExtensionMetadata_();
    const extToName = extensionMeta && extensionMeta.extToName ? extensionMeta.extToName : {};
    const callStatsByCanonical = {};
    if (callAnalytics && callAnalytics.byAccount && callAnalytics.byAccount.perExtension) {
      const perExtension = callAnalytics.byAccount.perExtension;
      Object.keys(perExtension).forEach(ext => {
        const meta = perExtension[ext] || {};
        const label = extToName[ext] || (meta.label || '');
        const canonical = canonicalTechnicianName_(label);
        if (!canonical) return;
        const metrics = meta.metrics || {};
        const bucket = callStatsByCanonical[canonical] || { totalCalls: 0, talkingSeconds: 0, callDurationSeconds: 0 };
        bucket.totalCalls += Number(metrics.total_calls) || 0;
        const talkingSeconds = metrics.talking_duration != null ? parseDurationSeconds_(metrics.talking_duration) :
          (metrics['talking_duration (s)'] != null ? parseDurationSeconds_(metrics['talking_duration (s)']) : 0);
        const callDurationSeconds = metrics.call_duration != null ? parseDurationSeconds_(metrics.call_duration) :
          (metrics['call_duration (s)'] != null ? parseDurationSeconds_(metrics['call_duration (s)']) : 0);
        bucket.talkingSeconds += talkingSeconds;
        bucket.callDurationSeconds += callDurationSeconds;
        callStatsByCanonical[canonical] = bucket;
      });
    }

    const effectivenessRows = Object.keys(techStats)
      .map(name => {
        const stats = techStats[name];
        const avgDuration = stats.durationCount > 0 ? stats.totalDuration / stats.durationCount : 0;
        const avgPickup = stats.pickupCount > 0 ? stats.totalPickup / stats.pickupCount : 0;
        const workHours = stats.workSeconds / 3600;
        const activeHours = stats.activeSeconds / 3600;
        const sessionsPerHour = stats.workSeconds > 0 ? (stats.sessions / (stats.workSeconds / 3600)) : 0;
        const slaPct = stats.pickupCount > 0 ? (stats.slaHits / stats.pickupCount) : 0;
        const canonical = canonicalTechnicianName_(name);
        const callStats = callStatsByCanonical[canonical] || { totalCalls: 0, talkingSeconds: 0, callDurationSeconds: 0 };
        const avgCallDuration = callStats.totalCalls > 0 ? callStats.callDurationSeconds / callStats.totalCalls : 0;
        const talkHours = callStats.talkingSeconds / 3600;
        return {
          name,
          sessions: stats.sessions,
          avgDuration,
          avgPickup,
          workHours,
          activeHours,
          sessionsPerHour,
          slaPct,
          totalCalls: callStats.totalCalls,
          avgCallDuration,
          talkHours
        };
      })
      .sort((a, b) => b.sessions - a.sessions || a.name.localeCompare(b.name))
      .slice(0, 20) // cap to top 20 to keep section compact
      .map(entry => [
        entry.name,
        entry.sessions,
        entry.avgDuration / 86400,
        entry.avgPickup / 86400,
        entry.workHours,
        entry.activeHours,
        entry.sessionsPerHour,
        entry.slaPct,
        entry.totalCalls,
        entry.avgCallDuration / 86400,
        entry.talkHours
      ]);

    const headersRowValues = [['Technician', 'Sessions', 'Avg Duration', 'Avg Pickup', 'Work Hours', 'Active Hours', 'Sessions / Hour', 'SLA %', 'Total Calls', 'Avg Call Duration', 'Talk Hours']];
    const sectionTitleRow = startRow;
    const headerRow = startRow + 1;
    const dataStartRow = headerRow + 1;

    styleAnalyticsSectionHeader_(sheet, sectionTitleRow, headersRowValues[0].length);
    sheet.getRange(sectionTitleRow, 1).setValue('Technician Effectiveness Comparison');
    sheet.getRange(headerRow, 1, 1, headersRowValues[0].length).setValues(headersRowValues);

    if (effectivenessRows.length) {
      sheet.getRange(dataStartRow, 1, effectivenessRows.length, headersRowValues[0].length).setValues(effectivenessRows);
      registerAnalyticsTable_(styleRegistry, headerRow, 1, headersRowValues[0].length, effectivenessRows.length);

      try {
        sheet.getRange(dataStartRow, 2, effectivenessRows.length, 1).setNumberFormat('0'); // Sessions
        sheet.getRange(dataStartRow, 3, effectivenessRows.length, 2).setNumberFormat('hh:mm:ss'); // Avg duration/pickup
        sheet.getRange(dataStartRow, 5, effectivenessRows.length, 2).setNumberFormat('0.0'); // Work/Active hours
        sheet.getRange(dataStartRow, 7, effectivenessRows.length, 1).setNumberFormat('0.0'); // Sessions / Hour
        sheet.getRange(dataStartRow, 8, effectivenessRows.length, 1).setNumberFormat('0.0%'); // SLA %
        sheet.getRange(dataStartRow, 9, effectivenessRows.length, 1).setNumberFormat('0'); // Total Calls
        sheet.getRange(dataStartRow, 10, effectivenessRows.length, 1).setNumberFormat('hh:mm:ss'); // Avg Call Duration
        sheet.getRange(dataStartRow, 11, effectivenessRows.length, 1).setNumberFormat('0.0'); // Talk Hours
      } catch (e) { /* non-fatal */ }
    } else {
      sheet.getRange(dataStartRow, 1).setValue('No technician session data for the selected range.')
        .setFontColor(ANALYTICS_THEME.muted);
    }

    return dataStartRow + Math.max(effectivenessRows.length, 1) + 2;
  } catch (e) {
    Logger.log('createTechnicianEffectiveness_ error: ' + e.toString());
    return startRow + 10;
  }
}
function createRepeatCustomerAnalysis_(sheet, startRow, startDate, endDate, styleRegistry) {
  try {
    sheet.getRange(startRow, 1).setValue('Repeat Customer Analysis (coming soon)');
    sheet.getRange(startRow, 1).setFontStyle('italic').setFontColor(ANALYTICS_THEME.muted);
  } catch (e) {
    Logger.log('createRepeatCustomerAnalysis_ error: ' + e.toString());
  }
  return startRow + 2;
}
function createTrendAnalysis_(sheet, startRow, startDate, endDate, styleRegistry, callAnalytics) {
  try {
    styleRegistry = styleRegistry || [];
    const colSpan = 8;
    sheet.getRange(startRow, 1).setValue('📈 Trend Analysis');
    sheet.getRange(startRow, 1, 1, colSpan).merge();
    styleAnalyticsSectionHeader_(sheet, startRow, colSpan);
    sheet.getRange(startRow + 1, 1, 1, colSpan)
      .setValue('Daily session totals compared with mapped-extension call metrics')
      .setFontColor(ANALYTICS_THEME.text)
      .setFontStyle('italic');

    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) {
      sheet.getRange(startRow + 3, 1).setValue('Sessions sheet not found; unable to build trend analysis.')
        .setFontColor(ANALYTICS_THEME.muted);
      return startRow + 5;
    }

    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      sheet.getRange(startRow + 3, 1).setValue('No session data available for the selected date range.')
        .setFontColor(ANALYTICS_THEME.muted);
      return startRow + 5;
    }

    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const values = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const startIdx = analyticsGetHeaderIndex_(headers, ['start_time', 'start time', 'Start Time']);
    const pickupIdx = analyticsGetHeaderIndex_(headers, ['waiting time', 'waiting_time', 'pickup_seconds', 'pickup time']);
    if (startIdx < 0) {
      sheet.getRange(startRow + 3, 1).setValue('Start time column missing in sessions data.')
        .setFontColor(ANALYTICS_THEME.muted);
      return startRow + 5;
    }

    const tz = (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function')
      ? Session.getScriptTimeZone()
      : 'Etc/UTC';
    const startMillis = startDate.getTime();
    const endMillis = endDate.getTime();

    const dailyStats = {};
    const durationToSeconds = (value) => {
      if (value == null || value === '') return 0;
      if (value instanceof Date) {
        return value.getHours() * 3600 + value.getMinutes() * 60 + value.getSeconds();
      }
      if (typeof value === 'number') {
        if (value > 0 && value < 1) return Math.round(value * 86400);
        return value;
      }
      const str = String(value).trim();
      if (!str) return 0;
      if (str.includes(':')) return parseDurationSeconds_(str);
      const num = Number(str);
      if (!isNaN(num)) {
        if (num > 0 && num < 1) return Math.round(num * 86400);
        return num;
      }
      return 0;
    };

    values.forEach(row => {
      const rawDate = row[startIdx];
      if (!rawDate) return;
      const rowDate = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (!(rowDate instanceof Date) || isNaN(rowDate)) return;
      const rowMillis = rowDate.getTime();
      if (rowMillis < startMillis || rowMillis > endMillis) return;
      const key = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd');
      if (!dailyStats[key]) {
        dailyStats[key] = { sessions: 0, pickupSeconds: 0, pickupCount: 0 };
      }
      dailyStats[key].sessions += 1;
      if (pickupIdx >= 0 && row[pickupIdx] != null && row[pickupIdx] !== '') {
        const seconds = durationToSeconds(row[pickupIdx]);
        if (seconds > 0) {
          dailyStats[key].pickupSeconds += seconds;
          dailyStats[key].pickupCount += 1;
        }
      }
    });

    const callDaily = {};
    if (callAnalytics && callAnalytics.byDate && callAnalytics.byDate.ok && Array.isArray(callAnalytics.byDate.rows)) {
      const callDates = callAnalytics.byDate.dates || callAnalytics.byDate.categories || [];
      const callRowMap = {};
      callAnalytics.byDate.rows.forEach(row => {
        callRowMap[String(row[0] || '').trim().toLowerCase()] = row;
      });
      const totalRow = callRowMap['total calls'];
      const talkRow = callRowMap['talking duration (s)'] || callRowMap['talking duration'];
      callDates.forEach((dateKey, idx) => {
        if (!callDaily[dateKey]) callDaily[dateKey] = { totalCalls: 0, talkSeconds: 0 };
        if (totalRow && totalRow.length > idx + 1) {
          callDaily[dateKey].totalCalls = Number(totalRow[idx + 1]) || 0;
        }
        if (talkRow && talkRow.length > idx + 1) {
          callDaily[dateKey].talkSeconds = Number(talkRow[idx + 1]) || 0;
        }
      });
    }

    const dayList = [];
    const cursor = new Date(startDate.getTime());
    cursor.setHours(0, 0, 0, 0);
    const endCursor = new Date(endDate.getTime());
    endCursor.setHours(0, 0, 0, 0);
    while (cursor.getTime() <= endCursor.getTime()) {
      dayList.push(Utilities.formatDate(cursor, tz, 'yyyy-MM-dd'));
      cursor.setDate(cursor.getDate() + 1);
    }

    if (!dayList.length) {
      sheet.getRange(startRow + 3, 1).setValue('Invalid date range supplied.')
        .setFontColor(ANALYTICS_THEME.muted);
      return startRow + 5;
    }

    const headersRow = ['Date', 'Sessions', 'Calls', 'Calls per Session', 'Avg Pickup', 'Avg Talk / Call', 'Sessions Δ%', 'Calls Δ%'];
    const headerRowIndex = startRow + 3;
    const dataStartRow = headerRowIndex + 1;
    sheet.getRange(headerRowIndex, 1, 1, headersRow.length).setValues([headersRow]);
    sheet.getRange(headerRowIndex, 1, 1, headersRow.length)
      .setFontWeight('bold')
      .setFontColor(ANALYTICS_THEME.accent);

    const dataRows = [];
    let prevSessions = null;
    let prevCalls = null;
    dayList.forEach(dateKey => {
      const sessionStats = dailyStats[dateKey] || { sessions: 0, pickupSeconds: 0, pickupCount: 0 };
      const callStats = callDaily[dateKey] || { totalCalls: 0, talkSeconds: 0 };
      const sessionCount = sessionStats.sessions || 0;
      const totalCalls = callStats.totalCalls || 0;
      const avgPickupSeconds = sessionStats.pickupCount > 0 ? sessionStats.pickupSeconds / sessionStats.pickupCount : 0;
      const avgTalkSeconds = totalCalls > 0 ? (callStats.talkSeconds || 0) / totalCalls : 0;
      const ratio = sessionCount > 0 ? totalCalls / sessionCount : 0;
      const sessionDelta = (prevSessions != null && prevSessions > 0)
        ? (sessionCount - prevSessions) / prevSessions
        : 0;
      const callDelta = (prevCalls != null && prevCalls > 0)
        ? (totalCalls - prevCalls) / prevCalls
        : 0;
      const displayDate = Utilities.formatDate(new Date(dateKey + 'T00:00:00'), tz, 'MMM d');
      dataRows.push([
        displayDate,
        sessionCount,
        totalCalls,
        ratio,
        avgPickupSeconds / 86400,
        avgTalkSeconds / 86400,
        sessionDelta,
        callDelta
      ]);
      prevSessions = sessionCount;
      prevCalls = totalCalls;
    });

    if (dataRows.length) {
      sheet.getRange(dataStartRow, 1, dataRows.length, headersRow.length).setValues(dataRows);
      registerAnalyticsTable_(styleRegistry, headerRowIndex, 1, headersRow.length, dataRows.length);

      try {
        sheet.getRange(dataStartRow, 2, dataRows.length, 2).setNumberFormat('0');
        sheet.getRange(dataStartRow, 4, dataRows.length, 1).setNumberFormat('0.00');
        sheet.getRange(dataStartRow, 5, dataRows.length, 2).setNumberFormat('hh:mm:ss');
        sheet.getRange(dataStartRow, 7, dataRows.length, 2).setNumberFormat('0.0%');
      } catch (e) { /* non-fatal */ }
    } else {
      sheet.getRange(dataStartRow, 1).setValue('No matching session data for the selected range.')
        .setFontColor(ANALYTICS_THEME.muted);
    }

    return dataStartRow + Math.max(dataRows.length, 1) + 2;
  } catch (e) {
    Logger.log('createTrendAnalysis_ error: ' + e.toString());
    return startRow + 6;
  }
}
// Styling helper for Advanced_Analytics to improve readability without changing data
function applyAdvancedAnalyticsStyling_(sheet, styleRegistry) {
  try {
    if (!sheet) return;
    const lastRow = Math.max(sheet.getLastRow() || 1, 40);
    const lastCol = Math.max(sheet.getLastColumn() || 1, 8);

    const baseRange = sheet.getRange(1, 1, lastRow, lastCol);
    baseRange
      .setBackground(ANALYTICS_THEME.background)
      .setFontColor(ANALYTICS_THEME.text)
      .setFontFamily('Roboto')
      .setFontSize(10)
      .setVerticalAlignment('middle');

    try { sheet.setFrozenRows(3); } catch (e) {}

    // Header + meta rows
    try {
      sheet.getRange(1, 1, 1, lastCol)
        .setBackground(ANALYTICS_THEME.header)
        .setFontColor(ANALYTICS_THEME.accent)
        .setFontSize(18)
        .setFontWeight('bold');
      sheet.getRange(2, 1, 1, lastCol)
        .setBackground(ANALYTICS_THEME.subheader)
        .setFontColor(ANALYTICS_THEME.text)
        .setFontSize(11)
        .setFontWeight('normal');
    } catch (e) {}

    // Column widths tuned for dashboard layout
    try {
      const widths = [240, 170, 150, 140, 160, 140, 140, 140];
      for (let c = 1; c <= Math.min(widths.length, lastCol); c++) sheet.setColumnWidth(c, widths[c - 1]);
      for (let c = widths.length + 1; c <= lastCol; c++) {
        if (sheet.getColumnWidth(c) < 120) sheet.setColumnWidth(c, 120);
      }
    } catch (e) { /* non-fatal */ }

    // Apply registered table styling and highlights
    if (Array.isArray(styleRegistry) && styleRegistry.length) {
      styleRegistry.filter(entry => entry.type === 'table')
        .forEach(entry => applyAnalyticsTableTheme_(sheet, entry));
      styleRegistry.filter(entry => entry.type !== 'table')
        .forEach(entry => applyAnalyticsHighlight_(sheet, entry));
    }

    // Ensure wrap and alignment for top meta rows
    try { sheet.getRange(1, 1, 3, lastCol).setWrap(true).setHorizontalAlignment('left'); } catch (e) {}

  } catch (e) {
    Logger.log('applyAdvancedAnalyticsStyling_ error: ' + e.toString());
  }
}
// 8. Predictive Analytics
function createPredictiveAnalytics_(sheet, startRow, styleRegistry, startDate, endDate, callAnalytics) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sessionsSheet = ss.getSheetByName(SHEETS_SESSIONS_TABLE);
    if (!sessionsSheet) return startRow;
    
    const dataRange = sessionsSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return startRow;
    
    const allDataRaw = sessionsSheet.getRange(2, 1, dataRange.getNumRows() - 1, dataRange.getNumColumns()).getValues();
    const headers = sessionsSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0];
    const allData = filterOutExcludedTechnicians_(headers, allDataRaw);
    const startIdx = analyticsGetHeaderIndex_(headers, ['start_time', 'start time', 'Start Time']);
    if (startIdx < 0) {
      Logger.log('createPredictiveAnalytics_: start_time header missing');
      return startRow;
    }
    
    const sessionDailyCounts = {};
    allData.forEach(row => {
      if (!row[startIdx]) return;
      try {
        const date = new Date(row[startIdx]);
        const key = date.toISOString().split('T')[0];
        sessionDailyCounts[key] = (sessionDailyCounts[key] || 0) + 1;
      } catch (e) { /* ignore */ }
    });

    const sessionDatesSorted = Object.keys(sessionDailyCounts).sort();
    if (sessionDatesSorted.length < 7) {
      sheet.getRange(startRow, 1).setValue('🔮 Predictive Analytics');
      sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold');
      sheet.getRange(startRow + 2, 1).setValue('Insufficient historical data for predictions');
      return startRow + 5;
    }
    
    const lookbackWindow = Math.min(45, sessionDatesSorted.length);
    const recentSessionDates = sessionDatesSorted.slice(-lookbackWindow);
    const sessionSeries = recentSessionDates.map(date => sessionDailyCounts[date]);

    let callSeries = [];
    let callSeriesDates = [];
    if (callAnalytics && callAnalytics.byDate && callAnalytics.byDate.ok && Array.isArray(callAnalytics.byDate.rows)) {
      const callRow = callAnalytics.byDate.rows.find(r => /total calls/i.test(String(r[0] || '')));
      if (callRow) {
        const callDates = callAnalytics.byDate.dates || [];
        const callMap = {};
        callDates.forEach((date, idx) => { callMap[date] = Number(callRow[idx + 1]) || 0; });
        const matchingDates = recentSessionDates.filter(date => Object.prototype.hasOwnProperty.call(callMap, date));
        if (matchingDates.length >= 7) {
          callSeriesDates = matchingDates;
          callSeries = matchingDates.map(date => callMap[date]);
        } else {
          const availableDates = Object.keys(callMap).sort();
          const callLookback = Math.min(lookbackWindow, availableDates.length);
          callSeriesDates = availableDates.slice(-callLookback);
          callSeries = callSeriesDates.map(date => callMap[date]);
        }
      }
    }

    if (callSeries.length === 0) {
      callSeries = sessionSeries.map(() => 0);
      callSeriesDates = recentSessionDates.slice();
    }

    const linearForecast = (series, periods) => {
      if (!series || series.length < 2) {
        const fallback = series && series.length ? series[series.length - 1] : 0;
        return {
          slope: 0,
          intercept: fallback,
          predictions: Array(periods).fill(fallback)
        };
      }
      const n = series.length;
      const indices = Array.from({ length: n }, (_, i) => i);
      const sumX = indices.reduce((sum, x) => sum + x, 0);
      const sumY = series.reduce((sum, y) => sum + y, 0);
      const sumXY = series.reduce((sum, y, i) => sum + i * y, 0);
      const sumX2 = indices.reduce((sum, x) => sum + x * x, 0);
      const denominator = n * sumX2 - sumX * sumX;
      const slope = denominator !== 0 ? (n * sumXY - sumX * sumY) / denominator : 0;
      const intercept = (sumY - slope * sumX) / n;
      const predictions = [];
      for (let i = 1; i <= periods; i++) {
        const xFuture = n - 1 + i;
        let value = slope * xFuture + intercept;
        if (value < 0) value = 0;
        predictions.push(value);
      }
      return { slope, intercept, predictions };
    };

    const forecastHorizon = 7;
    const sessionForecast = linearForecast(sessionSeries, forecastHorizon);
    const callForecast = linearForecast(callSeries, forecastHorizon);

    const futureDates = [];
    const base = new Date(endDate);
    for (let i = 1; i <= forecastHorizon; i++) {
      const next = new Date(base);
      next.setDate(next.getDate() + i);
      futureDates.push(Utilities.formatDate(next, Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Etc/GMT', 'MMM d'));
    }

    const forecastTable = futureDates.map((label, idx) => [
      label,
      Math.round(sessionForecast.predictions[idx]),
      Math.round(callForecast.predictions[idx])
    ]);
    const averageSessions = sessionSeries.reduce((a, b) => a + b, 0) / sessionSeries.length;
    const averageCalls = callSeries.reduce((a, b) => a + b, 0) / callSeries.length;

    sheet.getRange(startRow, 1).setValue('🔮 Predictive Analytics');
    sheet.getRange(startRow, 1).setFontSize(16).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    sheet.getRange(startRow, 1, 1, 6).merge();
    
    const infoRow = startRow + 2;
    sheet.getRange(infoRow, 1).setValue(`Using the last ${lookbackWindow} days of Support session data${callSeries.some(v => v > 0) ? ' and Digium call volume' : ''}.`);
    sheet.getRange(infoRow, 1).setFontStyle('italic').setFontColor(ANALYTICS_THEME.muted);

    const tableHeaderRow = infoRow + 2;
    const forecastHeaders = ['Date', 'Projected Sessions', 'Projected Calls'];
    sheet.getRange(tableHeaderRow, 1, 1, forecastHeaders.length).setValues([forecastHeaders]);
    sheet.getRange(tableHeaderRow + 1, 1, forecastTable.length, forecastHeaders.length).setValues(forecastTable);
    registerAnalyticsTable_(styleRegistry, tableHeaderRow, 1, forecastHeaders.length, forecastTable.length);

    const insightRow = tableHeaderRow + forecastTable.length + 3;
    const sessionTrend = sessionForecast.slope > 0 ? 'increasing' : (sessionForecast.slope < 0 ? 'softening' : 'steady');
    const callTrend = callForecast.slope > 0 ? 'increasing' : (callForecast.slope < 0 ? 'softening' : 'steady');
    const recommendations = [
      `Sessions trend: ${sessionTrend}. Average daily sessions: ${averageSessions.toFixed(1)}.`,
      `Calls trend: ${callTrend}. Average daily calls: ${averageCalls.toFixed(1)}.`,
      `Plan staffing for peak forecast day: ${futureDates[sessionForecast.predictions.indexOf(Math.max(...sessionForecast.predictions))]}`
    ];

    sheet.getRange(insightRow, 1).setValue('Recommendations');
    sheet.getRange(insightRow, 1).setFontWeight('bold').setFontColor(ANALYTICS_THEME.accent);
    recommendations.forEach((rec, idx) => {
      sheet.getRange(insightRow + 1 + idx, 1).setValue('• ' + rec);
    });
    
    return insightRow + recommendations.length + 2;
  } catch (e) {
    Logger.log('createPredictiveAnalytics_ error: ' + e.toString());
 
    return startRow + 50;
  }
}