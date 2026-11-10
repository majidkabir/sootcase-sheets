/**
 * Sootcase for Google Sheets: freight emissions (GLEC Framework / ISO 14083)
 * from the Sootcase API, as a cell function and for a whole sheet of shipments.
 *
 * https://github.com/majidkabir/sootcase-sheets (MIT License)
 * Calculations: https://sootcase.org (https://github.com/majidkabir/sootcase)
 */

const SOOTCASE_DEFAULT_URL = 'https://sootcase.org';
const KEY_PROP = 'SOOTCASE_API_KEY';
const URL_PROP = 'SOOTCASE_URL';
const CHUNK = 500;              // shipments per request from SOOTCASE()
const MAX_SHEET_SHIPMENTS = 10000; // the API's limit per report
const CACHE_SECONDS = 21600;    // 6 hours, the cache's maximum

const DEFAULT_VEHICLES = {
  road: 'hgv',
  rail: 'freight_train',
  sea: 'container_ship',
  air: 'air_freight',
  inland_waterway: 'barge',
};

const MODE_ALIASES = {
  truck: 'road', lorry: 'road',
  train: 'rail',
  ship: 'sea', ocean: 'sea', vessel: 'sea',
  plane: 'air', flight: 'air',
  barge: 'inland_waterway', river: 'inland_waterway', waterway: 'inland_waterway', iww: 'inland_waterway',
};

const VALUES = {
  wtw: 'wtw_kg', wtw_kg: 'wtw_kg',
  ttw: 'ttw_kg', ttw_kg: 'ttw_kg',
  wtt: 'wtt_kg', wtt_kg: 'wtt_kg',
  tkm: 'tonne_km', tonne_km: 'tonne_km',
  intensity: 'intensity_g_per_tkm', intensity_g_per_tkm: 'intensity_g_per_tkm',
  distance: 'distance_km', distance_km: 'distance_km',
};

/**
 * Freight emissions of one transport leg, in kg CO2e well-to-wheel (GLEC Framework / ISO 14083).
 *
 * Works on single cells or whole columns: =SOOTCASE(A2:A100, B2:B100, C2:C100, "road").
 *
 * @param {string} from Start: a UN/LOCODE (NLRTM), an IATA code (CPH), a place name (Copenhagen) or "lat,lon".
 * @param {string} to End, the same way.
 * @param {number} tonnes Cargo weight in tonnes.
 * @param {string} mode road, rail, sea, air or inland_waterway.
 * @param {string} vehicle Optional vehicle type, e.g. hgv_artic, freight_train_electric, container_ship. Default per mode.
 * @param {number} distance_km Optional known distance in km; otherwise Sootcase routes it.
 * @param {string} value Optional: wtw (default), ttw, wtt, tonne_km, intensity (g CO2e per tonne-km) or distance_km.
 * @return The value asked for, per row.
 * @customfunction
 */
function SOOTCASE(from, to, tonnes, mode, vehicle, distance_km, value) {
  const args = [from, to, tonnes, mode, vehicle, distance_km, value];
  const isRange = args.some(Array.isArray);
  const rows = isRange ? Math.max(...args.filter(Array.isArray).map((a) => a.length)) : 1;
  const cell = (arg, i) => (Array.isArray(arg) ? (arg[i] || [])[0] : arg);

  const requests = [];
  for (let i = 0; i < rows; i++) {
    const r = args.map((a) => cell(a, i));
    if (blank_(r[0]) && blank_(r[1]) && blank_(r[2])) {
      requests.push(null); // empty row in a range
      continue;
    }
    try {
      requests.push({ input: legInput_(r[0], r[1], r[2], r[3], r[4], r[5]), field: valueField_(r[6]) });
    } catch (e) {
      requests.push({ error: e.message });
    }
  }

  calculateLegs_(requests);

  const out = requests.map((q) => {
    if (q === null) return '';
    if (q.error) {
      if (!isRange) throw new Error(q.error);
      return 'Error: ' + q.error;
    }
    return q.result[q.field];
  });
  return isRange ? out.map((v) => [v]) : out[0];
}

/** Adds the Sootcase menu. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sootcase')
    .addItem('Calculate shipments on this sheet', 'calculateSheet')
    .addItem('Insert shipments template', 'insertTemplate')
    .addSeparator()
    .addItem('Set my API key…', 'setApiKey')
    .addItem('Set API key for this spreadsheet…', 'setSpreadsheetApiKey')
    .addItem('Use another Sootcase server…', 'setServer')
    .addSeparator()
    .addItem('About', 'about')
    .addToUi();
}

/** For the add-on: show the menu right after installing. */
function onInstall(e) {
  onOpen(e);
}

/**
 * Calculates the shipments on the active sheet (Sootcase's CSV columns, one row per leg or hub stop)
 * and writes three sheets: per shipment, per leg, and a summary report.
 */
function calculateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  if ([RESULTS_SHEET, LEGS_SHEET, SUMMARY_SHEET].indexOf(sheet.getName()) >= 0) {
    ui.alert('Sootcase', 'Open the sheet with your shipments first (not a results sheet).', ui.ButtonSet.OK);
    return;
  }
  const values = sheet.getDataRange().getValues();
  const header = (values[0] || []).map((h) => String(h).trim().toLowerCase());
  if (header.indexOf('shipment_id') < 0) {
    ui.alert('Sootcase',
      'The first row needs Sootcase\'s column names, at least shipment_id. ' +
      'Use Sootcase > Insert shipments template to start from an example.', ui.ButtonSet.OK);
    return;
  }
  // Empty rows stay in the CSV (Sootcase skips them) so the line numbers in its
  // messages are the sheet's row numbers.
  const rows = values.slice(1);
  const idCol = header.indexOf('shipment_id');
  const ids = new Set();
  for (let i = 0; i < rows.length; i++) {
    if (!rows[i].some((v) => !blank_(v))) continue;
    if (blank_(rows[i][idCol])) {
      ui.alert('Sootcase', 'Row ' + (i + 2) + ' has no shipment_id. Every leg and hub row needs one.', ui.ButtonSet.OK);
      return;
    }
    ids.add(String(rows[i][idCol]).trim());
  }
  if (ids.size === 0) {
    ui.alert('Sootcase', 'There are no shipments below the header row.', ui.ButtonSet.OK);
    return;
  }
  if (ids.size > MAX_SHEET_SHIPMENTS) {
    ui.alert('Sootcase', 'This sheet has ' + ids.size + ' shipments; one run takes up to ' +
      MAX_SHEET_SHIPMENTS + '. Split it over several sheets.', ui.ButtonSet.OK);
    return;
  }

  ss.toast('Calculating ' + ids.size + ' shipments…', 'Sootcase', -1);
  const tz = ss.getSpreadsheetTimeZone();
  const csv = [header].concat(rows).map((r) => r.map((v) => csvCell_(v, tz)).join(',')).join('\n');
  const q = 'results=true&title=' + encodeURIComponent(sheet.getName());
  let report;
  try {
    report = request_('post', '/v1/report?' + q, csv, 'text/csv');
  } catch (e) {
    ss.toast('Not calculated.', 'Sootcase', 3); // replaces the "Calculating…" toast
    ui.alert('Sootcase', e.message, ui.ButtonSet.OK);
    return;
  }

  writeResults_(ss, report);
  writeLegs_(ss, report);
  writeSummary_(ss, report, sheet.getName());
  ss.setActiveSheet(ss.getSheetByName(RESULTS_SHEET));

  const c = report.coverage || {};
  ss.toast(c.calculated + ' calculated, ' + (c.failed || 0) + ' failed. Total ' +
    round_(report.total.wtw_kg / 1000, 3) + ' t CO2e.', 'Sootcase', 10);
}

/** Inserts a "Shipments" sheet with Sootcase's columns and a few example rows. */
function insertTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let name = 'Shipments';
  for (let n = 2; ss.getSheetByName(name); n++) name = 'Shipments ' + n;
  const sheet = ss.insertSheet(name);
  const rows = TEMPLATE_ROWS;
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}

const TEMPLATE_ROWS = [
  ['shipment_id', 'weight_kg', 'leg_id', 'mode', 'vehicle_type', 'from_code', 'to_code', 'distance_km', 'hub_type'],
  ['CPH-RTM-1', 12000, 'truck', 'road', 'hgv', 'DKCPH', 'DEHAM', '', ''],
  ['CPH-RTM-1', '', 'feeder', 'sea', 'container_ship', 'DEHAM', 'NLRTM', '', ''],
  ['CPH-FRA-AIR', 800, 'flight', 'air', 'air_freight', 'CPH', 'FRA', '', ''],
  ['HAM-MIL-RAIL', 20000, 'train', 'rail', 'freight_train', 'DEHAM', 'ITMIL', '', ''],
  ['HAM-MIL-RAIL', '', 'dc', '', '', '', '', '', 'warehouse'],
];

/** Stores the user's own API key (from sootcase.org). */
function setApiKey() {
  promptProperty_(PropertiesService.getUserProperties(), KEY_PROP,
    'Your Sootcase API key (get one at sootcase.org). Only you can read it. Leave empty to remove it.');
}

/**
 * Stores an API key with this spreadsheet. SOOTCASE() cells only get the spreadsheet owner's own
 * settings (Google's rule for custom functions); without an owner's key they use this one. Anyone
 * who can edit the spreadsheet's script can read it.
 */
function setSpreadsheetApiKey() {
  promptProperty_(PropertiesService.getDocumentProperties(), KEY_PROP,
    'API key for this spreadsheet. Anyone who can edit its script can read it. Leave empty to remove it.');
}

/** Points the add-in at another Sootcase server, e.g. your own `sootcase serve`. */
function setServer() {
  promptProperty_(PropertiesService.getUserProperties(), URL_PROP,
    'Sootcase server URL, starting with https:// or http:// (now ' + baseUrl_() + '). Leave empty for ' +
    SOOTCASE_DEFAULT_URL + '.', (v) => /^https?:\/\/[^\s/]+(\/\S*)?$/.test(v));
}

function about() {
  SpreadsheetApp.getUi().alert('Sootcase for Google Sheets',
    'Freight emissions following the GLEC Framework v3.3 and ISO 14083, calculated by ' + baseUrl_() + '.\n\n' +
    'In a cell: =SOOTCASE("Copenhagen", "Hamburg", 12, "road")\n' +
    'For many shipments: Sootcase > Insert shipments template, fill it, then Calculate.\n\n' +
    'Results are calculations for your own reporting; Sootcase is not certified against ISO 14083.\n' +
    'https://github.com/majidkabir/sootcase-sheets',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ---- cell function internals ----

/** Builds a one-leg shipment from SOOTCASE()'s arguments. */
function legInput_(from, to, tonnes, mode, vehicle, distanceKm) {
  const t = Number(tonnes);
  if (blank_(tonnes) || !isFinite(t) || t <= 0) throw new Error('tonnes must be a number above 0');
  const m = modeOf_(mode);
  const leg = {
    mode: m,
    vehicle_type: blank_(vehicle) ? DEFAULT_VEHICLES[m] : String(vehicle).trim(),
    from: placeInput_(from, 'from'),
    to: placeInput_(to, 'to'),
  };
  if (!blank_(distanceKm)) {
    const d = Number(distanceKm);
    if (!isFinite(d) || d <= 0) throw new Error('distance_km must be a number above 0');
    leg.distance_km = d;
  }
  return { weight_kg: t * 1000, legs: [leg] };
}

function modeOf_(mode) {
  if (blank_(mode)) throw new Error('mode is required: road, rail, sea, air or inland_waterway');
  const m = String(mode).trim().toLowerCase().replace(/[\s-]+/g, '_');
  const r = MODE_ALIASES[m] || m;
  if (!DEFAULT_VEHICLES[r]) throw new Error('unknown mode "' + mode + '": use road, rail, sea, air or inland_waterway');
  return r;
}

function valueField_(value) {
  if (blank_(value)) return 'wtw_kg';
  const f = VALUES[String(value).trim().toLowerCase()];
  if (!f) throw new Error('unknown value "' + value + '": use wtw, ttw, wtt, tonne_km, intensity or distance_km');
  return f;
}

/**
 * A place as the API takes it: {code} for an upper-case UN/LOCODE or IATA code, {lat, lon} for
 * "lat,lon", or {lookup} for a name to look up (resolved in calculateLegs_).
 */
function placeInput_(v, which) {
  if (blank_(v)) throw new Error(which + ' is required');
  const s = String(v).trim();
  const ll = s.match(/^(-?\d+(?:\.\d+)?)\s*[,;]\s*(-?\d+(?:\.\d+)?)$/);
  if (ll) {
    const lat = Number(ll[1]);
    const lon = Number(ll[2]);
    if (Math.abs(lat) > 90 || Math.abs(lon) > 180) throw new Error(which + ': "' + s + '" is not a valid lat,lon');
    return { lat: lat, lon: lon };
  }
  if (/^[A-Z]{2}[A-Z2-9]{3}$/.test(s) || /^[A-Z]{3}$/.test(s)) return { code: s };
  return { lookup: s };
}

/** Fills in each request's result (or error), using the cache and as few API calls as possible. */
function calculateLegs_(requests) {
  const cache = CacheService.getScriptCache();
  const todo = requests.filter((q) => q && !q.error);

  // Place names -> places.
  const names = [];
  todo.forEach((q) => ['from', 'to'].forEach((k) => {
    const p = q.input.legs[0][k];
    if (p.lookup && names.indexOf(p.lookup) < 0) names.push(p.lookup);
  }));
  const places = lookupPlaces_(names, cache);
  todo.forEach((q) => {
    ['from', 'to'].forEach((k) => {
      const p = q.input.legs[0][k];
      if (!p.lookup) return;
      const found = places[p.lookup];
      if (typeof found === 'string') q.error = found;
      else q.input.legs[0][k] = found;
    });
  });

  // Shipments -> results, cached by server and input. All chunks go out at
  // once: a cell function has 30 seconds.
  const ready = todo.filter((q) => !q.error);
  ready.forEach((q) => { q.key = 'calc:' + digest_(baseUrl_() + JSON.stringify(q.input)); });
  const hits = cacheGet_(cache, ready.map((q) => q.key));
  const pending = [];
  ready.forEach((q) => {
    if (hits[q.key]) q.result = JSON.parse(hits[q.key]);
    else pending.push(q);
  });
  if (!pending.length) return;
  const parts = [];
  for (let i = 0; i < pending.length; i += CHUNK) parts.push(pending.slice(i, i + CHUNK));
  const resps = UrlFetchApp.fetchAll(parts.map((part) => ({
    url: baseUrl_() + '/v1/calculate/batch',
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ shipments: part.map((q, j) => Object.assign({ id: 'r' + j }, q.input)) }),
    headers: headers_(),
    muteHttpExceptions: true,
  })));
  const store = {};
  resps.forEach((resp, n) => {
    const part = parts[n];
    let body;
    try {
      body = parse_(resp);
    } catch (e) {
      part.forEach((q) => { q.error = e.message; }); // other chunks keep their answers
      return;
    }
    const byId = {};
    (body.results || []).forEach((item, j) => { byId[item.shipment_id || 'r' + j] = item; });
    part.forEach((q, j) => {
      const item = byId['r' + j];
      if (!item) q.error = 'no result from Sootcase';
      else if (item.error) q.error = item.error + (item.retryable ? ' (try again)' : '');
      else {
        q.result = summary_(item.result);
        store[q.key] = JSON.stringify(q.result);
      }
    });
  });
  cachePut_(cache, store);
}

/** The numbers SOOTCASE() can return, from one API result. */
function summary_(r) {
  return {
    wtw_kg: r.wtw_kg,
    ttw_kg: r.ttw_kg,
    wtt_kg: r.wtt_kg,
    tonne_km: r.tonne_km,
    intensity_g_per_tkm: r.intensity_g_per_tkm,
    distance_km: (r.legs || []).reduce((s, l) => s + (l.distance_km || 0), 0),
  };
}

/**
 * Looks up place names: name -> the place as the API takes it ({code}, or {name, lat, lon}
 * for a place without one), or an error message string.
 */
function lookupPlaces_(names, cache) {
  const out = {};
  const missing = [];
  const key = (n) => 'place:' + digest_(baseUrl_() + '|' + n.toLowerCase());
  const hits = cacheGet_(cache, names.map(key));
  names.forEach((n) => {
    const hit = hits[key(n)];
    if (hit) out[n] = JSON.parse(hit);
    else missing.push(n);
  });
  if (!missing.length) return out;
  const resps = UrlFetchApp.fetchAll(missing.map((n) => ({
    url: baseUrl_() + '/v1/places?limit=1&q=' + encodeURIComponent(n),
    method: 'get',
    headers: headers_(),
    muteHttpExceptions: true,
  })));
  const store = {};
  resps.forEach((resp, i) => {
    const n = missing[i];
    let p;
    try {
      p = (parse_(resp).places || [])[0];
    } catch (e) {
      out[n] = e.message; // only rows with this place fail
      return;
    }
    if (!p) {
      out[n] = 'no place found for "' + n + '": try a UN/LOCODE (DKCPH) or "lat,lon"';
      return;
    }
    out[n] = p.code ? { code: p.code } : { name: p.name + (p.country ? ', ' + p.country : ''), lat: p.lat, lon: p.lon };
    store[key(n)] = JSON.stringify(out[n]);
  });
  cachePut_(cache, store);
  return out;
}

/**
 * Reads many cache keys in one call (the cache keeps about 1,000 items, so a large range is
 * only partly cached; that is fine, the rest is calculated again).
 */
function cacheGet_(cache, keys) {
  return keys.length ? cache.getAll(keys) || {} : {};
}

function cachePut_(cache, store) {
  const keys = Object.keys(store);
  for (let i = 0; i < keys.length; i += 100) {
    const part = {};
    keys.slice(i, i + 100).forEach((k) => { part[k] = store[k]; });
    cache.putAll(part, CACHE_SECONDS);
  }
}

// ---- sheet output ----

const RESULTS_SHEET = 'Sootcase results';
const LEGS_SHEET = 'Sootcase legs';
const SUMMARY_SHEET = 'Sootcase summary';

function writeResults_(ss, report) {
  const rows = [['shipment_id', 'weight_kg', 'legs', 'distance_km', 'tonne_km', 'ttw_kg', 'wtt_kg',
    'hub_wtw_kg', 'wtw_kg', 'intensity_g_per_tkm', 'primary_data_share', 'added_road_legs', 'warnings', 'error']];
  (report.results || []).forEach((item) => {
    const r = item.result;
    if (!r) {
      rows.push([item.shipment_id, '', '', '', '', '', '', '', '', '', '', '', '',
        (item.error || 'not calculated') + (item.retryable ? ' (send again)' : '')]);
      return;
    }
    const legs = r.legs || [];
    const hubs = (r.hubs || []).reduce((s, h) => s + (h.wtw_kg || 0), 0);
    rows.push([item.shipment_id, r.weight_kg, legs.length,
      round_(legs.reduce((s, l) => s + (l.distance_km || 0), 0), 1),
      round_(r.tonne_km, 1), round_(r.ttw_kg, 3), round_(r.wtt_kg, 3), round_(hubs, 3),
      round_(r.wtw_kg, 3), round_(r.intensity_g_per_tkm, 2), round_(r.primary_data_share, 3),
      legs.filter((l) => l.added).length, (r.warnings || []).join(' | '), '']);
  });
  fillSheet_(ss, RESULTS_SHEET, rows, [0, 12, 13]);
}

function writeLegs_(ss, report) {
  const rows = [['shipment_id', 'leg_id', 'mode', 'vehicle_type', 'factor_vehicle', 'distance_km',
    'distance_method', 'distance_type', 'daf', 'method', 'data_type', 'tonne_km', 'ttw_kg', 'wtt_kg',
    'wtw_kg', 'intensity_g_per_tkm', 'country_km', 'factor_source', 'note']];
  (report.results || []).forEach((item) => {
    const r = item.result;
    if (!r) return;
    (r.legs || []).forEach((l) => {
      const f = l.factor || {};
      const note = [l.added ? 'added road leg (' + l.added + ')' : '', l.note || '', l.factor_note || '']
        .filter(String).join(' ');
      rows.push([item.shipment_id, l.leg_id || '', l.mode, l.vehicle_type, f.vehicle_type || '',
        round_(l.distance_km, 1), l.distance_method || '', l.distance_type || '', l.daf || '',
        l.method || '', l.data_type || '', round_(l.tonne_km, 1), round_(l.ttw_kg, 3),
        round_(l.wtt_kg, 3), round_(l.wtw_kg, 3), round_(l.intensity_g_per_tkm, 2), countryKm_(l.country_km),
        f.source || '', note]);
    });
    (r.hubs || []).forEach((h) => {
      rows.push([item.shipment_id, h.hub_id || '', 'hub', h.hub_type || '', '', '', '', '', '', '',
        h.data_type || '', '', round_(h.ttw_kg, 3), round_(h.wtt_kg, 3), round_(h.wtw_kg, 3), '', '',
        h.source || '', h.note || '']);
    });
  });
  fillSheet_(ss, LEGS_SHEET, rows, [0, 1, 2, 3, 4, 6, 7, 9, 10, 16, 17, 18]);
}

function writeSummary_(ss, report, sourceName) {
  const t = report.total || {};
  const c = report.coverage || {};
  const dq = report.data_quality || {};
  const g = report.ghg_protocol || {};
  const W = 6;
  const pad = (r) => r.concat(Array(Math.max(0, W - r.length)).fill(''));
  const rows = [];
  const add = (...r) => rows.push(pad(r));

  add('Freight GHG emissions, ' + sourceName);
  add(report.statement || '');
  add('Calculated', report.generated || '', 'Dataset', (report.datasets || []).join(', '));
  add('');
  add('Total', 'kg CO2e');
  add('Well-to-wheel (all)', round_(t.wtw_kg, 1));
  add('Transport, operational (TTW)', round_(t.ttw_kg, 1));
  add('Transport, energy provision (WTT)', round_(t.wtt_kg, 1));
  add('Hubs (WTW)', round_(t.hub_wtw_kg || 0, 1));
  add('Tonne-km', round_(t.tonne_km, 1));
  add('Intensity, g CO2e per tonne-km', round_(t.intensity_g_per_tkm, 2));
  add('GHG Protocol scope 3 (for the shipper)', round_(g.scope3_shipper_kg, 1));
  add('');
  add('Coverage');
  add('Shipments', c.shipments, 'Calculated', c.calculated, 'Failed', c.failed || 0);
  add('Legs', c.legs, 'Hub stops', c.hub_stops || 0, 'Added road legs', c.added_road_legs || 0);
  add('');
  add('Mode', 'Legs', 'Tonne-km', 'TTW kg', 'WTT kg', 'WTW kg');
  (report.modes || []).forEach((m) => add(m.mode, m.legs, round_(m.tonne_km, 1), round_(m.ttw_kg, 1),
    round_(m.wtt_kg, 1), round_(m.wtw_kg, 1)));
  if ((report.hubs || []).length) {
    add('');
    add('Hub type', 'Stops', 'Tonnes', 'WTW kg');
    report.hubs.forEach((h) => add(h.hub_type, h.stops, round_(h.tonnes, 1), round_(h.wtw_kg, 1)));
  }
  add('');
  add('Data quality', 'Share of WTW');
  add('Primary data', round_(dq.primary_share, 3));
  add('Modelled', round_(dq.modelled_share, 3));
  add('Default values', round_(dq.default_share, 3));
  (dq.default_sources || []).forEach((s) => add('Source: ' + s.source, round_(s.share, 3)));
  if ((report.notes || []).length) {
    add('');
    add('Notes');
    report.notes.forEach((n) => add(n));
  }
  fillSheet_(ss, SUMMARY_SHEET, rows, [0]);
}

/**
 * Replaces a sheet's contents (creating the sheet when needed). textCols (0-based) are set to
 * plain text first, so IDs like 00123 keep their zeros and text starting with = isn't a formula.
 */
function fillSheet_(ss, name, rows, textCols) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();
  if (!rows.length) return;
  (textCols || []).forEach((c) => sheet.getRange(1, c + 1, rows.length, 1).setNumberFormat('@'));
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// ---- HTTP and helpers ----

function baseUrl_() {
  const u = PropertiesService.getUserProperties().getProperty(URL_PROP);
  return (u || SOOTCASE_DEFAULT_URL).replace(/\/+$/, '');
}

function apiKey_() {
  return PropertiesService.getUserProperties().getProperty(KEY_PROP) ||
    PropertiesService.getDocumentProperties().getProperty(KEY_PROP) || '';
}

function headers_() {
  const key = apiKey_();
  return key ? { Authorization: 'Bearer ' + key } : {};
}

/** Calls the API and returns its JSON, or throws a readable error. */
function request_(method, path, payload, contentType) {
  const resp = UrlFetchApp.fetch(baseUrl_() + path, {
    method: method,
    contentType: contentType,
    payload: payload,
    headers: headers_(),
    muteHttpExceptions: true,
  });
  return parse_(resp);
}

function parse_(resp) {
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  let body = {};
  try {
    body = JSON.parse(text);
  } catch (e) {
    body = { error: text.slice(0, 200) };
  }
  if (code >= 200 && code < 300) return body;
  const msg = body.error || ('HTTP ' + code);
  if (code === 401 || code === 403) {
    throw new Error('Sootcase didn\'t accept the request (' + msg + '). Set your API key in the Sootcase menu ' +
      '(for =SOOTCASE() cells: Set API key for this spreadsheet).');
  }
  if (code === 429) throw new Error('Sootcase quota reached (' + msg + '). Try again later.');
  if (code === 503) throw new Error('Sootcase is busy or a routing service didn\'t answer. Try again in a minute.');
  throw new Error('Sootcase: ' + msg);
}

function promptProperty_(props, prop, text, valid) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Sootcase', text, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const v = r.getResponseText().trim();
  if (v && valid && !valid(v)) {
    ui.alert('Sootcase', '"' + v + '" isn\'t a valid value; nothing was changed.', ui.ButtonSet.OK);
    return;
  }
  if (v) props.setProperty(prop, v);
  else props.deleteProperty(prop);
  SpreadsheetApp.getActiveSpreadsheet().toast(v ? 'Saved.' : 'Removed.', 'Sootcase', 3);
}

/** "DE 210.4; DK 35.1", largest first. */
function countryKm_(m) {
  if (!m) return '';
  return Object.keys(m).sort((a, b) => m[b] - m[a]).map((c) => c + ' ' + round_(m[c], 1)).join('; ');
}

function csvCell_(v, tz) {
  let s;
  if (v instanceof Date) s = Utilities.formatDate(v, tz || 'UTC', 'yyyy-MM-dd');
  else s = v === null || v === undefined ? '' : String(v);
  return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function digest_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s));
}

function blank_(v) {
  return v === undefined || v === null || String(v).trim() === '';
}

function round_(v, d) {
  if (typeof v !== 'number' || !isFinite(v)) return '';
  const p = Math.pow(10, d);
  return Math.round(v * p) / p;
}
