// Runs src/Code.js in Node with small stand-ins for the Apps Script services it uses.
// HTTP goes to a real Sootcase server (SOOTCASE_TEST_URL, default http://localhost:8080) through curl,
// because UrlFetchApp is synchronous.
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const SERVER = (process.env.SOOTCASE_TEST_URL || 'http://localhost:8080').replace(/\/+$/, '');

function response(raw) {
  // curl -w appends the status code on its own last line.
  const i = raw.lastIndexOf('\n');
  return {
    getResponseCode: () => Number(raw.slice(i + 1)),
    getContentText: () => raw.slice(0, i),
  };
}

function makeEnv() {
  const calls = [];
  function fetch(url, opts) {
    opts = opts || {};
    calls.push({ url, method: opts.method || 'get', headers: opts.headers || {} });
    const args = ['-s', '-m', '60', '-w', '\n%{http_code}', '-X', (opts.method || 'get').toUpperCase()];
    Object.entries(opts.headers || {}).forEach(([k, v]) => args.push('-H', k + ': ' + v));
    if (opts.payload !== undefined) {
      args.push('-H', 'Content-Type: ' + (opts.contentType || 'application/x-www-form-urlencoded'));
      args.push('--data-binary', '@-');
    }
    args.push(url);
    const raw = execFileSync('curl', args, { input: opts.payload || '', maxBuffer: 1 << 28 }).toString();
    return response(raw);
  }

  const store = (data) => ({
    getProperty: (k) => (k in data ? data[k] : null),
    setProperty: (k, v) => { data[k] = v; },
    deleteProperty: (k) => { delete data[k]; },
    data,
  });
  const userProps = store({ SOOTCASE_URL: SERVER });
  const docProps = store({});

  const cacheData = {};
  const cache = {
    get: (k) => (k in cacheData ? cacheData[k] : null),
    getAll: (keys) => { const o = {}; keys.forEach((k) => { if (k in cacheData) o[k] = cacheData[k]; }); return o; },
    putAll: (o) => Object.assign(cacheData, o),
    data: cacheData,
  };

  // A spreadsheet: sheets of 2-D arrays.
  function makeSheet(name, values) {
    const s = {
      name,
      values: values || [],
      textCols: [],
      getName: () => s.name,
      getDataRange: () => ({ getValues: () => s.values.map((r) => r.slice()) }),
      clear: () => { s.values = []; },
      getRange: (row, col, nr, nc) => ({
        setValues: (v) => {
          if (v.length !== nr || v.some((r) => r.length !== nc)) {
            throw new Error('setValues: data is ' + v.length + 'x' + (v[0] || []).length +
              ', range is ' + nr + 'x' + nc + ' on ' + s.name);
          }
          v.forEach((r, i) => { s.values[row - 1 + i] = r.slice(); });
        },
        setFontWeight: () => {},
        setNumberFormat: (f) => { if (f === '@') s.textCols.push(col); },
      }),
      setFrozenRows: () => {},
      autoResizeColumns: () => {},
    };
    return s;
  }
  const sheets = [];
  let active = null;
  const toasts = [];
  const answers = [];
  const alerts = [];
  const ss = {
    getActiveSheet: () => active,
    getSheetByName: (n) => sheets.find((s) => s.name === n) || null,
    insertSheet: (n) => { const s = makeSheet(n); sheets.push(s); active = s; return s; },
    toast: (msg) => toasts.push(msg),
    getSpreadsheetTimeZone: () => 'Europe/Copenhagen',
  };
  const ui = {
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
    Button: { OK: 'OK' },
    alert: (title, msg) => alerts.push(msg || title),
    // Answers prompts from env.answers (button OK) in order.
    prompt: () => ({ getSelectedButton: () => 'OK', getResponseText: () => answers.shift() }),
  };

  const sandbox = {
    console,
    UrlFetchApp: { fetch, fetchAll: (reqs) => reqs.map((r) => fetch(r.url, r)) },
    PropertiesService: { getUserProperties: () => userProps, getDocumentProperties: () => docProps },
    CacheService: { getScriptCache: () => cache },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      computeDigest: (alg, s) => Array.from(crypto.createHash(alg).update(s).digest()),
      base64Encode: (bytes) => Buffer.from(bytes.map((b) => b & 255)).toString('base64'),
      formatDate: (d, tz) => new Intl.DateTimeFormat('en-CA', { timeZone: tz }).format(d),
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => ui },
  };
  vm.createContext(sandbox);
  const code = fs.readFileSync(path.join(__dirname, '..', 'src', 'Code.js'), 'utf8');
  // Top-level const/function declarations become reachable through this object.
  vm.runInContext(code + '\n;globalThis.__api = { SOOTCASE, calculateSheet, insertTemplate, setServer, setApiKey, TEMPLATE_ROWS, placeInput_, legInput_, csvCell_, countryKm_ };', sandbox);

  return {
    api: sandbox.__api,
    run: (js) => vm.runInContext(js, sandbox),
    calls,
    cache,
    userProps,
    docProps,
    sheets,
    toasts,
    alerts,
    answers,
    addSheet: (name, values) => { const s = makeSheet(name, values); sheets.push(s); active = s; return s; },
    sheet: (n) => ss.getSheetByName(n),
  };
}

module.exports = { makeEnv, SERVER };
