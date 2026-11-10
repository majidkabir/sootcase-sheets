// Tests for src/Code.js against a running Sootcase server (see harness.js).
// No routing servers are needed: road and rail legs give distance_km, air uses great-circle.
'use strict';

const assert = require('assert');
const { makeEnv, SERVER } = require('./harness');

const tests = [];
const test = (name, fn) => tests.push({ name, fn });
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 1e-6, (msg || '') + ': ' + a + ' != ' + b);

test('road leg by place names and a given distance', () => {
  const { api } = makeEnv();
  // GLEC v3.3 European HGV starting point: 125 g CO2e/t-km WTW. 12 t x 332 km x 125 g = 498 kg.
  near(api.SOOTCASE('Copenhagen', 'Hamburg', 12, 'road', '', 332), 498, 'wtw');
  near(api.SOOTCASE('Copenhagen', 'Hamburg', 12, 'truck', 'hgv', 332, 'ttw'), 382.464, 'ttw');
  near(api.SOOTCASE('Copenhagen', 'Hamburg', 12, 'road', '', 332, 'wtt'), 115.536, 'wtt');
  near(api.SOOTCASE('Copenhagen', 'Hamburg', 12, 'road', '', 332, 'tkm'), 3984, 'tonne-km');
  near(api.SOOTCASE('Copenhagen', 'Hamburg', 12, 'road', '', 332, 'intensity'), 125, 'intensity');
  near(api.SOOTCASE('Copenhagen', 'Hamburg', 12, 'road', '', 332, 'distance'), 332, 'distance');
});

test('air leg by IATA codes matches the API', () => {
  const { api } = makeEnv();
  const got = api.SOOTCASE('CPH', 'FRA', 0.8, 'air');
  const raw = require('child_process').execFileSync('curl', ['-s', '-X', 'POST', SERVER + '/v1/calculate',
    '--data', JSON.stringify({ weight_kg: 800, legs: [{ mode: 'air', vehicle_type: 'air_freight', from: { code: 'CPH' }, to: { code: 'FRA' } }] })]);
  near(got, JSON.parse(raw).wtw_kg, 'wtw');
  assert.ok(got > 0);
});

test('ranges: one request, empty rows stay empty, bad rows show their error', () => {
  const env = makeEnv();
  const col = (...v) => v.map((x) => [x]);
  const out = env.api.SOOTCASE(
    col('Copenhagen', '', 'CPH', 'CPH'),
    col('Hamburg', '', 'FRA', 'FRA'),
    col(12, '', 0.8, 0),
    'road',
    col('hgv', '', 'air_freight', ''),
    col(332, '', '', ''),
  );
  assert.strictEqual(out.length, 4);
  near(out[0][0], 498, 'row 1');
  assert.strictEqual(out[1][0], '');
  assert.ok(/Error/.test(out[2][0]), 'row 3 is air_freight on a road leg: ' + out[2][0]);
  assert.ok(/tonnes/.test(out[3][0]), 'row 4 has 0 tonnes: ' + out[3][0]);
  const batches = env.calls.filter((c) => /\/v1\/calculate\/batch$/.test(c.url));
  assert.strictEqual(batches.length, 1, 'one batch request for the whole range');
});

test('a large range goes out in 500-row chunks, every row answered in order', () => {
  const env = makeEnv();
  const n = 1201;
  const col = (f) => Array.from({ length: n }, (_, i) => [f(i)]);
  const out = env.api.SOOTCASE(col(() => 'CPH'), col(() => 'FRA'), col((i) => 1 + i), 'road', 'hgv', col(() => 100));
  assert.strictEqual(out.length, n);
  out.forEach((r, i) => near(r[0], (1 + i) * 100 * 0.125, 'row ' + i)); // 125 g/tkm
  assert.strictEqual(env.calls.filter((c) => /\/batch$/.test(c.url)).length, 3);
});

test('one failing request fails only its own rows', () => {
  const env = makeEnv();
  const fetchAll = env.sandbox.UrlFetchApp.fetchAll;
  let n = 0;
  env.sandbox.UrlFetchApp.fetchAll = (reqs) => fetchAll(reqs).map((r) => (/batch$/.test(reqs[0].url) && n++ === 0
    ? { getResponseCode: () => 429, getContentText: () => '{"error":"100000 shipments a day"}' } : r));
  const col = (f) => Array.from({ length: 600 }, (_, i) => [f(i)]);
  const out = env.api.SOOTCASE(col(() => 'CPH'), col(() => 'FRA'), col(() => 1), 'road', 'hgv', col(() => 100));
  assert.ok(/quota reached/.test(out[0][0]), out[0][0]);
  assert.ok(/quota reached/.test(out[499][0]));
  near(out[500][0], 12.5, 'second chunk answered');
});

test('cached: the same cell again makes no requests', () => {
  const env = makeEnv();
  env.api.SOOTCASE('CPH', 'FRA', 1, 'air');
  const n = env.calls.length;
  env.api.SOOTCASE('CPH', 'FRA', 1, 'air');
  assert.strictEqual(env.calls.length, n);
});

test('single cell errors are thrown', () => {
  const { api } = makeEnv();
  assert.throws(() => api.SOOTCASE('CPH', 'FRA', 1, 'bicycle'), /unknown mode/);
  assert.throws(() => api.SOOTCASE('', 'FRA', 1, 'air'), /from is required/);
  assert.throws(() => api.SOOTCASE('CPH', 'FRA', 1, 'air', '', '', 'co2'), /unknown value/);
  assert.throws(() => api.SOOTCASE('Xqzvkplm Nowhere', 'FRA', 1, 'air'), /no place found/);
});

test('place inputs', () => {
  const { api } = makeEnv();
  assert.deepStrictEqual(JSON.parse(JSON.stringify(api.placeInput_('NLRTM', 'from'))), { code: 'NLRTM' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(api.placeInput_('CPH', 'from'))), { code: 'CPH' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(api.placeInput_('Paris', 'from'))), { lookup: 'Paris' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(api.placeInput_('cph', 'from'))), { lookup: 'cph' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(api.placeInput_('55.67, 12.56', 'from'))), { lat: 55.67, lon: 12.56 });
  assert.throws(() => api.placeInput_('95,10', 'from'), /not a valid/);
});

test('API key is sent as a bearer token', () => {
  const env = makeEnv();
  env.docProps.setProperty('SOOTCASE_API_KEY', 'doc-key');
  env.api.SOOTCASE('CPH', 'FRA', 2, 'air');
  assert.strictEqual(env.calls[env.calls.length - 1].headers.Authorization, 'Bearer doc-key');
  env.userProps.setProperty('SOOTCASE_API_KEY', 'my-key');
  env.api.SOOTCASE('CPH', 'FRA', 3, 'air');
  assert.strictEqual(env.calls[env.calls.length - 1].headers.Authorization, 'Bearer my-key', 'own key first');
});

test('calculate a sheet of shipments', () => {
  const env = makeEnv();
  env.addSheet('Q1', [
    ['shipment_id', 'weight_kg', 'leg_id', 'mode', 'vehicle_type', 'distance_km', 'from_code', 'to_code', 'hub_type'],
    ['A1', 12000, 'l1', 'road', 'hgv', 332, '', '', ''],
    ['A1', '', 'l2', 'sea', 'container_ship', 560, '', '', ''],
    ['A2', 3500, 'l1', 'air', 'air_freight', '', 'CPH', 'FRA', ''],
    ['A3', 20000, 'l1', 'rail', 'freight_train', 330, '', '', ''],
    ['A3', '', 'h1', '', '', '', '', '', 'warehouse'],
    ['', '', '', '', '', '', '', '', ''],
    ['BAD', 'abc', 'l1', 'road', 'hgv', 10, '', '', ''],
  ]);
  env.api.calculateSheet();
  assert.deepStrictEqual(env.alerts, []);

  const res = env.sheet('Sootcase results').values;
  assert.strictEqual(res.length, 5, 'header + 4 shipments');
  const h = res[0];
  const row = (id) => res.find((r) => r[0] === id);
  const v = (id, col) => row(id)[h.indexOf(col)];
  near(v('A1', 'wtw_kg'), 547.325, 'A1 = 498 road + 49.3248 sea');
  assert.strictEqual(v('A1', 'legs'), 2);
  near(v('A3', 'hub_wtw_kg'), 710, 'A3 warehouse, 20 t x 35.5 kg/t');
  assert.ok(/not a number/.test(v('BAD', 'error')));
  assert.ok(/line 8\b/.test(v('BAD', 'error')), 'line number is the sheet row: ' + v('BAD', 'error'));
  assert.ok(env.sheet('Sootcase results').textCols.indexOf(1) >= 0, 'shipment_id column is plain text');

  const legs = env.sheet('Sootcase legs').values;
  assert.strictEqual(legs.length, 1 + 4 + 1, 'header + 4 legs + 1 hub');
  assert.ok(legs.slice(1).every((r) => r.length === legs[0].length));
  assert.ok(legs.some((r) => r[2] === 'hub' && r[3] === 'warehouse'));

  const sum = env.sheet('Sootcase summary').values;
  const total = sum.find((r) => r[0] === 'Well-to-wheel (all)');
  const expected = 498 + 49.3248 + v('A2', 'wtw_kg') + v('A3', 'wtw_kg');
  assert.ok(Math.abs(total[1] - expected) < 0.2, 'summary total ' + total[1] + ' vs ' + expected);
  assert.ok(sum.some((r) => r[0] === 'air'), 'per mode rows');
  assert.ok(/1 failed/.test(env.toasts[env.toasts.length - 1]));
  assert.strictEqual(res[0].indexOf('warnings'), res[0].length - 2, 'warnings column before error');
  assert.ok(legs[0].indexOf('country_km') > 0);
});

test('a failed sheet calculation says why and clears the progress toast', () => {
  const env = makeEnv();
  env.userProps.setProperty('SOOTCASE_URL', 'http://localhost:1'); // nothing listens
  env.addSheet('Q3', [['shipment_id', 'weight_kg', 'mode', 'vehicle_type', 'distance_km'], ['A', 1000, 'road', 'hgv', 10]]);
  env.api.calculateSheet(); // the curl stand-in throws on a refused connection, as UrlFetchApp does
  assert.strictEqual(env.alerts.length, 1, 'one alert with the reason');
  assert.strictEqual(env.toasts[env.toasts.length - 1], 'Not calculated.');
  assert.strictEqual(env.sheet('Sootcase results'), null);
});

test('the results sheet is shown after a calculation', () => {
  const env = makeEnv();
  env.addSheet('Q4', [['shipment_id', 'weight_kg', 'mode', 'vehicle_type', 'distance_km'], ['A', 1000, 'road', 'hgv', 10]]);
  env.api.calculateSheet();
  assert.strictEqual(env.active().getName(), 'Sootcase results');
});

test('a row without shipment_id is refused with its row number', () => {
  const env = makeEnv();
  env.addSheet('Q2', [
    ['shipment_id', 'weight_kg', 'mode', 'vehicle_type', 'distance_km'],
    ['A', 1000, 'road', 'hgv', 10],
    ['', '', 'rail', 'freight_train', 20],
  ]);
  env.api.calculateSheet();
  assert.ok(/Row 3 has no shipment_id/.test(env.alerts[0]), env.alerts[0]);
  assert.strictEqual(env.sheet('Sootcase results'), null);
});

test('a sheet without Sootcase columns is refused', () => {
  const env = makeEnv();
  env.addSheet('Other', [['name', 'weight'], ['x', 1]]);
  env.api.calculateSheet();
  assert.ok(/shipment_id/.test(env.alerts[0]));
  assert.strictEqual(env.sheet('Sootcase results'), null);
});

test('CSV cells: dates in the spreadsheet time zone, quoting', () => {
  const env = makeEnv();
  // Midnight 1 March in Copenhagen is 28 February in UTC.
  const d = env.run('new Date("2026-02-28T23:00:00Z")');
  assert.strictEqual(env.api.csvCell_(d, 'Europe/Copenhagen'), '2026-03-01');
  assert.strictEqual(env.api.csvCell_('a "b", c'), '"a ""b"", c"');
  assert.strictEqual(env.api.csvCell_(null), '');
  assert.strictEqual(env.api.csvCell_(0.5), '0.5');
});

test('km per country, largest first', () => {
  const { api } = makeEnv();
  assert.strictEqual(api.countryKm_({ DK: 35.14, DE: 210.44 }), 'DE 210.4; DK 35.1');
  assert.strictEqual(api.countryKm_(undefined), '');
});

test('a place found by name is sent as its code', () => {
  const env = makeEnv();
  env.api.SOOTCASE('Hamburg', 'Rotterdam', 1, 'sea', '', 500);
  const batch = env.calls.find((c) => /\/batch$/.test(c.url));
  assert.ok(batch, 'batch call');
  assert.ok(env.cache.data && Object.values(env.cache.data).some((v) => v === '{"code":"DEHAM"}'), 'cached as a code');
});

test('server URL must start with http(s)://; the API key is saved and removed', () => {
  const env = makeEnv();
  env.answers.push('sootcase.org');
  env.api.setServer();
  assert.ok(/isn't a valid value/.test(env.alerts[0]));
  assert.strictEqual(env.userProps.getProperty('SOOTCASE_URL'), require('./harness').SERVER, 'unchanged');
  env.answers.push('k-123', '');
  env.api.setApiKey();
  assert.strictEqual(env.userProps.getProperty('SOOTCASE_API_KEY'), 'k-123');
  env.api.setApiKey();
  assert.strictEqual(env.userProps.getProperty('SOOTCASE_API_KEY'), null);
});

test('template sheet', () => {
  const env = makeEnv();
  env.api.insertTemplate();
  const s = env.sheet('Shipments');
  assert.strictEqual(JSON.stringify(s.values), JSON.stringify(env.api.TEMPLATE_ROWS));
  env.api.insertTemplate();
  assert.ok(env.sheet('Shipments 2'), 'a second template gets a new name');
});

let failed = 0;
for (const t of tests) {
  try {
    t.fn();
    console.log('ok   ' + t.name);
  } catch (e) {
    failed++;
    console.log('FAIL ' + t.name + '\n     ' + (e.stack || e).toString().split('\n').slice(0, 3).join('\n     '));
  }
}
console.log(failed ? failed + ' failed' : 'all ' + tests.length + ' passed');
process.exit(failed ? 1 : 0);
