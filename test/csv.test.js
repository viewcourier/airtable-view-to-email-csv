'use strict';
// Run: npm test   (Node 18+, no dependencies). Pure helpers of Code.gs: CSV, URL, templates, retry decision.
const test = require('node:test');
const assert = require('node:assert');
const { load } = require('./harness');

const BOM = '\uFEFF';
const { ctx } = load();
const plain = (x) => JSON.parse(JSON.stringify(x));

test('csvEscape_: plain text is left alone', () => {
  assert.strictEqual(ctx.csvEscape_('hello world'), 'hello world');
});

test('csvEscape_: comma forces quotes', () => {
  assert.strictEqual(ctx.csvEscape_('a,b'), '"a,b"');
});

test('csvEscape_: inner quotes are doubled', () => {
  assert.strictEqual(ctx.csvEscape_('say "hi"'), '"say ""hi"""');
});

test('csvEscape_: LF and CR and CRLF force quotes and are kept', () => {
  assert.strictEqual(ctx.csvEscape_('line1\nline2'), '"line1\nline2"');
  assert.strictEqual(ctx.csvEscape_('line1\rline2'), '"line1\rline2"');
  assert.strictEqual(ctx.csvEscape_('l1\r\nl2'), '"l1\r\nl2"');
});

test('csvEscape_: empty, null and undefined become an empty field', () => {
  assert.strictEqual(ctx.csvEscape_(''), '');
  assert.strictEqual(ctx.csvEscape_(null), '');
  assert.strictEqual(ctx.csvEscape_(undefined), '');
});

test('csvEscape_: arrays are joined with ", " and then quoted', () => {
  assert.strictEqual(ctx.csvEscape_(['Red', 'Blue']), '"Red, Blue"');
  assert.strictEqual(ctx.csvEscape_(['solo']), 'solo');
  assert.strictEqual(ctx.csvEscape_([]), '');
  assert.strictEqual(ctx.csvEscape_(['a', null, '', 'b']), '"a, b"');
  assert.strictEqual(ctx.csvEscape_([1, 2.5]), '"1, 2.5"');
});

test('csvEscape_: numbers and booleans, including 0 and false', () => {
  assert.strictEqual(ctx.csvEscape_(42), '42');
  assert.strictEqual(ctx.csvEscape_(-3.5), '-3.5');
  assert.strictEqual(ctx.csvEscape_(0), '0');
  assert.strictEqual(ctx.csvEscape_(true), 'true');
  assert.strictEqual(ctx.csvEscape_(false), 'false');
});

test('csvEscape_: unicode is passed through unchanged', () => {
  assert.strictEqual(ctx.csvEscape_('Zoë Müller – 東京 ☕'), 'Zoë Müller – 東京 ☕');
  assert.strictEqual(ctx.csvEscape_('café, crème'), '"café, crème"');
});

test('csvEscape_: attachment-like objects become their URL, other objects JSON', () => {
  assert.strictEqual(ctx.csvEscape_({ url: 'https://example.com/a.pdf', filename: 'a.pdf' }), 'https://example.com/a.pdf');
  assert.strictEqual(ctx.csvEscape_({ a: 1 }), '"{""a"":1}"');
});

test('buildCsv_: starts with a UTF-8 BOM and uses CRLF line endings, including the last line', () => {
  const csv = ctx.buildCsv_([{ fields: { Name: 'A', Qty: 1 } }, { fields: { Name: 'B', Qty: 2 } }], []);
  assert.strictEqual(csv, BOM + 'Name,Qty\r\nA,1\r\nB,2\r\n');
  assert.strictEqual(csv.charCodeAt(0), 0xFEFF);
  assert.strictEqual(csv.replace(/\r\n/g, '').includes('\n'), false, 'no bare LF outside quoted fields');
});

test('buildCsv_: BOM encodes to EF BB BF in UTF-8', () => {
  const bytes = Buffer.from(ctx.buildCsv_([{ fields: { N: 'é' } }], []), 'utf8');
  assert.deepStrictEqual([...bytes.slice(0, 3)], [0xEF, 0xBB, 0xBF]);
});

test('buildCsv_: a multi-line value stays in one quoted field', () => {
  const csv = ctx.buildCsv_([{ fields: { Note: 'one\ntwo', Who: 'Ann, "A"' } }], []);
  assert.strictEqual(csv, BOM + 'Note,Who\r\n"one\ntwo","Ann, ""A"""\r\n');
});

test('buildCsv_: header is the union of field names in first-seen order', () => {
  const records = [
    { fields: { Name: 'A', Status: 'Open' } },
    { fields: { Name: 'B', Owner: 'x' } },          // Status empty here, Owner new
    { fields: { Priority: 'high', Name: 'C' } },     // Priority new
  ];
  assert.deepStrictEqual(plain(ctx.buildHeader_(records, [])), ['Name', 'Status', 'Owner', 'Priority']);
  assert.strictEqual(ctx.buildCsv_(records, []),
    BOM + 'Name,Status,Owner,Priority\r\nA,Open,,\r\nB,,x,\r\nC,,,high\r\n');
});

test('buildCsv_: FIELDS decides column order and drops other fields', () => {
  const records = [{ fields: { Name: 'A', Status: 'Open', Extra: 'z' } }, { fields: { Name: 'B' } }];
  assert.strictEqual(ctx.buildCsv_(records, ['Status', 'Name']), BOM + 'Status,Name\r\nOpen,A\r\n,B\r\n');
});

test('buildCsv_: FIELDS with no records gives a header-only file; no FIELDS and no records gives just the BOM', () => {
  assert.strictEqual(ctx.buildCsv_([], ['A', 'B']), BOM + 'A,B\r\n');
  assert.strictEqual(ctx.buildCsv_([], []), BOM);
});

test('buildCsv_: header names are escaped too', () => {
  assert.strictEqual(ctx.buildCsv_([{ fields: { 'Cost, EUR': 1 } }], []), BOM + '"Cost, EUR"\r\n1\r\n');
});

test('buildUrl_: view, pageSize, cellFormat=string, timeZone, userLocale', () => {
  const cfg = { BASE_ID: 'appABC', TABLE: 'tblXYZ', VIEW: 'viwQ1', FIELDS: [], LOCALE: 'en-gb' };
  assert.strictEqual(ctx.buildUrl_(cfg, '', 'Europe/Amsterdam'),
    'https://api.airtable.com/v0/appABC/tblXYZ?pageSize=100&view=viwQ1&cellFormat=string&timeZone=Europe%2FAmsterdam&userLocale=en-gb');
});

test('buildUrl_: encodes names, adds fields[] and offset, omits empty view', () => {
  const cfg = { BASE_ID: 'appABC', TABLE: 'My Table', VIEW: '', FIELDS: ['Name', 'Due date'], LOCALE: '' };
  const url = ctx.buildUrl_(cfg, 'itrX/recY', 'Etc/UTC');
  assert.ok(url.startsWith('https://api.airtable.com/v0/appABC/My%20Table?pageSize=100&cellFormat=string'));
  assert.ok(!url.includes('view='));
  assert.ok(url.includes('userLocale=en'));
  assert.ok(url.includes('fields%5B%5D=Name&fields%5B%5D=Due%20date'));
  assert.ok(url.endsWith('&offset=itrX%2FrecY'));
});

test('applyTemplate_: {date} and {count} are replaced everywhere, unknown ones stay', () => {
  const vars = { date: '2026-10-29', count: 12, view: 'Open orders' };
  assert.strictEqual(ctx.applyTemplate_('Report {date}: {count} rows ({date})', vars), 'Report 2026-10-29: 12 rows (2026-10-29)');
  assert.strictEqual(ctx.applyTemplate_('{count} {other}', vars), '12 {other}');
  assert.strictEqual(ctx.applyTemplate_('{count}', { count: 0 }), '0');
});

test('buildFilename_: templating, forbidden characters and the .csv extension', () => {
  const vars = { date: '2026-10-29', count: 3, view: 'a/b' };
  assert.strictEqual(ctx.buildFilename_('orders-{date}.csv', vars), 'orders-2026-10-29.csv');
  assert.strictEqual(ctx.buildFilename_('orders-{date}', vars), 'orders-2026-10-29.csv');
  assert.strictEqual(ctx.buildFilename_('{view}-{date}.csv', vars), 'a-b-2026-10-29.csv');
  assert.strictEqual(ctx.buildFilename_('', vars), 'airtable-export-2026-10-29.csv');
});

test('parseRecipients_: trims, drops empties', () => {
  assert.deepStrictEqual(plain(ctx.parseRecipients_(' a@x.com, b@y.com ,, ')), ['a@x.com', 'b@y.com']);
  assert.deepStrictEqual(plain(ctx.parseRecipients_('')), []);
  assert.deepStrictEqual(plain(ctx.parseRecipients_(undefined)), []);
});

test('shouldRetry_: 429 and 5xx retry until the last attempt, 4xx never', () => {
  assert.strictEqual(ctx.shouldRetry_(429, 0), true);
  assert.strictEqual(ctx.shouldRetry_(500, 0), true);
  assert.strictEqual(ctx.shouldRetry_(503, 3), true);
  assert.strictEqual(ctx.shouldRetry_(429, 4), false, 'fifth try is the last one');
  for (const code of [400, 401, 403, 404, 422]) assert.strictEqual(ctx.shouldRetry_(code, 0), false);
  assert.strictEqual(ctx.shouldRetry_(200, 0), false);
});

test('backoffMs_: doubling with a 30 s cap, Retry-After wins (capped at 60 s)', () => {
  assert.deepStrictEqual([0, 1, 2, 3, 4, 5].map((a) => ctx.backoffMs_(a, {})), [2000, 4000, 8000, 16000, 30000, 30000]);
  assert.strictEqual(ctx.backoffMs_(0, { 'Retry-After': '7' }), 7000);
  assert.strictEqual(ctx.backoffMs_(0, { 'retry-after': '500' }), 60000);
  assert.strictEqual(ctx.backoffMs_(1, { 'Retry-After': 'soon' }), 4000);
});

test('describeError_: messages for 401, 403, 404', () => {
  const body = (type, message) => JSON.stringify({ error: { type, message } });
  assert.match(ctx.describeError_(401, body('AUTHENTICATION_REQUIRED', 'Authentication required')), /HTTP 401.*AIRTABLE_TOKEN/);
  assert.match(ctx.describeError_(403, body('INVALID_PERMISSIONS_OR_MODEL_NOT_FOUND', 'Invalid permissions, or the requested model was not found.')), /HTTP 403.*data\.records:read/);
  assert.match(ctx.describeError_(404, body('NOT_FOUND', 'Could not find what you are looking for')), /HTTP 404.*BASE_ID, TABLE and VIEW/);
  assert.match(ctx.describeError_(500, 'not json at all'), /HTTP 500: not json at all\./);
});

test('weekdayName_ and checkHour_ validate their input', () => {
  assert.strictEqual(ctx.weekdayName_('monday'), 'MONDAY');
  assert.strictEqual(ctx.weekdayName_('Fri'), 'FRIDAY');
  assert.strictEqual(ctx.weekdayName_(7), 'SUNDAY');
  assert.throws(() => ctx.weekdayName_('funday'), /Unknown weekday/);
  assert.throws(() => ctx.weekdayName_(0), /Unknown weekday/);
  assert.strictEqual(ctx.checkHour_(0), 0);
  assert.strictEqual(ctx.checkHour_(23), 23);
  assert.throws(() => ctx.checkHour_(24), /0 to 23/);
  assert.throws(() => ctx.checkHour_(7.5), /0 to 23/);
  assert.throws(() => ctx.checkHour_('7'), /0 to 23/);
});

test('Code.gs exports its pure helpers when loaded with require()', () => {
  const lib = require('../Code.gs');
  assert.strictEqual(typeof lib.buildCsv_, 'function');
  assert.strictEqual(lib.csvEscape_('a,b'), '"a,b"');
});
