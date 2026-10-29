'use strict';
// Run: npm test   (Node 18+, no dependencies).
// Loads Code.gs in a vm with stand-ins for UrlFetchApp, MailApp, Utilities, PropertiesService, Session, ScriptApp.
// The Airtable responses are hand-written in the shape the Airtable REST API documents.
const test = require('node:test');
const assert = require('node:assert');
const { load, page } = require('./harness');

const BOM = '\uFEFF';
const BASE = { BASE_ID: 'appTEST1234567890', TABLE: 'Orders', VIEW: 'Open orders', RECIPIENTS: 'ops@example.com' };
const plain = (x) => JSON.parse(JSON.stringify(x));
const err = (code, type, message) => ({ code, body: { error: { type, message } } });

test('pagination: follows offset across 3 pages and sends all records in one CSV', () => {
  const pages = [
    page([{ Name: 'r1' }, { Name: 'r2' }], 'itr1/rec2'),
    page([{ Name: 'r3' }, { Name: 'r4' }], 'itr2/rec4'),
    page([{ Name: 'r5' }]),
  ];
  const { ctx, calls } = load({ config: BASE, fetchImpl: (url, o, n) => ({ code: 200, body: pages[n - 1] }) });
  const result = ctx.sendViewAsCsv();
  assert.strictEqual(calls.fetch.length, 3);
  assert.ok(!calls.fetch[0].url.includes('offset='));
  assert.ok(calls.fetch[1].url.endsWith('&offset=itr1%2Frec2'));
  assert.ok(calls.fetch[2].url.endsWith('&offset=itr2%2Frec4'));
  assert.strictEqual(result.records, 5);
  assert.strictEqual(calls.mail[0].attachments[0].data, BOM + 'Name\r\nr1\r\nr2\r\nr3\r\nr4\r\nr5\r\n');
  assert.deepStrictEqual(calls.sleep, [250, 250], 'short pause between pages only');
});

test('requests carry the bearer token, GET, and the view/cellFormat/timeZone/userLocale parameters', () => {
  const { ctx, calls } = load({ config: { ...BASE, TIMEZONE: 'Europe/Berlin', LOCALE: 'de' }, fetchImpl: () => ({ code: 200, body: page([{ A: 1 }]) }) });
  ctx.sendViewAsCsv();
  const { url, options } = calls.fetch[0];
  assert.strictEqual(options.headers.Authorization, 'Bearer test-token');
  assert.strictEqual(options.method, 'get');
  assert.strictEqual(options.muteHttpExceptions, true);
  assert.ok(url.startsWith('https://api.airtable.com/v0/appTEST1234567890/Orders?pageSize=100&view=Open%20orders'));
  assert.ok(url.includes('cellFormat=string&timeZone=Europe%2FBerlin&userLocale=de'));
});

test('429 is retried with backoff, then the run succeeds', () => {
  const { ctx, calls } = load({
    config: BASE,
    fetchImpl: (url, o, n) => (n <= 2 ? err(429, 'RATE_LIMIT_REACHED', 'Rate limit') : { code: 200, body: page([{ A: 'x' }]) }),
  });
  const result = ctx.sendViewAsCsv();
  assert.strictEqual(result.sent, true);
  assert.strictEqual(calls.fetch.length, 3);
  assert.deepStrictEqual(calls.sleep, [2000, 4000]);
  assert.strictEqual(calls.mail.length, 1);
});

test('429 honours Retry-After', () => {
  const { ctx, calls } = load({
    config: BASE,
    fetchImpl: (url, o, n) => (n === 1 ? { ...err(429, 'RATE_LIMIT_REACHED', 'Rate limit'), headers: { 'Retry-After': '9' } } : { code: 200, body: page([{ A: 'x' }]) }),
  });
  ctx.sendViewAsCsv();
  assert.deepStrictEqual(calls.sleep, [9000]);
});

test('5xx is retried too', () => {
  const { ctx, calls } = load({
    config: BASE,
    fetchImpl: (url, o, n) => (n === 1 ? { code: 503, body: 'Service Unavailable' } : { code: 200, body: page([{ A: 'x' }]) }),
  });
  assert.strictEqual(ctx.sendViewAsCsv().sent, true);
  assert.strictEqual(calls.fetch.length, 2);
});

test('persistent 429 gives up after 5 tries, sends nothing and says so', () => {
  const { ctx, calls } = load({ config: BASE, fetchImpl: () => err(429, 'RATE_LIMIT_REACHED', 'Rate limit reached') });
  assert.throws(() => ctx.sendViewAsCsv(), /HTTP 429: Rate limit reached/);
  assert.strictEqual(calls.fetch.length, 5);
  assert.deepStrictEqual(calls.sleep, [2000, 4000, 8000, 16000]);
  assert.strictEqual(calls.mail.length, 0);
});

test('401: clear message, no retry, no email, token not echoed', () => {
  const { ctx, calls } = load({ config: BASE, token: 'super-secret-token', fetchImpl: () => err(401, 'AUTHENTICATION_REQUIRED', 'Authentication required') });
  let message = '';
  try { ctx.sendViewAsCsv(); } catch (e) { message = e.message; }
  assert.match(message, /HTTP 401: Authentication required\. Check the AIRTABLE_TOKEN script property\./);
  assert.ok(!message.includes('super-secret-token'));
  assert.strictEqual(calls.fetch.length, 1);
  assert.strictEqual(calls.mail.length, 0);
});

test('403 and 404: clear messages, no retry', () => {
  for (const [code, pattern] of [[403, /HTTP 403.*data\.records:read/], [404, /HTTP 404.*BASE_ID, TABLE and VIEW/]]) {
    const { ctx, calls } = load({ config: BASE, fetchImpl: () => err(code, 'X', 'denied or missing') });
    assert.throws(() => ctx.sendViewAsCsv(), pattern);
    assert.strictEqual(calls.fetch.length, 1);
  }
});

test('missing AIRTABLE_TOKEN script property: error before any request', () => {
  const { ctx, calls } = load({ config: BASE, token: null, fetchImpl: () => { throw new Error('should not be called'); } });
  assert.throws(() => ctx.sendViewAsCsv(), /AIRTABLE_TOKEN is missing/);
  assert.strictEqual(calls.fetch.length, 0);
});

test('placeholder BASE_ID: error before any request', () => {
  const { ctx, calls } = load({ fetchImpl: () => { throw new Error('should not be called'); } });
  assert.throws(() => ctx.sendViewAsCsv(), /Set BASE_ID/);
  assert.strictEqual(calls.fetch.length, 0);
});

test('CONFIG.AIRTABLE_TOKEN reads the script property and the source has no token literal', () => {
  const { ctx } = load({ token: 'from-properties', fetchImpl: () => ({ code: 200, body: page([]) }) });
  assert.strictEqual(ctx.CONFIG.AIRTABLE_TOKEN, 'from-properties');
  const source = require('fs').readFileSync(require('path').join(__dirname, '..', 'Code.gs'), 'utf8');
  assert.ok(!/pat[A-Za-z0-9]{10,}\./.test(source));
});

test('MailApp.sendEmail gets the CSV as a text/csv blob attachment with BOM, CRLF and the templated name', () => {
  const records = [{ Name: 'Zoë', Note: 'a, b' }, { Name: 'Ann "A"', Note: 'x\ny' }];
  const { ctx, calls } = load({ config: BASE, now: '2026-10-29T08:00:00Z', fetchImpl: () => ({ code: 200, body: page(records) }) });
  const result = ctx.sendViewAsCsv();
  assert.strictEqual(calls.mail.length, 1);
  const mail = calls.mail[0];
  assert.strictEqual(mail.to, 'ops@example.com');
  assert.strictEqual(mail.subject, 'Airtable export: Open orders (2 records) - 2026-10-29');
  assert.match(mail.body, /Attached: 2 records from the Airtable view "Open orders"/);
  assert.strictEqual(mail.attachments.length, 1);
  const blob = mail.attachments[0];
  assert.strictEqual(blob.name, 'airtable-export-2026-10-29.csv');
  assert.strictEqual(blob.contentType, 'text/csv');
  assert.strictEqual(blob.getDataAsString(), BOM + 'Name,Note\r\nZoë,"a, b"\r\n"Ann ""A""","x\ny"\r\n');
  assert.deepStrictEqual(blob.getBytes().slice(0, 3).map((b) => b & 0xFF), [0xEF, 0xBB, 0xBF]);
  assert.deepStrictEqual(plain(result), { sent: true, records: 2, bytes: blob.getBytes().length, filename: 'airtable-export-2026-10-29.csv' });
});

test('SUBJECT and FILENAME templates: {date} and {count}, in the configured time zone', () => {
  const { ctx, calls } = load({
    config: { ...BASE, SUBJECT: '{count} orders on {date}', FILENAME: 'orders_{date}', TIMEZONE: 'Asia/Tokyo' },
    now: '2026-10-29T23:30:00Z',   // already 30 October in Tokyo
    fetchImpl: () => ({ code: 200, body: page([{ A: 1 }, { A: 2 }, { A: 3 }]) }),
  });
  ctx.sendViewAsCsv();
  assert.strictEqual(calls.mail[0].subject, '3 orders on 2026-10-30');
  assert.strictEqual(calls.mail[0].attachments[0].name, 'orders_2026-10-30.csv');
});

test('empty TIMEZONE falls back to the script time zone', () => {
  const { ctx, calls } = load({
    config: { ...BASE, TIMEZONE: '' }, scriptTimeZone: 'Pacific/Auckland', now: '2026-10-29T12:30:00Z',  // 30 Oct 01:30 there
    fetchImpl: () => ({ code: 200, body: page([{ A: 1 }]) }),
  });
  ctx.sendViewAsCsv();
  assert.ok(calls.fetch[0].url.includes('timeZone=Pacific%2FAuckland'));
  assert.strictEqual(calls.mail[0].attachments[0].name, 'airtable-export-2026-10-30.csv');
});

test('multiple recipients are trimmed and joined; empty RECIPIENTS falls back to the script owner', () => {
  let { ctx, calls } = load({ config: { ...BASE, RECIPIENTS: ' a@example.com , b@example.com,' }, fetchImpl: () => ({ code: 200, body: page([{ A: 1 }]) }) });
  ctx.sendViewAsCsv();
  assert.strictEqual(calls.mail[0].to, 'a@example.com,b@example.com');
  ({ ctx, calls } = load({ config: { ...BASE, RECIPIENTS: '' }, userEmail: 'owner@example.com', fetchImpl: () => ({ code: 200, body: page([{ A: 1 }]) }) }));
  ctx.sendViewAsCsv();
  assert.strictEqual(calls.mail[0].to, 'owner@example.com');
});

test('SKIP_IF_EMPTY true and 0 records: nothing is sent, and it is logged', () => {
  const { ctx, calls } = load({ config: { ...BASE, SKIP_IF_EMPTY: true }, fetchImpl: () => ({ code: 200, body: page([]) }) });
  const result = ctx.sendViewAsCsv();
  assert.deepStrictEqual(plain(result), { sent: false, records: 0 });
  assert.strictEqual(calls.mail.length, 0);
  assert.ok(calls.logs.some((l) => /No email sent/.test(l)));
});

test('SKIP_IF_EMPTY false and 0 records: sends a header-only CSV when FIELDS is set', () => {
  const { ctx, calls } = load({ config: { ...BASE, SKIP_IF_EMPTY: false, FIELDS: ['Name', 'Status'] }, fetchImpl: () => ({ code: 200, body: page([]) }) });
  assert.strictEqual(ctx.sendViewAsCsv().sent, true);
  assert.strictEqual(calls.mail[0].attachments[0].data, BOM + 'Name,Status\r\n');
  assert.match(calls.mail[0].subject, /\(0 records\)/);
  assert.match(calls.mail[0].body, /Attached: 0 records/);
});

test('FIELDS restricts the request and fixes the column order', () => {
  const { ctx, calls } = load({
    config: { ...BASE, FIELDS: ['Status', 'Name'] },
    fetchImpl: () => ({ code: 200, body: page([{ Name: 'A', Status: 'Open', Hidden: 'x' }]) }),
  });
  ctx.sendViewAsCsv();
  assert.ok(calls.fetch[0].url.includes('fields%5B%5D=Status&fields%5B%5D=Name'));
  assert.strictEqual(calls.mail[0].attachments[0].data, BOM + 'Status,Name\r\nOpen,A\r\n');
});

test('an attachment over the size limit is refused with a clear error', () => {
  const { ctx, calls } = load({ config: BASE, tunables: { MAX_ATTACHMENT_BYTES: 20 }, fetchImpl: () => ({ code: 200, body: page([{ Name: 'a fairly long value here' }]) }) });
  assert.throws(() => ctx.sendViewAsCsv(), /over the ~25 MB email limit/);
  assert.strictEqual(calls.mail.length, 0);
});

test('the run stops with a clear error when the runtime budget is used up', () => {
  const { ctx, calls } = load({ config: BASE, tunables: { MAX_RUNTIME_MS: -1 }, fetchImpl: () => ({ code: 200, body: page([{ A: 1 }]) }) });
  assert.throws(() => ctx.sendViewAsCsv(), /Nothing was emailed/);
  assert.strictEqual(calls.mail.length, 0);
});

test('installDailyTrigger(7): one time-based daily trigger for sendViewAsCsv at 7', () => {
  const { ctx, calls } = load();
  ctx.installDailyTrigger(7);
  assert.strictEqual(calls.triggers.length, 1);
  assert.strictEqual(calls.triggers[0].fn, 'sendViewAsCsv');
  assert.deepStrictEqual(plain(calls.triggers[0].steps), [['timeBased'], ['everyDays', 1], ['atHour', 7]]);
});

test('installDailyTrigger: default hour 7, bad hour rejected, TIMEZONE applied', () => {
  let { ctx, calls } = load();
  ctx.installDailyTrigger();
  assert.deepStrictEqual(plain(calls.triggers[0].steps).slice(-1), [['atHour', 7]]);
  assert.throws(() => ctx.installDailyTrigger(25), /0 to 23/);
  ({ ctx, calls } = load({ config: { TIMEZONE: 'Europe/Amsterdam' } }));
  ctx.installDailyTrigger(6);
  assert.deepStrictEqual(plain(calls.triggers[0].steps).slice(-1), [['inTimezone', 'Europe/Amsterdam']]);
});

test('installWeeklyTrigger("monday", 8): weekly trigger on Monday at 8', () => {
  const { ctx, calls } = load();
  ctx.installWeeklyTrigger('monday', 8);
  assert.deepStrictEqual(plain(calls.triggers[0].steps), [['timeBased'], ['onWeekDay', 'MONDAY'], ['atHour', 8]]);
  ctx.installWeeklyTrigger(5, 17);
  assert.deepStrictEqual(plain(calls.triggers[1].steps), [['timeBased'], ['onWeekDay', 'FRIDAY'], ['atHour', 17]]);
  assert.throws(() => ctx.installWeeklyTrigger('someday', 8), /Unknown weekday/);
});

test('installing a trigger replaces the old one; removeTriggers leaves unrelated triggers alone', () => {
  const { ctx, calls, existingTriggers } = load();
  existingTriggers.push({ getHandlerFunction: () => 'somethingElse' });
  ctx.installDailyTrigger(7);
  ctx.installWeeklyTrigger('friday', 9);
  const mine = existingTriggers.filter((t) => t.getHandlerFunction() === 'sendViewAsCsv');
  assert.strictEqual(mine.length, 1, 'no duplicates');
  assert.deepStrictEqual(plain(mine[0].spec.steps)[1], ['onWeekDay', 'FRIDAY']);
  assert.strictEqual(calls.deleted.length, 1);
  ctx.removeTriggers();
  assert.strictEqual(existingTriggers.length, 1);
  assert.strictEqual(existingTriggers[0].getHandlerFunction(), 'somethingElse');
});
