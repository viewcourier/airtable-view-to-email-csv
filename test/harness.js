'use strict';
// Loads Code.gs into a Node vm context with stand-ins for the Apps Script services.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const CODE = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');

/**
 * fetchImpl(url, options, callNumber) -> { code, body, headers }   (body may be an object; it is JSON-encoded)
 * config:   values assigned onto CONFIG, e.g. { BASE_ID: 'app...' }
 * tunables: values assigned to the top-level tuning variables of Code.gs, e.g. { PAGE_SIZE: 3 }
 * now:      the fixed value of `new Date()` inside the script (ISO string)
 * token:    value of the AIRTABLE_TOKEN script property (null = not set)
 */
function load({ fetchImpl, config = {}, tunables = {}, now = '2026-10-29T08:00:00Z', token = 'test-token',
  scriptTimeZone = 'Europe/Amsterdam', userEmail = 'owner@example.com' } = {}) {
  const calls = { fetch: [], sleep: [], mail: [], logs: [], triggers: [], deleted: [] };
  const existingTriggers = [];
  const props = token === null ? {} : { AIRTABLE_TOKEN: token };
  const fixedNow = new Date(now).getTime();

  class FakeDate extends Date {
    constructor(...args) { if (args.length === 0) super(fixedNow); else super(...args); }
    static now() { return fixedNow; }
  }

  const sandbox = {
    console,
    Date: FakeDate,
    Logger: { log: (m) => calls.logs.push(String(m)) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null) }) },
    Session: {
      getScriptTimeZone: () => scriptTimeZone,
      getEffectiveUser: () => ({ getEmail: () => userEmail }),
    },
    Utilities: {
      sleep: (ms) => calls.sleep.push(ms),
      newBlob(data, contentType, name) {
        const bytes = Buffer.from(String(data), 'utf8');
        return {
          contentType, name, data: String(data),
          getBytes: () => Array.from(bytes),
          getName: () => name,
          getContentType: () => contentType,
          getDataAsString: () => bytes.toString('utf8'),
        };
      },
      // Only the 'yyyy-MM-dd' pattern is needed.
      formatDate(date, timeZone, pattern) {
        if (pattern !== 'yyyy-MM-dd') throw new Error('unsupported pattern in stand-in: ' + pattern);
        return new Intl.DateTimeFormat('en-CA', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
      },
    },
    UrlFetchApp: {
      fetch(url, options) {
        calls.fetch.push({ url, options });
        const r = fetchImpl(url, options, calls.fetch.length);
        const body = typeof r.body === 'string' ? r.body : JSON.stringify(r.body);
        return { getResponseCode: () => r.code, getContentText: () => body, getHeaders: () => r.headers || {} };
      },
    },
    MailApp: { sendEmail: (message) => calls.mail.push(message) },
    ScriptApp: {
      WeekDay: { MONDAY: 'MONDAY', TUESDAY: 'TUESDAY', WEDNESDAY: 'WEDNESDAY', THURSDAY: 'THURSDAY', FRIDAY: 'FRIDAY', SATURDAY: 'SATURDAY', SUNDAY: 'SUNDAY' },
      getProjectTriggers: () => existingTriggers.slice(),
      deleteTrigger: (t) => { calls.deleted.push(t); existingTriggers.splice(existingTriggers.indexOf(t), 1); },
      newTrigger(fn) {
        const spec = { fn, steps: [] };
        const step = (name) => (...args) => { spec.steps.push([name, ...args]); return b; };
        const b = {
          timeBased: step('timeBased'), everyDays: step('everyDays'), onWeekDay: step('onWeekDay'),
          atHour: step('atHour'), inTimezone: step('inTimezone'),
          create: () => { calls.triggers.push(spec); existingTriggers.push({ getHandlerFunction: () => fn, spec }); return spec; },
        };
        return b;
      },
    },
  };

  const ctx = vm.createContext(sandbox);
  vm.runInContext(CODE, ctx, { filename: 'Code.gs' });
  Object.assign(ctx.CONFIG, config);
  for (const [k, v] of Object.entries(tunables)) ctx[k] = v;
  return { ctx, calls, existingTriggers };
}

// A page of Airtable records in the shape of the REST API response.
function page(records, offset) {
  const body = { records: records.map((r, i) => ({ id: r.id || 'rec' + String(i).padStart(14, '0'), createdTime: '2026-01-01T00:00:00.000Z', fields: r.fields || r })) };
  if (offset) body.offset = offset;
  return body;
}

module.exports = { load, page };
