/**
 * Airtable view -> email with CSV attachment (Google Apps Script)
 * See README.md for setup.
 *
 * Fetches every record of one Airtable view, builds a CSV file and emails it
 * as an attachment. Run sendViewAsCsv() by hand, or run installDailyTrigger()
 * or installWeeklyTrigger() once to send it on a schedule.
 *
 * The Airtable personal access token is NOT stored in this file. Add it under
 * Project Settings > Script properties as AIRTABLE_TOKEN.
 * Required token scope: data.records:read, with access to the base.
 *
 * Copyright (c) 2026 BASIT B.V. - MIT License
 */

// ---------------------------------------------------------------------------
// Configuration: edit these
// ---------------------------------------------------------------------------
var CONFIG = {
  // Read from Script Properties (key AIRTABLE_TOKEN). Never put the token here.
  get AIRTABLE_TOKEN() {
    return PropertiesService.getScriptProperties().getProperty('AIRTABLE_TOKEN');
  },
  BASE_ID: 'appXXXXXXXXXXXXXX',   // starts with "app"; find it in the Airtable API docs URL
  TABLE: 'Table 1',               // table name or table ID (tblXXXXXXXXXXXXXX)
  VIEW: 'Grid view',              // view name or ID (viwXXXXXXXXXXXXXX); only its records, in its order
  FIELDS: [],                     // e.g. ['Name', 'Status']; [] = every field, in order of first appearance
  RECIPIENTS: '',                 // comma separated; '' = the account that runs the script
  SUBJECT: 'Airtable export: {view} ({count} records) - {date}',  // placeholders: {date} {count} {view}
  FILENAME: 'airtable-export-{date}.csv',                         // placeholder: {date}
  SKIP_IF_EMPTY: true,            // true = send nothing when the view has no records
  TIMEZONE: '',                   // IANA name, e.g. 'Europe/Amsterdam'; '' = the script's time zone
  LOCALE: 'en'                    // Airtable userLocale for number and date text, e.g. 'en', 'en-gb', 'nl'
};

// Tuning: normally leave as is
var PAGE_SIZE = 100;                              // Airtable maximum
var PAUSE_MS = 250;                               // pause between pages, keeps under 5 requests per second
var MAX_ATTEMPTS = 5;                             // tries per request for 429 and 5xx responses
var MAX_RUNTIME_MS = 5 * 60 * 1000;               // stop before Apps Script's 6 minute limit
var MAX_ATTACHMENT_BYTES = 24 * 1024 * 1024;      // Gmail allows 25 MB per message; keep a margin
var LIST_SEPARATOR = ', ';                        // joins array values in a cell
var BOM = '\uFEFF';                               // makes Excel read the file as UTF-8

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/** Fetch the view from Airtable and email it as a CSV attachment. */
function sendViewAsCsv() {
  var token = CONFIG.AIRTABLE_TOKEN;
  if (!token) {
    throw new Error('AIRTABLE_TOKEN is missing. Add it under Project Settings > Script properties.');
  }
  if (!CONFIG.BASE_ID || CONFIG.BASE_ID === 'appXXXXXXXXXXXXXX') {
    throw new Error('Set BASE_ID in CONFIG at the top of the script.');
  }

  var started = Date.now();
  var timeZone = resolveTimeZone_(CONFIG.TIMEZONE);
  var records = fetchAllRecords_(token, started, timeZone);

  if (records.length === 0 && CONFIG.SKIP_IF_EMPTY) {
    Logger.log('View "' + CONFIG.VIEW + '" has no records and SKIP_IF_EMPTY is true. No email sent.');
    return { sent: false, records: 0 };
  }

  var csv = buildCsv_(records, CONFIG.FIELDS);
  var date = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  var vars = { date: date, count: records.length, view: CONFIG.VIEW || CONFIG.TABLE };
  var subject = applyTemplate_(CONFIG.SUBJECT, vars);
  var filename = buildFilename_(CONFIG.FILENAME, vars);
  var recipients = parseRecipients_(CONFIG.RECIPIENTS);
  if (!recipients.length) recipients = [Session.getEffectiveUser().getEmail()];

  var blob = Utilities.newBlob(csv, 'text/csv', filename);
  var bytes = blob.getBytes().length;
  if (bytes > MAX_ATTACHMENT_BYTES) {
    throw new Error('The CSV is ' + Math.round(bytes / 1048576) + ' MB, over the ~25 MB email limit. ' +
      'Use a smaller VIEW or set FIELDS to fewer columns.');
  }

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: buildBody_(vars, CONFIG.TABLE),
    attachments: [blob]
  });

  Logger.log('Sent ' + records.length + ' records (' + bytes + ' bytes) to ' + recipients.length + ' recipient(s).');
  return { sent: true, records: records.length, bytes: bytes, filename: filename };
}

/** Run once: sends every day at `hour` (0-23, default 7), replacing existing triggers of this script. */
function installDailyTrigger(hour) {
  hour = checkHour_(hour === undefined ? 7 : hour);
  removeTriggers();
  var builder = ScriptApp.newTrigger('sendViewAsCsv').timeBased().everyDays(1).atHour(hour);
  if (CONFIG.TIMEZONE) builder = builder.inTimezone(CONFIG.TIMEZONE);
  builder.create();
}

/**
 * Run once: sends every week on `weekday` at `hour`, replacing existing triggers of this script.
 * weekday: 'MONDAY' ... 'SUNDAY' (any case, or 3-letter 'mon'), or a number 1 (Monday) to 7 (Sunday).
 */
function installWeeklyTrigger(weekday, hour) {
  var day = ScriptApp.WeekDay[weekdayName_(weekday)];
  hour = checkHour_(hour === undefined ? 7 : hour);
  removeTriggers();
  var builder = ScriptApp.newTrigger('sendViewAsCsv').timeBased().onWeekDay(day).atHour(hour);
  if (CONFIG.TIMEZONE) builder = builder.inTimezone(CONFIG.TIMEZONE);
  builder.create();
}

/** Removes every trigger that runs sendViewAsCsv. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'sendViewAsCsv') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ---------------------------------------------------------------------------
// Pure helpers (no Google services): unit-tested with Node, see test/
// ---------------------------------------------------------------------------

/** One CSV field per RFC 4180: quote if it holds a comma, quote, CR or LF; double inner quotes. */
function csvEscape_(value) {
  var text;
  if (value === null || value === undefined) text = '';
  else if (Array.isArray(value)) text = value.map(csvText_).filter(function (t) { return t !== ''; }).join(LIST_SEPARATOR);
  else text = csvText_(value);
  return /[",\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}

function csvText_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    if (typeof value.url === 'string') return value.url;
    if (typeof value.name === 'string' && value.name) return value.name;
    return JSON.stringify(value);
  }
  return String(value);
}

/**
 * Header: FIELDS if given, else the union of field names in order of first appearance
 * (Airtable leaves empty fields out of a record, so records can differ).
 */
function buildHeader_(records, fieldNames) {
  var header = (fieldNames || []).slice();
  if (header.length) return header;
  var seen = {};
  records.forEach(function (record) {
    Object.keys(record.fields || {}).forEach(function (name) {
      if (!Object.prototype.hasOwnProperty.call(seen, name)) {
        seen[name] = true;
        header.push(name);
      }
    });
  });
  return header;
}

/** Whole CSV text: UTF-8 BOM, header row, one row per record, CRLF line endings. */
function buildCsv_(records, fieldNames) {
  var header = buildHeader_(records, fieldNames);
  if (!header.length) return BOM;
  var lines = [header.map(csvEscape_).join(',')];
  records.forEach(function (record) {
    var fields = record.fields || {};
    lines.push(header.map(function (name) { return csvEscape_(fields[name]); }).join(','));
  });
  return BOM + lines.join('\r\n') + '\r\n';
}

/** URL for one page of records. `cfg` is CONFIG (or a copy of it). */
function buildUrl_(cfg, offset, timeZone) {
  var enc = encodeURIComponent;
  var query = ['pageSize=' + PAGE_SIZE];
  if (cfg.VIEW) query.push('view=' + enc(cfg.VIEW));
  query.push('cellFormat=string');
  query.push('timeZone=' + enc(timeZone));
  query.push('userLocale=' + enc(cfg.LOCALE || 'en'));
  (cfg.FIELDS || []).forEach(function (name) { query.push('fields%5B%5D=' + enc(name)); });
  if (offset) query.push('offset=' + enc(offset));
  return 'https://api.airtable.com/v0/' + enc(cfg.BASE_ID) + '/' + enc(cfg.TABLE) + '?' + query.join('&');
}

/** Replace {name} placeholders. Unknown placeholders are left as they are. */
function applyTemplate_(template, vars) {
  return String(template).replace(/\{(\w+)\}/g, function (whole, key) {
    return Object.prototype.hasOwnProperty.call(vars, key) ? String(vars[key]) : whole;
  });
}

/** File name from the template: characters that file systems reject become '-', '.csv' is added if missing. */
function buildFilename_(template, vars) {
  var name = applyTemplate_(template, vars).replace(/[\\\/:*?"<>|\r\n]+/g, '-').trim();
  if (!name) name = 'airtable-export-' + vars.date;
  return /\.csv$/i.test(name) ? name : name + '.csv';
}

function buildBody_(vars, table) {
  return 'Attached: ' + vars.count + ' record' + (vars.count === 1 ? '' : 's') + ' from the Airtable view "' +
    vars.view + '" (table "' + table + '") as of ' + vars.date + '.\n\n' +
    'Sent by a Google Apps Script on a schedule.';
}

function parseRecipients_(list) {
  return String(list || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s !== ''; });
}

function resolveTimeZone_(configured) {
  return configured || Session.getScriptTimeZone() || 'Etc/UTC';
}

/** True if a response with this status should be tried again. `attempt` is 0 for the first try. */
function shouldRetry_(code, attempt) {
  return (code === 429 || code >= 500) && attempt < MAX_ATTEMPTS - 1;
}

/** Wait before retry number `attempt` (0-based): 2, 4, 8, 16, 30 s, or Retry-After (capped at 60 s). */
function backoffMs_(attempt, headers) {
  var retryAfter = headers && (headers['Retry-After'] || headers['retry-after']);
  var seconds = parseInt(retryAfter, 10);
  if (seconds > 0) return Math.min(seconds, 60) * 1000;
  return Math.min(2000 * Math.pow(2, attempt), 30000);
}

function describeError_(code, body) {
  var message = '';
  try {
    var parsed = JSON.parse(body);
    message = parsed.error && (parsed.error.message || parsed.error.type || parsed.error);
    if (typeof message !== 'string') message = JSON.stringify(message);
  } catch (e) {
    message = String(body).slice(0, 200);
  }
  var hint = '';
  if (code === 401) hint = ' Check the AIRTABLE_TOKEN script property.';
  if (code === 403) hint = ' Check that the token has the data.records:read scope and access to this base.';
  if (code === 404) hint = ' Check BASE_ID, TABLE and VIEW in CONFIG (names are case sensitive).';
  if (code === 422) hint = ' Check VIEW and FIELDS in CONFIG.';
  message = String(message).replace(/[.\s]+$/, '');
  return 'Airtable returned HTTP ' + code + ': ' + message + '.' + hint;
}

function weekdayName_(weekday) {
  var names = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'];
  if (typeof weekday === 'number' && weekday >= 1 && weekday <= 7 && weekday % 1 === 0) return names[weekday - 1];
  if (typeof weekday === 'string') {
    var key = weekday.trim().toUpperCase();
    for (var i = 0; i < names.length; i++) {
      if (names[i] === key || names[i].slice(0, 3) === key) return names[i];
    }
  }
  throw new Error('Unknown weekday "' + weekday + '". Use \'MONDAY\' ... \'SUNDAY\' or a number 1 (Monday) to 7 (Sunday).');
}

function checkHour_(hour) {
  if (typeof hour !== 'number' || hour % 1 !== 0 || hour < 0 || hour > 23) {
    throw new Error('hour must be a whole number from 0 to 23, got ' + hour + '.');
  }
  return hour;
}

// ---------------------------------------------------------------------------
// Airtable access
// ---------------------------------------------------------------------------

/** GET one URL, retrying 429 and 5xx with backoff. Returns parsed JSON. */
function fetchJson_(url, token) {
  for (var attempt = 0; ; attempt++) {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 200) return JSON.parse(response.getContentText());
    if (shouldRetry_(code, attempt)) {
      Utilities.sleep(backoffMs_(attempt, response.getHeaders && response.getHeaders()));
      continue;
    }
    throw new Error(describeError_(code, response.getContentText()));
  }
}

/** Follow Airtable's offset cursor until the last page. */
function fetchAllRecords_(token, started, timeZone) {
  var records = [];
  var offset = '';
  do {
    if (Date.now() - started > MAX_RUNTIME_MS) {
      throw new Error('Stopped after ' + Math.round(MAX_RUNTIME_MS / 60000) + ' minutes with ' + records.length +
        ' records loaded. Nothing was emailed. Use a smaller VIEW or fewer FIELDS.');
    }
    var page = fetchJson_(buildUrl_(CONFIG, offset, timeZone), token);
    Array.prototype.push.apply(records, page.records || []);
    offset = page.offset || '';
    if (offset) Utilities.sleep(PAUSE_MS);
  } while (offset);
  return records;
}

// Lets Node load the pure helpers. Apps Script has no `module`, so this is skipped there.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    csvEscape_: csvEscape_, buildHeader_: buildHeader_, buildCsv_: buildCsv_, buildUrl_: buildUrl_,
    applyTemplate_: applyTemplate_, buildFilename_: buildFilename_, parseRecipients_: parseRecipients_,
    shouldRetry_: shouldRetry_, backoffMs_: backoffMs_, describeError_: describeError_,
    weekdayName_: weekdayName_, checkHour_: checkHour_
  };
}
