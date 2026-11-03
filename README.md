# Airtable View to Email (CSV) on a Schedule

A single Google Apps Script file (`Code.gs`) that emails one Airtable view as a CSV attachment on a schedule. No add-on, no third-party service, no dependencies.

## Why

Airtable has no built-in way to email a CSV of a view on a schedule, and the "Send email" action in Airtable automations cannot attach a CSV of a view. This script fills that gap with a time-driven trigger in Google Apps Script.

- Reads the Airtable token from Script Properties, never from the code.
- Pages through the Airtable API (100 records per request, `offset` cursor), so views of any size within the limits below work.
- Asks Airtable for text values (`cellFormat=string`, with your time zone and locale), so linked records, lookups and dates arrive as readable text.
- Builds an RFC 4180 CSV: fields with commas, quotes or line breaks are quoted, inner quotes are doubled, lines end with CRLF.
- Starts the file with a UTF-8 byte order mark, so Excel shows accented characters correctly.
- Sends it with `MailApp.sendEmail` to one or more recipients, with a subject and file name you can template with the date and record count.
- Can skip the email when the view is empty.
- Retries HTTP 429 (rate limit) and 5xx responses with backoff, and gives a clear error for 401, 403 and 404.
- Includes `installDailyTrigger()`, `installWeeklyTrigger()` and `removeTriggers()`.

## Setup

1. **Create an Airtable token.** At [airtable.com/create/tokens](https://airtable.com/create/tokens) create a personal access token with the scope `data.records:read` and access to the base you want to export. Copy the token.
2. **Create an Apps Script project.** Go to [script.google.com](https://script.google.com), choose *New project*, and paste the contents of `Code.gs` over the default file.
3. **Add the token.** Open *Project Settings* (gear icon), scroll to *Script properties*, click *Add script property*, and set the name `AIRTABLE_TOKEN` and the value to your token.
4. **Edit `CONFIG` at the top of `Code.gs`.** Set `BASE_ID` (starts with `app`), `TABLE`, `VIEW` and `RECIPIENTS`. The other settings are optional, see the table below.
5. **Run it once.** Pick `sendViewAsCsv` in the function list and click *Run*. Google asks you to authorize the script (access to send email and to connect to an external service). You should receive the email within a minute.
6. **Schedule it.** Pick `installDailyTrigger` in the function list and click *Run*. Without an argument it sends every day around 07:00 (the same as `installDailyTrigger(7)`). Run `removeTriggers` to stop the schedule.

Running an install function again replaces the existing trigger of this script, so you never get duplicates.

The editor's *Run* button cannot pass arguments. For another hour, or a weekly email, add a small wrapper function and run it once:

```js
function scheduleDailyAt9() { installDailyTrigger(9); }
function scheduleMondayAt8() { installWeeklyTrigger('MONDAY', 8); }
```

## Configuration

All settings are in the `CONFIG` object at the top of `Code.gs`.

| Setting | Meaning | Default |
| --- | --- | --- |
| `AIRTABLE_TOKEN` | Read from the Script Property of the same name. Not a value you edit | n/a |
| `BASE_ID` | Airtable base ID, starts with `app` | placeholder, must be set |
| `TABLE` | Table name or ID (`tbl...`) | `Table 1` |
| `VIEW` | View name or ID (`viw...`). Only that view's records, in that view's order | `Grid view` |
| `FIELDS` | List of field names to include, in that order | `[]` (every field) |
| `RECIPIENTS` | Comma separated email addresses | `''` (the account that runs the script) |
| `SUBJECT` | Subject template. `{date}` is the send date (`yyyy-MM-dd`), `{count}` the number of records, `{view}` the view name | `Airtable export: {view} ({count} records) - {date}` |
| `FILENAME` | Attachment name template. `{date}` is replaced; `.csv` is added if missing | `airtable-export-{date}.csv` |
| `SKIP_IF_EMPTY` | `true` = send nothing when the view has no records (a line is written to the execution log) | `true` |
| `TIMEZONE` | IANA time zone such as `Europe/Amsterdam`. Used for `{date}`, for Airtable's date text, and for the trigger hour | `''` (the script's time zone) |
| `LOCALE` | Airtable `userLocale` for number and date text, such as `en`, `en-gb`, `nl` | `en` |

Without `FIELDS`, the columns are the union of the field names that appear in the returned records, in order of first appearance. Airtable leaves empty fields out of a record, so a field that is empty in every record has no column. Set `FIELDS` for a fixed set and order of columns (also useful when `SKIP_IF_EMPTY` is `false`, so an empty view still gets a header row).

## Limits and quotas

These are Google's and Airtable's limits, not the script's. Check Google's Apps Script quotas page for current numbers.

- **Email recipients per day:** `MailApp` allows about 100 recipients per day on consumer Google accounts, and more on Google Workspace accounts. One email to three recipients counts as three.
- **Attachment size:** Gmail caps a message at 25 MB. The script refuses to send a CSV over 24 MB and tells you to use a smaller view or fewer fields.
- **Execution time:** 6 minutes per run. The script stops itself after 5 minutes without sending anything. A page of 100 records takes about a second, so tens of thousands of records are the practical ceiling.
- **Airtable API:** 5 requests per second per base. The script pauses 250 ms between pages and retries a 429 after 2, 4, 8 and 16 seconds (or the `Retry-After` value if Airtable sends one), up to 5 tries per request.
- **Trigger timing:** Time-driven triggers run within about an hour of the chosen hour, not at the exact minute.

## Tests

The Apps Script cannot run outside Google, so the tests load `Code.gs` into a Node `vm` with stand-ins for `UrlFetchApp`, `MailApp`, `Utilities`, `PropertiesService`, `Session` and `ScriptApp`. The Airtable responses are hand-written in the shape the Airtable REST API documents. They cover CSV quoting, BOM and CRLF, header order, pagination across three pages, retry on 429, error messages, `SKIP_IF_EMPTY`, templating and trigger installation.

```
npm test
```

Needs Node 18 or newer (`node --test`), nothing to install. The Google-side behaviour is exercised through the stand-ins only, so run the script once from the editor before relying on the schedule.

## FAQ

**Excel shows broken characters like `Ã©` instead of accents.** The file starts with a UTF-8 byte order mark exactly for this, so double-clicking the CSV in Excel should show `é` correctly. If you import the file through *Data > From Text/CSV*, choose UTF-8 as the file origin.

**Does it work with more than 100 records?** Yes. The script follows Airtable's `offset` cursor until the last page. The limits are the execution time and the 25 MB attachment size above.

**What do attachment fields look like in the CSV?** As text that includes the file's Airtable URL. Those URLs expire after about two hours, so a link in an email you open tomorrow may no longer work. See [why Airtable attachment links expire and what to do about it](https://viewcourier.com/guides/airtable-csv-attachment-links-expired?utm_source=github&utm_medium=readme&utm_campaign=email_csv_repo).

**How are linked records, lookups and dates shown?** As readable text, because the script requests `cellFormat=string`. Multiple values are joined with `, ` inside one quoted cell. Dates and numbers follow `TIMEZONE` and `LOCALE`.

**I get "HTTP 403" or "HTTP 404".** The token has no access to that base, lacks the `data.records:read` scope, or `BASE_ID`, `TABLE` or `VIEW` is wrong. Names are case sensitive; IDs (`tbl...`, `viw...`) are safer than names.

**"Authorization is required" when running.** Run `sendViewAsCsv` from the editor once and accept the Google permission prompt. A time-driven trigger cannot show the prompt.

**Where is the token stored?** In Script Properties of the Apps Script project, not in the code. Anyone who can edit the script project can read it, so share the project with care. Use a token limited to `data.records:read` on the one base.

**Can it export several views?** Use one Apps Script project per view, or copy `sendViewAsCsv` into a second function that sets different `CONFIG` values before calling it.

**Who is the sender?** The Google account that authorized the script and owns the trigger. `MailApp` cannot change the sender address.

## Prefer no code?

[ViewCourier](https://viewcourier.com/?utm_source=github&utm_medium=readme&utm_campaign=email_csv_repo) schedules Airtable view exports as CSV, Excel or PDF to any inbox, with a free plan. No script to maintain. A step-by-step guide to the no-code route: [Email an Airtable view on a schedule](https://viewcourier.com/guides/email-airtable-view-on-a-schedule?utm_source=github&utm_medium=readme&utm_campaign=email_csv_repo).

## License

MIT. Copyright (c) 2026 ViewCourier
