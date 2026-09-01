# Jira timesheet gap finder

Finds the holes in your Jira board. Asks Jira what you have already logged,
reads your Outlook calendar for the same period, and tells you exactly how many
hours are still missing on each day — so when you fill your Excel sheet you top
up by 2 hours instead of blindly adding 8 and double-counting.

Nothing here writes to Jira. It is read-only. The output is a spreadsheet you
feed into the upload script you already have.

---

## 1. One-time setup on Windows

Open Command Prompt in the folder containing these files.

```bat
pip install -r requirements.txt
```

If `pip` is not recognised, use `py -m pip install -r requirements.txt`.

### Get a Jira token

**Jira Cloud** (your URL ends in `.atlassian.net`):
go to <https://id.atlassian.com/manage-profile/security/api-tokens>, create an
API token, and copy it. Set `deployment = cloud` and fill in your `email`.

**Jira Server / Data Center** (self-hosted): open your Jira profile menu →
Personal Access Tokens → Create token. Set `deployment = server` and leave
`email` blank.

Either way you only need your normal account. No admin rights, no app
registration, no approval from a platform team — the token acts as you and sees
exactly what you already see in the browser.

### Fill in `config.ini`

Every line marked `CHANGE ME` needs your value. At minimum: `base_url`,
`deployment`, `email` (Cloud only), `api_token`, and `csv_path`.

If you would rather not leave the token in a file:

```bat
setx JIRA_API_TOKEN "your-token-here"
```

Close and reopen Command Prompt afterwards. The script prefers the environment
variable over `config.ini`.

### Check it works

```bat
python build_gap_report.py --check
```

You should see `Connected to Jira as: Your Name`.

---

## 2. Export your Outlook calendar

Do this each time you want a fresh report. Classic Outlook desktop:

1. **File → Open & Export → Import/Export**
2. Choose **Export to a file** → Next
3. Choose **Comma Separated Values** → Next
4. Pick your **Calendar** → Next
5. Choose where to save it → Next
6. Click **Map Custom Fields** and confirm at minimum: Subject, Start Date,
   Start Time, End Date, End Time, All day event, Show time as, Location,
   Meeting Organizer
7. Click **Finish**. Outlook asks for a date range — set it to cover the whole
   period you are reporting on

New Outlook / web has no CSV export. Options: switch to classic Outlook for the
export, or in Outlook on the web go to Settings → Calendar → Shared calendars →
**Publish a calendar**, publish as ICS, and save the file. If you end up with an
ICS instead of a CSV, tell me and I will add an ICS reader — it is a small
addition.

Recurring meetings are exported as individual occurrences, which is what you
want. Save the file to the path you put in `csv_path`.

---

## 3. Run it

```bat
python build_gap_report.py --from 2026-07-01 --to 2026-08-31
```

With no dates it defaults to the first day of last month through today.

| Flag | What it does |
|---|---|
| `--check` | Test the Jira connection, then stop |
| `--from` / `--to` | Date window, `YYYY-MM-DD` |
| `--outlook PATH` | Use a different CSV than the one in config |
| `--out PATH` | Write the report somewhere else |
| `--no-jira` | Build from Outlook alone (useful before your token works) |
| `--no-outlook` | Show only what Jira already holds |

---

## 4. Reading the output

The workbook has five sheets:

**Read Me** — a legend of the run: window, account, target hours, source file.

**Daily Summary** — one row per day. The important column is **Gap**: hours
still missing after what Jira already holds. Orange rows are short, green are on
target, yellow means you have logged *more* than the target that day.

**Jira Worklogs** — every entry already sitting in Jira under your name. This is
the evidence behind the Gap column. If a number looks wrong, check here first.

**Outlook Meetings** — every meeting the parser kept, with its duration.

**Upload Template** — the sheet you act on. Columns: Issue Key, Work Date,
Hours, Comment, Source. Yellow cells need a value from you.

### How hours get suggested

For each working day it takes the gap, fills it with that day's meetings oldest
first, and whatever is left over becomes a single "remaining hours" row.

So a Monday where you already logged 6 hours manually and sat in 4 hours of
calls produces **2 hours** of meeting entries — not 4, and not 8. That was the
whole point.

Two details worth knowing:

- **Overlapping meetings are counted once.** Two calls both running 10:00–11:00
  is 1 hour of your day, not 2. Without this your meeting totals inflate on
  double-booked days and you overshoot 8 hours.
- **Meetings crossing midnight are split** across the two days they touch.

---

## 5. Wiring it into your existing script

Your current script reads an Excel sheet and creates Jira tickets and worklogs.
The Upload Template sheet is built to slot into that with minimal fuss.

If your script expects different column headers, either rename the headers on
that sheet, or change them at the source — in `build_gap_report.py`, find:

```python
_style_header(ws, ["Issue Key", "Work Date", "Hours", "Comment", "Source"],
```

and put your own names in that list.

If you set `default_meeting_issue` and `default_work_issue` in `config.ini`, the
Issue Key column comes out pre-filled and there is nothing to type at all.

**Before your first upload,** run the report, upload from it, then run the
report again. Every day should come back green. That round trip confirms the two
halves agree with each other before you trust it with two months of backfill.

---

## Troubleshooting

**401 Unauthorized** — on Cloud the token must be an API token, not your
password. On Server it must be a Personal Access Token, not an API token.

**403 Forbidden** — usually a CAPTCHA lockout after failed logins. Sign into
Jira in a browser once, then retry.

**SSL errors** — corporate proxy intercepting HTTPS. Set `verify_ssl = false`.

**Dates come out wrong** — set `date_format` to `us` or `eu` instead of `auto`.
Auto-detection needs at least one date where the day is above 12 to be certain.

**Meeting hours look too high** — add subjects to `exclude_subjects`, or set
`skip_tentative = true`.

**A day shows zero logged but you know you logged it** — check the Jira Worklogs
sheet. If the entry is missing, the worklog is probably under a different
account, or on an issue in a project your account cannot browse.

## Files

| File | Purpose |
|---|---|
| `build_gap_report.py` | Main script — run this |
| `jira_worklogs.py` | Read-only Jira API client |
| `outlook_csv.py` | Calendar CSV parser and overlap merging |
| `config.ini` | Your settings |
| `requirements.txt` | Dependencies |
