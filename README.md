# Master Timesheet

Chromag hours tracking web app backed by Google Sheets and Google Apps Script.

## Multiple punches per day

Each day always has a permanent punch **1** row. Use **+** on that row to add
another punch; additional punches receive the next numeric ID for that date,
and **×** removes only additional punches. IDs are saved with the rows so a
different browser can load and update the same punches without duplicating
them. Punch hours are included in day and pay-period totals.

When updating the app, deploy the latest `appscript.gs` as a new version of the
Apps Script web app as well as serving the updated `index.html`. The script adds
an `Entry ID` column to employee tabs automatically when they are next read or
written.
