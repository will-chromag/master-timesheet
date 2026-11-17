# Master Timesheet

Chromag hours tracking web app backed by Google Sheets and Google Apps Script.

## Multiple punches per day

Each day always has a permanent punch **1** row. Use **+** on that row to add
another punch; additional punches receive the next numeric ID for that date,
and **×** removes only additional punches. IDs are saved with the rows so a
different browser can load and update the same punches without duplicating
them. Each punch row has its own **In**, **Out**, and **Clear** controls.
Punch hours are included in day and pay-period totals.

The Payroll Summary **30 Day Avg** column groups punches by day over the latest
30 calendar days and averages only days with at least five total hours.

Time inputs remain native picker fields. When entering times by keyboard, try
`9` for 09:00, `4p` for 16:00, `930a` for 09:30, `9:30p` for 21:30, or `16:30`.
The value is normalized when you leave the field.

When updating the app, deploy the latest `appscript.gs` as a new version of the
Apps Script web app as well as serving the updated `index.html`. The script adds
an `Entry ID` column to employee tabs automatically when they are next read or
written, and populates the Payroll Summary average formulas when the roster is
loaded.
