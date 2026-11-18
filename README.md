# Master Timesheet

Chromag hours tracking web app backed by Google Sheets and Google Apps Script.

## Multiple punches per day

Each day always has a permanent punch **1** row. Use **+** on that row to add
another punch; additional punches receive the next numeric ID for that date,
and **×** removes only additional punches. IDs are saved with the rows so a
different browser can load and update the same punches without duplicating
them. Each punch row has its own **In**, **Out**, and **Clear** controls.
Punch hours are included in day and pay-period totals.
On mobile, the controls appear above the time fields, with Notes and the
add/remove button on the row below.

The Payroll Summary **30 Day Avg** column groups punches by day over the latest
30 calendar days and averages only days with at least five total hours.
Choose a preset pay period, **Year to Date**, or **Custom Range** from the
selector in B1. Year to Date covers January 1 through today. For a custom
range, enter the start and end dates in B2 and B3; the purple heading and
payroll totals update to match. The selected range is shown in the purple
merged heading above the summary table.

The **All Payroll** tab lists each staff punch for its selected period, followed
by a yellow per-staff hours subtotal and a grand total. Its period dropdown and
custom date pickers work like Payroll Summary. The last column shows when each
entry was last updated. Custom date inputs are labeled in A2:A3; the grand
total row uses blue fill and white text.

Time inputs remain native picker fields. When entering times by keyboard, try
`9` for 09:00, `4p` for 16:00, `930a` for 09:30, `9:30p` for 21:30, or `16:30`.
The value is normalized when you leave the field.

When updating the app, deploy the latest `appscript.gs` as a new version of the
Apps Script web app as well as serving the updated `index.html`. The script adds
an `Entry ID` column to employee tabs automatically when they are next read or
written, and populates the Payroll Summary average formulas when the roster is
loaded.

Employee PINs are verified by the Apps Script against the current `Staff` sheet
on each PIN prompt. The app does not accept a cached or built-in PIN when the
sheet cannot be reached; deploy both the script and app updates for this behavior.
