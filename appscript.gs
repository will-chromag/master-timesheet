const DEFAULT_STAFF = [
  "Claire Buchar", "Dave Thoss", "Emilie de Crombrugghe", "Emily Slaco",
  "Ian Ritz", "James Hawkins", "JM Josse", "Joe Swann",
  "Julian Hine", "Kraig Kinsman", "Lachie O’Connor", "Monique Joyce",
  "Pete Fowler", "Rebecca Ritz", "Sara Hrubesova", "Steven Janney", "Tim Wilson",
  "William Westwood"
];

function getStoredAdminPassword() {
  return PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || "Upupawayjoey!";
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = e.parameter.action;

    // 1. Fetch Staff Roster and PINs
    if (action === "getStaff") {
      let staffSheet = ss.getSheetByName("Staff");
      if (!staffSheet) {
        staffSheet = ss.insertSheet("Staff");
        staffSheet.appendRow(["Staff Name", "PIN"]);
        staffSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
        staffSheet.getRange("B:B").setNumberFormat("@");

        const rows = DEFAULT_STAFF.map(name => [name, "'1234"]);
        staffSheet.getRange(2, 1, rows.length, 2).setValues(rows);
      }

      ensurePayrollAverageColumn(ss);
      ensurePayrollRangeControls(ss);
      ensurePayrollSummaryFormulas(ss);

      const lastRow = staffSheet.getLastRow();
      const staffList = [];
      if (lastRow > 1) {
        const raw = staffSheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (let i = 0; i < raw.length; i++) {
          const name = String(raw[i][0]).trim();
          let pin = String(raw[i][1]).replace(/^'/, "").trim();
          if (name) {
            staffList.push({ name: name, pin: pin || "1234" });
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "success", staff: staffList }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Fetch User Timesheet Data
    const staffName = (e.parameter.staffName || "").trim();
    const sheet = ss.getSheetByName(staffName);

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ status: "not_found", entries: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    ensureEntryIds(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", entries: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getRange(2, 2, lastRow - 1, 6).getValues();
    const timeZone = Session.getScriptTimeZone() || "GMT";
    const entries = [];

    function formatTimeValue(val) {
      if (!val) return "";
      if (val instanceof Date) {
        return Utilities.formatDate(val, timeZone, "HH:mm");
      }
      const str = String(val).trim();
      const match = str.match(/^(\d{1,2}):(\d{2})/);
      if (match) {
        return match[1].padStart(2, "0") + ":" + match[2];
      }
      return "";
    }

    for (let i = 0; i < data.length; i++) {
      let val = data[i][0];
      let dateStr = "";
      if (val instanceof Date) {
        dateStr = Utilities.formatDate(val, timeZone, "yyyy-MM-dd");
      } else {
        dateStr = String(val).replace(/^'/, "").trim();
      }

      if (dateStr) {
        entries.push({
          dateStr: dateStr,
          startTime: formatTimeValue(data[i][1]),
          endTime: formatTimeValue(data[i][2]),
          totalHours: Number(data[i][3]) || 0,
          notes: String(data[i][4] || ""),
          entryId: String(data[i][5] || "").trim() || "1"
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", entries: entries }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const validAdminPassword = getStoredAdminPassword();

    // 1. Verify Admin Password Endpoint
    if (data.action === "verifyAdminPassword") {
      if (String(data.password || "").trim() === validAdminPassword) {
        return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Incorrect password" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 2. Admin Action: Save Updated Staff Roster and PINs
    if (data.action === "saveStaff") {
      if (String(data.adminPassword || "").trim() !== validAdminPassword) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unauthorized: Invalid password" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      let staffSheet = ss.getSheetByName("Staff");
      if (!staffSheet) {
        staffSheet = ss.insertSheet("Staff");
        staffSheet.appendRow(["Staff Name", "PIN"]);
        staffSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
        staffSheet.getRange("B:B").setNumberFormat("@");
      }

      const currentLast = staffSheet.getLastRow();
      if (currentLast > 1) {
        staffSheet.getRange(2, 1, currentLast - 1, 2).clearContent();
      }

      const list = data.staffList || [];
      if (list.length > 0) {
        const rows = list.map(item => [item.name.trim(), "'" + String(item.pin || "1234").trim()]);
        staffSheet.getRange(2, 1, rows.length, 2).setValues(rows);

        // Ensure both tab and summary row exist for every employee
        list.forEach(item => {
          const cleanName = item.name.trim();
          ensureStaffTab(ss, cleanName);
          registerStaffInSummary(ss, cleanName);
        });
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Shift Logging
    const staffName = String(data.staffName || "").trim();
    const targetDate = String(data.dateStr || "").trim();
    const entryId = data.entryId === undefined || data.entryId === null || data.entryId === ""
      ? "1"
      : String(data.entryId).trim();

    if (!staffName || !targetDate || !/^[1-9]\d*$/.test(entryId) || !Number.isSafeInteger(Number(entryId))) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid staff name, date, or entry ID" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "deleteEntry") {
      if (entryId === "1") {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "The main entry cannot be removed" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const sheet = ss.getSheetByName(staffName);
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ status: "not_found" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      ensureEntryIds(sheet);

      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const rows = sheet.getRange(2, 2, lastRow - 1, 6).getValues();
        const timeZone = Session.getScriptTimeZone() || "GMT";
        for (let i = 0; i < rows.length; i++) {
          const dateValue = rows[i][0];
          const rowDate = dateValue instanceof Date
            ? Utilities.formatDate(dateValue, timeZone, "yyyy-MM-dd")
            : String(dateValue).replace(/^'/, "").trim();
          const rowEntryId = String(rows[i][5] || "").trim() || "1";
          if (rowDate === targetDate && rowEntryId === entryId) {
            sheet.deleteRow(i + 2);
            SpreadsheetApp.flush();
            return ContentService.createTextOutput(JSON.stringify({ status: "success", deleted: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "not_found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let sheet = ensureStaffTab(ss, staffName);
    ensureEntryIds(sheet);
    registerStaffInSummary(ss, staffName);

    const lastRow = sheet.getLastRow();
    let rowIndex = -1;

    if (lastRow > 1) {
      const existingRows = sheet.getRange(2, 2, lastRow - 1, 6).getValues();
      for (let i = 0; i < existingRows.length; i++) {
        let val = existingRows[i][0];
        let valStr = "";
        if (val instanceof Date) {
          valStr = Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd");
        } else {
          valStr = String(val).replace(/^'/, "").trim();
        }
        const existingEntryId = String(existingRows[i][5] || "").trim() || "1";

        if (valStr === targetDate && existingEntryId === entryId) {
          rowIndex = i + 2;
          break;
        }
      }
    }

    const rowValues = [
      new Date(),
      "'" + targetDate,
      data.startTime || "",
      data.endTime || "",
      data.totalHours !== undefined ? data.totalHours : 0,
      data.notes || "",
      Number(entryId)
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, 7).setValues([rowValues]);
    } else {
      const targetRow = lastRow + 1;
      sheet.getRange(targetRow, 2).setNumberFormat("@");
      sheet.getRange(targetRow, 1, 1, 7).setValues([rowValues]);
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ status: "success", row: rowIndex }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function ensureStaffTab(ss, staffName) {
  let sheet = ss.getSheetByName(staffName);
  if (!sheet) {
    sheet = ss.insertSheet(staffName);
    sheet.appendRow(["Last Updated", "Date", "Start Time", "End Time", "Total Hours", "Notes"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.getRange("B:B").setNumberFormat("@");
  }
  ensureEntryIds(sheet);
  return sheet;
}

function ensureEntryIds(sheet) {
  if (!sheet.getRange(1, 7).getValue()) {
    sheet.getRange(1, 7).setValue("Entry ID").setFontWeight("bold");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const rows = sheet.getRange(2, 2, lastRow - 1, 6).getValues();
  const timeZone = Session.getScriptTimeZone() || "GMT";
  const entriesByDate = new Map();
  const outputIds = rows.map(row => [row[5]]);

  rows.forEach((row, index) => {
    const dateValue = row[0];
    if (!dateValue) return;
    const dateStr = dateValue instanceof Date
      ? Utilities.formatDate(dateValue, timeZone, "yyyy-MM-dd")
      : String(dateValue).replace(/^'/, "").trim();
    if (!dateStr) return;
    const entries = entriesByDate.get(dateStr) || [];
    entries.push({ index, existingId: String(row[5] || "").trim() });
    entriesByDate.set(dateStr, entries);
  });

  entriesByDate.forEach(entries => {
    const usedIds = new Set();
    const pending = [];

    entries.forEach(entry => {
      const currentId = entry.existingId;
      if (/^[1-9]\d*$/.test(currentId) && Number.isSafeInteger(Number(currentId)) && !usedIds.has(Number(currentId))) {
        usedIds.add(Number(currentId));
        outputIds[entry.index][0] = Number(currentId);
      } else {
        pending.push(entry);
      }
    });

    pending.sort((a, b) => (a.existingId === "main" ? -1 : 0) - (b.existingId === "main" ? -1 : 0));
    pending.forEach(entry => {
      let nextId = 1;
      while (usedIds.has(nextId)) nextId += 1;
      usedIds.add(nextId);
      outputIds[entry.index][0] = nextId;
    });
  });

  sheet.getRange(2, 7, outputIds.length, 1).setValues(outputIds);
}

function registerStaffInSummary(ss, staffName) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;
  ensureAverageHeader(summarySheet);

  const dataA = summarySheet.getRange("A:A").getValues();
  let totalRowIndex = -1;

  for (let i = 0; i < dataA.length; i++) {
    const val = String(dataA[i][0]).trim().toLowerCase();
    if (val.includes("total payroll") || val === "total") {
      totalRowIndex = i + 1;
      break;
    }
  }

  const searchLimit = totalRowIndex > 0 ? totalRowIndex - 1 : dataA.length;
  for (let r = 5; r < searchLimit; r++) {
    if (String(dataA[r][0]).trim().toLowerCase() === staffName.toLowerCase()) {
      setPayrollRangeFormulas(summarySheet, r + 1);
      setAverageFormula(summarySheet, r + 1);
      return;
    }
  }

  const insertAtRow = totalRowIndex > 0 ? totalRowIndex : summarySheet.getLastRow() + 1;
  if (totalRowIndex > 0) {
    summarySheet.insertRowBefore(insertAtRow);
  }

  summarySheet.getRange(insertAtRow, 1).setValue(staffName);
  setPayrollRangeFormulas(summarySheet, insertAtRow);
  setAverageFormula(summarySheet, insertAtRow);

  SpreadsheetApp.flush();
}

function ensurePayrollAverageColumn(ss) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

  ensureAverageHeader(summarySheet);
  const dataA = summarySheet.getRange("A:A").getValues();
  let totalRowIndex = dataA.findIndex(row => {
    const value = String(row[0]).trim().toLowerCase();
    return value.includes("total payroll") || value === "total";
  });
  totalRowIndex = totalRowIndex >= 0 ? totalRowIndex + 1 : summarySheet.getLastRow() + 1;

  for (let row = 6; row < totalRowIndex; row++) {
    if (String(dataA[row - 1]?.[0] || "").trim()) {
      setAverageFormula(summarySheet, row);
    }
  }
}

function ensurePayrollRangeControls(ss) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

  summarySheet.getRange("D1:D3").setValues([
    ["Summary Range"],
    ["Custom Start"],
    ["Custom End"]
  ]);

  const rangeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Selected Period", "Custom Dates"], true)
    .setAllowInvalid(false)
    .setHelpText("Choose the selected pay period or enter a custom date range.")
    .build();
  const rangeCell = summarySheet.getRange("E1");
  rangeCell.setDataValidation(rangeRule);
  if (!rangeCell.getValue()) rangeCell.setValue("Selected Period");

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Choose a custom start or end date.")
    .build();

  summarySheet.getRange("E2:E3")
    .setDataValidation(dateRule)
    .setNumberFormat("dd-mmm-yyyy");
}

function ensurePayrollSummaryFormulas(ss) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

  const dataA = summarySheet.getRange("A:A").getValues();
  let totalRowIndex = dataA.findIndex(row => {
    const value = String(row[0]).trim().toLowerCase();
    return value.includes("total payroll") || value === "total";
  });
  totalRowIndex = totalRowIndex >= 0 ? totalRowIndex + 1 : summarySheet.getLastRow() + 1;

  for (let row = 6; row < totalRowIndex; row++) {
    if (String(dataA[row - 1]?.[0] || "").trim()) {
      setPayrollRangeFormulas(summarySheet, row);
    }
  }
}

function setPayrollRangeFormulas(summarySheet, row) {
  const dateRange = `INDIRECT("'" & A${row} & "'!B2:B")`;
  const hoursRange = `INDIRECT("'" & A${row} & "'!E2:E")`;
  const notesRange = `INDIRECT("'" & A${row} & "'!F2:F")`;
  const startDate = `TEXT(IF($E$1="Custom Dates",$E$2,$B$2),"yyyy-mm-dd")`;
  const endDate = `TEXT(IF($E$1="Custom Dates",$E$3,$B$3),"yyyy-mm-dd")`;

  const hoursFormula = `=IFERROR(SUM(FILTER(${hoursRange}, ${dateRange} >= ${startDate}, ${dateRange} <= ${endDate})), 0)`;
  summarySheet.getRange(row, 2).setFormula(hoursFormula);

  const notesFormula = `=IFERROR(TEXTJOIN(" // ", TRUE, FILTER(${notesRange}, ${dateRange} >= ${startDate}, ${dateRange} <= ${endDate}, ${notesRange} <> "")), "")`;
  summarySheet.getRange(row, 3).setFormula(notesFormula);
}

function ensureAverageHeader(summarySheet) {
  summarySheet.getRange(5, 3).copyTo(
    summarySheet.getRange(5, 4),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );
  summarySheet.getRange(5, 4).setValue("30 Day Avg");
}

function setAverageFormula(summarySheet, row) {
  const dateRange = `INDIRECT("'" & A${row} & "'!B2:B")`;
  const hoursRange = `INDIRECT("'" & A${row} & "'!E2:E")`;
  const dailyHours = `QUERY({${dateRange},${hoursRange}}, "select Col1, sum(Col2) where Col1 >= '" & TEXT(TODAY()-29, "yyyy-mm-dd") & "' and Col1 <= '" & TEXT(TODAY(), "yyyy-mm-dd") & "' and Col1 is not null group by Col1 label sum(Col2) ''", 0)`;
  const formula = `=IFERROR(AVERAGE(QUERY(${dailyHours}, "select Col2 where Col2 >= 5", 0)), "")`;
  summarySheet.getRange(row, 4).setFormula(formula).setNumberFormat("0.00");
}
