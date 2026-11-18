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

      ensurePayrollSummaryLayout(ss);
      ensurePayrollAverageColumn(ss);
      ensurePayrollSummaryFormulas(ss);
      ensureAllPayrollSheet(ss);
      refreshAllPayroll(ss);

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

    // Verify employee PINs against the current Staff sheet, never a browser copy.
    if (data.action === "verifyStaffPin") {
      const staffName = String(data.staffName || "").trim();
      const pin = String(data.pin || "").trim();
      const staffSheet = ss.getSheetByName("Staff");

      if (!staffName || !/^\d{4}$/.test(pin)) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Incorrect name or PIN" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      if (!staffSheet) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Staff roster is unavailable" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const lastRow = staffSheet.getLastRow();
      const rows = lastRow > 1 ? staffSheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
      const matches = rows.filter(row => String(row[0] || "").trim() === staffName);
      const storedPin = matches.length === 1
        ? String(matches[0][1] || "").replace(/^'/, "").trim()
        : "";

      if (storedPin && storedPin === pin) {
        return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Incorrect name or PIN" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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

      ensureAllPayrollSheet(ss);
      refreshAllPayroll(ss);
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
            refreshAllPayroll(ss);
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

    refreshAllPayroll(ss);
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

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "All Payroll") return;
  if (e.range.getColumn() !== 2 || e.range.getRow() < 1 || e.range.getRow() > 3) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    refreshAllPayroll(e.source);
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
  ensurePayrollSummaryLayout(ss);
  ensureAverageHeader(summarySheet);
  const headerRow = getPayrollTableHeaderRow(summarySheet);

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
  for (let r = headerRow; r < searchLimit; r++) {
    if (String(dataA[r][0]).trim().toLowerCase() === staffName.toLowerCase()) {
      setPayrollPeriodFormulas(summarySheet, r + 1);
      setAverageFormula(summarySheet, r + 1);
      return;
    }
  }

  const insertAtRow = totalRowIndex > 0 ? totalRowIndex : summarySheet.getLastRow() + 1;
  if (totalRowIndex > 0) {
    summarySheet.insertRowBefore(insertAtRow);
  }

  summarySheet.getRange(insertAtRow, 1).setValue(staffName);
  setPayrollPeriodFormulas(summarySheet, insertAtRow);
  setAverageFormula(summarySheet, insertAtRow);

  SpreadsheetApp.flush();
}

function ensurePayrollAverageColumn(ss) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

  ensureAverageHeader(summarySheet);
  const headerRow = getPayrollTableHeaderRow(summarySheet);
  const dataA = summarySheet.getRange("A:A").getValues();
  let totalRowIndex = dataA.findIndex(row => {
    const value = String(row[0]).trim().toLowerCase();
    return value.includes("total payroll") || value === "total";
  });
  totalRowIndex = totalRowIndex >= 0 ? totalRowIndex + 1 : summarySheet.getLastRow() + 1;

  for (let row = headerRow + 1; row < totalRowIndex; row++) {
    if (String(dataA[row - 1]?.[0] || "").trim()) {
      setAverageFormula(summarySheet, row);
    }
  }
}

function ensurePayrollSummaryLayout(ss) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

  const headerValues = summarySheet.getRange("A5:A6").getValues().flat();
  if (String(headerValues[0]).trim().toLowerCase() === "employee") {
    summarySheet.insertRowBefore(5);
  }

  const tableHeaderRow = getPayrollTableHeaderRow(summarySheet);
  const periodHeading = summarySheet.getRange("A5:D5");
  periodHeading.breakApart();
  summarySheet.getRange(tableHeaderRow, 1, 1, 4).copyTo(
    periodHeading,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );
  periodHeading.merge();
  periodHeading.setFormula(getPayrollPeriodHeadingFormula("$B$1", "$B$2", "$B$3"));
  const oldPeriodHeading = summarySheet.getRange("A4:D4");
  oldPeriodHeading.breakApart();
  oldPeriodHeading.clearContent();

  const lastRow = summarySheet.getLastRow();
  const periodLabels = lastRow >= 2
    ? summarySheet.getRange(2, 6, lastRow - 1, 1).getValues()
      .map(row => String(row[0]).trim())
      .filter(Boolean)
    : [];
  if (!periodLabels.length) {
    throw new Error("No payroll periods were found in Payroll Summary column F.");
  }

  const currentPeriod = periodLabels.find(label => label.toLowerCase().includes("current period")) || periodLabels[0];
  const otherPeriods = periodLabels.filter(label =>
    label !== currentPeriod && label.toLowerCase() !== "year to date"
  );
  const selectorRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([currentPeriod, "Year to Date", ...otherPeriods, "Custom Range"], true)
    .setAllowInvalid(false)
    .build();
  summarySheet.getRange("B1").setDataValidation(selectorRule);

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Enter a custom start or end date when Custom Range is selected.")
    .build();
  summarySheet.getRange("B2:B3")
    .setDataValidation(dateRule)
    .setNumberFormat("dd-mmm-yyyy");
}

function ensureAllPayrollSheet(ss) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

  let allPayrollSheet = ss.getSheetByName("All Payroll");
  if (!allPayrollSheet) {
    allPayrollSheet = ss.insertSheet("All Payroll");
  }

  allPayrollSheet.getRange("A1:A3").setValues([
    ["Pay Period:"],
    ["Custom Start Date:"],
    ["Custom End Date:"]
  ]);
  allPayrollSheet.getRange("A4:H4").breakApart().merge();
  allPayrollSheet.getRange("A5:H5").setValues([[
    "Notes", "User", "Date", "Start Time", "End Time",
    "Duration (h)", "Duration (decimal)", "Last Updated"
  ]]).setFontWeight("bold");
  allPayrollSheet.setFrozenRows(5);

  const lastSummaryRow = summarySheet.getLastRow();
  const periodLabels = (lastSummaryRow >= 2 ? summarySheet.getRange(2, 6, lastSummaryRow - 1, 1).getValues() : [])
    .map(row => String(row[0]).trim())
    .filter(Boolean);
  if (!periodLabels.length) {
    throw new Error("No payroll periods were found in Payroll Summary column F.");
  }

  const currentPeriod = periodLabels.find(label => label.toLowerCase().includes("current period")) || periodLabels[0];
  const otherPeriods = periodLabels.filter(label =>
    label !== currentPeriod && label.toLowerCase() !== "year to date"
  );
  const selectorRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([currentPeriod, "Year to Date", ...otherPeriods, "Custom Range"], true)
    .setAllowInvalid(false)
    .build();
  allPayrollSheet.getRange("B1").setDataValidation(selectorRule);

  const summarySelection = String(summarySheet.getRange("B1").getValue() || "").trim();
  if (!allPayrollSheet.getRange("B1").getValue()) {
    allPayrollSheet.getRange("B1").setValue(
      summarySelection || currentPeriod
    );
    if (summarySelection === "Custom Range") {
      allPayrollSheet.getRange("B2:B3").setValues(summarySheet.getRange("B2:B3").getValues());
    }
  }

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Enter a custom start or end date when Custom Range is selected.")
    .build();
  allPayrollSheet.getRange("B2:B3")
    .setDataValidation(dateRule)
    .setNumberFormat("dd-mmm-yyyy");
}

function refreshAllPayroll(ss) {
  const allPayrollSheet = ss.getSheetByName("All Payroll");
  const staffSheet = ss.getSheetByName("Staff");
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!allPayrollSheet || !staffSheet || !summarySheet) return;

  const timeZone = Session.getScriptTimeZone() || "GMT";
  const selection = String(allPayrollSheet.getRange("B1").getValue() || "").trim();
  const range = getAllPayrollDateRange(summarySheet, selection, allPayrollSheet, timeZone);
  const staffLastRow = staffSheet.getLastRow();
  const staffNames = staffLastRow > 1
    ? staffSheet.getRange(2, 1, staffLastRow - 1, 1).getValues()
      .map(row => String(row[0] || "").trim())
      .filter(Boolean)
    : [];
  const output = [];
  const subtotalRows = [];
  let grandTotal = 0;

  staffNames.forEach(staffName => {
    const staffTab = ss.getSheetByName(staffName);
    if (!staffTab || staffTab.getLastRow() <= 1) return;

    const rows = staffTab.getRange(2, 1, staffTab.getLastRow() - 1, 7).getValues();
    const entries = [];
    rows.forEach(row => {
      const date = parsePayrollDate(row[1], timeZone);
      if (!date || !range.start || !range.end || date < range.start || date > range.end) return;

      const hours = Number(row[4]) || 0;
      entries.push({
        date,
        startTime: formatPayrollTime(row[2], timeZone),
        endTime: formatPayrollTime(row[3], timeZone),
        hours,
        notes: String(row[5] || ""),
        lastUpdated: row[0] instanceof Date ? row[0] : ""
      });
    });
    if (!entries.length) return;

    entries.sort((a, b) => a.date - b.date);
    let staffTotal = 0;
    entries.forEach(entry => {
      staffTotal += entry.hours;
      output.push([
        entry.notes,
        staffName,
        entry.date,
        entry.startTime,
        entry.endTime,
        entry.hours / 24,
        entry.hours,
        entry.lastUpdated
      ]);
    });
    grandTotal += staffTotal;
    output.push(["", "", "", "", "", "", staffTotal, ""]);
    subtotalRows.push(output.length);
    output.push(["", "", "", "", "", "", "", ""]);
  });

  if (output.length && output[output.length - 1].every(value => value === "")) {
    output.pop();
  }
  if (output.length) {
    output.push(["", "Total Payroll Hours", "", "", "", "", grandTotal, ""]);
  }

  const previousLastRow = Math.max(allPayrollSheet.getLastRow(), 6);
  const rowsToClear = Math.max(previousLastRow - 5, output.length);
  if (rowsToClear > 0) {
    allPayrollSheet.getRange(6, 1, rowsToClear, 8).clearContent();
    allPayrollSheet.getRange(6, 1, rowsToClear, 8)
      .setBackground("#ffffff")
      .setFontColor("#000000")
      .setFontWeight("normal");
  }
  if (output.length) {
    allPayrollSheet.getRange(6, 1, output.length, 8).setValues(output);
    allPayrollSheet.getRange(6, 3, output.length, 1).setNumberFormat("dd/mm/yyyy");
    allPayrollSheet.getRange(6, 4, output.length, 2).setNumberFormat("hh:mm:ss AM/PM");
    allPayrollSheet.getRange(6, 6, output.length, 1).setNumberFormat("[hh]:mm:ss");
    allPayrollSheet.getRange(6, 7, output.length, 1).setNumberFormat("0.00");
    allPayrollSheet.getRange(6, 8, output.length, 1).setNumberFormat("dd/mm/yyyy");
    subtotalRows.forEach(row => {
      allPayrollSheet.getRange(5 + row, 7)
        .setBackground("#ffff00")
        .setFontColor("#000000")
        .setFontWeight("bold");
    });
    allPayrollSheet.getRange(5 + output.length, 1, 1, 8)
      .setBackground("#1a73e8")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
  }

  const heading = !range.start || !range.end
    ? selection === "Custom Range"
      ? "Period: enter custom start and end dates"
      : "Period: select a valid range"
    : `Period: ${Utilities.formatDate(range.start, timeZone, "dd-MMM-yyyy")} --> ${Utilities.formatDate(range.end, timeZone, "dd-MMM-yyyy")}`;
  allPayrollSheet.getRange("A4").setValue(heading);
}

function getAllPayrollDateRange(summarySheet, selection, allPayrollSheet, timeZone) {
  if (selection === "Custom Range") {
    return {
      start: parsePayrollDate(allPayrollSheet.getRange("B2").getValue(), timeZone),
      end: parsePayrollDate(allPayrollSheet.getRange("B3").getValue(), timeZone)
    };
  }
  if (selection === "Year to Date") {
    const today = parsePayrollDate(new Date(), timeZone);
    return { start: new Date(today.getFullYear(), 0, 1), end: today };
  }

  const lastRow = summarySheet.getLastRow();
  const periodRows = lastRow >= 2 ? summarySheet.getRange(2, 6, lastRow - 1, 3).getValues() : [];
  const selectedPeriod = periodRows.find(row => String(row[0] || "").trim() === selection);
  if (!selectedPeriod) return { start: null, end: null };
  return {
    start: parsePayrollDate(selectedPeriod[1], timeZone),
    end: parsePayrollDate(selectedPeriod[2], timeZone)
  };
}

function parsePayrollDate(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const parts = Utilities.formatDate(value, timeZone, "yyyy-MM-dd").split("-").map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  const text = String(value || "").replace(/^'/, "").trim();
  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
  return null;
}

function formatPayrollTime(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone, "hh:mm:ss a");
  }
  const text = String(value || "").trim();
  const match = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return text;
  const hour = Number(match[1]);
  const suffix = hour >= 12 ? "PM" : "AM";
  const hour12 = hour % 12 || 12;
  return `${String(hour12).padStart(2, "0")}:${match[2]}:${match[3] || "00"} ${suffix}`;
}

function getPayrollPeriodHeadingFormula(selectorCell, startCell, endCell) {
  return `=IF(${selectorCell}="Custom Range",IF(OR(${startCell}="",${endCell}=""),"Period: enter custom start and end dates","Period: "&TEXT(${startCell},"dd-mmm-yyyy")&" --> "&TEXT(${endCell},"dd-mmm-yyyy")),IF(${selectorCell}="Year to Date","Period: "&TEXT(DATE(YEAR(TODAY()),1,1),"dd-mmm-yyyy")&" --> "&TEXT(TODAY(),"dd-mmm-yyyy"),IFERROR("Period: "&TEXT(INDEX($G$2:$G,MATCH(${selectorCell},$F$2:$F,0)),"dd-mmm-yyyy")&" --> "&TEXT(INDEX($H$2:$H,MATCH(${selectorCell},$F$2:$F,0)),"dd-mmm-yyyy"),"")))`;
}

function ensurePayrollSummaryFormulas(ss) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

  const controlLabels = summarySheet.getRange("D1:D3").getValues().flat();
  if (controlLabels[0] === "Summary Range" &&
      controlLabels[1] === "Custom Start" &&
      controlLabels[2] === "Custom End") {
    summarySheet.getRange("D1:E3").clearContent().clearDataValidations();
  }

  const dataA = summarySheet.getRange("A:A").getValues();
  let totalRowIndex = dataA.findIndex(row => {
    const value = String(row[0]).trim().toLowerCase();
    return value.includes("total payroll") || value === "total";
  });
  totalRowIndex = totalRowIndex >= 0 ? totalRowIndex + 1 : summarySheet.getLastRow() + 1;
  const headerRow = getPayrollTableHeaderRow(summarySheet);

  for (let row = headerRow + 1; row < totalRowIndex; row++) {
    if (String(dataA[row - 1]?.[0] || "").trim()) {
      setPayrollPeriodFormulas(summarySheet, row);
    }
  }
}

function setPayrollPeriodFormulas(summarySheet, row) {
  const dateRange = `INDIRECT("'" & A${row} & "'!B2:B")`;
  const hoursRange = `INDIRECT("'" & A${row} & "'!E2:E")`;
  const notesRange = `INDIRECT("'" & A${row} & "'!F2:F")`;
  const startDate = `TEXT(IF($B$1="Custom Range",$B$2,IF($B$1="Year to Date",DATE(YEAR(TODAY()),1,1),INDEX($G$2:$G,MATCH($B$1,$F$2:$F,0)))),"yyyy-mm-dd")`;
  const endDate = `TEXT(IF($B$1="Custom Range",$B$3,IF($B$1="Year to Date",TODAY(),INDEX($H$2:$H,MATCH($B$1,$F$2:$F,0)))),"yyyy-mm-dd")`;

  const hoursFormula = `=IFERROR(SUM(FILTER(${hoursRange}, ${dateRange} >= ${startDate}, ${dateRange} <= ${endDate})), 0)`;
  summarySheet.getRange(row, 2).setFormula(hoursFormula);

  const notesFormula = `=IFERROR(TEXTJOIN(" // ", TRUE, FILTER(${notesRange}, ${dateRange} >= ${startDate}, ${dateRange} <= ${endDate}, ${notesRange} <> "")), "")`;
  summarySheet.getRange(row, 3).setFormula(notesFormula);
}

function ensureAverageHeader(summarySheet) {
  const headerRow = getPayrollTableHeaderRow(summarySheet);
  summarySheet.getRange(headerRow, 3).copyTo(
    summarySheet.getRange(headerRow, 4),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );
  summarySheet.getRange(headerRow, 4).setValue("30 Day Avg");
}

function getPayrollTableHeaderRow(summarySheet) {
  const value = String(summarySheet.getRange("A6").getValue()).trim().toLowerCase();
  return value === "employee" ? 6 : 5;
}

function setAverageFormula(summarySheet, row) {
  const dateRange = `INDIRECT("'" & A${row} & "'!B2:B")`;
  const hoursRange = `INDIRECT("'" & A${row} & "'!E2:E")`;
  const dailyHours = `QUERY({${dateRange},${hoursRange}}, "select Col1, sum(Col2) where Col1 >= '" & TEXT(TODAY()-29, "yyyy-mm-dd") & "' and Col1 <= '" & TEXT(TODAY(), "yyyy-mm-dd") & "' and Col1 is not null group by Col1 label sum(Col2) ''", 0)`;
  const formula = `=IFERROR(AVERAGE(QUERY(${dailyHours}, "select Col2 where Col2 >= 5", 0)), "")`;
  summarySheet.getRange(row, 4).setFormula(formula).setNumberFormat("0.00");
}
