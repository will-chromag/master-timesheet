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

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", entries: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getRange(2, 2, lastRow - 1, 5).getValues();
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
          notes: String(data[i][4] || "")
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
    const staffName = (data.staffName || "Unknown").trim();
    const targetDate = String(data.dateStr).trim();

    let sheet = ensureStaffTab(ss, staffName);
    registerStaffInSummary(ss, staffName);

    const lastRow = sheet.getLastRow();
    let rowIndex = -1;

    if (lastRow > 1) {
      const dateValues = sheet.getRange(1, 2, lastRow, 1).getValues();
      for (let i = 1; i < dateValues.length; i++) {
        let val = dateValues[i][0];
        let valStr = "";
        if (val instanceof Date) {
          valStr = Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd");
        } else {
          valStr = String(val).replace(/^'/, "").trim();
        }

        if (valStr === targetDate) {
          rowIndex = i + 1;
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
      data.notes || ""
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, 6).setValues([rowValues]);
    } else {
      const targetRow = lastRow + 1;
      sheet.getRange(targetRow, 2).setNumberFormat("@");
      sheet.getRange(targetRow, 1, 1, 6).setValues([rowValues]);
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
  return sheet;
}

function registerStaffInSummary(ss, staffName) {
  const summarySheet = ss.getSheetByName("Payroll Summary");
  if (!summarySheet) return;

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
      return;
    }
  }

  const insertAtRow = totalRowIndex > 0 ? totalRowIndex : summarySheet.getLastRow() + 1;
  if (totalRowIndex > 0) {
    summarySheet.insertRowBefore(insertAtRow);
  }

  summarySheet.getRange(insertAtRow, 1).setValue(staffName);

  const hoursFormula = `=IFERROR(SUM(FILTER(INDIRECT("'" & A${insertAtRow} & "'!E2:E"), INDIRECT("'" & A${insertAtRow} & "'!B2:B") >= TEXT($B$2, "yyyy-mm-dd"), INDIRECT("'" & A${insertAtRow} & "'!B2:B") <= TEXT($B$3, "yyyy-mm-dd"))), 0)`;
  summarySheet.getRange(insertAtRow, 2).setFormula(hoursFormula);

  const notesFormula = `=IFERROR(TEXTJOIN(" // ", TRUE, FILTER(INDIRECT("'" & A${insertAtRow} & "'!F2:F"), INDIRECT("'" & A${insertAtRow} & "'!B2:B") >= TEXT($B$2, "yyyy-mm-dd"), INDIRECT("'" & A${insertAtRow} & "'!B2:B") <= TEXT($B$3, "yyyy-mm-dd"), INDIRECT("'" & A${insertAtRow} & "'!F2:F") <> "")), "")`;
  summarySheet.getRange(insertAtRow, 3).setFormula(notesFormula);

  SpreadsheetApp.flush();
}
