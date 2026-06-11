const SETTINGS = {
  WORKERS_SHEET: "Workers",
  LOG_SHEET: "Time Log",
  MANUAL_SHEET: "Manual Adjustments",
  TIMEZONE: "America/Toronto",
};

function setup() {
  setupSheets_(SpreadsheetApp.getActiveSpreadsheet(), true);
}

function refreshAllWorkerSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets_(spreadsheet, false);
  saveAllManualAdjustments_(spreadsheet);

  getWorkerNames_(spreadsheet).forEach((workerName) => {
    rebuildWorkerSheet_(spreadsheet, workerName);
  });
}

function installStableWorkerHoursFix() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const workerNames = getWorkerNames_(spreadsheet);
  const activeWorkerNames = getActiveWorkerNames_(spreadsheet)
    .filter((workerName) => !isRepairProtectedSheet_(getWorkerSheetName_(workerName)));

  backupWorkerSheets_(spreadsheet, workerNames);
  setupManualAdjustmentsSheet_(spreadsheet);
  clearManualAdjustments_(spreadsheet);

  activeWorkerNames.forEach((workerName) => {
    rebuildWorkerSheet_(spreadsheet, workerName, { preserveManualValues: false });
  });
}

function doGet(e) {
  const callback = sanitizeCallback_(e.parameter.callback || "callback");
  const response = handleRequest_(e.parameter);
  return ContentService.createTextOutput(`${callback}(${JSON.stringify(response)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handleRequest_(params) {
  try {
    const action = String(params.action || "").trim();
    const name = String(params.name || "").trim();
    const pin = String(params.pin || "").trim();

    if (action !== "CHECK_IN" && action !== "CHECK_OUT" && action !== "CHECK_STATUS") {
      return fail_("Choose check in or check out.");
    }

    if (!pin) {
      return fail_("Worker PIN is required.");
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets_(spreadsheet, false);

    const worker = findWorker_(spreadsheet, name, pin);
    if (!worker) {
      return fail_("Name or PIN is incorrect.");
    }

    const currentStatus = getWorkerStatus_(spreadsheet, worker.name);
    if (action === "CHECK_STATUS") {
      return {
        ok: true,
        name: worker.name,
        lastAction: currentStatus.lastAction,
        nextAction: currentStatus.lastAction === "Check In" ? "CHECK_OUT" : "CHECK_IN",
        message: currentStatus.lastAction === "Check In" ? "Ready to check out." : "Ready to check in.",
      };
    }

    if (action === "CHECK_IN" && currentStatus.lastAction === "Check In") {
      return fail_(`${worker.name} is already checked in. Please check out first.`, "CHECK_OUT");
    }

    if (action === "CHECK_OUT" && currentStatus.lastAction !== "Check In") {
      return fail_(`${worker.name} is not checked in. Please check in first.`, "CHECK_IN");
    }

    const now = new Date();
    const displayDate = Utilities.formatDate(now, SETTINGS.TIMEZONE, "EEE MMM d yyyy");
    const displayTime = Utilities.formatDate(now, SETTINGS.TIMEZONE, "HH:mm:ss");
    const lat = String(params.latitude || "").trim();
    const lng = String(params.longitude || "").trim();
    const mapLink = lat && lng ? `https://www.google.com/maps?q=${lat},${lng}` : "";

    spreadsheet.getSheetByName(SETTINGS.LOG_SHEET).appendRow([
      now,
      displayDate,
      displayTime,
      action === "CHECK_IN" ? "Check In" : "Check Out",
      worker.name,
      String(params.job || "").trim(),
      String(params.notes || "").trim(),
      lat,
      lng,
      String(params.accuracyMeters || "").trim(),
      mapLink,
      String(params.clientLocalTime || "").trim(),
      String(params.timezone || "").trim(),
      String(params.userAgent || "").trim(),
    ]);
    formatLogSheet_(spreadsheet.getSheetByName(SETTINGS.LOG_SHEET));
    rebuildWorkerSheet_(spreadsheet, worker.name);

    return {
      ok: true,
      name: worker.name,
      action,
      date: displayDate,
      time: displayTime,
      nextAction: action === "CHECK_IN" ? "CHECK_OUT" : "CHECK_IN",
      message: "Saved.",
    };
  } catch (error) {
    return fail_(error.message || "Could not save entry.");
  }
}

function setupSheets_(spreadsheet, populateWorkerSheets) {
  let workersSheet = spreadsheet.getSheetByName(SETTINGS.WORKERS_SHEET);
  if (!workersSheet) {
    workersSheet = spreadsheet.insertSheet(SETTINGS.WORKERS_SHEET);
  }

  if (workersSheet.getLastRow() === 0) {
    workersSheet.appendRow(["Name", "PIN", "Active"]);
    workersSheet.appendRow(["Example Worker", "1234", "TRUE"]);
  }

  let logSheet = spreadsheet.getSheetByName(SETTINGS.LOG_SHEET);
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(SETTINGS.LOG_SHEET);
  }

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "Timestamp",
      "Date",
      "Time",
      "Action",
      "Name",
      "Job",
      "Notes",
      "Latitude",
      "Longitude",
      "Accuracy Meters",
      "Map Link",
      "Worker Local Time",
      "Worker Timezone",
      "Device",
    ]);
    logSheet.setFrozenRows(1);
  }
  formatLogSheet_(logSheet);
  setupManualAdjustmentsSheet_(spreadsheet);

  if (populateWorkerSheets) {
    refreshAllWorkerSheets();
  }
}

function rebuildWorkerSheet_(spreadsheet, workerName, options) {
  const shouldPreserveManualValues = !options || options.preserveManualValues !== false;
  const sheetName = getWorkerSheetName_(workerName);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (shouldPreserveManualValues) {
    saveManualAdjustmentsForWorker_(spreadsheet, workerName);
  }
  const manualValues = shouldPreserveManualValues ? readManualAdjustments_(spreadsheet, workerName) : {};
  const logRange = spreadsheet.getSheetByName(SETTINGS.LOG_SHEET).getDataRange();
  const logValues = logRange.getValues();
  const logDisplayValues = logRange.getDisplayValues();
  const entries = [];
  const days = {};
  const targetName = normalizeName_(workerName);

  for (let row = 1; row < logValues.length; row += 1) {
    const timestamp = logValues[row][0];
    const date = String(logDisplayValues[row][1] || "").trim();
    const time = String(logDisplayValues[row][2] || "").trim();
    const action = String(logValues[row][3] || "").trim();
    const name = String(logValues[row][4] || "").trim();
    const job = String(logValues[row][5] || "").trim();
    const notes = String(logValues[row][6] || "").trim();
    const lat = String(logValues[row][7] || "").trim();
    const lng = String(logValues[row][8] || "").trim();
    const accuracy = String(logValues[row][9] || "").trim();
    const mapLink = String(logValues[row][10] || "").trim();

    if (normalizeName_(name) !== targetName) {
      continue;
    }

    entries.push([date, time, action, job, notes, lat, lng, accuracy, mapLink]);

    if (!date || !(timestamp instanceof Date)) {
      continue;
    }

    const dateKey = Utilities.formatDate(timestamp, SETTINGS.TIMEZONE, "yyyy-MM-dd");
    if (!days[dateKey]) {
      days[dateKey] = {
        displayDate: date,
        events: [],
        jobs: [],
        notes: [],
      };
    }

    days[dateKey].events.push({ timestamp, action });
    if (job) days[dateKey].jobs.push(job);
    if (notes) days[dateKey].notes.push(notes);
  }

  const summaryRows = Object.keys(days)
    .sort()
    .map((dateKey) => {
      const day = days[dateKey];
      const date = day.displayDate;
      day.events.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());

      let openCheckIn = null;
      let firstCheckIn = null;
      let lastCheckOut = null;
      let totalMs = 0;

      day.events.forEach((event) => {
        if (event.action === "Check In") {
          openCheckIn = event.timestamp;
          if (!firstCheckIn) firstCheckIn = event.timestamp;
        }

        if (event.action === "Check Out") {
          lastCheckOut = event.timestamp;
          if (openCheckIn && event.timestamp > openCheckIn) {
            totalMs += event.timestamp.getTime() - openCheckIn.getTime();
            openCheckIn = null;
          }
        }
      });

      const manual = manualValues[getManualKey_(workerName, date)] || {};
      const autoHours = totalMs ? Math.round((totalMs / 36e5) * 100) / 100 : "";
      const autoPayableHours = autoHours;
      const status = openCheckIn ? "Checked In" : "Complete";

      return [
        date,
        firstCheckIn ? Utilities.formatDate(firstCheckIn, SETTINGS.TIMEZONE, "h:mm a") : "",
        lastCheckOut ? Utilities.formatDate(lastCheckOut, SETTINGS.TIMEZONE, "h:mm a") : "",
        autoHours,
        autoPayableHours,
        manual.manualHours || "",
        "",
        manual.payRate || "",
        "",
        status,
        uniqueText_(day.jobs),
        manual.payNotes || "",
      ];
    });

  sheet.clearContents();
  writeWorkerHeaders_(sheet);

  if (entries.length) {
    sheet.getRange(3, 1, entries.length, entries[0].length).setValues(entries);
  }

  if (summaryRows.length) {
    sheet.getRange(3, 11, summaryRows.length, summaryRows[0].length).setValues(summaryRows);
    const finalHourFormulas = summaryRows.map((row, index) => {
      const rowNumber = index + 3;
      return [`=IF(OR(O${rowNumber}<>"",P${rowNumber}<>""),N(O${rowNumber})+N(P${rowNumber}),"")`];
    });
    const payAmountFormulas = summaryRows.map((row, index) => {
      const rowNumber = index + 3;
      return [`=IF(R${rowNumber}<>"",Q${rowNumber}*R${rowNumber},"")`];
    });
    sheet.getRange(3, 17, finalHourFormulas.length, 1).setFormulas(finalHourFormulas);
    sheet.getRange(3, 19, payAmountFormulas.length, 1).setFormulas(payAmountFormulas);
  }

  insertWeeklyGapRows_(sheet);
  sheet.setFrozenRows(2);
  formatWorkerSheet_(sheet);
  sheet.autoResizeColumns(1, 22);
}

function formatLogSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange("B:C").setNumberFormat("@");
  sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat("M/d/yyyy HH:mm:ss");
}

function formatWorkerSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 3);
  sheet.getRange("A:B").setNumberFormat("@");
  sheet.getRange("K:M").setNumberFormat("@");
  sheet.getRange(3, 1, lastRow - 2, 1).setNumberFormat("@");
}

function writeWorkerHeaders_(sheet) {
  sheet.getRange(1, 1).setValue("Time Entries");
  sheet.getRange(1, 11).setValue("Daily Summary");
  sheet.getRange(2, 1, 1, 9).setValues([[
    "Date",
    "Time",
    "Action",
    "Job",
    "Notes",
    "Latitude",
    "Longitude",
    "Accuracy Meters",
    "Map Link",
  ]]);
  sheet.getRange(2, 11, 1, 12).setValues([[
    "Date",
    "First Check In",
    "Last Check Out",
    "Auto Hours",
    "Auto Payable Hours",
    "Manual Hours",
    "Final Payable Hours",
    "Pay Rate",
    "Pay Amount",
    "Status",
    "Jobs",
    "Pay Notes",
  ]]);
  sheet.getRange(1, 1, 2, 22).setFontWeight("bold");
}

function setupManualAdjustmentsSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SETTINGS.MANUAL_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SETTINGS.MANUAL_SHEET);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Key", "Worker Name", "Date", "Manual Hours", "Pay Rate", "Pay Notes", "Updated At"]);
    sheet.setFrozenRows(1);
  }
}

function clearManualAdjustments_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SETTINGS.MANUAL_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

function saveAllManualAdjustments_(spreadsheet) {
  getWorkerNames_(spreadsheet).forEach((workerName) => {
    saveManualAdjustmentsForWorker_(spreadsheet, workerName);
  });
}

function saveManualAdjustmentsForWorker_(spreadsheet, workerName) {
  setupManualAdjustmentsSheet_(spreadsheet);
  const sheet = spreadsheet.getSheetByName(getWorkerSheetName_(workerName));
  if (!sheet || sheet.getLastRow() < 3) {
    return;
  }

  const values = sheet.getDataRange().getDisplayValues();
  const updates = {};

  for (let row = 2; row < values.length; row += 1) {
    const date = String(values[row][10] || "").trim();
    if (!date) continue;

    const manualHours = String(values[row][15] || "").trim();
    const payRate = String(values[row][17] || "").trim();
    const payNotes = String(values[row][21] || "").trim();

    if (manualHours || payRate || payNotes) {
      updates[getManualKey_(workerName, date)] = {
        workerName,
        date,
        manualHours,
        payRate,
        payNotes,
      };
    }
  }

  upsertManualAdjustments_(spreadsheet, updates);
}

function readManualAdjustments_(spreadsheet, workerName) {
  setupManualAdjustmentsSheet_(spreadsheet);
  const sheet = spreadsheet.getSheetByName(SETTINGS.MANUAL_SHEET);
  const values = sheet.getDataRange().getDisplayValues();
  const targetName = normalizeName_(workerName);
  const adjustments = {};

  for (let row = 1; row < values.length; row += 1) {
    const key = String(values[row][0] || "").trim();
    const rowWorkerName = String(values[row][1] || "").trim();
    if (!key || normalizeName_(rowWorkerName) !== targetName) continue;

    adjustments[key] = {
      manualHours: String(values[row][3] || "").trim(),
      payRate: String(values[row][4] || "").trim(),
      payNotes: String(values[row][5] || "").trim(),
    };
  }

  return adjustments;
}

function upsertManualAdjustments_(spreadsheet, updates) {
  const keys = Object.keys(updates);
  if (!keys.length) return;

  const sheet = spreadsheet.getSheetByName(SETTINGS.MANUAL_SHEET);
  const values = sheet.getDataRange().getValues();
  const rowByKey = {};

  for (let row = 1; row < values.length; row += 1) {
    const key = String(values[row][0] || "").trim();
    if (key) rowByKey[key] = row + 1;
  }

  keys.forEach((key) => {
    const update = updates[key];
    const row = [
      key,
      update.workerName,
      update.date,
      update.manualHours,
      update.payRate,
      update.payNotes,
      new Date(),
    ];

    if (rowByKey[key]) {
      sheet.getRange(rowByKey[key], 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  });
}

function getManualKey_(workerName, date) {
  return `${normalizeName_(workerName)}|${String(date || "").trim()}`;
}

function getWorkerNames_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SETTINGS.WORKERS_SHEET);
  const values = sheet.getDataRange().getValues();
  const names = [];

  for (let row = 1; row < values.length; row += 1) {
    const workerName = String(values[row][0] || "").trim();
    if (workerName) names.push(workerName);
  }

  return names;
}

function getActiveWorkerNames_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SETTINGS.WORKERS_SHEET);
  const values = sheet.getDataRange().getValues();
  const names = [];

  for (let row = 1; row < values.length; row += 1) {
    const workerName = String(values[row][0] || "").trim();
    const active = String(values[row][2] || "TRUE").trim().toUpperCase();
    if (workerName && active !== "FALSE") names.push(workerName);
  }

  return names;
}

function backupWorkerSheets_(spreadsheet, workerNames) {
  const backupStamp = Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, "yyyyMMdd HHmmss");

  workerNames.forEach((workerName) => {
    const sheetName = getWorkerSheetName_(workerName);
    if (isRepairProtectedSheet_(sheetName)) return;

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    sheet.copyTo(spreadsheet).setName(getUniqueBackupSheetName_(spreadsheet, sheetName, backupStamp));
  });
}

function getUniqueBackupSheetName_(spreadsheet, sheetName, backupStamp) {
  const baseName = `${sheetName} Backup ${backupStamp}`.slice(0, 90);
  let backupName = baseName;
  let count = 2;

  while (spreadsheet.getSheetByName(backupName)) {
    backupName = `${baseName} ${count}`.slice(0, 99);
    count += 1;
  }

  return backupName;
}

function isRepairProtectedSheet_(sheetName) {
  return [
    SETTINGS.WORKERS_SHEET,
    SETTINGS.LOG_SHEET,
    "Rishav Early Sheet",
  ].indexOf(sheetName) !== -1 || /backup/i.test(sheetName);
}

function insertWeeklyGapRows_(sheet) {
  let row = 3;
  while (row <= sheet.getLastRow()) {
    const dateValue = String(sheet.getRange(row, 11).getDisplayValue() || "").trim();
    if (!dateValue) {
      row += 1;
      continue;
    }

    const date = parseDisplayDate_(dateValue);
    const nextDateValue = String(sheet.getRange(row + 1, 11).getDisplayValue() || "").trim();
    const nextDate = parseDisplayDate_(nextDateValue);

    if (date && nextDate && getMondayWeekStart_(nextDate) > getMondayWeekStart_(date)) {
      sheet.insertRowsAfter(row, 1);
      row += 2;
    } else {
      row += 1;
    }
  }
}

function parseDisplayDate_(dateText) {
  if (!dateText) return null;
  const date = new Date(dateText);
  return Number.isNaN(date.getTime()) ? null : date;
}

function getMondayWeekStart_(date) {
  const weekStart = new Date(date);
  const daysSinceMonday = (weekStart.getDay() + 6) % 7;
  weekStart.setDate(weekStart.getDate() - daysSinceMonday);
  weekStart.setHours(0, 0, 0, 0);
  return weekStart.getTime();
}

function getWorkerSheetName_(workerName) {
  const cleaned = String(workerName || "Worker")
    .replace(/[\[\]\*\?\/\\:]/g, "-")
    .trim()
    .slice(0, 90);
  const sheetName = cleaned || "Worker";
  const reservedNames = [SETTINGS.WORKERS_SHEET, SETTINGS.LOG_SHEET];
  return reservedNames.indexOf(sheetName) === -1 ? sheetName : `${sheetName} Worker`;
}

function uniqueText_(items) {
  return [...new Set(items)].join("; ");
}

function findWorker_(spreadsheet, name, pin) {
  const sheet = spreadsheet.getSheetByName(SETTINGS.WORKERS_SHEET);
  const values = sheet.getDataRange().getValues();
  const wantedName = normalizeName_(name);
  const wantedPin = String(pin);

  for (let row = 1; row < values.length; row += 1) {
    const workerName = String(values[row][0] || "").trim();
    const workerPin = String(values[row][1] || "").trim();
    const active = String(values[row][2] || "TRUE").trim().toUpperCase();

    const nameMatches = !wantedName || normalizeName_(workerName) === wantedName;
    if (nameMatches && workerPin === wantedPin && active !== "FALSE") {
      return { name: workerName };
    }
  }

  return null;
}

function getWorkerStatus_(spreadsheet, workerName) {
  const sheet = spreadsheet.getSheetByName(SETTINGS.LOG_SHEET);
  const values = sheet.getDataRange().getValues();
  const targetName = normalizeName_(workerName);

  for (let row = values.length - 1; row >= 1; row -= 1) {
    const name = String(values[row][4] || "").trim();
    if (normalizeName_(name) === targetName) {
      return {
        lastAction: String(values[row][3] || "").trim(),
        lastTimestamp: values[row][0],
      };
    }
  }

  return {
    lastAction: "",
    lastTimestamp: null,
  };
}

function normalizeName_(name) {
  return String(name || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function sanitizeCallback_(callback) {
  return /^[a-zA-Z_$][\w$]*(\.[a-zA-Z_$][\w$]*)*$/.test(callback) ? callback : "callback";
}

function fail_(message, nextAction) {
  return {
    ok: false,
    message,
    nextAction: nextAction || "",
  };
}
