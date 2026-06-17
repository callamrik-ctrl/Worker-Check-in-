const SETTINGS = {
  WORKERS_SHEET: "Workers",
  LOG_SHEET: "Time Log",
  TIMEZONE: "America/Toronto",
};
const AUTO_CLOSED_BACKGROUND = "#f4cccc";
const AUTO_CLOSED_NOTE = "Worker forgot to check out. System closed this work date at 11:59 PM.";

function setup() {
  setupSheets_(SpreadsheetApp.getActiveSpreadsheet(), true);
}

function refreshAllWorkerSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets_(spreadsheet, false);

  getWorkerNames_(spreadsheet).forEach((workerName) => {
    rebuildWorkerSheet_(spreadsheet, workerName);
  });
}

function installStableWorkerHoursFix() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets_(spreadsheet, false);
  deleteManualAdjustmentsSheet_(spreadsheet);
  const activeWorkerNames = getActiveWorkerNames_(spreadsheet)
    .filter((workerName) => !isRepairProtectedSheet_(getWorkerSheetName_(workerName)));

  activeWorkerNames.forEach((workerName) => {
    rebuildWorkerSheet_(spreadsheet, workerName);
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
    setupSheets_(spreadsheet, false, { formatLog: action !== "CHECK_STATUS" });

    const worker = findWorker_(spreadsheet, name, pin);
    if (worker && worker.blocked) {
      return fail_("Your access is blocked. Please contact the company.");
    }

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

function setupSheets_(spreadsheet, populateWorkerSheets, options) {
  const shouldFormatLog = !options || options.formatLog !== false;
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
  if (shouldFormatLog) {
    formatLogSheet_(logSheet);
  }

  if (populateWorkerSheets) {
    refreshAllWorkerSheets();
  }
}

function rebuildWorkerSheet_(spreadsheet, workerName) {
  const sheetName = getWorkerSheetName_(workerName);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const annotations = captureWorkerAnnotations_(sheet, workerName);
  const logRange = spreadsheet.getSheetByName(SETTINGS.LOG_SHEET).getDataRange();
  const logValues = logRange.getValues();
  const logDisplayValues = logRange.getDisplayValues();
  const entries = [];
  const days = {};
  const targetName = normalizeName_(workerName);

  for (let row = 1; row < logValues.length; row += 1) {
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

    const eventTime = parseLogDateTime_(date, time);
    if (!eventTime) {
      continue;
    }

    const dateKey = Utilities.formatDate(eventTime, SETTINGS.TIMEZONE, "yyyy-MM-dd");
    if (!days[dateKey]) {
      days[dateKey] = {
        displayDate: date,
        events: [],
        jobs: [],
        notes: [],
      };
    }

    days[dateKey].events.push({ timestamp: eventTime, action });
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
      let autoClosed = false;

      day.events.forEach((event) => {
        if (event.action === "Check In") {
          if (!openCheckIn) {
            openCheckIn = event.timestamp;
            if (!firstCheckIn) firstCheckIn = event.timestamp;
          }
        }

        if (event.action === "Check Out") {
          lastCheckOut = event.timestamp;
          if (openCheckIn && event.timestamp > openCheckIn) {
            totalMs += event.timestamp.getTime() - openCheckIn.getTime();
            openCheckIn = null;
          }
        }
      });

      if (openCheckIn && shouldAutoCloseWorkDate_(dateKey)) {
        totalMs += getWorkDateEnd_(dateKey).getTime() - openCheckIn.getTime();
        lastCheckOut = getWorkDateEnd_(dateKey);
        openCheckIn = null;
        autoClosed = true;
      }

      const annotation = annotations.summaryValues[getSummaryAnnotationKey_(workerName, dateKey)] || {};
      const workHours = totalMs ? Math.round((totalMs / 36e5) * 100) / 100 : "";
      const breakHours = calculateBreakHours_(workHours);
      const payableHours = workHours ? Math.max(Math.round((workHours - breakHours) * 100) / 100, 0) : "";
      const status = autoClosed ? "Auto Closed" : (openCheckIn ? "Checked In" : "Complete");

      return [
        date,
        firstCheckIn ? Utilities.formatDate(firstCheckIn, SETTINGS.TIMEZONE, "h:mm a") : "",
        lastCheckOut ? Utilities.formatDate(lastCheckOut, SETTINGS.TIMEZONE, "h:mm a") : "",
        workHours,
        breakHours || "",
        payableHours,
        annotation.payRate || "",
        "",
        status,
        uniqueText_(day.jobs),
        annotation.notes || "",
      ];
    });

  clearRebuiltSheet_(sheet);
  writeWorkerHeaders_(sheet);

  const entryRows = addWeeklyGapRows_(entries, 0, 9);
  const summaryRowsWithGaps = addWeeklyGapRows_(summaryRows, 0, 11);

  if (entryRows.length) {
    sheet.getRange(3, 1, entryRows.length, entryRows[0].length).setValues(entryRows);
  }

  if (summaryRowsWithGaps.length) {
    sheet.getRange(3, 11, summaryRowsWithGaps.length, summaryRowsWithGaps[0].length).setValues(summaryRowsWithGaps);
    const payAmountFormulas = summaryRowsWithGaps.map((row, index) => {
      const rowNumber = index + 3;
      return [`=IF(Q${rowNumber}<>"",P${rowNumber}*Q${rowNumber},"")`];
    });
    sheet.getRange(3, 18, payAmountFormulas.length, 1).setFormulas(payAmountFormulas);
  }

  applyWorkerAnnotations_(sheet, workerName, annotations);
  applyAutoClosedFormatting_(sheet);
  sheet.setFrozenRows(2);
  formatWorkerSheet_(sheet);
  sheet.autoResizeColumns(1, 22);
}

function clearRebuiltSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 3);
  sheet.getRange(1, 1, lastRow, 22).clearContent();

  if (lastRow > 2) {
    const bodyRange = sheet.getRange(3, 1, lastRow - 2, 22);
    bodyRange.clearFormat();
    bodyRange.clearNote();
  }
}

function formatLogSheet_(sheet) {
  removeBlankLogRows_(sheet);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange("B:C").setNumberFormat("@");
  sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat("M/d/yyyy HH:mm:ss");
  insertLogWeeklyGapRows_(sheet);
  clearAutoClosedLogFormatting_(sheet);
  applyLogAutoClosedFormatting_(sheet);
}

function removeBlankLogRows_(sheet) {
  for (let row = sheet.getLastRow(); row >= 2; row -= 1) {
    const values = sheet.getRange(row, 1, 1, 14).getDisplayValues()[0];
    if (values.every((value) => !String(value || "").trim())) {
      sheet.deleteRow(row);
    }
  }
}

function insertLogWeeklyGapRows_(sheet) {
  for (let row = sheet.getLastRow() - 1; row >= 2; row -= 1) {
    const date = parseDisplayDate_(sheet.getRange(row, 2).getDisplayValue());
    const nextDate = parseDisplayDate_(sheet.getRange(row + 1, 2).getDisplayValue());
    if (date && nextDate && getMondayWeekStart_(nextDate) > getMondayWeekStart_(date)) {
      sheet.insertRowsAfter(row, 1);
    }
  }
}

function applyLogAutoClosedFormatting_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 14);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const openRows = {};

  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const timestamp = row[0];
    const workDate = getIsoDateFromDisplayDate_(row[1]);
    const action = String(row[3] || "").trim();
    const workerName = String(row[4] || "").trim();
    if (!workerName || !workDate) return;

    const key = `${normalizeName_(workerName)}|${workDate}`;
    if (action === "Check In" && !openRows[key]) {
      openRows[key] = { rowNumber, workDate, backgrounds: backgrounds[index] };
    }

    if (action === "Check Out" && openRows[key]) {
      delete openRows[key];
    }
  });

  Object.keys(openRows).forEach((key) => {
    const open = openRows[key];
    if (!shouldAutoCloseWorkDate_(open.workDate) || !isPlainLogRowBackground_(open.backgrounds)) return;

    sheet.getRange(open.rowNumber, 1, 1, 14).setBackground(AUTO_CLOSED_BACKGROUND);
    const actionCell = sheet.getRange(open.rowNumber, 4);
    if (!actionCell.getNote()) {
      actionCell.setNote(AUTO_CLOSED_NOTE);
    }
  });
}

function clearAutoClosedLogFormatting_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 14);
  const backgrounds = range.getBackgrounds();
  const notes = range.getNotes();

  backgrounds.forEach((row, rowIndex) => {
    row.forEach((background, columnIndex) => {
      if (String(background || "").trim().toLowerCase() === AUTO_CLOSED_BACKGROUND) {
        sheet.getRange(rowIndex + 2, columnIndex + 1).setBackground(null);
      }
    });

    if (notes[rowIndex][3] === AUTO_CLOSED_NOTE) {
      sheet.getRange(rowIndex + 2, 4).setNote("");
    }
  });
}

function isPlainLogRowBackground_(backgrounds) {
  return backgrounds.every((background) => {
    const color = String(background || "").trim().toLowerCase();
    return !color || color === "#ffffff" || color === AUTO_CLOSED_BACKGROUND;
  });
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
  sheet.getRange(2, 11, 1, 11).setValues([[
    "Date",
    "First Check In",
    "Last Check Out",
    "Work Hours",
    "Break",
    "Payable Hours",
    "Pay Rate",
    "Pay Amount",
    "Status",
    "Jobs",
    "Notes",
  ]]);
  sheet.getRange(1, 1, 2, 21).setFontWeight("bold");
}

function captureWorkerAnnotations_(sheet, workerName) {
  const annotations = { cells: {}, summaryValues: {} };
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return annotations;

  const range = sheet.getRange(3, 1, lastRow - 2, 22);
  const values = range.getDisplayValues();
  const notes = range.getNotes();
  const backgrounds = range.getBackgrounds();
  const headers = sheet.getRange(2, 1, 1, 22).getDisplayValues()[0];
  const payRateIndex = headers.indexOf("Pay Rate");
  const notesIndex = headers.indexOf("Notes", 10) !== -1 ? headers.indexOf("Notes", 10) : headers.indexOf("Pay Notes");

  values.forEach((row, rowIndex) => {
    const entryKey = getEntryAnnotationKey_(workerName, row);
    if (entryKey) captureRowAnnotations_(annotations.cells, entryKey, row, notes[rowIndex], backgrounds[rowIndex], 1, 9);

    const summaryKey = getSummaryAnnotationKey_(workerName, getIsoDateFromDisplayDate_(row[10]));
    if (summaryKey) {
      captureRowAnnotations_(annotations.cells, summaryKey, row, notes[rowIndex], backgrounds[rowIndex], 11, 21);
      annotations.summaryValues[summaryKey] = {
        payRate: payRateIndex === -1 ? "" : String(row[payRateIndex] || "").trim(),
        notes: notesIndex === -1 ? "" : String(row[notesIndex] || "").trim(),
      };
    }
  });

  return annotations;
}

function captureRowAnnotations_(annotations, key, values, notes, backgrounds, startColumn, endColumn) {
  for (let column = startColumn; column <= endColumn; column += 1) {
    const index = column - 1;
    const note = String(notes[index] || "").trim();
    const background = getPreservedBackground_(backgrounds[index]);
    if (!note && !background) continue;
    if (!annotations[key]) annotations[key] = {};
    annotations[key][column] = { note, background };
  }
}

function applyWorkerAnnotations_(sheet, workerName, annotations) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const values = sheet.getRange(3, 1, lastRow - 2, 21).getDisplayValues();
  values.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + 3;
    applyRowAnnotations_(sheet, rowNumber, annotations.cells[getEntryAnnotationKey_(workerName, row)]);

    const summaryKey = getSummaryAnnotationKey_(workerName, getIsoDateFromDisplayDate_(row[10]));
    applyRowAnnotations_(sheet, rowNumber, annotations.cells[summaryKey]);
  });
}

function applyRowAnnotations_(sheet, rowNumber, annotations) {
  if (!annotations) return;

  Object.keys(annotations).forEach((column) => {
    const annotation = annotations[column];
    const range = sheet.getRange(rowNumber, Number(column));
    if (annotation.note) range.setNote(annotation.note);
    if (annotation.background) range.setBackground(annotation.background);
  });
}

function getEntryAnnotationKey_(workerName, row) {
  const date = String(row[0] || "").trim();
  const time = String(row[1] || "").trim();
  const action = String(row[2] || "").trim();
  if (!date || !time || !action) return "";
  return `${normalizeName_(workerName)}|entry|${getIsoDateFromDisplayDate_(date)}|${time}|${action}`;
}

function getSummaryAnnotationKey_(workerName, workDate) {
  return workDate ? `${normalizeName_(workerName)}|summary|${workDate}` : "";
}

function getPreservedBackground_(background) {
  const color = String(background || "").trim().toLowerCase();
  if (!color || color === "#ffffff" || color === AUTO_CLOSED_BACKGROUND) return "";
  return background;
}

function applyAutoClosedFormatting_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const statusValues = sheet.getRange(3, 19, lastRow - 2, 1).getDisplayValues();
  statusValues.forEach((row, index) => {
    if (String(row[0] || "").trim() !== "Auto Closed") return;

    const rowNumber = index + 3;
    sheet.getRange(rowNumber, 11, 1, 11).setBackground(AUTO_CLOSED_BACKGROUND);
    sheet.getRange(rowNumber, 19).setNote(AUTO_CLOSED_NOTE);
  });
}

function calculateBreakHours_(workHours) {
  return workHours > 5 ? 0.5 : 0;
}

function parseLogDateTime_(dateText, timeText) {
  const date = parseDisplayDate_(dateText);
  if (!date || !timeText) return null;

  const parsed = new Date(`${dateText} ${timeText}`);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  const match = String(timeText).trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (!match) return null;

  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] || 0);
  const meridiem = String(match[4] || "").toUpperCase();
  if (meridiem === "PM" && hours < 12) hours += 12;
  if (meridiem === "AM" && hours === 12) hours = 0;

  date.setHours(hours, minutes, seconds, 0);
  return date;
}

function getWorkDateEnd_(dateKey) {
  const parts = String(dateKey || "").split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 23, 59, 0);
}

function shouldAutoCloseWorkDate_(dateKey) {
  return String(dateKey || "") < getIsoDateFromTimestamp_(new Date());
}

function getIsoDateFromDisplayDate_(dateText) {
  const date = parseDisplayDate_(dateText);
  return date ? Utilities.formatDate(date, SETTINGS.TIMEZONE, "yyyy-MM-dd") : "";
}

function deleteManualAdjustmentsSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("Manual Adjustments");
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
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
    if (workerName && isWorkerActive_(values[row][2])) names.push(workerName);
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

function addWeeklyGapRows_(rows, dateIndex, width) {
  const rowsWithGaps = [];

  rows.forEach((row, index) => {
    rowsWithGaps.push(row);

    const date = parseDisplayDate_(row[dateIndex]);
    const nextRow = rows[index + 1];
    const nextDate = nextRow ? parseDisplayDate_(nextRow[dateIndex]) : null;
    if (date && nextDate && getMondayWeekStart_(nextDate) > getMondayWeekStart_(date)) {
      rowsWithGaps.push(new Array(width).fill(""));
    }
  });

  return rowsWithGaps;
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
    const active = isWorkerActive_(values[row][2]);

    const nameMatches = !wantedName || normalizeName_(workerName) === wantedName;
    if (nameMatches && workerPin === wantedPin) {
      if (!active) {
        return { name: workerName, blocked: true };
      }
      return { name: workerName };
    }
  }

  return null;
}

function isWorkerActive_(value) {
  if (value === "" || value === null || typeof value === "undefined") return true;
  return String(value).trim().toUpperCase() !== "FALSE";
}

function getWorkerStatus_(spreadsheet, workerName) {
  const sheet = spreadsheet.getSheetByName(SETTINGS.LOG_SHEET);
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  const targetName = normalizeName_(workerName);

  for (let row = values.length - 1; row >= 1; row -= 1) {
    const name = String(values[row][4] || "").trim();
    if (normalizeName_(name) === targetName) {
      const lastAction = String(values[row][3] || "").trim();
      const lastTimestamp = parseLogDateTime_(displayValues[row][1], displayValues[row][2]);
      const lastWorkDate = getIsoDateFromDisplayDate_(displayValues[row][1]);
      if (lastAction === "Check In" && lastWorkDate && shouldAutoCloseWorkDate_(lastWorkDate)) {
        return {
          lastAction: "Check Out",
          lastTimestamp,
          autoClosed: true,
        };
      }

      return {
        lastAction,
        lastTimestamp,
      };
    }
  }

  return {
    lastAction: "",
    lastTimestamp: null,
  };
}

function getIsoDateFromTimestamp_(timestamp) {
  return Utilities.formatDate(timestamp, SETTINGS.TIMEZONE, "yyyy-MM-dd");
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
