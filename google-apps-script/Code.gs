const SETTINGS = {
  WORKERS_SHEET: "Workers",
  LOG_SHEET: "Time Log",
  TIMEZONE: "America/Toronto",
};

function setup() {
  setupSheets_(SpreadsheetApp.getActiveSpreadsheet(), true);
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

    if (action !== "CHECK_IN" && action !== "CHECK_OUT") {
      return fail_("Choose check in or check out.");
    }

    if (!name || !pin) {
      return fail_("Name and PIN are required.");
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets_(spreadsheet, false);

    const worker = findWorker_(spreadsheet, name, pin);
    if (!worker) {
      return fail_("Name or PIN is incorrect.");
    }

    const now = new Date();
    const displayDate = Utilities.formatDate(now, SETTINGS.TIMEZONE, "yyyy-MM-dd");
    const displayTime = Utilities.formatDate(now, SETTINGS.TIMEZONE, "h:mm a");
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
    rebuildWorkerSheet_(spreadsheet, worker.name);

    return {
      ok: true,
      name: worker.name,
      action,
      date: displayDate,
      time: displayTime,
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

  if (populateWorkerSheets) {
    const workerValues = workersSheet.getDataRange().getValues();
    for (let row = 1; row < workerValues.length; row += 1) {
      const workerName = String(workerValues[row][0] || "").trim();
      if (workerName) {
        rebuildWorkerSheet_(spreadsheet, workerName);
      }
    }
  }
}

function rebuildWorkerSheet_(spreadsheet, workerName) {
  const sheetName = getWorkerSheetName_(workerName);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const manualValues = readManualSummaryValues_(sheet);
  const logValues = spreadsheet.getSheetByName(SETTINGS.LOG_SHEET).getDataRange().getValues();
  const entries = [];
  const days = {};

  for (let row = 1; row < logValues.length; row += 1) {
    const timestamp = logValues[row][0];
    const date = String(logValues[row][1] || "").trim();
    const time = String(logValues[row][2] || "").trim();
    const action = String(logValues[row][3] || "").trim();
    const name = String(logValues[row][4] || "").trim();
    const job = String(logValues[row][5] || "").trim();
    const notes = String(logValues[row][6] || "").trim();
    const lat = String(logValues[row][7] || "").trim();
    const lng = String(logValues[row][8] || "").trim();
    const accuracy = String(logValues[row][9] || "").trim();
    const mapLink = String(logValues[row][10] || "").trim();

    if (name !== workerName) {
      continue;
    }

    entries.push([date, time, action, job, notes, lat, lng, accuracy, mapLink]);

    if (!date || !(timestamp instanceof Date)) {
      continue;
    }

    if (!days[date]) {
      days[date] = {
        events: [],
        jobs: [],
        notes: [],
      };
    }

    days[date].events.push({ timestamp, action });
    if (job) days[date].jobs.push(job);
    if (notes) days[date].notes.push(notes);
  }

  const summaryRows = Object.keys(days)
    .sort()
    .map((date) => {
      const day = days[date];
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

      const manual = manualValues[date] || {};
      const autoHours = totalMs ? Math.round((totalMs / 36e5) * 100) / 100 : "";
      const status = openCheckIn ? "Checked In" : "Complete";

      return [
        date,
        firstCheckIn ? Utilities.formatDate(firstCheckIn, SETTINGS.TIMEZONE, "h:mm a") : "",
        lastCheckOut ? Utilities.formatDate(lastCheckOut, SETTINGS.TIMEZONE, "h:mm a") : "",
        autoHours,
        manual.manualHours || "",
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
    const formulas = summaryRows.map((row, index) => {
      const rowNumber = index + 3;
      return [`=IF(P${rowNumber}<>"",IF(O${rowNumber}<>"",O${rowNumber},N${rowNumber})*P${rowNumber},"")`];
    });
    sheet.getRange(3, 17, formulas.length, 1).setFormulas(formulas);
  }

  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, 20);
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
  sheet.getRange(2, 11, 1, 10).setValues([[
    "Date",
    "First Check In",
    "Last Check Out",
    "Auto Hours",
    "Manual Hours",
    "Pay Rate",
    "Pay Amount",
    "Status",
    "Jobs",
    "Pay Notes",
  ]]);
  sheet.getRange(1, 1, 2, 20).setFontWeight("bold");
}

function readManualSummaryValues_(sheet) {
  const values = sheet.getDataRange().getValues();
  const manualValues = {};

  for (let row = 2; row < values.length; row += 1) {
    const date = String(values[row][10] || "").trim();
    if (!date) {
      continue;
    }

    manualValues[date] = {
      manualHours: values[row][14],
      payRate: values[row][15],
      payNotes: values[row][19],
    };
  }

  return manualValues;
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
  const wantedName = name.toLowerCase();
  const wantedPin = String(pin);

  for (let row = 1; row < values.length; row += 1) {
    const workerName = String(values[row][0] || "").trim();
    const workerPin = String(values[row][1] || "").trim();
    const active = String(values[row][2] || "TRUE").trim().toUpperCase();

    if (workerName.toLowerCase() === wantedName && workerPin === wantedPin && active !== "FALSE") {
      return { name: workerName };
    }
  }

  return null;
}

function sanitizeCallback_(callback) {
  return /^[a-zA-Z_$][\w$]*(\.[a-zA-Z_$][\w$]*)*$/.test(callback) ? callback : "callback";
}

function fail_(message) {
  return {
    ok: false,
    message,
  };
}
