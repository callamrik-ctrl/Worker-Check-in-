const SETTINGS = {
  WORKERS_SHEET: "Workers",
  LOG_SHEET: "Time Log",
  HOURS_SHEET: "Daily Hours",
  TIMEZONE: "America/Toronto",
};

function setup() {
  setupSheets_(SpreadsheetApp.getActiveSpreadsheet());
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
    setupSheets_(spreadsheet);

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
    rebuildDailyHours_(spreadsheet);

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

function setupSheets_(spreadsheet) {
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

  let hoursSheet = spreadsheet.getSheetByName(SETTINGS.HOURS_SHEET);
  if (!hoursSheet) {
    hoursSheet = spreadsheet.insertSheet(SETTINGS.HOURS_SHEET);
  }

  if (hoursSheet.getLastRow() === 0) {
    hoursSheet.appendRow([
      "Date",
      "Name",
      "First Check In",
      "Last Check Out",
      "Total Hours",
      "Status",
      "Jobs",
      "Notes",
    ]);
    hoursSheet.setFrozenRows(1);
  }
}

function rebuildDailyHours_(spreadsheet) {
  const logSheet = spreadsheet.getSheetByName(SETTINGS.LOG_SHEET);
  const hoursSheet = spreadsheet.getSheetByName(SETTINGS.HOURS_SHEET);
  const values = logSheet.getDataRange().getValues();
  const groups = {};

  for (let row = 1; row < values.length; row += 1) {
    const timestamp = values[row][0];
    const date = String(values[row][1] || "").trim();
    const action = String(values[row][3] || "").trim();
    const name = String(values[row][4] || "").trim();
    const job = String(values[row][5] || "").trim();
    const notes = String(values[row][6] || "").trim();

    if (!date || !name || !(timestamp instanceof Date)) {
      continue;
    }

    const key = `${date}|${name}`;
    if (!groups[key]) {
      groups[key] = {
        date,
        name,
        events: [],
        jobs: [],
        notes: [],
      };
    }

    groups[key].events.push({ timestamp, action });
    if (job) groups[key].jobs.push(job);
    if (notes) groups[key].notes.push(notes);
  }

  const rows = Object.keys(groups)
    .sort()
    .map((key) => {
      const group = groups[key];
      group.events.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());

      let openCheckIn = null;
      let firstCheckIn = null;
      let lastCheckOut = null;
      let totalMs = 0;

      group.events.forEach((event) => {
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

      const totalHours = totalMs ? Math.round((totalMs / 36e5) * 100) / 100 : "";
      const status = openCheckIn ? "Checked In" : "Complete";

      return [
        group.date,
        group.name,
        firstCheckIn ? Utilities.formatDate(firstCheckIn, SETTINGS.TIMEZONE, "h:mm a") : "",
        lastCheckOut ? Utilities.formatDate(lastCheckOut, SETTINGS.TIMEZONE, "h:mm a") : "",
        totalHours,
        status,
        uniqueText_(group.jobs),
        uniqueText_(group.notes),
      ];
    });

  hoursSheet.clearContents();
  hoursSheet.appendRow([
    "Date",
    "Name",
    "First Check In",
    "Last Check Out",
    "Total Hours",
    "Status",
    "Jobs",
    "Notes",
  ]);

  if (rows.length) {
    hoursSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  hoursSheet.setFrozenRows(1);
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
