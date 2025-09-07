function rebuildExpandedEventsAndExportJSON() {
  // === CONFIG ===
  var SOURCE_SHEET_NAME = "Form Responses 1";
  var TARGET_SPREADSHEET_ID = "1Xi7SOBiQ_VjeOVW12pkSVldVBHb7OwBcRjGXD2OMPvQ";
  var TARGET_SHEET_NAME = "Events Expanded";

  // If a row has no explicit recurrence end date, how many months ahead should we expand?
  var DEFAULT_MONTHS_AHEAD = 12;

  // Optional: cap rows written to the Google Sheet (does NOT affect JSON)
  var SHEET_ROWS_CAP = 1000;

  // Timezone: use the script's timezone (set in File → Project properties),
  // or hardcode (e.g. "America/Chicago")
  var TIMEZONE = Session.getScriptTimeZone();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  var targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var targetSheet = targetSs.getSheetByName(TARGET_SHEET_NAME);

  // Prepare target sheet
  targetSheet.clearContents();
  targetSheet.appendRow([
    "Show name",
    "Date & Time",
    "Date",
    "Stage",
    "Show Description",
    "Show image",
    "Genre",
    "Email",
    "Instagram",
    "Website",
    "Duration"
  ]);

  var data = sourceSheet.getRange(2, 1, Math.max(0, sourceSheet.getLastRow() - 1), sourceSheet.getLastColumn()).getValues();
  var now = new Date();
  now.setHours(0,0,0,0);

  var rowsToWrite = [];
  var jsonEvents = [];

  data.forEach(function (row) {
    var timestamp    = row[0]; // form timestamp
    var showName     = safeStr(row[1]);
    var startDateRaw = row[2]; // Date (Google Sheets date)
    var stage        = safeStr(row[3]);
    var description  = safeStr(row[4]);
    var image        = safeStr(row[5]);
    var endDateRaw   = row[6]; // Date (Google Sheets date) - optional
    var timeVal      = row[7]; // could be "8:00 PM", "20:00", a Date, or a number (fraction of day)
    var recurringNum = safeStr(row[8]); // e.g. "1,3" or "1st and 3rd"
    var recurringDay = safeStr(row[9]); // e.g. "Wednesday"
    var genre        = safeStr(row[10]); // comma-separated tags
    var email        = safeStr(row[11]);
    var instagram    = safeStr(row[12]);
    var website      = safeStr(row[13]);
    var duration     = safeStr(row[14]);

    if (!startDateRaw) return; // must have a start date

    var dateTime = combineDateTime(startDateRaw, timeVal);
    var dateOnly = stripTime(dateTime);

    // Build link cells for the sheet (leave blank if not provided)
    var emailCell = email ? ('<a href="mailto:' + email + '">Email Link</a>') : "";
    var instaCell = instagram ? ('<a href="https://instagram.com/' + instagram.replace(/^@/, "") + '">IG Link</a>') : "";
    var webCell   = website ? ('<a href="' + normalizeUrl(website) + '">Website</a>') : "";

    // Recurrence handling
    var nonRecurring = !recurringNum || !recurringDay ||
                       equalsIgnoreCase(recurringNum, "Does not recur") ||
                       equalsIgnoreCase(recurringDay, "Does not recur");

    if (nonRecurring) {
      if (dateTime >= now) {
        rowsToWrite.push([showName, dateTime, dateOnly, stage, description, image, genre, emailCell, instaCell, webCell, duration]);
        jsonEvents.push(createEventObject(showName, dateTime, stage, description, image, genre, email, instagram, website, duration, TIMEZONE));
      }
      return;
    }

    // Monthly recurrence: expand until the provided end date, else DEFAULT_MONTHS_AHEAD
    var ordinals = parseOrdinals(recurringNum);       // e.g. [1,3]
    var targetDow = parseWeekday(recurringDay);       // 0..6 (Sun..Sat)
    if (ordinals.length === 0 || targetDow == null) return;

    var startDate = new Date(startDateRaw);
    startDate.setHours(0,0,0,0);

    var endLimit = endDateRaw ? new Date(endDateRaw) : addMonths(startDate, DEFAULT_MONTHS_AHEAD);
    endLimit.setHours(23,59,59,999);

    var cur = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    while (cur <= endLimit) {
      var y = cur.getFullYear();
      var m = cur.getMonth();

      ordinals.forEach(function (ord) {
        var showDate = getNthWeekdayOfMonth(y, m, targetDow, ord);
        if (showDate) {
          var dt = combineDateTime(showDate, timeVal); // same time each month
          if (dt >= now && dt <= endLimit && dt >= dateTime) { // don’t generate before initial start
            rowsToWrite.push([showName, dt, stripTime(dt), stage, description, image, genre, emailCell, instaCell, webCell, duration]);
            jsonEvents.push(createEventObject(showName, dt, stage, description, image, genre, email, instagram, website, duration, TIMEZONE));
          }
        }
      });

      cur = addMonths(cur, 1);
    }
  });

  // Sort by date (earliest first)
  rowsToWrite.sort(function (a, b) { return new Date(a[1]) - new Date(b[1]); });
  jsonEvents.sort(function (a, b) { return new Date(a.date) - new Date(b.date); });

  // Optional: cap rows written to SHEET (does NOT affect JSON)
  if (rowsToWrite.length > SHEET_ROWS_CAP) {
    rowsToWrite = rowsToWrite.slice(0, SHEET_ROWS_CAP);
  }

  // Write to spreadsheet
  if (rowsToWrite.length > 0) {
    targetSheet.getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
    targetSheet.getRange(2, 2, rowsToWrite.length, 1).setNumberFormat("MM/dd/yyyy hh:mm AM/PM");
    targetSheet.getRange(2, 3, rowsToWrite.length, 1).setNumberFormat("MM/dd/yyyy");
  }

  // === EXPORT TO JSON ===  (local ISO strings with timezone offset)
  var jsonOutput = {
    lastUpdated: Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    tz: TIMEZONE,
    totalEvents: jsonEvents.length,
    events: jsonEvents
  };

  var jsonString = JSON.stringify(jsonOutput, null, 2);

  // Save to Google Drive (publicly readable)
  saveJSONToGoogleDrive(jsonString, 'clubhouse-events.json');

  Logger.log('Export completed. ' + jsonEvents.length + ' events exported.');
}

/* ================= Helpers ================= */

function safeStr(v) {
  return (v == null) ? "" : String(v).trim();
}

function equalsIgnoreCase(a, b) {
  return String(a).toLowerCase() === String(b).toLowerCase();
}

function normalizeUrl(url) {
  if (!url) return "";
  if (!/^https?:\/\//i.test(url)) return "https://" + url;
  return url;
}

function stripTime(dt) {
  if (!dt) return null;
  var d = new Date(dt);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Robust combo: supports strings like "8:00 PM", "20:00", a Date, or a number (fraction of day).
 */
function combineDateTime(dateVal, timeVal) {
  if (!dateVal) return null;
  var d = new Date(dateVal);
  if (!timeVal && timeVal !== 0) {
    // If time missing, leave whatever hours are in dateVal (usually 00:00)
    return d;
  }

  if (Object.prototype.toString.call(timeVal) === "[object Date]") {
    d.setHours(timeVal.getHours());
    d.setMinutes(timeVal.getMinutes());
    return d;
  }

  if (typeof timeVal === "number") {
    var msInDay = 24 * 60 * 60 * 1000;
    var ms = Math.round(timeVal * msInDay);
    d.setHours(0, 0, 0, 0);
    return new Date(d.getTime() + ms);
  }

  // String case
  var s = String(timeVal).trim();
  // Try hh:mm AM/PM
  var m12 = s.match(/^(\d{1,2}):(\d{2})\s*([AaPp][Mm])$/);
  if (m12) {
    var hh = parseInt(m12[1], 10) % 12;
    var mm = parseInt(m12[2], 10);
    if (/[Pp]/.test(m12[3])) hh += 12;
    d.setHours(hh, mm, 0, 0);
    return d;
  }
  // Try 24h "HH:mm"
  var m24 = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m24) {
    d.setHours(parseInt(m24[1], 10), parseInt(m24[2], 10), 0, 0);
    return d;
  }

  // Fallback: leave date hours as-is
  return d;
}

/**
 * Parse ordinals like "1,3", "1st and 3rd", "First, Third"
 * Returns array of integers [1..5]. Supports "last" => -1 (not used by UI, but harmless).
 */
function parseOrdinals(str) {
  if (!str) return [];
  var s = String(str).toLowerCase();

  // replace words with numbers
  var map = {
    "first": 1, "1st": 1,
    "second": 2, "2nd": 2,
    "third": 3, "3rd": 3,
    "fourth": 4, "4th": 4,
    "fifth": 5, "5th": 5,
    "last": -1
  };
  Object.keys(map).forEach(function (k) {
    s = s.replace(new RegExp("\\b" + k + "\\b", "g"), String(map[k]));
  });

  // split on commas/space/and
  var parts = s.split(/[^-?\d]+/).filter(Boolean).map(function (x) { return parseInt(x, 10); });
  // keep only allowed values
  parts = parts.filter(function (n) { return n >= -1 && n <= 5 && n !== 0 && !isNaN(n); });

  // uniq & sort (keep original order is fine)
  var seen = {};
  var out = [];
  parts.forEach(function (n) { if (!seen[n]) { seen[n] = 1; out.push(n); } });
  return out;
}

/**
 * Sunday=0 .. Saturday=6
 */
function parseWeekday(s) {
  if (!s) return null;
  var name = String(s).trim().toLowerCase();
  var map = {
    "sun": 0, "sunday": 0,
    "mon": 1, "monday": 1,
    "tue": 2, "tuesday": 2,
    "wed": 3, "wednesday": 3,
    "thu": 4, "thursday": 4,
    "fri": 5, "friday": 5,
    "sat": 6, "saturday": 6
  };
  return (name in map) ? map[name] : null;
}

/**
 * nth weekday of a month (1..5). If n = -1, returns the last weekday of month.
 */
function getNthWeekdayOfMonth(year, month, weekday, n) {
  if (n === -1) {
    // last weekday of the month
    var d = new Date(year, month + 1, 0); // last day of month
    while (d.getDay() !== weekday) d.setDate(d.getDate() - 1);
    return new Date(d);
  }
  if (n < 1 || n > 5) return null;

  var d0 = new Date(year, month, 1);
  var firstDow = d0.getDay(); // 0..6
  var delta = (weekday - firstDow + 7) % 7;
  var day = 1 + delta + (n - 1) * 7;
  var d = new Date(year, month, day);
  if (d.getMonth() !== month) return null; // overflow (e.g. 5th weekday that doesn't exist)
  return d;
}

function addMonths(d, n) {
  return new Date(d.getFullYear(), d.getMonth() + n, d.getDate());
}

/**
 * Create event object with LOCAL ISO string (includes timezone offset).
 */
function createEventObject(title, dateTime, stage, description, image, genre, email, instagram, website, duration, tz) {
  var localISO = Utilities.formatDate(dateTime, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
  return {
    id: Utilities.getUuid(),
    title: title,
    date: localISO,
    stage: stage || "",
    genre: genre || "",
    description: description || "",
    image: image || "",
    email: email || "",
    instagram: instagram || "",
    website: website || "",
    duration: duration || ""
  };
}

/* ===== Outputs ===== */

function saveJSONToGoogleDrive(jsonString, filename) {
  try {
    var files = DriveApp.getFilesByName(filename);
    while (files.hasNext()) files.next().setTrashed(true);

    var blob = Utilities.newBlob(jsonString, 'application/json', filename);
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('JSON file saved to Google Drive: ' + file.getUrl());
    return file.getUrl();
  } catch (err) {
    Logger.log('Error saving JSON: ' + err);
  }
}
