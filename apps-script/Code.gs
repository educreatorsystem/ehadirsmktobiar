const CONFIG = {
  SPREADSHEET_ID: "1tEdH0WqNW7s3X0CRksPlBsB1M23AvITvQxtoWhzWaT4",
  ROSTER_SHEET_ID: 0,
  SUBJECT_SHEET_ID: 368172363,
  SHEET_HARIAN: "Rekod_Harian",
  SHEET_TAMBAHAN: "Rekod_Tambahan"
};

const ATTENDANCE_HEADERS = [
  "Timestamp",
  "Tarikh",
  "Jenis",
  "Kelas",
  "Subjek",
  "Jumlah Murid",
  "Bil Hadir",
  "Bil TH",
  "ID Murid TH JSON",
  "Nama Murid TH JSON",
  "ID Murid TA JSON",
  "Nama Murid TA JSON",
  "Rekod Key"
];

function doGet(e) {
  try {
    const action = String((e.parameter && e.parameter.action) || "fetchAllData");

    if (action === "ping") {
      return output_({
        status: "success",
        message: "Apps Script aktif",
        serverTime: new Date().toISOString()
      }, e);
    }

    if (action === "fetchAllData") {
      return output_({
        status: "success",
        data: getAllData_(),
        serverTime: new Date().toISOString()
      }, e);
    }

    if (action === "fetchRosters") {
      return output_({
        status: "success",
        data: getRosterData_(),
        serverTime: new Date().toISOString()
      }, e);
    }

    if (action === "submitAttendance") {
      const payloadText = String((e.parameter && e.parameter.payload) || "{}");
      const payload = JSON.parse(payloadText);
      const saved = saveAttendance_(payload);
      return output_({ status: "success", saved: saved }, e);
    }

    if (action === "setup") {
      setupSheets_();
      return output_({ status: "success", message: "Sheet rekod telah disediakan." }, e);
    }

    return output_({ status: "error", message: "Action tidak dikenali: " + action }, e);
  } catch (err) {
    return output_({ status: "error", message: err.message, stack: err.stack }, e);
  }
}

function doPost(e) {
  try {
    const payload = parsePostPayload_(e);
    const action = String(payload.action || "");

    if (action === "submitAttendance") {
      const saved = saveAttendance_(payload);
      return output_({ status: "success", saved: saved });
    }

    if (action === "syncAll") {
      syncAllAttendance_(payload);
      return output_({ status: "success", message: "Semua rekod telah disegerakkan." });
    }

    if (action === "ping") {
      return output_({ status: "success", message: "POST diterima." });
    }

    return output_({ status: "error", message: "Action POST tidak dikenali: " + action });
  } catch (err) {
    return output_({ status: "error", message: err.message, stack: err.stack });
  }
}

function getAllData_() {
  setupSheets_();

  return {
    recordsHarian: readAttendance_(CONFIG.SHEET_HARIAN),
    recordsTambahan: readAttendance_(CONFIG.SHEET_TAMBAHAN)
  };
}

function getRosterData_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const rosterSheet = ss.getSheetById(CONFIG.ROSTER_SHEET_ID);
  const subjectSheet = ss.getSheetById(CONFIG.SUBJECT_SHEET_ID);
  if (!rosterSheet) throw new Error("Tab roster murid tidak ditemui.");
  if (!subjectSheet) throw new Error("Tab subjek tidak ditemui.");

  const rosterValues = rosterSheet.getDataRange().getDisplayValues();
  const subjectValues = subjectSheet.getDataRange().getDisplayValues();
  return {
    rosterRows: rosterValues.slice(1).map(function(row) {
      return [cleanText_(row[0]), cleanText_(row[1])];
    }).filter(function(row) {
      return row[0] && row[1];
    }),
    subjectRows: subjectValues.slice(1).map(function(row) {
      return [cleanText_(row[0])];
    }).filter(function(row) {
      return row[0];
    })
  };
}

function saveAttendance_(payload) {
  const type = String(payload.type || "harian").toLowerCase() === "tambahan" ? "tambahan" : "harian";
  const sheetName = type === "tambahan" ? CONFIG.SHEET_TAMBAHAN : CONFIG.SHEET_HARIAN;
  const sheet = ensureAttendanceSheet_(sheetName);

  const date = cleanText_(payload.date);
  const className = cleanText_(payload.className);
  const subject = type === "tambahan" ? cleanText_(payload.subject) : "";
  if (!date || !className) throw new Error("Tarikh dan kelas diperlukan.");
  if (type === "tambahan" && !subject) throw new Error("Subjek diperlukan untuk kelas tambahan.");

  const absentStudentIds = normalizeIdArray_(payload.absentStudentIds);
  const absentNames = normalizeTextArray_(payload.absentNames);
  const excludedStudentIds = normalizeIdArray_(payload.excludedStudentIds);
  const excludedNames = normalizeTextArray_(payload.excludedNames);

  const totalStudents = Number(payload.totalStudents || 0);
  const absentCount = absentStudentIds.length;
  const presentCount = Math.max(totalStudents - absentCount - excludedStudentIds.length, 0);
  const key = makeRecordKey_(type, date, className, subject);

  const row = [
    new Date(),
    date,
    type,
    className,
    subject,
    totalStudents,
    presentCount,
    absentCount,
    JSON.stringify(absentStudentIds),
    JSON.stringify(absentNames),
    JSON.stringify(excludedStudentIds),
    JSON.stringify(excludedNames),
    key
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rowIndex = findRecordRow_(sheet, key);
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }

  return {
    key: key,
    type: type,
    date: date,
    className: className,
    subject: subject,
    totalStudents: totalStudents,
    presentCount: presentCount,
    absentCount: absentCount,
    absentStudentIds: absentStudentIds,
    absentNames: absentNames,
    excludedStudentIds: excludedStudentIds,
    excludedNames: excludedNames
  };
}

function syncAllAttendance_(payload) {
  replaceAttendanceSheet_(CONFIG.SHEET_HARIAN, "harian", payload.recordsHarian || []);
  replaceAttendanceSheet_(CONFIG.SHEET_TAMBAHAN, "tambahan", payload.recordsTambahan || []);
}

function replaceAttendanceSheet_(sheetName, type, records) {
  const sheet = ensureAttendanceSheet_(sheetName);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, ATTENDANCE_HEADERS.length).setValues([ATTENDANCE_HEADERS]);

    if (!records.length) return;

    const rows = records.map(function(record) {
      const className = cleanText_(record.className || record.kelas);
      const date = cleanText_(record.date || record.tarikh);
      const subject = type === "tambahan" ? cleanText_(record.subject || record.subjek) : "";
      const ids = normalizeIdArray_(record.absentStudentIds || record.idsTH);
      const names = normalizeTextArray_(record.absentNames);
      const excludedIds = normalizeIdArray_(record.excludedStudentIds || record.idsTA);
      const excludedNames = normalizeTextArray_(record.excludedNames);
      const total = Number(record.totalStudents || 0);

      return [
        new Date(),
        date,
        type,
        className,
        subject,
        total,
        Math.max(total - ids.length - excludedIds.length, 0),
        ids.length,
        JSON.stringify(ids),
        JSON.stringify(names),
        JSON.stringify(excludedIds),
        JSON.stringify(excludedNames),
        makeRecordKey_(type, date, className, subject)
      ];
    }).filter(function(row) {
      return row[1] && row[3];
    });

    if (rows.length) {
      sheet.getRange(2, 1, rows.length, ATTENDANCE_HEADERS.length).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
}

function readAttendance_(sheetName) {
  const sheet = ensureAttendanceSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1).map(function(row) {
    return {
      date: formatDateValue_(row[1]),
      className: cleanText_(row[3]),
      subject: cleanText_(row[4]),
      absentStudentIds: parseJsonArray_(row[8]).map(String),
      excludedStudentIds: parseJsonArray_(row[10]).map(String)
    };
  }).filter(function(record) {
    return record.date && record.className;
  });
}

function setupSheets_() {
  ensureAttendanceSheet_(CONFIG.SHEET_HARIAN);
  ensureAttendanceSheet_(CONFIG.SHEET_TAMBAHAN);
}

function ensureAttendanceSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, ATTENDANCE_HEADERS.length).setValues([ATTENDANCE_HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 10)).getValues()[0];
    if (currentHeaders[4] === "Jumlah Murid") {
      sheet.insertColumnAfter(4);
    }
    const headersAfterSubject = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 11)).getValues()[0];
    if (headersAfterSubject[10] === "Rekod Key") {
      sheet.insertColumnsBefore(11, 2);
    }
    const headers = sheet.getRange(1, 1, 1, ATTENDANCE_HEADERS.length).getValues()[0];
    if (headers.join("|") !== ATTENDANCE_HEADERS.join("|")) {
      sheet.getRange(1, 1, 1, ATTENDANCE_HEADERS.length).setValues([ATTENDANCE_HEADERS]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function findRecordRow_(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const keys = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === key) return i + 2;
  }
  return -1;
}

function parsePostPayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("POST body kosong.");
  }
  return JSON.parse(e.postData.contents);
}

function output_(payload, e) {
  const callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function makeRecordKey_(type, date, className, subject) {
  return [type, date, className, subject || ""].join("|");
}

function normalizeIdArray_(value) {
  if (Array.isArray(value)) return value.map(String).filter(Boolean);
  if (!value) return [];
  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value);
      if (Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
    } catch (err) {
      return value.split(/[,\n|]+/).map(cleanText_).filter(Boolean);
    }
  }
  return [];
}

function normalizeTextArray_(value) {
  if (Array.isArray(value)) return value.map(cleanText_).filter(Boolean);
  if (!value) return [];
  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value);
      if (Array.isArray(parsed)) return parsed.map(cleanText_).filter(Boolean);
    } catch (err) {
      return value.split(/[,\n|]+/).map(cleanText_).filter(Boolean);
    }
  }
  return [];
}

function parseJsonArray_(value) {
  if (Array.isArray(value)) return value;
  if (!value) return [];
  try {
    const parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return String(value).split(/[,\n|]+/).map(cleanText_).filter(Boolean);
  }
}

function cleanText_(value) {
  return String(value == null ? "" : value).trim();
}

function formatDateValue_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return cleanText_(value);
}
