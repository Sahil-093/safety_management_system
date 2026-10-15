/**
 * INDUCTION LOG
 * Reads and appends the SBARC induction register.
 */
var INDUCTION_SPREADSHEET_ID = "1BxOtq6QKLEwsSrhiv2cIRZdtNiS5k2pt2uZaFIep4pY";
var INDUCTION_SHEET_GID = 578338714;

function inductionTab_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === INDUCTION_SHEET_GID) return sheets[i];
  }
  return null;
}

function readInductionRows_() {
  try {
    const ss = SpreadsheetApp.openById(INDUCTION_SPREADSHEET_ID);
    const sheet = inductionTab_(ss);
    if (!sheet) throw new Error("Induction sheet tab was not found.");
    return sheet.getDataRange().getValues();
  } catch (err) {
    const url = "https://docs.google.com/spreadsheets/d/" + INDUCTION_SPREADSHEET_ID +
      "/export?format=csv&gid=" + INDUCTION_SHEET_GID;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() !== 200) throw err;
    return Utilities.parseCsv(response.getContentText());
  }
}

function inductionIdText_(value) {
  if (typeof value === "number" && isFinite(value)) return String(Math.round(value));
  return String(value || "").trim();
}

function inductionHeaderIndex_(rows) {
  const limit = Math.min(rows.length, 6);
  for (let i = 0; i < limit; i++) {
    if (String(rows[i][1] || "").trim().toLowerCase() === "name") return i;
  }
  return 1;
}

function inductionDateText_(row) {
  const day = inductionIdText_(row[5]);
  const month = String(row[6] || "").trim();
  const year = inductionIdText_(row[7]);
  return [day, month, year].filter(Boolean).join(" ");
}

function inductionIso_(row) {
  const day = parseInt(row[5], 10);
  const monthRaw = String(row[6] || "").trim().toLowerCase();
  const year = parseInt(row[7], 10);
  const full = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  let month = full.indexOf(monthRaw);
  if (month < 0 && monthRaw.length >= 3) {
    const short = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    month = short.indexOf(monthRaw.slice(0, 3));
  }
  if (!day || month < 0 || !year) return "";
  return year + "-" + ("0" + (month + 1)).slice(-2) + "-" + ("0" + day).slice(-2);
}

function inductionRecord_(row, sheetRow) {
  return {
    sheetRow: sheetRow,
    serial: inductionIdText_(row[0]),
    name: String(row[1] || "").trim(),
    sticker: String(row[2] || "").trim(),
    idNumber: inductionIdText_(row[3]),
    position: String(row[4] || "").trim(),
    day: inductionIdText_(row[5]),
    month: String(row[6] || "").trim(),
    year: inductionIdText_(row[7]),
    inductionDate: inductionDateText_(row),
    dateIso: inductionIso_(row),
    company: String(row[8] || "").trim(),
    trainer: String(row[9] || "").trim(),
    notes: String(row[10] || "").trim()
  };
}

function inductionDateParts_(iso) {
  const parts = String(iso || "").split("-");
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  if (parts.length !== 3 || parts[0].length !== 4) return null;
  const month = months[Number(parts[1]) - 1];
  if (!month) return null;
  return { day: Number(parts[2]), month: month, year: Number(parts[0]) };
}

function inductionWriteError_(err) {
  const message = String(err);
  if (/permission|access/i.test(message)) {
    return { success: false, message: "Cannot write the induction sheet. Share it with sahilkhattak093@gmail.com as Editor." };
  }
  return { success: false, message: message };
}

function listInductionEmployees() {
  const rows = readInductionRows_();
  const start = inductionHeaderIndex_(rows) + 1;
  const people = [];
  for (let i = start; i < rows.length; i++) {
    const name = String(rows[i][1] || "").trim();
    const id = inductionIdText_(rows[i][3]);
    if (!name && !id) continue;
    people.push(inductionRecord_(rows[i], i + 1));
  }
  people.reverse();
  return people;
}

function inductionQueryHit_(name, id, q, qDigits) {
  const nameHit = String(name || "").toLowerCase().indexOf(q) !== -1;
  const idDigits = String(id || "").replace(/\D/g, "");
  let idHit = false;
  if (id && qDigits.length >= 8) idHit = idDigits === qDigits;
  else if (id && qDigits.length >= 2) idHit = idDigits.indexOf(qDigits) !== -1;
  else if (id) idHit = String(id).toLowerCase().indexOf(q) !== -1;
  return nameHit || idHit;
}

function searchInductionLog(query) {
  const q = String(query || "").trim().toLowerCase();
  if (q.length < 2) return [];
  const rows = readInductionRows_();
  const start = inductionHeaderIndex_(rows) + 1;
  const qDigits = q.replace(/\D/g, "");
  const inductionIds = {};
  const byId = {};
  const records = [];
  for (let i = start; i < rows.length; i++) {
    const name = String(rows[i][1] || "").trim();
    const id = inductionIdText_(rows[i][3]);
    if (!name && !id) continue;
    const rec = inductionRecord_(rows[i], i + 1);
    rec.source = "induction";
    rec.inducted = true;
    const idDigits = id.replace(/\D/g, "");
    if (idDigits) {
      inductionIds[idDigits] = true;
      if (!byId[idDigits]) byId[idDigits] = [];
      byId[idDigits].push(rec);
    }
    if (inductionQueryHit_(name, id, q, qDigits)) records.push(rec);
  }
  records.reverse();

  const roster = readEmployeeRoster_();
  const seen = {};
  records.forEach(function (rec) { seen[rec.sheetRow] = true; });
  const extras = [];
  roster.forEach(function (person) {
    const idDigits = String(person.idNumber || "").replace(/\D/g, "");
    const hit = inductionQueryHit_(person.name, person.idNumber, q, qDigits);
    if (idDigits && inductionIds[idDigits]) {
      (byId[idDigits] || []).forEach(function (rec) {
        if (person.photoUrl) rec.photoUrl = person.photoUrl;
        if (person.contact) rec.contact = person.contact;
        if (person.email) rec.email = person.email;
        if (person.submittedBy) rec.submittedBy = person.submittedBy;
        if (hit && !seen[rec.sheetRow]) {
          records.push(rec);
          seen[rec.sheetRow] = true;
        }
      });
      return;
    }
    if (!hit) return;
    extras.push(person);
  });
  extras.reverse();
  return attachInductionPhotos_(publishRosterPhotos_(extras.concat(records).slice(0, 40)));
}

function publishRosterPhotos_(records) {
  const cache = CacheService.getScriptCache();
  records.forEach(function (rec) {
    const match = String(rec.photoUrl || "").match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) return;
    const key = "photo-public-" + match[1];
    if (cache.get(key)) return;
    try {
      DriveApp.getFileById(match[1]).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      cache.put(key, "1", 21600);
    } catch (err) {
      try { cache.put(key, "0", 300); } catch (ignore) {}
    }
  });
  return records;
}

var INDUCTION_PHOTO_FOLDER = "SBARC Induction Photos";

function inductionPhotoFolder_(create) {
  const folders = DriveApp.getFoldersByName(INDUCTION_PHOTO_FOLDER);
  if (folders.hasNext()) return folders.next();
  if (!create) return null;
  return DriveApp.createFolder(INDUCTION_PHOTO_FOLDER);
}

function inductionPhotoFolderLink_() {
  try {
    const folder = inductionPhotoFolder_(true);
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {}
    let email = "";
    try { email = Session.getEffectiveUser().getEmail(); } catch (err) {}
    return { success: true, url: folder.getUrl(), name: folder.getName(), email: email };
  } catch (err) {
    const message = String(err);
    if (/permission|authorization|drive/i.test(message)) {
      return { success: false, message: "Google Drive access is needed. Open the Apps Script project once and allow Drive." };
    }
    return { success: false, message: message };
  }
}

function inductionPhotoFileName_(serial, idNumber) {
  const serialText = String(serial || "").trim();
  if (serialText) return "sn-" + serialText + ".jpg";
  const id = String(idNumber || "").replace(/[^\w.-]/g, "");
  return "id-" + (id || "unknown") + ".jpg";
}

function inductionPhotoUrl_(fileId) {
  return "https://lh3.googleusercontent.com/d/" + fileId;
}

function inductionPhotoMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("induction-photos");
  if (cached) {
    try { return JSON.parse(cached); } catch (err) {}
  }
  const folder = inductionPhotoFolder_(false);
  const map = {};
  if (folder) {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      map[file.getName()] = file.getId();
    }
  }
  try { cache.put("induction-photos", JSON.stringify(map), 300); } catch (err) {}
  return map;
}

function attachInductionPhotos_(records) {
  if (!records.length) return records;
  let map = {};
  try { map = inductionPhotoMap_(); } catch (err) { map = {}; }
  for (let i = 0; i < records.length; i++) {
    if (records[i].photoUrl) continue;
    const name = inductionPhotoFileName_(records[i].serial, records[i].idNumber);
    records[i].photoUrl = map[name] ? inductionPhotoUrl_(map[name]) : "";
  }
  return records;
}

function uploadInductionPhoto(formData) {
  try {
    const row = parseInt(formData.sheetRow, 10);
    if (!row || row < 3) return { success: false, message: "Invalid induction row." };
    const image = String(formData.image || "").replace(/^data:image\/\w+;base64,/, "");
    if (!image) return { success: false, message: "Choose a picture first." };
    if (image.length > 120000) return { success: false, message: "Picture is too large. Choose a smaller photo." };

    const ss = SpreadsheetApp.openById(INDUCTION_SPREADSHEET_ID);
    const sheet = inductionTab_(ss);
    if (!sheet) return { success: false, message: "Induction sheet tab was not found." };

    const currentName = String(sheet.getRange(row, 2).getValue() || "").trim();
    const currentId = inductionIdText_(sheet.getRange(row, 4).getValue());
    const currentSerial = inductionIdText_(sheet.getRange(row, 1).getValue());
    const sameRow = currentName === String(formData.originalName || "").trim() &&
      currentId === inductionIdText_(formData.originalId) &&
      currentSerial === String(formData.originalSerial || "").trim();
    if (!sameRow) {
      return { success: false, message: "This row changed in the sheet. Search again and try again." };
    }

    const folder = inductionPhotoFolder_(true);
    const name = inductionPhotoFileName_(currentSerial, currentId);
    const existing = folder.getFilesByName(name);
    while (existing.hasNext()) existing.next().setTrashed(true);
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(image), "image/jpeg", name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    try { CacheService.getScriptCache().remove("induction-photos"); } catch (err) {}
    return { success: true, message: "Picture saved.", photoUrl: inductionPhotoUrl_(file.getId()) };
  } catch (err) {
    const message = String(err);
    if (/permission|authorization|drive/i.test(message)) {
      return { success: false, message: "Google Drive access is needed. Open the Apps Script project once and allow Drive, then try again." };
    }
    return inductionWriteError_(err);
  }
}

function recordInduction(formData, user) {
  try {
    const name = String(formData.name || "").trim();
    const idNumber = String(formData.idNumber || "").trim();
    if (!name) return { success: false, message: "Employee name is required." };

    const ss = SpreadsheetApp.openById(INDUCTION_SPREADSHEET_ID);
    const sheet = inductionTab_(ss);
    if (!sheet) return { success: false, message: "Induction sheet tab was not found." };

    const date = inductionDateParts_(formData.inductionDate) || { day: "", month: "", year: "" };

    const rows = sheet.getDataRange().getValues();
    let serial = 0;
    for (let i = 0; i < rows.length; i++) {
      const n = parseInt(rows[i][0], 10);
      if (!isNaN(n) && n > serial) serial = n;
    }

    sheet.appendRow([
      serial + 1,
      name,
      formData.sticker || "",
      idNumber,
      formData.position || "",
      date.day,
      date.month,
      date.year,
      formData.company || "",
      formData.trainer || (user && user.name) || "",
      formData.notes || ""
    ]);
    const last = sheet.getLastRow();
    if (idNumber) sheet.getRange(last, 4).setNumberFormat("@").setValue(idNumber);
    return { success: true, message: "Induction saved." };
  } catch (err) {
    return inductionWriteError_(err);
  }
}

function updateInductionRecord(formData) {
  try {
    const row = parseInt(formData.sheetRow, 10);
    const name = String(formData.name || "").trim();
    if (!row || row < 3) return { success: false, message: "Invalid induction row." };
    if (!name) return { success: false, message: "Employee name is required." };

    const ss = SpreadsheetApp.openById(INDUCTION_SPREADSHEET_ID);
    const sheet = inductionTab_(ss);
    if (!sheet) return { success: false, message: "Induction sheet tab was not found." };

    const currentName = String(sheet.getRange(row, 2).getValue() || "").trim();
    const currentId = inductionIdText_(sheet.getRange(row, 4).getValue());
    const sameRow = currentName === String(formData.originalName || "").trim() &&
      currentId === inductionIdText_(formData.originalId);
    if (!sameRow) {
      return { success: false, message: "This row changed in the sheet. Search again and try again." };
    }

    const idNumber = String(formData.idNumber || "").trim();
    sheet.getRange(row, 2).setValue(name);
    sheet.getRange(row, 3).setValue(formData.sticker || "");
    sheet.getRange(row, 4).setNumberFormat("@").setValue(idNumber);
    sheet.getRange(row, 5).setValue(formData.position || "");
    const date = inductionDateParts_(formData.inductionDate);
    if (date) {
      sheet.getRange(row, 6).setValue(date.day);
      sheet.getRange(row, 7).setValue(date.month);
      sheet.getRange(row, 8).setValue(date.year);
    }
    sheet.getRange(row, 9).setValue(formData.company || "");
    sheet.getRange(row, 10).setValue(formData.trainer || "");
    sheet.getRange(row, 11).setValue(formData.notes || "");
    return { success: true, message: "Induction updated." };
  } catch (err) {
    return inductionWriteError_(err);
  }
}

function deleteInductionRecord(formData) {
  try {
    const row = parseInt(formData.sheetRow, 10);
    if (!row || row < 3) return { success: false, message: "Invalid induction row." };

    const ss = SpreadsheetApp.openById(INDUCTION_SPREADSHEET_ID);
    const sheet = inductionTab_(ss);
    if (!sheet) return { success: false, message: "Induction sheet tab was not found." };

    const currentName = String(sheet.getRange(row, 2).getValue() || "").trim();
    const currentId = inductionIdText_(sheet.getRange(row, 4).getValue());
    const currentSerial = String(sheet.getRange(row, 1).getValue() || "").trim();
    const sameRow = currentName === String(formData.originalName || "").trim() &&
      currentId === inductionIdText_(formData.originalId) &&
      currentSerial === String(formData.originalSerial || "").trim();
    if (!sameRow) {
      return { success: false, message: "This row changed in the sheet. Search again and try again." };
    }

    sheet.deleteRow(row);
    return { success: true, message: "Induction deleted." };
  } catch (err) {
    return inductionWriteError_(err);
  }
}

var EMPLOYEE_SPREADSHEET_ID = "1KhkV4-yBQi_PgYedq0VZaTPhviB2_zzxRrdKjJk-VIg";
var EMPLOYEE_SHEET_GID = 1794678974;

function employeeRosterTab_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === EMPLOYEE_SHEET_GID) return sheets[i];
  }
  return null;
}

function readEmployeeRosterRows_() {
  try {
    const ss = SpreadsheetApp.openById(EMPLOYEE_SPREADSHEET_ID);
    const sheet = employeeRosterTab_(ss);
    if (!sheet) throw new Error("Employee sheet tab was not found.");
    return sheet.getDataRange().getDisplayValues();
  } catch (err) {
    const url = "https://docs.google.com/spreadsheets/d/" + EMPLOYEE_SPREADSHEET_ID +
      "/export?format=csv&gid=" + EMPLOYEE_SHEET_GID;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() !== 200) throw err;
    return Utilities.parseCsv(response.getContentText());
  }
}

function employeeHeaderMap_(header) {
  const map = {};
  for (let i = 0; i < header.length; i++) {
    map[String(header[i] || "").trim().toLowerCase()] = i;
  }
  return map;
}

function employeeCell_(row, map, name) {
  const index = map[name];
  if (index === undefined) return "";
  return row[index];
}

function employeePhotoUrl_(value) {
  const text = String(value || "").trim();
  const open = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  const file = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const id = open ? open[1] : (file ? file[1] : "");
  if (id) return "https://lh3.googleusercontent.com/d/" + id + "=w400";
  if (/^https?:\/\//i.test(text)) return text;
  return "";
}

function readEmployeeRoster_() {
  let rows = [];
  try { rows = readEmployeeRosterRows_(); } catch (err) { return []; }
  if (!rows.length) return [];
  let headerAt = 0;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    if (rows[i].join(" ").toLowerCase().indexOf("employee name") !== -1) {
      headerAt = i;
      break;
    }
  }
  const map = employeeHeaderMap_(rows[headerAt]);
  const people = [];
  for (let i = headerAt + 1; i < rows.length; i++) {
    const name = String(employeeCell_(rows[i], map, "employee name") || "").trim();
    const idNumber = inductionIdText_(employeeCell_(rows[i], map, "employee id/iqama"));
    if (!name && !idNumber) continue;
    people.push({
      sheetRow: 0,
      source: "roster",
      inducted: false,
      serial: "",
      name: name,
      sticker: "",
      idNumber: idNumber,
      position: String(employeeCell_(rows[i], map, "employee profession") || "").trim(),
      day: "",
      month: "",
      year: "",
      inductionDate: "",
      dateIso: "",
      company: String(employeeCell_(rows[i], map, "company name") || "").trim(),
      trainer: "",
      notes: "",
      contact: String(employeeCell_(rows[i], map, "employee contact") || "").trim(),
      email: String(employeeCell_(rows[i], map, "email address") || "").trim(),
      submittedBy: String(employeeCell_(rows[i], map, "submitted by") || "").trim(),
      submittedAt: String(employeeCell_(rows[i], map, "timestamp") || "").trim(),
      photoUrl: employeePhotoUrl_(employeeCell_(rows[i], map, "employee photo"))
    });
  }
  return people;
}
