/**
 * HSE Observations from the form responses sheet.
 * Closeout columns are added only when someone closes an item:
 * Closeout Response, Closed By, Closeout Photos.
 */
var OBSERVATION_SPREADSHEET_ID = "1pNhVJAlOjTbfkTJetF_d05n2QOlayV8brC-x6XlN2QY";
var OBSERVATION_SHEET_GID = 1533864783;
var OBSERVATION_PHOTO_FOLDER = "SBARC Observation Closeouts";

function observationSheet_() {
  const ss = SpreadsheetApp.openById(OBSERVATION_SPREADSHEET_ID);
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === OBSERVATION_SHEET_GID) return sheets[i];
  }
  return ss.getSheetByName("Form Responses 1");
}

function observationHeader_(sheet) {
  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  const map = {};
  headers.forEach(function (name, index) {
    const key = String(name || "").trim().toLowerCase();
    if (key) map[key] = index;
  });
  return { headers: headers, map: map };
}

function observationCol_(map, name) {
  const key = String(name || "").trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(map, key) ? map[key] : -1;
}

function observationText_(value) {
  return String(value == null ? "" : value).trim();
}

function observationDate_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, "Asia/Riyadh", "yyyy-MM-dd");
  }
  const text = observationText_(value);
  if (!text) return "";
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, "Asia/Riyadh", "yyyy-MM-dd");
  return "";
}

function observationWhen_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, "Asia/Riyadh", "d MMM yyyy, h:mm a");
  }
  return observationText_(value);
}

function observationCell_(row, index) {
  if (index < 0) return "";
  return row[index];
}

function observationFileId_(url) {
  const text = String(url || "");
  let match = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  match = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return "";
}

function observationPhotoList_(cell) {
  const photos = [];
  const seen = {};
  String(cell || "").split(/\s+/).forEach(function (chunk) {
    const id = observationFileId_(chunk);
    if (!id || seen[id]) return;
    seen[id] = true;
    photos.push({
      id: id,
      openUrl: "https://drive.google.com/file/d/" + id + "/view"
    });
  });
  return photos;
}

function observationPublish_(id) {
  if (!id) return "";
  const cache = CacheService.getScriptCache();
  const key = "obs-photo-" + id;
  if (cache.get(key) === "1") return "https://lh3.googleusercontent.com/d/" + id + "=w800";
  try {
    DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    cache.put(key, "1", 21600);
    return "https://lh3.googleusercontent.com/d/" + id + "=w800";
  } catch (err) {
    try { cache.put(key, "0", 300); } catch (ignore) {}
    return "";
  }
}

function observationPhotos_(cell) {
  return observationPhotoList_(cell).map(function (photo) {
    return {
      imageUrl: observationPublish_(photo.id),
      openUrl: photo.openUrl
    };
  });
}

function listObservations(token) {
  const actor = chatActor_(token);
  if (!actor) return { success: false, message: "Sign in again to open observations." };
  try {
    const sheet = observationSheet_();
    if (!sheet) return { success: false, message: "Observations sheet was not found." };
    const header = observationHeader_(sheet);
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, records: [] };
    const values = sheet.getRange(2, 1, last - 1, Math.max(sheet.getLastColumn(), 1)).getValues();
    const col = function (name) { return observationCol_(header.map, name); };
    const records = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const ref = observationText_(observationCell_(row, col("ref. no.")));
      const title = observationText_(observationCell_(row, col("observation title")));
      if (!ref && !title) continue;
      records.push({
        sheetRow: i + 2,
        ref: ref,
        status: observationText_(observationCell_(row, col("status"))) || "Open",
        date: observationDate_(observationCell_(row, col("timestamp"))),
        when: observationWhen_(observationCell_(row, col("timestamp"))),
        contractor: observationText_(observationCell_(row, col("contractor/subcontractor"))),
        zone: observationText_(observationCell_(row, col("zone"))),
        building: observationText_(observationCell_(row, col("building/cluster"))),
        title: title,
        villa: observationText_(observationCell_(row, col("villa/part"))),
        description: observationText_(observationCell_(row, col("observation description"))),
        responsible: observationText_(observationCell_(row, col("responsible person name"))),
        email: observationText_(observationCell_(row, col("email address"))),
        submittedBy: observationText_(observationCell_(row, col("submitted by"))),
        photos: observationPhotos_(observationCell_(row, col("photos"))),
        closeout: observationText_(observationCell_(row, col("closeout response"))),
        closedBy: observationText_(observationCell_(row, col("closed by"))),
        closePhotos: observationPhotos_(observationCell_(row, col("closeout photos")))
      });
    }
    return { success: true, records: records };
  } catch (err) {
    const message = String(err);
    if (/permission|access|authorization/i.test(message)) {
      return { success: false, message: "Share the observations sheet with the portal Google account as Editor." };
    }
    return { success: false, message: message };
  }
}

function observationEnsureColumn_(sheet, header, name) {
  const existing = observationCol_(header.map, name);
  if (existing >= 0) return existing;
  const col = Math.max(sheet.getLastColumn(), header.headers.length) + 1;
  sheet.getRange(1, col).setValue(name).setFontWeight("bold");
  header.map[name.toLowerCase()] = col - 1;
  header.headers.push(name);
  return col - 1;
}

function observationCloseFolder_() {
  const folders = DriveApp.getFoldersByName(OBSERVATION_PHOTO_FOLDER);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(OBSERVATION_PHOTO_FOLDER);
}

function observationSavePhoto_(ref, index, image) {
  const raw = String(image || "").replace(/^data:image\/\w+;base64,/, "");
  if (!raw) return { url: "" };
  if (raw.length > 120000) return { success: false, message: "A picture is too large. Choose a smaller photo." };
  const folder = observationCloseFolder_();
  const safeRef = String(ref || "observation").replace(/[^\w.-]/g, "").slice(0, 40) || "observation";
  const name = safeRef + "-" + (index + 1) + "-" + new Date().getTime() + ".jpg";
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(raw), "image/jpeg", name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: "https://lh3.googleusercontent.com/d/" + file.getId() + "=w800" };
}

function closeObservation(token, form) {
  const actor = chatActor_(token);
  if (!actor) return { success: false, message: "Sign in again to close an observation." };
  const incoming = form || {};
  const row = parseInt(incoming.sheetRow, 10);
  const response = observationText_(incoming.response);
  if (!row || row < 2) return { success: false, message: "That observation was not found." };
  if (!response) return { success: false, message: "Write the closeout response first." };
  if (response.length > 2000) return { success: false, message: "Closeout response is too long." };
  try {
    const sheet = observationSheet_();
    if (!sheet) return { success: false, message: "Observations sheet was not found." };
    const header = observationHeader_(sheet);
    const refCol = observationCol_(header.map, "ref. no.");
    const statusCol = observationCol_(header.map, "status");
    if (refCol < 0 || statusCol < 0) return { success: false, message: "Observation columns were not found." };
    const currentRef = observationText_(sheet.getRange(row, refCol + 1).getValue());
    if (!currentRef || currentRef !== observationText_(incoming.ref)) {
      return { success: false, message: "This row changed in the sheet. Refresh and try again." };
    }
    const images = Array.isArray(incoming.images) ? incoming.images.slice(0, 3) : [];
    const photoUrls = [];
    for (let i = 0; i < images.length; i++) {
      const saved = observationSavePhoto_(currentRef, i, images[i]);
      if (saved && saved.success === false) return saved;
      if (saved && saved.url) photoUrls.push(saved.url);
    }
    const responseCol = observationEnsureColumn_(sheet, header, "Closeout Response");
    const byCol = observationEnsureColumn_(sheet, header, "Closed By");
    const photoCol = observationEnsureColumn_(sheet, header, "Closeout Photos");
    const stamp = Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyy-MM-dd HH:mm");
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet.getRange(row, statusCol + 1).setValue("Closed");
      sheet.getRange(row, responseCol + 1).setValue(response);
      sheet.getRange(row, byCol + 1).setValue(actor.name + " (" + stamp + ")");
      if (photoUrls.length) sheet.getRange(row, photoCol + 1).setValue(photoUrls.join("\n"));
    } finally {
      lock.releaseLock();
    }
    return { success: true, message: "Observation " + currentRef + " closed." };
  } catch (err) {
    const message = String(err);
    if (/permission|access|authorization|drive/i.test(message)) {
      return { success: false, message: "Share the observations sheet with the portal Google account as Editor, and allow Drive if a photo was added." };
    }
    return { success: false, message: message };
  }
}
