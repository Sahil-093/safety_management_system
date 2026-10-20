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
  if (cache.get(key) === "1") return "https://drive.google.com/thumbnail?id=" + id + "&sz=w800";
  try {
    DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    cache.put(key, "1", 21600);
    return "https://drive.google.com/thumbnail?id=" + id + "&sz=w800";
  } catch (err) {
    try { cache.put(key, "0", 300); } catch (ignore) {}
    return "";
  }
}

function observationMime_(response) {
  const headers = response.getHeaders() || {};
  return String(headers["Content-Type"] || headers["content-type"] || "");
}

function observationFetchImage_(url, token) {
  const response = UrlFetchApp.fetch(url, {
    headers: token ? { Authorization: "Bearer " + token } : {},
    muteHttpExceptions: true,
    followRedirects: true
  });
  const mime = observationMime_(response);
  if (response.getResponseCode() !== 200 || mime.indexOf("image/") !== 0) return null;
  const bytes = response.getContent();
  if (!bytes || !bytes.length || bytes.length > 180000) return null;
  return {
    type: mime.split(";")[0],
    bytes: bytes
  };
}

function observationInlineImage_(id) {
  if (!id) return "";
  const cache = CacheService.getScriptCache();
  const key = "obs-inline-" + id;
  const saved = cache.get(key);
  if (saved) return saved;
  let image = null;
  try {
    const token = ScriptApp.getOAuthToken();
    const meta = UrlFetchApp.fetch(
      "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(id) + "?fields=thumbnailLink",
      { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true }
    );
    const info = JSON.parse(meta.getContentText() || "{}");
    let link = String(info.thumbnailLink || "");
    if (link) {
      const large = /=s\d+/.test(link) ? link.replace(/=s\d+/, "=s800") : link;
      image = observationFetchImage_(large, "") || observationFetchImage_(large, token) || observationFetchImage_(link, "");
    }
  } catch (ignore) {}
  if (!image) {
    try {
      const thumb = DriveApp.getFileById(id).getThumbnail();
      if (thumb) {
        const bytes = thumb.getBytes();
        if (bytes && bytes.length && bytes.length <= 180000) {
          image = { type: thumb.getContentType() || "image/png", bytes: bytes };
        }
      }
    } catch (ignore) {}
  }
  if (!image) return observationPublish_(id);
  const data = "data:" + image.type + ";base64," + Utilities.base64Encode(image.bytes);
  if (data.length < 99000) {
    try { cache.put(key, data, 21600); } catch (ignore) {}
  }
  return data;
}

function observationPhotos_(cell) {
  return observationPhotoList_(cell).map(function (photo) {
    return {
      imageUrl: observationInlineImage_(photo.id),
      openUrl: photo.openUrl
    };
  });
}

function observationLoginZones_(actor) {
  const sheet = teamSheet_();
  if (!sheet || !actor) return [];
  const data = sheet.getDataRange().getValues();
  const me = chatNorm_(actor.username);
  for (let i = 1; i < data.length; i++) {
    if (chatNorm_(data[i][2]) !== me) continue;
    return observationText_(data[i][4]).split(",").map(observationText_).filter(Boolean);
  }
  return [];
}

function observationZoneAllowed_(zone, zones) {
  if (!zones.length) return false;
  const broad = zones.some(function (item) {
    const key = item.toLowerCase().replace(/\s+/g, "");
    return key === "all" || key === "allsite";
  });
  if (broad) return true;
  const value = observationText_(zone).toLowerCase();
  return zones.some(function (item) {
    const key = item.toLowerCase();
    return value === key || value === "zone " + key || ("zone " + value) === key;
  });
}

function observationOpenedMap_() {
  const raw = PropertiesService.getScriptProperties().getProperty("obs-opened-refs") || "";
  const map = {};
  raw.split(",").forEach(function (ref) {
    const key = observationText_(ref);
    if (key) map[key] = true;
  });
  return map;
}

function observationRememberOpened_(refs) {
  const map = observationOpenedMap_();
  refs.forEach(function (ref) {
    const key = observationText_(ref);
    if (key) map[key] = true;
  });
  const keys = Object.keys(map).slice(-800);
  PropertiesService.getScriptProperties().setProperty("obs-opened-refs", keys.join(","));
}

function observationStatusKey_(value) {
  const text = observationText_(value).toLowerCase();
  return text === "closed" || text === "close" ? "closed" : "open";
}

function observationRecord_(row, col, index, withPhotos) {
  const photosCell = observationCell_(row, col("photos"));
  const closeCell = observationCell_(row, col("closeout photos"));
  return {
    sheetRow: index + 2,
    ref: observationText_(observationCell_(row, col("ref. no."))),
    status: observationText_(observationCell_(row, col("status"))) || "Open",
    date: observationDate_(observationCell_(row, col("timestamp"))),
    when: observationWhen_(observationCell_(row, col("timestamp"))),
    contractor: observationText_(observationCell_(row, col("contractor/subcontractor"))),
    zone: observationText_(observationCell_(row, col("zone"))),
    building: observationText_(observationCell_(row, col("building/cluster"))),
    title: observationText_(observationCell_(row, col("observation title"))),
    villa: observationText_(observationCell_(row, col("villa/part"))),
    description: observationText_(observationCell_(row, col("observation description"))),
    responsible: observationText_(observationCell_(row, col("responsible person name"))),
    submittedBy: observationText_(observationCell_(row, col("submitted by"))),
    photoCount: observationPhotoList_(photosCell).length,
    closePhotoCount: observationPhotoList_(closeCell).length,
    photos: withPhotos ? observationPhotos_(photosCell) : [],
    closeout: observationText_(observationCell_(row, col("closeout response"))),
    closedBy: observationText_(observationCell_(row, col("closed by"))),
    closePhotos: withPhotos ? observationPhotos_(closeCell) : []
  };
}

function observationFilterActive_(form) {
  const incoming = form || {};
  const status = observationText_(incoming.status).toLowerCase();
  return !!((status && status !== "all") || observationText_(incoming.from) || observationText_(incoming.to) || observationText_(incoming.contractor) || observationText_(incoming.zone) || observationText_(incoming.building) || observationText_(incoming.q));
}

function observationMatchesFilter_(record, form) {
  const incoming = form || {};
  const status = observationText_(incoming.status).toLowerCase();
  if (status && status !== "all" && observationStatusKey_(record.status) !== (status === "close" ? "closed" : status)) return false;
  if (observationText_(incoming.contractor) && record.contractor !== observationText_(incoming.contractor)) return false;
  if (observationText_(incoming.zone) && record.zone !== observationText_(incoming.zone)) return false;
  if (observationText_(incoming.building) && record.building !== observationText_(incoming.building)) return false;
  if (observationText_(incoming.from) && (!record.date || record.date < observationText_(incoming.from))) return false;
  if (observationText_(incoming.to) && (!record.date || record.date > observationText_(incoming.to))) return false;
  const query = observationText_(incoming.q).toLowerCase();
  if (!query) return true;
  const haystack = [record.ref, record.title, record.description, record.contractor, record.zone, record.building, record.villa, record.submittedBy, record.responsible].join(" ").toLowerCase();
  return haystack.indexOf(query) !== -1;
}

function observationRows_(token) {
  const actor = chatActor_(token);
  if (!actor) return { success: false, message: "Sign in again to open observations." };
  const zones = observationLoginZones_(actor);
  const sheet = observationSheet_();
  if (!sheet) return { success: false, message: "Observations sheet was not found." };
  const header = observationHeader_(sheet);
  const last = sheet.getLastRow();
  const values = last < 2 ? [] : sheet.getRange(2, 1, last - 1, Math.max(sheet.getLastColumn(), 1)).getValues();
  const col = function (name) { return observationCol_(header.map, name); };
  const records = [];
  const choices = { contractors: {}, zones: {}, buildings: {} };
  for (let i = 0; i < values.length; i++) {
    const zone = observationText_(observationCell_(values[i], col("zone")));
    const title = observationText_(observationCell_(values[i], col("observation title")));
    const ref = observationText_(observationCell_(values[i], col("ref. no.")));
    if (!ref && !title) continue;
    if (!observationZoneAllowed_(zone, zones)) continue;
    const record = observationRecord_(values[i], col, i, false);
    records.push(record);
    if (record.contractor) choices.contractors[record.contractor] = true;
    if (record.zone) choices.zones[record.zone] = true;
    if (record.building) choices.buildings[record.building] = true;
  }
  records.sort(function (a, b) { return String(b.date || "").localeCompare(String(a.date || "")); });
  return {
    success: true,
    records: records,
    choices: {
      contractors: Object.keys(choices.contractors).sort(),
      zones: Object.keys(choices.zones).sort(),
      buildings: Object.keys(choices.buildings).sort()
    },
    values: values,
    col: col
  };
}

function listObservations(token, form) {
  const incoming = form || {};
  try {
    const book = observationRows_(token);
    if (!book.success) return book;
    let records = [];
    if (incoming.recent) {
      const opened = observationOpenedMap_();
      records = book.records.filter(function (row) { return !row.ref || !opened[row.ref]; });
      observationRememberOpened_(records.map(function (row) { return row.ref; }));
    } else if (observationFilterActive_(incoming)) {
      records = book.records.filter(function (row) { return observationMatchesFilter_(row, incoming); });
    }
    return { success: true, records: records, choices: book.choices };
  } catch (err) {
    const message = String(err);
    if (/permission|access|authorization/i.test(message)) {
      return { success: false, message: "Share the observations sheet with the portal Google account as Editor." };
    }
    return { success: false, message: message };
  }
}

function getObservation(token, form) {
  const incoming = form || {};
  const ref = observationText_(incoming.ref);
  const sheetRow = parseInt(incoming.sheetRow, 10);
  if (!ref && !sheetRow) return { success: false, message: "That observation was not found." };
  try {
    const book = observationRows_(token);
    if (!book.success) return book;
    let found = -1;
    for (let i = 0; i < book.records.length; i++) {
      if (ref && book.records[i].ref === ref) { found = i; break; }
      if (!ref && book.records[i].sheetRow === sheetRow) { found = i; break; }
    }
    if (found < 0) return { success: false, message: "That observation was not found." };
    const light = book.records[found];
    const sourceIndex = light.sheetRow - 2;
    const record = observationRecord_(book.values[sourceIndex], book.col, sourceIndex, true);
    record.photosLoaded = true;
    return { success: true, record: record };
  } catch (err) {
    const message = String(err);
    if (/permission|access|authorization|drive/i.test(message)) {
      return { success: false, message: "Share the observations sheet with the portal Google account as Editor, and allow Drive if a photo was added." };
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

function observationFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function observationActorEmail_(actor) {
  if (!actor || !actor.username) return "";
  const sheet = teamSheet_();
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  const header = observationText_(data[0][6]).toLowerCase();
  if (header && header !== "email") return "";
  const me = chatNorm_(actor.username);
  for (let i = 1; i < data.length; i++) {
    if (chatNorm_(data[i][2]) !== me) continue;
    return observationText_(data[i][6]);
  }
  return "";
}

function observationCode_(contractor) {
  const letters = observationText_(contractor).toUpperCase().replace(/[^A-Z]/g, "");
  return (letters + "XXX").slice(0, 3);
}

function observationSavePhoto_(ref, index, image, folderName) {
  const raw = String(image || "").replace(/^data:image\/\w+;base64,/, "");
  if (!raw) return { url: "" };
  if (raw.length > 120000) return { success: false, message: "A picture is too large. Choose a smaller photo." };
  const folder = observationFolder_(folderName || OBSERVATION_PHOTO_FOLDER);
  const safeRef = String(ref || "observation").replace(/[^\w.-]/g, "").slice(0, 40) || "observation";
  const name = safeRef + "-" + (index + 1) + "-" + new Date().getTime() + ".jpg";
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(raw), "image/jpeg", name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    url: "https://lh3.googleusercontent.com/d/" + file.getId() + "=w800",
    id: file.getId()
  };
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
    const emailCol = observationEnsureColumn_(sheet, header, "Closeout Email");
    const stamp = Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyy-MM-dd HH:mm");
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet.getRange(row, statusCol + 1).setValue("Closed");
      sheet.getRange(row, responseCol + 1).setValue(response);
      sheet.getRange(row, byCol + 1).setValue(actor.name + " (" + stamp + ")");
      sheet.getRange(row, emailCol + 1).setValue(observationActorEmail_(actor));
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

function createObservation(token, form) {
  const actor = chatActor_(token);
  if (!actor) return { success: false, message: "Sign in again to issue an observation." };
  const incoming = form || {};
  const contractor = observationText_(incoming.contractor);
  const zone = observationText_(incoming.zone);
  const building = observationText_(incoming.building);
  const title = observationText_(incoming.title);
  const villa = observationText_(incoming.villa);
  const description = observationText_(incoming.description);
  const responsible = observationText_(incoming.responsible);
  const email = observationActorEmail_(actor);
  if (!contractor) return { success: false, message: "Enter the contractor or subcontractor." };
  if (!zone) return { success: false, message: "Enter the zone." };
  if (!title) return { success: false, message: "Enter the observation title." };
  if (!description) return { success: false, message: "Enter the observation description." };
  if (title.length > 200 || description.length > 2000 || contractor.length > 80 || zone.length > 40 || building.length > 80 || villa.length > 80 || responsible.length > 80) {
    return { success: false, message: "One of the fields is too long." };
  }
  try {
    const sheet = observationSheet_();
    if (!sheet) return { success: false, message: "Observations sheet was not found." };
    const header = observationHeader_(sheet);
    const refCol = observationCol_(header.map, "ref. no.");
    if (refCol < 0 || observationCol_(header.map, "status") < 0) {
      return { success: false, message: "Observation columns were not found." };
    }
    const images = Array.isArray(incoming.images) ? incoming.images.slice(0, 3) : [];
    const photoUrls = [];
    for (let i = 0; i < images.length; i++) {
      const saved = observationSavePhoto_("issue", i, images[i], "SBARC Observations");
      if (saved && saved.success === false) return saved;
      if (saved && saved.id) photoUrls.push("https://drive.google.com/open?id=" + saved.id);
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    let ref = "";
    try {
      const last = sheet.getLastRow();
      let max = 0;
      if (last >= 2) {
        const refs = sheet.getRange(2, refCol + 1, last - 1, 1).getValues();
        refs.forEach(function (item) {
          const match = observationText_(item[0]).match(/(\d+)\s*$/);
          if (match) max = Math.max(max, parseInt(match[1], 10) || 0);
        });
      }
      ref = "SO-MAK-" + observationCode_(contractor) + "-" + ("0000" + (max + 1)).slice(-4);
      const width = Math.max(sheet.getLastColumn(), header.headers.length, 1);
      const rowValues = [];
      for (let c = 0; c < width; c++) rowValues.push("");
      const put = function (name, value) {
        const col = observationCol_(header.map, name);
        if (col >= 0 && col < rowValues.length) rowValues[col] = value;
      };
      put("timestamp", new Date());
      put("ref. no.", ref);
      put("status", "Open");
      put("contractor/subcontractor", contractor);
      put("zone", zone);
      put("building/cluster", building);
      put("observation title", title);
      put("villa/part", villa);
      put("observation description", description);
      if (photoUrls.length) put("photos", photoUrls.join("\n"));
      put("responsible person name", responsible);
      put("email address", email);
      put("submitted by", actor.name);
      sheet.appendRow(rowValues);
    } finally {
      lock.releaseLock();
    }
    return { success: true, ref: ref, message: "Observation " + ref + " issued." };
  } catch (err) {
    const message = String(err);
    if (/permission|access|authorization|drive/i.test(message)) {
      return { success: false, message: "Share the observations sheet with the portal Google account as Editor, and allow Drive if a photo was added." };
    }
    return { success: false, message: message };
  }
}
