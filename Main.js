/**
 * MAIN CONTROLLER
 * Handles Routing, Core Configuration, and Authentication
 */
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const TEAM_SPREADSHEET_ID = "1Rip6zKnc9gbHpX-9P3GF69jeYQqu3uH3p6MRF2zTr18";

function teamSpreadsheet_() {
  return SpreadsheetApp.openById(TEAM_SPREADSHEET_ID);
}

function teamSheet_() {
  return teamSpreadsheet_().getSheetByName("HSE_TEAM");
}

var UPDATED_BY_HEADER = "Updated By";
var UPDATED_AT_HEADER = "Updated Date & Time";

// Columns are found by header name and added after the last used column, never before minColumn.
function stampRowUpdate_(sheet, row, headerRow, who, minColumn) {
  const width = Math.max(sheet.getLastColumn(), minColumn || 0, 1);
  if (width > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  const header = sheet.getRange(headerRow, 1, 1, width).getValues()[0].map(function (cell) {
    return String(cell || "").trim().toLowerCase();
  });
  let byCol = header.indexOf(UPDATED_BY_HEADER.toLowerCase()) + 1;
  let atCol = header.indexOf(UPDATED_AT_HEADER.toLowerCase()) + 1;
  let next = width;
  const need = (byCol ? 0 : 1) + (atCol ? 0 : 1);
  if (need && next + need > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), next + need - sheet.getMaxColumns());
  }
  if (!byCol) {
    byCol = ++next;
    sheet.getRange(headerRow, byCol).setValue(UPDATED_BY_HEADER).setFontWeight("bold");
  }
  if (!atCol) {
    atCol = ++next;
    sheet.getRange(headerRow, atCol).setValue(UPDATED_AT_HEADER).setFontWeight("bold");
  }
  const stamp = Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyy-MM-dd HH:mm");
  sheet.getRange(row, byCol).setValue(String(who || "").trim());
  sheet.getRange(row, atCol).setNumberFormat("@").setValue(stamp);
}

function updaterName_(user) {
  return String((user && (user.name || user.username)) || "").trim();
}

var UPDATE_LOG_SHEET = "UPDATE_LOG";
var UPDATE_LOG_HEADERS = ["Date & Time", "Updated By", "Username", "Section", "Sheet / Tab", "Row", "Record", "Changes"];
var UPDATE_LOG_VALUE_MAX = 300;

function rowSnapshot_(sheet, row) {
  const width = sheet.getLastColumn();
  if (!width || row > sheet.getMaxRows()) return [];
  return sheet.getRange(row, 1, 1, width).getDisplayValues()[0];
}

function updateLogSheet_() {
  const ss = teamSpreadsheet_();
  let sheet = ss.getSheetByName(UPDATE_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(UPDATE_LOG_SHEET);
    sheet.getRange(1, 1, 1, UPDATE_LOG_HEADERS.length).setValues([UPDATE_LOG_HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function updateLogValue_(value) {
  const text = String(value === undefined || value === null ? "" : value).trim();
  if (!text) return "(blank)";
  return text.length > UPDATE_LOG_VALUE_MAX ? text.slice(0, UPDATE_LOG_VALUE_MAX) + "..." : text;
}

// Login codes are never written to the log, only that they changed.
function updateLogChanges_(headers, before, after) {
  const skip = [UPDATED_BY_HEADER.toLowerCase(), UPDATED_AT_HEADER.toLowerCase()];
  const changes = [];
  const width = Math.max(headers.length, before.length, after.length);
  for (let i = 0; i < width; i++) {
    const label = String(headers[i] || "").trim() || "Column " + (i + 1);
    const key = label.toLowerCase();
    if (skip.indexOf(key) !== -1) continue;
    const oldValue = String(before[i] === undefined ? "" : before[i]).trim();
    const newValue = String(after[i] === undefined ? "" : after[i]).trim();
    if (oldValue === newValue) continue;
    if (key === "pin" || key === "login code") {
      changes.push(label + ": changed");
      continue;
    }
    changes.push(label + ": " + updateLogValue_(oldValue) + " → " + updateLogValue_(newValue));
  }
  return changes;
}

function logRowUpdate_(sheet, row, headerRow, before, info) {
  const after = rowSnapshot_(sheet, row);
  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, width).getDisplayValues()[0];
  const changes = updateLogChanges_(headers, before || [], after);
  updateLogSheet_().appendRow([
    Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyy-MM-dd HH:mm:ss"),
    info.who || "",
    info.username || "",
    info.section || "",
    sheet.getParent().getName() + " / " + sheet.getName(),
    row,
    info.record || "",
    changes.length ? changes.join("\n") : "No field changed"
  ]);
}

// Overwrites the row's Updated By / Updated Date & Time and appends the full change to UPDATE_LOG.
function recordRowUpdate_(sheet, row, headerRow, before, info, minColumn) {
  stampRowUpdate_(sheet, row, headerRow, info.who, minColumn);
  try {
    logRowUpdate_(sheet, row, headerRow, before, info);
  } catch (err) {
    Logger.log("Update log failed: " + err);
  }
}

// Web App entry point. action=... serves JSON for the mobile page.
function doPost(e) {
  let params = {};
  try {
    params = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  } catch (err) {
    params = {};
  }
  return apiResponse_(params);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action) {
    return apiResponse_(params);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Safety Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function apiResponse_(params) {
  let result;
  try {
    const user = params.user ? JSON.parse(params.user) : {};
    const payload = params.payload ? JSON.parse(params.payload) : {};
    if (params.action === "login") {
      result = authenticateUser(params.username || "", params.pin || "");
    } else if (params.action === "actions") {
      result = getZoneActionItems(user);
    } else if (params.action === "closeout") {
      result = submitCloseout(payload, user);
    } else if (params.action === "inspection") {
      result = recordGeneratorInspection(payload, user);
    } else if (params.action === "equipment") {
      result = recordEquipmentInspection(payload, user);
    } else if (params.action === "equipmentList") {
      result = getEquipmentRegister();
    } else if (params.action === "registers") {
      const equipment = getEquipmentRegister();
      const generator = getGeneratorRegister();
      const scaffold = getScaffoldRegister();
      if (equipment && equipment.success === false) result = equipment;
      else if (generator && generator.success === false) result = generator;
      else if (scaffold && scaffold.success === false) result = scaffold;
      else result = { equipment: equipment, generator: generator, scaffold: scaffold };
    } else if (params.action === "equipmentUpdate") {
      result = updateEquipmentRecord(payload, user);
    } else if (params.action === "equipmentRemove") {
      result = removeEquipmentRecord(payload, user);
    } else if (params.action === "generatorUpdate") {
      result = updateGeneratorRecord(payload, user);
    } else if (params.action === "inductionSearch") {
      result = searchInductionLog(params.q || "");
    } else if (params.action === "inductionList") {
      result = { records: listInductionEmployees() };
    } else if (params.action === "induction") {
      result = recordInduction(payload, user);
    } else if (params.action === "inductionUpdate") {
      result = updateInductionRecord(payload, user);
    } else if (params.action === "inductionDelete") {
      result = deleteInductionRecord(payload);
    } else if (params.action === "inductionPhoto") {
      result = uploadInductionPhoto(payload);
    } else if (params.action === "inductionPhotoFolder") {
      result = inductionPhotoFolderLink_();
    } else if (params.action === "chatPeople") {
      result = listChatPeople(params.token || "");
    } else if (params.action === "chatThread") {
      result = listChatThread(params.token || "", params.with || "");
    } else if (params.action === "chatSend") {
      result = sendChatMessage(params.token || "", params.with || "", params.text || "");
    } else if (params.action === "teamList") {
      result = listTeamMembers(params.token || "");
    } else if (params.action === "teamSave") {
      result = saveTeamMember(params.token || "", payload);
    } else if (params.action === "teamRemove") {
      result = removeTeamMember(params.token || "", params.username || "");
    } else if (params.action === "profileGet") {
      result = getProfile(params.token || "");
    } else if (params.action === "profileSave") {
      result = saveProfile(params.token || "", payload);
    } else if (params.action === "credentialRequest") {
      result = requestCredential(params.lookup || "");
    } else if (params.action === "credentialPending") {
      result = listCredentialRequests(params.token || "");
    } else if (params.action === "credentialClear") {
      result = clearCredentialRequest(params.token || "", params.username || "");
    } else if (params.action === "observationList") {
      result = listObservations(params.token || "", payload);
    } else if (params.action === "observationGet") {
      result = getObservation(params.token || "", payload);
    } else if (params.action === "observationClose") {
      result = closeObservation(params.token || "", payload);
    } else if (params.action === "observationIssue") {
      result = createObservation(params.token || "", payload);
    } else {
      result = { success: false, message: "Unknown action." };
    }
  } catch (err) {
    result = { success: false, message: String(err) };
  }

  const body = JSON.stringify(result);
  const callback = String(params.callback || "");
  if (/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + body + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

// Global user authentication (from HSE_TEAM sheet)
function authenticateUser(username, pin) {
  const sheet = teamSheet_();
  if (!sheet) return { success: false, message: "HSE_TEAM tab was not found." };
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const uName = String(data[i][2]).trim().toLowerCase();
    const uPin = String(data[i][3]).trim();
    
    if (uName === username.trim().toLowerCase() && uPin === pin.trim()) {
      let sameUsername = 0;
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][2]).trim().toLowerCase() === uName) sameUsername++;
      }
      if (sameUsername > 1) {
        return { success: false, message: "This username is used by more than one person. Ask HSE DC to give you a different username." };
      }
      const assignedZonesStr = String(data[i][4]);
      const user = {
        id: data[i][0],
        name: data[i][1],
        username: data[i][2],
        assignedZones: assignedZonesStr.split(',').map(s => s.trim()),
        role: data[i][5],
        email: String(data[i][6] || "").trim(),
        photo: String(data[i][7] || "").trim()
      };
      if (String(data[i][5] || "").trim().toLowerCase() === "hse dc") {
        let pending = 0;
        for (let r = 1; r < data.length; r++) {
          if (String(data[r][8] || "").trim()) pending++;
        }
        user.credentialRequests = pending;
      }
      try {
        user.token = issueChatSession_(data[i][2], data[i][1], data[i][5]);
      } catch (err) {
        user.token = "";
      }
      return { success: true, user: user };
    }
  }
  return { success: false, message: "Invalid username or PIN." };
}