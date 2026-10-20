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
    .setTitle('HSE Inspection & Compliance Portal')
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
      result = { records: searchInductionLog(params.q || "") };
    } else if (params.action === "inductionList") {
      result = { records: listInductionEmployees() };
    } else if (params.action === "induction") {
      result = recordInduction(payload, user);
    } else if (params.action === "inductionUpdate") {
      result = updateInductionRecord(payload);
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
      const assignedZonesStr = String(data[i][4]);
      const user = {
        id: data[i][0],
        name: data[i][1],
        username: data[i][2],
        assignedZones: assignedZonesStr.split(',').map(s => s.trim()),
        role: data[i][5]
      };
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