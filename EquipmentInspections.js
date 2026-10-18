/**
 * EQUIPMENT INSPECTION
 * Writes to the separate equipment register spreadsheet.
 * Failed checks also open an item on the main Action_Tracker.
 */
const EQUIPMENT_SPREADSHEET_ID = "12M-quo3ISq6KNg6c2jjtFW30EjuwofUJ-0jkQqtcLPk";
const EQUIPMENT_SHEET_GID = 555123456;

function recordEquipmentInspection(formData, user) {
  try {
    const equipSs = SpreadsheetApp.openById(EQUIPMENT_SPREADSHEET_ID);
    const equipSheet = equipSs.getSheets().find(function (sheet) {
      return sheet.getSheetId() === EQUIPMENT_SHEET_GID;
    });
    if (!equipSheet) {
      return { success: false, message: "Equipment sheet tab was not found." };
    }

    const now = new Date();
    const timeFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "M/d/yyyy H:mm:ss");
    const dateFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

    equipSheet.appendRow([
      timeFormatted,
      formData.contractor,
      formData.equipmentType,
      formData.brand,
      formData.plate,
      formData.equipThirdParty,
      formData.stickerNo,
      formatUsDate_(formData.equipIssue),
      formatUsDate_(formData.equipExpiry),
      formData.operatorName,
      formData.operatorId,
      formData.licence,
      formatUsDate_(formData.licenceExpiry),
      "",
      formData.operatorThirdParty,
      formatUsDate_(formData.operatorIssue),
      formatUsDate_(formData.operatorExpiry),
      formData.herc,
      formData.insurance,
      formData.contact,
      formData.inspectedBy || (user && user.name) || "",
      formData.remarks
    ]);

    const trackerSheet = SPREADSHEET.getSheetByName("Action_Tracker");
    let lastRow = trackerSheet.getLastRow();
    const place = [formData.contractor, formData.equipmentType, formData.plate].filter(Boolean).join(" / ");
    const checks = [
      ["3rd Party", formData.equipThirdParty],
      ["Licence", formData.licence],
      ["Operator 3rd Party", formData.operatorThirdParty],
      ["HERC", formData.herc],
      ["Insurance", formData.insurance]
    ];

    checks.forEach(function (check) {
      if (!isEquipmentFail_(check[1])) return;
      lastRow++;
      trackerSheet.appendRow([
        "ACT-" + (1000 + lastRow),
        dateFormatted,
        formData.zone,
        place,
        "Equipment " + check[0],
        String(check[1] || "") + (formData.remarks ? ". " + formData.remarks : ""),
        formData.targetDate || "",
        "Open",
        "",
        ""
      ]);
    });

    return { success: true, message: "Equipment inspection recorded successfully!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getEquipmentRegister() {
  try {
    const rows = readEquipmentRows_();
    const records = [];
    const headerRow = equipmentHeaderIndex_(rows);
    const outCol = equipmentColumnIndex_(rows[headerRow], "out of service");
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i];
      const contractor = String(row[1] || "").trim();
      const equipmentType = String(row[2] || "").trim();
      if (!contractor && !equipmentType && !String(row[4] || "").trim()) continue;
      const equipThirdParty = String(row[5] || "").trim();
      const licence = String(row[11] || "").trim();
      const operatorThirdParty = String(row[14] || "").trim();
      const herc = String(row[17] || "").trim();
      const insurance = String(row[18] || "").trim();
      const pending = isEquipmentFail_(equipThirdParty) || isEquipmentFail_(licence) ||
        isEquipmentFail_(operatorThirdParty) || isEquipmentFail_(herc) || isEquipmentFail_(insurance);
      records.push({
        sheetRow: i + 1,
        pending: pending,
        timestamp: formatCell_(row[0]),
        contractor: contractor,
        equipmentType: equipmentType,
        brand: String(row[3] || "").trim(),
        plate: String(row[4] || "").trim(),
        equipThirdParty: equipThirdParty,
        stickerNo: String(row[6] || "").trim(),
        equipIssue: formatCell_(row[7]),
        equipIssueIso: cellIso_(row[7]),
        equipExpiry: formatCell_(row[8]),
        equipExpiryIso: cellIso_(row[8]),
        operatorName: String(row[9] || "").trim(),
        operatorId: String(row[10] || "").trim(),
        licence: licence,
        licenceExpiry: formatCell_(row[12]),
        licenceExpiryIso: cellIso_(row[12]),
        operatorThirdParty: operatorThirdParty,
        operatorIssue: formatCell_(row[15]),
        operatorIssueIso: cellIso_(row[15]),
        operatorExpiry: formatCell_(row[16]),
        operatorExpiryIso: cellIso_(row[16]),
        herc: herc,
        insurance: insurance,
        contact: String(row[19] || "").trim(),
        inspectedBy: String(row[20] || "").trim(),
        remarks: String(row[21] || "").trim(),
        outOfService: outCol >= 0 && isOutOfService_(row[outCol]),
        category: "Equipment",
        statuses: [
          { label: "Equipment 3rd party", value: equipThirdParty },
          { label: "Licence", value: licence },
          { label: "Operator 3rd party", value: operatorThirdParty },
          { label: "Insurance", value: insurance },
          { label: "HERC", value: herc }
        ]
      });
    }
    records.reverse();
    return records;
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function updateEquipmentRecord(formData, user) {
  try {
    const row = parseInt(formData.sheetRow, 10);
    if (!row || row < 2) return { success: false, message: "Invalid equipment row." };

    const equipSs = SpreadsheetApp.openById(EQUIPMENT_SPREADSHEET_ID);
    const equipSheet = equipSs.getSheets().find(function (sheet) {
      return sheet.getSheetId() === EQUIPMENT_SHEET_GID;
    });
    if (!equipSheet) return { success: false, message: "Equipment sheet tab was not found." };

    const currentPlate = String(equipSheet.getRange(row, 5).getValue() || "").trim();
    const originalPlate = String(formData.originalPlate || "").trim();
    if (originalPlate && currentPlate && currentPlate !== originalPlate) {
      return { success: false, message: "This row changed in the sheet. Refresh and try again." };
    }

    const values = [
      [2, formData.contractor],
      [3, formData.equipmentType],
      [4, formData.brand],
      [5, formData.plate],
      [6, formData.equipThirdParty],
      [7, formData.stickerNo],
      [8, formatUsDate_(formData.equipIssue)],
      [9, formatUsDate_(formData.equipExpiry)],
      [10, formData.operatorName],
      [11, formData.operatorId],
      [12, formData.licence],
      [13, formatUsDate_(formData.licenceExpiry)],
      [15, formData.operatorThirdParty],
      [16, formatUsDate_(formData.operatorIssue)],
      [17, formatUsDate_(formData.operatorExpiry)],
      [18, formData.herc],
      [19, formData.insurance],
      [20, formData.contact],
      [21, formData.inspectedBy || (user && user.name) || ""],
      [22, formData.remarks]
    ];
    values.forEach(function (pair) {
      equipSheet.getRange(row, pair[0]).setValue(pair[1] || "");
    });
    writeOutOfService_(equipSheet, row, formData.outOfService);
    return { success: true, message: "Equipment record updated." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function removeEquipmentRecord(formData, user) {
  try {
    if (!equipmentActorIsManager_(user)) {
      return { success: false, message: "Only a manager can remove equipment." };
    }
    const row = parseInt(formData.sheetRow, 10);
    if (!row || row < 2) return { success: false, message: "Invalid equipment row." };

    const equipSs = SpreadsheetApp.openById(EQUIPMENT_SPREADSHEET_ID);
    const equipSheet = equipSs.getSheets().find(function (sheet) {
      return sheet.getSheetId() === EQUIPMENT_SHEET_GID;
    });
    if (!equipSheet) return { success: false, message: "Equipment sheet tab was not found." };

    const width = Math.max(equipSheet.getLastColumn(), 19);
    const values = equipSheet.getRange(row, 1, 1, width).getValues()[0];
    const currentPlate = String(values[4] || "").trim();
    const originalPlate = String(formData.originalPlate || "").trim();
    if (originalPlate && currentPlate && currentPlate !== originalPlate) {
      return { success: false, message: "This row changed in the sheet. Refresh and try again." };
    }
    const failed = isEquipmentFail_(values[5]) || isEquipmentFail_(values[11]) ||
      isEquipmentFail_(values[14]) || isEquipmentFail_(values[17]) || isEquipmentFail_(values[18]);
    if (!failed) {
      return { success: false, message: "Requirements are complete, so this equipment cannot be removed." };
    }

    equipSheet.deleteRow(row);
    return { success: true, message: "Equipment removed." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function equipmentActorIsManager_(user) {
  const username = String((user && user.username) || "").trim().toLowerCase();
  if (!username) return false;
  const sheet = SPREADSHEET.getSheetByName("HSE_TEAM");
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2] || "").trim().toLowerCase() !== username) continue;
    return String(data[i][5] || "").toLowerCase().indexOf("manager") !== -1;
  }
  return false;
}

function equipmentColumnIndex_(header, name) {
  const target = String(name || "").trim().toLowerCase();
  const cells = header || [];
  for (let i = 0; i < cells.length; i++) {
    if (String(cells[i] || "").trim().toLowerCase() === target) return i;
  }
  return -1;
}

function isOutOfService_(value) {
  const text = String(value || "").trim().toLowerCase();
  return text === "out of service" || text === "yes" || text === "true";
}

function writeOutOfService_(sheet, row, value) {
  const mark = isOutOfService_(value);
  const rows = sheet.getDataRange().getValues();
  const headerRow = equipmentHeaderIndex_(rows);
  let col = equipmentColumnIndex_(rows[headerRow], "out of service");
  if (col < 0) {
    if (!mark) return;
    col = Math.max(sheet.getLastColumn(), (rows[headerRow] || []).length);
    sheet.getRange(headerRow + 1, col + 1).setValue("Out of Service");
  }
  sheet.getRange(row, col + 1).setValue(mark ? "Out of Service" : "");
}

function cellIso_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const text = String(value).trim();
  let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    return match[3] + "-" + ("0" + match[1]).slice(-2) + "-" + ("0" + match[2]).slice(-2);
  }
  return "";
}

function equipmentHeaderIndex_(rows) {
  const limit = Math.min(rows.length, 6);
  for (let i = 0; i < limit; i++) {
    const stamp = String(rows[i][0] || "").trim().toLowerCase();
    const contractor = String(rows[i][1] || "").trim().toLowerCase();
    if (stamp === "timestamp" || contractor === "contractor name") return i;
  }
  return 0;
}

function readEquipmentRows_() {
  try {
    const equipSs = SpreadsheetApp.openById(EQUIPMENT_SPREADSHEET_ID);
    const equipSheet = equipSs.getSheets().find(function (sheet) {
      return sheet.getSheetId() === EQUIPMENT_SHEET_GID;
    });
    if (!equipSheet) throw new Error("Equipment sheet tab was not found.");
    return equipSheet.getDataRange().getValues();
  } catch (err) {
    const url = "https://docs.google.com/spreadsheets/d/" + EQUIPMENT_SPREADSHEET_ID +
      "/export?format=csv&gid=" + EQUIPMENT_SHEET_GID;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() !== 200) throw err;
    return Utilities.parseCsv(response.getContentText());
  }
}

function formatCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "d MMM yyyy");
  }
  return String(value || "").trim();
}

function formatUsDate_(iso) {
  if (!iso) return "";
  const parts = String(iso).split("-");
  if (parts.length !== 3) return String(iso);
  return Number(parts[1]) + "/" + Number(parts[2]) + "/" + parts[0];
}

function isEquipmentFail_(status) {
  const value = String(status || "").trim().toLowerCase();
  return value === "expired" || value === "not available" || value === "no";
}
