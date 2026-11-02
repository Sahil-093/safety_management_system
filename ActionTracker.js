/**
 * ACTION TRACKER MODULE
 * Category-wise filtering (Generator, Scaffolding, Equipment) & Closeout Engine
 */

function getZoneActionItems(user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Action_Tracker");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const filtered = [];
    
    const userRole = String(user.role || "").trim().toLowerCase();
    const userZones = Array.isArray(user.assignedZones) 
      ? user.assignedZones.map(z => String(z).trim().toLowerCase()) 
      : [String(user.assignedZones || "").trim().toLowerCase()];

    const isAdmin = userRole === "hse dc" || 
                    userRole === "hse manager" || 
                    userRole === "admin" || 
                    userZones.includes("all");
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[7] || "").trim(); // Col H: Status
      const zone = String(row[2] || "").trim();   // Col C: Zone
      const zoneLower = zone.toLowerCase();

      if (status === "Open" || status === "In Progress") {
        const hasAccess = isAdmin || userZones.includes(zoneLower);

        if (hasAccess) {
          // Safe Date Formatting for Date Reported
          let repDate = row[1];
          if (repDate instanceof Date) {
            repDate = Utilities.formatDate(repDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            repDate = String(repDate || "");
          }

          // Safe Date Formatting for Target Date
          let tgtDate = row[6];
          if (tgtDate instanceof Date) {
            tgtDate = Utilities.formatDate(tgtDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            tgtDate = String(tgtDate || "");
          }

          const deficiencyItem = String(row[4] || "").trim();
          const clusterDesc = String(row[3] || "").trim();

          // Auto Category Identification
          let category = "Generator"; // default for current data
          const combinedText = (deficiencyItem + " " + clusterDesc).toLowerCase();
          
          if (combinedText.includes("scaffold") || combinedText.includes("tag") || combinedText.includes("plank")) {
            category = "Scaffolding";
          } else if (combinedText.includes("excavator") || combinedText.includes("crane") || combinedText.includes("bobcat") || combinedText.includes("equipment") || combinedText.includes("roller")) {
            category = "Equipment";
          } else {
            category = "Generator";
          }

          filtered.push({
            rowIndex: i + 1,
            actionId: String(row[0] || ""),
            dateReported: repDate,
            zone: zone,
            cluster: clusterDesc,
            category: category,
            item: deficiencyItem,
            finding: String(row[5] || ""),
            targetDate: tgtDate,
            status: status
          });
        }
      }
    }
    return filtered;
  } catch (err) {
    Logger.log("Error in getZoneActionItems: " + err.toString());
    return [];
  }
}

function submitCloseout(formData, user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Action_Tracker");
    const row = parseInt(formData.rowIndex);
    const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    const before = rowSnapshot_(sheet, row);
    sheet.getRange(row, 8).setValue(formData.status);
    sheet.getRange(row, 9).setValue(formData.remarks);
    sheet.getRange(row, 10).setValue(user.name + " (" + timeStamp + ")");
    recordRowUpdate_(sheet, row, 1, before, {
      who: updaterName_(user),
      username: (user && user.username) || "",
      section: "Action Tracker",
      record: String(formData.actionId || "")
    });

    return { 
      success: true, 
      message: `Action ${formData.actionId} marked as ${formData.status}.` 
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}