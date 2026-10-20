/**
 * Admin-only edits for HSE_TEAM logins.
 * Columns: ID, Name, Username, PIN, Zones, Role.
 */

function teamText_(value) {
  return String(value == null ? "" : value).trim();
}

function teamRoleIsAdmin_(value) {
  return teamText_(value).toLowerCase() === "hse dc";
}

function teamAdminError_(actor) {
  if (!actor || actor.denied === "signin") {
    return { success: false, message: "Sign in again to manage team logins." };
  }
  if (actor.denied === "sheet") return { success: false, message: "HSE_TEAM tab was not found." };
  if (actor.denied) return { success: false, message: "Only HSE DC can open team logins." };
  return null;
}

function teamAdminActor_(token) {
  const actor = chatActor_(token);
  if (!actor) return { denied: "signin" };
  const sheet = teamSheet_();
  if (!sheet) return { denied: "sheet" };
  const data = sheet.getDataRange().getValues();
  const me = chatNorm_(actor.username);
  for (let i = 1; i < data.length; i++) {
    if (chatNorm_(data[i][2]) !== me) continue;
    if (!teamRoleIsAdmin_(data[i][5])) return { denied: "role" };
    return { username: teamText_(data[i][2]), name: teamText_(data[i][1]) };
  }
  return { denied: "role" };
}

function teamBook_() {
  const sheet = teamSheet_();
  if (!sheet) return null;
  return { sheet: sheet, values: sheet.getDataRange().getValues() };
}

function teamFindRow_(values, username) {
  const key = chatNorm_(username);
  if (!key) return -1;
  for (let i = 1; i < values.length; i++) {
    if (chatNorm_(values[i][2]) === key) return i;
  }
  return -1;
}

function teamAdminCount_(values, skipUsername) {
  const skip = chatNorm_(skipUsername);
  let count = 0;
  for (let i = 1; i < values.length; i++) {
    if (skip && chatNorm_(values[i][2]) === skip) continue;
    if (teamRoleIsAdmin_(values[i][5])) count++;
  }
  return count;
}

function listTeamMembers(token) {
  const actor = teamAdminActor_(token);
  const denied = teamAdminError_(actor);
  if (denied) return denied;
  const book = teamBook_();
  if (!book) return { success: false, message: "HSE_TEAM tab was not found." };
  const people = [];
  for (let i = 1; i < book.values.length; i++) {
    const username = teamText_(book.values[i][2]);
    const name = teamText_(book.values[i][1]);
    if (!username && !name) continue;
    people.push({
      id: teamText_(book.values[i][0]),
      name: name,
      username: username,
      pin: teamText_(book.values[i][3]),
      zones: teamText_(book.values[i][4]),
      role: teamText_(book.values[i][5])
    });
  }
  people.sort(function (a, b) {
    return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
  });
  return { success: true, people: people };
}

function saveTeamMember(token, form) {
  const actor = teamAdminActor_(token);
  const denied = teamAdminError_(actor);
  if (denied) return denied;
  const book = teamBook_();
  if (!book) return { success: false, message: "HSE_TEAM tab was not found." };
  const incoming = form || {};
  const id = teamText_(incoming.id);
  const name = teamText_(incoming.name);
  const username = teamText_(incoming.username);
  const pin = teamText_(incoming.pin);
  const zones = teamText_(incoming.zones);
  const role = teamText_(incoming.role);
  const original = teamText_(incoming.originalUsername);
  if (!name || !username || !pin || !role) {
    return { success: false, message: "Name, username, login code, and role are required." };
  }
  if (name.length > 80 || username.length > 40 || pin.length > 20 || role.length > 40 || zones.length > 120) {
    return { success: false, message: "One of the fields is too long." };
  }
  const originalIndex = original ? teamFindRow_(book.values, original) : -1;
  if (original && originalIndex < 0) {
    return { success: false, message: "That login was not found. Refresh and try again." };
  }
  for (let i = 1; i < book.values.length; i++) {
    if (i === originalIndex) continue;
    if (chatNorm_(book.values[i][2]) === chatNorm_(username)) {
      return { success: false, message: "That username is already used." };
    }
  }
  const wasAdmin = originalIndex >= 0 && teamRoleIsAdmin_(book.values[originalIndex][5]);
  if (wasAdmin && !teamRoleIsAdmin_(role) && teamAdminCount_(book.values, original) < 1) {
    return { success: false, message: "Keep at least one HSE DC login." };
  }
  const rowNumber = originalIndex >= 0 ? originalIndex + 1 : book.sheet.getLastRow() + 1;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    book.sheet.getRange(rowNumber, 1, 1, 6).setNumberFormat("@");
    book.sheet.getRange(rowNumber, 1, 1, 6).setValues([[id, name, username, pin, zones, role]]);
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

function removeTeamMember(token, originalUsername) {
  const actor = teamAdminActor_(token);
  const denied = teamAdminError_(actor);
  if (denied) return denied;
  const key = teamText_(originalUsername);
  if (!key) return { success: false, message: "Choose a person first." };
  if (chatNorm_(key) === chatNorm_(actor.username)) {
    return { success: false, message: "You cannot remove your own login." };
  }
  const book = teamBook_();
  if (!book) return { success: false, message: "HSE_TEAM tab was not found." };
  const index = teamFindRow_(book.values, key);
  if (index < 1) return { success: false, message: "That login was not found. Refresh and try again." };
  if (teamRoleIsAdmin_(book.values[index][5]) && teamAdminCount_(book.values, key) < 1) {
    return { success: false, message: "Keep at least one HSE DC login." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    book.sheet.deleteRow(index + 1);
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}
