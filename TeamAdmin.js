/**
 * HSE_TEAM logins and each person's profile.
 * Columns: ID, Name, Username, PIN, Zones, Role, Email, Photo, Credential request.
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

function teamHeader_(sheet, column, label) {
  const cell = sheet.getRange(1, column);
  if (!teamText_(cell.getValue())) cell.setValue(label).setFontWeight("bold");
}

function teamRowProfile_(row) {
  return {
    id: teamText_(row[0]),
    name: teamText_(row[1]),
    username: teamText_(row[2]),
    pin: teamText_(row[3]),
    zones: teamText_(row[4]),
    role: teamText_(row[5]),
    email: teamText_(row[6]),
    photo: teamText_(row[7]),
    credentialRequest: teamText_(row[8])
  };
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
    people.push(teamRowProfile_(book.values[i]));
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
  const email = teamText_(incoming.email);
  const original = teamText_(incoming.originalUsername);
  if (!name || !username || !pin || !role) {
    return { success: false, message: "Name, username, login code, and role are required." };
  }
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, message: "Enter a valid email, or leave it blank." };
  }
  if (name.length > 80 || username.length > 40 || pin.length > 20 || role.length > 40 || zones.length > 120 || email.length > 120) {
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
    if (!teamText_(book.sheet.getRange(1, 7).getValue())) {
      book.sheet.getRange(1, 7).setValue("Email").setFontWeight("bold");
    }
    book.sheet.getRange(rowNumber, 1, 1, 7).setNumberFormat("@");
    book.sheet.getRange(rowNumber, 1, 1, 7).setValues([[id, name, username, pin, zones, role, email]]);
    if (original && original !== username) chatRenameUser_(original, username, name);
    if (original && chatNorm_(actor.username) === chatNorm_(original)) {
      refreshChatSession_(token, username, name, role);
    }
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

function profileSavePhoto_(username, image) {
  const raw = String(image || "").replace(/^data:image\/\w+;base64,/, "");
  if (!raw) return { url: "" };
  if (raw.length > 120000) return { success: false, message: "The picture is too large. Choose a smaller photo." };
  try {
    const folder = observationFolder_("SBARC Profile Photos");
    const safe = String(username || "profile").replace(/[^\w.-]/g, "").slice(0, 40) || "profile";
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(raw), "image/jpeg", safe + "-" + new Date().getTime() + ".jpg"));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: "https://lh3.googleusercontent.com/d/" + file.getId() + "=w400" };
  } catch (err) {
    return { success: false, message: "The photo could not be saved. Allow Drive for the portal Google account." };
  }
}

function getProfile(token) {
  const actor = chatActor_(token);
  if (!actor) return { success: false, message: "Sign in again to open your profile." };
  const book = teamBook_();
  if (!book) return { success: false, message: "HSE_TEAM tab was not found." };
  const index = teamFindRow_(book.values, actor.username);
  if (index < 1) return { success: false, message: "Your login was not found." };
  const profile = teamRowProfile_(book.values[index]);
  delete profile.credentialRequest;
  return { success: true, profile: profile };
}

function saveProfile(token, form) {
  const actor = chatActor_(token);
  if (!actor) return { success: false, message: "Sign in again to update your profile." };
  const incoming = form || {};
  const name = teamText_(incoming.name);
  const username = teamText_(incoming.username);
  const pin = teamText_(incoming.pin);
  const email = teamText_(incoming.email);
  if (!name || !username || !pin) {
    return { success: false, message: "Name, username, and login code are required." };
  }
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, message: "Enter a valid email. Observation emails are taken from your profile." };
  }
  if (name.length > 80 || username.length > 40 || pin.length > 20 || email.length > 120) {
    return { success: false, message: "One of the fields is too long." };
  }
  let photoUrl = null;
  if (incoming.removePhoto) photoUrl = "";
  if (incoming.image) {
    const saved = profileSavePhoto_(username, incoming.image);
    if (saved && saved.success === false) return saved;
    if (saved && saved.url) photoUrl = saved.url;
  }
  const sheet = teamSheet_();
  if (!sheet) return { success: false, message: "HSE_TEAM tab was not found." };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const values = sheet.getDataRange().getValues();
    const index = teamFindRow_(values, actor.username);
    if (index < 1) return { success: false, message: "Your login was not found." };
    for (let i = 1; i < values.length; i++) {
      if (i === index) continue;
      if (chatNorm_(values[i][2]) === chatNorm_(username)) {
        return { success: false, message: "That username is already used." };
      }
    }
    const rowNumber = index + 1;
    const id = teamText_(values[index][0]);
    const zones = teamText_(values[index][4]);
    const role = teamText_(values[index][5]);
    const oldUsername = teamText_(values[index][2]);
    if (photoUrl === null) photoUrl = teamText_(values[index][7]);
    teamHeader_(sheet, 7, "Email");
    teamHeader_(sheet, 8, "Photo");
    sheet.getRange(rowNumber, 2, 1, 3).setNumberFormat("@");
    sheet.getRange(rowNumber, 2, 1, 3).setValues([[name, username, pin]]);
    sheet.getRange(rowNumber, 7).setNumberFormat("@").setValue(email);
    sheet.getRange(rowNumber, 8).setValue(photoUrl);
    if (oldUsername !== username || teamText_(values[index][1]) !== name) {
      chatRenameUser_(oldUsername, username, name);
    }
    refreshChatSession_(token, username, name, role);
    return {
      success: true,
      profile: {
        id: id,
        name: name,
        username: username,
        pin: pin,
        zones: zones,
        role: role,
        email: email,
        photo: photoUrl
      }
    };
  } catch (err) {
    return { success: false, message: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function requestCredential(lookup) {
  const key = teamText_(lookup);
  if (!key || key.length > 80) return { success: false, message: "Enter your username or name." };
  const sheet = teamSheet_();
  if (!sheet) return { success: false, message: "HSE_TEAM tab was not found." };
  const cache = CacheService.getScriptCache();
  const guard = "credreq-" + chatNorm_(key).slice(0, 40);
  if (cache.get(guard)) {
    return { success: true, message: "Request sent to HSE DC. They can give you the login code." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const values = sheet.getDataRange().getValues();
    let index = teamFindRow_(values, key);
    if (index < 1) {
      const matches = [];
      for (let i = 1; i < values.length; i++) {
        if (chatNorm_(values[i][1]) === chatNorm_(key)) matches.push(i);
      }
      if (matches.length > 1) return { success: false, message: "More than one person has that name. Enter the username." };
      if (matches.length === 1) index = matches[0];
    }
    if (index < 1) return { success: false, message: "No login matched that name or username." };
    teamHeader_(sheet, 9, "Credential request");
    const stamp = Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyy-MM-dd HH:mm");
    sheet.getRange(index + 1, 9).setNumberFormat("@").setValue(stamp);
    cache.put(guard, "1", 60);
    return { success: true, message: "Request sent to HSE DC. They can give you the login code." };
  } catch (err) {
    return { success: false, message: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function listCredentialRequests(token) {
  const actor = teamAdminActor_(token);
  const denied = teamAdminError_(actor);
  if (denied) return denied;
  const book = teamBook_();
  if (!book) return { success: false, message: "HSE_TEAM tab was not found." };
  const people = [];
  for (let i = 1; i < book.values.length; i++) {
    const requested = teamText_(book.values[i][8]);
    if (!requested) continue;
    const row = teamRowProfile_(book.values[i]);
    if (!row.username && !row.name) continue;
    people.push({
      name: row.name,
      username: row.username,
      pin: row.pin,
      requestedAt: requested
    });
  }
  people.sort(function (a, b) {
    return a.requestedAt < b.requestedAt ? 1 : -1;
  });
  return { success: true, people: people };
}

function clearCredentialRequest(token, username) {
  const actor = teamAdminActor_(token);
  const denied = teamAdminError_(actor);
  if (denied) return denied;
  const key = teamText_(username);
  if (!key) return { success: false, message: "Choose a person first." };
  const sheet = teamSheet_();
  if (!sheet) return { success: false, message: "HSE_TEAM tab was not found." };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const values = sheet.getDataRange().getValues();
    const index = teamFindRow_(values, key);
    if (index < 1) return { success: false, message: "That login was not found. Refresh and try again." };
    sheet.getRange(index + 1, 9).setValue("");
  } catch (err) {
    return { success: false, message: String(err) };
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}
