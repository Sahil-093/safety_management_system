/**
 * Private text chat between two signed-in team members.
 * TEAM_CHAT columns: Time, From, To, From name, Text.
 */
var CHAT_SHEET_NAME = "TEAM_CHAT";
var CHAT_TEXT_MAX = 500;
var CHAT_SCAN_MAX = 5000;

function chatSessionKey_(token) {
  return "hsechat-" + token;
}

function issueChatSession_(username, name, role) {
  const token = Utilities.getUuid();
  const payload = JSON.stringify({
    username: String(username || "").trim(),
    name: String(name || username || "").trim(),
    role: String(role || "").trim()
  });
  CacheService.getScriptCache().put(chatSessionKey_(token), payload, 21600);
  return token;
}

function chatActor_(token) {
  const key = String(token || "").trim();
  if (!key || key.length > 80) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get(chatSessionKey_(key));
  if (!raw) return null;
  let actor;
  try {
    actor = JSON.parse(raw);
  } catch (err) {
    return null;
  }
  if (!actor || !actor.username) return null;
  cache.put(chatSessionKey_(key), raw, 21600);
  return {
    username: String(actor.username).trim(),
    name: String(actor.name || actor.username).trim(),
    role: String(actor.role || "").trim()
  };
}

function chatNorm_(value) {
  return String(value || "").trim().toLowerCase();
}

function refreshChatSession_(token, username, name, role) {
  const key = String(token || "").trim();
  if (!key) return;
  const payload = JSON.stringify({
    username: String(username || "").trim(),
    name: String(name || username || "").trim(),
    role: String(role || "").trim()
  });
  CacheService.getScriptCache().put(chatSessionKey_(key), payload, 21600);
}

function chatRenameUser_(fromUsername, toUsername, toName) {
  const sheet = teamSpreadsheet_().getSheetByName(CHAT_SHEET_NAME);
  if (!sheet) return;
  const last = sheet.getLastRow();
  if (last < 2) return;
  const values = sheet.getRange(2, 2, last - 1, 3).getValues();
  const oldKey = chatNorm_(fromUsername);
  const nextUser = String(toUsername || "").trim();
  const nextName = String(toName || "").trim();
  if (!oldKey || !nextUser) return;
  let changed = false;
  for (let i = 0; i < values.length; i++) {
    if (chatNorm_(values[i][0]) === oldKey) {
      if (values[i][0] !== nextUser) {
        values[i][0] = nextUser;
        changed = true;
      }
      if (nextName && values[i][2] !== nextName) {
        values[i][2] = nextName;
        changed = true;
      }
    }
    if (chatNorm_(values[i][1]) === oldKey && values[i][1] !== nextUser) {
      values[i][1] = nextUser;
      changed = true;
    }
  }
  if (changed) sheet.getRange(2, 2, last - 1, 3).setValues(values);
}

function chatDeny_() {
  return { success: false, message: "Sign in again to use chat." };
}

function chatSheet_() {
  const book = teamSpreadsheet_();
  let sheet = book.getSheetByName(CHAT_SHEET_NAME);
  if (sheet) return sheet;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet = book.getSheetByName(CHAT_SHEET_NAME);
    if (!sheet) {
      sheet = book.insertSheet(CHAT_SHEET_NAME);
      sheet.getRange(1, 1, 1, 5).setNumberFormat("@");
      sheet.getRange(1, 1, 1, 5).setValues([["Time", "From", "To", "From name", "Text"]]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }
  } finally {
    lock.releaseLock();
  }
  return sheet;
}

function chatTeam_() {
  const sheet = teamSpreadsheet_().getSheetByName("HSE_TEAM");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const people = [];
  const seen = {};
  for (let i = 1; i < data.length; i++) {
    const username = String(data[i][2] || "").trim();
    const key = chatNorm_(username);
    if (!key || seen[key]) continue;
    seen[key] = true;
    people.push({
      username: username,
      name: String(data[i][1] || username).trim() || username,
      role: String(data[i][5] || "").trim()
    });
  }
  return people;
}

function chatTimeIso_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value.toISOString();
  }
  const text = String(value || "").trim();
  if (!text) return "";
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return parsed.toISOString();
  return "";
}

function chatRows_() {
  const sheet = teamSpreadsheet_().getSheetByName(CHAT_SHEET_NAME);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - CHAT_SCAN_MAX + 1);
  const values = sheet.getRange(start, 1, last - start + 1, 5).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const at = chatTimeIso_(values[i][0]);
    const from = chatNorm_(values[i][1]);
    const to = chatNorm_(values[i][2]);
    const fromName = String(values[i][3] || "").trim();
    const text = String(values[i][4] || "");
    if (!at || !from || !to || !text) continue;
    rows.push({ at: at, from: from, to: to, fromName: fromName, text: text });
  }
  return rows;
}

function listChatPeople(token) {
  const actor = chatActor_(token);
  if (!actor) return chatDeny_();
  const me = chatNorm_(actor.username);
  const latest = {};
  chatRows_().forEach(function (row) {
    let peer = "";
    if (row.from === me) peer = row.to;
    else if (row.to === me) peer = row.from;
    else return;
    const prev = latest[peer];
    if (!prev || prev.at < row.at) latest[peer] = row;
  });
  const people = chatTeam_().filter(function (person) {
    return chatNorm_(person.username) !== me;
  }).map(function (person) {
    const hit = latest[chatNorm_(person.username)];
    return {
      username: person.username,
      name: person.name,
      role: person.role,
      lastText: hit ? String(hit.text).slice(0, 80) : "",
      lastAt: hit ? hit.at : "",
      lastFrom: hit ? hit.from : ""
    };
  });
  people.sort(function (a, b) {
    if (a.lastAt && b.lastAt && a.lastAt !== b.lastAt) return a.lastAt < b.lastAt ? 1 : -1;
    if (a.lastAt && !b.lastAt) return -1;
    if (!a.lastAt && b.lastAt) return 1;
    return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
  });
  return { success: true, people: people };
}

function listChatThread(token, withUser) {
  const actor = chatActor_(token);
  if (!actor) return chatDeny_();
  const me = chatNorm_(actor.username);
  const peer = chatNorm_(withUser);
  if (!peer || peer === me) return { success: false, message: "Choose a person on the team." };
  const known = chatTeam_().some(function (person) {
    return chatNorm_(person.username) === peer;
  });
  if (!known) return { success: false, message: "That person is not on the team." };
  const messages = chatRows_().filter(function (row) {
    return (row.from === me && row.to === peer) || (row.from === peer && row.to === me);
  }).map(function (row) {
    return {
      at: row.at,
      fromName: row.from === me ? actor.name : (row.fromName || peer),
      mine: row.from === me,
      text: row.text
    };
  });
  return { success: true, messages: messages };
}

function sendChatMessage(token, withUser, text) {
  const actor = chatActor_(token);
  if (!actor) return chatDeny_();
  const me = chatNorm_(actor.username);
  const peer = chatNorm_(withUser);
  const body = String(text || "").replace(/\u0000/g, "").trim();
  if (!peer || peer === me) return { success: false, message: "Choose a person on the team." };
  if (!body) return { success: false, message: "Write a message first." };
  if (body.length > CHAT_TEXT_MAX) {
    return { success: false, message: "Message is too long. Keep it under 500 characters." };
  }
  const person = chatTeam_().filter(function (row) {
    return chatNorm_(row.username) === peer;
  })[0];
  if (!person) return { success: false, message: "That person is not on the team." };
  const sheet = chatSheet_();
  const at = new Date().toISOString();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 5).setNumberFormat("@");
    sheet.getRange(row, 1, 1, 5).setValues([[at, me, peer, actor.name, body]]);
  } finally {
    lock.releaseLock();
  }
  return { success: true, at: at };
}
