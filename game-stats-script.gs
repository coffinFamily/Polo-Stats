// ═══════════════════════════════════════════════════════════
//  Water Polo Stats — Game Stats Apps Script
//
//  Paste into Extensions → Apps Script in your Game Stats Google Sheet.
//  Deploy as Web App (Execute as: Me | Who has access: Anyone).
//
//  IMPORTANT: after ANY edit, publish it with
//    Deploy → Manage deployments → pencil → Version: "New version" → Deploy
//  Saving alone does not change the live /exec URL.
//
//  Sheet tabs: Games, Events, Tokens, Sessions
//  Games columns: A GameID, B Date, C DarkTeam, D WhiteTeam,
//                 E Created, F CreatorEmail, G Status
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    switch (data.action) {
      case 'init':         initGame(ss, data);              invalidateScoreboard_(); break;
      case 'events':       appendEvents(ss, data.rows);     break;
      case 'delete-event': deleteEventRows_(ss, data.groupId); break;
      case 'update-event': updateEvent(ss, data);           break;
      case 'updateGameStatus': setGameStatus(ss, data);     invalidateScoreboard_(); break;
    }

    return jsonOut({status: 'ok'});

  } catch (err) {
    return jsonOut({status: 'error', message: err.toString()});
  }
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'games') {
    const sheet = ss.getSheetByName('Games');
    if (!sheet) return jsonOut({games: []});
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return jsonOut({games: []});
    const games = rows.slice(1).map(r => ({
      gameId:       String(r[0] || ''),
      date:         String(r[1] || ''),
      darkTeam:     String(r[2] || ''),
      whiteTeam:    String(r[3] || ''),
      creatorEmail: String(r[5] || ''),
      status:       String(r[6] || '') || 'active',
    })).filter(g => g.gameId);
    return jsonOut({games});
  }

  // Scores tab: active games + goal counts in ONE call (cached ~20s).
  if (action === 'scoreboard') return jsonOut(handleScoreboard_(ss));

  if (action === 'events') {
    const gameId = (e.parameter && e.parameter.gameId) || '';
    const sheet  = ss.getSheetByName('Events');
    if (!sheet) return jsonOut({events: []});
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return jsonOut({events: []});
    const events = rows.slice(1)
      .filter(r => String(r[0]) === String(gameId))
      .map(r => ({
        groupId:  String(r[1] || ''),
        q:        String(r[2] || ''),
        time:     String(r[3] || ''),
        event:    String(r[5] || ''),
        team:     String(r[6] || ''),
        cap:      String(r[7] || ''),
        playerId: String(r[8] || ''),
        zone:     String(r[9] || ''),
      }));
    return jsonOut({events});
  }

  // ── Auth actions ────────────────────────────────────────
  if (action === 'requestLogin')    return jsonOut(handleRequestLogin_(ss, e.parameter.email));
  if (action === 'verifyToken')     return jsonOut(handleVerifyToken_(ss, e.parameter.token));
  if (action === 'validateSession') return jsonOut(handleValidateSession_(ss, e.parameter.session));
  // ───────────────────────────────────────────────────────
  if (action === 'updateCreators') return jsonOut(handleUpdateCreators_(ss, e.parameter.gameId, e.parameter.creators));
  return jsonOut({status: 'ok', message: 'Game Stats connected'});
}

// ───────────────────────────────────────────────────────────
//  Game init — writes one row to the Games tab
// ───────────────────────────────────────────────────────────

function initGame(ss, data) {
  const sheet = getOrCreateSheet(ss, 'Games', [
    'GameID', 'Date', 'DarkTeam', 'WhiteTeam', 'Created', 'CreatorEmail'
  ]);

  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][0]) === String(data.gameId)) return;
  }

  sheet.appendRow([
    data.gameId       || '',
    data.date         || '',
    data.darkTeam     || '',
    data.whiteTeam    || '',
    new Date().toISOString(),
    data.creatorEmail || '',
  ]);
}

// ───────────────────────────────────────────────────────────
//  Game status (archive / restore). Matches the GameID header
//  case-insensitively and adds a Status column if it's missing.
// ───────────────────────────────────────────────────────────

function setGameStatus(ss, data) {
  const sheet = ss.getSheetByName('Games');
  if (!sheet) return;
  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const idCol   = Math.max(headers.indexOf('gameid'), 0);
  let col = headers.indexOf('status');
  if (col < 0) {
    col = headers.length;
    sheet.getRange(1, col + 1).setValue('Status');
  }
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]).trim() === String(data.gameId).trim()) {
      sheet.getRange(i + 1, col + 1).setValue(data.status);
      return;
    }
  }
}

// ───────────────────────────────────────────────────────────
//  Scoreboard — one pass over Games + Events
//  Returns:
//    games:    active games (newest first) with darkScore / whiteScore
//    archived: inactive games, metadata only (so they can be restored)
//  A goal is any Events row whose Event starts with "Goal (" — the same
//  rule the app uses — credited to the row's Team.
//  Cached for 20 seconds so many viewers cost one sheet read.
// ───────────────────────────────────────────────────────────

const SCOREBOARD_CACHE_KEY = 'scoreboard_v1';
const SCOREBOARD_TTL_SEC   = 20;
const SCOREBOARD_MAX_GAMES = 150;

function invalidateScoreboard_() {
  try { CacheService.getScriptCache().remove(SCOREBOARD_CACHE_KEY); } catch (err) {}
}

function handleScoreboard_(ss) {
  const cache = CacheService.getScriptCache();
  try {
    const hit = cache.get(SCOREBOARD_CACHE_KEY);
    if (hit) return JSON.parse(hit);
  } catch (err) {}

  const tz     = ss.getSpreadsheetTimeZone();
  const gSheet = ss.getSheetByName('Games');
  if (!gSheet || gSheet.getLastRow() < 2) return {games: [], archived: []};
  const gRows = gSheet.getDataRange().getValues().slice(1);

  // Goals per game per team name (Events: A GameID … F Event, G Team)
  const counts = {};
  const eSheet = ss.getSheetByName('Events');
  if (eSheet && eSheet.getLastRow() >= 2) {
    const eRows = eSheet.getRange(2, 1, eSheet.getLastRow() - 1, 7).getValues();
    for (let i = 0; i < eRows.length; i++) {
      const r = eRows[i];
      if (!/^Goal \(/i.test(String(r[5] || ''))) continue;
      const gid  = String(r[0] || '');
      const team = String(r[6] || '').trim();
      if (!gid) continue;
      const g = counts[gid] || (counts[gid] = {});
      g[team] = (g[team] || 0) + 1;
    }
  }

  const active = [], archived = [];
  gRows.forEach(r => {
    const gameId = String(r[0] || '');
    if (!gameId) return;
    const dark  = String(r[2] || '').trim();
    const white = String(r[3] || '').trim();
    const rec = {
      gameId:       gameId,
      date:         sheetDate_(r[1], tz),
      darkTeam:     dark,
      whiteTeam:    white,
      creatorEmail: String(r[5] || ''),
      status:       String(r[6] || '').trim().toLowerCase() === 'inactive' ? 'inactive' : 'active',
      _d:           dateKey_(r[1]),
      _c:           Date.parse(String(r[4] || '')) || 0,
    };
    if (rec.status === 'inactive') { archived.push(rec); return; }
    const c = counts[gameId] || {};
    rec.darkScore  = c[dark]  || 0;
    rec.whiteScore = c[white] || 0;
    active.push(rec);
  });

  // Newest game date first; games on the same date: most recently created first
  const newest = (a, b) => (b._d - a._d) || (b._c - a._c);
  active.sort(newest);
  archived.sort(newest);
  active.concat(archived).forEach(g => { delete g._d; delete g._c; });

  const result = {
    games:    active.slice(0, SCOREBOARD_MAX_GAMES),
    archived: archived.slice(0, SCOREBOARD_MAX_GAMES),
  };
  try { cache.put(SCOREBOARD_CACHE_KEY, JSON.stringify(result), SCOREBOARD_TTL_SEC); } catch (err) {}
  return result;
}

// Date cell → 'yyyy-MM-dd' (cells formatted as dates arrive as Date objects)
function sheetDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v || '');
}

// Date cell → sortable number (ms since epoch), accepts Date, YYYY-MM-DD, M/D/YYYY
function dateKey_(v) {
  let t = 0;
  if (v instanceof Date) t = v.getTime();
  else {
    const s = String(v || '').trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) t = new Date(+m[1], +m[2] - 1, +m[3]).getTime();
    else if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) t = new Date(+m[3], +m[1] - 1, +m[2]).getTime();
    else t = Date.parse(s) || 0;
  }
  return t || 0;
}

// ───────────────────────────────────────────────────────────
//  Append / delete / update
// ───────────────────────────────────────────────────────────

function appendEvents(ss, rows) {
  if (!rows || !rows.length) return;
  const sheet = getOrCreateSheet(ss, 'Events', [
    'GameID', 'GroupID', 'Q', 'Time', 'Date',
    'Event', 'Team', 'Cap', 'PlayerID', 'Zone', 'GM'
  ]);
  const toAppend = rows.map(r => [
    r.gameId || '', r.groupId || '', r.q || '', r.time || '', r.date || '',
    r.event || '', r.team || '', r.cap || '', r.playerId || '', r.zone || '', r.gm || '',
  ]);
  if (toAppend.length === 1) {
    sheet.appendRow(toAppend[0]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length)
         .setValues(toAppend);
  }
}

function deleteEventRows_(ss, groupId) {
  if (!groupId) return;
  const sheet = ss.getSheetByName('Events');
  if (!sheet) return;
  const data   = sheet.getDataRange().getValues();
  const col    = data[0].findIndex(h => String(h).toLowerCase() === 'groupid');
  if (col < 0) return;
  const toDelete = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]) === String(groupId)) toDelete.push(i + 1);
  }
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
}

function updateEvent(ss, data) {
  deleteEventRows_(ss, data.groupId);
  appendEvents(ss, data.rows);
}

// ───────────────────────────────────────────────────────────
//  Helpers
// ───────────────────────────────────────────────────────────

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    const hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setFontWeight('bold');
    hdr.setBackground('#1a3a5a');
    hdr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════
//  Magic-link authentication
// ═══════════════════════════════════════════════════════════

// Where the app's HTML is publicly hosted (GitHub Pages)
const APP_URL = 'https://coffinfamily.github.io/Polo-Stats/';

function handleRequestLogin_(ss, email) {
  if (!email || !email.includes('@')) return {error: 'Invalid email'};
  const token = Utilities.getUuid();
  const now   = new Date();
  const exp   = new Date(now.getTime() + 15 * 60 * 1000); // 15 minutes
  const sh    = getOrCreateSheet(ss, 'Tokens', ['Token','Email','Created','Expires','Used']);
  sh.appendRow([token, email, now.toISOString(), exp.toISOString(), 'false']);
  const link  = APP_URL + '#ml=' + token;
  MailApp.sendEmail(
    email,
    'Your Polo Stats sign-in link',
    'Tap the link below to sign in (expires in 15 minutes):\n\n' + link +
    '\n\nIf you did not request this, you can ignore this email.'
  );
  return {ok: true};
}

function handleVerifyToken_(ss, token) {
  if (!token) return {error: 'No token'};
  const sh   = ss.getSheetByName('Tokens');
  if (!sh)   return {error: 'Token store missing'};
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== token) continue;
    if (rows[i][4] === 'true')             return {error: 'Token already used'};
    if (new Date() > new Date(rows[i][3])) return {error: 'Token expired'};
    sh.getRange(i + 1, 5).setValue('true'); // mark used
    const email      = rows[i][1];
    const sessionTok = Utilities.getUuid();
    const exp        = new Date(new Date().getTime() + 30 * 24 * 60 * 60 * 1000); // 30 days
    const ssh        = getOrCreateSheet(ss, 'Sessions', ['Token','Email','Created','Expires']);
    ssh.appendRow([sessionTok, email, new Date().toISOString(), exp.toISOString()]);
    return {ok: true, session: sessionTok, email: email, expires: exp.toISOString()};
  }
  return {error: 'Token not found'};
}

function handleValidateSession_(ss, session) {
  if (!session) return {error: 'No session'};
  const sh   = ss.getSheetByName('Sessions');
  if (!sh)   return {error: 'Session store missing'};
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== session) continue;
    if (new Date() > new Date(rows[i][3])) return {error: 'Session expired'};
    return {ok: true, email: rows[i][1]};
  }
  return {error: 'Session not found'};
}

function handleUpdateCreators_(ss, gameId, creators) {
  if (!gameId) return {error: 'No gameId'};
  const sh = ss.getSheetByName('Games');
  if (!sh) return {error: 'No Games sheet'};
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(gameId)) {
      sh.getRange(i + 1, 6).setValue(creators || '');
      return {ok: true};
    }
  }
  return {error: 'Game not found'};
}
