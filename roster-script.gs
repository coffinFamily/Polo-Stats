// ═══════════════════════════════════════════════════════════
//  Water Polo Stats — Master Roster Apps Script
//  Paste this into Extensions → Apps Script in your
//  Master Roster Google Sheet, then deploy as a Web App:
//    Execute as: Me
//    Who has access: Anyone
// ═══════════════════════════════════════════════════════════

// GET — called by any browser to load the shared roster
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'getRoster') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = getRosterData(ss);
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default: connection test
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Master Roster connected' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST — called by the app to write teams/players
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    switch (data.action) {
      case 'upsert-team':   upsertTeam(ss, data.team);          break;
      case 'delete-team':   deleteById(ss, 'Teams', data.id);   break;
      case 'upsert-player': upsertPlayer(ss, data.player);      break;
      case 'delete-player': deleteById(ss, 'Players', data.id); break;
      case 'ping':          break;
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ───────────────────────────────────────────────────────────
//  Read roster (used by doGet)
// ───────────────────────────────────────────────────────────

function getRosterData(ss) {
  const teamsSheet   = ss.getSheetByName('Teams');
  const playersSheet = ss.getSheetByName('Players');

  const teams = teamsSheet
    ? getRows(teamsSheet).map(r => ({ id: String(r[0]), name: r[1] }))
    : [];

  const players = playersSheet
    ? getRows(playersSheet).map(r => ({
        id:       String(r[0]),
        first:    r[1],
        last:     r[2],
        dob:      r[3] || '',
        teamIds:  r[4] ? String(r[4]).split(',').map(s => s.trim()).filter(Boolean) : [],
      }))
    : [];

  return { teams, players };
}

// Returns data rows (skips header, skips blank ID rows)
function getRows(sheet) {
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(row => row[0] !== '' && row[0] !== null);
}

// ───────────────────────────────────────────────────────────
//  Sheet helpers
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

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function deleteById(ss, sheetName, id) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const row = findRowById(sheet, id);
  if (row > 0) sheet.deleteRow(row);
}

// ───────────────────────────────────────────────────────────
//  Teams
// ───────────────────────────────────────────────────────────

function upsertTeam(ss, team) {
  const sheet = getOrCreateSheet(ss, 'Teams', ['ID', 'Name', 'Created', 'Updated']);
  const now   = new Date().toISOString();
  const row   = findRowById(sheet, team.id);

  if (row > 0) {
    const created = sheet.getRange(row, 3).getValue();
    sheet.getRange(row, 1, 1, 4).setValues([[team.id, team.name, created, now]]);
  } else {
    sheet.appendRow([team.id, team.name, now, now]);
  }
}

// ───────────────────────────────────────────────────────────
//  Players
// ───────────────────────────────────────────────────────────

function upsertPlayer(ss, player) {
  const sheet     = getOrCreateSheet(ss, 'Players', ['ID', 'First', 'Last', 'DOB', 'TeamIDs', 'TeamNames', 'Created', 'Updated']);
  const now       = new Date().toISOString();
  const teamIds   = (player.teamIds   || []).join(', ');
  const teamNames = (player.teamNames || []).join(', ');
  const row       = findRowById(sheet, player.id);

  if (row > 0) {
    const created = sheet.getRange(row, 7).getValue();
    sheet.getRange(row, 1, 1, 8).setValues([[
      player.id, player.first, player.last, player.dob || '',
      teamIds, teamNames, created, now
    ]]);
  } else {
    sheet.appendRow([player.id, player.first, player.last, player.dob || '', teamIds, teamNames, now, now]);
  }
}
