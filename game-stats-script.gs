// ═══════════════════════════════════════════════════════════
//  Water Polo Stats — Game Stats Apps Script
//  Paste into Extensions → Apps Script in your Game Stats
//  Google Sheet. Deploy as Web App:
//    Execute as: Me  |  Who has access: Anyone
//
//  Sheet structure (auto-created):
//  "Events" tab — one row per player action
//  "Games"  tab — one row per game (metadata)
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    switch (data.action) {
      case 'init':   initGame(ss, data);       break;
      case 'events': appendEvents(ss, data.rows); break;
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

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Game Stats connected' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────────────────────────────────────
//  Game init — writes one row to the Games tab
// ───────────────────────────────────────────────────────────

function initGame(ss, data) {
  const sheet = getOrCreateSheet(ss, 'Games', [
    'GameID', 'Date', 'DarkTeam', 'WhiteTeam', 'Created'
  ]);

  // Avoid duplicate init rows for the same game
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][0]) === String(data.gameId)) return;
  }

  sheet.appendRow([
    data.gameId || '',
    data.date   || '',
    data.darkTeam  || '',
    data.whiteTeam || '',
    new Date().toISOString(),
  ]);
}

// ───────────────────────────────────────────────────────────
//  Append event rows — one row per player action
// ───────────────────────────────────────────────────────────

function appendEvents(ss, rows) {
  if (!rows || !rows.length) return;

  const sheet = getOrCreateSheet(ss, 'Events', [
    'GameID', 'GroupID', 'Q', 'Time', 'Date',
    'Event', 'Team', 'Cap', 'PlayerID', 'Zone', 'GM'
  ]);

  const toAppend = rows.map(r => [
    r.gameId   || '',
    r.groupId  || '',
    r.q        || '',
    r.time     || '',
    r.date     || '',
    r.event    || '',
    r.team     || '',
    r.cap      || '',
    r.playerId || '',
    r.zone     || '',
    r.gm       || '',
  ]);

  // Batch append all rows at once (faster than one at a time)
  if (toAppend.length === 1) {
    sheet.appendRow(toAppend[0]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length)
         .setValues(toAppend);
  }
}

// ───────────────────────────────────────────────────────────
//  Sheet helper
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
