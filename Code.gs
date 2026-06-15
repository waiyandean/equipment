// ═══════════════════════════════════════════════════════════════════════════
//  Knife Registry — Google Apps Script
//  Covers: knife-check.html (daily checks) + knife-setup.html (register)
//
//  HOW TO DEPLOY
//  1. Open your Google Spreadsheet
//  2. Extensions → Apps Script → paste this entire file → Save
//  3. Deploy → New deployment → Web app
//     • Execute as: Me
//     • Who has access: Anyone
//  4. Copy the Web App URL into Settings on both knife pages
//
//  SHEETS CREATED AUTOMATICALLY
//  • "Knife Register"      — one row per knife, updated whenever setup changes
//  • "Knife Daily Checks"  — one row per day, opening + closing side by side
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_CHECKS   = 'Knife Daily Checks';
const SHEET_REGISTER = 'Knife Register';

// ── Entry points ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'mergeKnifeCheck':   return json(mergeKnifeCheck(body.check, body.knives));
      case 'saveKnifeRegister': return json(saveKnifeRegister(body.knives));
      default: return json({ result: 'error', message: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    return json({ result: 'error', message: err.toString() });
  }
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'getKnifeRegister')    return json(readKnifeRegister());
  if (action === 'getTodayCheckStatus') {
    const date = (e.parameter.date) ||
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return json(getTodayCheckStatus(date));
  }
  return json({ result: 'ok', message: 'Knife Registry Apps Script is running' });
}

// ═══════════════════════════════════════════════════════════════════════════
//  TODAY CHECK STATUS  (getTodayCheckStatus)
//
//  Returns whether opening and/or closing checks have been submitted to the
//  sheet for a given date. Used by the check page to show live status
//  instead of relying on device-local records.
//
//  Response shape:
//  {
//    result: 'success',
//    opening: { done:true, time:'09:30', staff:'Alice', allPresent:true, allGoodCondition:true } | null,
//    closing: { done:true, time:'17:45', staff:'Bob',   allPresent:true, allGoodCondition:false } | null
//  }
// ═══════════════════════════════════════════════════════════════════════════

function getTodayCheckStatus(dateStr) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CHECKS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { result: 'success', opening: null, closing: null };
  }

  const headerCount = sheet.getLastColumn();
  const headers     = sheet.getRange(1, 1, 1, headerCount).getValues()[0];

  // Find the row for this date
  const lastRow = sheet.getLastRow();
  const dateVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowNum = -1;
  for (let i = 0; i < dateVals.length; i++) {
    const cell = dateVals[i][0];
    const str  = cell instanceof Date
      ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(cell).trim();
    if (str === dateStr) { rowNum = i + 2; break; }
  }

  if (rowNum < 0) return { result: 'success', opening: null, closing: null };

  const row = sheet.getRange(rowNum, 1, 1, headerCount).getValues()[0];
  function val(name) {
    const idx = headers.indexOf(name);
    return idx >= 0 ? String(row[idx] || '') : '';
  }

  const openTime  = val('Opening Time');
  const closeTime = val('Closing Time');

  return {
    result: 'success',
    opening: openTime ? {
      done:             true,
      time:             openTime,
      staff:            val('Opening Staff'),
      allPresent:       val('Opening All Present').startsWith('✓'),
      allGoodCondition: val('Opening All Good Condition').startsWith('✓')
    } : null,
    closing: closeTime ? {
      done:             true,
      time:             closeTime,
      staff:            val('Closing Staff'),
      allPresent:       val('Closing All Present').startsWith('✓'),
      allGoodCondition: val('Closing All Good Condition').startsWith('✓')
    } : null
  };
}

// ═══════════════════════════════════════════════════════════════════════════
//  DAILY CHECKS  (mergeKnifeCheck)
//
//  One row per calendar day. The same row is used for both opening and
//  closing checks — the first check creates the row, the second fills it in.
//
//  Column layout:
//    Date | Site |
//    Opening Time | Opening Staff |
//      Opening K001 | Opening K001 Note | Opening K002 | ...  (one pair per knife)
//    Opening All Present | Opening All Good Condition |
//    Opening Production Held | Opening Notes |
//    Closing Time | Closing Staff |
//      Closing K001 | Closing K001 Note | ...
//    Closing All Present | Closing All Good Condition |
//    Closing Production Held | Closing Notes |
//    Day Complete
//
//  New knives are added as new columns automatically — existing data is never
//  deleted.
// ═══════════════════════════════════════════════════════════════════════════

function mergeKnifeCheck(check, activeKnives) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = getOrCreateSheet(ss, SHEET_CHECKS);
  const ids    = (activeKnives || []).map(k => k.id);

  ensureCheckHeaders(sheet, ids);

  const rowNum = findOrCreateDateRow(sheet, check.date, check.site);
  writeCheckData(sheet, rowNum, check);
  updateDayComplete(sheet, rowNum);

  return { result: 'success' };
}

// ── Header management ─────────────────────────────────────────────────────

function buildExpectedHeaders(knifeIds) {
  const h = ['Date', 'Site'];
  h.push('Opening Time', 'Opening Staff');
  knifeIds.forEach(id => h.push('Opening ' + id, 'Opening ' + id + ' Note'));
  h.push('Opening All Present', 'Opening All Good Condition',
         'Opening Production Held', 'Opening Notes');
  h.push('Closing Time', 'Closing Staff');
  knifeIds.forEach(id => h.push('Closing ' + id, 'Closing ' + id + ' Note'));
  h.push('Closing All Present', 'Closing All Good Condition',
         'Closing Production Held', 'Closing Notes');
  h.push('Day Complete');
  return h;
}

function ensureCheckHeaders(sheet, knifeIds) {
  const expected = buildExpectedHeaders(knifeIds);
  const lastCol  = sheet.getLastColumn();
  const lastRow  = sheet.getLastRow();

  if (lastCol === 0 || lastRow === 0) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    styleHeaderRange(sheet.getRange(1, 1, 1, expected.length));
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    return;
  }

  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                        .map(v => String(v));
  const missing  = expected.filter(h => !existing.includes(h));
  if (missing.length > 0) {
    const startCol = lastCol + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
    styleHeaderRange(sheet.getRange(1, startCol, 1, missing.length));
  }
}

function styleHeaderRange(range) {
  range.setBackground('#1a1916')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setFontSize(10)
       .setVerticalAlignment('middle');
}

// ── Row management ────────────────────────────────────────────────────────

function findOrCreateDateRow(sheet, dateStr, site) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dateVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < dateVals.length; i++) {
      const cell = dateVals[i][0];
      const str  = cell instanceof Date
        ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(cell).trim();
      if (str === dateStr) return i + 2;
    }
  }
  const newRow  = lastRow + 1;
  const dateCell = sheet.getRange(newRow, 1);
  dateCell.setValue(new Date(dateStr + 'T00:00:00'));
  dateCell.setNumberFormat('dd/MM/yyyy');
  sheet.getRange(newRow, 2).setValue(site || '');
  return newRow;
}

// ── Writing check data ────────────────────────────────────────────────────

function writeCheckData(sheet, rowNum, check) {
  const headerCount = sheet.getLastColumn();
  const headers     = sheet.getRange(1, 1, 1, headerCount).getValues()[0];

  function col(name) {
    const idx = headers.indexOf(name);
    return idx >= 0 ? idx + 1 : null;
  }
  function write(colName, value, bgColour, fgColour) {
    const c = col(colName);
    if (c == null) return;
    const cell = sheet.getRange(rowNum, c);
    cell.setValue(value);
    if (bgColour) cell.setBackground(bgColour);
    if (fgColour) cell.setFontColor(fgColour);
  }

  const prefix = check.checkType === 'opening' ? 'Opening' : 'Closing';

  write(prefix + ' Time',  check.time  || '');
  write(prefix + ' Staff', check.staff || '');

  (check.entries || []).forEach(entry => {
    const label          = knifeStatusLabel(entry.status, entry.productionHeld);
    const [bg, fg]       = knifeStatusColours(entry.status);
    write(prefix + ' ' + entry.knifeId,               label, bg, fg);
    write(prefix + ' ' + entry.knifeId + ' Note', entry.note || '');
  });

  const presentBg = check.allPresent ? '#edf7ed' : '#fdecea';
  const presentFg = check.allPresent ? '#2d6e2d' : '#c0392b';
  write(prefix + ' All Present',
        check.allPresent ? '✓ Yes' : '✗ No', presentBg, presentFg);

  if (check.allGoodCondition !== undefined) {
    const goodBg = check.allGoodCondition ? '#edf7ed' : '#fdf6e3';
    const goodFg = check.allGoodCondition ? '#2d6e2d' : '#8a6500';
    write(prefix + ' All Good Condition',
          check.allGoodCondition ? '✓ Yes' : '⚠ Issues', goodBg, goodFg);
  }

  if (check.productionHeld) {
    write(prefix + ' Production Held', '⚠ Yes', '#fdf6e3', '#8a6500');
  } else {
    write(prefix + ' Production Held', 'No');
  }

  write(prefix + ' Notes', check.notes || '');
}

function knifeStatusLabel(status, productionHeld) {
  switch (status) {
    case 'present-good': return '✓ Present & Good';
    case 'damaged':      return '⚠ Present – Damaged' + (productionHeld ? ' — HELD' : '');
    case 'missing':      return '✗ Missing'           + (productionHeld ? ' — PRODUCTION HELD' : '');
    case 'present':      return '✓ Present';
    default:             return status || '';
  }
}

function knifeStatusColours(status) {
  switch (status) {
    case 'present-good': return ['#edf7ed', '#2d6e2d'];
    case 'damaged':      return ['#fdf6e3', '#8a6500'];
    case 'missing':      return ['#fdecea', '#c0392b'];
    case 'present':      return ['#edf7ed', '#2d6e2d'];
    default:             return [null, null];
  }
}

// ── Day Complete column ───────────────────────────────────────────────────

function updateDayComplete(sheet, rowNum) {
  const headerCount = sheet.getLastColumn();
  const headers     = sheet.getRange(1, 1, 1, headerCount).getValues()[0];
  const doneCol     = headers.indexOf('Day Complete') + 1;
  if (!doneCol) return;

  const openCol  = headers.indexOf('Opening Time') + 1;
  const closeCol = headers.indexOf('Closing Time') + 1;
  if (!openCol || !closeCol) return;

  const row       = sheet.getRange(rowNum, 1, 1, headerCount).getValues()[0];
  const openDone  = row[openCol - 1]  !== '';
  const closeDone = row[closeCol - 1] !== '';
  const cell      = sheet.getRange(rowNum, doneCol);

  if (openDone && closeDone) {
    cell.setValue('✓ Complete')
        .setBackground('#edf7ed').setFontColor('#2d6e2d').setFontWeight('bold');
  } else if (openDone) {
    cell.setValue('Opening only')
        .setBackground('#fdf6e3').setFontColor('#8a6500').setFontWeight('normal');
  }
}


// ═══════════════════════════════════════════════════════════════════════════
//  KNIFE REGISTER  (saveKnifeRegister + readKnifeRegister)
// ═══════════════════════════════════════════════════════════════════════════

function readKnifeRegister() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_REGISTER);
  if (!sheet || sheet.getLastRow() <= 1) return { result: 'success', knives: [] };

  const rows   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const knives = rows
    .filter(r => r[0])
    .map(r => ({
      id:              String(r[0]),
      type:            String(r[1]),
      color:           String(r[2]),
      description:     String(r[3]),
      acquired:        r[4] instanceof Date
                         ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy-MM-dd')
                         : String(r[4] || ''),
      status:          String(r[5] || 'active'),
      withdrawnDate:   r[6] instanceof Date
                         ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
                         : '',
      withdrawnReason: String(r[7] || ''),
      withdrawnNotes:  String(r[8] || '')
    }));

  return { result: 'success', knives };
}

function saveKnifeRegister(knives) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = getOrCreateSheet(ss, SHEET_REGISTER);
  const headers = [
    'ID', 'Type', 'Colour', 'Description', 'Date Acquired',
    'Status', 'Withdrawn Date', 'Withdrawn Reason', 'Withdrawn Notes'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  styleHeaderRange(headerRange);
  sheet.setFrozenRows(1);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clear();
  }

  (knives || []).forEach((k, i) => {
    const rowNum = i + 2;
    const row = [
      k.id             || '',
      k.type           || '',
      k.color          || '',
      k.description    || '',
      k.acquired       ? new Date(k.acquired      + 'T00:00:00') : '',
      k.status         || 'active',
      k.withdrawnDate  ? new Date(k.withdrawnDate + 'T00:00:00') : '',
      k.withdrawnReason || '',
      k.withdrawnNotes  || ''
    ];
    sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);

    if (k.acquired)      sheet.getRange(rowNum, 5).setNumberFormat('dd/MM/yyyy');
    if (k.withdrawnDate) sheet.getRange(rowNum, 7).setNumberFormat('dd/MM/yyyy');

    if (k.status === 'withdrawn') {
      sheet.getRange(rowNum, 1, 1, headers.length)
           .setFontColor('#b0ada5')
           .setBackground('#fafaf8');
    } else {
      sheet.getRange(rowNum, 1, 1, headers.length)
           .setFontColor('#1a1916')
           .setBackground('#ffffff');
    }
  });

  sheet.autoResizeColumns(1, headers.length);
  return { result: 'success' };
}

// ═══════════════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════════════

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
