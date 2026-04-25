// Paste this entire file into your Google Apps Script editor.
// Deploy as a Web App: Execute as "Me", access "Anyone".

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'get';
  if (action === 'add')    return addTrade(e.parameter);
  if (action === 'update') return updateTrade(e.parameter);
  if (action === 'delete') return deleteTrade(e.parameter);
  return getTrades();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Trades');
  if (!sh) {
    sh = ss.insertSheet('Trades');
    sh.appendRow(['Date','Ticker','Type','Risk','Stop Type','Result','Close Type','Reason','Mood','Notes']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold');
  } else {
    const lastCol = sh.getLastColumn();
    if (lastCol < 7)  sh.getRange(1, 7).setValue('Close Type').setFontWeight('bold');
    if (lastCol < 8)  sh.getRange(1, 8).setValue('Reason').setFontWeight('bold');
    if (lastCol < 9)  sh.getRange(1, 9).setValue('Mood').setFontWeight('bold');
    if (lastCol < 10) sh.getRange(1, 10).setValue('Notes').setFontWeight('bold');
  }
  return sh;
}

function addTrade(p) {
  try {
    const risk      = parseFloat(p.risk);
    const resultNum = (p.result !== undefined && p.result !== '') ? parseFloat(p.result) : null;
    const resultStr = resultNum !== null ? (resultNum >= 0 ? '+' : '') + resultNum + 'R' : '';
    const date      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMMM yy');

    getSheet_().appendRow([
      date, p.ticker.toUpperCase(), p.type, risk + 'R', p.stopType,
      resultStr, p.closeType || '', p.reason || '', p.mood || '', p.notes || ''
    ]);
    return json_({ success: true });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function updateTrade(p) {
  try {
    const sh  = getSheet_();
    const row = parseInt(p.row);
    if (isNaN(row) || row < 2) return json_({ error: 'Invalid row' });

    const risk      = parseFloat(p.risk);
    const resultNum = (p.result !== undefined && p.result !== '') ? parseFloat(p.result) : null;
    const resultStr = resultNum !== null ? (resultNum >= 0 ? '+' : '') + resultNum + 'R' : '';

    sh.getRange(row, 2).setValue(p.ticker.toUpperCase());
    sh.getRange(row, 3).setValue(p.type);
    sh.getRange(row, 4).setValue(risk + 'R');
    sh.getRange(row, 5).setValue(p.stopType);
    sh.getRange(row, 6).setValue(resultStr);
    sh.getRange(row, 7).setValue(p.closeType || '');
    sh.getRange(row, 8).setValue(p.reason    || '');
    sh.getRange(row, 9).setValue(p.mood      || '');
    sh.getRange(row, 10).setValue(p.notes    || '');

    return json_({ success: true });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function deleteTrade(p) {
  try {
    const sh  = getSheet_();
    const row = parseInt(p.row);
    if (isNaN(row) || row < 2) return json_({ error: 'Invalid row' });
    sh.deleteRow(row);
    return json_({ success: true });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function getTrades() {
  try {
    const sh   = getSheet_();
    const last = sh.getLastRow();
    if (last < 2) return json_([]);
    const rows = sh.getRange(2, 1, last - 1, 10).getValues();
    const tz = Session.getScriptTimeZone();
    return json_(rows.map((r, i) => ({
      row:       i + 2,
      date:      r[0] instanceof Date
                   ? Utilities.formatDate(r[0], tz, 'd MMMM yy')
                   : String(r[0] || ''),
      ticker:    r[1] || '',
      type:      r[2] || '',
      risk:      r[3] || '',
      stopType:  r[4] || '',
      result:    r[5] || '',
      closeType: r[6] || '',
      reason:    r[7] || '',
      mood:      r[8] || '',
      notes:     r[9] || ''
    })));
  } catch (err) {
    return json_({ error: err.message });
  }
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
