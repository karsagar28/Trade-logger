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
    sh.appendRow(['Date','Ticker','Type','Risk','Stop Type','Result','Close Type','Reason','Compliant','Analysis','Alternate Result','Strategy Type']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 12).setFontWeight('bold');
  } else {
    const lastCol = sh.getLastColumn();
    if (lastCol < 7)  sh.getRange(1, 7).setValue('Close Type').setFontWeight('bold');
    if (lastCol < 8)  sh.getRange(1, 8).setValue('Reason').setFontWeight('bold');
    if (lastCol < 9)  sh.getRange(1, 9).setValue('Compliant').setFontWeight('bold');
    if (lastCol < 10) sh.getRange(1, 10).setValue('Analysis').setFontWeight('bold');
    if (lastCol < 11) sh.getRange(1, 11).setValue('Alternate Result').setFontWeight('bold');
    if (lastCol < 12) sh.getRange(1, 12).setValue('Strategy Type').setFontWeight('bold');
    sh.getRange(1, 9).setValue('Compliant').setFontWeight('bold');
    sh.getRange(1, 10).setValue('Analysis').setFontWeight('bold');
    sh.getRange(1, 11).setValue('Alternate Result').setFontWeight('bold');
    sh.getRange(1, 12).setValue('Strategy Type').setFontWeight('bold');
  }
  return sh;
}

function addTrade(p) {
  try {
    const risk      = 1;
    const resultNum = (p.result !== undefined && p.result !== '') ? parseFloat(p.result) : null;
    const resultStr = resultNum !== null ? (resultNum >= 0 ? '+' : '') + resultNum + 'R' : '';
    const date      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMMM yy');

    getSheet_().appendRow([
      date, p.ticker.toUpperCase(), p.type, risk + 'R', p.stopType,
      resultStr, p.closeType || '', p.reason || '', normalizeCompliant_(p.compliant), p.analysis || '', formatR_(p.alternateResult), p.strategyType || ''
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

    const risk      = 1;
    const resultNum = (p.result !== undefined && p.result !== '') ? parseFloat(p.result) : null;
    const resultStr = resultNum !== null ? (resultNum >= 0 ? '+' : '') + resultNum + 'R' : '';

    sh.getRange(row, 2).setValue(p.ticker.toUpperCase());
    sh.getRange(row, 3).setValue(p.type);
    sh.getRange(row, 4).setValue(risk + 'R');
    sh.getRange(row, 5).setValue(p.stopType);
    sh.getRange(row, 6).setValue(resultStr);
    sh.getRange(row, 7).setValue(p.closeType || '');
    sh.getRange(row, 8).setValue(p.reason    || '');
    sh.getRange(row, 9).setValue(normalizeCompliant_(p.compliant));
    sh.getRange(row, 10).setValue(p.analysis || '');
    sh.getRange(row, 11).setValue(formatR_(p.alternateResult));
    sh.getRange(row, 12).setValue(p.strategyType || '');

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
    const rows = sh.getRange(2, 1, last - 1, 12).getValues();
    const tz = Session.getScriptTimeZone();
    return json_(rows.map((r, i) => ({
      row:       i + 2,
      date:      r[0] instanceof Date
                   ? Utilities.formatDate(r[0], tz, 'd MMMM yy')
                   : String(r[0] || ''),
      ticker:    r[1] || '',
      type:      r[2] || '',
      risk:      normalizeRisk_(r[3]),
      stopType:  r[4] || '',
      result:    r[5] || '',
      closeType: r[6] || '',
      reason:    r[7] || '',
      compliant: normalizeCompliant_(r[8]),
      analysis:  r[9] || '',
      alternateResult: r[10] || '',
      strategyType: r[11] || ''
    })));
  } catch (err) {
    return json_({ error: err.message });
  }
}

function normalizeCompliant_(value) {
  const v = String(value || '').toLowerCase();
  return (v === 'no' || v === 'false' || v === 'non-compliant' || v === 'noncompliant')
    ? 'No'
    : 'Yes';
}

function normalizeRisk_(value) {
  const n = parseFloat(String(value || '').replace(/R$/i, ''));
  return isNaN(n) ? '1R' : n + 'R';
}

function formatR_(value) {
  if (value === undefined || value === null || value === '') return '';
  const n = parseFloat(value);
  if (isNaN(n)) return '';
  return (n >= 0 ? '+' : '') + n + 'R';
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
