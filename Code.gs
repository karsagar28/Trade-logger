// Paste this entire file into your Google Apps Script editor.
// Deploy as a Web App: Execute as "Me", access "Anyone".

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'get';
  if (action === 'add') return addTrade(e.parameter);
  return getTrades();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Trades');
  if (!sh) {
    sh = ss.insertSheet('Trades');
    sh.appendRow(['Date', 'Ticker', 'Type', 'Risk', 'Stop Type', 'Result']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sh;
}

function addTrade(p) {
  try {
    const risk   = parseFloat(p.risk);
    const result = parseFloat(p.result);
    const date   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d yyyy, h:mm a');
    getSheet_().appendRow([
      date,
      p.ticker.toUpperCase(),
      p.type,
      risk + 'R',
      p.stopType,
      (result >= 0 ? '+' : '') + result + 'R'
    ]);
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
    const rows = sh.getRange(2, 1, last - 1, 6).getValues();
    return json_(rows.map(r => ({
      date:     r[0],
      ticker:   r[1],
      type:     r[2],
      risk:     r[3],
      stopType: r[4],
      result:   r[5]
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
