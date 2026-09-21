/**
 * Backend for the "TrackTime" calorie tracker (multi-account).
 * Deploy this script (bound to a Google Sheet) as a Web App:
 *   Deploy -> New deployment -> Web app
 *     Execute as: Me
 *     Who has access: Anyone
 * Copy the resulting /exec URL into the app's Profile -> Google Sheets sync field.
 *
 * Data for every account is tagged with its email and kept in the same
 * sheets, so one deployment/URL serves all accounts that use the app.
 */

var SHEET_ENTRIES = 'Entries';
var SHEET_TRAINING = 'Training';
var SHEET_PROFILE = 'Profile';
var SHEET_META = 'Meta';

function doGet(e) {
  var email = (e.parameter.email || '').toLowerCase();
  if (!email) return respond({ error: 'missing email' });
  return respond(readState(email));
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var email = (body.email || '').toLowerCase();
  if (!email) return respond({ error: 'missing email' });
  writeState(email, body);
  return respond({ ok: true });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function rowsFor(sheet, email) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var rows = values.slice(1).filter(function (row) { return String(row[0]).toLowerCase() === email; });
  return { headers: headers, rows: rows };
}

function readState(email) {
  var entriesSheet = getSheet(SHEET_ENTRIES, ['email', 'date', 'id', 'name', 'category', 'kcal']);
  var trainingSheet = getSheet(SHEET_TRAINING, ['email', 'id', 'type', 'quantity', 'unit', 'days', 'exercises', 'notes']);
  var profileSheet = getSheet(SHEET_PROFILE, ['email', 'key', 'value']);
  var metaSheet = getSheet(SHEET_META, ['email', 'key', 'value']);

  var history = {};
  rowsFor(entriesSheet, email).rows.forEach(function (row) {
    var date = row[1], id = row[2], name = row[3], category = row[4], kcal = row[5];
    if (!date) return;
    var key = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!history[key]) history[key] = { entries: [] };
    history[key].entries.push({ id: String(id), name: name, category: category, kcal: Number(kcal), photo: '' });
  });

  var training = rowsFor(trainingSheet, email).rows
    .filter(function (row) { return row[1]; })
    .map(function (row) {
      return {
        id: String(row[1]),
        type: row[2],
        quantity: Number(row[3]),
        unit: row[4],
        days: row[5] ? String(row[5]).split(',') : [],
        exercises: row[6],
        notes: row[7]
      };
    });

  var profile = {};
  rowsFor(profileSheet, email).rows.forEach(function (row) {
    if (row[1]) profile[row[1]] = row[2];
  });

  var meta = {};
  rowsFor(metaSheet, email).rows.forEach(function (row) {
    if (row[1]) meta[row[1]] = row[2];
  });

  return {
    goal: meta.goal ? Number(meta.goal) : 2000,
    language: meta.language || 'uk',
    theme: meta.theme || 'light',
    history: history,
    training: training,
    profile: profile
  };
}

function replaceRowsFor(sheet, email, headers, newRows) {
  var values = sheet.getDataRange().getValues();
  var kept = values.slice(1).filter(function (row) { return String(row[0]).toLowerCase() !== email; });
  sheet.clearContents();
  sheet.appendRow(headers);
  var all = kept.concat(newRows);
  if (all.length) sheet.getRange(2, 1, all.length, headers.length).setValues(all);
}

function writeState(email, state) {
  var entriesSheet = getSheet(SHEET_ENTRIES, ['email', 'date', 'id', 'name', 'category', 'kcal']);
  var entryRows = [];
  Object.keys(state.history || {}).forEach(function (date) {
    (state.history[date].entries || []).forEach(function (entry) {
      entryRows.push([email, date, entry.id, entry.name, entry.category, entry.kcal]);
    });
  });
  replaceRowsFor(entriesSheet, email, ['email', 'date', 'id', 'name', 'category', 'kcal'], entryRows);

  var trainingSheet = getSheet(SHEET_TRAINING, ['email', 'id', 'type', 'quantity', 'unit', 'days', 'exercises', 'notes']);
  var trainingRows = (state.training || []).map(function (tr) {
    return [email, tr.id, tr.type, tr.quantity, tr.unit, (tr.days || []).join(','), tr.exercises, tr.notes];
  });
  replaceRowsFor(trainingSheet, email, ['email', 'id', 'type', 'quantity', 'unit', 'days', 'exercises', 'notes'], trainingRows);

  var profileSheet = getSheet(SHEET_PROFILE, ['email', 'key', 'value']);
  var profile = state.profile || {};
  var profileRows = [
    [email, 'name', profile.name || ''],
    [email, 'age', profile.age || ''],
    [email, 'height', profile.height || ''],
    [email, 'weight', profile.weight || '']
  ];
  replaceRowsFor(profileSheet, email, ['email', 'key', 'value'], profileRows);

  var metaSheet = getSheet(SHEET_META, ['email', 'key', 'value']);
  var metaRows = [
    [email, 'goal', state.goal || 2000],
    [email, 'language', state.language || 'uk'],
    [email, 'theme', state.theme || 'light']
  ];
  replaceRowsFor(metaSheet, email, ['email', 'key', 'value'], metaRows);
}
