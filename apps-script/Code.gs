/**
 * Backend for index.html, bound to the Google Sheet with the training log.
 *
 * One TRAINING (e.g. "FBW A", "Push") = one sheet (tab). Layout of a tab:
 *
 *   Blok | Ćwiczenie | Seria |  2026-11-10   |  2026-11-24   | ...
 *        |           |       | Powt./czas | Kg | Powt./czas | Kg | ...
 *    A   | Wioślarz  |   1   |   120 s    |    |            |    |
 *    ...
 *
 * Exercises and sets are rows; every saved session adds two columns on the right
 * (reps or seconds, and kg). Reps are coloured against the previous date:
 * green = more, red = fewer, blue = the same. Remarks sit in cell notes:
 * the exercise remark on the first set's reps cell, the session remark on the date.
 *
 * GET  ?action=list  -> { ok, version, trainings: [{ name, lastDate, exercises }] }
 *                       exercises = the last session of the training (used to prefill the form):
 *                       [{ name, mode, note, group, sets: [{ reps, seconds, weight }] }]
 * POST (JSON as text/plain)
 *   { training, create, date: "YYYY-MM-DD", notes,
 *     exercises: [{ name, mode: "reps" | "time", note, group (optional block letter, e.g. "A"),
 *                   sets: [{ reps, seconds, weight }] }] }
 *   create=true makes a new tab, create=false adds a column pair to an existing one.
 *   mode "reps" uses sets[].reps, mode "time" uses sets[].seconds;
 *   weight is optional: a number (kg), "BW" or empty.
 */

// Bumped when the response format or sheet layout changes; index.html checks it to detect an outdated deployment.
var VERSION = 4;
var LEFT_HEADER = ['Blok', 'Ćwiczenie', 'Seria'];
var HEADER_ROWS = 2;
var FIRST_SESSION_COL = 4;
var TIME_FORMAT = '0" s"';
var REPS_FORMAT = '0';
var WEIGHT_FORMAT = '0.0#';
var COLOR_NAVY = '#0f2a4a';
var COLOR_SUBHEADER = '#e4e9f2';
// Pastel fills: block column and reps comparison with the previous date.
var GROUP_COLORS = { A: '#c6ebdc', B: '#fbd9b8', C: '#cbdcf9', D: '#efcff2', E: '#f8edb0' };
var COLOR_MORE = '#c8ecd3';
var COLOR_LESS = '#f7c9c5';
var COLOR_SAME = '#cddffa';

function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (action === 'list') {
      return json_({ ok: true, version: VERSION, trainings: listTrainings_() });
    }
    return json_({ ok: true, version: VERSION, message: 'Trainings endpoint is running.' });
  } catch (error) {
    return json_({ ok: false, error: message_(error) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    validate_(data);

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timeZone = spreadsheet.getSpreadsheetTimeZone();
    var sheet;

    if (data.create) {
      var name = cleanSheetName_(data.training);
      if (spreadsheet.getSheetByName(name)) {
        throw new Error('Zakładka "' + name + '" już istnieje.');
      }
      sheet = createTrainingSheet_(spreadsheet, name);
    } else {
      sheet = spreadsheet.getSheetByName(data.training);
      if (!sheet || !isTrainingSheet_(sheet)) {
        throw new Error('Nie znaleziono treningu "' + data.training + '".');
      }
    }

    appendSession_(sheet, data, timeZone);
    return json_({ ok: true, sheet: sheet.getName() });
  } catch (error) {
    return json_({ ok: false, error: message_(error) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- reading ----------

function isTrainingSheet_(sheet) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 3) return false;
  var head = sheet.getRange(1, 1, 1, 3).getValues()[0];
  return head[0] === LEFT_HEADER[0] && head[1] === LEFT_HEADER[1] && head[2] === LEFT_HEADER[2];
}

function listTrainings_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = spreadsheet.getSpreadsheetTimeZone();
  return spreadsheet.getSheets().filter(isTrainingSheet_).map(function (sheet) {
    var layout = readLayout_(sheet, timeZone);
    var last = layout.sessions.length ? layout.sessions[layout.sessions.length - 1] : null;
    return {
      name: sheet.getName(),
      lastDate: last ? last.date : '',
      exercises: last ? sessionExercises_(sheet, layout, last) : []
    };
  });
}

/** Rows (one per exercise set) and sessions (one per date column pair) of a training tab. */
function readLayout_(sheet, timeZone) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var rows = [];
  var sessions = [];

  if (lastRow > HEADER_ROWS) {
    var left = sheet.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, 3).getValues();
    left.forEach(function (row, index) {
      var name = String(row[1]).trim();
      if (!name) return;
      rows.push({
        row: HEADER_ROWS + 1 + index,
        group: String(row[0] || '').trim().toUpperCase(),
        name: name,
        set: Number(row[2]) || 1
      });
    });
  }

  if (lastCol >= FIRST_SESSION_COL) {
    var header = sheet.getRange(1, FIRST_SESSION_COL, 1, lastCol - FIRST_SESSION_COL + 1).getValues()[0];
    header.forEach(function (value, index) {
      if (value === '' || value === null) return;
      sessions.push({ col: FIRST_SESSION_COL + index, date: formatDate_(value, timeZone) });
    });
  }
  return { rows: rows, sessions: sessions };
}

function formatDate_(value, timeZone) {
  if (value instanceof Date) return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  return String(value);
}

/** Exercises (with their sets) that have results in the given session column pair. */
function sessionExercises_(sheet, layout, session) {
  if (!layout.rows.length) return [];
  var lastRow = layout.rows[layout.rows.length - 1].row;
  var count = lastRow - HEADER_ROWS;
  var values = sheet.getRange(HEADER_ROWS + 1, session.col, count, 2).getValues();
  var formats = sheet.getRange(HEADER_ROWS + 1, session.col, count, 1).getNumberFormats();
  var notes = sheet.getRange(HEADER_ROWS + 1, session.col, count, 1).getNotes();

  var order = [];
  var byKey = {};
  layout.rows.forEach(function (entry) {
    var index = entry.row - HEADER_ROWS - 1;
    var main = values[index][0];
    var weight = values[index][1];
    if ((main === '' || main === null) && (weight === '' || weight === null)) return;

    var key = entry.name.toLowerCase();
    if (!byKey[key]) {
      byKey[key] = { name: entry.name, mode: 'reps', note: '', group: entry.group, sets: [] };
      order.push(key);
    }
    var exercise = byKey[key];
    var timed = String(formats[index][0]).indexOf(' s"') >= 0;
    if (timed) exercise.mode = 'time';
    if (!exercise.note && notes[index][0]) exercise.note = String(notes[index][0]);
    var amount = Number(main) || 0;
    exercise.sets.push({
      reps: timed ? 0 : amount,
      seconds: timed ? amount : 0,
      weight: weight === '' || weight === null ? null : (isBodyweight_(weight) ? 'BW' : Number(weight))
    });
  });
  return order.map(function (key) { return byKey[key]; });
}

// ---------- writing ----------

function validate_(data) {
  if (!data || typeof data !== 'object') throw new Error('Brak danych.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.date))) throw new Error('Nieprawidłowa data.');
  if (!data.training || !String(data.training).trim()) throw new Error('Brak nazwy treningu.');
  if (!data.exercises || !data.exercises.length) throw new Error('Brak ćwiczeń.');
  if (data.exercises.length > 50) throw new Error('Za dużo ćwiczeń.');
  data.exercises.forEach(function (exercise) {
    var label = '"' + exercise.name + '"';
    if (!exercise.name || !String(exercise.name).trim()) throw new Error('Ćwiczenie bez nazwy.');
    if (!exercise.sets || !exercise.sets.length || exercise.sets.length > 30) {
      throw new Error('Nieprawidłowa liczba serii w ćwiczeniu ' + label + '.');
    }
    if (exercise.group && !/^[A-Za-z]$/.test(String(exercise.group))) {
      throw new Error('Nieprawidłowy blok w ćwiczeniu ' + label + '.');
    }
    var timed = exercise.mode === 'time';
    exercise.sets.forEach(function (set) {
      var value = Number(timed ? set.seconds : set.reps);
      var hasWeight = set.weight !== null && set.weight !== undefined && set.weight !== '';
      if (!(value > 0)) throw new Error('Brak ' + (timed ? 'czasu' : 'powtórzeń') + ' w ćwiczeniu ' + label + '.');
      if (hasWeight && !isBodyweight_(set.weight) && !(Number(set.weight) >= 0)) {
        throw new Error('Nieprawidłowy ciężar w ćwiczeniu ' + label + '.');
      }
    });
  });
}

function cleanSheetName_(name) {
  var cleaned = String(name).replace(/[\[\]*\/\\?:]/g, '-').trim().substring(0, 90);
  if (!cleaned) throw new Error('Nieprawidłowa nazwa treningu.');
  return cleaned;
}

function createTrainingSheet_(spreadsheet, name) {
  var sheet = spreadsheet.insertSheet(name);
  sheet.getRange(1, 1, 1, 3).setValues([LEFT_HEADER]);
  [1, 2, 3].forEach(function (col) { sheet.getRange(1, col, HEADER_ROWS, 1).merge(); });
  sheet.getRange(1, 1, HEADER_ROWS, 3)
    .setBackground(COLOR_NAVY).setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(HEADER_ROWS);
  sheet.setFrozenColumns(3);
  [60, 340, 60].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.setTabColor(COLOR_NAVY);
  return sheet;
}

/**
 * Returns the sheet rows (in set order) for an exercise, inserting rows when the exercise is
 * new or has fewer sets than needed. New exercises are placed under the last row of their block.
 */
function ensureRows_(sheet, timeZone, name, group, setCount) {
  var key = name.toLowerCase();
  var layout = readLayout_(sheet, timeZone);
  var own = layout.rows.filter(function (entry) { return entry.name.toLowerCase() === key; });
  var missing = setCount - own.length;

  if (missing > 0) {
    var after;
    if (own.length) {
      after = own[own.length - 1].row;
    } else {
      var sameBlock = group ? layout.rows.filter(function (entry) { return entry.group === group; }) : [];
      after = sameBlock.length ? sameBlock[sameBlock.length - 1].row
        : (layout.rows.length ? layout.rows[layout.rows.length - 1].row : HEADER_ROWS);
    }
    if (after >= sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), missing);
    else sheet.insertRowsAfter(after, missing);
    var fill = [];
    for (var i = 0; i < missing; i++) {
      fill.push([group || '', own.length ? own[0].name : name, own.length + i + 1]);
    }
    sheet.getRange(after + 1, 1, missing, 3).setValues(fill);
    layout = readLayout_(sheet, timeZone);
    own = layout.rows.filter(function (entry) { return entry.name.toLowerCase() === key; });
  }

  // A block chosen on the page also updates the block of an existing exercise.
  if (group) {
    own.forEach(function (entry) {
      if (entry.group !== group) sheet.getRange(entry.row, 1).setValue(group);
    });
  }
  return own.slice(0, setCount).map(function (entry) { return entry.row; });
}

function appendSession_(sheet, data, timeZone) {
  var date = Utilities.parseDate(data.date, timeZone, 'yyyy-MM-dd');

  // 1. Make sure every exercise has rows for all of its sets.
  var plan = data.exercises.map(function (exercise) {
    var timed = exercise.mode === 'time';
    var name = String(exercise.name).trim();
    var group = String(exercise.group || '').trim().toUpperCase();
    var rows = ensureRows_(sheet, timeZone, name, group, exercise.sets.length);
    return { exercise: exercise, timed: timed, name: name, rows: rows };
  });

  // 2. The new column pair goes to the right of the last session.
  var layout = readLayout_(sheet, timeZone);
  var col = Math.max(sheet.getLastColumn(), FIRST_SESSION_COL - 1) + 1;
  var lastRow = Math.max(sheet.getLastRow(), HEADER_ROWS + 1);
  var count = lastRow - HEADER_ROWS;

  var previous = col > FIRST_SESSION_COL
    ? sheet.getRange(HEADER_ROWS + 1, FIRST_SESSION_COL, count, col - FIRST_SESSION_COL).getValues()
    : [];
  var sessionCols = layout.sessions.map(function (session) { return session.col - FIRST_SESSION_COL; });

  var values = [];
  var formats = [];
  var backgrounds = [];
  var notes = [];
  for (var i = 0; i < count; i++) {
    values.push(['', '']);
    formats.push([REPS_FORMAT, WEIGHT_FORMAT]);
    backgrounds.push([null, null]);
    notes.push(['', '']);
  }

  plan.forEach(function (item) {
    item.exercise.sets.forEach(function (set, index) {
      var rowIndex = item.rows[index] - HEADER_ROWS - 1;
      var amount = Number(item.timed ? set.seconds : set.reps);
      var hasWeight = set.weight !== null && set.weight !== undefined && set.weight !== '';
      values[rowIndex] = [amount, hasWeight ? (isBodyweight_(set.weight) ? 'BW' : Number(set.weight)) : ''];
      formats[rowIndex][0] = item.timed ? TIME_FORMAT : REPS_FORMAT;
      backgrounds[rowIndex][0] = compareColor_(previous[rowIndex], sessionCols, amount);
      if (index === 0 && item.exercise.note && String(item.exercise.note).trim()) {
        notes[rowIndex][0] = String(item.exercise.note).trim();
      }
    });
  });

  var body = sheet.getRange(HEADER_ROWS + 1, col, count, 2);
  body.setValues(values);
  body.setNumberFormats(formats);
  body.setBackgrounds(backgrounds);
  sheet.getRange(HEADER_ROWS + 1, col, count, 1).setNotes(notes.map(function (n) { return [n[0]]; }));
  body.setHorizontalAlignment('center');

  // 3. Header of the new date.
  var head = sheet.getRange(1, col, 1, 2);
  head.merge();
  head.setNumberFormat('yyyy-mm-dd')
    .setBackground(COLOR_NAVY).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(1, col).setValue(date);
  if (data.notes && String(data.notes).trim()) sheet.getRange(1, col).setNote(String(data.notes).trim());
  sheet.getRange(2, col, 1, 2).setValues([['Powt./czas', 'Kg']])
    .setBackground(COLOR_SUBHEADER).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setColumnWidth(col, 90);
  sheet.setColumnWidth(col + 1, 70);

  formatLeft_(sheet, timeZone);
}

/** Colour of a reps cell against the latest earlier value in the same row (null = nothing to compare). */
function compareColor_(previousRow, sessionCols, amount) {
  if (!previousRow) return null;
  for (var i = sessionCols.length - 1; i >= 0; i--) {
    var before = previousRow[sessionCols[i]];
    if (before !== '' && before !== null && !isNaN(Number(before))) {
      before = Number(before);
      return amount > before ? COLOR_MORE : (amount < before ? COLOR_LESS : COLOR_SAME);
    }
  }
  return null;
}

/** Block colours, exercise names and separators in the three left columns. */
function formatLeft_(sheet, timeZone) {
  var layout = readLayout_(sheet, timeZone);
  var lastCol = sheet.getLastColumn();
  var previousName = null;
  layout.rows.forEach(function (entry) {
    var key = entry.name.toLowerCase();
    var first = key !== previousName;
    previousName = key;
    var block = sheet.getRange(entry.row, 1);
    block.setBackground(GROUP_COLORS[entry.group] || null).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(entry.row, 2).setFontColor(first ? '#000000' : '#9ca3af').setFontWeight(first ? 'bold' : 'normal');
    sheet.getRange(entry.row, 3).setHorizontalAlignment('center');
    if (first) {
      sheet.getRange(entry.row, 1, 1, lastCol)
        .setBorder(true, null, null, null, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    }
  });
}

// ---------- helpers ----------

function isBodyweight_(value) {
  return String(value).trim().toUpperCase() === 'BW';
}

function json_(object) {
  return ContentService.createTextOutput(JSON.stringify(object)).setMimeType(ContentService.MimeType.JSON);
}

function message_(error) {
  return String(error && error.message ? error.message : error);
}
