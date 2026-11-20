/**
 * Backend for index.html, bound to the Google Sheet with the training log.
 *
 * One TRAINING (e.g. "FBW A", "Push") = one sheet (tab). Every time the form is
 * saved, the exercises are appended to that tab as a new dated session, one row
 * per set, so the tab becomes a log from which progress is easy to read.
 *
 * GET  ?action=list  -> { ok, trainings: [{ name, lastDate, exercises }] }
 *                       exercises = the last session of the training (used to prefill the form):
 *                       [{ name, mode, note, sets: [{ reps, seconds, weight }] }]
 * POST (JSON as text/plain)
 *   { training, create, date: "YYYY-MM-DD", notes,
 *     exercises: [{ name, mode: "reps" | "time", note,
 *                   sets: [{ reps, seconds, weight }] }] }
 *   create=true makes a new tab, create=false appends to an existing one.
 *   notes = remarks about the whole session, note = remarks about one exercise.
 *   mode "reps" uses sets[].reps, mode "time" uses sets[].seconds;
 *   weight is optional: a number (kg), "BW" or empty.
 */

var HEADER = ['Data', 'Ćwiczenie', 'Seria', 'Powtórzenia', 'Czas (s)', 'Ciężar (kg)', 'Progres', 'Uwagi', 'Notatki z treningu'];
var COLOR_NAVY = '#0f2a4a';
var COLOR_UP = '#15803d';
var COLOR_DOWN = '#b91c1c';
var COLOR_SAME = '#6b7280';

function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (action === 'list') {
      return json_({ ok: true, trainings: listTrainings_() });
    }
    return json_({ ok: true, message: 'Trainings endpoint is running.' });
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
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 2) return false;
  var head = sheet.getRange(1, 1, 1, 2).getValues()[0];
  return head[0] === HEADER[0] && head[1] === HEADER[1];
}

function listTrainings_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = spreadsheet.getSpreadsheetTimeZone();
  return spreadsheet.getSheets().filter(isTrainingSheet_).map(function (sheet) {
    var last = lastSession_(readEntries_(sheet, timeZone));
    return { name: sheet.getName(), lastDate: last.date, exercises: last.exercises };
  });
}

/** One entry per row (= per set), in sheet order. */
function readEntries_(sheet, timeZone) {
  var rowCount = sheet.getLastRow() - 1;
  if (rowCount < 1) return [];
  var values = sheet.getRange(2, 1, rowCount, 8).getValues();
  var entries = [];
  values.forEach(function (row) {
    var name = String(row[1]).trim();
    if (!name) return;
    entries.push({
      date: formatDate_(row[0], timeZone),
      name: name,
      reps: Number(row[3]) || 0,
      seconds: Number(row[4]) || 0,
      weight: row[5] === '' ? null : (isBodyweight_(row[5]) ? 'BW' : Number(row[5])),
      note: String(row[7] || '')
    });
  });
  return entries;
}

function formatDate_(value, timeZone) {
  if (value instanceof Date) return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  return String(value);
}

function modeOf_(sets) {
  var allTimed = sets.length > 0 && sets.every(function (set) { return !(set.reps > 0) && set.seconds > 0; });
  return allTimed ? 'time' : 'reps';
}

/** Groups entries of one date into exercises (sets in sheet order). */
function groupExercises_(entries) {
  var order = [];
  var byKey = {};
  entries.forEach(function (entry) {
    var key = entry.name.toLowerCase();
    if (!byKey[key]) {
      byKey[key] = { name: entry.name, note: '', sets: [] };
      order.push(key);
    }
    var exercise = byKey[key];
    if (entry.note && !exercise.note) exercise.note = entry.note;
    exercise.sets.push({ reps: entry.reps, seconds: entry.seconds, weight: entry.weight });
  });
  return order.map(function (key) {
    var exercise = byKey[key];
    exercise.mode = modeOf_(exercise.sets);
    return exercise;
  });
}

/** Exercises of the most recent date in the log. */
function lastSession_(entries) {
  if (!entries.length) return { date: '', exercises: [] };
  var date = entries[entries.length - 1].date;
  var sameDay = entries.filter(function (entry) { return entry.date === date; });
  return { date: date, exercises: groupExercises_(sameDay) };
}

/** For every exercise name: its sets from the most recent date it was logged. */
function previousByName_(entries) {
  var byKey = {};
  entries.forEach(function (entry) {
    var key = entry.name.toLowerCase();
    if (!byKey[key] || byKey[key].date !== entry.date) {
      byKey[key] = { date: entry.date, name: entry.name, note: '', sets: [] };
    }
    byKey[key].sets.push({ reps: entry.reps, seconds: entry.seconds, weight: entry.weight });
  });
  Object.keys(byKey).forEach(function (key) { byKey[key].mode = modeOf_(byKey[key].sets); });
  return byKey;
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
  sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER])
    .setBackground(COLOR_NAVY).setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  var widths = [100, 300, 60, 110, 90, 110, 200, 280, 280];
  widths.forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.setTabColor(COLOR_NAVY);
  sheet.getRange(1, 1, sheet.getMaxRows(), HEADER.length).createFilter(); // filter by exercise to see its history
  return sheet;
}

function metrics_(exercise) {
  var timed = exercise.mode === 'time';
  var top = 0;
  var total = 0;
  exercise.sets.forEach(function (set) {
    top = Math.max(top, typeof set.weight === 'number' ? set.weight : 0); // "BW" and empty count as 0
    total += timed ? set.seconds : set.reps;
  });
  return { top: top, total: total, count: exercise.sets.length, mode: exercise.mode };
}

/** Compares an exercise with its previous entry; returns text and direction (1, 0, -1). */
function progress_(previous, current) {
  if (!previous) return { text: 'nowe ćwiczenie', direction: 0 };
  var before = metrics_(previous);
  var now = metrics_(current);
  var changes = [{ diff: now.top - before.top, unit: 'kg' }];
  if (before.mode === now.mode) {
    changes.push({ diff: now.total - before.total, unit: now.mode === 'time' ? 's' : 'powt.' });
  }
  changes.push({ diff: now.count - before.count, unit: 'serii' });

  var parts = [];
  var direction = 0;
  changes.forEach(function (change) {
    var diff = Math.round(change.diff * 100) / 100;
    if (diff === 0) return;
    if (direction === 0) direction = diff > 0 ? 1 : -1;
    parts.push((diff > 0 ? '+' : '') + diff + ' ' + change.unit);
  });
  return { text: parts.length ? parts.join(', ') : 'bez zmian', direction: direction };
}

function appendSession_(sheet, data, timeZone) {
  var previous = previousByName_(readEntries_(sheet, timeZone));
  var date = Utilities.parseDate(data.date, timeZone, 'yyyy-MM-dd');
  var startRow = sheet.getLastRow() + 1;
  var rows = [];
  var blocks = []; // first row, row count and progress of every exercise

  data.exercises.forEach(function (exercise) {
    var timed = exercise.mode === 'time';
    var current = {
      name: String(exercise.name).trim(),
      mode: timed ? 'time' : 'reps',
      sets: exercise.sets.map(function (set) {
        var hasWeight = set.weight !== null && set.weight !== undefined && set.weight !== '';
        return {
          reps: timed ? 0 : Number(set.reps),
          seconds: timed ? Number(set.seconds) : 0,
          weight: hasWeight ? (isBodyweight_(set.weight) ? 'BW' : Number(set.weight)) : null
        };
      })
    };
    var result = progress_(previous[current.name.toLowerCase()], current);
    blocks.push({ row: startRow + rows.length, count: current.sets.length, result: result });

    current.sets.forEach(function (set, index) {
      rows.push([
        date,
        current.name,
        index + 1,
        set.reps || '',
        set.seconds || '',
        set.weight === null ? '' : set.weight,
        index === 0 ? result.text : '',
        index === 0 ? String(exercise.note || '').trim() : '',
        rows.length === 0 ? String(data.notes || '') : ''
      ]);
    });
  });

  sheet.getRange(startRow, 1, rows.length, HEADER.length).setValues(rows);
  formatRows_(sheet, startRow, rows.length, blocks);
}

function formatRows_(sheet, startRow, rowCount, blocks) {
  sheet.getRange(startRow, 1, rowCount, 1).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  sheet.getRange(startRow, 3, rowCount, 4).setHorizontalAlignment('center');
  sheet.getRange(startRow, 6, rowCount, 1).setNumberFormat('0.0#');
  sheet.getRange(startRow, 8, rowCount, 2).setWrap(true);
  sheet.getRange(startRow, 2, rowCount, 1).setFontColor('#6b7280'); // later sets of an exercise are muted

  blocks.forEach(function (block) {
    sheet.getRange(block.row, 2).setFontColor('#000000').setFontWeight('bold');
    sheet.getRange(block.row, 1, 1, HEADER.length)
      .setBorder(true, null, null, null, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    var color = block.result.direction > 0 ? COLOR_UP : (block.result.direction < 0 ? COLOR_DOWN : COLOR_SAME);
    sheet.getRange(block.row, 7).setFontColor(color).setFontWeight(block.result.direction === 0 ? 'normal' : 'bold');
  });

  // A stronger line above the first exercise separates sessions (days).
  sheet.getRange(startRow, 1, 1, HEADER.length)
    .setBorder(true, null, null, null, null, null, '#64748b', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
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
