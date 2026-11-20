/**
 * Backend for index.html, bound to the Google Sheet with the training log.
 *
 * One TRAINING (e.g. "FBW A", "Push") = one sheet (tab). Layout of a tab:
 *
 *   Blok | Ćwiczenie | Seria |  2026-11-10   |  2026-11-24   | ...
 *        |           |       | Powt./czas | Kg | Powt./czas | Kg | ...
 *   A0   | Wioślarz  |   1   |   120 s    |    |            |    |
 *    ...
 *
 * "Blok" holds the exercise code from the plan: block letter + section number + optional
 * variant letter, e.g. A1, C2a, C2b. Variants of one section (C2a / C2b) are alternatives:
 * the form shows one tile per section and lets you switch between its variants.
 *
 * How the day went (energy, sleep, food; 1-5) is kept in three extra rows at the bottom,
 * marked "Stan" in the Blok column, with a value in the reps column of the session's date.
 *
 * Exercises and sets are rows; every saved session adds two columns on the right
 * (reps or seconds, and kg). Reps are coloured against the previous date:
 * green = more, red = fewer, blue = the same. Remarks sit in cell notes:
 * the exercise remark on the first set's reps cell, the session remark (and links to photos
 * and videos) on the date.
 *
 * GET  ?action=list  -> { ok, version, trainings: [{ name, lastDate, exercises }] }
 *                       exercises = every exercise of the training with its most recent results:
 *                       [{ name, code, group, mode, note, date, inLast, sets: [{ reps, seconds, weight }] }]
 *                       inLast = it was done in the training's last session.
 * POST (JSON as text/plain)
 *   save a session:
 *   { training, create, date: "YYYY-MM-DD", notes,
 *     wellness: { energy, sleep, food } (each 1-5, all optional),
 *     exercises: [{ name, mode: "reps" | "time", note, code (e.g. "A1", "C2b"; optional),
 *                   sets: [{ reps, seconds, weight }] }],
 *     emails: ["a@b.pl"], media: [{ id }], pin }
 *   create=true makes a new tab, create=false adds a column pair to an existing one.
 *   mode "reps" uses sets[].reps, mode "time" uses sets[].seconds;
 *   weight is optional: a number (kg), "BW" or empty.
 *   emails -> a summary of the session is mailed from the owner's account (needs pin).
 *   media  -> files already uploaded to the "Treningi – media" Drive folder (see below).
 *   upload a photo or video to Drive in chunks (all need pin):
 *   { action: "uploadStart", name, mime, size, pin }            -> { ok, uploadId }
 *   { action: "uploadChunk", uploadId, offset, data (base64), pin } -> { ok, done, id?, url? }
 *
 * Mail and Drive are only available when the script property PIN is set
 * (Project settings -> Script properties); the page sends it with every such request.
 * Run authorize() once from the editor to grant the Drive and Mail permissions.
 */

// Bumped when the response format or sheet layout changes; index.html checks it to detect an outdated deployment.
var VERSION = 6;
var LEFT_HEADER = ['Blok', 'Ćwiczenie', 'Seria'];
var HEADER_ROWS = 2;
var FIRST_SESSION_COL = 4;
var TIME_FORMAT = '0" s"';
var REPS_FORMAT = '0';
var WEIGHT_FORMAT = '0.0#';
var COLOR_NAVY = '#0f2a4a';
var COLOR_SUBHEADER = '#e4e9f2';
var MEDIA_FOLDER = 'Treningi – media';
var WELLNESS_BLOCK = 'Stan';
var WELLNESS = [
  { key: 'energy', label: 'Energia (1–5)', mail: 'energia' },
  { key: 'sleep', label: 'Sen (1–5)', mail: 'sen' },
  { key: 'food', label: 'Jedzenie (1–5)', mail: 'jedzenie' }
];
// 1 (worst) .. 5 (best)
var WELLNESS_COLORS = ['#f7c9c5', '#fbd9b8', '#f8edb0', '#d9efc8', '#c8ecd3'];
var MAX_EMAILS = 5;
var MAX_MEDIA = 20;
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
  try {
    var data = JSON.parse(e.postData.contents);
    if (data && data.action === 'uploadStart') return json_(uploadStart_(data));
    if (data && data.action === 'uploadChunk') return json_(uploadChunk_(data));
    return json_(saveSession_(data));
  } catch (error) {
    return json_({ ok: false, error: message_(error) });
  }
}

function saveSession_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    validate_(data);
    var emails = data.emails || [];
    var media = data.media || [];
    if (emails.length || media.length) requirePin_(data);

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

    var files = media.length ? mediaFiles_(media) : [];
    var before = emails.length ? previousByName_(sheet, timeZone) : {};
    appendSession_(sheet, data, timeZone, files);

    var result = { ok: true, sheet: sheet.getName(), media: files.length };
    if (emails.length) {
      try {
        shareFiles_(files, emails);
        sendSummary_(data, sheet.getName(), before, files, emails);
        result.emailed = emails.length;
      } catch (error) {
        result.emailError = message_(error);
      }
    }
    return result;
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
      exercises: last ? latestExercises_(sheet, layout) : []
    };
  });
}

/** Rows (one per exercise set) and sessions (one per date column pair) of a training tab. */
function readLayout_(sheet, timeZone) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var rows = [];
  var sessions = [];
  var wellness = {};

  if (lastRow > HEADER_ROWS) {
    var left = sheet.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, 3).getValues();
    left.forEach(function (row, index) {
      var name = String(row[1]).trim();
      if (!name) return;
      if (String(row[0]).trim() === WELLNESS_BLOCK) {
        WELLNESS.forEach(function (item) { if (item.label === name) wellness[item.key] = HEADER_ROWS + 1 + index; });
        return;
      }
      var code = normalizeCode_(row[0]);
      rows.push({
        row: HEADER_ROWS + 1 + index,
        code: code,
        group: code.charAt(0),
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
  return { rows: rows, sessions: sessions, wellness: wellness };
}

function formatDate_(value, timeZone) {
  if (value instanceof Date) return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  return String(value);
}

/**
 * Every exercise of the tab with its most recent results (the latest session in which it has data).
 * inLast tells whether that session is the tab's last one.
 */
function latestExercises_(sheet, layout) {
  if (!layout.rows.length || !layout.sessions.length) return [];
  var lastRow = layout.rows[layout.rows.length - 1].row;
  var count = lastRow - HEADER_ROWS;
  var firstCol = layout.sessions[0].col;
  var lastSession = layout.sessions[layout.sessions.length - 1];
  var width = lastSession.col + 2 - firstCol;
  var range = sheet.getRange(HEADER_ROWS + 1, firstCol, count, width);
  var values = range.getValues();
  var formats = range.getNumberFormats();
  var notes = range.getNotes();

  var order = [];
  var byKey = {};
  layout.rows.forEach(function (entry) {
    var key = entry.name.toLowerCase();
    if (!byKey[key]) { byKey[key] = { entry: entry, rows: [] }; order.push(key); }
    byKey[key].rows.push(entry);
  });

  var result = [];
  order.forEach(function (key) {
    var item = byKey[key];
    for (var s = layout.sessions.length - 1; s >= 0; s--) {
      var session = layout.sessions[s];
      var c = session.col - firstCol;
      var exercise = { name: item.entry.name, code: item.entry.code, group: item.entry.group, mode: 'reps', note: '', date: session.date, inLast: s === layout.sessions.length - 1, sets: [] };
      item.rows.forEach(function (entry) {
        var index = entry.row - HEADER_ROWS - 1;
        var main = values[index][c];
        var weight = values[index][c + 1];
        if ((main === '' || main === null) && (weight === '' || weight === null)) return;
        var timed = String(formats[index][c]).indexOf(' s"') >= 0;
        if (timed) exercise.mode = 'time';
        if (!exercise.note && notes[index][c]) exercise.note = String(notes[index][c]);
        var amount = Number(main) || 0;
        exercise.sets.push({
          reps: timed ? 0 : amount,
          seconds: timed ? amount : 0,
          weight: weight === '' || weight === null ? null : (isBodyweight_(weight) ? 'BW' : Number(weight))
        });
      });
      if (exercise.sets.length) { result.push(exercise); return; }
    }
  });
  return result;
}

/** Results from before the new session, by lower-case exercise name (for the e-mail comparison). */
function previousByName_(sheet, timeZone) {
  var map = {};
  latestExercises_(sheet, readLayout_(sheet, timeZone)).forEach(function (exercise) {
    map[exercise.name.toLowerCase()] = exercise;
  });
  return map;
}

// ---------- writing ----------

function validate_(data) {
  if (!data || typeof data !== 'object') throw new Error('Brak danych.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.date))) throw new Error('Nieprawidłowa data.');
  if (!data.training || !String(data.training).trim()) throw new Error('Brak nazwy treningu.');
  if (!data.exercises || !data.exercises.length) throw new Error('Brak ćwiczeń.');
  if (data.exercises.length > 50) throw new Error('Za dużo ćwiczeń.');
  if (data.emails !== undefined && !Array.isArray(data.emails)) throw new Error('Nieprawidłowa lista e-maili.');
  if ((data.emails || []).length > MAX_EMAILS) throw new Error('Maksymalnie ' + MAX_EMAILS + ' adresów e-mail.');
  (data.emails || []).forEach(function (address) {
    if (!/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(String(address))) {
      throw new Error('Nieprawidłowy adres e-mail: ' + address);
    }
  });
  if (data.wellness !== undefined && data.wellness !== null) {
    if (typeof data.wellness !== 'object') throw new Error('Nieprawidłowe samopoczucie.');
    WELLNESS.forEach(function (item) {
      var value = data.wellness[item.key];
      if (value !== undefined && value !== null && !(value === Math.round(value) && value >= 1 && value <= 5)) {
        throw new Error('Samopoczucie: wartość ' + item.mail + ' musi być od 1 do 5.');
      }
    });
  }
  if (data.media !== undefined && !Array.isArray(data.media)) throw new Error('Nieprawidłowa lista plików.');
  if ((data.media || []).length > MAX_MEDIA) throw new Error('Maksymalnie ' + MAX_MEDIA + ' plików.');
  (data.media || []).forEach(function (item) {
    if (!item || !/^[A-Za-z0-9_-]{10,}$/.test(String(item.id))) throw new Error('Nieprawidłowy plik.');
  });
  data.exercises.forEach(function (exercise) {
    var label = '"' + exercise.name + '"';
    if (!exercise.name || !String(exercise.name).trim()) throw new Error('Ćwiczenie bez nazwy.');
    if (!exercise.sets || !exercise.sets.length || exercise.sets.length > 30) {
      throw new Error('Nieprawidłowa liczba serii w ćwiczeniu ' + label + '.');
    }
    var rawCode = exercise.code || exercise.group;
    if (rawCode && !normalizeCode_(rawCode)) {
      throw new Error('Nieprawidłowy kod (np. A1, C2b) w ćwiczeniu ' + label + '.');
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
 * new or has fewer sets than needed. New exercises are placed under the last row of their
 * section (A1 for A1b), else of their block (A), else at the end.
 */
function ensureRows_(sheet, timeZone, name, code, setCount) {
  var key = name.toLowerCase();
  var layout = readLayout_(sheet, timeZone);
  var own = layout.rows.filter(function (entry) { return entry.name.toLowerCase() === key; });
  var missing = setCount - own.length;

  if (missing > 0) {
    var after;
    if (own.length) {
      after = own[own.length - 1].row;
    } else {
      var slot = slotOf_(code);
      var sameSlot = slot ? layout.rows.filter(function (entry) { return slotOf_(entry.code) === slot; }) : [];
      var sameBlock = code ? layout.rows.filter(function (entry) { return entry.group === code.charAt(0); }) : [];
      after = sameSlot.length ? sameSlot[sameSlot.length - 1].row
        : sameBlock.length ? sameBlock[sameBlock.length - 1].row
        : (layout.rows.length ? layout.rows[layout.rows.length - 1].row : HEADER_ROWS);
    }
    if (after >= sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), missing);
    else sheet.insertRowsAfter(after, missing);
    var fill = [];
    for (var i = 0; i < missing; i++) {
      fill.push([code || '', own.length ? own[0].name : name, own.length + i + 1]);
    }
    sheet.getRange(after + 1, 1, missing, 3).setValues(fill);
    layout = readLayout_(sheet, timeZone);
    own = layout.rows.filter(function (entry) { return entry.name.toLowerCase() === key; });
  }

  // A code chosen on the page also updates the code of an existing exercise.
  if (code) {
    own.forEach(function (entry) {
      if (entry.code !== code) sheet.getRange(entry.row, 1).setValue(code);
    });
  }
  return own.slice(0, setCount).map(function (entry) { return entry.row; });
}

/** Rows of the wellness block, created at the bottom the first time they are needed. */
function ensureWellnessRows_(sheet, timeZone, keys) {
  var layout = readLayout_(sheet, timeZone);
  var missing = WELLNESS.filter(function (item) { return keys.indexOf(item.key) >= 0 && !layout.wellness[item.key]; });
  if (missing.length) {
    var start = Math.max(sheet.getLastRow(), HEADER_ROWS) + 1;
    var end = start + missing.length - 1;
    if (end > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), end - sheet.getMaxRows());
    sheet.getRange(start, 1, missing.length, 3).setValues(missing.map(function (item) { return [WELLNESS_BLOCK, item.label, '']; }));
    sheet.getRange(start, 1, missing.length, 1)
      .setBackground(COLOR_SUBHEADER).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(start, 2, missing.length, 1).setFontColor('#000000').setFontWeight('bold');
    sheet.getRange(start, 1, 1, Math.max(sheet.getLastColumn(), 3))
      .setBorder(true, null, null, null, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    layout = readLayout_(sheet, timeZone);
  }
  return layout.wellness;
}

function appendSession_(sheet, data, timeZone, files) {
  var date = Utilities.parseDate(data.date, timeZone, 'yyyy-MM-dd');

  // 1. Make sure every exercise has rows for all of its sets.
  var plan = data.exercises.map(function (exercise) {
    var timed = exercise.mode === 'time';
    var name = String(exercise.name).trim();
    var code = normalizeCode_(exercise.code || exercise.group);
    var rows = ensureRows_(sheet, timeZone, name, code, exercise.sets.length);
    return { exercise: exercise, timed: timed, name: name, rows: rows };
  });

  var mood = data.wellness || {};
  var moodKeys = WELLNESS.filter(function (item) { return mood[item.key]; }).map(function (item) { return item.key; });
  var moodRows = moodKeys.length ? ensureWellnessRows_(sheet, timeZone, moodKeys) : {};

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

  moodKeys.forEach(function (key) {
    var rowIndex = moodRows[key] - HEADER_ROWS - 1;
    values[rowIndex] = [mood[key], ''];
    formats[rowIndex][0] = REPS_FORMAT;
    backgrounds[rowIndex][0] = WELLNESS_COLORS[mood[key] - 1];
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
  var headNote = [];
  if (data.notes && String(data.notes).trim()) headNote.push(String(data.notes).trim());
  if (files && files.length) {
    headNote.push('Zdjęcia i filmy:\n' + files.map(function (file) { return '- ' + file.name + ': ' + file.url; }).join('\n'));
  }
  if (headNote.length) sheet.getRange(1, col).setNote(headNote.join('\n\n'));
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

// ---------- photos and videos ----------

function requirePin_(data) {
  var pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin) {
    throw new Error('Maile i pliki są wyłączone: dodaj w skrypcie właściwość PIN (Ustawienia projektu → Właściwości skryptu).');
  }
  if (String(data.pin || '') !== pin) throw new Error('Zły kod dostępu.');
}

function mediaFolder_() {
  var folders = DriveApp.getFoldersByName(MEDIA_FOLDER);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(MEDIA_FOLDER);
}

/** Opens a resumable upload session in Drive; the page then sends the file in chunks. */
function uploadStart_(data) {
  requirePin_(data);
  var size = Number(data.size);
  if (!(size > 0)) throw new Error('Pusty plik.');
  var name = String(data.name || 'plik').replace(/[\\/\\\\:*?"<>|]/g, '-').substring(0, 150);
  var mime = String(data.mime || 'application/octet-stream');
  var response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'X-Upload-Content-Type': mime,
      'X-Upload-Content-Length': String(size)
    },
    payload: JSON.stringify({ name: name, mimeType: mime, parents: [mediaFolder_().getId()] }),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Dysk Google odrzucił plik (' + response.getResponseCode() + '): ' + response.getContentText().substring(0, 200));
  }
  var headers = response.getHeaders();
  var uri = headers['Location'] || headers['location'];
  if (!uri) throw new Error('Dysk Google nie zwrócił adresu wysyłki.');
  var uploadId = Utilities.getUuid();
  CacheService.getScriptCache().put(uploadId, JSON.stringify({ uri: uri, size: size, mime: mime, name: name }), 21600);
  return { ok: true, uploadId: uploadId };
}

/** Forwards one chunk (multiples of 256 KiB, except the last) to the Drive upload session. */
function uploadChunk_(data) {
  requirePin_(data);
  var raw = CacheService.getScriptCache().get(String(data.uploadId));
  if (!raw) throw new Error('Wysyłka wygasła, dodaj plik jeszcze raz.');
  var info = JSON.parse(raw);
  var offset = Number(data.offset);
  var bytes = Utilities.base64Decode(String(data.data));
  if (!(offset >= 0) || !bytes.length || offset + bytes.length > info.size) throw new Error('Nieprawidłowy fragment pliku.');
  var response = UrlFetchApp.fetch(info.uri, {
    method: 'put',
    contentType: info.mime,
    headers: { 'Content-Range': 'bytes ' + offset + '-' + (offset + bytes.length - 1) + '/' + info.size },
    payload: bytes,
    followRedirects: false,
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code === 308) return { ok: true, done: false };
  if (code === 200 || code === 201) {
    var id = JSON.parse(response.getContentText()).id;
    return { ok: true, done: true, id: id, name: info.name, url: fileUrl_(id) };
  }
  throw new Error('Wysyłka na Dysk nie powiodła się (' + code + ').');
}

function fileUrl_(id) {
  return 'https://drive.google.com/file/d/' + id + '/view';
}

/** The uploaded files named by the page, only if they are really inside the media folder. */
function mediaFiles_(media) {
  var folderId = mediaFolder_().getId();
  return media.map(function (item) {
    var file = DriveApp.getFileById(String(item.id));
    var parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === folderId) return { id: file.getId(), name: file.getName(), url: fileUrl_(file.getId()), file: file };
    }
    throw new Error('Plik "' + file.getName() + '" nie jest w folderze "' + MEDIA_FOLDER + '".');
  });
}

/** Lets the recipients open the files (Google accounts only; others just can't, which is fine). */
function shareFiles_(files, emails) {
  files.forEach(function (item) {
    emails.forEach(function (address) {
      try { item.file.addViewer(address); } catch (error) { /* not a Google account */ }
    });
  });
}

// ---------- e-mail summary ----------

function sendSummary_(data, sheetName, before, files, emails) {
  var summary = buildSummary_(data, sheetName, before, files);
  MailApp.sendEmail({
    to: emails.join(','),
    subject: summary.subject,
    body: summary.text,
    htmlBody: summary.html,
    name: 'Treningi'
  });
}

function buildSummary_(data, sheetName, before, files) {
  var parts = String(data.date).split('-');
  var shownDate = parts[2] + '.' + parts[1] + '.' + parts[0];
  var subject = 'Trening ' + sheetName + ' – ' + shownDate;
  var text = [subject, ''];
  var html = ['<div style="font-family:Arial,sans-serif;color:#17212b;max-width:560px">',
    '<h2 style="margin:0 0 12px">' + esc_(subject) + '</h2>'];

  if (data.notes && String(data.notes).trim()) {
    text.push('Notatka: ' + String(data.notes).trim(), '');
    html.push('<p style="margin:0 0 14px"><i>' + esc_(String(data.notes).trim()) + '</i></p>');
  }

  var mood = WELLNESS.filter(function (item) { return data.wellness && data.wellness[item.key]; })
    .map(function (item) { return item.mail + ' ' + data.wellness[item.key] + '/5'; });
  if (mood.length) {
    text.push('Samopoczucie: ' + mood.join(' · '), '');
    html.push('<p style="margin:0 0 14px">Samopoczucie: <b>' + esc_(mood.join(' · ')) + '</b></p>');
  }

  data.exercises.forEach(function (exercise) {
    var mode = exercise.mode === 'time' ? 'time' : 'reps';
    var code = normalizeCode_(exercise.code || exercise.group);
    var sets = exercise.sets.map(function (set) {
      return { value: Number(mode === 'time' ? set.seconds : set.reps), weight: setWeight_(set.weight) };
    });
    var now = describeSets_(mode, sets);
    var prev = before[String(exercise.name).trim().toLowerCase()];
    var prevText = '';
    var diff = '';
    if (prev) {
      var prevSets = prev.sets.map(function (set) {
        return { value: prev.mode === 'time' ? set.seconds : set.reps, weight: setWeight_(set.weight) };
      });
      prevText = describeSets_(prev.mode, prevSets);
      diff = diffText_(summarize_(prev.mode, prevSets), summarize_(mode, sets));
    }
    var title = (code ? code + '  ' : '') + String(exercise.name).trim();
    text.push(title, '  ' + now + (prevText ? '   (poprzednio: ' + prevText + (diff ? ', ' + diff : ', bez zmian') + ')' : ''));
    var line = '<b>' + esc_(now) + '</b>';
    if (prevText) line += ' <span style="color:#6b7684">· poprzednio: ' + esc_(prevText) + '</span> ' + (diff ? diffHtml_(diff) : '<span style="color:#6b7684">bez zmian</span>');
    html.push('<div style="margin:0 0 12px"><div style="font-weight:700">' + esc_(title) + '</div><div>' + line + '</div>');
    if (exercise.note && String(exercise.note).trim()) {
      text.push('  Uwaga: ' + String(exercise.note).trim());
      html.push('<div style="color:#6b7684">Uwaga: ' + esc_(String(exercise.note).trim()) + '</div>');
    }
    text.push('');
    html.push('</div>');
  });

  if (files && files.length) {
    text.push('Zdjęcia i filmy:');
    html.push('<div style="margin-top:6px"><b>Zdjęcia i filmy</b><ul style="margin:4px 0 0;padding-left:20px">');
    files.forEach(function (file) {
      text.push('- ' + file.name + ': ' + file.url);
      html.push('<li><a href="' + esc_(file.url) + '">' + esc_(file.name) + '</a></li>');
    });
    html.push('</ul></div>');
  }
  html.push('</div>');
  return { subject: subject, text: text.join('\n'), html: html.join('') };
}

function setWeight_(weight) {
  if (weight === null || weight === undefined || weight === '') return 'BW';
  return isBodyweight_(weight) ? 'BW' : Number(weight);
}

function fmtNumber_(n) {
  return String(Math.round(n * 100) / 100).replace('.', ',');
}

function fmtTime_(seconds) {
  if (seconds < 60) return seconds + '"';
  var rest = seconds % 60;
  return Math.floor(seconds / 60) + "'" + (rest ? ('0' + rest).slice(-2) + '"' : '');
}

/** "3 × 10 × 30 kg · 8 × BW": equal neighbouring sets are merged. */
function describeSets_(mode, sets) {
  var groups = [];
  sets.forEach(function (set) {
    var last = groups[groups.length - 1];
    if (last && last.value === set.value && last.weight === set.weight) last.count += 1;
    else groups.push({ value: set.value, weight: set.weight, count: 1 });
  });
  return groups.map(function (group) {
    var text = mode === 'time' ? fmtTime_(group.value) : String(group.value);
    if (group.weight === 'BW') text += ' × BW';
    else if (group.weight > 0) text += ' × ' + fmtNumber_(group.weight) + ' kg';
    return (group.count > 1 ? group.count + ' × ' : '') + text;
  }).join(' · ');
}

function summarize_(mode, sets) {
  var top = 0;
  var total = 0;
  sets.forEach(function (set) {
    if (typeof set.weight === 'number') top = Math.max(top, set.weight);
    total += set.value;
  });
  return { mode: mode, top: top, total: total, count: sets.length };
}

/** Same wording as the page: "▲ +2,5 kg, +2 powt." (empty when nothing changed). */
function diffText_(before, now) {
  var items = [{ diff: now.top - before.top, unit: 'kg' }];
  if (before.mode === now.mode) items.push({ diff: now.total - before.total, unit: now.mode === 'time' ? 's' : 'powt.' });
  items.push({ diff: now.count - before.count, unit: 'serii' });
  items = items.filter(function (item) { return Math.round(item.diff * 100) !== 0; });
  if (!items.length) return '';
  return (items[0].diff > 0 ? '▲ ' : '▼ ') + items.map(function (item) {
    return (item.diff > 0 ? '+' : '') + fmtNumber_(item.diff) + ' ' + item.unit;
  }).join(', ');
}

function diffHtml_(text) {
  var color = text.charAt(0) === '▲' ? '#17784a' : '#b83a2e';
  return '<b style="color:' + color + '">' + esc_(text) + '</b>';
}

function esc_(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Run once from the editor (select "authorize" and press Run) to grant Drive and Mail access. */
function authorize() {
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis', { muteHttpExceptions: true });
  ScriptApp.getOAuthToken();
}

// ---------- helpers ----------

/** "c2B" -> "C2b"; "" for anything that is not letter + up to two digits + optional variant letter (needs digits). */
function normalizeCode_(value) {
  var text = String(value === null || value === undefined ? '' : value).trim();
  var match = /^([A-Za-z])(\d{0,2})([A-Za-z]?)$/.exec(text);
  if (!match || (match[3] && !match[2])) return '';
  return match[1].toUpperCase() + match[2] + match[3].toLowerCase();
}

/** Section of a code: "C2b" -> "C2" (variants of a section are alternatives); "" when there is no number. */
function slotOf_(code) {
  var match = /^([A-Z]\d+)[a-z]?$/.exec(String(code || ''));
  return match ? match[1] : '';
}

function isBodyweight_(value) {
  return String(value).trim().toUpperCase() === 'BW';
}

function json_(object) {
  return ContentService.createTextOutput(JSON.stringify(object)).setMimeType(ContentService.MimeType.JSON);
}

function message_(error) {
  return String(error && error.message ? error.message : error);
}
