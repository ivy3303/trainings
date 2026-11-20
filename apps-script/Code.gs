/**
 * Backend for index.html, bound to the Google Sheet with the training log.
 *
 * One TRAINING (e.g. "FBW A", "Push") = one sheet (tab). Layout of a tab:
 *
 *   Blok | Ćwiczenie    | Seria |  2026-11-10   |  2026-11-24   | ...
 *   Stan | Samopoczucie |       |    ● ● ●      |    ● ● ●      | ...   (energy, sleep, food)
 *        |              |       | Powt./czas | Kg | Powt./czas | Kg | ...
 *   A0   | Wioślarz  |   1   |   120 s    |    |            |    |
 *    ...
 *   UWAGI (merged A:C)     |  notes of the whole training, one merged cell per date   (always the last row)
 *
 * "Blok" holds the exercise code from the plan: block letter + section number + optional
 * variant letter, e.g. A1, C2a, C2b. Variants of one section (C2a / C2b) are alternatives:
 * the form shows one tile per section and lets you switch between its variants.
 *
 * How the day went is the second header row, right under the date: three coloured dots (energy, sleep,
 * food; red / yellow / green = 1 / 2 / 3, grey = not rated) in one merged cell, with the words and the
 * free-text remark in that cell's note.
 *
 * Exercises and sets are rows; every saved session adds two columns on the right
 * (reps or seconds, and kg). Reps are coloured against the previous date:
 * green = more, red = fewer, blue = the same. Remarks sit in cell notes:
 * the exercise remark on the first set's reps cell, the links to photos and videos on the date.
 * The remark on the whole training goes to the UWAGI row at the bottom.
 *
 * A successful save answers { ok, sheet, media, summary (HTML of the session summary), emailed?, emailError? }.
 *
 * GET  ?action=list  -> { ok, version, trainings: [{ name, lastDate, exercises }] }
 *                       exercises = every exercise of the training with its most recent results:
 *                       [{ name, code, group, mode, note, date, inLast, sets: [{ reps, seconds, weight }] }]
 *                       inLast = it was done in the training's last session.
 * POST (JSON as text/plain)
 *   save a session:
 *   { training, create, date: "YYYY-MM-DD", notes,
 *     wellness: { energy, sleep, food (each 1-3), note (text) } (all optional),
 *     exercises: [{ name, mode: "reps" | "time", note, code (e.g. "A1", "C2b"; optional),
 *                   sets: [{ reps, seconds, weight }] }],
 *     emails: ["a@b.pl"], media: [{ id, exercise (name of an exercise of this session; optional) }], pin }
 *   create=true makes a new tab, create=false adds a column pair to an existing one.
 *   mode "reps" uses sets[].reps, mode "time" uses sets[].seconds;
 *   weight is optional: a number (kg), "BW" or empty.
 *   emails -> a summary of the session is mailed from the owner's account (needs pin).
 *   media  -> files already uploaded to the "trainings" Drive folder (see below).
 *   upload a photo or video to Drive in chunks (all need pin):
 *   { action: "uploadStart", name, mime, size, pin }            -> { ok, uploadId }
 *   { action: "uploadChunk", uploadId, offset, data (base64), pin } -> { ok, done, id?, url? }
 *
 * Mail and Drive are only available when the script property PIN is set
 * (Project settings -> Script properties); the page sends it with every such request.
 * Run authorize() once from the editor to grant the Drive and Mail permissions.
 */

// Bumped when the response format or sheet layout changes; index.html checks it to detect an outdated deployment.
var VERSION = 11;
var LEFT_HEADER = ['Blok', 'Ćwiczenie', 'Seria'];
var HEADER_ROWS = 3;
var WELLNESS_ROW = 2;
var FIRST_SESSION_COL = 4;
var TIME_FORMAT = '0" s"';
var REPS_FORMAT = '0';
var WEIGHT_FORMAT = '0.0#';
var COLOR_NAVY = '#0f2a4a';
var COLOR_SUBHEADER = '#e4e9f2';
var MEDIA_FOLDER = 'trainings';
var OLD_MEDIA_FOLDER = 'Treningi – media'; // where photos went before the folder was renamed
var NOTES_LABEL = 'UWAGI';
var COLOR_NOTES = '#f6f7fb';
var WELLNESS_BLOCK = 'Stan';
var WELLNESS_LABEL = 'Samopoczucie';
var WELLNESS = [
  { key: 'energy', label: 'Energia', mail: 'energia', words: ['Mało sił', 'Średnio', 'Forma mutant'] },
  { key: 'sleep', label: 'Sen', mail: 'sen', words: ['Słabo', 'Średnio', 'Dobrze'] },
  { key: 'food', label: 'Jedzenie', mail: 'jedzenie', words: ['Kiepsko', 'Średnio', 'Dobrze'] }
];
// 1 (worst) .. 3 (best); grey = not rated
var WELLNESS_DOT_COLORS = ['#e5534b', '#f2c230', '#34c47c'];
var WELLNESS_DOT_EMPTY = '#b8bfca';
var WELLNESS_NOTE_MAX = 500;
var MAX_EMAILS = 5;
var MAX_MEDIA = 20;
// Pastel fills: block column and reps comparison with the previous date.
var GROUP_COLORS = { A: '#c6ebdc', B: '#fbd9b8', C: '#cbdcf9', D: '#efcff2', E: '#f8edb0' };
// Colours of the summary mail, as on the page: the three mood dots and the block chip/strip.
var MAIL_MOOD_COLORS = ['#f0746c', '#f1df8e', '#4fd09a'];
var MAIL_CHIP_COLORS = { A: '#9fdcc3', B: '#f6c796', C: '#a9c6f4', D: '#e3b0ea', E: '#f4a3b8' };
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
    var before = previousByName_(sheet, timeZone);
    appendSession_(sheet, data, timeZone, files);

    // The same summary is mailed and shown on the page after saving (summary = the HTML without the mail background).
    var summary = null;
    try { summary = buildSummary_(data, sheet.getName(), before, files); } catch (error) { summary = null; }
    var result = { ok: true, sheet: sheet.getName(), media: files.length, summary: summary ? summary.inner : '' };
    if (emails.length) {
      try {
        if (!summary) throw new Error('Nie udało się zbudować podsumowania.');
        shareFiles_(files, emails);
        sendSummary_(summary, emails);
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
      exercises: last ? latestExercises_(sheet, layout) : [],
      wellness: last ? readWellness_(sheet, last.col) : {}
    };
  });
}

/**
 * How the day went in a session: { energy, sleep, food (1-3), note }, read back from the note of the
 * dots cell ("Energia: Średnio (2/3)" ... "Uwagi: ..."); only what is there.
 */
function readWellness_(sheet, col) {
  var text = String(sheet.getRange(WELLNESS_ROW, col).getNote() || '');
  var out = {};
  var cut = text.indexOf('Uwagi: ');
  var scales = cut >= 0 ? text.slice(0, cut) : text;
  if (cut >= 0 && text.slice(cut + 7).trim()) out.note = text.slice(cut + 7).trim();
  scales.split('\n').forEach(function (line) {
    WELLNESS.forEach(function (item) {
      if (line.indexOf(item.label + ':') !== 0) return;
      var match = /\(([123])\/3\)/.exec(line);
      if (match) out[item.key] = Number(match[1]);
    });
  });
  return out;
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
  return { rows: rows, sessions: sessions };
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
      if (value !== undefined && value !== null && !(value === Math.round(value) && value >= 1 && value <= 3)) {
        throw new Error('Samopoczucie: wartość ' + item.mail + ' musi być od 1 do 3.');
      }
    });
    var remark = data.wellness.note;
    if (remark !== undefined && remark !== null && (typeof remark !== 'string' || remark.length > WELLNESS_NOTE_MAX)) {
      throw new Error('Uwagi do samopoczucia: maksymalnie ' + WELLNESS_NOTE_MAX + ' znaków.');
    }
  }
  if (data.media !== undefined && !Array.isArray(data.media)) throw new Error('Nieprawidłowa lista plików.');
  if ((data.media || []).length > MAX_MEDIA) throw new Error('Maksymalnie ' + MAX_MEDIA + ' plików.');
  (data.media || []).forEach(function (item) {
    if (!item || !/^[A-Za-z0-9_-]{10,}$/.test(String(item.id))) throw new Error('Nieprawidłowy plik.');
    if (item.exercise !== undefined && typeof item.exercise !== 'string') throw new Error('Nieprawidłowe ćwiczenie przy pliku.');
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
  sheet.getRange(1, 1, 1, 3).setValues([LEFT_HEADER])
    .setBackground(COLOR_NAVY).setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(WELLNESS_ROW, 1, 1, 3).setValues([[WELLNESS_BLOCK, WELLNESS_LABEL, '']]);
  sheet.getRange(WELLNESS_ROW, 1, HEADER_ROWS - 1, 3)
    .setBackground(COLOR_SUBHEADER).setFontColor('#000000').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(HEADER_ROWS);
  sheet.setFrozenColumns(3);
  [60, 340, 60].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.setTabColor(COLOR_NAVY);
  ensureNotesRow_(sheet);
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

/** Row of the closing UWAGI row (notes of whole trainings), 0 when the tab has none yet. */
function findNotesRow_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROWS) return 0;
  var left = sheet.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, 1).getValues();
  for (var i = left.length - 1; i >= 0; i--) {
    if (String(left[i][0]).trim().toUpperCase() === NOTES_LABEL) return HEADER_ROWS + 1 + i;
  }
  return 0;
}

/** The UWAGI row is the last row of a tab; tabs made before it existed get it under the last exercise. */
function ensureNotesRow_(sheet) {
  var row = findNotesRow_(sheet);
  if (row) return row;
  row = Math.max(sheet.getLastRow(), HEADER_ROWS) + 1;
  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  var label = sheet.getRange(row, 1, 1, 3);
  label.merge();
  sheet.getRange(row, 1).setValue(NOTES_LABEL);
  label.setBackground(COLOR_SUBHEADER).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 3))
    .setBorder(true, null, null, null, null, null, COLOR_NAVY, SpreadsheetApp.BorderStyle.SOLID);
  return row;
}

/** The note on the whole training of a date: one merged cell (reps + kg columns) in the UWAGI row. */
function writeNotes_(sheet, row, col, text) {
  var note = text ? String(text).trim() : '';
  var cell = sheet.getRange(row, col, 1, 2);
  cell.merge();
  cell.setNumberFormat('@'); // plain text, so a note starting with "=" is never a formula
  sheet.getRange(row, col).setValue(note);
  cell.setBackground(COLOR_NOTES).setFontWeight('normal').setWrap(true)
    .setHorizontalAlignment('left').setVerticalAlignment('top');
  if (note) {
    // merged cells do not auto-fit: about 22 characters fit on a line of the two columns
    var lines = 0;
    note.split('\n').forEach(function (paragraph) { lines += Math.max(1, Math.ceil(paragraph.length / 22)); });
    sheet.setRowHeight(row, Math.min(Math.max(sheet.getRowHeight(row), lines * 16 + 10), 400));
  }
}

/** The second header row of a new date: three coloured dots (energy, sleep, food), words and remark in the note. */
function writeWellness_(sheet, col, mood) {
  var remark = String(mood.note || '').trim();
  var rated = WELLNESS.filter(function (item) { return mood[item.key]; });
  var cell = sheet.getRange(WELLNESS_ROW, col, 1, 2);
  cell.merge();
  cell.setBackground(COLOR_SUBHEADER).setHorizontalAlignment('center').setVerticalAlignment('middle');
  if (!rated.length && !remark) return;

  var builder = SpreadsheetApp.newRichTextValue().setText('● ● ●');
  WELLNESS.forEach(function (item, index) {
    var color = mood[item.key] ? WELLNESS_DOT_COLORS[mood[item.key] - 1] : WELLNESS_DOT_EMPTY;
    builder.setTextStyle(index * 2, index * 2 + 1, SpreadsheetApp.newTextStyle().setForegroundColor(color).build());
  });
  var top = sheet.getRange(WELLNESS_ROW, col);
  top.setRichTextValue(builder.build());
  top.setFontSize(16);

  var lines = rated.map(function (item) { return item.label + ': ' + item.words[mood[item.key] - 1] + ' (' + mood[item.key] + '/3)'; });
  if (remark) lines.push('', 'Uwagi: ' + remark);
  top.setNote(lines.join('\n'));
}

function appendSession_(sheet, data, timeZone, files) {
  var date = Utilities.parseDate(data.date, timeZone, 'yyyy-MM-dd');

  // 1. Every exercise gets rows for all of its sets.
  var plan = data.exercises.map(function (exercise) {
    var timed = exercise.mode === 'time';
    var name = String(exercise.name).trim();
    var code = normalizeCode_(exercise.code || exercise.group);
    var rows = ensureRows_(sheet, timeZone, name, code, exercise.sets.length);
    return { exercise: exercise, timed: timed, name: name, rows: rows };
  });

  // 2. The new column pair goes to the right of the last session.
  var notesRow = ensureNotesRow_(sheet);
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
  writeNotes_(sheet, notesRow, col, data.notes);
  // 3. Header of the new date.
  var head = sheet.getRange(1, col, 1, 2);
  head.merge();
  head.setNumberFormat('yyyy-mm-dd')
    .setBackground(COLOR_NAVY).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(1, col).setValue(date);
  var headNote = [];
  if (files && files.length) {
    headNote.push('Zdjęcia i filmy:\n' + files.map(function (file) { return fileLine_(file, data); }).join('\n'));
  }
  if (headNote.length) sheet.getRange(1, col).setNote(headNote.join('\n\n'));
  writeWellness_(sheet, col, data.wellness || {});
  sheet.getRange(HEADER_ROWS, col, 1, 2).setValues([['Powt./czas', 'Kg']])
    .setBackground(COLOR_SUBHEADER).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setColumnWidth(col, 90);
  sheet.setColumnWidth(col + 1, 70);

  formatLeft_(sheet, timeZone);
}

/** "- foto.jpg [A1]: link" - the exercise code (or name) in brackets when the file belongs to an exercise. */
function fileLine_(file, data) {
  var label = '';
  var match = fileExercise_(file, data.exercises);
  if (match) label = normalizeCode_(match.code || match.group) || String(match.name).trim();
  return '- ' + file.name + (label ? ' [' + label + ']' : '') + ': ' + file.url;
}

/** The exercise of this session a file was assigned to, or null (whole training). */
function fileExercise_(file, exercises) {
  if (!file.exercise) return null;
  var key = String(file.exercise).toLowerCase();
  for (var i = 0; i < exercises.length; i++) {
    if (String(exercises[i].name).trim().toLowerCase() === key) return exercises[i];
  }
  return null;
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

/** The uploaded files named by the page, only if they are really inside a media folder (current or old name). */
function mediaFiles_(media) {
  var folderId = mediaFolder_().getId();
  return media.map(function (item) {
    var file = DriveApp.getFileById(String(item.id));
    var parents = file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (parent.getId() === folderId || parent.getName() === MEDIA_FOLDER || parent.getName() === OLD_MEDIA_FOLDER) {
        return { id: file.getId(), name: file.getName(), url: fileUrl_(file.getId()), file: file, exercise: String(item.exercise || '').trim().slice(0, 200) };
      }
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

function sendSummary_(summary, emails) {
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
  var blocks = [];

  var note = data.notes ? String(data.notes).trim() : '';
  if (note) {
    text.push('Notatka: ' + note, '');
    blocks.push(mailCard_('', '<div style="font-size:15px;line-height:22px;color:#f4f6ff;font-style:italic">' + esc_(note) + '</div>', '#a9bcff'));
  }

  var wellness = data.wellness || {};
  var mood = WELLNESS.filter(function (item) { return wellness[item.key]; })
    .map(function (item) { return item.mail + ' ' + wellness[item.key] + '/3'; });
  var moodNote = wellness.note ? String(wellness.note).trim() : '';
  if (mood.length) text.push('Samopoczucie: ' + mood.join(' · '), '');
  if (moodNote) text.push('Uwagi do samopoczucia: ' + moodNote, '');
  if (mood.length || moodNote) blocks.push(mailMood_(wellness, moodNote));

  var sections = [];
  var general = (files || []).filter(function (file) { return !fileExercise_(file, data.exercises); });
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
    var name = String(exercise.name).trim();
    var remark = exercise.note ? String(exercise.note).trim() : '';
    var own = (files || []).filter(function (file) { return fileExercise_(file, data.exercises) === exercise; });
    text.push((code ? code + '  ' : '') + name, '  ' + now + (prevText ? '   (poprzednio: ' + prevText + (diff ? ', ' + diff : ', bez zmian') + ')' : ''));
    if (remark) text.push('  Uwaga: ' + remark);
    own.forEach(function (file) { text.push('  Zdjęcie: ' + file.name + ': ' + file.url); });
    text.push('');
    // Neighbouring exercises of one block (A1, A2 …) share a single card.
    var letter = code ? code.charAt(0) : '';
    var last = sections[sections.length - 1];
    var row = mailExercise_(code, name, now, prevText, diff, remark, own);
    if (letter && last && last.letter === letter) last.rows.push(row);
    else sections.push({ letter: letter, rows: [row] });
  });
  sections.forEach(function (section) { blocks.push(mailSection_(section.letter, section.rows)); });

  if (general.length) {
    text.push('Zdjęcia i filmy:');
    general.forEach(function (file) { text.push('- ' + file.name + ': ' + file.url); });
    blocks.push(mailMedia_(general));
  }

  var count = data.exercises.length;
  var counted = count + (count === 1 ? ' ćwiczenie' : count > 1 && count < 5 ? ' ćwiczenia' : ' ćwiczeń');
  var pill = 'display:inline-block;padding:5px 14px;border-radius:999px;font-size:13px;line-height:18px;font-weight:700;';
  var inner =
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:560px;margin:0 auto">' +
    '<tr><td style="padding:6px 8px 18px">' +
      '<div style="font-size:14px;color:#aab4d8">Treningi</div>' +
      '<div style="font-size:36px;line-height:42px;font-weight:800;color:#ffffff;margin-top:2px">' + esc_(sheetName) + '</div>' +
      '<div style="margin-top:12px">' +
        '<span style="' + pill + 'background:#6c7dff;background-image:linear-gradient(135deg,#7c5cff,#33c6ff);color:#ffffff">' + esc_(shownDate) + '</span> ' +
        '<span style="' + pill + 'border:1px solid #3a4268;border:1px solid rgba(255,255,255,.22);color:#dfe6ff">' + esc_(counted) + '</span>' +
      '</div>' +
    '</td></tr>' +
    '<tr><td>' + blocks.join('') + '</td></tr>' +
    '</table>';
  var html = '<div style="background:#0b0f1f;background-image:radial-gradient(60% 38% at 8% 0%,rgba(124,92,255,.55),transparent 62%),radial-gradient(50% 36% at 100% 22%,rgba(0,205,210,.32),transparent 62%),radial-gradient(70% 45% at 40% 100%,rgba(255,92,150,.34),transparent 62%);padding:22px 10px;font-family:\'DM Sans\',Arial,Helvetica,sans-serif;color:#f4f6ff">' + inner +
    '<div style="max-width:560px;margin:0 auto;text-align:center;padding:8px 0 2px;font-size:12px;color:#8590bd">Wysłano z formularza Treningi</div></div>';
  return { subject: subject, text: text.join('\n'), html: html, inner: inner };
}

/** A frosted-glass card of the summary mail; accent = colour of the strip (and glow) on its left. */
function mailCard_(title, body, accent) {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 12px;background:#171c37;background:rgba(255,255,255,.085);border:1px solid #333a5e;border:1px solid rgba(255,255,255,.17);border-left:7px solid ' + accent + ';border-radius:24px;box-shadow:-6px 0 16px -8px ' + accent + '">' +
    '<tr><td style="padding:14px 16px">' + (title || '') + body + '</td></tr></table>';
}

function mailMood_(wellness, remark) {
  var cells = WELLNESS.map(function (item) {
    var value = wellness[item.key];
    var dot = value
      ? '<div style="width:40px;height:40px;margin:0 auto;border-radius:50%;background:' + MAIL_MOOD_COLORS[value - 1] + ';box-shadow:0 0 16px ' + MAIL_MOOD_COLORS[value - 1] + '99"></div>'
      : '<div style="width:36px;height:36px;margin:0 auto;border-radius:50%;border:2px dashed #5b6490"></div>';
    return '<td align="center" width="33%" style="padding:2px 4px">' + dot +
      '<div style="font-size:12px;color:#aab4d8;margin-top:8px">' + esc_(item.label) + '</div>' +
      '<div style="font-size:14px;font-weight:700;color:#f4f6ff">' + (value ? esc_(item.words[value - 1]) : '–') + '</div></td>';
  }).join('');
  var body = '<div style="font-size:15px;font-weight:700;color:#f4f6ff;margin:0 0 10px">Samopoczucie</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' + cells + '</tr></table>';
  if (remark) body += '<div style="margin-top:12px;padding-top:10px;border-top:1px solid #333a5e;border-top:1px solid rgba(255,255,255,.14);font-size:14px;line-height:20px;color:#c6cdee;font-style:italic">' + esc_(remark) + '</div>';
  return mailCard_('', body, '#7c5cff');
}

/** One card per training block: its exercises one under another, split by thin lines. */
function mailSection_(letter, rows) {
  var divider = '<div style="height:1px;line-height:1px;font-size:1px;margin:14px 0;background:#333a5e;background:rgba(255,255,255,.14)">&nbsp;</div>';
  return mailCard_('', rows.join(divider), MAIL_CHIP_COLORS[letter] || '#8e99c8');
}

/** One exercise inside a section card. */
function mailExercise_(code, name, now, prevText, diff, remark, files) {
  var color = MAIL_CHIP_COLORS[code.charAt(0)] || '#8e99c8';
  var chip = code
    ? '<span style="display:inline-block;background:' + color + ';border-radius:999px;padding:1px 10px;margin-right:8px;font-size:12px;line-height:20px;font-weight:700;color:#12162a">' + esc_(code) + '</span>'
    : '';
  var body = '<div style="font-size:16px;line-height:24px;font-weight:700;color:#f4f6ff">' + chip + esc_(name) + '</div>' +
    '<div style="font-size:20px;line-height:28px;font-weight:800;color:#ffffff;margin-top:6px">' + esc_(now) + '</div>';
  if (prevText || diff) {
    body += '<div style="font-size:13px;line-height:24px;color:#aab4d8;margin-top:2px">' +
      (prevText ? 'poprzednio: ' + esc_(prevText) : '') + (diff ? ' ' + diffHtml_(diff) : '') + '</div>';
  }
  if (remark) body += '<div style="margin-top:8px;font-size:13px;line-height:19px;color:#c6cdee;font-style:italic">' + esc_(remark) + '</div>';
  if (files && files.length) body += '<div style="margin-top:10px">' + mailLinks_(files) + '</div>';
  return body;
}

/** Pills linking to photos and videos. */
function mailLinks_(files) {
  return files.map(function (file) {
    return '<a href="' + esc_(file.url) + '" style="display:inline-block;margin:0 6px 8px 0;padding:8px 14px;border:1.5px dashed #6f7fc0;border-radius:999px;color:#a9bcff;font-size:13px;font-weight:700;text-decoration:none">&#128247; ' + esc_(file.name) + '</a>';
  }).join('');
}

/** Card with the photos and videos of the whole training (those not assigned to an exercise). */
function mailMedia_(files) {
  return mailCard_('', '<div style="font-size:15px;font-weight:700;color:#f4f6ff;margin:0 0 8px">Zdjęcia i filmy</div>' + mailLinks_(files), '#33c6ff');
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

/** Same wording as the page: "▲ +2,5 kg, ▼ -1 powt." - every change has its own arrow (empty when nothing changed). */
function diffText_(before, now) {
  var items = [{ diff: now.top - before.top, unit: 'kg' }];
  if (before.mode === now.mode) items.push({ diff: now.total - before.total, unit: now.mode === 'time' ? 's' : 'powt.' });
  items.push({ diff: now.count - before.count, unit: 'serii' });
  return items.filter(function (item) { return Math.round(item.diff * 100) !== 0; }).map(function (item) {
    return (item.diff > 0 ? '▲ +' : '▼ ') + fmtNumber_(item.diff) + ' ' + item.unit;
  }).join(', ');
}

/** One pill per change: green for more, red for less; fewer reps or sets with more kilograms is progress, so it is grey. */
function diffHtml_(text) {
  var parts = text.split(', ');
  var heavier = /^▲ \+[\d,.]+ kg$/.test(parts[0]);
  return parts.map(function (part, i) {
    var up = part.charAt(0) === '▲';
    var flat = !up && heavier && i > 0;
    var bg = up ? 'rgba(40,200,140,.22)' : flat ? 'rgba(160,170,200,.18)' : 'rgba(255,90,90,.22)';
    var fg = up ? '#5ef0b0' : flat ? '#b8c2dc' : '#ff8a80';
    var solid = up ? '#17503f' : flat ? '#3a4258' : '#5a2a35';
    return '<span style="display:inline-block;margin:0 4px 0 2px;padding:1px 10px;border-radius:999px;font-size:12px;font-weight:700;background:' + solid + ';background:' + bg + ';color:' + fg + '">' + esc_(part) + '</span>';
  }).join('');
}

function esc_(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Run once from the editor (select "authorize" and press Run) to grant Drive and Mail access. */
function authorize() {
  mediaFolder_(); // creates the media folder, which asks for full Drive access
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
