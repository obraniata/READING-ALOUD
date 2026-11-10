/**
 * READING ALOUD ASSESSMENT -- reusable "Make a copy" template (v1.0)
 * Container-bound Apps Script: create it from the Google Sheet itself
 * (Extensions > Apps Script) so it is copied together with the Sheet.
 *
 * Everything a teacher customizes lives in the Sheet, never in code:
 *   Settings        titles, class names, leaderboard on/off, access method, PIN
 *   Passages        one row per reading (paste text, questions, key terms)
 *   Benchmarks      WCPM tiers
 *   Student Access  what students typed to protect their own reports
 *   Results - ...   one results tab per class (created automatically)
 *
 * Scoring, tagging, rubrics, PDF reports and the Sheet round trip carry over
 * unchanged from the Southwest Asia edition. New in this version: Passages and
 * Settings tabs, automatic sound-focus detection, a passage snapshot saved with
 * every session, a leaderboard on/off switch, student self-registered access
 * (birthdate or secret code) stored in the "Student Access" tab, and an
 * optional teacher PIN.
 */

var TAB_START = 'Start Here';
var TAB_SETTINGS = 'Settings';
var TAB_PASSAGES = 'Passages';
var TAB_BENCH = 'Benchmarks';
var TAB_ACCESS = 'Student Access';
var RESULTS_PREFIX = 'Results - ';
var SUBMIT_LOCK_WAIT_MS = 30000;

// ---------- Results-tab schema (first 31 columns identical to the earlier edition) ----------
var HEADER_ROW = [
  'Timestamp', 'Student ID', 'Full Name', 'Section',
  'Total Words', 'Miscues Tagged (True Errors)', 'Accuracy (%)', 'Session Notes',
  'Miscue Tally (JSON)', 'Responses (JSON)',
  'Reading Duration (Sec)', 'WCPM', 'Self-Corrections',
  'Reading Fluency Band (CEFR-insp.)',
  'Prosody: Expression & Volume', 'Prosody: Phrasing & Syntax',
  'Prosody: Smoothness & Pausing', 'Prosody: Punctuation Adherence',
  'Prosody Overall Level',
  'WCPM Tier',
  'Passage ID', 'Passage Title',
  'Sound: L & R', 'Sound: W & V', 'Sound: Final S/ES', 'Sound: -ED Endings',
  'Sound Focus Overall Level',
  'Comprehension: Main Idea', 'Comprehension: Key Details',
  'Comprehension: Key Terms', 'Comprehension Overall Level',
  'Passage Text Snapshot', 'Sound Focus Was On'
];
var COL_TIMESTAMP = 1, COL_STUDENT_ID = 2, COL_NAME = 3, COL_SECTION = 4, COL_TOTAL_WORDS = 5,
    COL_MISCUE_COUNT = 6, COL_ACCURACY = 7, COL_SESSION_NOTES = 8, COL_TALLY_JSON = 9,
    COL_RESPONSES_JSON = 10, COL_DURATION_SECONDS = 11, COL_WCPM = 12, COL_SELF_CORRECTIONS = 13,
    COL_READING_FLUENCY_BAND = 14, COL_PROSODY_EXPRESSION = 15, COL_PROSODY_OVERALL = 19,
    COL_WCPM_TIER = 20, COL_PASSAGE_ID = 21, COL_PASSAGE_TITLE = 22, COL_SOUND_LR = 23,
    COL_SOUND_OVERALL = 27, COL_COMP_MAIN_IDEA = 28, COL_COMP_OVERALL = 31,
    COL_SNAPSHOT = 32, COL_SOUND_ON = 33;

// ---------- Settings tab definitions: [key, default value, help text] ----------
var DEFAULT_CAVEAT = 'Working benchmark adapted from published EFL/ESL reading-fluency research, not an officially validated norm.';
var SETTINGS_DEFS = [
  ['App Title', 'Reading Aloud Assessment', 'Shown at the top of the teacher and student pages.'],
  ['Subject / Course', 'English Reading', 'Shown under the title and on printed reports.'],
  ['School / Teacher Name', '', 'Optional. Printed on reports.'],
  ['Sections', 'Class A, Class B', 'Your classes, separated by commas. Each class gets its own Results tab.'],
  ['Show Leaderboard To Students', true, 'Tick = students see the leaderboard. Untick = students only see their own reports.'],
  ['Student Access Method', 'Birthdate', 'Birthdate, Secret code, or ID only. How students unlock their own reports.'],
  ['Student ID Length', 5, 'How many digits a student ID has. Use 0 to allow any ID (letters and numbers).'],
  ['Default Sound Focus', true, 'Tick = Sound Focus is ON for passages whose Sound Focus cell is blank.'],
  ['Teacher PIN (optional)', '', 'If filled in, the teacher page asks for this PIN before it opens.'],
  ['Benchmark Caveat', DEFAULT_CAVEAT, 'Text shown beside the WCPM tier. Keep it honest about where your benchmark comes from.']
];
var DEFAULT_TIERS = [
  { min: 0, label: 'Emerging', note: 'Needs focused fluency support.' },
  { min: 50, label: 'Developing', note: 'Building toward the expected range.' },
  { min: 80, label: 'Within Expected Range', note: 'Typical pace for this level.' },
  { min: 110, label: 'Above Expected Range', note: 'Strong fluency for this level.' }
];
var PASSAGE_HEADERS = [
  'Passage Title', 'Short Name', 'Subtitle', 'Description',
  'Passage Text (paste here)', 'Question 1', 'Question 2', 'Question 3',
  'Key Terms (comma separated)', 'Sound Focus (On / Off, blank = default)',
  'Extra Focus Words (comma separated)', 'Hide this passage'
];
var ACCESS_HEADERS = ['Timestamp', 'Student ID', 'Full Name', 'Section', 'Method', 'Birthdate / Secret Code', 'Last Used', 'Teacher Note'];

// =====================================================================
// Spreadsheet + settings helpers
// =====================================================================
function getSS_() {
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active) return active;
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('This copy is not set up yet. Open the Google Sheet and choose Reading Aloud > Set Up.');
  return SpreadsheetApp.openById(id);
}

function toBool_(v) {
  var t = String(v).trim().toLowerCase();
  return v === true || t === 'true' || t === 'yes' || t === 'on';
}
function splitList_(v) {
  return String(v === null || v === undefined ? '' : v).split(/[,;\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
}
function countWords_(text) {
  return String(text || '').split(/\s+/).filter(function (w) { return w !== ''; }).length;
}

function getSettings_() {
  var sheet = getSS_().getSheetByName(TAB_SETTINGS);
  if (!sheet) throw new Error('This copy is not set up yet. Open the Google Sheet and choose Reading Aloud > Set Up.');
  var map = {};
  SETTINGS_DEFS.forEach(function (d) { map[d[0]] = d[1]; });
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]).trim();
      if (map.hasOwnProperty(k)) map[k] = r[1];
    });
  }
  var sections = splitList_(map['Sections']).filter(function (x, i, a) { return a.indexOf(x) === i; });
  if (!sections.length) sections = ['Class A'];
  var m = String(map['Student Access Method']).trim().toLowerCase();
  var method = 'birthdate';
  if (m.indexOf('code') !== -1 || m.indexOf('secret') !== -1) method = 'code';
  else if (m.indexOf('id') !== -1 || m.indexOf('none') !== -1) method = 'none';
  return {
    appTitle: String(map['App Title'] || '').trim() || 'Reading Aloud Assessment',
    subject: String(map['Subject / Course'] || '').trim(),
    schoolName: String(map['School / Teacher Name'] || '').trim(),
    sections: sections,
    showLeaderboard: toBool_(map['Show Leaderboard To Students']),
    accessMethod: method,
    idLength: Math.max(0, parseInt(map['Student ID Length'], 10) || 0),
    defaultSound: toBool_(map['Default Sound Focus']),
    pin: String(map['Teacher PIN (optional)'] || '').trim(),
    caveat: String(map['Benchmark Caveat'] || '').trim() || DEFAULT_CAVEAT
  };
}

function getBenchmarks_() {
  var tiers = [];
  var sheet = getSS_().getSheetByName(TAB_BENCH);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function (r) {
      var min = Number(r[0]);
      var label = String(r[1] || '').trim();
      if (r[0] !== '' && !isNaN(min) && label) tiers.push({ min: min, label: label, note: String(r[2] || '') });
    });
  }
  if (!tiers.length) tiers = DEFAULT_TIERS.slice();
  tiers.sort(function (a, b) { return a.min - b.min; });
  return tiers;
}

function publicConfig_(s) {
  return {
    appTitle: s.appTitle, subject: s.subject, schoolName: s.schoolName, sections: s.sections,
    showLeaderboard: s.showLeaderboard, accessMethod: s.accessMethod, idLength: s.idLength,
    caveat: s.caveat
  };
}

function sheetNameForSection_(section) {
  return (RESULTS_PREFIX + String(section).replace(/[\[\]\*\?:\/\\]/g, '-')).slice(0, 100);
}
function isValidSection_(section, s) { return s.sections.indexOf(section) !== -1; }

function getResultsSheet_(section) {
  var ss = getSS_();
  var name = sheetNameForSection_(section);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADER_ROW);
    sheet.setFrozenRows(1);
    formatResultsSheet_(sheet);
  } else {
    ensureHeaderRow_(sheet);
  }
  return sheet;
}

/** Student IDs are stored as plain text so IDs such as 01234 keep their leading zero. */
function formatResultsSheet_(sheet) {
  sheet.getRange(2, COL_STUDENT_ID, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
}

function ensureHeaderRow_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
    sheet.setFrozenRows(1);
    return;
  }
  var width = Math.max(sheet.getLastColumn(), HEADER_ROW.length);
  var existing = sheet.getRange(1, 1, 1, width).getValues()[0];
  var changed = false;
  for (var i = 0; i < HEADER_ROW.length; i++) {
    if (existing[i] === '' || existing[i] === null || existing[i] === undefined) {
      existing[i] = HEADER_ROW[i];
      changed = true;
    }
  }
  if (changed) sheet.getRange(1, 1, 1, existing.length).setValues([existing]);
  if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
}

// =====================================================================
// Passages (read from the Passages tab) + automatic sound-focus detection
// =====================================================================
var ED_STOP = ['need', 'needs', 'speed', 'feed', 'seed', 'bleed', 'indeed', 'proceed', 'succeed', 'exceed',
  'breed', 'weed', 'heed', 'hundred', 'sacred', 'naked', 'wicked', 'kindred', 'shed', 'sled', 'bred', 'shred'];
var S_STOP = ['this', 'was', 'his', 'its', 'has', 'does', 'always', 'perhaps', 'across', 'less', 'plus', 'yes', 'thus', 'whereas'];

function edSound_(w) {
  var stem = w.slice(0, -2);
  if (/[td]$/.test(stem)) return '\u026ad';
  if (/(p|k|f|x|ch|sh|ss|gh|c)$/.test(stem)) return 't';
  return 'd';
}

/**
 * Finds the words worth listening for, using simple spelling patterns. It is
 * close but not perfect (for example "based" is guessed as /d/). Teachers can
 * add or fix words with the "Extra Focus Words" column of the Passages tab.
 */
function detectFocus_(text) {
  var seen = {}, res = { lr: [], wv: [], s: [], ed: [] };
  text.split(/\s+/).forEach(function (raw) {
    var shown = raw.replace(/^[^A-Za-z]+|[^A-Za-z']+$/g, '');
    var w = shown.toLowerCase().replace(/[^a-z]/g, '');
    if (w.length < 4 || seen[w]) return;
    seen[w] = true;
    if (/^[lr]/.test(w) || /^(bl|br|cl|cr|dr|fl|fr|gl|gr|pl|pr|sl|tr|thr|scr|spr|str|spl|shr)/.test(w) || /[aeiou][lr][aeiou]/.test(w)) res.lr.push(shown);
    if (/^[wv]/.test(w) || /v/.test(w)) res.wv.push(shown);
    if (/s$/.test(w) && !/(ss|us|is|ous)$/.test(w) && S_STOP.indexOf(w) === -1) res.s.push(shown);
    if (/ed$/.test(w) && ED_STOP.indexOf(w) === -1) res.ed.push(shown + ' /' + edSound_(w) + '/');
  });
  return res;
}

function getPassages_(s) {
  var sheet = getSS_().getSheetByName(TAB_PASSAGES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PASSAGE_HEADERS.length).getValues();
  var out = [], used = {};
  rows.forEach(function (r, i) {
    var text = String(r[4] || '').replace(/\s+/g, ' ').trim();
    if (!text || toBool_(r[11])) return;
    var title = String(r[0] || '').trim() || ('Passage ' + (i + 1));
    var id = title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || ('passage-' + (i + 2));
    if (used[id]) id += '-' + (i + 2);
    used[id] = true;
    var flag = String(r[9] || '').trim().toLowerCase();
    var sound = s.defaultSound;
    if (flag === 'on' || flag === 'yes' || flag === 'true') sound = true;
    else if (flag === 'off' || flag === 'no' || flag === 'false') sound = false;
    var focus = sound ? detectFocus_(text) : { lr: [], wv: [], s: [], ed: [] };
    focus.custom = splitList_(r[10]);
    out.push({
      id: id,
      short: String(r[1] || '').trim() || (title.length > 28 ? title.slice(0, 27) + '\u2026' : title),
      title: title,
      subtitle: String(r[2] || '').trim(),
      blurb: String(r[3] || '').trim(),
      text: text,
      questions: [r[5], r[6], r[7]].map(function (q) { return String(q || '').trim(); }).filter(Boolean),
      terms: splitList_(r[8]),
      soundFocus: sound,
      focus: focus,
      wordCount: countWords_(text)
    });
  });
  return out;
}

// =====================================================================
// Scoring (unchanged rules from the original system)
// =====================================================================
var TRUE_ERROR_CATEGORIES = ['Mispronunciation', 'Substitution', 'Omission', 'Hesitation'];
var NON_ERROR_CATEGORY = 'Self-Correction';

function countTrueErrors_(tally) {
  if (!tally || typeof tally !== 'object') return 0;
  var total = 0;
  TRUE_ERROR_CATEGORIES.forEach(function (c) {
    var n = Number(tally[c]);
    if (!isNaN(n) && n > 0) total += n;
  });
  return total;
}
function countSelfCorrections_(tally) {
  if (!tally || typeof tally !== 'object') return 0;
  var n = Number(tally[NON_ERROR_CATEGORY]);
  return isNaN(n) || n < 0 ? 0 : n;
}
function computeWcpm_(correctWords, durationSeconds) {
  var seconds = Number(durationSeconds), correct = Number(correctWords);
  if (!seconds || seconds <= 0 || isNaN(correct) || correct < 0) return 0;
  return Math.round(correct / (seconds / 60));
}

var FLUENCY_BAND_LABELS = ['A1', 'A2', 'B1', 'B2', 'C1'];
var FLUENCY_BAND_ACCURACY_CUTOFFS = [0, 70, 85, 95, 98];
var FLUENCY_BAND_WCPM_CUTOFFS = [0, 30, 60, 100, 130];

function bandIndexFor_(value, cutoffs) {
  var idx = 0;
  for (var i = 0; i < cutoffs.length; i++) { if (value >= cutoffs[i]) idx = i; }
  return idx;
}
function computeReadingFluencyBand_(accuracyPercent, wcpm) {
  var accIdx = bandIndexFor_(Number(accuracyPercent) || 0, FLUENCY_BAND_ACCURACY_CUTOFFS);
  if (!wcpm || wcpm <= 0) return FLUENCY_BAND_LABELS[accIdx] + ' (reading-only est., accuracy only, not timed)';
  var overall = Math.min(accIdx, bandIndexFor_(Number(wcpm) || 0, FLUENCY_BAND_WCPM_CUTOFFS));
  return FLUENCY_BAND_LABELS[overall] + ' (reading-only est.)';
}
function computeWcpmTier_(wcpm, s, tiers) {
  if (!wcpm || wcpm <= 0) {
    return { label: 'Not Timed', note: 'This session was not timed, so a speed tier could not be determined.', caveat: s.caveat };
  }
  var chosen = tiers[0];
  tiers.forEach(function (t) { if (wcpm >= t.min) chosen = t; });
  return { label: chosen.label, note: chosen.note, caveat: s.caveat };
}

// =====================================================================
// Web app entry points
// =====================================================================
function doGet(e) {
  var params = (e && e.parameter) || {};
  var mode = String(params.mode || '').trim().toLowerCase();
  var settings;
  try {
    settings = getSettings_();
  } catch (err) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px;">' + String(err.message).replace(/</g, '&lt;') + '</p>');
  }
  var student = mode === 'student';
  return HtmlService.createTemplateFromFile(student ? 'student' : 'index')
    .evaluate()
    .setTitle(settings.appTitle + (student ? ' - Student' : ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =====================================================================
// Teacher functions (all require the teacher PIN when one is set)
// =====================================================================
function requireTeacher_(pin) {
  var s = getSettings_();
  if (s.pin && String(pin || '').trim() !== s.pin) throw new Error('PIN required. Please reload the page and enter your PIN.');
  return s;
}

function getTeacherBoot(pin) {
  try {
    var s = getSettings_();
    if (s.pin && String(pin || '').trim() !== s.pin) {
      return { success: false, needPin: true, appTitle: s.appTitle, message: pin ? 'That PIN is not correct.' : '' };
    }
    var config = publicConfig_(s);
    config.tiers = getBenchmarks_();
    return { success: true, config: config, passages: getPassages_(s) };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function isValidStudentId_(id, len) {
  id = String(id || '').trim();
  if (len > 0) return new RegExp('^\\d{' + len + '}$').test(id);
  return /^[A-Za-z0-9\-]{1,20}$/.test(id);
}

function submitReadingResult(pin, payload) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  try {
    var s = requireTeacher_(pin);
    payload = payload || {};
    var studentId = String(payload.studentId || '').trim();
    var name = String(payload.name || '').trim();
    var section = payload.section;
    var sessionNotes = String(payload.sessionNotes || '').trim();
    if (!isValidStudentId_(studentId, s.idLength)) {
      return { success: false, message: 'Invalid Student ID' + (s.idLength ? ' (it must be exactly ' + s.idLength + ' digits).' : '.') };
    }
    if (!name) return { success: false, message: 'Missing student name.' };
    if (!isValidSection_(section, s)) return { success: false, message: 'Invalid section.' };
    var passage = getPassages_(s).filter(function (p) { return p.id === payload.passageId; })[0];
    if (!passage) return { success: false, message: 'Missing or unknown reading passage. Check the Passages tab.' };

    try { lockAcquired = lock.tryLock(SUBMIT_LOCK_WAIT_MS); } catch (lockErr) { lockAcquired = false; }
    if (!lockAcquired) {
      return { success: false, retryable: true, message: 'The server is busy recording another session. Please try again in a moment.' };
    }

    var sheet = getResultsSheet_(section);
    var tiers = getBenchmarks_();
    var totalWords = passage.wordCount;
    var tally = payload.tally || {};
    var responses = payload.responses || [];
    var miscueCount = countTrueErrors_(tally);
    var accuracyPercent = totalWords ? Math.round(((totalWords - miscueCount) / totalWords) * 100) : 0;
    var selfCorrections = countSelfCorrections_(tally);
    var durationSeconds = Math.max(0, Math.round(Number(payload.durationSeconds) || 0));
    var wcpm = computeWcpm_(Math.max(0, totalWords - miscueCount), durationSeconds);
    var band = computeReadingFluencyBand_(accuracyPercent, wcpm);
    var tier = computeWcpmTier_(wcpm, s, tiers);

    sheet.appendRow([
      new Date(), studentId, name, section, totalWords, miscueCount, accuracyPercent, sessionNotes,
      JSON.stringify(tally), JSON.stringify(responses), durationSeconds, wcpm, selfCorrections, band,
      '', '', '', '', '',
      tier.label, passage.id, passage.title,
      '', '', '', '', '',
      '', '', '', '',
      passage.text, passage.soundFocus
    ]);
    SpreadsheetApp.flush();

    return {
      success: true, accuracyPercent: accuracyPercent, totalWords: totalWords, miscueCount: miscueCount,
      tally: tally, rowNumber: sheet.getLastRow(), durationSeconds: durationSeconds, wcpm: wcpm,
      selfCorrections: selfCorrections, section: section, readingFluencyBand: band,
      wcpmTierLabel: tier.label, wcpmTierNote: tier.note, wcpmTierCaveat: tier.caveat,
      passageId: passage.id, passageTitle: passage.title
    };
  } catch (err) {
    return { success: false, retryable: true, message: 'Server error: ' + err.message };
  } finally {
    if (lockAcquired) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

function clampRubricScore_(value) {
  var n = Math.round(Number(value));
  return (isNaN(n) || n < 1 || n > 4) ? 0 : n;
}
function averageRubricScores_(scores) {
  var rated = scores.filter(function (n) { return n > 0; });
  if (!rated.length) return 0;
  return Math.round(rated.reduce(function (a, b) { return a + b; }, 0) / rated.length);
}

/** Saves one optional rubric (prosody, sound, or comp) for an already-saved session row. */
function saveRubric(pin, section, rowNumber, rubricId, scores) {
  try {
    var s = requireTeacher_(pin);
    var cols = { prosody: [COL_PROSODY_EXPRESSION, 4], sound: [COL_SOUND_LR, 4], comp: [COL_COMP_MAIN_IDEA, 3] }[rubricId];
    if (!cols) return { success: false, message: 'Unknown rubric.' };
    if (!isValidSection_(section, s)) return { success: false, message: 'Invalid section.' };
    var sheet = getSS_().getSheetByName(sheetNameForSection_(section));
    if (!sheet || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      return { success: false, message: 'That session could not be found -- it may have been moved or deleted on the sheet.' };
    }
    var clamped = [];
    for (var i = 0; i < cols[1]; i++) clamped.push(clampRubricScore_((scores || [])[i]));
    var overall = averageRubricScores_(clamped);
    sheet.getRange(rowNumber, cols[0], 1, cols[1] + 1).setValues([clamped.concat([overall])]);
    SpreadsheetApp.flush();
    return { success: true, scores: clamped, overallLevel: overall };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// =====================================================================
// Reading sessions back out (leaderboard + one-session detail)
// =====================================================================
function readAllSessions_(s) {
  var ss = getSS_(), out = [];
  s.sections.forEach(function (section) {
    var sheet = ss.getSheetByName(sheetNameForSection_(section));
    if (!sheet || sheet.getLastRow() < 2) return;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_ROW.length).getValues().forEach(function (row, i) {
      if (!row[COL_STUDENT_ID - 1]) return;
      var title = String(row[COL_PASSAGE_TITLE - 1] || '');
      out.push({
        timestamp: row[COL_TIMESTAMP - 1] ? new Date(row[COL_TIMESTAMP - 1]).toISOString() : '',
        rowNumber: i + 2,
        studentId: String(row[COL_STUDENT_ID - 1]).trim(),
        name: String(row[COL_NAME - 1]).trim(),
        section: String(row[COL_SECTION - 1]).trim() || section,
        totalWords: Number(row[COL_TOTAL_WORDS - 1]) || 0,
        miscueCount: Number(row[COL_MISCUE_COUNT - 1]) || 0,
        accuracyPercent: Number(row[COL_ACCURACY - 1]) || 0,
        durationSeconds: Number(row[COL_DURATION_SECONDS - 1]) || 0,
        wcpm: Number(row[COL_WCPM - 1]) || 0,
        passageTitle: title,
        passageShort: title.length > 28 ? title.slice(0, 27) + '\u2026' : title
      });
    });
  });
  return out;
}

function buildLeaderboard_(s) {
  var entries = readAllSessions_(s);
  entries.sort(function (a, b) {
    if (b.accuracyPercent !== a.accuracyPercent) return b.accuracyPercent - a.accuracyPercent;
    if (a.miscueCount !== b.miscueCount) return a.miscueCount - b.miscueCount;
    return new Date(a.timestamp) - new Date(b.timestamp);
  });
  return entries.slice(0, 100).map(function (e, idx) {
    e.rank = idx + 1;
    return e;
  });
}

function getLeaderboard(pin) {
  try {
    var s = requireTeacher_(pin);
    return { success: true, leaderboard: buildLeaderboard_(s) };
  } catch (err) {
    return { success: false, message: err.message, leaderboard: [] };
  }
}

function buildDetail_(sheet, rowNumber, s) {
  var row = sheet.getRange(rowNumber, 1, 1, HEADER_ROW.length).getValues()[0];
  var tally = {}, responses = [];
  try { tally = JSON.parse(row[8] || '{}'); } catch (e1) { tally = {}; }
  try { responses = JSON.parse(row[9] || '[]'); } catch (e2) { responses = []; }
  var wcpm = Number(row[COL_WCPM - 1]) || 0;
  var tier = computeWcpmTier_(wcpm, s, getBenchmarks_());
  var passageId = String(row[COL_PASSAGE_ID - 1] || '').trim();
  var text = String(row[COL_SNAPSHOT - 1] || '');
  if (!text) {
    var p = getPassages_(s).filter(function (x) { return x.id === passageId; })[0];
    text = p ? p.text : '';
  }
  var n = function (c) { return Number(row[c - 1]) || 0; };
  return {
    success: true,
    appTitle: s.appTitle, subject: s.subject, schoolName: s.schoolName,
    timestamp: row[0] ? new Date(row[0]).toISOString() : '',
    studentId: String(row[COL_STUDENT_ID - 1]).trim(),
    name: String(row[COL_NAME - 1]).trim(),
    section: String(row[COL_SECTION - 1]).trim(),
    totalWords: n(COL_TOTAL_WORDS), miscueCount: n(COL_MISCUE_COUNT), accuracyPercent: n(COL_ACCURACY),
    sessionNotes: String(row[COL_SESSION_NOTES - 1] || ''),
    tally: tally, responses: responses,
    durationSeconds: n(COL_DURATION_SECONDS), wcpm: wcpm, selfCorrections: n(COL_SELF_CORRECTIONS),
    readingFluencyBand: String(row[COL_READING_FLUENCY_BAND - 1] || ''),
    prosodyExpression: n(15), prosodyPhrasing: n(16), prosodySmoothness: n(17), prosodyPunctuation: n(18),
    prosodyOverallLevel: n(COL_PROSODY_OVERALL),
    wcpmTierLabel: tier.label, wcpmTierNote: tier.note, wcpmTierCaveat: tier.caveat,
    passageId: passageId, passageTitle: String(row[COL_PASSAGE_TITLE - 1] || ''), passageText: text,
    soundFocusEnabled: toBool_(row[COL_SOUND_ON - 1]),
    soundLR: n(23), soundWV: n(24), soundFinals: n(25), soundEd: n(26), soundOverallLevel: n(COL_SOUND_OVERALL),
    compMainIdea: n(28), compDetails: n(29), compTerms: n(30), compOverallLevel: n(COL_COMP_OVERALL)
  };
}

function loadDetail_(s, section, rowNumber) {
  if (!isValidSection_(section, s)) return { success: false, message: 'Invalid section.' };
  var sheet = getSS_().getSheetByName(sheetNameForSection_(section));
  rowNumber = Number(rowNumber);
  if (!sheet || !rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return { success: false, message: 'That session could not be found -- it may have been moved or deleted on the sheet.' };
  }
  return buildDetail_(sheet, rowNumber, s);
}

function getSessionDetail(pin, section, rowNumber) {
  try {
    return loadDetail_(requireTeacher_(pin), section, rowNumber);
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// =====================================================================
// Student functions (no teacher access; each one checks the Settings tab)
// =====================================================================
function getStudentBoot() {
  try {
    return { success: true, config: publicConfig_(getSettings_()) };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getStudentLeaderboard() {
  try {
    var s = getSettings_();
    if (!s.showLeaderboard) return { success: true, hidden: true, leaderboard: [] };
    return { success: true, leaderboard: buildLeaderboard_(s) };
  } catch (err) {
    return { success: false, message: err.message, leaderboard: [] };
  }
}

/** Printing from the public leaderboard works only while the leaderboard is switched on. */
function getPublicSessionDetail(section, rowNumber) {
  try {
    var s = getSettings_();
    if (!s.showLeaderboard) return { success: false, message: 'The leaderboard is turned off. Use "Find my report" instead.' };
    return loadDetail_(s, section, rowNumber);
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getAccessSheet_() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(TAB_ACCESS);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_ACCESS);
    sheet.appendRow(ACCESS_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(2, 2, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
    sheet.getRange(2, 6, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
  }
  return sheet;
}

function normalizeSecret_(method, v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var t = String(v === null || v === undefined ? '' : v).trim();
  return method === 'code' ? t.toLowerCase() : t;
}

function validSecretFormat_(method, v) {
  v = String(v || '').trim();
  if (method === 'code') return v.length >= 4 && v.length <= 20;
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
  if (!m) return false;
  var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  var dt = new Date(y, mo - 1, d);
  return y >= 1950 && y <= new Date().getFullYear() && dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d;
}

function findAccessRow_(id) {
  var sheet = getAccessSheet_();
  if (sheet.getLastRow() < 2) return null;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim() === id) {
      return { rowIndex: i + 2, method: String(data[i][4]).trim().toLowerCase(), secret: data[i][5] };
    }
  }
  return null;
}

function isLocked_(id) { return Number(CacheService.getScriptCache().get('fail_' + id) || 0) >= 5; }
function noteFail_(id) {
  var cache = CacheService.getScriptCache();
  cache.put('fail_' + id, String(Number(cache.get('fail_' + id) || 0) + 1), 900);
  Utilities.sleep(700);
}

function sessionsFor_(id, s) {
  return readAllSessions_(s).filter(function (e) { return e.studentId === id; })
    .sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
}

/** Returns { ok:true } when the ID + secret match, otherwise { ok:false, message }. */
function verifyStudent_(id, secret, s) {
  if (!isValidStudentId_(id, s.idLength)) return { ok: false, message: 'Please enter a valid Student ID.' };
  if (isLocked_(id)) return { ok: false, message: 'Too many wrong tries. Please ask your teacher.' };
  if (s.accessMethod === 'none') return { ok: true };
  var acc = findAccessRow_(id);
  if (!acc || acc.method !== s.accessMethod) return { ok: false, message: 'Please set up your access first.' };
  if (normalizeSecret_(s.accessMethod, secret) !== normalizeSecret_(s.accessMethod, acc.secret)) {
    noteFail_(id);
    return { ok: false, message: 'That does not match. Please try again.' };
  }
  CacheService.getScriptCache().remove('fail_' + id);
  return { ok: true, acc: acc };
}

function studentLookup(id) {
  try {
    var s = getSettings_();
    id = String(id || '').trim();
    if (!isValidStudentId_(id, s.idLength)) return { success: false, message: 'Please enter a valid Student ID.' };
    if (isLocked_(id)) return { success: false, message: 'Too many wrong tries. Please ask your teacher.' };
    if (!sessionsFor_(id, s).length) {
      Utilities.sleep(500);
      return { success: true, status: 'none', message: 'No reports found for that ID yet. Ask your teacher.' };
    }
    if (s.accessMethod === 'none') return { success: true, status: 'open', method: 'none' };
    var acc = findAccessRow_(id);
    var ready = acc && acc.method === s.accessMethod;
    return { success: true, status: ready ? 'ready' : 'needs_setup', method: s.accessMethod };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function studentRegister(id, secret, confirmSecret) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var s = getSettings_();
    id = String(id || '').trim();
    if (s.accessMethod === 'none') return { success: false, message: 'No setup is needed for this class.' };
    if (!isValidStudentId_(id, s.idLength)) return { success: false, message: 'Please enter a valid Student ID.' };
    var mine = sessionsFor_(id, s);
    if (!mine.length) return { success: false, message: 'No reports found for that ID yet. Ask your teacher.' };
    if (normalizeSecret_(s.accessMethod, secret) !== normalizeSecret_(s.accessMethod, confirmSecret)) {
      return { success: false, message: 'The two entries do not match. Please type them again.' };
    }
    if (!validSecretFormat_(s.accessMethod, secret)) {
      return { success: false, message: s.accessMethod === 'code' ? 'Your secret code must be 4 to 20 characters.' : 'Please choose a real birthdate.' };
    }
    locked = lock.tryLock(15000);
    if (!locked) return { success: false, message: 'The server is busy. Please try again.' };
    var existing = findAccessRow_(id);
    if (existing && existing.method === s.accessMethod) {
      return { success: false, message: 'This ID is already set up. If you forgot your ' + (s.accessMethod === 'code' ? 'secret code' : 'birthdate entry') + ', ask your teacher.' };
    }
    var sheet = getAccessSheet_();
    var rowIndex = existing ? existing.rowIndex : sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 2).setNumberFormat('@');
    sheet.getRange(rowIndex, 6).setNumberFormat('@');
    sheet.getRange(rowIndex, 1, 1, 8).setValues([[
      new Date(), id, mine[0].name, mine[0].section, s.accessMethod, normalizeSecret_(s.accessMethod, secret), '', ''
    ]]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

function studentListReports(id, secret) {
  try {
    var s = getSettings_();
    id = String(id || '').trim();
    var v = verifyStudent_(id, secret, s);
    if (!v.ok) return { success: false, message: v.message };
    if (v.acc) getAccessSheet_().getRange(v.acc.rowIndex, 7).setValue(new Date());
    var reports = sessionsFor_(id, s).map(function (e) {
      return {
        section: e.section, rowNumber: e.rowNumber, timestamp: e.timestamp, passageTitle: e.passageTitle,
        accuracyPercent: e.accuracyPercent, wcpm: e.wcpm, miscueCount: e.miscueCount, totalWords: e.totalWords,
        name: e.name
      };
    });
    return { success: true, reports: reports };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function studentGetReport(id, secret, section, rowNumber) {
  try {
    var s = getSettings_();
    id = String(id || '').trim();
    var v = verifyStudent_(id, secret, s);
    if (!v.ok) return { success: false, message: v.message };
    var detail = loadDetail_(s, section, rowNumber);
    if (!detail.success) return detail;
    if (detail.studentId !== id) return { success: false, message: 'That report does not belong to this Student ID.' };
    return detail;
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// =====================================================================
// Menu + one-click setup (this is what a teacher runs after "Make a copy")
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Reading Aloud')
    .addItem('1. Set Up (run this first)', 'setupSheets')
    .addItem('2. Show my teacher and student links', 'showLinks')
    .addToUi();
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  // ----- Settings -----
  var st = ss.getSheetByName(TAB_SETTINGS);
  if (!st) {
    st = ss.insertSheet(TAB_SETTINGS);
    st.appendRow(['Setting', 'Value', 'What it does']);
    st.setFrozenRows(1);
  }
  var existing = {};
  if (st.getLastRow() > 1) {
    st.getRange(2, 1, st.getLastRow() - 1, 1).getValues().forEach(function (r, i) { existing[String(r[0]).trim()] = i + 2; });
  }
  SETTINGS_DEFS.forEach(function (d) {
    var row = existing[d[0]];
    if (!row) {
      row = st.getLastRow() + 1;
      st.getRange(row, 1).setValue(d[0]);
      st.getRange(row, 3).setValue(d[2]);
      if (d[0] === 'Teacher PIN (optional)') st.getRange(row, 2).setNumberFormat('@');
      if (typeof d[1] === 'boolean') {
        st.getRange(row, 2).insertCheckboxes();
        st.getRange(row, 2).setValue(d[1]);
      } else {
        st.getRange(row, 2).setValue(d[1]);
      }
      if (d[0] === 'Student Access Method') {
        st.getRange(row, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Birthdate', 'Secret code', 'ID only'], true).build());
      }
    }
  });
  st.setColumnWidth(1, 240); st.setColumnWidth(2, 260); st.setColumnWidth(3, 560);
  st.getRange('A:A').setFontWeight('bold');

  // ----- Passages -----
  var ps = ss.getSheetByName(TAB_PASSAGES);
  if (!ps) {
    ps = ss.insertSheet(TAB_PASSAGES);
    ps.getRange(1, 1, 1, PASSAGE_HEADERS.length).setValues([PASSAGE_HEADERS]).setFontWeight('bold').setBackground('#1f4e8c').setFontColor('#ffffff').setWrap(true);
    ps.setFrozenRows(1);
    ps.getRange(2, 1, 1, PASSAGE_HEADERS.length).setValues([[
      'Mesopotamia: The Land Between the Rivers', 'Mesopotamia', 'The Tigris, the Euphrates and the Fertile Crescent',
      'Sample passage. Replace it with your own, or paste new rows underneath.',
      'Mesopotamia means the land between the rivers. It lay between the Tigris and Euphrates Rivers, in a region called the Fertile Crescent. Each spring the rivers flooded and left rich soil, so farmers raised wheat and barley and watered their fields with canals. Villages grew into large cities such as Ur, Uruk, and Babylon. The Sumerians created one of the world\'s first writing systems, called cuneiform, which they pressed into wet clay with reeds. Later, a Babylonian ruler named Hammurabi recorded his laws on a stone pillar so everyone would know the rules. Traders traveled along the rivers to exchange grain, wool, and metal. Because of these achievements, historians often call Mesopotamia the cradle of civilization.',
      'What does the word Mesopotamia mean, and which two rivers surround it?',
      'How did the yearly floods and canals help farmers in this region?',
      'Why do historians call Mesopotamia the cradle of civilization?',
      'Tigris, Euphrates, Fertile Crescent, cuneiform, Hammurabi, civilization', '', 'cuneiform, Hammurabi', false
    ]]);
    ps.getRange(2, 12, 200, 1).insertCheckboxes();
    ps.getRange(2, 12).setValue(false);
    ps.setColumnWidths(1, 4, 190); ps.setColumnWidth(5, 520); ps.setColumnWidths(6, 3, 220);
    ps.setColumnWidths(9, 4, 180);
    ps.getRange(2, 1, 200, PASSAGE_HEADERS.length).setWrap(true).setVerticalAlignment('top');
    ps.getRange(2, 10, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['On', 'Off'], true).setAllowInvalid(true).build());
  }

  // ----- Benchmarks -----
  var bs = ss.getSheetByName(TAB_BENCH);
  if (!bs) {
    bs = ss.insertSheet(TAB_BENCH);
    bs.getRange(1, 1, 1, 3).setValues([['Minimum WCPM', 'Tier Label', 'Note']]).setFontWeight('bold');
    bs.setFrozenRows(1);
    bs.getRange(2, 1, DEFAULT_TIERS.length, 3).setValues(DEFAULT_TIERS.map(function (t) { return [t.min, t.label, t.note]; }));
    bs.getRange(7, 1).setValue('Edit the numbers and labels to match your students. Add or delete rows freely. The default tiers are a conservative EFL working benchmark for lower-secondary students, not an official norm.');
    bs.setColumnWidth(1, 130); bs.setColumnWidth(2, 220); bs.setColumnWidth(3, 340);
  }

  // ----- Student Access + results tabs -----
  getAccessSheet_();
  var s = getSettings_();
  s.sections.forEach(function (section) { formatResultsSheet_(getResultsSheet_(section)); });

  // ----- Start Here -----
  if (!ss.getSheetByName(TAB_START)) {
    var sh = ss.insertSheet(TAB_START, 0);
    var lines = [
      ['READING ALOUD ASSESSMENT: START HERE'],
      [''],
      ['1. Menu Reading Aloud > Set Up. Approve the Google permission prompt. If Google says the app is not verified, choose Advanced > Go to project (unsafe). This is normal for a script you copied yourself.'],
      ['2. Open your teacher link and click Manage. In Settings type your App Title, Subject, and class names (separated by commas), and decide whether students may see the leaderboard. Set a Teacher PIN so students cannot open Manage. (The Settings tab in this Sheet holds the same values.)'],
      ['3. In Manage > Passages click Add passage, paste your reading text, and add a title, optional questions and key terms. You can hide, edit or delete a reading there. (The Passages tab in this Sheet holds the same readings.)'],
      ['4. In the script editor choose Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Click Deploy and copy the link. (Do this once. After later code changes use Deploy > Manage deployments.)'],
      ['5. Menu Reading Aloud > Show my teacher and student links. Keep the teacher link for yourself and give students the student link.'],
      [''],
      ['STUDENTS FORGOT THEIR BIRTHDATE OR CODE?'],
      ['Open your teacher link, click Manage, then Students. Each student\'s ID, name, and what they typed are listed there. Tell them what it says, or click Reset so they can set it up again. (The Student Access tab in this Sheet shows the same list.)'],
      ['If you change the Student Access Method in Settings, clear the Student Access rows so students can set up again.'],
      [''],
      ['GOOD TO KNOW'],
      ['Changes to Settings and Passages show up the next time a page is opened or reloaded. No redeploy is needed.'],
      ['Each saved session stores a copy of the passage text it used, so old reports stay correct even if you later edit or delete a passage.'],
      ['Sound Focus (L/R, W/V, final S/ES, -ED) suits English pronunciation practice. Turn it Off for a passage in the Passages tab if it does not help.'],
      ['Use passages of roughly 80 to 200 words. Student results are in the Results tabs. Please do not delete or reorder rows there, because reports find sessions by row.']
    ];
    sh.getRange(1, 1, lines.length, 1).setValues(lines).setWrap(true).setVerticalAlignment('top');
    sh.setColumnWidth(1, 760);
    sh.getRange('A1').setFontSize(16).setFontWeight('bold');
    sh.getRange('A9').setFontWeight('bold'); sh.getRange('A13').setFontWeight('bold');
  }

  var status = { success: true, spreadsheetName: ss.getName(), sections: s.sections };
  Logger.log(JSON.stringify(status));
  try { SpreadsheetApp.getUi().alert('Set up complete. Next: fill in Settings and Passages, then deploy the web app (see the Start Here tab).'); } catch (e) {}
  return status;
}

function showLinks() {
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  var html = url
    ? '<div style="font-family:sans-serif;font-size:14px"><p><b>Teacher link</b><br><input style="width:100%" readonly value="' + url + '"></p>' +
      '<p><b>Student link</b><br><input style="width:100%" readonly value="' + url + '?mode=student"></p>' +
      '<p style="color:#555">Copy each link. The student link is the one to share with students.</p></div>'
    : '<div style="font-family:sans-serif;font-size:14px"><p>No web app is deployed yet.</p><p>In the script editor choose <b>Deploy > New deployment > Web app</b>, set <b>Execute as: Me</b> and <b>Who has access: Anyone</b>, click Deploy, then run this menu item again.</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(280), 'Your links');
}

// =====================================================================
// Manage screen (teacher-only): edit Settings, Passages, Benchmarks and
// Student Access from the web app. The Sheet tabs remain the storage, so
// editing there still works and both always agree.
// =====================================================================
var METHOD_LABELS = { birthdate: 'Birthdate', code: 'Secret code', none: 'ID only' };

function getPassageRows_() {
  var sheet = getSS_().getSheetByName(TAB_PASSAGES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PASSAGE_HEADERS.length).getValues();
  var out = [];
  rows.forEach(function (r, i) {
    if (!String(r[0]).trim() && !String(r[4]).trim()) return;
    var flag = String(r[9] || '').trim().toLowerCase();
    out.push({
      rowNumber: i + 2, title: String(r[0]), shortName: String(r[1]), subtitle: String(r[2]), description: String(r[3]),
      text: String(r[4]), q1: String(r[5]), q2: String(r[6]), q3: String(r[7]), terms: String(r[8]),
      sound: flag === 'on' ? 'On' : (flag === 'off' ? 'Off' : ''), extra: String(r[10]), hidden: toBool_(r[11]),
      wordCount: countWords_(r[4])
    });
  });
  return out;
}

function getManageData(pin) {
  try {
    var s = requireTeacher_(pin);
    return {
      success: true,
      settings: {
        appTitle: s.appTitle, subject: s.subject, schoolName: s.schoolName, sections: s.sections.join(', '),
        showLeaderboard: s.showLeaderboard, accessMethod: s.accessMethod, idLength: s.idLength,
        defaultSound: s.defaultSound, pin: s.pin, caveat: s.caveat
      },
      passages: getPassageRows_(),
      tiers: getBenchmarks_()
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function writeSetting_(sheet, key, value) {
  var def = null, row = 0;
  SETTINGS_DEFS.forEach(function (d) { if (d[0] === key) def = d; });
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
      if (String(r[0]).trim() === key) row = i + 2;
    });
  }
  if (!row) {
    row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(key);
    if (def) sheet.getRange(row, 3).setValue(def[2]);
  }
  var cell = sheet.getRange(row, 2);
  if (def && typeof def[1] === 'boolean') {
    cell.insertCheckboxes();
    cell.setValue(!!value);
  } else {
    if (key === 'Teacher PIN (optional)') cell.setNumberFormat('@');
    cell.setValue(value);
  }
}

function saveSettings(pin, o) {
  try {
    requireTeacher_(pin);
    o = o || {};
    var sections = splitList_(o.sections).filter(function (x, i, a) { return a.indexOf(x) === i; });
    if (!sections.length) return { success: false, message: 'Please enter at least one section (class name).' };
    var idLen = parseInt(o.idLength, 10);
    if (isNaN(idLen) || idLen < 0 || idLen > 20) return { success: false, message: 'Student ID length must be a number from 0 to 20.' };
    var method = METHOD_LABELS[o.accessMethod] ? o.accessMethod : 'birthdate';
    var sheet = getSS_().getSheetByName(TAB_SETTINGS);
    if (!sheet) return { success: false, message: 'The Settings tab is missing. Run Reading Aloud > Set Up in the Sheet.' };
    writeSetting_(sheet, 'App Title', String(o.appTitle || '').trim() || 'Reading Aloud Assessment');
    writeSetting_(sheet, 'Subject / Course', String(o.subject || '').trim());
    writeSetting_(sheet, 'School / Teacher Name', String(o.schoolName || '').trim());
    writeSetting_(sheet, 'Sections', sections.join(', '));
    writeSetting_(sheet, 'Show Leaderboard To Students', !!o.showLeaderboard);
    writeSetting_(sheet, 'Student Access Method', METHOD_LABELS[method]);
    writeSetting_(sheet, 'Student ID Length', idLen);
    writeSetting_(sheet, 'Default Sound Focus', !!o.defaultSound);
    writeSetting_(sheet, 'Teacher PIN (optional)', String(o.pin || '').trim());
    writeSetting_(sheet, 'Benchmark Caveat', String(o.caveat || '').trim() || DEFAULT_CAVEAT);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function savePassage(pin, p) {
  try {
    requireTeacher_(pin);
    p = p || {};
    var text = String(p.text || '').replace(/\s+/g, ' ').trim();
    var title = String(p.title || '').trim();
    if (!title) return { success: false, message: 'Please give the passage a title.' };
    if (!text) return { success: false, message: 'Please paste the passage text.' };
    var sheet = getSS_().getSheetByName(TAB_PASSAGES);
    if (!sheet) return { success: false, message: 'The Passages tab is missing. Run Reading Aloud > Set Up in the Sheet.' };
    var rowNumber = Number(p.rowNumber) || 0;
    if (rowNumber < 2 || rowNumber > Math.max(2, sheet.getLastRow())) {
      var last = 1;
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(function (r, i) {
          if (String(r[0]).trim() || String(r[4]).trim()) last = i + 2;
        });
      }
      rowNumber = last + 1;
    }
    var sound = p.sound === 'On' || p.sound === 'Off' ? p.sound : '';
    sheet.getRange(rowNumber, 1, 1, 11).setValues([[
      title, String(p.shortName || '').trim(), String(p.subtitle || '').trim(), String(p.description || '').trim(), text,
      String(p.q1 || '').trim(), String(p.q2 || '').trim(), String(p.q3 || '').trim(),
      String(p.terms || '').trim(), sound, String(p.extra || '').trim()
    ]]);
    sheet.getRange(rowNumber, 10).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['On', 'Off'], true).setAllowInvalid(true).build());
    sheet.getRange(rowNumber, 12).insertCheckboxes();
    sheet.getRange(rowNumber, 12).setValue(!!p.hidden);
    sheet.getRange(rowNumber, 1, 1, PASSAGE_HEADERS.length).setWrap(true).setVerticalAlignment('top');
    SpreadsheetApp.flush();
    return { success: true, rowNumber: rowNumber };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function deletePassage(pin, rowNumber) {
  try {
    requireTeacher_(pin);
    var sheet = getSS_().getSheetByName(TAB_PASSAGES);
    rowNumber = Number(rowNumber);
    if (!sheet || rowNumber < 2 || rowNumber > sheet.getLastRow()) return { success: false, message: 'That passage could not be found.' };
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function saveBenchmarks(pin, tiers) {
  try {
    requireTeacher_(pin);
    var clean = [];
    (tiers || []).forEach(function (t) {
      var min = Number(t.min), label = String(t.label || '').trim();
      if (t.min === '' || isNaN(min) || min < 0 || !label) return;
      clean.push({ min: min, label: label, note: String(t.note || '').trim() });
    });
    if (!clean.length) return { success: false, message: 'Please keep at least one tier with a minimum WCPM and a label.' };
    clean.sort(function (a, b) { return a.min - b.min; });
    var sheet = getSS_().getSheetByName(TAB_BENCH);
    if (!sheet) return { success: false, message: 'The Benchmarks tab is missing. Run Reading Aloud > Set Up in the Sheet.' };
    sheet.getRange(2, 1, Math.max(1, sheet.getMaxRows() - 1), 3).clearContent();
    sheet.getRange(2, 1, clean.length, 3).setValues(clean.map(function (t) { return [t.min, t.label, t.note]; }));
    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function getStudentAccessList(pin) {
  try {
    requireTeacher_(pin);
    var sheet = getAccessSheet_();
    if (sheet.getLastRow() < 2) return { success: true, students: [] };
    var students = [];
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues().forEach(function (r, i) {
      if (!String(r[1]).trim()) return;
      var method = String(r[4]).trim().toLowerCase();
      students.push({
        rowNumber: i + 2, studentId: String(r[1]).trim(), name: String(r[2]), section: String(r[3]),
        method: method, secret: normalizeSecret_(method, r[5]),
        lastUsed: r[6] ? new Date(r[6]).toISOString() : ''
      });
    });
    return { success: true, students: students };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/** Removes a student's access entry so they can choose a new birthdate or code. */
function resetStudentAccess(pin, rowNumber) {
  try {
    requireTeacher_(pin);
    var sheet = getAccessSheet_();
    rowNumber = Number(rowNumber);
    if (rowNumber < 2 || rowNumber > sheet.getLastRow()) return { success: false, message: 'That entry could not be found.' };
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}