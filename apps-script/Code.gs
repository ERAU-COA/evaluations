var WEIGHTS = { c1: 20, c2: 20, c3: 20, c4: 10, c5: 10, c6: 20 };
var ITEMS = {
  c1: ['c1a', 'c1b', 'c1c', 'c1d', 'c1e'],
  c2: ['c2a', 'c2b', 'c2c', 'c2d', 'c2e'],
  c3: ['c3a', 'c3b', 'c3c', 'c3d', 'c3e'],
  c4: ['c4a', 'c4b', 'c4c', 'c4d', 'c4e'],
  c5: ['c5a', 'c5b', 'c5c', 'c5d'],
  c6: ['c6a', 'c6b', 'c6c', 'c6d']
};
var TEAMS = { G1: 'Gordon & Romane', G2: 'Christian & Aaron', G3: 'Gal & Max', G4: 'Sukwoo & Ishaa' };
var TOPICS = {
  G1: 'Aircraft Cannibalization as an Inventory & Procurement Strategy',
  G2: 'Airbus A320 Family Production Ramp-Up Strategy',
  G3: 'Production Leveling; Supply Chain Planning at Embraer (Praetor 600E)',
  G4: 'GE’s Lean Evolution & Post-Split Aerospace Production Capacity'
};
var FIXED = { c3a: 1, c3c: 1, c3d: 1, c4c: 1, c4d: 1, c4e: 1, c5d: 1 };
var YESNO = { c4c: 1, c4d: 1 };
var FREQ3 = { c3a: 1, c3c: 1, c3d: 1 };
var SHEET = 'Responses';
var SHEET_ID = '1BEWXV9MhyCAvf2Z5hAN_ZdPAcw2rgsSYvH75pMt7NK4';
function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); }

function allItems_() {
  var out = [];
  Object.keys(ITEMS).forEach(function (c) { out = out.concat(ITEMS[c]); });
  return out;
}

function headers_() {
  return ['Date', 'Time', 'Phone local time', 'Time zone', 'Team', 'Group members', 'Topic']
    .concat(allItems_())
    .concat(['C1 (20)', 'C2 (20)', 'C3 (20)', 'C4 (10)', 'C5 (10)', 'C6 (20)', 'Total (100)'])
    .concat(['Comment C1', 'Comment C2', 'Comment C3', 'Comment C4', 'Comment C5', 'Comment C6', 'Overall comment'])
    .concat(['Seconds spent', 'Rating SD', 'Distinct values', 'Flat criteria', 'Straight-line flag',
             'Straight-line reason', 'Too-fast flag', 'Duplicate flag', 'Device token', 'Submission ID', 'Item order',
             'Attention check', 'Attention answer', 'Rating styles', 'Phone ISO time']);
}

function setup() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  var h = headers_();
  sh.getRange(1, 1, 1, sh.getMaxColumns()).clearContent();
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  sh.setFrozenRows(1);

  var sum = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  sum.clear();
  var totalCol = colLetter_(h.indexOf('Total (100)') + 1);
  var flagCol = colLetter_(h.indexOf('Straight-line flag') + 1);
  var attCol = colLetter_(h.indexOf('Attention check') + 1);
  var dupCol = colLetter_(h.indexOf('Duplicate flag') + 1);
  var rows = [['Team', 'Group members', 'Evaluations', 'Mean total', 'Mean total (clean)', 'Min', 'Max', 'Flagged rows']];
  Object.keys(TEAMS).forEach(function (t, i) {
    var r = i + 2;
    var teamCol = colLetter_(h.indexOf('Team') + 1);
    var tr = SHEET + '!' + teamCol + ':' + teamCol, tot = SHEET + '!' + totalCol + ':' + totalCol;
    var fl = SHEET + '!' + flagCol + ':' + flagCol, du = SHEET + '!' + dupCol + ':' + dupCol;
    var at = SHEET + '!' + attCol + ':' + attCol;
    rows.push([t, TEAMS[t],
      '=COUNTIF(' + tr + ',A' + r + ')',
      '=IFERROR(ROUND(AVERAGEIF(' + tr + ',A' + r + ',' + tot + '),1),"")',
      '=IFERROR(ROUND(AVERAGEIFS(' + tot + ',' + tr + ',A' + r + ',' + fl + ',"",' + du + ',"",' + at + ',"<>FAIL"),1),"")',
      '=IF(C' + r + '=0,"",MINIFS(' + tot + ',' + tr + ',A' + r + '))',
      '=IF(C' + r + '=0,"",MAXIFS(' + tot + ',' + tr + ',A' + r + '))',
      '=C' + r + '-COUNTIFS(' + tr + ',A' + r + ',' + fl + ',"",' + du + ',"",' + at + ',"<>FAIL")']);
  });
  sum.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sum.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sum.getRange(rows.length + 2, 1).setValue('Mean total (clean) leaves out straight-line, duplicate and failed-attention-check submissions.');
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function doGet(e) {
  var device = e && e.parameter && e.parameter.device;
  var sub = e && e.parameter && e.parameter.sub;
  if (!device) return json_({ ok: true, status: 'MSA 641 evaluation endpoint is running.' });
  var sh = ss_().getSheetByName(SHEET), h = headers_(), teams = {}, found = false;
  var last = sh ? sh.getLastRow() : 0;
  if (last > 1) {
    var devs = sh.getRange(2, h.indexOf('Device token') + 1, last - 1, 1).getValues();
    var tms = sh.getRange(2, h.indexOf('Team') + 1, last - 1, 1).getValues();
    for (var i = 0; i < devs.length; i++) if (String(devs[i][0]) === clean_(device)) teams[tms[i][0]] = 1;
    if (sub) {
      var subs = sh.getRange(2, h.indexOf('Submission ID') + 1, last - 1, 1).getValues();
      for (var k = 0; k < subs.length; k++) if (String(subs[k][0]) === clean_(sub)) found = true;
    }
  }
  return json_({ ok: true, teams: Object.keys(teams), found: found });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    var team = TEAMS[d.team] ? d.team : '';

    var ss = ss_();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) { setup(); sh = ss.getSheetByName(SHEET); }
    var h = headers_();

    var subCol = h.indexOf('Submission ID') + 1, devCol = h.indexOf('Device token') + 1;
    var last = sh.getLastRow();
    var dup = '';
    if (last > 1) {
      var subs = sh.getRange(2, subCol, last - 1, 1).getValues();
      for (var i = 0; i < subs.length; i++) if (String(subs[i][0]) === clean_(d.subId)) return json_({ ok: true, repeat: true });
      var devs = sh.getRange(2, devCol, last - 1, 1).getValues();
      var times = sh.getRange(2, h.indexOf('Phone ISO time') + 1, last - 1, 1).getValues();
      var t0 = new Date(d.localTime).getTime();
      for (var k = 0; k < devs.length; k++) {
        if (String(devs[k][0]) !== clean_(d.device)) continue;
        var tk = new Date(times[k][0]).getTime();
        if (!isNaN(t0) && !isNaN(tk) && Math.abs(t0 - tk) < 10 * 60 * 1000) dup = 'YES';
      }
    }

    var items = allItems_(), vals = [], crit = {}, total = 0;
    for (var j = 0; j < items.length; j++) {
      var v = Number(d.ratings && d.ratings[items[j]]);
      if (!(v >= 1 && v <= 5 && v === Math.round(v))) return json_({ ok: false, error: 'Missing or invalid rating.' });
      if (YESNO[items[j]] && v !== 1 && v !== 5) return json_({ ok: false, error: 'Invalid yes/no answer.' });
      if (FREQ3[items[j]] && v !== 1 && v !== 3 && v !== 5) return json_({ ok: false, error: 'Invalid how-often answer.' });
      vals.push(v);
    }
    Object.keys(ITEMS).forEach(function (c) {
      var s = 0; ITEMS[c].forEach(function (key) { s += Number(d.ratings[key]) - 1; });
      var raw = WEIGHTS[c] * (s / ITEMS[c].length) / 4;
      total += raw;
      crit[c] = Math.round(raw * 100) / 100;
    });
    total = Math.round(total * 10) / 10;

    var rv = items.filter(function (k) { return !FIXED[k]; }).map(function (k) { return Number(d.ratings[k]); });
    var mean = rv.reduce(function (a, b) { return a + b; }, 0) / rv.length;
    var sd = Math.sqrt(rv.reduce(function (a, b) { return a + (b - mean) * (b - mean); }, 0) / rv.length);
    var distinct = Object.keys(rv.reduce(function (o, x) { o[x] = 1; return o; }, {})).length;
    var flat = distinct === 1 ? 'YES' : '';
    var fast = Number(d.durationSec) < 60 ? 'YES' : '';
    var c = d.comments || {};
    var order = d.order ? Object.keys(d.order).map(function (k) { return k + ':' + d.order[k]; }).join(' | ') : '';

    var now = new Date(), stz = Session.getScriptTimeZone();
    var a = d.attn || {};
    var attnRes = (Number(a.value) === Number(a.target) && (Number(a.target) === 2 || Number(a.target) === 4)) ? 'PASS' : 'FAIL';
    var row = [Utilities.formatDate(now, stz, 'yyyy-MM-dd'), Utilities.formatDate(now, stz, 'HH:mm:ss'),
               clean_(d.localDisplay), clean_(d.tz), team, TEAMS[team] || '', TOPICS[team] || '']
      .concat(vals)
      .concat([crit.c1, crit.c2, crit.c3, crit.c4, crit.c5, crit.c6, total])
      .concat([c.c1, c.c2, c.c3, c.c4, c.c5, c.c6, c.overall].map(clean_))
      .concat([Number(d.durationSec) || '', Math.round(sd * 1000) / 1000, distinct, Number(d.flatCriteria) || 0, flat,
               clean_(d.flatReason), fast, dup, clean_(d.device), clean_(d.subId), clean_(order),
               attnRes, Number(a.value) ? (Number(a.value) + ' (asked ' + Number(a.target) + ', ' + clean_(a.style) + ')') : '',
               clean_(d.styles), clean_(d.localTime)]);
    sh.appendRow(row);
    return json_({ ok: true, total: total });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function clean_(s) {
  s = String(s || '').slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
