var WEIGHTS = { c1: 20, c2: 20, c3: 20, c4: 10, c5: 10, c6: 20 };
var ITEMS = {
  c1: ['c1a', 'c1b', 'c1c', 'c1d', 'c1e'],
  c2: ['c2a', 'c2b', 'c2c', 'c2d', 'c2e'],
  c3: ['c3a', 'c3b', 'c3c', 'c3d', 'c3e'],
  c4: ['c4a', 'c4b', 'c4c', 'c4d', 'c4e'],
  c5: ['c5a', 'c5b', 'c5c', 'c5d'],
  c6: ['c6a', 'c6b', 'c6c', 'c6d']
};
var GAP_MIN = 6;
var REPEAT_MIN = 8;
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
  return ['Date', 'Time', 'Phone local time', 'Time zone', 'Presentation', 'Presenters', 'Topic']
    .concat(allItems_())
    .concat(['C1 (20)', 'C2 (20)', 'C3 (20)', 'C4 (10)', 'C5 (10)', 'C6 (20)', 'Total (100)'])
    .concat(['Comment C1', 'Comment C2', 'Comment C3', 'Comment C4', 'Comment C5', 'Comment C6', 'Overall comment'])
    .concat(['Seconds spent', 'Rating SD', 'Distinct values', 'Flat criteria', 'Straight-line flag',
             'Straight-line reason', 'Too-fast flag', 'Duplicate flag', 'Device token', 'Submission ID', 'Item order',
             'Attention check', 'Attention answer', 'Rating styles', 'Phone ISO time',
             'Fingerprint ID', 'Device model', 'OS and browser', 'Screen', 'GPU', 'Device details',
             'First answer time', 'Last answer time', 'Minutes since previous (same phone)', 'Submitted (epoch ms)',
             'Offset from presentation mean']);
}

function setup() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  var h = headers_();
  if (sh.getMaxColumns() < h.length) sh.insertColumnsAfter(sh.getMaxColumns(), h.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, sh.getMaxColumns()).clearContent();
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, sh.getMaxRows() - 1, h.length).clearDataValidations();
  sh.getRange(2, h.indexOf('Presentation') + 1, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(false).setHelpText('Presentation must be a number.').build());

  var sum = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  sum.clear();
  var L = function (name) { return colLetter_(h.indexOf(name) + 1); };
  var R = function (name) { return SHEET + '!' + L(name) + ':' + L(name); };
  var pr = R('Presentation'), tot = R('Total (100)'), fl = R('Straight-line flag'), du = R('Duplicate flag');
  var at = R('Attention check'), names = R('Presenters'), dt = R('Date');
  var P = L('Presentation'), T = L('Total (100)'), D = L('Device token'), O = L('Offset from presentation mean');
  sum.getRange(1, 1).setFormula('=QUERY(' + SHEET + '!A:' + O + ',"select ' + P + ', count(' + T + '), avg(' + T + '), min(' + T + '), max(' + T +
    ') where ' + P + ' is not null group by ' + P + ' order by ' + P + ' label ' + P + ' \'Presentation\', count(' + T + ') \'Evaluations\', avg(' + T +
    ') \'Mean total\', min(' + T + ') \'Min\', max(' + T + ') \'Max\' format avg(' + T + ') \'0.0\'",1)');
  sum.getRange(1, 6, 1, 4).setValues([['Mean total (clean)', 'Flagged rows', 'Presenters', 'Date']]);
  var f = [];
  for (var r = 2; r <= 301; r++) {
    f.push(['=IF(A' + r + '="","",IFERROR(ROUND(AVERAGEIFS(' + tot + ',' + pr + ',A' + r + ',' + fl + ',"",' + du + ',"",' + at + ',"<>FAIL"),1),""))',
            '=IF(A' + r + '="","",B' + r + '-COUNTIFS(' + pr + ',A' + r + ',' + fl + ',"",' + du + ',"",' + at + ',"<>FAIL"))',
            '=IF(A' + r + '="","",IFERROR(INDEX(FILTER(' + names + ',' + pr + '=A' + r + ',' + names + '<>""),1),""))',
            '=IF(A' + r + '="","",IFERROR(TEXT(INDEX(FILTER(' + dt + ',' + pr + '=A' + r + '),1),"yyyy-mm-dd"),""))']);
  }
  sum.getRange(2, 6, f.length, 4).setFormulas(f);
  sum.getRange(1, 11).setFormula('=IFERROR(QUERY(QUERY(' + SHEET + '!A:' + O + ',"select ' + D + ', count(' + O + '), avg(' + O + ') where ' + D +
    ' is not null and ' + O + ' is not null group by ' + D + ' label ' + D + ' \'Phone (device token)\', count(' + O +
    ') \'Evaluations\', avg(' + O + ') \'Leniency (avg points above class)\'",1),"select * where Col2 >= 3 order by Col3 desc format Col3 \'0.0\'",1),"Leniency appears once a phone has 3 clean ratings.")');
  sum.getRange(1, 1, 1, 13).setFontWeight('bold');
  sum.getRange(304, 1, 4, 1).setValues([
    ['Rows are grouped into presentations automatically. A new number starts on a new day, after a ' + GAP_MIN + '-minute pause in submissions, or when a phone that already rated the current presentation submits again at least ' + REPEAT_MIN + ' minutes later.'],
    ['To move a row, change its number in the Presentation column of Responses (numbers only). Type the presenters’ names in the Presenters column of any one row.'],
    ['Mean total (clean) leaves out straight-line, duplicate and failed-attention-check submissions.'],
    ['Leniency: how many points a phone gives above (+) or below (-) the other clean ratings of the same presentation, averaged over its clean ratings; shown from 3 ratings.']]);
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
    var tms = sh.getRange(2, h.indexOf('Presentation') + 1, last - 1, 1).getValues();
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
    lock.waitLock(28000);
    var d = JSON.parse(e.postData.contents);

    var ss = ss_();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) { setup(); sh = ss.getSheetByName(SHEET); }
    var h = headers_();

    var subCol = h.indexOf('Submission ID') + 1, devCol = h.indexOf('Device token') + 1;
    var last = sh.getLastRow();
    var nowMs = Date.now(), prevMs = 0, pres = 1, ps = [], tots = [], atts = [], fls = [];
    var presC = h.indexOf('Presentation') + 1;
    if (last > 1) {
      var subs = sh.getRange(2, subCol, last - 1, 1).getValues();
      for (var i = 0; i < subs.length; i++) if (String(subs[i][0]) === clean_(d.subId)) return json_({ ok: true, repeat: true });
      var devs = sh.getRange(2, devCol, last - 1, 1).getValues();
      var eps = sh.getRange(2, h.indexOf('Submitted (epoch ms)') + 1, last - 1, 1).getValues();
      ps = sh.getRange(2, presC, last - 1, 1).getValues();
      tots = sh.getRange(2, h.indexOf('Total (100)') + 1, last - 1, 1).getValues();
      atts = sh.getRange(2, h.indexOf('Attention check') + 1, last - 1, 1).getValues();
      fls = sh.getRange(2, h.indexOf('Straight-line flag') + 1, last - 1, 1).getValues();
      var maxP = 0, lastP = 0, lastEp = 0, seen = 0, k, ek, pk;
      for (k = 0; k < eps.length; k++) {
        ek = Number(eps[k][0]) || 0; pk = Number(ps[k][0]) || 0;
        if (pk > maxP) maxP = pk;
        if (pk && ek >= lastEp) { lastEp = ek; lastP = pk; }
      }
      for (k = 0; k < devs.length; k++) {
        ek = Number(eps[k][0]) || 0; pk = Number(ps[k][0]) || 0;
        if (String(devs[k][0]) === clean_(d.device)) {
          if (ek > prevMs) prevMs = ek;
          if (lastP && pk === lastP && ek > seen) seen = ek;
        }
      }
      var day = function (ms) { return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'yyyy-MM-dd'); };
      var newDay = lastEp > 0 && day(lastEp) !== day(nowMs);
      pres = (!lastP || newDay || nowMs - lastEp > GAP_MIN * 60000 || (seen && nowMs - seen >= REPEAT_MIN * 60000)) ? maxP + 1 : lastP;
    }
    var r = last + 1, presL = colLetter_(presC), devL = colLetter_(devCol), totL = colLetter_(h.indexOf('Total (100)') + 1);
    var epL = colLetter_(h.indexOf('Submitted (epoch ms)') + 1), flL = colLetter_(h.indexOf('Straight-line flag') + 1);
    var dpL = colLetter_(h.indexOf('Duplicate flag') + 1), atL = colLetter_(h.indexOf('Attention check') + 1);
    var dup = '=IF(' + presL + r + '="","",IF(COUNTIFS(' + presL + ':' + presL + ',' + presL + r + ',' + devL + ':' + devL + ',' + devL + r + ',' + epL + ':' + epL + ',"<"&' + epL + r + ')>0,"YES",""))';
    var crit_ = presL + ':' + presL + ',' + presL + r + ',' + flL + ':' + flL + ',"",' + dpL + ':' + dpL + ',"",' + atL + ':' + atL + ',"<>FAIL"';
    var mine = 'AND(' + flL + r + '="",' + dpL + r + '="",' + atL + r + '<>"FAIL")';
    var offset = '=IF(OR(' + presL + r + '="",' + totL + r + '="",NOT(' + mine + ')),"",IFERROR(ROUND(' + totL + r + '-(SUMIFS(' + totL + ':' + totL + ',' + crit_ + ')-' + totL + r + ')/(COUNTIFS(' + crit_ + ')-1),1),""))';
    var hm = function (ms) { ms = Number(ms); return ms > 0 ? Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'HH:mm:ss') : ''; };

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
    var fp = (d.fp && typeof d.fp === 'object') ? d.fp : {};
    var attnRes = (Number(a.value) === Number(a.target) && (Number(a.target) === 2 || Number(a.target) === 4)) ? 'PASS' : 'FAIL';
    var row = [Utilities.formatDate(now, stz, 'yyyy-MM-dd'), Utilities.formatDate(now, stz, 'HH:mm:ss'),
               clean_(d.localDisplay), clean_(d.tz), pres, '', '']
      .concat(vals)
      .concat([crit.c1, crit.c2, crit.c3, crit.c4, crit.c5, crit.c6, total])
      .concat([c.c1, c.c2, c.c3, c.c4, c.c5, c.c6, c.overall].map(text_))
      .concat([Number(d.durationSec) || '', Math.round(sd * 1000) / 1000, distinct, Number(d.flatCriteria) || 0, flat,
               text_(d.flatReason), fast, dup, clean_(d.device), clean_(d.subId), clean_(order),
               attnRes, Number(a.value) ? (Number(a.value) + ' (asked ' + Number(a.target) + ', ' + clean_(a.style) + ')') : '',
               clean_(d.styles), clean_(d.localTime),
               clean_(fp.id), clean_(fp.model), clean_(fp.ua),
               clean_(fp.screen ? fp.screen + ' @' + fp.dpr + 'x' : ''), clean_([fp.glVendor, fp.glRenderer].filter(Boolean).join(' / ')),
               clean_(JSON.stringify(fp)),
               hm(d.firstAt), hm(d.lastAt), prevMs ? Math.round((nowMs - prevMs) / 6000) / 10 : '', nowMs, offset]);
    sh.appendRow(row);
    var okMe = attnRes !== 'FAIL' && !flat;
    var sumT = okMe ? total : 0, n = okMe ? 1 : 0;
    for (var q = 0; q < ps.length; q++) if (Number(ps[q][0]) === pres && atts[q][0] !== 'FAIL' && fls[q][0] !== 'YES' && tots[q][0] !== '') { sumT += Number(tots[q][0]); n++; }
    return json_({ ok: true, total: total, presentation: pres, classN: n, classAvg: n ? Math.round(sumT / n * 10) / 10 : null });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function text_(s) {
  s = String(s || '').slice(0, 2000);
  return s ? "'" + s : '';
}

function clean_(s) {
  s = String(s || '').slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
