/**
 * 競技会申込書メーカー 利用統計（匿名）
 * ・個人情報（犬名・飼い主名・住所・電話番号・メールアドレスなど）は一切受け取らない・保存しない。
 * ・受け取るのは「セッションID / イベント種別 / 大会名 / 滞在秒数 / 端末種別」だけ。
 * ・doPost: アプリ（GitHub Pages）からの匿名ログ追記
 * ・doGet : キーが一致したときだけ集計JSONを返す（茂樹さん専用の統計パネル用）
 */

var SHEET_NAME = 'ログ';
var HEADER = ['日時', 'セッションID', 'イベント', '大会名', '滞在秒数', '端末'];
var STATS_KEY = 'funstyle2026';
var EVENTS = ['visit', 'pdf', 'transfer', 'leave'];
var MAX_DURATION = 7200;   // 滞在秒数の上限（2時間）
var TZ = 'Asia/Tokyo';

/* ログシートを取得（無ければ作成。1行目にヘッダーを入れる） */
function getLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    var sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() === 0) {
      sh = sheets[0];
      sh.setName(SHEET_NAME);
    } else {
      sh = ss.insertSheet(SHEET_NAME);
    }
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* 初回セットアップ（エディタから1回実行するとシートが整い、承認も済む） */
function setup() {
  var sh = getLogSheet_();
  Logger.log('シート準備OK: ' + sh.getName() + ' / 行数 ' + sh.getLastRow());
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 受信（アプリからの匿名ログ） ---------- */
function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      body = e.parameter;
    }
    var event = String(body.event || '');
    if (EVENTS.indexOf(event) < 0) return jsonOut_({ ok: false, error: 'bad_event' });

    var sid = String(body.sessionId || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
    var compName = String(body.compName || '').slice(0, 80);
    var dur = Number(body.durationSec || 0);
    if (!isFinite(dur) || dur < 0) dur = 0;
    dur = Math.min(Math.round(dur), MAX_DURATION);
    var device = String(body.device || '').slice(0, 12);

    getLogSheet_().appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      sid, event, compName, dur, device
    ]);
    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/* ---------- 集計（茂樹さん専用） ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (String(p.action || '') === 'ping') {
    return jsonOut_({ ok: true, time: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss') });
  }
  if (String(p.key || '') !== STATS_KEY) {
    return jsonOut_({ ok: false, error: 'unauthorized' });
  }
  return jsonOut_(buildStats_());
}

function buildStats_() {
  var sh = getLogSheet_();
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, HEADER.length).getValues() : [];

  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var sessions = {};   // visit したセッション
  var maxDur = {};     // セッションごとの最大滞在秒数
  var comps = {};      // 大会名 -> {name, visits, pdfs}
  var days = {};       // yyyy-MM-dd -> {date, visits, pdfs}
  var todaySessions = {};
  var totalVisits = 0, totalPdf = 0, totalTransfer = 0, totalLeave = 0;
  var todayVisits = 0, todayPdf = 0;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var stamp = (r[0] instanceof Date)
      ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm:ss')
      : String(r[0] || '');
    var day = stamp.slice(0, 10);
    var sid = String(r[1] || '');
    var ev = String(r[2] || '');
    var comp = String(r[3] || '');
    var dur = Number(r[4] || 0);

    if (ev === 'visit') {
      totalVisits++;
      if (sid) sessions[sid] = true;
      if (day === today) { todayVisits++; if (sid) todaySessions[sid] = true; }
    } else if (ev === 'pdf') {
      totalPdf++;
      if (day === today) todayPdf++;
    } else if (ev === 'transfer') {
      totalTransfer++;
    } else if (ev === 'leave') {
      totalLeave++;
    }

    if (dur > 0 && dur <= MAX_DURATION && sid) {
      if (!maxDur[sid] || dur > maxDur[sid]) maxDur[sid] = dur;
    }

    if (ev === 'visit' || ev === 'pdf') {
      var c = comps[comp] || { name: comp || '(大会名なし)', visits: 0, pdfs: 0 };
      if (ev === 'visit') c.visits++; else c.pdfs++;
      comps[comp] = c;
      var d = days[day] || { date: day, visits: 0, pdfs: 0 };
      if (ev === 'visit') d.visits++; else d.pdfs++;
      days[day] = d;
    }
  }

  /* 平均滞在時間（セッションごとの最大滞在秒数の平均） */
  var durSum = 0, durN = 0;
  for (var k in maxDur) { durSum += maxDur[k]; durN++; }
  var avgDurationSec = durN ? Math.round(durSum / durN) : 0;

  /* 大会別 */
  var byComp = [];
  for (var k2 in comps) byComp.push(comps[k2]);
  byComp.sort(function (a, b) {
    return (b.pdfs - a.pdfs) || (b.visits - a.visits) || (a.name < b.name ? -1 : 1);
  });

  /* 直近7日間（今日を含む） */
  var recentDays = [];
  var base = new Date();
  for (var k3 = 6; k3 >= 0; k3--) {
    var dt = new Date(base.getFullYear(), base.getMonth(), base.getDate() - k3, 12, 0, 0);
    var key = Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
    var dd = days[key] || { date: key, visits: 0, pdfs: 0 };
    recentDays.push({ date: key, visits: dd.visits, pdfs: dd.pdfs });
  }

  return {
    ok: true,
    updatedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    totalSessions: Object.keys(sessions).length,
    todaySessions: Object.keys(todaySessions).length,
    totalVisits: totalVisits,
    todayVisits: todayVisits,
    totalPdf: totalPdf,
    todayPdf: todayPdf,
    totalTransfer: totalTransfer,
    totalLeave: totalLeave,
    avgDurationSec: avgDurationSec,
    byComp: byComp.slice(0, 40),
    recentDays: recentDays
  };
}
