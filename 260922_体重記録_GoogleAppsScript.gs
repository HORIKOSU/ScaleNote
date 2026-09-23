/**
 * 体重記録ツール（WeightSync）用 Google Apps Script
 * ------------------------------------------------
 * 役割：
 *  1. iPhoneショートカットからのPOSTを受け取り、スプレッドシートに1行追記する（自動連携の受け口）
 *  2. Webダッシュボードからのfetch(GET)に、記録データをJSONで返す
 *
 * 事前準備：
 *  1. Googleスプレッドシートを新規作成する
 *  2. シート名を「records」にする（1行目に見出し：date, weight, bodyFat, source, savedAt）
 *  3. 拡張機能 → Apps Script を開き、このコードを貼り付けて保存
 *  4. SECRET_KEY を自分だけの適当な文字列に書き換える（推測されにくいランダムな文字列にする）
 *  5. 右上「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」
 *     - アクセスできるユーザー：全員（ここが「全員」でないとiPhone側から呼び出せない）
 *     - デプロイ後に表示される「ウェブアプリのURL」を控えておく（ダッシュボード側・ショートカット側で使う）
 */

// ここを自分だけのランダムな文字列に変更してください（他人に推測されない値にする）
var SECRET_KEY = 'ここに自分だけの合言葉を設定';

var SHEET_NAME = 'records';

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.key !== SECRET_KEY) {
      return jsonOutput({ ok: false, error: 'invalid key' });
    }
    var sheet = getSheet_();
    var date = body.date; // 'YYYY-MM-DD'
    var weight = Number(body.weight);
    var bodyFat = body.bodyFat !== undefined && body.bodyFat !== null && body.bodyFat !== ''
      ? Number(body.bodyFat) : '';
    if (!date || isNaN(weight)) {
      return jsonOutput({ ok: false, error: 'date/weight is required' });
    }
    sheet.appendRow([date, weight, bodyFat, body.source || 'shortcut', new Date()]);
    return jsonOutput({ ok: true });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var key = e.parameter.key;
  if (key !== SECRET_KEY) {
    return jsonOutput({ ok: false, error: 'invalid key' });
  }
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    var dateVal = r[0] instanceof Date
      ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(r[0]);
    rows.push({
      date: dateVal,
      weight: Number(r[1]),
      bodyFat: r[2] === '' ? null : Number(r[2]),
      source: r[3] || ''
    });
  }
  return jsonOutput({ ok: true, records: rows });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['date', 'weight', 'bodyFat', 'source', 'savedAt']);
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
