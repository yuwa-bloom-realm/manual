/**
 * 追いかけない集客 説明会 申込フォーム 受信スクリプト
 * ------------------------------------------------------------
 * ■ 使い方
 * 1. 受け取り用のスプレッドシートを新規作成して開く
 * 2. 拡張機能 → Apps Script を開く
 * 3. 中身を全部消して、このコードを丸ごと貼り付け
 * 4. NOTIFY_TO を自分のメアドに（すでに設定済み）
 * 5. 上のメニュー「デプロイ」→「新しいデプロイ」
 *    - 種類：ウェブアプリ
 *    - 実行するユーザー：自分
 *    - アクセスできるユーザー：全員
 * 6. 発行された「ウェブアプリのURL」をコピー
 * 7. form.html の GAS_URL にそのURLを貼る
 * ------------------------------------------------------------
 */

// ▼ 通知メールの宛先（ゆわのアドレス）
const NOTIFY_TO = "ここに自分のメールアドレスを入れる";

// ▼ 記録先シート名（自動で作られる）
const SHEET_NAME = "申込";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    const now = new Date();

    // シートへ1行追加
    sheet.appendRow([
      now,
      data.name || "",
      data.email || "",
      data.job || "",
      data.worry || "",
      data.career || "",
      data.usedBefore || "",
      data.usingNow || "",
      data.reasonUsing || "",
      data.reasonQuit || "",
      data.want || ""
    ]);

    // ゆわへ通知メール
    sendNotify_(data, now);

    return json_({ ok: true });
  } catch (err) {
    // エラーも自分に飛ばしておくと気づける
    try {
      MailApp.sendEmail(NOTIFY_TO, "【申込フォーム】エラー発生", String(err));
    } catch (_) {}
    return json_({ ok: false, error: String(err) });
  }
}

// 動作確認用（URLをブラウザで開くと表示される）
function doGet() {
  return ContentService.createTextOutput("フォーム受信スクリプトは稼働中です。");
}

/** シートを取得（なければ作成＋見出し行） */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "受信日時", "お名前", "メールアドレス", "今のお仕事",
      "悩んでいること", "活動歴", "自動化の利用経験",
      "いまも使っているか", "見てみようと思った理由", "やめた理由",
      "知りたいこと"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** ゆわへの通知メール */
function sendNotify_(d, now) {
  const tz = Session.getScriptTimeZone();
  const ts = Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm");

  let body = "";
  body += "説明会の新しいお申し込みが届きました。\n\n";
  body += "──────────────────────\n";
  body += "受信日時：" + ts + "\n";
  body += "お名前　：" + (d.name || "") + "\n";
  body += "メール　：" + (d.email || "") + "\n";
  body += "──────────────────────\n";
  body += "【今のお仕事】\n" + (d.job || "") + "\n\n";
  body += "【悩んでいること】\n" + (d.worry || "") + "\n\n";
  body += "【活動歴】\n" + (d.career || "") + "\n\n";
  body += "【自動化の利用経験】" + (d.usedBefore || "") + "\n";
  if (d.usedBefore === "ある") {
    body += "【いまも使っているか】" + (d.usingNow || "") + "\n";
    if (d.usingNow === "使っている" && d.reasonUsing) {
      body += "【見てみようと思った理由】\n" + d.reasonUsing + "\n";
    }
    if (d.usingNow === "やめた" && d.reasonQuit) {
      body += "【やめた理由】\n" + d.reasonQuit + "\n";
    }
  }
  body += "\n【知りたいこと】\n" + (d.want || "（記入なし）") + "\n";
  body += "──────────────────────\n";

  MailApp.sendEmail(
    NOTIFY_TO,
    "【説明会お申し込み】" + (d.name || "お名前未記入") + " 様",
    body
  );
}

/** JSONレスポンス */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
