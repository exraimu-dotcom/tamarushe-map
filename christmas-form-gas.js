/**
 * 玉城移動商店街 CHRISTMAS MARKET 2026 出店申込 受付スクリプト（Google Apps Script）
 *
 * christmas-form.html から送られた申込内容を、Google スプレッドシートに1行ずつ追加します。
 *
 * 【設定手順】
 * 1. christmas-market-2026-entries.xlsx を Google ドライブにアップロードして開き、
 *    「ファイル」→「Google スプレッドシートとして保存」する
 *    （見出し・管理用の列・プルダウンが入った状態で始められます。新規作成でも動きます）
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. 最初から入っているコードを全部消して、このファイルの中身を貼り付けて保存
 * 4. 右上「デプロイ」→「新しいデプロイ」
 *      種類の選択：ウェブアプリ
 *      次のユーザーとして実行：自分
 *      アクセスできるユーザー：全員
 *    →「デプロイ」を押し、Googleアカウントのアクセスを許可する
 * 5. 表示された「ウェブアプリのURL」（https://script.google.com/macros/s/…/exec）をコピーし、
 *    christmas-form.html の GAS_URL = '' の '' の中に貼り付ける
 *
 * ※コードを修正した場合は「デプロイを管理」→ 編集 → バージョン「新バージョン」で更新してください
 *   （URLはそのまま使えます）。
 */

// 申込があったときに通知メールを送る宛先（不要なら空のまま）
const NOTIFY_EMAIL = '';

const SHEET_NAME = '申込一覧';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const record = Object.assign({ '受付日時': new Date() }, data);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // 見出し行を用意し、未登録の項目があれば列を追加する
    let headers = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];
    Object.keys(record).forEach(key => {
      if (!headers.includes(key)) headers.push(key);
    });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    sheet.appendRow(headers.map(h => (h in record ? record[h] : '')));

    if (NOTIFY_EMAIL) {
      const body = Object.keys(data).map(k => `【${k}】\n${data[k]}`).join('\n\n');
      MailApp.sendEmail(NOTIFY_EMAIL,
        `【クリスマスマーケット出店申込】${data['店舗名'] || ''}`,
        body + '\n\n' + ss.getUrl());
    }

    return ContentService.createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
