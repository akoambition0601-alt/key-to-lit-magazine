/**
 * KEY TO LIT 雑誌まとめ：サイト内アンケート受信
 * Google Apps Script の新規プロジェクトへ、このファイルの内容を貼り付けて利用する。
 * ウェブアプリとして「次のユーザーとして実行：自分」「アクセスできるユーザー：全員」に設定する。
 * 回答先スプレッドシート自体は共有しない。
 */
const FEEDBACK_SPREADSHEET_ID = '18P-A4CDlyx-a53Nalx78JLhqaodyrGzeYYupDvn0c_k';
const FEEDBACK_TAB_NAME = '回答一覧';

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.website) return reply_('IGNORED');
    const id = String(p.id || '').trim();
    if (!/^[0-9a-f-]{36}$/i.test(id)) return reply_('INVALID_ID');
    const ease = String(p.ease || '').slice(0, 40);
    const useful = String(p.useful || '').slice(0, 40);
    const shop = String(p.shop || '').slice(0, 50);
    const comment = String(p.comment || '').slice(0, 1000).trim();
    if (!ease && !useful && !shop && !comment) return reply_('EMPTY');
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sheet = SpreadsheetApp.openById(FEEDBACK_SPREADSHEET_ID).getSheetByName(FEEDBACK_TAB_NAME);
      if (!sheet) throw Error('回答一覧のタブが見つかりません');
      if (findResponse_(sheet, id)) return reply_('DUPLICATE');
      // ユーザー入力がスプレッドシートの数式として処理されないようにする。
      const plain = x => /^[=+@\-]/.test(x) ? "'" + x : x;
      sheet.appendRow([new Date(), id, plain(ease), plain(useful), plain(shop), plain(comment), '2026-launch']);
      return reply_('OK');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
    return reply_('ERROR');
  }
}

/**
 * 送信IDが保存されているかのみ返す確認用エンドポイント。
 * 回答本文、日時、投稿者の情報は公開しない。
 */
function doGet(e) {
  const id = String((e && e.parameter && e.parameter.check) || '');
  let saved = false;
  if (/^[0-9a-f-]{36}$/i.test(id)) {
    try {
      const sheet = SpreadsheetApp.openById(FEEDBACK_SPREADSHEET_ID).getSheetByName(FEEDBACK_TAB_NAME);
      saved = Boolean(sheet && findResponse_(sheet, id));
    } catch (err) { console.error(err); }
  }
  // callback名は固定。利用者の入力をスクリプトに埋め込まない。
  return ContentService.createTextOutput('window.__hokureaFeedbackAck('+JSON.stringify(id)+','+(saved?'true':'false')+');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function findResponse_(sheet, id) {
  const count = sheet.getLastRow() - 1;
  if (count <= 0) return false;
  return Boolean(sheet.getRange(2, 2, count, 1).createTextFinder(id).matchEntireCell(true).findNext());
}

function reply_(message) {
  return ContentService.createTextOutput(message).setMimeType(ContentService.MimeType.TEXT);
}
