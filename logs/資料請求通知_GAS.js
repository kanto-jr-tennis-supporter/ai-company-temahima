/**
 * 資料請求／資料送付 通知スクリプト（テマヒマ・ラボ × アド・クリックス様）
 * ------------------------------------------------------------------
 * 対象スプレッドシートの「コール状況」列に「資料請求」または「資料送付」が
 * 入力されたら、その行の 企業名・住所・電話番号（＋社内向けにシート名・行番号）を
 * ① LINE（クライアント向け・スッキリ）② メール（社内＋クライアント）③ Slack（社内）
 * に通知します。届け先ごとに内容を出し分けます。
 *
 * ★はじめに：下の RR_CONFIG を埋めて、setupTriggerRR() を1回だけ実行してください。
 * ★詳しい手順は「資料請求通知_セットアップ手順.md」を参照。
 */

// ================== 設定（ここだけ触ればOK） ==================
const RR_CONFIG = {
  // 監視するシート名（このシートの編集だけ見ます）
  TARGET_SHEETS: ['リスト（寺社仏閣）', 'リスト（社福）'],

  // 列はヘッダー名で探します（列を並び替えても壊れません）。
  // 実際のシートの見出しと1文字違わず合わせてください。
  HEADERS: {
    callStatus: 'コール状況', // 監視する列
    company:    '企業名',
    address:    '住所',
    phone:      '電話番号',
    notifiedAt: '通知日時',   // 任意。この見出しの列を作ると「二重通知」を防げます（無ければ無視）
  },

  // このどれかが入力されたら通知（Q1＝両方）
  TRIGGER_KEYWORDS: ['資料請求', '資料送付'],

  // ① LINE（Messaging API・クライアント向け）──設定が済むまで enabled:false のまま
  LINE: {
    enabled: false,
    channelAccessToken: '★チャネルアクセストークン（長期）を貼る',
    to: '★送信先ID（クライアントとのグループのグループID）を貼る',
  },

  // ② メール（社内＋クライアント様。複数OK）
  EMAIL: {
    enabled: true,
    to: ['★クライアント様のアドレス', '★社内のアドレス'],
  },

  // ③ Slack（社内・Pj-アドクリックス）
  SLACK: {
    enabled: true,
    webhookUrl: '★Slack Incoming Webhook の URL を貼る',
  },

  // 不具合が起きたときに知らせる社内アドレス（サイレント失敗を防ぐ）
  ADMIN_EMAIL: '★社内の管理用アドレス',
};
// ==========================================================


/** 編集トリガー本体（installableトリガーから呼ばれます） */
function onResourceRequestEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (RR_CONFIG.TARGET_SHEETS.indexOf(sheetName) === -1) return;

    const lastCol = sheet.getLastColumn();
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idx = rrBuildHeaderIndex_(header);
    if (idx.callStatus === -1) return; // コール状況列が無ければ何もしない

    const callCol = idx.callStatus + 1; // 1始まり
    // 編集範囲がコール状況列を含むか（1セル編集・貼り付けの複数セル編集どちらも対応）
    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    if (callCol < startCol || callCol > endCol) return;

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    for (let r = startRow; r < startRow + numRows; r++) {
      if (r === 1) continue; // ヘッダー行はスキップ

      const value = String(sheet.getRange(r, callCol).getValue()).trim();
      if (!rrMatchesKeyword_(value)) continue;

      // 二重通知の防止（通知日時列がある場合のみ）
      if (idx.notifiedAt !== -1) {
        const already = sheet.getRange(r, idx.notifiedAt + 1).getValue();
        if (already) continue;
      }

      const rowData = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
      const info = {
        sheetName: sheetName,
        row: r,
        keyword: value,
        company: rrGetCell_(rowData, idx.company),
        address: rrGetCell_(rowData, idx.address),
        phone:   rrGetCell_(rowData, idx.phone),
      };

      const errors = rrNotifyAll_(info);

      // 成功した通知が1つでもあれば通知日時を記録
      if (idx.notifiedAt !== -1 && errors.length < rrCountEnabled_()) {
        const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
        sheet.getRange(r, idx.notifiedAt + 1).setValue(stamp);
      }
      if (errors.length) throw new Error(errors.join(' / '));
    }
  } catch (err) {
    rrNotifyError_(err);
  }
}


/** 3つの通知先へ送る。1つ失敗しても他は続行。失敗内容の配列を返す */
function rrNotifyAll_(info) {
  const errors = [];
  if (RR_CONFIG.LINE.enabled) {
    try { rrSendLine_(info); } catch (e) { errors.push('LINE: ' + e.message); }
  }
  if (RR_CONFIG.EMAIL.enabled) {
    try { rrSendEmail_(info); } catch (e) { errors.push('メール: ' + e.message); }
  }
  if (RR_CONFIG.SLACK.enabled) {
    try { rrSendSlack_(info); } catch (e) { errors.push('Slack: ' + e.message); }
  }
  return errors;
}

function rrCountEnabled_() {
  let n = 0;
  if (RR_CONFIG.LINE.enabled) n++;
  if (RR_CONFIG.EMAIL.enabled) n++;
  if (RR_CONFIG.SLACK.enabled) n++;
  return n;
}


// ---------- 各チャネルの送信 ----------

/** ① LINE：クライアント向け（企業名・住所・電話番号のみ。社内用の行番号・シート名は入れない） */
function rrSendLine_(info) {
  const text =
    '📩 テレアポにて「' + info.keyword + '」のご承諾をいただきました。\n' +
    '下記のお客様へ資料のご送付をお願いいたします。\n\n' +
    '■ 企業名：' + info.company + '\n' +
    '■ 住所：'  + info.address + '\n' +
    '■ 電話番号：' + info.phone;

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + RR_CONFIG.LINE.channelAccessToken },
    payload: JSON.stringify({
      to: RR_CONFIG.LINE.to,
      messages: [{ type: 'text', text: text }],
    }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200) throw new Error('送信失敗(' + code + ') ' + res.getContentText());
}

/** ② メール：社内＋クライアント。追跡用にシート名・行番号も入れる */
function rrSendEmail_(info) {
  const subject = '【' + info.keyword + '】' + info.company + ' 様（' + info.sheetName + '）';
  MailApp.sendEmail({
    to: RR_CONFIG.EMAIL.to.join(','),
    subject: subject,
    body: rrBuildInternalBody_(info),
  });
}

/** ③ Slack：社内。追跡用にシート名・行番号も入れる */
function rrSendSlack_(info) {
  const res = UrlFetchApp.fetch(RR_CONFIG.SLACK.webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: rrBuildInternalBody_(info) }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200) throw new Error('送信失敗(' + code + ') ' + res.getContentText());
}

/** 社内向け本文（メール・Slack共通。行番号・シート名つき） */
function rrBuildInternalBody_(info) {
  return [
    '【資料請求／送付の通知】',
    'キーワード：' + info.keyword,
    'シート：' + info.sheetName,
    '行番号：' + info.row + '行目',
    '企業名：' + info.company,
    '住所：' + info.address,
    '電話番号：' + info.phone,
  ].join('\n');
}


// ---------- ヘルパー ----------

function rrBuildHeaderIndex_(header) {
  const find = name => header.indexOf(name);
  return {
    callStatus: find(RR_CONFIG.HEADERS.callStatus),
    company:    find(RR_CONFIG.HEADERS.company),
    address:    find(RR_CONFIG.HEADERS.address),
    phone:      find(RR_CONFIG.HEADERS.phone),
    notifiedAt: find(RR_CONFIG.HEADERS.notifiedAt),
  };
}

function rrGetCell_(rowData, idx) {
  if (idx === -1) return '(見出しが見つかりません)';
  const v = rowData[idx];
  return (v === '' || v === null || v === undefined) ? '(空欄)' : String(v).trim();
}

function rrMatchesKeyword_(value) {
  if (!value) return false;
  return RR_CONFIG.TRIGGER_KEYWORDS.some(k => value.indexOf(k) !== -1);
}

function rrNotifyError_(err) {
  try {
    MailApp.sendEmail(
      RR_CONFIG.ADMIN_EMAIL,
      '【資料請求通知】エラーが発生しました',
      '通知処理でエラーが発生しました。内容をご確認ください。\n\n' + String((err && err.stack) || err)
    );
  } catch (e) {
    console.error('通知エラー（メールも失敗）: ' + err);
  }
}


// ---------- セットアップ・テスト（手動実行） ----------

/** 最初に1回だけ実行：編集トリガー（installable）を作成します */
function setupTriggerRR() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onResourceRequestEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onResourceRequestEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getActive().toast('通知トリガーを設定しました', '資料請求通知', 5);
}

/** 動作確認：ダミーデータで有効な通知先すべてに1通ずつ送ります */
function testNotifyRR() {
  const errors = rrNotifyAll_({
    sheetName: RR_CONFIG.TARGET_SHEETS[0],
    row: 999,
    keyword: '資料送付',
    company: 'テスト商事株式会社',
    address: '東京都千代田区テスト1-2-3',
    phone: '03-0000-0000',
  });
  if (errors.length) throw new Error('テスト送信に失敗: ' + errors.join(' / '));
  SpreadsheetApp.getActive().toast('テスト通知を送信しました', '資料請求通知', 5);
}
