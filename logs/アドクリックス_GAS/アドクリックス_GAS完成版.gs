/*************************************************
 * アド・クリックス様 スプレッドシート GAS【完成版・1ファイル】
 * テマヒマ・ラボ（担当：ハック）
 * ------------------------------------------------
 *  ⚠️ このファイル1本だけを Apps Script に置いてください。
 *     旧ファイル（架電ログ／資料送付リスト転記）が残っていると
 *     「SS_CFG は宣言済み」等のエラーで全部止まります。
 *
 *  ■ 中身（メニュー「コールログ」から使えます）
 *   01 架電ログ再構築      rebuildCallLog
 *   02 報告シート月別更新  updateReportSheet_ ／ 毎日22時 dailyUpdateAll
 *   03 メニュー            onOpen
 *   04 メルマガ送信        sendNewsletterWithConfirm ほか
 *   05 資料送付リスト転記  onEdit（自動）／ syncShiryoSofuAll（一括）※資料送付・資料請求の両方
 *   06 資料請求／送付通知  onEditNotify（自動）／ setupNotifyTrigger ／ testNotify
 *
 *  ■ 設定を触るのは 06 の NOTIFY_CFG（★の箇所）だけ。
 *  ■ 元のスクリプトからの変更は3点だけ：
 *     ・01 架電ログ：③〜⑤回目の「連絡先」「メール」列を①②と同じ並びに（シート側の並び替えに合わせた）
 *     ・05 転記：「資料請求」も資料送付リストへ転記（F列に資料送付／資料請求を記入）
 *     ・01 キャッチ判定：「【資料請求】」もキャッチに数える
 *************************************************/

/*************************************************
 * 01〜04 架電ログ・報告シート・メニュー・メルマガ送信
 *************************************************/

function rebuildCallLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sources = [
    { name: 'リスト（寺社仏閣）', prefix: '寺' },
    { name: 'リスト（社福）',   prefix: '福' }
  ];
  const rounds = [
    { n:1, caller:10, date:11, time:12, status:13, nextAction:14, nextDate:15, comment:16, contact:17, mail:18 },
    { n:2, caller:19, date:20, time:21, status:22, nextAction:23, nextDate:24, comment:25, contact:26, mail:27 },
    { n:3, caller:28, date:29, time:30, status:31, nextAction:32, nextDate:33, comment:34, contact:35, mail:36 },
    { n:4, caller:37, date:38, time:39, status:40, nextAction:41, nextDate:42, comment:43, contact:44, mail:45 },
    { n:5, caller:46, date:47, time:48, status:49, nextAction:50, nextDate:51, comment:52, contact:53, mail:54 }
  ];
  const output = [];

  sources.forEach(src => {
    const sh = ss.getSheetByName(src.name);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const data = sh.getRange(2, 1, lastRow - 1, 54).getValues();
    data.forEach((row, i) => {
      const sheetRow = i + 2;
      const company = row[1];

      rounds.forEach(r => {
        const statusVal = row[r.status - 1];
        if (statusVal === '' || statusVal === null) return;

        const dateVal = row[r.date - 1];
        const timeVal = row[r.time - 1];
        const hourSlot = parseHour(timeVal);
        const commentVal = row[r.comment - 1] || '';
        const cls = classifyCallStatus_(statusVal, commentVal);

        output.push([
          src.name, sheetRow, company,
          row[r.contact - 1] || '',
          row[r.caller - 1] || '',
          dateVal, timeVal || '', hourSlot,
          statusVal,
          row[r.nextAction - 1] || '',
          row[r.nextDate - 1] || '',
          commentVal,
          row[r.mail - 1] || '',
          r.n,
          src.prefix + sheetRow + '-' + r.n,
          cls.connected,
          cls.caught
        ]);
      });
    });
  });

  const logSheet = ss.getSheetByName('架電ログ');
  const lastLogRow = logSheet.getLastRow();
  if (lastLogRow > 1) {
    logSheet.getRange(2, 1, lastLogRow - 1, 17).clearContent();
  }
  if (output.length > 0) {
    logSheet.getRange(2, 1, output.length, 17).setValues(output);
  }
  logSheet.getRange(1, 16).setValue('接電');
  logSheet.getRange(1, 17).setValue('キャッチ');
  SpreadsheetApp.getActiveSpreadsheet().toast('架電ログを更新しました（' + output.length + '件）', '完了', 5);
}

/********* 接電・キャッチ判定 *********/
const CATCH_ALWAYS = ['【アポ獲得】', '【資料送付】', '【資料請求】', '【定期フォロー】', '【関心あり】', '【決済者断り】'];
const CATCH_NEVER = ['【受付断り】', '【不在】', 'その他'];
const CATCH_LEGACY_AMBIGUOUS = ['【断り】', '【タイミングが悪い】'];
const DECISION_MAKER_KEYWORDS = ['住職'];
const NOT_CONNECTED_STATUSES = ['番号不通'];

function classifyCallStatus_(status, comment) {
  const s = (status || '').toString().trim();
  const c = (comment || '').toString();
  const connected = NOT_CONNECTED_STATUSES.indexOf(s) === -1;

  let caught = false;
  if (CATCH_ALWAYS.indexOf(s) !== -1) {
    caught = true;
  } else if (CATCH_LEGACY_AMBIGUOUS.indexOf(s) !== -1) {
    caught = DECISION_MAKER_KEYWORDS.some(kw => c.indexOf(kw) !== -1);
  } else {
    caught = false;
  }
  return { connected: connected, caught: caught };
}

function parseHour(timeVal) {
  if (!timeVal) return '';
  if (timeVal instanceof Date) return timeVal.getHours();
  const s = String(timeVal).replace('：', ':').trim();
  const m = s.match(/^(\d{1,2}):(\d{1,2})/);
  if (!m) return '';
  const h = parseInt(m[1], 10);
  return isNaN(h) ? '' : h;
}

/********* 報告シート：月別推移の自動行追加＋更新 *********/
function updateReportSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('報告シート');
  const startRow = 3;
  const lastCol = 14;

  let r = startRow;
  while (sh.getRange(r, 1).getValue() instanceof Date) r++;
  let lastRow = r - 1;
  if (lastRow < startRow) return;

  const now = new Date();
  const currentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  let lastMonth = sh.getRange(lastRow, 1).getValue();

  while (lastMonth < currentMonth) {
    const nextMonth = new Date(lastMonth.getFullYear(), lastMonth.getMonth() + 1, 1);
    sh.insertRowAfter(lastRow);
    const newRow = lastRow + 1;
    sh.getRange(lastRow, 1, 1, lastCol).copyTo(sh.getRange(newRow, 1, 1, lastCol));
    sh.getRange(newRow, 1).setValue(nextMonth);
    lastRow = newRow;
    lastMonth = nextMonth;
  }

  for (let row = startRow; row <= lastRow; row++) {
    sh.getRange(row, 4).setFormula(
      "=COUNTIFS('架電ログ'!$F:$F,\">=\"&$A" + row + ",'架電ログ'!$F:$F,\"<\"&EDATE($A" + row + ",1),'架電ログ'!$P:$P,TRUE)"
    );
    sh.getRange(row, 6).setFormula(
      "=COUNTIFS('架電ログ'!$F:$F,\">=\"&$A" + row + ",'架電ログ'!$F:$F,\"<\"&EDATE($A" + row + ",1),'架電ログ'!$Q:$Q,TRUE)"
    );
    sh.getRange(row, 7).setFormula("=F" + row);
  }
}

/********* 毎日22時の自動更新 *********/
function dailyUpdateAll() {
  rebuildCallLog();
  updateReportSheet_();
}

function installDailyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyUpdateAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyUpdateAll').timeBased().atHour(22).everyDays(1).create();
  SpreadsheetApp.getUi().alert('毎日22時に「架電ログ更新→報告シート更新」が自動実行されるように設定しました。');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('コールログ')
    .addItem('架電ログ＋報告シートを今すぐ更新', 'dailyUpdateAll')
    .addItem('架電ログだけ更新', 'rebuildCallLog')
    .addSeparator()
    .addItem('毎日22時の自動更新を設定する（初回のみ）', 'installDailyReportTrigger')
    .addSeparator()
    .addItem('資料送付リストへ一括転記', 'syncShiryoSofuAll')
    .addSeparator()
    .addItem('資料請求通知：自動通知をONにする（初回のみ）', 'setupNotifyTrigger')
    .addItem('資料請求通知：テスト送信', 'testNotify')
    .addToUi();
}

/*************************************************
 * 04_メルマガ送信（アド・クリックス様用）
 * 資料送付リスト → メルマガ送信リスト → Gmail送信（ボタン起動）
 * ヘッダー画像リンク・件名は表（Q2・R2）から読み取る
 *************************************************/

const SRC_SHEET  = "資料送付リスト";
const DEST_SHEET = "メルマガ送信リスト";
const HEADER_ROW = 1;

const SRC_DATE_COL = 1;
const SRC_CHECK_COL = 2;

const DEST_COPY_START_COL = 6;
const DEST_COPY_COLS = 5;
const DEST_SENDDATE_HEADERS = ["送信日", "送信日付", "送付日", "配信日"];

const HDR_COMPANY = "会社名";
const HDR_DEPT    = "部署";
const HDR_TITLE   = "役職";
const HDR_PERSON  = "氏名";
const HDR_MAIL    = "Mail";

const BODY_CELL = "K2";
const ATTACH_CELL = "L2";
const TEST_MODE_CELL = "M2";
const TEST_TO_CELL = "N2";
const TEST_COMPANY_CELL = "O2";
const TEST_PERSON_CELL = "P2";
const HEADER_IMAGE_CELL = "Q2";
const SUBJECT_CELL = "R2";

/********* ① 資料送付リスト：B列にチェックボックス *********/
function setupSourceCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SRC_SHEET);
  const ui = SpreadsheetApp.getUi();
  if (!sh) return ui.alert("元シートが見つかりません：" + SRC_SHEET);

  const lastRow = sh.getLastRow();
  if (lastRow < HEADER_ROW) return ui.alert("シートが空です。");

  sh.getRange(HEADER_ROW, SRC_DATE_COL).setValue("転記日");
  sh.getRange(HEADER_ROW, SRC_CHECK_COL).setValue("転記");

  if (lastRow > HEADER_ROW) {
    const rng = sh.getRange(HEADER_ROW + 1, SRC_CHECK_COL, lastRow - HEADER_ROW, 1);
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    rng.setDataValidation(rule);
    const vals = rng.getValues().map(r => [r[0] === true]);
    rng.setValues(vals);
  }
  ui.alert("B列にチェックボックスを設定しました。");
}

/********* ② 上書き転記：メルマガ送信リストのF〜Jだけ全消し→上書き *********/
function overwriteToDest_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SRC_SHEET);
  const dest = ss.getSheetByName(DEST_SHEET);

  if (!src) return ui.alert("元シートが見つかりません：" + SRC_SHEET);
  if (!dest) return ui.alert("転記先シートが見つかりません：" + DEST_SHEET);

  const srcLastRow = src.getLastRow();
  const srcLastCol = src.getLastColumn();
  if (srcLastRow <= HEADER_ROW) return ui.alert("資料送付リストにデータがありません。");

  const srcHeader = src.getRange(HEADER_ROW, 1, 1, srcLastCol).getValues()[0].map(x => String(x).trim());
  const destHeader = dest.getRange(HEADER_ROW, DEST_COPY_START_COL, 1, DEST_COPY_COLS).getValues()[0].map(x => String(x).trim());

  const srcValues = src.getRange(HEADER_ROW + 1, 1, srcLastRow - HEADER_ROW, srcLastCol).getValues();

  const checked = [];
  const checkedRowNumbers = [];
  for (let i = 0; i < srcValues.length; i++) {
    const row = srcValues[i];
    if (row[SRC_CHECK_COL - 1] === true) {
      checked.push(row);
      checkedRowNumbers.push(HEADER_ROW + 1 + i);
    }
  }

  const destLastRow = dest.getLastRow();
  if (destLastRow > HEADER_ROW) {
    dest.getRange(HEADER_ROW + 1, DEST_COPY_START_COL, destLastRow - HEADER_ROW, DEST_COPY_COLS).clearContent();
  }

  if (checked.length === 0) {
    ui.alert("チェック済み0件だったので、メルマガ送信リストのF〜Jを空にしました。");
    return;
  }

  const out = [];
  for (const row of checked) {
    const newRow = new Array(DEST_COPY_COLS).fill("");
    for (let d = 0; d < destHeader.length; d++) {
      const dh = destHeader[d];
      if (!dh) continue;
      const sIdx = srcHeader.indexOf(dh);
      if (sIdx === -1) continue;
      if (sIdx === SRC_DATE_COL - 1) continue;
      if (sIdx === SRC_CHECK_COL - 1) continue;
      newRow[d] = row[sIdx];
    }
    out.push(newRow);
  }

  dest.getRange(HEADER_ROW + 1, DEST_COPY_START_COL, out.length, DEST_COPY_COLS).setValues(out);

  const today = new Date();
  src.getRangeList(checkedRowNumbers.map(r => `A${r}`)).setValue(today);

  ui.alert(`上書き転記しました：${out.length}件`);
}

/********* ③ 資料送付リスト：B列チェック解除 *********/
function clearSourceChecks() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SRC_SHEET);
  if (!sh) return ui.alert("元シートが見つかりません：" + SRC_SHEET);

  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  sh.getRange(HEADER_ROW + 1, SRC_CHECK_COL, lastRow - HEADER_ROW, 1).setValue(false);
  ui.alert("B列のチェックをすべて解除しました。");
}

/********* ⓪ 資料送付リスト：全行にチェックを入れる *********/
function checkAllSourceRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SRC_SHEET);
  const ui = SpreadsheetApp.getUi();
  if (!sh) return ui.alert("元シートが見つかりません：" + SRC_SHEET);

  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return ui.alert("データがありません。");

  const numRows = lastRow - HEADER_ROW;
  const checkRange = sh.getRange(HEADER_ROW + 1, SRC_CHECK_COL, numRows, 1);
  checkRange.setValue(true);
  ui.alert(`全${numRows}件にチェックを入れました。`);
}

/********* ④ メルマガ送信（テスト/本番・ボタンから起動） *********/
function sendNewsletterWithConfirm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEST_SHEET);
  const ui = SpreadsheetApp.getUi();
  if (!sheet) return ui.alert("送信対象シートが見つかりません：" + DEST_SHEET);

  const bodyTemplate = String(sheet.getRange(BODY_CELL).getDisplayValue()).trim();
  if (!bodyTemplate) return ui.alert(`本文が空です。${BODY_CELL} に本文を入力してください。`);

  const subject = String(sheet.getRange(SUBJECT_CELL).getDisplayValue()).trim();
  if (!subject) return ui.alert(`件名が空です。${SUBJECT_CELL} に件名を入力してください。`);

  const urls = getDriveUrlsFromCell_(sheet, ATTACH_CELL);
  if (urls.length === 0) return ui.alert(`添付リンクが空です。${ATTACH_CELL} にDriveの共有リンクを入れてください。`);

  const { blobs: attachmentBlobs, names: attachmentNames } = buildPdfAttachmentsFromUrls_(urls);
  if (attachmentBlobs.length === 0) return ui.alert("PDFが取得できませんでした。添付リンクがPDFファイルの共有リンクか確認してください。");

  const headerImageBlob = getInlineImageBlob_(sheet, HEADER_IMAGE_CELL);

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  const idxCompany = header.indexOf(HDR_COMPANY);
  const idxDept    = header.indexOf(HDR_DEPT);
  const idxTitle   = header.indexOf(HDR_TITLE);
  const idxPerson  = header.indexOf(HDR_PERSON);
  const idxMail    = header.indexOf(HDR_MAIL);

  if (idxMail === -1) return ui.alert(`メール列が見つかりません。ヘッダー「${HDR_MAIL}」が必要です。`);
  if (idxCompany === -1) return ui.alert(`会社名列が見つかりません。ヘッダー「${HDR_COMPANY}」が必要です。`);
  if (idxPerson === -1) return ui.alert(`氏名列が見つかりません。ヘッダー「${HDR_PERSON}」が必要です。`);

  let idxSendDate = findHeaderIndex_(header, DEST_SENDDATE_HEADERS);
  if (idxSendDate === -1) idxSendDate = 0;

  const testMode = Boolean(sheet.getRange(TEST_MODE_CELL).getValue());

  if (testMode) {
    const testTo = getTestTo_(sheet, TEST_TO_CELL);
    const testCompany = String(sheet.getRange(TEST_COMPANY_CELL).getDisplayValue()).trim();
    const testPerson  = String(sheet.getRange(TEST_PERSON_CELL).getDisplayValue()).trim();

    const atenaBlock = buildAtenaBlock_(testCompany, "", "", testPerson);
    const { html, plain } = buildMailBody_(atenaBlock, bodyTemplate, headerImageBlob);

    const res = ui.alert(
      "テスト送信の確認",
      `テストモードがONです。\n件名：${subject}\n宛先：${testTo}\n※テストは1通のみ送信します。\n\n【添付PDF】\n・${attachmentNames.join("\n・")}\n\n送信しますか？`,
      ui.ButtonSet.OK_CANCEL
    );
    if (res !== ui.Button.OK) return;

    sendMail_(testTo, subject, plain, html, attachmentBlobs, headerImageBlob);
    ui.alert(`テスト送信完了（1通）\n宛先：${testTo}\n添付：${attachmentBlobs.length}件`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < HEADER_ROW + 1) return ui.alert("データ行がありません。");

  const numRows = lastRow - HEADER_ROW;
  const values = sheet.getRange(HEADER_ROW + 1, 1, numRows, lastCol).getValues();

  const targets = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[idxSendDate]) continue;

    const email = String(row[idxMail] ?? "").trim();
    if (!email || email.indexOf("@") === -1) continue;

    targets.push({
      rowNumber: HEADER_ROW + 1 + i,
      email,
      company: String(row[idxCompany] ?? "").trim(),
      dept: idxDept !== -1 ? String(row[idxDept] ?? "").trim() : "",
      title: idxTitle !== -1 ? String(row[idxTitle] ?? "").trim() : "",
      person: String(row[idxPerson] ?? "").trim(),
    });
  }

  if (targets.length === 0) {
    return ui.alert("送信対象が0件でした（送信日が空欄、かつメールが有効な行がありません）。");
  }

  const res = ui.alert(
    "一括送信の確認",
    `テストモードは OFF です。\n件名：${subject}\n送信対象：${targets.length} 件\n\n【添付PDF】\n・${attachmentNames.join("\n・")}\n\nこの内容で一括送信していいですか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  let sent = 0, failed = 0;

  for (const t of targets) {
    try {
      const atenaBlock = buildAtenaBlock_(t.company, t.dept, t.title, t.person);
      const { html, plain } = buildMailBody_(atenaBlock, bodyTemplate, headerImageBlob);

      sendMail_(t.email, subject, plain, html, attachmentBlobs, headerImageBlob);

      sheet.getRange(t.rowNumber, idxSendDate + 1).setValue(new Date());
      sent++;
    } catch (e) {
      failed++;
      console.error(`Row ${t.rowNumber} 送信失敗: ${e && e.message ? e.message : e}`);
    }
  }

  sheet.getRange(TEST_MODE_CELL).setValue(true);
  ui.alert(`本送信完了：${sent}件\n失敗：${failed}件\n\nテストモード（${TEST_MODE_CELL}）は自動でONに戻しました。`);
}

/********* メール本文の組み立て（*装飾 → HTML／ヘッダー画像埋め込み） *********/
function buildMailBody_(atenaBlock, bodyTemplate, headerImageBlob) {
  const plain = `${atenaBlock}\n\n${bodyTemplate}`;
  const bodyHtml = decorateText_(`${atenaBlock}\n\n${bodyTemplate}`);
  const headerHtml = headerImageBlob ? `<img src="cid:headerImage" style="max-width:100%;"><br><br>` : "";
  const html = `<div style="font-family:sans-serif;font-size:14px;line-height:1.7;">${headerHtml}${bodyHtml}</div>`;
  return { html, plain };
}

// *文字*     → 大きな文字＋太字
// **文字**   → オレンジ色（サイズは他と同じ）
// ***文字*** → 大きな文字（太さは普通）
// ****文字**** → サイズは他と同じ＋太字
// ★判定は4個→3個→2個→1個の順で行うので、誤判定しません
// [[文字|URL]]  → オレンジ色のボタン（リンク付き）に変換
// __文字__    → 下線（アンダーライン）
function decorateText_(text) {
  let html = String(text)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/\[\[(.+?)\|(https?:\/\/[^\]]+)\]\]/g, '<a href="$2" target="_blank" style="display:inline-block;background-color:#e67e22;color:#ffffff;padding:10px 28px;border-radius:6px;text-decoration:none;font-weight:bold;">$1</a>')
    .replace(/\*{4}(.+?)\*{4}/g, "<b>$1</b>")
    .replace(/\*{3}(.+?)\*{3}/g, '<span style="font-size:1.5em;">$1</span>')
    .replace(/\*{2}(.+?)\*{2}/g, '<span style="color:#e67e22;">$1</span>')
    .replace(/\*(.+?)\*/g, '<b style="font-size:1.5em;">$1</b>')
    .replace(/_{2}(.+?)_{2}/g, '<u>$1</u>')
    .replace(/\n/g, "<br>");
  return html;
}

function sendMail_(to, subject, plainBody, htmlBody, attachmentBlobs, headerImageBlob) {
  const options = { htmlBody: htmlBody, attachments: attachmentBlobs };
  if (headerImageBlob) {
    options.inlineImages = { headerImage: headerImageBlob };
  }
  GmailApp.sendEmail(to, subject, plainBody, options);
}

function getInlineImageBlob_(sheet, a1) {
  const raw = String(sheet.getRange(a1).getDisplayValue() || "").trim();
  if (!raw) return null;
  const fileId = extractDriveFileId_(raw);
  if (!fileId) return null;
  try {
    const file = DriveApp.getFileById(fileId);
    const mime = file.getMimeType();

    // 画像ファイルならそのまま使う
    if (/^image\/(png|jpe?g|gif)$/i.test(mime)) {
      return file.getBlob().setName("header." + mime.split("/")[1].replace("jpeg", "jpg"));
    }

    // PDFなど画像以外 → Driveのプレビュー画像（PNG・高解像度）に変換して使う
    const token = ScriptApp.getOAuthToken();
    const meta = JSON.parse(UrlFetchApp.fetch(
      "https://www.googleapis.com/drive/v3/files/" + fileId + "?fields=thumbnailLink&supportsAllDrives=true",
      { headers: { Authorization: "Bearer " + token } }
    ).getContentText());
    if (!meta.thumbnailLink) return null;
    const bigUrl = meta.thumbnailLink.replace(/=s\d+$/, "=s1600");
    const img = UrlFetchApp.fetch(bigUrl, { headers: { Authorization: "Bearer " + token } }).getBlob();
    return img.setName("header.png");
  } catch (e) {
    console.error("ヘッダー画像の取得に失敗: " + (e && e.message ? e.message : e));
    return null;
  }
}

/********* 共通：宛名ブロック（会社名／部署＋役職／氏名） *********/
function buildAtenaBlock_(company, dept, title, person) {
  const lines = [];
  if (company) lines.push(company);

  const roleParts = [dept, title].filter(Boolean).join(" ");
  let secondLine = "";
  if (person) {
    secondLine = (roleParts ? roleParts + " " : "") + person + " 様";
  } else if (roleParts) {
    secondLine = roleParts + " 御中";
  } else {
    secondLine = "ご担当者様";
  }
  lines.push(secondLine);
  return lines.join("\n");
}

function getTestTo_(sheet, testToCellA1) {
  const explicit = String(sheet.getRange(testToCellA1).getDisplayValue()).trim();
  if (explicit) return explicit;
  const me = Session.getActiveUser().getEmail();
  if (me) return me;
  throw new Error(`テスト送信先が空です。${testToCellA1} に自分のメールアドレスを入れてください。`);
}

function getDriveUrlsFromCell_(sheet, a1) {
  const r = sheet.getRange(a1);
  const rt = r.getRichTextValue();

  if (rt) {
    const direct = rt.getLinkUrl();
    if (direct) return [direct];
    const runs = rt.getRuns();
    const runUrls = [];
    for (const run of runs) {
      const u = run.getLinkUrl();
      if (u) runUrls.push(u);
    }
    if (runUrls.length) return Array.from(new Set(runUrls));
  }

  const raw = String(r.getDisplayValue() || "").trim();
  if (!raw) return [];
  const lines = raw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  return lines.filter(s => /^https?:\/\//i.test(s) || s.includes("drive.google.com"));
}

function buildPdfAttachmentsFromUrls_(urls) {
  const blobs = [];
  const names = [];
  for (const url of urls) {
    const fileId = extractDriveFileId_(url);
    if (!fileId) continue;
    const file = DriveApp.getFileById(fileId);
    if (file.getMimeType() !== MimeType.PDF) continue;
    blobs.push(file.getBlob().setName(file.getName()));
    names.push(file.getName());
  }
  return { blobs, names };
}

function extractDriveFileId_(s) {
  if (!s) return "";
  const m1 = s.match(/\/file\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m1) return m1[1];
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  return "";
}

function findHeaderIndex_(headerRow, candidates) {
  const norm = headerRow.map(x => String(x || "").trim().toLowerCase());
  for (const c of candidates) {
    const idx = norm.indexOf(String(c).trim().toLowerCase());
    if (idx !== -1) return idx;
  }
  return -1;
}


/*************************************************
 * 05_資料送付リスト 自動転記（アド・クリックス様用）
 * リスト（寺社仏閣）・リスト（社福）の「コール状況①〜⑤」に
 * 「資料送付」または「資料請求」が入った行を「資料送付リスト」へ転記する。
 *  - 自動：onEdit（入力した瞬間に転記）
 *  - 一括：メニュー「コールログ」→「資料送付リストへ一括転記」
 *  - 同じ企業（企業名＋TEL）は重複させず、最新内容で上書き
 *  - E列＝資料送付／資料請求になった日付、F列＝「資料送付」or「資料請求」（最新の回）、G列＝転記元シート名
 *  - 資料送付リストのA〜D列には書き込まない
 *************************************************/

const SS_CFG = {
  sources: [
    { name: 'リスト（寺社仏閣）', category: '寺社仏閣' },
    { name: 'リスト（社福）',     category: '社福' }
  ],
  dest: '資料送付リスト',
  keywords: ['資料送付', '資料請求'],   // どちらかが入ったら転記
  headerRow: 1,
  fixedHeaders: ['企業名', '住所', '担当者名', '役職', '部署', 'TEL', 'Mail'],
  blockStartHeader: '担当①',   // ここから右（①〜⑤の架電履歴）はまとめて転記
  statusPrefix: 'コール状況',
  destCol: {                   // 資料送付リストの列番号（1始まり）
    sofuDate: 5,               // E：資料送付になった日付
    label: 6,                  // F：「資料送付」or「資料請求」
    sourceSheet: 7             // G：転記元シート名
  }
};

/********* 自動転記（シンプルトリガー） *********/
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const src = SS_CFG.sources.find(s => s.name === sh.getName());
    if (!src) return;

    const startRow = Math.max(e.range.getRow(), SS_CFG.headerRow + 1);
    const endRow = e.range.getLastRow();
    if (endRow < startRow) return;

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(10000)) return;
    try {
      const items = collectItems_(sh, src, startRow, endRow - startRow + 1);
      if (items.length === 0) return;
      const res = upsertToDest_(items);
      if (res.added > 0) {
        sh.getParent().toast(res.added + '件を資料送付リストに転記しました', '資料送付', 3);
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error('資料送付 自動転記エラー: ' + err);
  }
}

/********* 一括転記（既存分もまとめて） *********/
function syncShiryoSofuAll() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let items = [];
    SS_CFG.sources.forEach(src => {
      const sh = ss.getSheetByName(src.name);
      if (!sh) throw new Error('シートが見つかりません：' + src.name);
      const lastRow = sh.getLastRow();
      if (lastRow <= SS_CFG.headerRow) return;
      items = items.concat(collectItems_(sh, src, SS_CFG.headerRow + 1, lastRow - SS_CFG.headerRow));
    });
    const res = upsertToDest_(items);
    ui.alert('資料送付リストへ転記しました\n\n新規追加：' + res.added + '件\n更新：' + res.updated + '件');
  } catch (err) {
    ui.alert('エラーが発生しました：\n' + err);
  }
}

/********* 元シートから「資料送付」の行を集める *********/
function collectItems_(sh, src, startRow, numRows) {
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(SS_CFG.headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const fixedIdx = SS_CFG.fixedHeaders.map(h => header.indexOf(h)); // 最初に出てくる列を採用
  const blockStart = header.indexOf(SS_CFG.blockStartHeader);
  const statusIdx = [];
  header.forEach((h, i) => { if (h.indexOf(SS_CFG.statusPrefix) === 0) statusIdx.push(i); });
  if (blockStart < 0 || statusIdx.length === 0) throw new Error(src.name + ' のヘッダーが想定と違います');

  const values = sh.getRange(startRow, 1, numRows, lastCol).getValues();
  const items = [];
  values.forEach(row => {
    const matchKw = v => SS_CFG.keywords.find(k => String(v).indexOf(k) !== -1);
    const hit = statusIdx.some(i => matchKw(row[i]));
    if (!hit) return;

    const fixed = {};
    SS_CFG.fixedHeaders.forEach((h, j) => { fixed[h] = fixedIdx[j] >= 0 ? row[fixedIdx[j]] : ''; });
    if (String(fixed['企業名']).trim() === '') return;

    // 「資料送付／資料請求」になった日付と種類（複数回ある場合は一番右＝最新の回）
    let sofuDate = '';
    let label = '';
    statusIdx.forEach(i => {
      const kw = matchKw(row[i]);
      if (kw) {
        sofuDate = row[i - 2]; // 担当｜日付｜時間｜コール状況 の並び
        label = kw;
      }
    });

    items.push({
      key: makeKey_(fixed['企業名'], fixed['TEL']),
      fixed: fixed,
      block: row.slice(blockStart),
      sofuDate: sofuDate,
      label: label,
      sourceSheet: src.name
    });
  });
  return items;
}

/********* 資料送付リストへ追加 or 上書き *********/
function upsertToDest_(items) {
  const result = { added: 0, updated: 0 };
  if (items.length === 0) return result;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dest = ss.getSheetByName(SS_CFG.dest);
  if (!dest) throw new Error('シートが見つかりません：' + SS_CFG.dest);

  const lastCol = dest.getLastColumn();
  const header = dest.getRange(SS_CFG.headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const fixedIdx = SS_CFG.fixedHeaders.map(h => header.indexOf(h));
  const companyIdx = header.indexOf('企業名');
  const telIdx = header.indexOf('TEL');
  const blockStart = header.indexOf(SS_CFG.blockStartHeader);
  if (companyIdx < 0 || blockStart < 0) throw new Error(SS_CFG.dest + ' のヘッダーが想定と違います');

  const C = SS_CFG.destCol;
  const writeFrom = C.sofuDate;              // E列から右だけ書き込む（A〜Dは触らない）
  const width = lastCol - writeFrom + 1;

  // 既存データのキー一覧
  const lastRow = dest.getLastRow();
  const keyToRow = {};
  let lastFilled = SS_CFG.headerRow;
  if (lastRow > SS_CFG.headerRow) {
    const data = dest.getRange(SS_CFG.headerRow + 1, 1, lastRow - SS_CFG.headerRow, lastCol).getValues();
    data.forEach((r, i) => {
      const company = String(r[companyIdx]).trim();
      if (company === '') return;
      const rowNum = SS_CFG.headerRow + 1 + i;
      keyToRow[makeKey_(company, telIdx >= 0 ? r[telIdx] : '')] = rowNum;
      lastFilled = rowNum;
    });
  }

  const newRows = [];
  const newKeys = {};

  items.forEach(it => {
    // E列以降の1行分をつくる
    const out = new Array(width).fill('');
    const put = (col1, v) => { const k = col1 - writeFrom; if (k >= 0 && k < width) out[k] = v; };
    put(C.sofuDate, it.sofuDate);
    put(C.label, it.label);
    put(C.sourceSheet, it.sourceSheet);
    SS_CFG.fixedHeaders.forEach((h, j) => { if (fixedIdx[j] >= 0) put(fixedIdx[j] + 1, it.fixed[h]); });
    const blockLen = Math.min(it.block.length, lastCol - blockStart);
    for (let b = 0; b < blockLen; b++) put(blockStart + 1 + b, it.block[b]);

    if (keyToRow[it.key]) {
      dest.getRange(keyToRow[it.key], writeFrom, 1, width).setValues([out]);
      result.updated++;
    } else if (newKeys[it.key] !== undefined) {
      newRows[newKeys[it.key]].out = out;     // 同じ企業が元リストに2回ある場合
    } else {
      newKeys[it.key] = newRows.length;
      newRows.push({ out: out });
    }
  });

  if (newRows.length > 0) {
    const start = lastFilled + 1;
    // E列以降をまとめて書き込み
    const block = newRows.map(n => n.out);
    dest.getRange(start, writeFrom, block.length, block[0].length).setValues(block);

    // B列のチェックボックスが無い行には追加（メルマガ転記用）
    const bRange = dest.getRange(start, 2, block.length, 1);
    const dv = bRange.getDataValidations();
    dv.forEach((r, i) => { if (!r[0]) dest.getRange(start + i, 2).insertCheckboxes(); });

    result.added = newRows.length;
  }
  return result;
}

/********* 重複判定キー（企業名＋TEL数字のみ） *********/
function makeKey_(company, tel) {
  const c = String(company || '').replace(/[\s　]/g, '');
  const t = String(tel || '').replace(/[^0-9０-９]/g, '');
  return c + '|' + t;
}

/*************************************************
 * 06_資料請求／資料送付 通知（アド・クリックス様用）
 * リスト（寺社仏閣）・リスト（社福）の「コール状況①〜⑤」に
 * 「資料請求」または「資料送付」が入ったら、その会社の情報を通知する。
 *  - LINE（クライアント様グループ）：企業名・住所・TEL だけ
 *  - メール／Slack（社内）：上記＋シート名・行番号・何回目・担当・日付
 *  - 同じ会社×同じ回（①〜⑤）は1回しか通知しない（並べ替えても二重通知しない）
 *  - メール・外部送信が必要なので「インストール型トリガー」で動く
 *    → メニュー「コールログ」→「資料請求通知：自動通知をONにする（初回のみ）」
 *  - ★ の箇所が埋まっていない通知先は自動でお休み（エラーにならない）
 *************************************************/

const NOTIFY_CFG = {
  sources: ['リスト（寺社仏閣）', 'リスト（社福）'],
  keywords: ['資料請求', '資料送付'],   // どちらかが含まれていれば通知（【資料送付】もOK）
  statusPrefix: 'コール状況',           // 「コール状況①〜⑤」をすべて監視
  headers: { company: '企業名', address: '住所', phone: 'TEL' },

  // ① LINE（クライアント様とのグループに push。一斉配信は使わない）
  LINE: {
    enabled: false,                     // Messaging API の設定が済んだら true に
    channelAccessToken: '★チャネルアクセストークン（長期）',
    to: '★クライアント様グループのグループID（Cから始まる）',
  },

  // ② メール（複数OK）
  EMAIL: {
    enabled: true,
    to: ['★クライアント様のアドレス', '★社内のアドレス'],
  },

  // ③ Slack（Pj-アドクリックス の Incoming Webhook）
  SLACK: {
    enabled: true,
    webhookUrl: '★https://hooks.slack.com/... を貼る',
  },

  // 通知でエラーが起きたときに知らせる社内アドレス
  ADMIN_EMAIL: '★社内の管理用アドレス',
};

/********* 編集トリガー本体（インストール型から呼ばれる） *********/
function onEditNotify(e) {
  const errors = [];
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (NOTIFY_CFG.sources.indexOf(sh.getName()) === -1) return;

    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const c0 = e.range.getColumn();
    const c1 = c0 + e.range.getNumColumns() - 1;
    const statusCols = [];
    header.forEach((h, i) => {
      const col = i + 1;
      if (h.indexOf(NOTIFY_CFG.statusPrefix) === 0 && col >= c0 && col <= c1) statusCols.push(col);
    });
    if (statusCols.length === 0) return; // コール状況の列以外の編集は無視

    const idx = {
      company: header.indexOf(NOTIFY_CFG.headers.company),
      address: header.indexOf(NOTIFY_CFG.headers.address),
      phone:   header.indexOf(NOTIFY_CFG.headers.phone),
    };
    const props = PropertiesService.getDocumentProperties();
    const r0 = Math.max(e.range.getRow(), 2);
    const r1 = e.range.getLastRow();

    for (let r = r0; r <= r1; r++) {
      const row = sh.getRange(r, 1, 1, lastCol).getValues()[0];
      statusCols.forEach(col => {
        const status = String(row[col - 1]).trim();
        const keyword = NOTIFY_CFG.keywords.find(k => status.indexOf(k) !== -1);
        if (!keyword) return;

        const company = notifyCell_(row, idx.company);
        const phone = notifyCell_(row, idx.phone);
        const doneKey = 'notified|' + makeKey_(company, phone) + '|' + header[col - 1];
        if (props.getProperty(doneKey)) return; // 通知済み

        const info = {
          sheetName: sh.getName(),
          row: r,
          round: header[col - 1],                  // 例：コール状況③
          keyword: keyword,
          company: company,
          address: notifyCell_(row, idx.address),
          phone: phone,
          caller: col >= 4 ? notifyCell_(row, col - 4) : '',   // 担当｜日付｜時間｜コール状況 の並び
          date:   col >= 3 ? notifyCell_(row, col - 3) : '',
        };
        const result = notifySendAll_(info);
        if (result.sent > 0) props.setProperty(doneKey, notifyStamp_());
        result.errors.forEach(x => errors.push(info.company + '：' + x));
      });
    }
  } catch (err) {
    errors.push(String((err && err.stack) || err));
  }
  if (errors.length) notifyReportError_(errors.join('\n'));
}

/********* 有効な通知先すべてに送る（1つ失敗しても他は続行） *********/
function notifySendAll_(info) {
  const ch = notifyChannels_();
  const out = { sent: 0, errors: [] };
  const run = (name, on, fn) => {
    if (!on) return;
    try { fn(info); out.sent++; } catch (e) { out.errors.push(name + ': ' + e.message); }
  };
  run('LINE', ch.line, notifySendLine_);
  run('メール', ch.email, notifySendEmail_);
  run('Slack', ch.slack, notifySendSlack_);
  return out;
}

/** ★が残っている（未設定の）通知先はお休みにする */
function notifyChannels_() {
  const set = v => !!v && String(v).indexOf('★') === -1;
  const C = NOTIFY_CFG;
  return {
    line:  C.LINE.enabled && set(C.LINE.channelAccessToken) && set(C.LINE.to),
    email: C.EMAIL.enabled && C.EMAIL.to.filter(set).length > 0,
    slack: C.SLACK.enabled && set(C.SLACK.webhookUrl),
  };
}

/** ① LINE：クライアント様向け（企業名・住所・TELだけ） */
function notifySendLine_(info) {
  const text =
    '📩 テレアポにて「' + info.keyword + '」のご承諾をいただきました。\n' +
    '下記のお客様へ資料のご送付をお願いいたします。\n\n' +
    '■ 企業名：' + info.company + '\n' +
    '■ 住所：' + info.address + '\n' +
    '■ 電話番号：' + info.phone;
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + NOTIFY_CFG.LINE.channelAccessToken },
    payload: JSON.stringify({ to: NOTIFY_CFG.LINE.to, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('送信失敗(' + res.getResponseCode() + ') ' + res.getContentText());
}

/** ② メール：追跡用の情報つき */
function notifySendEmail_(info) {
  const to = NOTIFY_CFG.EMAIL.to.filter(v => v && v.indexOf('★') === -1).join(',');
  const subject = '【' + info.keyword + '】' + info.company + ' 様（' + info.sheetName + '）';
  GmailApp.sendEmail(to, subject, notifyInternalBody_(info));
}

/** ③ Slack：追跡用の情報つき */
function notifySendSlack_(info) {
  const res = UrlFetchApp.fetch(NOTIFY_CFG.SLACK.webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: notifyInternalBody_(info) }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('送信失敗(' + res.getResponseCode() + ') ' + res.getContentText());
}

/** 社内向け本文（メール・Slack共通） */
function notifyInternalBody_(info) {
  return [
    '【' + info.keyword + 'の通知】',
    '企業名：' + info.company,
    '住所：' + info.address,
    'TEL：' + info.phone,
    '──────',
    'シート：' + info.sheetName + '（' + info.row + '行目）',
    '回：' + info.round + (info.date ? '／日付：' + info.date : '') + (info.caller ? '／担当：' + info.caller : ''),
  ].join('\n');
}

function notifyCell_(row, i) {
  if (i < 0 || i >= row.length) return '';
  const v = row[i];
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
  return (v === '' || v === null || v === undefined) ? '' : String(v).trim();
}

function notifyStamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
}

function notifyReportError_(message) {
  console.error('資料請求通知エラー: ' + message);
  const admin = NOTIFY_CFG.ADMIN_EMAIL;
  if (!admin || admin.indexOf('★') !== -1) return;
  try {
    GmailApp.sendEmail(admin, '【資料請求通知】エラーが発生しました',
      '通知処理でエラーが発生しました。内容をご確認ください。\n\n' + message);
  } catch (e) {
    console.error('エラー通知メールも失敗: ' + e);
  }
}

/********* 初回のみ：自動通知をONにする（メニュー or エディタから実行） *********/
function setupNotifyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'onEditNotify' || fn === 'onResourceRequestEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditNotify').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  const ch = notifyChannels_();
  const on = [ch.line && 'LINE', ch.email && 'メール', ch.slack && 'Slack'].filter(Boolean);
  SpreadsheetApp.getActive().toast(
    '自動通知をONにしました（送り先：' + (on.length ? on.join('・') : 'まだ未設定') + '）', '資料請求通知', 8);
}

/********* 動作確認：ダミーの会社で1通ずつ送る *********/
function testNotify() {
  const ch = notifyChannels_();
  if (!ch.line && !ch.email && !ch.slack) {
    SpreadsheetApp.getActive().toast('送り先がまだ設定されていません（NOTIFY_CFG の ★ を埋めてください）', '資料請求通知', 8);
    return;
  }
  const result = notifySendAll_({
    sheetName: NOTIFY_CFG.sources[0], row: 999, round: 'コール状況①（テスト）',
    keyword: '資料送付', company: 'テスト商事株式会社', address: '東京都千代田区テスト1-2-3',
    phone: '03-0000-0000', caller: 'テスト担当', date: notifyStamp_(),
  });
  const msg = 'テスト送信：成功 ' + result.sent + '件' + (result.errors.length ? '／失敗：' + result.errors.join(' / ') : '');
  SpreadsheetApp.getActive().toast(msg, '資料請求通知', 10);
  if (result.errors.length) throw new Error(msg);
}
