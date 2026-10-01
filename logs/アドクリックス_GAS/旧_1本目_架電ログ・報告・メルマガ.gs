function rebuildCallLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sources = [
    { name: 'リスト（寺社仏閣）', prefix: '寺' },
    { name: 'リスト（社福）',   prefix: '福' }
  ];
  const rounds = [
    { n:1, caller:10, date:11, time:12, status:13, nextAction:14, nextDate:15, comment:16, contact:17, mail:18 },
    { n:2, caller:19, date:20, time:21, status:22, nextAction:23, nextDate:24, comment:25, contact:26, mail:27 },
    { n:3, caller:28, date:29, time:30, status:31, nextAction:32, nextDate:33, comment:34, mail:35, contact:36 },
    { n:4, caller:37, date:38, time:39, status:40, nextAction:41, nextDate:42, comment:43, mail:44, contact:45 },
    { n:5, caller:46, date:47, time:48, status:49, nextAction:50, nextDate:51, comment:52, mail:53, contact:54 }
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
const CATCH_ALWAYS = ['【アポ獲得】', '【資料送付】', '【定期フォロー】', '【関心あり】', '【決済者断り】'];
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
