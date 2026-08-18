const SOURCE_SHEET = 'リスト';
const MAIL_SHEET = 'メルマガ送信リスト';

const CALLLOG_SHEET = '架電ログ';
const KPI_MONTHLY_SHEET = '月別推移';
const KPI_HOURLY_SHEET = '時間帯分析';
const KPI_STATUS_SHEET = '状況内訳';

// ▼追加：名刺リスト → メルマガ送信リスト 用
const CARD_SOURCE_SHEET = '名刺リスト';
const CARD_EXCLUDE_COL = 1;  // A列：不要はチェック
const CARD_COMPANY_COL = 3;  // C列：会社名
const CARD_DEPT_COL = 4;     // D列：部署
const CARD_ROLE_COL = 5;     // E列：役職
const CARD_NAME_COL = 6;     // F列：氏名
const CARD_EMAIL_COL = 9;    // I列：メールアドレス

const MAIL_BODY_CELL = 'K2';
const MAIL_FILE_LINK_CELL = 'L2';
const TEST_MODE_CELL = 'M2';
const TEST_EMAIL_CELL = 'N2';
const TEST_COMPANY_CELL = 'O2';
const TEST_PERSON_CELL = 'P2';

const CALL_STATUS_LIST = [
  '【関心あり】',
  '【タイミングが悪い】',
  '【断り】',
  '【定期フォロー】',
  '【アポ獲得】',
  '【資料送付】',
  'TEL禁',
  '【不在】',
  'その他'
];

const CALL_BLOCKS = [
  { staffCol: 10, dateCol: 11, timeCol: 12, statusCol: 13, nextActionCol: 14, nextActionDateCol: 15, commentCol: 16, mailCol: 17, contactCol: 18, attempt: 1 },
  { staffCol: 19, dateCol: 20, timeCol: 21, statusCol: 22, nextActionCol: 23, nextActionDateCol: 24, commentCol: 25, mailCol: 26, contactCol: 27, attempt: 2 },
  { staffCol: 28, dateCol: 29, timeCol: 30, statusCol: 31, nextActionCol: 32, nextActionDateCol: 33, commentCol: 34, mailCol: 35, contactCol: 36, attempt: 3 },
  { staffCol: 37, dateCol: 38, timeCol: 39, statusCol: 40, nextActionCol: 41, nextActionDateCol: 42, commentCol: 43, mailCol: 44, contactCol: 45, attempt: 4 },
  { staffCol: 46, dateCol: 47, timeCol: 48, statusCol: 49, nextActionCol: 50, nextActionDateCol: 51, commentCol: 52, mailCol: 53, contactCol: 54, attempt: 5 }
];

const KPI_CATCH_STATUSES = [
  '【関心あり】',
  '【タイミングが悪い】',
  '【断り】',
  '【定期フォロー】',
  '【アポ獲得】',
  '【資料送付】',
  'TEL禁'
];

const KPI_INCOMING_STATUSES = [
  '【関心あり】',
  '【タイミングが悪い】',
  '【断り】',
  '【定期フォロー】',
  '【アポ獲得】',
  '【資料送付】',
  'TEL禁',
  'その他'
];

const KPI_CONNECTED_STATUSES = [
  '【関心あり】',
  '【タイミングが悪い】',
  '【断り】',
  '【定期フォロー】',
  '【アポ獲得】',
  '【資料送付】'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('メルマガ')
    .addItem('メルマガを送信', 'sendNewsletter')
    .addSeparator()
    .addItem('名刺リストを一括転記（初回のみ）', 'transferAllCardListToMailList')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;

  updateCallLogByEdit_(e);
  updateKpiReports_();

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const startRow = e.range.getRow();
  const startCol = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();

  if (startRow < 2) return;

  // リスト → メルマガ送信リスト（既存）
  if (sheetName === SOURCE_SHEET) {
    if (startCol <= 1 && 1 <= startCol + numCols - 1) {
      for (let r = 0; r < numRows; r++) {
        transferListToMailList_(startRow + r);
      }
    }
  }

  // 名刺リスト → メルマガ送信リスト（追加）
  if (sheetName === CARD_SOURCE_SHEET) {
    if (startCol <= CARD_EXCLUDE_COL && CARD_EXCLUDE_COL <= startCol + numCols - 1) {
      for (let r = 0; r < numRows; r++) {
        transferCardListToMailList_(startRow + r);
      }
    }
  }
}

/* =========================
   リスト → メルマガ送信リスト
========================= */

function transferListToMailList_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(SOURCE_SHEET);
  const dest = ss.getSheetByName(MAIL_SHEET);

  const checkbox = source.getRange(row, 1).getValue();
  if (checkbox !== true) return;

  const values = [
    source.getRange(row, 2).getValue(), // B 企業名 → F
    source.getRange(row, 4).getValue(), // D 担当者名 → G
    source.getRange(row, 5).getValue(), // E 役職 → H
    source.getRange(row, 6).getValue(), // F 部署 → I
    source.getRange(row, 8).getValue()  // H Mail → J
  ];

  const destRow = findFirstEmptyRow_(dest, 6, 2);
  dest.getRange(destRow, 6, 1, values.length).setValues([values]);
}

/* =========================
   名刺リスト → メルマガ送信リスト（追加）
========================= */

function transferCardListToMailList_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(CARD_SOURCE_SHEET);
  const dest = ss.getSheetByName(MAIL_SHEET);
  if (!source || !dest) return;

  const excludeChecked = source.getRange(row, CARD_EXCLUDE_COL).getValue(); // A 不要はチェック
  if (excludeChecked === true) return; // 不要にチェックがある人は対象外

  const email = source.getRange(row, CARD_EMAIL_COL).getValue(); // I メールアドレス
  if (!email) return; // メールアドレスが無い人は送りようがないので対象外

  const values = [
    source.getRange(row, CARD_COMPANY_COL).getValue(), // C 会社名 → F
    source.getRange(row, CARD_DEPT_COL).getValue(),     // D 部署 → G
    source.getRange(row, CARD_ROLE_COL).getValue(),     // E 役職 → H
    source.getRange(row, CARD_NAME_COL).getValue(),     // F 氏名 → I
    email                                                // I メールアドレス → J
  ];

  const destRow = findFirstEmptyRow_(dest, 6, 2);
  dest.getRange(destRow, 6, 1, values.length).setValues([values]);
}

// 名刺リストに既に入っている分をまとめて転記するための手動メニュー用関数
// ※重複登録を防ぐため、実行は1回だけにしてください
function transferAllCardListToMailList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(CARD_SOURCE_SHEET);
  if (!source) {
    SpreadsheetApp.getUi().alert(`「${CARD_SOURCE_SHEET}」シートが見つかりませんでした。`);
    return;
  }

  const lastRow = source.getLastRow();
  if (lastRow < 2) return;

  for (let row = 2; row <= lastRow; row++) {
    transferCardListToMailList_(row);
  }

  SpreadsheetApp.getUi().alert('名刺リストの一括転記が完了しました。');
}

/* =========================
   メルマガ本文の装飾（追加）
========================= */

/**
 * メール本文のHTML特殊文字エスケープ（&, <, > 等がタグ崩れしないように）
 */
function escapeMailHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * 本文中の記号で囲まれた部分を装飾する。
 * ・**文章**（アスタリスク2つ）→ 赤字・太字・大きめ（色もサイズも変えて強調したいとき）
 * ・*文章*（アスタリスク1つ）  → 色はそのまま・太字・大きめ（サイズだけ大きくしたいとき）
 * 例：本日はお時間いただきありがとうございました。**先着5社様限定**のご案内です。
 *     *来月から新プランを開始します*
 * ※ ** を先に処理してから残った単独の * を処理するので、**と*を混ぜて使っても崩れない
 * ※ 文字サイズは font-size の倍率（例：1.4em）を変えるだけで簡単に調整できる
 */
function applyMailHighlightMarkup_(escapedText) {
  let result = escapedText.replace(
    /\*\*([\s\S]+?)\*\*/g,
    '<span style="color:#d32f2f;font-weight:bold;font-size:1.4em;line-height:1.4;">$1</span>'
  );

  result = result.replace(
    /\*([\s\S]+?)\*/g,
    '<span style="font-weight:bold;font-size:1.4em;line-height:1.4;">$1</span>'
  );

  return result;
}

/**
 * プレーン本文（挨拶文込み）から、装飾込みのHTMLメール本文を組み立てる
 */
function buildMailHtmlBody_(plainBody) {
  let html = escapeMailHtml_(plainBody);
  html = applyMailHighlightMarkup_(html);
  html = html.replace(/\n/g, '<br>');

  return (
    `<div style="font-family:'Hiragino Kaku Gothic ProN','Meiryo',sans-serif;` +
    `font-size:14px;line-height:1.8;color:#333333;">${html}</div>`
  );
}

/* =========================
   メルマガ送信
========================= */

function sendNewsletter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAIL_SHEET);

  const testMode = sheet.getRange(TEST_MODE_CELL).getValue() === true;
  const testEmail = sheet.getRange(TEST_EMAIL_CELL).getValue();
  const bodyText = sheet.getRange(MAIL_BODY_CELL).getValue();
  const fileLink = sheet.getRange(MAIL_FILE_LINK_CELL).getValue();

  const subject = '【件名】';
  const attachments = getAttachmentsFromDriveLink_(fileLink);

  if (testMode) {
    const testCompany = sheet.getRange(TEST_COMPANY_CELL).getValue();
    const testPerson = sheet.getRange(TEST_PERSON_CELL).getValue();

    const body = `${testCompany}\n${testPerson} 様\n\n${bodyText}`;
    const htmlBody = buildMailHtmlBody_(body);

    GmailApp.sendEmail(testEmail, subject, body, {
      attachments: attachments,
      htmlBody: htmlBody
    });

    SpreadsheetApp.getUi().alert('テストメールを送信しました。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const today = new Date();
  let sentCount = 0;

  values.forEach((row, i) => {
    const rowNumber = i + 2;

    const sentDate = row[0]; // A
    const company = row[5];  // F
    const person = row[6];   // G
    const email = row[9];    // J

    if (sentDate || !email) return;

    const body = `${company}\n${person} 様\n\n${bodyText}`;
    const htmlBody = buildMailHtmlBody_(body);

    GmailApp.sendEmail(email, subject, body, {
      attachments: attachments,
      htmlBody: htmlBody
    });

    sheet.getRange(rowNumber, 1).setValue(today);
    sentCount++;
  });

  sheet.getRange(TEST_MODE_CELL).setValue(true);
  SpreadsheetApp.getUi().alert(`${sentCount}件送信しました。`);
}

/* =========================
   架電ログ 自動更新
========================= */

function updateCallLogByEdit_(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SOURCE_SHEET) return;

  const startRow = e.range.getRow();
  const startCol = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();

  if (startRow < 2) return;

  const editedCols = [];
  for (let c = 0; c < numCols; c++) {
    editedCols.push(startCol + c);
  }

  for (let r = 0; r < numRows; r++) {
    const row = startRow + r;

    CALL_BLOCKS.forEach(block => {
      const relatedCols = [
        block.staffCol,
        block.dateCol,
        block.timeCol,
        block.statusCol,
        block.nextActionCol,
        block.nextActionDateCol,
        block.commentCol,
        block.mailCol,
        block.contactCol
      ];

      if (!relatedCols.some(col => editedCols.includes(col))) return;

      upsertCallLogRow_(sheet, row, block);
    });
  }
}

function upsertCallLogRow_(source, row, block) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(CALLLOG_SHEET);
  if (!log) log = ss.insertSheet(CALLLOG_SHEET);

  ensureCallLogSheet_(log);

  const company = source.getRange(row, 2).getValue(); // B
  const baseContact = source.getRange(row, 4).getValue(); // D

  const staff = source.getRange(row, block.staffCol).getValue();
  const date = source.getRange(row, block.dateCol).getValue();
  const time = source.getRange(row, block.timeCol).getValue();
  const status = normalizeCallStatus_(source.getRange(row, block.statusCol).getValue());
  const nextAction = source.getRange(row, block.nextActionCol).getValue();
  const nextActionDate = source.getRange(row, block.nextActionDateCol).getValue();
  const comment = source.getRange(row, block.commentCol).getValue();
  const mail = source.getRange(row, block.mailCol).getValue();
  const contact = source.getRange(row, block.contactCol).getValue();

  const key = `${SOURCE_SHEET}-${row}-${block.attempt}`;
  const logRow = findCallLogRowByKey_(log, key);

  if (!staff && !date && !time && !status && !nextAction && !nextActionDate && !comment && !mail && !contact) {
    if (logRow) log.deleteRow(logRow);
    return;
  }

  const timeText = formatTime_(time);
  const hourBand = getHourBand_(timeText);

  const values = [[
    SOURCE_SHEET,
    row,
    company,
    contact || baseContact,
    staff,
    date,
    timeText,
    hourBand,
    status || 'その他',
    nextAction,
    nextActionDate,
    comment,
    mail,
    block.attempt,
    key
  ]];

  if (logRow) {
    log.getRange(logRow, 1, 1, values[0].length).setValues(values);
  } else {
    const destRow = findFirstEmptyRow_(log, 1, 2);
    log.getRange(destRow, 1, 1, values[0].length).setValues(values);
  }

  log.getRange('F:F').setNumberFormat('yyyy/mm/dd');
  log.getRange('K:K').setNumberFormat('yyyy/mm/dd');
}

function ensureCallLogSheet_(sheet) {
  const header = [
    '元シート',
    '元行',
    '企業名',
    '担当者名',
    '架電担当',
    '日付',
    '時間',
    '時間帯',
    'コール状況',
    '次回アクション',
    '次回アクション日',
    'コメント',
    'Mail',
    '回数',
    'キー'
  ];

  const current = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const isEmpty = current.every(v => v === '');

  if (isEmpty) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
}

function findCallLogRowByKey_(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  // O列：キー
  const keys = sheet.getRange(2, 15, lastRow - 1, 1).getValues();

  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === String(key)) {
      return i + 2;
    }
  }

  return null;
}

/* =========================
   KPI 自動更新
========================= */

function updateKpiReports_() {
  const records = getCallLogRecordsForKpi_();

  buildMonthlyTrend_(records);
  buildHourlyAnalysis_(records);
  buildStatusBreakdown_(records);
}

function getCallLogRecordsForKpi_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CALLLOG_SHEET);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

  return values.map(row => {
    const date = row[5];     // F
    const time = row[6];     // G
    const hourBand = row[7]; // H
    const status = row[8];   // I

    return {
      date,
      time,
      hourBand: String(hourBand || '').trim(),
      status: String(status || '').trim(),
      ym: toYearMonthForKpi_(date),
      hour: extractHourForKpi_(time, hourBand)
    };
  }).filter(r => r.status);
}

function buildMonthlyTrend_(records) {
  const sheet = getOrCreateSheet_(KPI_MONTHLY_SHEET);
  const months = [...new Set(records.map(r => r.ym).filter(Boolean))].sort();

  const header = [
    '月',
    'コール数',
    'キャッチ数',
    'キャッチ率',
    'アポ取得数',
    'アポ化率(キャッチ→アポ)',
    '目標値',
    'コール数',
    'アポ取得数',
    'アポ率',
    '資料送付数',
    '資料送付率',
    '未分類率',
    '着電数',
    'アポ数',
    '着電率',
    '接電数',
    'アポ数',
    '接電率'
  ];

  const rows = [header];

  months.forEach(month => {
    const monthRecords = records.filter(r => r.ym === month);
    rows.push(makeKpiRow_(month, monthRecords));
  });

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, header.length).setValues(rows);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');

  if (rows.length > 1) {
    [4, 6, 10, 12, 13, 16, 19].forEach(col => {
      sheet.getRange(2, col, rows.length - 1, 1).setNumberFormat('0.0%');
    });
  }

  sheet.autoResizeColumns(1, header.length);
}

function buildHourlyAnalysis_(records) {
  const sheet = getOrCreateSheet_(KPI_HOURLY_SHEET);

  const header = [
    '時間帯',
    'コール数',
    'キャッチ数',
    'キャッチ率',
    'アポ取得数',
    'アポ化率(キャッチ→アポ)',
    '資料送付数',
    '資料送付率',
    '未分類率',
    '着電数',
    'アポ数',
    '着電率',
    '接電数',
    'アポ数',
    '接電率'
  ];

  const rows = [header];

  for (let h = 0; h < 24; h++) {
    const hour = String(h).padStart(2, '0');
    const label = `${h}:00`;

    const hourRecords = records.filter(r => r.hour === hour);
    const kpi = calcKpi_(hourRecords);

    rows.push([
      label,
      kpi.callCount,
      kpi.catchCount,
      kpi.catchRate,
      kpi.apoCount,
      kpi.apoRateFromCatch,
      kpi.mailCount,
      kpi.mailRate,
      kpi.unclassifiedRate,
      kpi.incomingCount,
      kpi.apoCount,
      kpi.incomingApoRate,
      kpi.connectedCount,
      kpi.apoCount,
      kpi.connectedApoRate
    ]);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, header.length).setValues(rows);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');

  [4, 6, 8, 9, 12, 15].forEach(col => {
    sheet.getRange(2, col, 24, 1).setNumberFormat('0.0%');
  });

  sheet.autoResizeColumns(1, header.length);
}

function buildStatusBreakdown_(records) {
  const sheet = getOrCreateSheet_(KPI_STATUS_SHEET);

  const rows = [['コール状況', '件数']];

  CALL_STATUS_LIST.forEach(status => {
    rows.push([
      status,
      records.filter(r => r.status === status).length
    ]);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange('A1:B1').setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
}

function makeKpiRow_(label, records) {
  const kpi = calcKpi_(records);

  return [
    label,
    kpi.callCount,
    kpi.catchCount,
    kpi.catchRate,
    kpi.apoCount,
    kpi.apoRateFromCatch,
    '',
    kpi.callCount,
    kpi.apoCount,
    kpi.apoRate,
    kpi.mailCount,
    kpi.mailRate,
    kpi.unclassifiedRate,
    kpi.incomingCount,
    kpi.apoCount,
    kpi.incomingApoRate,
    kpi.connectedCount,
    kpi.apoCount,
    kpi.connectedApoRate
  ];
}

function calcKpi_(records) {
  const callCount = records.length;

  const catchCount = records.filter(r => KPI_CATCH_STATUSES.includes(r.status)).length;
  const apoCount = records.filter(r => r.status === '【アポ獲得】').length;
  const mailCount = records.filter(r => r.status === '【資料送付】').length;
  const unclassifiedCount = records.filter(r => r.status === 'その他').length;
  const incomingCount = records.filter(r => KPI_INCOMING_STATUSES.includes(r.status)).length;
  const connectedCount = records.filter(r => KPI_CONNECTED_STATUSES.includes(r.status)).length;

  return {
    callCount,
    catchCount,
    apoCount,
    mailCount,
    unclassifiedCount,
    incomingCount,
    connectedCount,
    catchRate: callCount > 0 ? catchCount / callCount : 0,
    apoRateFromCatch: catchCount > 0 ? apoCount / catchCount : 0,
    apoRate: callCount > 0 ? apoCount / callCount : 0,
    mailRate: callCount > 0 ? mailCount / callCount : 0,
    unclassifiedRate: callCount > 0 ? unclassifiedCount / callCount : 0,
    incomingApoRate: incomingCount > 0 ? apoCount / incomingCount : 0,
    connectedApoRate: connectedCount > 0 ? apoCount / connectedCount : 0
  };
}

/* =========================
   共通関数
========================= */

function normalizeCallStatus_(value) {
  let text = String(value || '').trim();
  if (!text) return '';

  const matched = CALL_STATUS_LIST.find(status => text.includes(status));
  if (matched) return matched;

  if (text.includes('アポ')) return '【アポ獲得】';
  if (text.includes('資料')) return '【資料送付】';
  if (text.includes('不在')) return '【不在】';
  if (text.includes('断り')) return '【断り】';
  if (text.includes('関心')) return '【関心あり】';
  if (text.includes('定期')) return '【定期フォロー】';
  if (text.includes('TEL禁')) return 'TEL禁';

  return 'その他';
}

function formatTime_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
      'HH:mm'
    );
  }

  if (typeof value === 'number') {
    const totalMinutes = Math.round(value * 24 * 60);
    const hour = Math.floor(totalMinutes / 60) % 24;
    const minute = totalMinutes % 60;
    return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
  }

  const text = String(value);
  const match = text.match(/(\d{1,2})[:：](\d{2})/);
  if (match) return `${String(match[1]).padStart(2, '0')}:${match[2]}`;

  return '';
}

function getHourBand_(timeText) {
  if (!timeText) return '時間未入力';

  const hour = Number(timeText.split(':')[0]);
  if (isNaN(hour)) return '時間未入力';

  return `${hour}:00台`;
}

function toYearMonthForKpi_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
      'yyyy/MM'
    );
  }

  const text = String(value).trim();
  const match = text.match(/^(\d{4})[\/\-](\d{1,2})/);
  if (!match) return '';

  return `${match[1]}/${String(match[2]).padStart(2, '0')}`;
}

function extractHourForKpi_(time, hourBand) {
  if (time instanceof Date) {
    return Utilities.formatDate(
      time,
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
      'HH'
    );
  }

  const timeText = String(time || '').trim();
  const match = timeText.match(/^(\d{1,2})[:：]/);
  if (match) return String(match[1]).padStart(2, '0');

  const bandText = String(hourBand || '').trim();
  const bandMatch = bandText.match(/^(\d{1,2})/);
  if (bandMatch) return String(bandMatch[1]).padStart(2, '0');

  return '';
}

function getAttachmentsFromDriveLink_(link) {
  if (!link) return [];

  const fileId = extractDriveFileId_(link);
  if (!fileId) return [];

  const file = DriveApp.getFileById(fileId);
  return [file.getBlob()];
}

function extractDriveFileId_(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/
  ];

  for (const pattern of patterns) {
    const match = String(url).match(pattern);
    if (match) return match[1];
  }

  return null;
}

function findFirstEmptyRow_(sheet, targetCol, startRow) {
  const maxRows = sheet.getMaxRows();
  const values = sheet
    .getRange(startRow, targetCol, maxRows - startRow + 1, 1)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') return startRow + i;
  }

  return maxRows + 1;
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}
