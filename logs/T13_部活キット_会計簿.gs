/************************************************************************
 * 🎾 部活動まるごと管理キット【出欠＋会計】
 * セットアップ＆毎日の運用スクリプト（このファイル1つで完結）v1.1
 *
 * ── これは何？ ─────────────────────────────────────────────
 * 新しい空のスプレッドシートにこのファイルを貼り付けて、
 * 「setup」を1回実行すると、9枚のシート
 * （はじめに／名簿／今日の出欠／出欠簿／出欠サマリ／
 *   会計簿／部費徴収／年度末報告／設定）が自動で作られます。
 *
 * 毎日の出欠記録は、シート上部のメニュー「🎾部活キット」→
 * 「✅ 今日の出欠を記録する」を押すだけです。
 * 会計簿・部費徴収・年度末報告は、ボタン操作は不要です。
 * 数式が自動で計算するので、セルに入力するだけで動きます。
 *
 * ※ このファイルは v1.0（T12・出欠管理のみ）に「会計」機能を追加した
 *   ものです。大会エントリー管理・連絡テンプレ文例集は次工程で別途追加します。
 *
 * ── 使う人（先生）へ ─────────────────────────────────────
 * ・このコードを編集する必要はありません。貼り付けて実行するだけです。
 * ・外部のサービスには一切つながりません。あなたのGoogleアカウントの
 *   このスプレッドシートの中だけで動きます。
 * ・初回実行時に出る「このアプリは Google で確認されていません」という
 *   画面は正常です（自作スクリプト共通の表示）。「詳細」→「移動」で
 *   進めてください。
 *
 * ── 困ったとき ───────────────────────────────────────────
 * ・シートを壊してしまったら：壊れたシートを削除して、メニュー
 *   「🛠 初期セットアップ」を再実行すると、そのシートだけ作り直されます。
 *   （※「出欠簿」を消すと記録も消えます。消すときは「出欠サマリ」も
 *     一緒に削除してから再実行してください。会計簿・部費徴収・年度末報告は
 *     それぞれ独立しているので、1枚だけ削除→再実行でも安全です）
 ************************************************************************/

/* ===== 基本設定（シート名・行位置・色）。通常は変更不要 ===== */

var SHEET = {
  GUIDE:   'はじめに',
  ROSTER:  '名簿',
  TODAY:   '今日の出欠',
  LEDGER:  '出欠簿',
  SUMMARY: '出欠サマリ',
  ACCOUNT: '会計簿',
  DUES:    '部費徴収',
  REPORT:  '年度末報告',
  CONFIG:  '設定'
};

var ROSTER_FIRST_ROW = 3;    // 名簿の部員データが始まる行
var ROSTER_MAX       = 60;   // 対応する最大部員数
var ROSTER_LAST_ROW  = ROSTER_FIRST_ROW + ROSTER_MAX - 1;  // = 62行目
var LEDGER_FIRST_DATE_COL = 4;   // 出欠簿の最初の日付列（D列）
var LEDGER_MAX_COLS  = 320;      // 出欠簿の列数（活動日317日分。1年度に十分）
var SUMMARY_FIRST_ROW = 8;       // 出欠サマリの部員データが始まる行

// --- 会計簿（B-1） -------------------------------------------------------
var ACCT_FIRST_ROW = 3;                                  // 会計簿：記帳データが始まる行
var ACCT_MAX       = 300;                                // 会計簿：対応する記帳件数
var ACCT_LAST_ROW  = ACCT_FIRST_ROW + ACCT_MAX - 1;      // = 302行目

// --- 部費徴収（B-2） -----------------------------------------------------
var DUES_FIRST_TERM_COL = 4;                                          // 徴収回の最初の列（D列）
var DUES_TERMS_MAX      = 8;                                          // 対応する徴収回数（第1回〜第8回）
var DUES_LAST_COL       = DUES_FIRST_TERM_COL + DUES_TERMS_MAX - 1;   // = 11列目（K列）
var DUES_TERM_NAME_ROW  = 13;   // 徴収回名の行
var DUES_AMOUNT_ROW     = 14;   // 徴収額（1人あたり）の行
var DUES_DEADLINE_ROW   = 15;   // 締切の行
var DUES_PAID_ROW       = 16;   // 入金済み合計（自動計算）の行
var DUES_UNPAID_ROW     = 17;   // 未納人数（自動計算）の行
var DUES_HEADER_ROW     = 18;   // 学年・組・氏名の見出し行
var DUES_FIRST_DATA_ROW = 19;   // 部員データが始まる行
var DUES_LAST_ROW       = DUES_FIRST_DATA_ROW + ROSTER_MAX - 1;  // = 78行目

// --- 年度末報告（B-3） ---------------------------------------------------
var REPORT_LAST_ROW = 24;
var REPORT_LAST_COL = 7;

var COLOR = {
  CREAM:        '#fff8e1',   // 入力してよいセル（全シート共通ルール）
  RED_LIGHT:    '#f4cccc',   // 欠席の行／未納セル／アラート
  YEL_LIGHT:    '#fff2cc',   // 遅刻・早退・遅早の行／案内バー
  HEADER:       '#efefef',   // 見出し行の背景
  MONTH_ODD:    '#e8f0fe',   // 出欠簿の日付見出し（奇数月）
  MONTH_EVEN:   '#e6f4ea',   // 出欠簿の日付見出し（偶数月）
  INCOME_LIGHT: '#e6f4ea',   // 会計簿：収入の行
  EXPENSE_LIGHT:'#fce8e6'    // 会計簿：支出の行／残高マイナス
};

/* =====================================================================
 * メニュー（スプレッドシートを開くたびに自動で追加されます）
 * =================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎾部活キット')
    .addItem('✅ 今日の出欠を記録する', 'recordToday')
    .addItem('🧹 今日の入力だけクリアする（記録しない）', 'clearTodayInput')
    .addSeparator()
    .addItem('📖 使い方を開く（はじめに）', 'showGuide')
    .addItem('🛠 初期セットアップ（初回のみ）', 'setup')
    .addToUi();
}

/** メニュー「📖 使い方」→「はじめに」シートへ移動するだけ */
function showGuide() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET.GUIDE);
  if (sh) ss.setActiveSheet(sh);
}

/* =====================================================================
 * ① 初期セットアップ（初回に1回だけ実行）
 * ---------------------------------------------------------------------
 * ・9つのシートを自動生成し、見出し・プルダウン・色・数式・見本データを
 *   すべて用意します。
 * ・【安全設計】すでに存在するシートには一切触りません（再実行しても
 *   データは壊れません）。無いシートだけを作り足します。
 * =================================================================== */

function setup() {
  var ss = SpreadsheetApp.getActive();
  var made = [];     // 今回つくったシート
  var skipped = [];  // すでにあったのでスキップしたシート

  // --- 1) シートを（無ければ）タブ順に作る -------------------------
  var guide   = ensureSheet_(ss, SHEET.GUIDE,   0);
  var roster  = ensureSheet_(ss, SHEET.ROSTER,  1);
  var today   = ensureSheet_(ss, SHEET.TODAY,   2);
  var ledger  = ensureSheet_(ss, SHEET.LEDGER,  3);
  var summary = ensureSheet_(ss, SHEET.SUMMARY, 4);
  var account = ensureSheet_(ss, SHEET.ACCOUNT, 5);
  var dues    = ensureSheet_(ss, SHEET.DUES,    6);
  var report  = ensureSheet_(ss, SHEET.REPORT,  7);
  var config  = ensureSheet_(ss, SHEET.CONFIG,  8);

  // --- 2) 新しく作ったシートだけ中身を入れる（依存順：設定が先）----
  if (config.created)  { buildConfig_(config.sheet);   made.push(SHEET.CONFIG); }  else skipped.push(SHEET.CONFIG);
  if (roster.created)  { buildRoster_(ss, roster.sheet);  made.push(SHEET.ROSTER); }  else skipped.push(SHEET.ROSTER);
  if (today.created)   { buildToday_(ss, today.sheet);    made.push(SHEET.TODAY); }   else skipped.push(SHEET.TODAY);
  if (ledger.created)  { buildLedger_(ledger.sheet);      made.push(SHEET.LEDGER); }  else skipped.push(SHEET.LEDGER);
  if (summary.created) { buildSummary_(summary.sheet);    made.push(SHEET.SUMMARY); } else skipped.push(SHEET.SUMMARY);
  if (account.created) { buildAcct_(ss, account.sheet);   made.push(SHEET.ACCOUNT); } else skipped.push(SHEET.ACCOUNT);
  if (dues.created)     { buildDues_(dues.sheet);          made.push(SHEET.DUES); }    else skipped.push(SHEET.DUES);
  if (report.created)  { buildReport_(report.sheet);      made.push(SHEET.REPORT); } else skipped.push(SHEET.REPORT);
  if (guide.created)   { buildGuide_(guide.sheet);        made.push(SHEET.GUIDE); }   else skipped.push(SHEET.GUIDE);

  // --- 3) 最初から入っている空の「シート1」を片づける ---------------
  removeDefaultSheet_(ss);

  // --- 4) ファイル名が「無題」のままなら商品名を付ける ---------------
  if (/無題|Untitled/i.test(ss.getName())) {
    ss.rename('部活動まるごと管理キット【出欠＋会計】');
  }

  // --- 5) 結果を報告して「はじめに」を表示 ---------------------------
  if (guide.sheet) ss.setActiveSheet(guide.sheet);
  var msg = 'セットアップが完了しました！🎾\n\n';
  if (made.length)    msg += '✅ 作成：' + made.join('・') + '\n';
  if (skipped.length) msg += '⏭ すでにあるため触りませんでした：' + skipped.join('・') + '\n';
  msg += '\nまず「はじめに」シートをお読みください。';
  alert_(msg);
}

/* =====================================================================
 * ② 毎日の記録：「✅ 今日の出欠を記録する」
 * ---------------------------------------------------------------------
 * 「今日の出欠」の内容を「出欠簿」の当日列へ記号で転記します。
 * ・区分が空欄の部員 → ○（出席）として記録（この部の標準運用）
 * ・欠席=×／遅刻=遅／早退=早／遅早=遅早（「設定」シートで変更・追加可。
 *   追加した区分に記号が無いときは、区分名の1文字目を記号にします）
 * ・理由はセルの「メモ（注釈）」として残ります
 * ・当日の列が無ければ日付順の位置に自動で追加。すでにあれば上書き確認
 * ・転記後、「今日の出欠」の入力欄をクリアして翌日に備えます（確認あり）
 * =================================================================== */

function recordToday() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var today  = ss.getSheetByName(SHEET.TODAY);
  var ledger = ss.getSheetByName(SHEET.LEDGER);
  var config = ss.getSheetByName(SHEET.CONFIG);

  if (!today || !ledger || !config) {
    alert_('必要なシートが見つかりません。\nメニュー「🛠 初期セットアップ」を先に実行してください。');
    return;
  }

  // --- 1) 記録する日付を読む（「今日の出欠」B1セル） -----------------
  var rawDate = today.getRange('B1').getValue();
  if (!(rawDate instanceof Date)) {
    alert_('「今日の出欠」シートの B1 セルに日付が入っていません。\n「=TODAY()」または「2026/7/8」のように入力してください。');
    return;
  }
  var tz = ss.getSpreadsheetTimeZone();
  var d = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate()); // 時刻を切り捨てて日付だけに
  var dateKey   = Utilities.formatDate(d, tz, 'yyyyMMdd');  // 比較用（例 20260708）
  var dateLabel = Utilities.formatDate(d, tz, 'M/d');       // 表示用（例 7/8）
  var youbi     = '日月火水木金土'.charAt(d.getDay());

  // --- 2) 今日の出欠の入力内容を読む ---------------------------------
  // 列： A学年 / B組 / C氏名 / D担任 / E区分 / F理由
  var rows = today.getRange(ROSTER_FIRST_ROW, 1, ROSTER_MAX, 6).getDisplayValues();
  var members = rows.filter(function (r) { return r[2] !== '' && r[2].indexOf('※') !== 0; });
  if (members.length === 0) {
    alert_('「今日の出欠」に部員が表示されていません。\n「名簿」シートに部員が入っていて、「在籍」列が「在籍」になっているか確認してください。');
    return;
  }

  // --- 3) 区分→記号の対応表を「設定」シートから読む -------------------
  var symbolMap = buildSymbolMap_(config);

  // --- 4) 出欠簿の日付列を探す（同日あり→上書き確認／無ければ挿入）----
  var lastCol = ledger.getLastColumn();
  var headerDates = [];  // {col: 列番号, key: 'yyyyMMdd'}
  if (lastCol >= LEDGER_FIRST_DATE_COL) {
    var headerVals = ledger.getRange(1, LEDGER_FIRST_DATE_COL, 1, lastCol - LEDGER_FIRST_DATE_COL + 1).getValues()[0];
    for (var i = 0; i < headerVals.length; i++) {
      if (headerVals[i] instanceof Date) {
        headerDates.push({ col: LEDGER_FIRST_DATE_COL + i, key: Utilities.formatDate(headerVals[i], tz, 'yyyyMMdd') });
      }
    }
  }

  var targetCol = -1;
  var existing = headerDates.filter(function (h) { return h.key === dateKey; })[0];

  if (existing) {
    // 同じ日付の列がすでにある → 上書きしてよいか確認
    var ans = ui.alert('上書きの確認',
      dateLabel + '（' + youbi + '）の記録はすでにあります。\n今の「今日の出欠」の内容で上書きしますか？',
      ui.ButtonSet.YES_NO);
    if (ans !== ui.Button.YES) return;
    targetCol = existing.col;
    // 古い記録とメモを消してから書き直す
    ledger.getRange(ROSTER_FIRST_ROW, targetCol, ROSTER_MAX, 1).clearContent().clearNote();
  } else {
    // 新しい日付 → 日付順になる位置に列を入れる
    var insertBefore = null;
    for (var j = 0; j < headerDates.length; j++) {
      if (headerDates[j].key > dateKey) { insertBefore = headerDates[j].col; break; }
    }
    if (insertBefore !== null) {
      // 過去の日付をあとから記録するケース：途中に挿入
      ledger.insertColumnBefore(insertBefore);
      targetCol = insertBefore;
    } else {
      // ふつうのケース：右端に追加
      targetCol = LEDGER_FIRST_DATE_COL + headerDates.length;
      if (targetCol > ledger.getMaxColumns()) {
        ledger.insertColumnsAfter(ledger.getMaxColumns(), targetCol - ledger.getMaxColumns());
      }
    }
    // 日付見出しの見た目を整える（日付・曜日・月ごとの色・幅・中央ぞろえ）
    var monthColor = ((d.getMonth() + 1) % 2 === 1) ? COLOR.MONTH_ODD : COLOR.MONTH_EVEN;
    ledger.getRange(1, targetCol).setValue(d).setNumberFormat('M/d');
    ledger.getRange(2, targetCol).setValue(youbi);
    ledger.getRange(1, targetCol, 2, 1).setBackground(monthColor).setFontWeight('bold');
    ledger.getRange(ROSTER_FIRST_ROW, targetCol, ROSTER_MAX, 1).setBackground('#ffffff');
    ledger.getRange(1, targetCol, ROSTER_LAST_ROW, 1).setHorizontalAlignment('center');
    ledger.setColumnWidth(targetCol, 48);
  }

  // --- 5) 出欠簿の部員行を「学年|組|氏名」で対応づける -----------------
  var ledgerRows = ledger.getRange(ROSTER_FIRST_ROW, 1, ROSTER_MAX, 3).getDisplayValues();
  var rowByKey = {};
  for (var k = 0; k < ledgerRows.length; k++) {
    var key = memberKey_(ledgerRows[k][0], ledgerRows[k][1], ledgerRows[k][2]);
    if (key !== '||' && !(key in rowByKey)) rowByKey[key] = k;  // k = 出欠簿の上から何人目か
  }

  // --- 6) 記号と理由メモを組み立てて、一括で書き込む -------------------
  var values = [], notes = [];
  for (var v = 0; v < ROSTER_MAX; v++) { values.push(['']); notes.push(['']); }
  var counts = {};        // 記号ごとの人数（報告用）
  var unmatched = [];     // 出欠簿に見つからなかった部員（名簿を直後に変えた場合など）

  members.forEach(function (m) {
    var kubun  = String(m[4]).trim();   // E列：区分（空欄＝出席）
    var reason = String(m[5]).trim();   // F列：理由
    var symbol = (kubun === '') ? '○' : (symbolMap[kubun] || kubun.charAt(0));
    var idx = rowByKey[memberKey_(m[0], m[1], m[2])];
    if (idx === undefined) { unmatched.push(m[2]); return; }
    values[idx][0] = symbol;
    if (reason !== '') notes[idx][0] = '理由：' + reason;
    counts[symbol] = (counts[symbol] || 0) + 1;
  });

  var target = ledger.getRange(ROSTER_FIRST_ROW, targetCol, ROSTER_MAX, 1);
  target.setValues(values);
  target.setNotes(notes);

  // --- 7) 結果を報告 ---------------------------------------------------
  var order = ['○', '×', '遅', '早', '遅早'];
  var labels = { '○': '出席', '×': '欠席', '遅': '遅刻', '早': '早退', '遅早': '遅早' };
  var parts = [];
  order.forEach(function (s) { if (counts[s]) { parts.push(labels[s] + '(' + s + ') ' + counts[s] + '人'); delete counts[s]; } });
  Object.keys(counts).forEach(function (s) { parts.push(s + ' ' + counts[s] + '人'); });

  var report = '✅ ' + dateLabel + '（' + youbi + '）の出欠を記録しました。\n\n' + parts.join('　');
  if (unmatched.length) {
    report += '\n\n⚠️ 出欠簿に見つからず記録できなかった部員：' + unmatched.join('、') +
              '\n（名簿を変更した直後の可能性があります。表示が落ち着いてから再実行してください）';
  }
  alert_(report);

  // --- 8) 「今日の出欠」の入力欄をクリアして翌日に備える（確認あり）---
  var ans2 = ui.alert('入力欄のクリア',
    '「今日の出欠」の入力（区分・理由）をクリアして、明日の分に備えますか？\n（記録済みの内容は出欠簿に残っています）',
    ui.ButtonSet.YES_NO);
  if (ans2 === ui.Button.YES) {
    today.getRange(ROSTER_FIRST_ROW, 5, ROSTER_MAX, 2).clearContent();  // E・F列
    today.getRange('B1').setFormula('=TODAY()');                        // 日付を「今日」に戻す
    ss.toast('入力欄をクリアしました。おつかれさまでした！', '🎾部活キット', 5);
  }
}

/** メニュー「🧹 今日の入力だけクリアする」：記録せずに入力欄を消す */
function clearTodayInput() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var today = ss.getSheetByName(SHEET.TODAY);
  if (!today) return;
  var ans = ui.alert('確認',
    '「今日の出欠」の入力（区分・理由）を、出欠簿に記録せずにクリアします。よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  today.getRange(ROSTER_FIRST_ROW, 5, ROSTER_MAX, 2).clearContent();
  today.getRange('B1').setFormula('=TODAY()');
  ss.toast('入力欄をクリアしました。', '🎾部活キット', 5);
}

/* =====================================================================
 * 以下、シートの中身をつくる裏方の関数（読まなくて大丈夫です）
 * =================================================================== */

/** シートが無ければ作る。あれば触らずそのまま返す（再実行しても安全） */
function ensureSheet_(ss, name, index) {
  var sh = ss.getSheetByName(name);
  if (sh) return { sheet: sh, created: false };
  return { sheet: ss.insertSheet(name, index), created: true };
}

/** シートの行数・列数をきっちり整える（余計な行列を無くして軽くする） */
function resize_(sh, rows, cols) {
  var mr = sh.getMaxRows(), mc = sh.getMaxColumns();
  if (mr > rows) sh.deleteRows(rows + 1, mr - rows);
  if (mr < rows) sh.insertRowsAfter(mr, rows - mr);
  if (mc > cols) sh.deleteColumns(cols + 1, mc - cols);
  if (mc < cols) sh.insertColumnsAfter(mc, cols - mc);
}

/** 最初から入っている空の「シート1」を削除（他のシートがある場合のみ） */
function removeDefaultSheet_(ss) {
  ['シート1', 'Sheet1'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && ss.getSheets().length > 1 && sh.getLastRow() === 0 && sh.getLastColumn() === 0) {
      ss.deleteSheet(sh);
    }
  });
}

/** 「設定」シートの区分表から、区分→記号の対応表をつくる */
function buildSymbolMap_(config) {
  var table = config.getRange('A4:B13').getDisplayValues();
  var map = {};
  table.forEach(function (r) {
    var name = String(r[0]).trim();
    var sym  = String(r[1]).trim();
    if (name !== '') map[name] = (sym !== '') ? sym : name.charAt(0);
  });
  return map;
}

/** 部員の照合キー（学年|組|氏名）。空白ゆれを吸収 */
function memberKey_(grade, kumi, name) {
  return String(grade).trim() + '|' + String(kumi).trim() + '|' + String(name).trim();
}

/** 列番号→列文字（1=A, 2=B, ... , 26=Z）。部費徴収の徴収回列の組み立てに使用 */
function colLetter_(n) {
  return String.fromCharCode(64 + n);
}

/** ダイアログ表示（エディタから実行した場合などでも落ちないように） */
function alert_(msg) {
  try { SpreadsheetApp.getUi().alert('🎾部活キット', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }
}

/* --------------------------------------------------------------------
 * 「設定」シート：区分と記号・学年の選択肢・費目一覧・部活名など（編集OKな場所）
 * ------------------------------------------------------------------ */
function buildConfig_(sh) {
  resize_(sh, 30, 9);
  sh.setTabColor('#999999');

  sh.getRange('A1').setValue('⚙️ 設定 ― このシートは色のついたセルだけ編集OK').setFontWeight('bold').setFontSize(12);
  sh.getRange('A2').setValue('（ほかのシートの白いセルは自動計算です。触らないでください）').setFontColor('#999999');

  // 区分と記号の表（プルダウンの中身はここを見ています）
  sh.getRange('A3:C3').setValues([['区分（プルダウンに出る言葉）', '記号（出欠簿に書かれる印）', 'メモ']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.getRange('A4:C7').setValues([
    ['欠席', '×',   'お休み'],
    ['遅刻', '遅',  '遅れて参加'],
    ['早退', '早',  '途中で帰る'],
    ['遅早', '遅早', '遅刻して早退']
  ]);
  sh.getRange('C8').setValue('← 上の空き行（13行目まで）に学校独自の区分を追加できます（例：公欠・見学・別メニュー・テスト期間）。記号欄が空のときは区分名の1文字目が記号になります。')
    .setFontColor('#666666').setWrap(true);
  sh.getRange('A4:B13').setBackground(COLOR.CREAM);

  // 学年の選択肢（名簿のプルダウンの中身）
  sh.getRange('E3').setValue('学年の選択肢').setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.getRange('E4:E6').setValues([[1], [2], [3]]);
  sh.getRange('F4').setValue('← 中等教育学校などで4〜6年を使う場合は E7〜E9 に追記').setFontColor('#666666');
  sh.getRange('E4:E9').setBackground(COLOR.CREAM);

  // 費目一覧（会計簿のプルダウンの中身。編集可）
  sh.getRange('H3').setValue('費目一覧（収入向け）').setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.getRange('H4:H6').setValues([['部費'], ['補助金'], ['その他（収入）']]);
  sh.getRange('I3').setValue('費目一覧（支出向け）').setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.getRange('I4:I8').setValues([['消耗品'], ['大会参加費'], ['交通費'], ['登録料'], ['その他（支出）']]);
  sh.getRange('H4:I8').setBackground(COLOR.CREAM);
  sh.getRange('H9').setValue('← 空き行に学校独自の費目を追加できます（会計簿のプルダウンに自動で反映されます）')
    .setFontColor('#666666').setWrap(true);

  // 部活名・年度開始月
  sh.getRange('A16').setValue('部活名').setFontWeight('bold');
  sh.getRange('B16').setValue('○○部').setBackground(COLOR.CREAM);
  sh.getRange('C16').setValue('（お好きな名前に。会計簿の年度末報告などで使います）').setFontColor('#666666');
  sh.getRange('A17').setValue('年度開始月').setFontWeight('bold');
  sh.getRange('B17').setValue(4).setBackground(COLOR.CREAM);
  sh.getRange('C17').setValue('（通常は 4 のまま）').setFontColor('#666666');

  // 顧問名（年度末報告の差し込みに使用。匿名運用なら空欄でも可）
  sh.getRange('A18').setValue('顧問名').setFontWeight('bold');
  sh.getRange('B18').setValue('').setBackground(COLOR.CREAM);
  sh.getRange('C18').setValue('（年度末報告に印字されます。空欄でも構いません）').setFontColor('#666666');

  // 困ったとき（初期状態への戻し方）
  sh.getRange('A20').setValue('🆘 初期状態に戻したいとき').setFontWeight('bold');
  sh.getRange('A21').setValue('いちばん安全な方法：配布ページのリンクからもう一度コピーを作り、「名簿」だけ貼り直す。').setWrap(true);
  sh.getRange('A22').setValue('一部のシートが壊れた場合：壊れたシートを削除 → メニュー「🎾部活キット」→「🛠 初期セットアップ」で、そのシートだけ作り直せます。※「出欠簿」を削除すると出欠の記録も消えます。削除するときは「出欠サマリ」も一緒に削除してから再実行してください（サマリの数式が出欠簿を参照しているため）。会計簿・部費徴収・年度末報告は互いに独立しているので、1枚だけ削除→再実行しても安全です。').setWrap(true);

  sh.setColumnWidth(1, 230).setColumnWidth(2, 190).setColumnWidth(3, 330).setColumnWidth(5, 110).setColumnWidth(6, 330)
    .setColumnWidth(8, 160).setColumnWidth(9, 160);

  // うっかり編集の防止（警告付き保護：編集はできるが確認が出る）
  sh.protect().setWarningOnly(true).setDescription('設定シート（編集前に確認が出ます）');
}

/* --------------------------------------------------------------------
 * 「名簿」シート：唯一の初期設定場所。見本20名（架空の名前）入り
 * ------------------------------------------------------------------ */
function buildRoster_(ss, sh) {
  var config = ss.getSheetByName(SHEET.CONFIG);
  resize_(sh, ROSTER_LAST_ROW, 7);
  sh.setTabColor('#fbbc04');

  // 1行目：案内バー
  sh.getRange('A1:G1').merge().setValue('🟡 3行目からが名簿です。この見本（すべて架空の名前）を消して、あなたの部の部員に貼り替えてください（最大60名）')
    .setBackground(COLOR.YEL_LIGHT).setFontWeight('bold').setWrap(true);

  // 2行目：見出し
  sh.getRange('A2:G2').setValues([[
    '学年', '組', '番号', '氏名', '担任（任意）', '備考 ⚠️個人情報。記載は最小限に', '在籍'
  ]]).setFontWeight('bold').setBackground(COLOR.HEADER);

  // 見本データ（20名・完全に架空。1名だけ「引退」の例を入れてあります）
  var sample = [
    [3, 1, 12, '青井 陽太', '', '', '在籍'],
    [3, 1, 18, '石川 湊',   '', '', '在籍'],
    [3, 2,  5, '内田 蒼空', '', '', '在籍'],
    [3, 3, 21, '遠藤 楓',   '', '', '在籍'],
    [3, 4,  9, '大野 律',   '', '', '在籍'],
    [3, 2, 14, '加賀 美咲', '', '', '引退'],
    [2, 1,  3, '北村 悠真', '', '', '在籍'],
    [2, 1, 25, '久保 陽菜', '', '', '在籍'],
    [2, 2,  8, '小泉 蓮',   '', '', '在籍'],
    [2, 3, 17, '佐々木 澪', '', '', '在籍'],
    [2, 3, 30, '島田 大翔', '', '', '在籍'],
    [2, 4, 11, '瀬川 結衣', '', '', '在籍'],
    [2, 4, 22, '高木 颯',   '', '', '在籍'],
    [1, 1,  7, '津田 芽依', '', '', '在籍'],
    [1, 1, 28, '中西 陸斗', '', '', '在籍'],
    [1, 2,  2, '西野 心春', '', '', '在籍'],
    [1, 2, 19, '野口 樹',   '', '', '在籍'],
    [1, 3, 15, '浜田 咲良', '', '', '在籍'],
    [1, 4,  4, '藤井 隼人', '', '', '在籍'],
    [1, 4, 26, '松田 莉子', '', '', '在籍']
  ];
  sh.getRange(ROSTER_FIRST_ROW, 1, sample.length, 7).setValues(sample);

  // プルダウン（学年＝設定シート参照、在籍＝固定3択）
  if (config) {
    // ※「警告」方式（setAllowInvalid(true)）にしてあるのは、名簿の一括貼り付けを
    //   ブロックしないため。想定外の値には赤い三角の警告が付きます。
    var gradeRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(config.getRange('E4:E9'), true).setAllowInvalid(true)
      .setHelpText('学年を選んでください（選択肢は「設定」シートで変更できます）').build();
    sh.getRange(ROSTER_FIRST_ROW, 1, ROSTER_MAX, 1).setDataValidation(gradeRule);
  }
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['在籍', '引退', '退部'], true).setAllowInvalid(true)
    .setHelpText('引退・退部になっても行は消さないでください（過去の記録が残ります）').build();
  sh.getRange(ROSTER_FIRST_ROW, 7, ROSTER_MAX, 1).setDataValidation(statusRule);

  // 入力エリアはクリーム色（＝触ってよい場所）
  sh.getRange(ROSTER_FIRST_ROW, 1, ROSTER_MAX, 7).setBackground(COLOR.CREAM);
  sh.getRange(ROSTER_FIRST_ROW, 3, ROSTER_MAX, 1).setNumberFormat('0');

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 45).setColumnWidth(2, 45).setColumnWidth(3, 45)
    .setColumnWidth(4, 120).setColumnWidth(5, 90).setColumnWidth(6, 260).setColumnWidth(7, 60);

  // うっかり全消しの防止（警告付き保護）
  sh.protect().setWarningOnly(true).setDescription('名簿シート（編集前に確認が出ます）');
}

/* --------------------------------------------------------------------
 * 「今日の出欠」シート：毎日触るのはここだけ（スマホ入力対応・6列）
 * ------------------------------------------------------------------ */
function buildToday_(ss, sh) {
  var config = ss.getSheetByName(SHEET.CONFIG);
  resize_(sh, ROSTER_LAST_ROW, 6);
  sh.setTabColor('#34a853');

  // 1行目：日付と案内
  sh.getRange('A1').setValue('日付').setFontWeight('bold');
  sh.getRange('B1').setFormula('=TODAY()').setNumberFormat('yyyy/M/d')
    .setBackground(COLOR.CREAM).setFontWeight('bold');
  sh.getRange('C1').setFormula('=IF($B$1="","",CHOOSE(WEEKDAY($B$1),"日","月","火","水","木","金","土")&"曜日")');
  sh.getRange('D1:F1').merge().setValue('💡 空欄＝出席。欠席・遅刻などの人だけ「区分」を選べばOK')
    .setFontColor('#666666');

  // 2行目：見出し
  sh.getRange('A2:F2').setValues([['学年', '組', '氏名', '担任', '区分', '理由（任意）']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);

  // 名簿から「在籍」の部員だけを自動表示（この数式は触らない）
  sh.getRange('A3').setFormula(
    '=IFERROR(FILTER({名簿!A3:A62,名簿!B3:B62,名簿!D3:D62,名簿!E3:E62},名簿!G3:G62="在籍",名簿!D3:D62<>""),"※名簿に「在籍」の部員がいません")'
  );

  // 区分プルダウン（中身は「設定」シートの区分表）
  if (config) {
    var kubunRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(config.getRange('A4:A13'), true).setAllowInvalid(false)
      .setHelpText('空欄＝出席。欠席・遅刻・早退・遅早（追加は「設定」シートで）').build();
    sh.getRange(ROSTER_FIRST_ROW, 5, ROSTER_MAX, 1).setDataValidation(kubunRule);
  }

  // 入力してよいのは区分・理由（クリーム色）
  sh.getRange(ROSTER_FIRST_ROW, 5, ROSTER_MAX, 2).setBackground(COLOR.CREAM);

  // 条件付き書式：欠席＝赤系、遅刻・早退・遅早＝黄系（行にうすく色）
  var dataRange = sh.getRange(ROSTER_FIRST_ROW, 1, ROSTER_MAX, 6);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E3="欠席"')
      .setBackground(COLOR.RED_LIGHT).setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($E3="遅刻",$E3="早退",$E3="遅早")')
      .setBackground(COLOR.YEL_LIGHT).setRanges([dataRange]).build()
  ];
  sh.setConditionalFormatRules(rules);

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 45).setColumnWidth(2, 40).setColumnWidth(3, 120)
    .setColumnWidth(4, 60).setColumnWidth(5, 80).setColumnWidth(6, 240);
}

/* --------------------------------------------------------------------
 * 「出欠簿」シート：行＝部員（名簿連動）、列＝活動日（GASが自動追加）
 * ------------------------------------------------------------------ */
function buildLedger_(sh) {
  resize_(sh, ROSTER_LAST_ROW, LEDGER_MAX_COLS);
  sh.setTabColor('#4285f4');

  sh.getRange('A1:C1').merge()
    .setValue('（自動）→ 右へ活動日が増えます').setFontColor('#999999').setFontSize(9);
  sh.getRange('A2:C2').setValues([['学年', '組', '氏名']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);

  // 部員の行は名簿をそのまま鏡写し（引退・退部でも行は残る＝記録が消えない）
  sh.getRange('A3').setFormula('=ARRAYFORMULA(IF(名簿!$D$3:$D$62="","",名簿!$A$3:$A$62))');
  sh.getRange('B3').setFormula('=ARRAYFORMULA(IF(名簿!$D$3:$D$62="","",名簿!$B$3:$B$62))');
  sh.getRange('C3').setFormula('=ARRAYFORMULA(IF(名簿!$D$3:$D$62="","",名簿!$D$3:$D$62))');

  // 条件付き書式：×＝赤系、遅・早・遅早＝黄系（日付列ぜんぶに適用）
  var dataRange = sh.getRange(ROSTER_FIRST_ROW, LEDGER_FIRST_DATE_COL, ROSTER_MAX, LEDGER_MAX_COLS - 3);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('×').setBackground(COLOR.RED_LIGHT).setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR(D3="遅",D3="早",D3="遅早")')
      .setBackground(COLOR.YEL_LIGHT).setRanges([dataRange]).build()
  ];
  sh.setConditionalFormatRules(rules);

  sh.setFrozenRows(2);
  sh.setFrozenColumns(3);
  sh.setColumnWidth(1, 40).setColumnWidth(2, 36).setColumnWidth(3, 110);
}

/* --------------------------------------------------------------------
 * 「出欠サマリ」シート：全部数式の自動集計（年度累計＋当月＋アラート）
 * ------------------------------------------------------------------ */
function buildSummary_(sh) {
  var LAST = SUMMARY_FIRST_ROW + ROSTER_MAX - 1;  // = 67行目
  resize_(sh, LAST, 19);
  sh.setTabColor('#a142f4');

  sh.getRange('A1').setValue('📊 出欠サマリ（すべて自動計算。このシートは編集不要）')
    .setFontWeight('bold').setFontSize(12);

  // --- アラート2本（顧問目線の機能） ---------------------------------
  sh.getRange('A3').setValue('⚠️ 連続3回欠席：').setFontWeight('bold');
  sh.getRange('B3:K3').merge();
  sh.getRange('B3').setFormula(
    '=IF(COUNTIF($K$8:$K$67,"⚠️")=0,"なし",TEXTJOIN("、",TRUE,FILTER($C$8:$C$67,$K$8:$K$67="⚠️")))'
  );
  sh.getRange('A4').setValue('📝 記録状況：').setFontWeight('bold');
  sh.getRange('B4:K4').merge();
  sh.getRange('B4').setFormula(
    '=IF(COUNTA(出欠簿!$D$1:$1)=0,"まだ記録がありません（メニュー「🎾部活キット」→「✅ 今日の出欠を記録する」から始まります）",' +
    'IF(TODAY()-MAX(出欠簿!$D$1:$1)>=3,"⚠️ 最後の記録（"&TEXT(MAX(出欠簿!$D$1:$1),"M/d")&"）から3日以上あいています。つけ忘れはありませんか？",' +
    '"OK：最後の記録は "&TEXT(MAX(出欠簿!$D$1:$1),"M/d")&" です"))'
  );

  // --- 見出し（左＝年度累計、右＝当月） ------------------------------
  sh.getRange('A6').setValue('■ 年度累計').setFontWeight('bold');
  sh.getRange('M6').setValue('■ 当月（今月ぶんだけ）').setFontWeight('bold');
  sh.getRange('A7:S7').setValues([[
    '学年', '組', '氏名', '記録日数', '出席(○)', '欠席(×)', '遅刻', '早退', '遅早', '出席率', '連続欠席', '',
    '記録日数', '出席(○)', '欠席(×)', '遅刻', '早退', '遅早', '出席率'
  ]]).setFontWeight('bold').setBackground(COLOR.HEADER);

  // --- 部員名は名簿を鏡写し ------------------------------------------
  sh.getRange('A8').setFormula('=ARRAYFORMULA(IF(名簿!$D$3:$D$62="","",名簿!$A$3:$A$62))');
  sh.getRange('B8').setFormula('=ARRAYFORMULA(IF(名簿!$D$3:$D$62="","",名簿!$B$3:$B$62))');
  sh.getRange('C8').setFormula('=ARRAYFORMULA(IF(名簿!$D$3:$D$62="","",名簿!$D$3:$D$62))');

  // --- 部員ごとの集計数式を60行ぶん流し込む ---------------------------
  // ※出欠簿の行と 5行ずれで対応（サマリ8行目 ＝ 出欠簿3行目の部員）
  // ※「出欠簿!$D3:3」という書き方は「3行目のD列から右端まで」という意味。
  //   活動日が何日増えても数式を直さなくてよい書き方です。
  // ※出席率＝（記録日数−欠席）÷記録日数。遅刻・早退・遅早・追加区分は
  //   出席扱いで計算します（欠席×だけを欠席として数える）。
  var left = [], right = [];
  for (var i = 0; i < ROSTER_MAX; i++) {
    var r = SUMMARY_FIRST_ROW + i;   // サマリ側の行
    var L = ROSTER_FIRST_ROW + i;    // 出欠簿側の行
    var row = '出欠簿!$D' + L + ':' + L;      // その部員の記録ぜんぶ（右端まで）
    var hdr = '出欠簿!$D$1:$1';               // 日付見出しぜんぶ（右端まで）
    var nowKey = 'TEXT(TODAY(),"yyyymm")';

    left.push([
      '=IF($C' + r + '="","",COUNTA(' + row + '))',
      '=IF($C' + r + '="","",COUNTIF(' + row + ',"○"))',
      '=IF($C' + r + '="","",COUNTIF(' + row + ',"×"))',
      '=IF($C' + r + '="","",COUNTIF(' + row + ',"遅"))',
      '=IF($C' + r + '="","",COUNTIF(' + row + ',"早"))',
      '=IF($C' + r + '="","",COUNTIF(' + row + ',"遅早"))',
      '=IF(OR($C' + r + '="",$D' + r + '=0),"",($D' + r + '-$F' + r + ')/$D' + r + ')',
      // 連続欠席：いちばん右の3列（直近3回の活動日）がすべて×なら⚠️
      '=IF($C' + r + '="","",IF(COUNTA(' + hdr + ')<3,"",' +
        'IF(AND(INDEX(' + row + ',1,COUNTA(' + hdr + '))="×",' +
               'INDEX(' + row + ',1,COUNTA(' + hdr + ')-1)="×",' +
               'INDEX(' + row + ',1,COUNTA(' + hdr + ')-2)="×"),"⚠️","")))'
    ]);

    right.push([
      '=IF($C' + r + '="","",SUMPRODUCT((' + hdr + '<>"")*(TEXT(' + hdr + ',"yyyymm")=' + nowKey + ')*(' + row + '<>"")))',
      '=IF($C' + r + '="","",SUMPRODUCT((' + hdr + '<>"")*(TEXT(' + hdr + ',"yyyymm")=' + nowKey + ')*(' + row + '="○")))',
      '=IF($C' + r + '="","",SUMPRODUCT((' + hdr + '<>"")*(TEXT(' + hdr + ',"yyyymm")=' + nowKey + ')*(' + row + '="×")))',
      '=IF($C' + r + '="","",SUMPRODUCT((' + hdr + '<>"")*(TEXT(' + hdr + ',"yyyymm")=' + nowKey + ')*(' + row + '="遅")))',
      '=IF($C' + r + '="","",SUMPRODUCT((' + hdr + '<>"")*(TEXT(' + hdr + ',"yyyymm")=' + nowKey + ')*(' + row + '="早")))',
      '=IF($C' + r + '="","",SUMPRODUCT((' + hdr + '<>"")*(TEXT(' + hdr + ',"yyyymm")=' + nowKey + ')*(' + row + '="遅早")))',
      '=IF(OR($C' + r + '="",$M' + r + '=0),"",($M' + r + '-$O' + r + ')/$M' + r + ')'
    ]);
  }
  sh.getRange(SUMMARY_FIRST_ROW, 4, ROSTER_MAX, 8).setFormulas(left);   // D〜K列
  sh.getRange(SUMMARY_FIRST_ROW, 13, ROSTER_MAX, 7).setFormulas(right); // M〜S列

  // 出席率はパーセント表示
  sh.getRange(SUMMARY_FIRST_ROW, 10, ROSTER_MAX, 1).setNumberFormat('0.0%'); // J列
  sh.getRange(SUMMARY_FIRST_ROW, 19, ROSTER_MAX, 1).setNumberFormat('0.0%'); // S列

  // 連続欠席⚠️は赤く目立たせる
  var flagRange = sh.getRange(SUMMARY_FIRST_ROW, 11, ROSTER_MAX, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('⚠️').setBackground(COLOR.RED_LIGHT).setRanges([flagRange]).build()
  ]);

  sh.setFrozenRows(7);
  sh.setFrozenColumns(3);
  sh.setColumnWidth(1, 40).setColumnWidth(2, 36).setColumnWidth(3, 110).setColumnWidth(12, 20);
  for (var c = 4; c <= 11; c++) sh.setColumnWidth(c, 66);
  for (var c2 = 13; c2 <= 19; c2++) sh.setColumnWidth(c2, 66);
}

/* --------------------------------------------------------------------
 * 「会計簿」シート（B-1）：記帳するだけで残高が自動計算される
 * ------------------------------------------------------------------ */
function buildAcct_(ss, sh) {
  var config = ss.getSheetByName(SHEET.CONFIG);
  resize_(sh, ACCT_LAST_ROW, 7);
  sh.setTabColor('#f9ab00');

  sh.getRange('A1').setValue('期首残高（前年度からの繰越金）').setFontWeight('bold');
  sh.getRange('B1').setValue(0).setBackground(COLOR.CREAM).setNumberFormat('#,##0').setFontWeight('bold');
  sh.getRange('C1:E1').merge().setValue('（導入時・年度はじめに入力してください。以降の残高はここを起点に自動計算されます）')
    .setFontColor('#666666').setWrap(true);
  sh.getRange('F1:G1').merge().setValue('残高がマイナスの行は自動で赤字表示になります').setFontColor('#b45309').setFontSize(9).setWrap(true);

  sh.getRange('A2:G2').setValues([['日付', '種別', '費目', '内容', '金額', '残高', '領収書']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);

  // 種別プルダウン（固定2択。費目集計・年度末報告の条件に使う文字なので入力ミスを防ぐ厳格設定）
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['収入', '支出'], true).setAllowInvalid(false)
    .setHelpText('収入か支出かを選んでください').build();
  sh.getRange(ACCT_FIRST_ROW, 2, ACCT_MAX, 1).setDataValidation(typeRule);

  // 費目プルダウン（「設定」シートの費目一覧を参照。学校ごとに編集可）
  if (config) {
    var itemRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(config.getRange('H4:I8'), true).setAllowInvalid(false)
      .setHelpText('費目を選んでください（候補は「設定」シートで編集できます）').build();
    sh.getRange(ACCT_FIRST_ROW, 3, ACCT_MAX, 1).setDataValidation(itemRule);
  }

  // 領収書プルダウン（固定2択。監査対応の実務目線）
  var receiptRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['あり', 'なし'], true).setAllowInvalid(false)
    .setHelpText('領収書の有無を記録します').build();
  sh.getRange(ACCT_FIRST_ROW, 7, ACCT_MAX, 1).setDataValidation(receiptRule);

  // 残高列（F）：1行目はB1（期首残高）を起点に、2行目以降は1つ上の残高を引き継ぐ累計式
  // ※日付が空の行は空欄のまま（未記帳の行に0が並ばないようにするため）
  var balanceFormulas = [];
  for (var i = 0; i < ACCT_MAX; i++) {
    var r = ACCT_FIRST_ROW + i;
    var prev = (i === 0) ? '$B$1' : ('F' + (r - 1));
    balanceFormulas.push(['=IF($A' + r + '="","",' + prev + '+IF($B' + r + '="収入",$E' + r + ',IF($B' + r + '="支出",-$E' + r + ',0)))']);
  }
  sh.getRange(ACCT_FIRST_ROW, 6, ACCT_MAX, 1).setFormulas(balanceFormulas);

  // 入力してよいセルはクリーム色（残高＝F列だけは自動計算なので白のまま）
  sh.getRange(ACCT_FIRST_ROW, 1, ACCT_MAX, 5).setBackground(COLOR.CREAM);
  sh.getRange(ACCT_FIRST_ROW, 7, ACCT_MAX, 1).setBackground(COLOR.CREAM);

  sh.getRange(ACCT_FIRST_ROW, 1, ACCT_MAX, 1).setNumberFormat('yyyy/M/d');
  sh.getRange(ACCT_FIRST_ROW, 5, ACCT_MAX, 2).setNumberFormat('#,##0');

  // 条件付き書式：収入行＝うす緑／支出行＝うす赤（A〜E列）、残高マイナス＝赤字太字（F列）、
  // 「支出なのに領収書なし」＝黄色で注意喚起（G列・記帳もれチェック用）
  var rowRange     = sh.getRange(ACCT_FIRST_ROW, 1, ACCT_MAX, 5);
  var balanceRange = sh.getRange(ACCT_FIRST_ROW, 6, ACCT_MAX, 1);
  var receiptRange = sh.getRange(ACCT_FIRST_ROW, 7, ACCT_MAX, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B' + ACCT_FIRST_ROW + '="収入"')
      .setBackground(COLOR.INCOME_LIGHT).setRanges([rowRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B' + ACCT_FIRST_ROW + '="支出"')
      .setBackground(COLOR.EXPENSE_LIGHT).setRanges([rowRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(COLOR.EXPENSE_LIGHT).setFontColor('#cc0000').setBold(true).setRanges([balanceRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B' + ACCT_FIRST_ROW + '="支出",$G' + ACCT_FIRST_ROW + '="なし")')
      .setBackground(COLOR.YEL_LIGHT).setRanges([receiptRange]).build()
  ]);

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 90).setColumnWidth(2, 70).setColumnWidth(3, 130)
    .setColumnWidth(4, 220).setColumnWidth(5, 90).setColumnWidth(6, 100).setColumnWidth(7, 70);
}

/* --------------------------------------------------------------------
 * 「部費徴収」シート（B-2）：部員×徴収回のチェック管理。未納者は自動表示
 * ------------------------------------------------------------------ */
function buildDues_(sh) {
  resize_(sh, DUES_LAST_ROW, DUES_LAST_COL);
  sh.setTabColor('#cc0000');

  sh.getRange(1, 1, 1, DUES_LAST_COL).merge()
    .setValue('🟡 ' + DUES_TERM_NAME_ROW + '〜' + DUES_DEADLINE_ROW + '行目に「徴収回名・徴収額・締切」を入力してください（使う回だけでOK。徴収額が空欄の回は集計・未納判定の対象外です。見本として第1回を入力してあります）')
    .setBackground(COLOR.YEL_LIGHT).setFontWeight('bold').setWrap(true);

  sh.getRange('A2').setValue('⚠️ 未納者一覧（声かけリスト・自動更新）').setFontWeight('bold').setFontSize(12);

  // --- 未納者一覧（徴収回ごとに1行。回の名前・締切は下の設定行から自動反映） ---
  for (var i = 0; i < DUES_TERMS_MAX; i++) {
    var col = DUES_FIRST_TERM_COL + i;
    var L   = colLetter_(col);
    var r   = 3 + i;
    sh.getRange(r, 1).setFormula(
      '=IF(' + L + '$' + DUES_AMOUNT_ROW + '="","",' +
      L + '$' + DUES_TERM_NAME_ROW + '&"（〆切"&IF(' + L + '$' + DUES_DEADLINE_ROW + '="","未設定",TEXT(' + L + '$' + DUES_DEADLINE_ROW + ',"M/d"))&"）")'
    );
    sh.getRange(r, 2, 1, DUES_LAST_COL - 1).merge().setFormula(
      '=IF(' + L + '$' + DUES_AMOUNT_ROW + '="","（この回は未使用です）",' +
      'IF(COUNTIFS($C$' + DUES_FIRST_DATA_ROW + ':$C$' + DUES_LAST_ROW + ',"<>",' +
        L + DUES_FIRST_DATA_ROW + ':' + L + DUES_LAST_ROW + ',FALSE)=0,' +
      '"未納なし（全員納入済み）",' +
      '"未納："&TEXTJOIN("、",TRUE,FILTER($C$' + DUES_FIRST_DATA_ROW + ':$C$' + DUES_LAST_ROW + ',' +
        '$C$' + DUES_FIRST_DATA_ROW + ':$C$' + DUES_LAST_ROW + '<>"",' +
        L + DUES_FIRST_DATA_ROW + ':' + L + DUES_LAST_ROW + '=FALSE))))'
    );
  }

  // --- 徴収回の設定行（名前・金額・締切）と、自動集計2行 -------------
  sh.getRange('A' + DUES_TERM_NAME_ROW).setValue('徴収回名').setFontWeight('bold');
  sh.getRange('A' + DUES_AMOUNT_ROW).setValue('徴収額（1人あたり・円）').setFontWeight('bold');
  sh.getRange('A' + DUES_DEADLINE_ROW).setValue('締切').setFontWeight('bold');
  sh.getRange('A' + DUES_PAID_ROW).setValue('入金済み合計（円）').setFontWeight('bold').setFontColor('#666666');
  sh.getRange('A' + DUES_UNPAID_ROW).setValue('未納人数').setFontWeight('bold').setFontColor('#666666');

  sh.getRange(DUES_TERM_NAME_ROW, DUES_FIRST_TERM_COL, 1, DUES_TERMS_MAX).setBackground(COLOR.CREAM);
  sh.getRange(DUES_AMOUNT_ROW, DUES_FIRST_TERM_COL, 1, DUES_TERMS_MAX).setBackground(COLOR.CREAM).setNumberFormat('#,##0');
  sh.getRange(DUES_DEADLINE_ROW, DUES_FIRST_TERM_COL, 1, DUES_TERMS_MAX).setBackground(COLOR.CREAM).setNumberFormat('yyyy/M/d');

  // 見本：第1回（金額のみ入れる。締切は空欄のままにして「未設定」表示も確認できるようにする）
  sh.getRange(DUES_TERM_NAME_ROW, DUES_FIRST_TERM_COL).setValue('第1回');
  sh.getRange(DUES_AMOUNT_ROW, DUES_FIRST_TERM_COL).setValue(3000);

  for (var j = 0; j < DUES_TERMS_MAX; j++) {
    var c  = DUES_FIRST_TERM_COL + j;
    var Lj = colLetter_(c);
    sh.getRange(DUES_PAID_ROW, c).setFormula(
      '=IF(' + Lj + '$' + DUES_AMOUNT_ROW + '="","",COUNTIF(' + Lj + DUES_FIRST_DATA_ROW + ':' + Lj + DUES_LAST_ROW + ',TRUE)*' + Lj + '$' + DUES_AMOUNT_ROW + ')'
    );
    sh.getRange(DUES_UNPAID_ROW, c).setFormula(
      '=IF(' + Lj + '$' + DUES_AMOUNT_ROW + '="","",COUNTA($C$' + DUES_FIRST_DATA_ROW + ':$C$' + DUES_LAST_ROW + ')-COUNTIF(' + Lj + DUES_FIRST_DATA_ROW + ':' + Lj + DUES_LAST_ROW + ',TRUE))'
    );
  }
  sh.getRange(DUES_PAID_ROW, DUES_FIRST_TERM_COL, 1, DUES_TERMS_MAX).setNumberFormat('#,##0');

  // --- 見出し行（学年・組・氏名） -------------------------------------
  sh.getRange(DUES_HEADER_ROW, 1, 1, 3).setValues([['学年', '組', '氏名']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.getRange(DUES_HEADER_ROW, DUES_FIRST_TERM_COL, 1, DUES_TERMS_MAX).setBackground(COLOR.HEADER);

  // --- 部員データ（名簿の「在籍」だけを自動表示。チェックボックスで納入管理） -----
  sh.getRange(DUES_FIRST_DATA_ROW, 1).setFormula(
    '=IFERROR(FILTER({名簿!A' + ROSTER_FIRST_ROW + ':A' + ROSTER_LAST_ROW + ',名簿!B' + ROSTER_FIRST_ROW + ':B' + ROSTER_LAST_ROW + ',名簿!D' + ROSTER_FIRST_ROW + ':D' + ROSTER_LAST_ROW + '},' +
    '名簿!G' + ROSTER_FIRST_ROW + ':G' + ROSTER_LAST_ROW + '="在籍",名簿!D' + ROSTER_FIRST_ROW + ':D' + ROSTER_LAST_ROW + '<>""),"")'
  );
  sh.getRange(DUES_FIRST_DATA_ROW, DUES_FIRST_TERM_COL, ROSTER_MAX, DUES_TERMS_MAX).insertCheckboxes();

  // 未納セルを赤く（徴収額が入っている回だけ・チェックが入っていない人）
  var duesRange = sh.getRange(DUES_FIRST_DATA_ROW, DUES_FIRST_TERM_COL, ROSTER_MAX, DUES_TERMS_MAX);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(' + colLetter_(DUES_FIRST_TERM_COL) + '$' + DUES_AMOUNT_ROW + '<>"",' +
        colLetter_(DUES_FIRST_TERM_COL) + DUES_FIRST_DATA_ROW + '=FALSE)')
      .setBackground(COLOR.RED_LIGHT).setRanges([duesRange]).build()
  ]);

  sh.setFrozenRows(DUES_HEADER_ROW);
  sh.setFrozenColumns(3);
  sh.setColumnWidth(1, 40).setColumnWidth(2, 36).setColumnWidth(3, 120);
  for (var w = DUES_FIRST_TERM_COL; w <= DUES_LAST_COL; w++) sh.setColumnWidth(w, 78);
}

/* --------------------------------------------------------------------
 * 「年度末報告」シート（B-3）：全部数式・GAS不要。印刷してそのまま使える
 * ------------------------------------------------------------------ */
function buildReport_(sh) {
  resize_(sh, REPORT_LAST_ROW, REPORT_LAST_COL);
  sh.setTabColor('#795548');

  // 対象年度の開始日・終了日（年度開始月は「設定」シートB17。F2は対象年度の「開始年」）
  var fyStart = 'DATE($F$2,設定!$B$17,1)';
  var fyEnd   = 'EDATE(' + fyStart + ',12)';

  sh.getRange(1, 1, 1, REPORT_LAST_COL).merge()
    .setValue('＝　会計報告書　＝').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');

  sh.getRange('A2').setValue('部活名').setFontWeight('bold');
  sh.getRange('B2:D2').merge().setFormula('=設定!$B$16');
  sh.getRange('E2').setValue('対象年度（開始年）').setFontWeight('bold').setFontSize(9).setWrap(true);
  sh.getRange('F2').setFormula('=IF(MONTH(TODAY())>=設定!$B$17,YEAR(TODAY()),YEAR(TODAY())-1)')
    .setBackground(COLOR.CREAM).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('G2').setValue('年度').setFontColor('#666666');

  sh.getRange('A3').setValue('作成日').setFontWeight('bold');
  sh.getRange('B3').setFormula('=TODAY()').setNumberFormat('yyyy/M/d');
  sh.getRange('E3').setValue('顧問名').setFontWeight('bold');
  sh.getRange('F3:G3').merge().setFormula('=設定!$B$18');

  sh.getRange('A5:G5').merge().setValue('■ 収入の部').setFontWeight('bold').setBackground(COLOR.HEADER);

  var incomeRows = [
    { r: 6, item: '設定!$H$4' },
    { r: 7, item: '設定!$H$5' },
    { r: 8, item: '設定!$H$6' }
  ];
  incomeRows.forEach(function (row) {
    sh.getRange(row.r, 1, 1, 4).merge().setFormula('=' + row.item);
    sh.getRange(row.r, 6).setFormula(
      '=SUMIFS(会計簿!$E$' + ACCT_FIRST_ROW + ':$E$' + ACCT_LAST_ROW + ',' +
      '会計簿!$B$' + ACCT_FIRST_ROW + ':$B$' + ACCT_LAST_ROW + ',"収入",' +
      '会計簿!$C$' + ACCT_FIRST_ROW + ':$C$' + ACCT_LAST_ROW + ',$A' + row.r + ',' +
      '会計簿!$A$' + ACCT_FIRST_ROW + ':$A$' + ACCT_LAST_ROW + ',">="&' + fyStart + ',' +
      '会計簿!$A$' + ACCT_FIRST_ROW + ':$A$' + ACCT_LAST_ROW + ',"<"&' + fyEnd + ')'
    );
  });
  sh.getRange('A9:D9').merge().setValue('収入合計').setFontWeight('bold');
  sh.getRange('F9').setFormula('=SUM(F6:F8)').setFontWeight('bold').setBorder(true, false, false, false, false, false);

  sh.getRange('A11:G11').merge().setValue('■ 支出の部').setFontWeight('bold').setBackground(COLOR.HEADER);

  var expenseRows = [
    { r: 12, item: '設定!$I$4' },
    { r: 13, item: '設定!$I$5' },
    { r: 14, item: '設定!$I$6' },
    { r: 15, item: '設定!$I$7' },
    { r: 16, item: '設定!$I$8' }
  ];
  expenseRows.forEach(function (row) {
    sh.getRange(row.r, 1, 1, 4).merge().setFormula('=' + row.item);
    sh.getRange(row.r, 6).setFormula(
      '=SUMIFS(会計簿!$E$' + ACCT_FIRST_ROW + ':$E$' + ACCT_LAST_ROW + ',' +
      '会計簿!$B$' + ACCT_FIRST_ROW + ':$B$' + ACCT_LAST_ROW + ',"支出",' +
      '会計簿!$C$' + ACCT_FIRST_ROW + ':$C$' + ACCT_LAST_ROW + ',$A' + row.r + ',' +
      '会計簿!$A$' + ACCT_FIRST_ROW + ':$A$' + ACCT_LAST_ROW + ',">="&' + fyStart + ',' +
      '会計簿!$A$' + ACCT_FIRST_ROW + ':$A$' + ACCT_LAST_ROW + ',"<"&' + fyEnd + ')'
    );
  });
  sh.getRange('A17:D17').merge().setValue('支出合計').setFontWeight('bold');
  sh.getRange('F17').setFormula('=SUM(F12:F16)').setFontWeight('bold').setBorder(true, false, false, false, false, false);

  sh.getRange('A19:D19').merge().setValue('差引収支（収入－支出）').setFontWeight('bold');
  sh.getRange('F19').setFormula('=F9-F17').setFontWeight('bold');

  sh.getRange('A20:D20').merge().setValue('期首残高');
  sh.getRange('F20').setFormula(
    '=会計簿!$B$1' +
    '+SUMIFS(会計簿!$E$' + ACCT_FIRST_ROW + ':$E$' + ACCT_LAST_ROW + ',会計簿!$B$' + ACCT_FIRST_ROW + ':$B$' + ACCT_LAST_ROW + ',"収入",会計簿!$A$' + ACCT_FIRST_ROW + ':$A$' + ACCT_LAST_ROW + ',"<"&' + fyStart + ')' +
    '-SUMIFS(会計簿!$E$' + ACCT_FIRST_ROW + ':$E$' + ACCT_LAST_ROW + ',会計簿!$B$' + ACCT_FIRST_ROW + ':$B$' + ACCT_LAST_ROW + ',"支出",会計簿!$A$' + ACCT_FIRST_ROW + ':$A$' + ACCT_LAST_ROW + ',"<"&' + fyStart + ')'
  );

  sh.getRange('A21:D21').merge().setValue('期末残高').setFontWeight('bold').setFontSize(13);
  sh.getRange('F21').setFormula('=F20+F19').setFontWeight('bold').setFontSize(13)
    .setBorder(true, true, true, true, false, false);

  [9, 12, 13, 14, 15, 16, 17, 19, 20, 21].forEach(function (r) {
    sh.getRange(r, 6).setNumberFormat('#,##0"円"');
  });

  sh.getRange('A23:G23').merge()
    .setValue('※このシートはすべて自動集計です（編集不要）。印刷してそのままお使いいただけます（ファイル→印刷→用紙：A4→現在のシートのみ）。E2〜F2の「対象年度」を変えると、過去年度分もこの1枚で確認できます。')
    .setFontColor('#666666').setWrap(true).setFontSize(9);

  sh.setColumnWidths(1, 5, 90);
  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 60);
}

/* --------------------------------------------------------------------
 * 「はじめに」シート：購入者向けの使い方ガイド（商品に同梱される文面）
 * ------------------------------------------------------------------ */
function buildGuide_(sh) {
  resize_(sh, 80, 3);
  sh.setTabColor('#ea4335');

  var lines = [
    '🎾 部活動まるごと管理キット【出欠＋会計】',
    '毎日の出欠つけを「一覧を見て、休みの人だけ選んで、ボタン1回」の2〜3分に。会計簿・部費徴収・年度末報告もこの1つのファイルで完結します。現役の部活動顧問が自分の部で使っている仕組みです。',
    '※大会エントリー管理・連絡テンプレ文例集は今後追加予定です（このファイルには含まれていません）',
    '',
    '■ はじめかた（3ステップ・約5分）',
    '① このスプレッドシートをコピーする（もうお済みのはずです）',
    '② 「名簿」シートの見本（架空の名前）を消して、部員名簿を貼り付ける（学年・組・番号・氏名。「在籍」列は「在籍」のまま）',
    '③ 上のメニュー「🎾部活キット」→「✅ 今日の出欠を記録する」を一度実行して、許可（承認）画面を進める',
    '※途中で「このアプリは Google で確認されていません」という画面が出ますが、正常です。「詳細」→「（安全ではないページ）に移動」で進めてください。このシートはあなたのアカウントの中だけで動き、外部にデータは送られません。',
    '',
    '■ 毎日の使い方（2〜3分）',
    '1. 「今日の出欠」シートを開く（入力はスマホのスプレッドシートアプリでもOK）',
    '2. 欠席・遅刻・早退の部員だけ「区分」を選ぶ。全員そろっていれば何もしなくてOK（空欄＝出席）',
    '3. 必要なら「理由」をひとことメモ',
    '4. メニュー「🎾部活キット」→「✅ 今日の出欠を記録する」。出欠簿に○×が転記され、入力欄がリセットされます',
    '※メニューの実行だけはパソコンのブラウザから行ってください（スマホアプリではメニューが表示されない仕様のため）。コート脇でスマホ入力→職員室のPCで記録、の流れがおすすめです。',
    '',
    '■ 区分と記号のルール',
    '・標準の区分：欠席(×)・遅刻(遅)・早退(早)・遅早(遅早)。空欄は出席(○)として記録されます',
    '・区分は「設定」シートで自由に追加できます（例：公欠・見学・別メニュー・テスト期間）。記号欄が空のときは区分名の1文字目が記号になります',
    '・出席率は「欠席(×)だけを欠席として数える」計算です（遅刻・早退・追加区分は出席扱い）。学校の運用に合わせた計算式の調整はカスタム対応で承ります',
    '',
    '■ 知っておくと便利',
    '・「理由」は出欠簿のセルにメモ（注釈）として残ります。セルにマウスを乗せると読めます',
    '・過去の日付の記録もできます。「今日の出欠」の日付セル（B1）を変えて実行すると、日付順の正しい位置に列が入ります',
    '・同じ日にもう一度実行すると「上書きしますか？」と確認が出ます（つけ直しOK）',
    '・引退・退部した部員は「名簿」の「在籍」を変えるだけ。行は消さないでください（過去の記録が残ります）',
    '・名簿の変更（貼り替え・在籍の変更）は、その日の記録を終えてから行ってください',
    '',
    '■ 会計簿の使い方（記帳するだけでOK）',
    '・「会計簿」シートに、日付・種別（収入／支出）・費目・内容・金額・領収書の有無を入力するだけで、残高（F列）は自動で計算されます',
    '・B1セル「期首残高」は、導入時または年度はじめに、前年度からの繰越金を入力してください（初めて使う場合は0のままでOK）',
    '・費目の選択肢は「設定」シートで学校ごとに追加・変更できます',
    '・残高がマイナスになった行は自動で赤字表示になります。「支出」なのに「領収書：なし」の行は黄色で注意表示されます（記帳もれチェック用）',
    '',
    '■ 部費徴収の使い方',
    '・「部費徴収」シートの13〜15行目に、徴収回の名前・1人あたりの徴収額・締切を入力すると、その回が使用可能になります（徴収額が空欄の回は集計対象外です。見本として第1回・3,000円を入力してあります）',
    '・部員ごとにチェックボックスへ入金済みかどうかチェックしてください。未チェックの人は自動で赤くなり、シート上部に「未納者一覧」としてそのまま声かけリストが表示されます',
    '・入金済み合計・未納人数は16〜17行目に自動集計されます。この金額を確認してから、「会計簿」に収入として1行記帳してください（二重計上を防ぐため、部費徴収シートから会計簿への自動転記はしません）',
    '',
    '■ 年度末報告の使い方（印刷してそのまま提出できます）',
    '・「年度末報告」シートは全部自動集計です。入力・編集は不要です',
    '・F2セル「対象年度」で見たい年度（開始年。例：2026）を指定できます',
    '・部活名・顧問名は「設定」シートの内容がそのまま反映されます（匿名運用の場合、顧問名は空欄のままでもかまいません）',
    '・印刷は「ファイル→印刷→用紙：A4→現在のシートのみ」で、この1枚がそのまま職員会議・保護者会向けの会計報告書になります',
    '',
    '■ 色のルール（全シート共通）',
    '・うすいクリーム色のセル＝入力してOKな場所　／　白いセル＝自動計算（触らない）。迷ったら「色のついたところだけ触る」でOKです',
    '',
    '■ 年度更新（4月にやること）',
    '① ファイル →「コピーを作成」で今年度分をバックアップ',
    '② 「出欠簿」と「出欠サマリ」の2つのシートを削除する（必ず2つセットで）',
    '③ メニュー「🎾部活キット」→「🛠 初期セットアップ」を実行（2つのシートが空の状態で作り直されます）',
    '④ 「名簿」を新年度のものに貼り替える',
    '⑤ 「会計簿」はリセット不要です（年度をまたいで記帳を続けてください。「年度末報告」のF2セルで年度ごとの集計を確認できます）',
    '',
    '■ 個人情報の取り扱い（必ずお読みください）',
    '・氏名・出欠は個人情報です。各校の情報管理規程に従って運用してください',
    '・学校アカウントの共有ドライブ内で使い、共有設定は「制限付き」のまま使ってください（「リンクを知っている全員」にしない）',
    '・校外・私用アカウントへの持ち出しはしないでください',
    '・「備考」欄への記載は最小限に。保護者連絡先はこのキットではあえて扱いません（漏えいリスクを構造的に避けるための設計です）',
    '・「部費徴収」シートの氏名も名簿由来の個人情報です。出欠情報と同様の管理をお願いします',
    '',
    '■ 困ったとき',
    '・メニュー「🎾部活キット」が出ない → ページを再読み込みして数秒待つ。それでも出なければ、いったんブラウザで開き直す',
    '・シートを壊してしまった → 「設定」シートの「🆘 初期状態に戻したいとき」を参照',
    '・「今日の出欠」に部員が出ない → 「名簿」の「在籍」列が「在籍」になっているか確認',
    '・「部費徴収」に部員が出ない → 同じく「名簿」の「在籍」列を確認',
    '・「会計簿」の残高がおかしい → B1セル（期首残高）と、各行の「種別」「金額」の入力が正しいか確認'
  ];

  var values = lines.map(function (t) { return [t]; });
  sh.getRange(1, 1, values.length, 1).setValues(values).setWrap(true);

  // 見出し行を太字に、タイトルを大きく
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].indexOf('■') === 0) {
      sh.getRange(i + 1, 1).setFontWeight('bold').setFontSize(11).setBackground(COLOR.HEADER);
    }
    if (lines[i].indexOf('※') === 0) {
      sh.getRange(i + 1, 1).setFontColor('#b45309');
    }
  }
  sh.setColumnWidth(1, 780);
  sh.setHiddenGridlines(true);
}
