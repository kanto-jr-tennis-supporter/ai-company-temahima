/**
 * GAS スクリプト: ToGリスト（山元作業中）
 * テマヒマ・ラボ / ハック作成 / 2026-07-02
 *
 * 対象スプレッドシート: 「ToGリスト（山元作業中）」
 *
 * 【列構成】
 * コールシート (全県リスト統一):
 *   A(1)商号 B(2)代表者名 C(3)住所 D(4)電話番号
 *   E(5)担当① F(6)日付① G(7)時間① H(8)コール状況① I(9)コメント①
 *   J(10)担当② K(11)日付② L(12)時間② M(13)コール状況② N(14)コメント②
 *   O(15)担当③ P(16)日付③ Q(17)時間③ R(18)コール状況③ S(19)コメント③
 *   T(20)担当④ U(21)日付④ V(22)時間④ W(23)コール状況④ X(24)コメント④
 *
 * 資料送付リスト（G列にメールアドレスを追加した構成）:
 *   A(1)ランク B(2)元シート C(3)商号 D(4)代表者名 E(5)住所 F(6)電話番号 G(7)メールアドレス H(8)発送方法
 *   I(9)担当① J(10)日付① K(11)時間① L(12)コメント①
 *   M(13)担当② N(14)日付② O(15)時間② P(16)コメント②
 *   Q(17)担当③ R(18)日付③ S(19)時間③ T(20)コメント③
 *   U(21)担当④ V(22)日付④ W(23)時間④ X(24)コメント④
 */

// ============================================================
// 定数定義
// ============================================================

/** コールシートとして扱わないシート名の一覧 */
const EXCLUDE_SHEETS = [
  "アポイントカレンダー", "コール数集計", "アポ時間_ログ", "アポ時間_元データ",
  "アポ取得時間集計", "【編集厳禁】架電リスト元データ", "【アポ】管理シート",
  "追客対応リスト", "資料発送一覧", "プルダウン", "シート12", "_DEST_INDEX",
  "応対リスト", "資料送付リスト_千葉 静岡 茨城 福井 神奈川 愛知",
  "コールログ", "アポ時間_ログ"
];

/** 資料送付リストのシート名 */
const SHEET_MATERIAL = "資料送付リスト_千葉 静岡 茨城 福井 神奈川 愛知";

/** 再アプローチリスト（別ファイル）のスプレッドシートID とシート名 */
const SS_REAPPROACH_ID = "1WirnvlmYVOvXSSzCIKKlhQjrRwvDjlr86CNRgryhqj4";
const SHEET_REAPPROACH = "再アプローチリスト";

/** コールログのシート名 */
const SHEET_CALLLOG = "コールログ";

/** プルダウン選択肢 */
const DROPDOWN_VALUES = [
  "【受付拒否】", "【代表(担当)拒否】", "【不通】", "【番号相違】",
  "【TEL禁】", "【留守電】", "【186】", "【不在】",
  "【本社問い合わせ】", "【資料メール】", "【アポ獲得】"
];

/** コール状況列のインデックス（1始まり）※担当者・メールアドレス列追加後 */
const STATUS_COLS = [8, 15, 22, 29]; // H, O, V, AC

/** コメント列のインデックス（1始まり）※担当者・メールアドレス列追加後 */
const COMMENT_COLS = [9, 16, 23, 30]; // I, P, W, AD

/** 各コール回の列インデックス（担当者・メールアドレス列追加後）
 * 1回目: E(5)担当 F(6)日付 G(7)時間 H(8)コール状況 I(9)コメント J(10)担当者 K(11)メールアドレス
 * 2回目: L(12)〜R(18)
 * 3回目: S(19)〜Y(25)
 * 4回目: Z(26)〜AF(32)
 */
const CALL_ROUNDS = [
  { staff: 5,  date: 6,  time: 7,  status: 8,  comment: 9,  person: 10, email: 11 },
  { staff: 12, date: 13, time: 14, status: 15, comment: 16, person: 17, email: 18 },
  { staff: 19, date: 20, time: 21, status: 22, comment: 23, person: 24, email: 25 },
  { staff: 26, date: 27, time: 28, status: 29, comment: 30, person: 31, email: 32 },
];

/** データ開始行（ヘッダーの次） */
const DATA_START_ROW = 2;

// ============================================================
// ユーティリティ
// ============================================================

/** 「資料いらない」系のコメントが含まれるか判定 */
function hasNoMaterialComment_(rowData, maxCol) {
  const pattern = /資料いらない|資料不要|いらない|不要/;
  return CALL_ROUNDS.some(round =>
    round.comment <= maxCol && pattern.test(String(rowData[round.comment - 1] || ""))
  );
}

/** メールアドレスをテキストから抽出（最初の1件。無ければ空文字） */
function extractEmail_(text) {
  const m = String(text).match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/);
  return m ? m[0] : "";
}

/** コールシートの一覧を返す */
function getCallSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().filter(s => !EXCLUDE_SHEETS.includes(s.getName()));
}

/** 電話番号を数字のみに正規化 */
function normalizeTel_(tel) {
  return String(tel).replace(/\D/g, "");
}

/** 再アプローチリストの電話番号セットを返す（重複防止用） */
function getReapproachTels_() {
  const ss = SpreadsheetApp.openById(SS_REAPPROACH_ID);
  const sheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!sheet) return new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return new Set();
  const tels = sheet.getRange(DATA_START_ROW, 6, lastRow - DATA_START_ROW + 1, 1).getValues(); // F列=電話番号
  return new Set(tels.map(r => normalizeTel_(r[0])).filter(t => t.length > 0));
}

/** 資料送付リストの電話番号セットを返す（重複防止用） */
function getMaterialTels_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MATERIAL);
  if (!sheet) return new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return new Set();
  const tels = sheet.getRange(DATA_START_ROW, 6, lastRow - DATA_START_ROW + 1, 1).getValues(); // F列=電話番号
  return new Set(tels.map(r => normalizeTel_(r[0])).filter(t => t.length > 0));
}

// ============================================================
// 1. setup_dropdowns() ── プルダウン一括設定
// ============================================================

/**
 * 全コールシートのコール状況列（H/M/R/W）にプルダウンを設定する。
 * メニューまたは手動で1回実行すればOK。再実行で上書き更新も可能。
 */
function setup_dropdowns() {
  const sheets = getCallSheets_();
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES, true)
    .setAllowInvalid(true)
    .build();

  let count = 0;
  sheets.forEach(sheet => {
    const lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW);
    const numRows = lastRow - DATA_START_ROW + 1;
    if (numRows < 1) return;

    const maxCol = sheet.getLastColumn();
    STATUS_COLS.forEach(col => {
      if (col > maxCol) return; // 列が存在しない場合はスキップ
      sheet.getRange(DATA_START_ROW, col, numRows, 1).setDataValidation(rule);
    });
    count++;
  });

  SpreadsheetApp.getUi().alert(`プルダウン設定完了: ${count} シートに適用しました`);
}

// ============================================================
// 2. setup_conditional() ── 条件付き書式一括設定
// ============================================================

/**
 * 全コールシートにA〜X列の条件付き書式（色分け）を設定する。
 * 既存ルールをクリアしてから設定し直す。
 * 優先順位: アポ獲得（赤） > 資料メール（水色） > NG系（グレー）
 */
function setup_conditional() {
  const sheets = getCallSheets_();
  let count = 0;

  sheets.forEach(sheet => {
    const lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW + 10);
    const numRows = lastRow - DATA_START_ROW + 1;
    if (numRows < 1) return;

    // 既存の条件付き書式をクリア
    sheet.clearConditionalFormatRules();

    const numCols = Math.min(32, sheet.getLastColumn());
    if (numCols < 1) return;
    const range = sheet.getRange(DATA_START_ROW, 1, numRows, numCols); // A〜AF列（実在する列まで）
    const newRules = [];

    // NG系（グレー濃い #757575、文字白） ── 優先度最低なので最後に追加
    const ruleNG = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=REGEXMATCH(TEXTJOIN(" ",TRUE,$H${DATA_START_ROW},$O${DATA_START_ROW},$V${DATA_START_ROW},$AC${DATA_START_ROW}),"TEL禁|受付拒否|不通|番号相違|断り|186|代表")`
      )
      .setBackground("#757575")
      .setFontColor("#FFFFFF")
      .setRanges([range])
      .build();

    // 資料メール（水色 #00BCD4） ── 優先度中
    const ruleMaterial = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=REGEXMATCH(TEXTJOIN(" ",TRUE,$H${DATA_START_ROW},$O${DATA_START_ROW},$V${DATA_START_ROW},$AC${DATA_START_ROW}),"資料")`
      )
      .setBackground("#00BCD4")
      .setRanges([range])
      .build();

    // アポ獲得（赤 #FF0000、文字白） ── 優先度最高なので最初に追加
    const ruleApo = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=REGEXMATCH(TEXTJOIN(" ",TRUE,$H${DATA_START_ROW},$O${DATA_START_ROW},$V${DATA_START_ROW},$AC${DATA_START_ROW}),"アポ")`
      )
      .setBackground("#FF0000")
      .setFontColor("#FFFFFF")
      .setRanges([range])
      .build();

    // GASの条件付き書式は配列の先頭が最高優先度
    newRules.push(ruleApo, ruleMaterial, ruleNG);
    sheet.setConditionalFormatRules(newRules);
    count++;
  });

  SpreadsheetApp.getUi().alert(`条件付き書式設定完了: ${count} シートに適用しました`);
}

// ============================================================
// 3. onEdit(e) ── リアルタイム処理
// ============================================================

/**
 * セル編集時に自動実行される。
 * - コール状況列に「資料メール」→ 資料送付リストへ転記（電話番号で重複防止）
 * - 担当・日付・時間のいずれかが入力された → コールログに1行追記
 */
function onEditInstallable(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // 除外シートは処理しない
  if (EXCLUDE_SHEETS.includes(sheetName)) return;

  // 単一セル編集のみ対象
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  const row = range.getRow();
  const col = range.getColumn();

  // ヘッダー行は無視
  if (row < DATA_START_ROW) return;

  // その行のデータを表示値で取得（最大X列=24列分）
  const rowData = sheet.getRange(row, 1, 1, 32).getDisplayValues()[0];

  // --- コール状況列の編集処理 ---
  if (STATUS_COLS.includes(col)) {
    const newValue = String(e.value || "").trim();

    // 「資料メール」が入力されたら資料送付リスト＋再アプローチリストへ転記
    if (newValue === "【資料メール】") {
      transferToMaterialList_(sheet, row, col, rowData);
      transferToReapproach_(sheet, rowData);
    }

    // コール状況が入力されたらコールログにも記録
    const roundIndex = STATUS_COLS.indexOf(col);
    if (roundIndex >= 0) {
      const round = CALL_ROUNDS[roundIndex];
      appendCallLog_(sheetName, row, roundIndex + 1, rowData, round);
    }
    return;
  }

  // --- 担当・日付・時間列の編集処理 ---
  CALL_ROUNDS.forEach((round, idx) => {
    if ([round.staff, round.date, round.time].includes(col)) {
      // 担当・日付・時間どれかが入力されたらコールログへ
      appendCallLog_(sheetName, row, idx + 1, rowData, round);
    }
  });
}

/**
 * 資料送付リストへ転記する内部関数
 * @param {Sheet} srcSheet - 元シート
 * @param {number} row - 編集行
 * @param {number} statusCol - 編集されたコール状況列番号
 * @param {Array} rowData - 0始まりの行データ配列
 */
function transferToMaterialList_(srcSheet, row, statusCol, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = ss.getSheetByName(SHEET_MATERIAL);
  if (!destSheet) {
    Logger.log("資料送付リストシートが見つかりません: " + SHEET_MATERIAL);
    return;
  }

  // 「資料いらない」系コメントがあればスキップ
  if (hasNoMaterialComment_(rowData, 24)) {
    Logger.log("資料送付リスト: 資料不要のためスキップ " + rowData[0]);
    return;
  }

  // 電話番号で重複チェック
  const tel = normalizeTel_(rowData[3]); // D列(index3)=電話番号
  if (tel && getMaterialTels_().has(tel)) {
    Logger.log("資料送付リスト: 重複のためスキップ 電話=" + tel);
    return;
  }

  // 編集されたコール回のインデックスを特定
  const roundIndex = STATUS_COLS.indexOf(statusCol);
  if (roundIndex < 0) return;
  const round = CALL_ROUNDS[roundIndex];

  // 転記データを組み立て
  // A(ランク) B(元シート) C(商号) D(代表者名) E(住所) F(電話番号) G(メールアドレス) H(発送方法)
  // I(担当①) J(日付①) K(時間①) L(コメント①)
  const comment = rowData[round.comment - 1] || "";
  const email = extractEmail_(comment);
  const newRow = [
    "",                            // A: ランク（空）
    srcSheet.getName(),            // B: 元シート名
    rowData[0],                    // C: 商号
    rowData[1],                    // D: 代表者名
    rowData[2],                    // E: 住所
    rowData[3],                    // F: 電話番号
    email,                         // G: メールアドレス（コメントから抽出）
    "",                            // H: 発送方法（空）
    rowData[round.staff - 1],      // I: 担当
    rowData[round.date - 1],       // J: 日付
    rowData[round.time - 1],       // K: 時間
    comment,                       // L: コメント
  ];

  destSheet.appendRow(newRow);
  Logger.log("資料送付リストへ転記: " + rowData[0] + " (" + srcSheet.getName() + ")");
}

/**
 * コールログへ1行追記する内部関数
 * @param {string} sheetName - 元シート名
 * @param {number} row - 行番号
 * @param {number} roundNum - コール回数（1〜4）
 * @param {Array} rowData - 0始まりの行データ配列
 * @param {Object} round - CALL_ROUNDS の1エントリ
 */
function appendCallLog_(sheetName, row, roundNum, rowData, round) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(SHEET_CALLLOG);

  // コールログシートがなければ作成
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_CALLLOG);
    logSheet.appendRow([
      "記録日時", "元シート", "行", "回数", "商号", "電話番号",
      "担当", "日付", "時間", "コール状況", "コメント"
    ]);
  }

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  logSheet.appendRow([
    now,
    sheetName,
    row,
    roundNum,
    rowData[0],                    // 商号
    rowData[3],                    // 電話番号
    rowData[round.staff - 1],      // 担当
    rowData[round.date - 1],       // 日付
    rowData[round.time - 1],       // 時間
    rowData[round.status - 1],     // コール状況
    rowData[round.comment - 1],    // コメント
  ]);
}

// ============================================================
// 4. backfill_callStatus() ── 既存コール状況の一括推定補填（一回きり）
// ============================================================

/**
 * 既存データのコメント列（I/N/S/X）を読み、コール状況列（H/M/R/W）が
 * 空の場合のみパターンマッチで推定して書き込む。
 * 実行は1回だけ。完了後に件数をアラートで表示。
 */
function backfill_callStatus() {
  const sheets = getCallSheets_();
  let totalFilled = 0;

  /** コメントテキストからコール状況を推定 */
  function inferStatus(comment) {
    const c = String(comment);
    if (/アポ/.test(c)) return "【アポ獲得】";
    if (/資料/.test(c)) return "【資料メール】";
    if (/TEL禁/.test(c)) return "【TEL禁】";
    if (/受付拒否/.test(c)) return "【受付拒否】";
    if (/代表.*拒否|担当.*拒否|断り/.test(c)) return "【代表(担当)拒否】";
    if (/使われていない|現在使われ|番号相違|番号違/.test(c)) return "【番号相違】";
    if (/不通/.test(c)) return "【不通】";
    if (/留守電/.test(c)) return "【留守電】";
    if (/186/.test(c)) return "【186】";
    if (/本社/.test(c)) return "【本社問い合わせ】";
    if (/不在|外出|出張/.test(c)) return "【不在】";
    return null;
  }

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return;
    const numRows = lastRow - DATA_START_ROW + 1;

    // 全データを表示値で取得（A〜X = 24列）
    const allData = sheet.getRange(DATA_START_ROW, 1, numRows, Math.min(32, sheet.getLastColumn())).getDisplayValues();

    CALL_ROUNDS.forEach((round, idx) => {
      const statusColIdx = round.status - 1; // 0始まり
      const commentColIdx = round.comment - 1;

      allData.forEach((rowData, rIdx) => {
        const currentStatus = String(rowData[statusColIdx]).trim();
        const comment = String(rowData[commentColIdx]).trim();

        // コール状況が空でコメントがある場合のみ推定
        if (currentStatus === "" && comment !== "") {
          const inferred = inferStatus(comment);
          if (inferred) {
            const actualRow = DATA_START_ROW + rIdx;
            sheet.getRange(actualRow, round.status).setValue(inferred);
            totalFilled++;
          }
        }
      });
    });
  });

  SpreadsheetApp.getUi().alert(`コール状況 補填完了\n設定件数: ${totalFilled} 件`);
}

// ============================================================
// 5. backfill_callLog() ── 既存データからコールログ一括収集（一回きり）
// ============================================================

/**
 * 全コールシートの既存データをスキャンし、担当・日付・時間のいずれかが
 * 入力されている行をすべてコールログに収集する。
 * コールログシートを新規作成（既存があればクリア）。
 * 完了後に件数をアラートで表示。
 */
function backfill_callLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getCallSheets_();

  // コールログシートを準備（既存はクリアして作り直し）
  let logSheet = ss.getSheetByName(SHEET_CALLLOG);
  if (logSheet) {
    logSheet.clearContents();
  } else {
    logSheet = ss.insertSheet(SHEET_CALLLOG);
  }

  // ヘッダー行を書き込む
  logSheet.appendRow([
    "記録日時", "元シート", "行", "回数", "商号", "電話番号",
    "担当", "日付", "時間", "コール状況", "コメント"
  ]);

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  let totalRows = 0;
  const logBuffer = []; // まとめて書き込むためのバッファ

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return;
    const numRows = lastRow - DATA_START_ROW + 1;

    // 全データを表示値で取得（A〜X = 24列）
    const allData = sheet.getRange(DATA_START_ROW, 1, numRows, Math.min(32, sheet.getLastColumn())).getDisplayValues();

    allData.forEach((rowData, rIdx) => {
      const actualRow = DATA_START_ROW + rIdx;

      CALL_ROUNDS.forEach((round, idx) => {
        const staff   = String(rowData[round.staff - 1]).trim();
        const date    = String(rowData[round.date - 1]).trim();
        const time    = String(rowData[round.time - 1]).trim();
        const status  = String(rowData[round.status - 1]).trim();
        const comment = String(rowData[round.comment - 1]).trim();

        // 担当・日付・時間のどれかが入力されていれば収集
        if (staff !== "" || date !== "" || time !== "") {
          logBuffer.push([
            now,
            sheetName,
            actualRow,
            idx + 1,
            rowData[0],    // 商号
            rowData[3],    // 電話番号
            staff,
            date,
            time,
            status,
            comment,
          ]);
          totalRows++;
        }
      });
    });
  });

  // バッファをまとめてシートに書き込む（API呼び出し回数を削減）
  if (logBuffer.length > 0) {
    logSheet.getRange(2, 1, logBuffer.length, 11).setValues(logBuffer);
  }

  SpreadsheetApp.getUi().alert(`コールログ 収集完了\n収集件数: ${totalRows} 件`);
}

/**
 * 再アプローチリスト（別ファイル）へ転記する内部関数
 * 全4回分のコールデータをまとめて転記する
 */
function transferToReapproach_(srcSheet, rowData) {
  // 「資料いらない」系コメントがあればスキップ
  if (hasNoMaterialComment_(rowData, 24)) {
    Logger.log("再アプローチリスト: 資料不要のためスキップ " + rowData[0]);
    return;
  }

  const tel = normalizeTel_(rowData[3] || "");
  if (tel && getReapproachTels_().has(tel)) {
    Logger.log("再アプローチリスト: 重複のためスキップ 電話=" + tel);
    return;
  }

  const ss = SpreadsheetApp.openById(SS_REAPPROACH_ID);
  const destSheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!destSheet) {
    Logger.log("再アプローチリストシートが見つかりません: " + SHEET_REAPPROACH);
    return;
  }

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  const len = rowData.length;
  function get(col1) { return col1 <= len ? (rowData[col1 - 1] || "") : ""; }

  // メールアドレスを全コメント列から抽出（I/N/S/X = 9/14/19/24列目）
  const email = extractEmail_(get(9)) || extractEmail_(get(14)) ||
                extractEmail_(get(19)) || extractEmail_(get(24));

  // A:転記日時 B:元シート C:商号 D:代表者名 E:住所 F:電話番号 G:メールアドレス
  // H-L: ①担当,日付,時間,コール状況,コメント
  // M-Q: ②  R-V: ③  W-AA: ④
  const newRow = [
    now, srcSheet.getName(),
    get(1), get(2), get(3), get(4),
    email,
    get(5),  get(6),  get(7),  get(8),  get(9),
    get(10), get(11), get(12), get(13), get(14),
    get(15), get(16), get(17), get(18), get(19),
    get(20), get(21), get(22), get(23), get(24),
  ];

  destSheet.appendRow(newRow);
  Logger.log("再アプローチリストへ転記: " + get(1));
}

// ============================================================
// 6. backfill_materialList() ── 資料送付リストへの一括転記（一回きり）
// ============================================================

/**
 * 全コールシートをスキャンし、コール状況が「資料メール」になっている行を
 * 資料送付リストに転記する。電話番号で重複チェックを行い、未登録のものだけ追加。
 * 既存データを上書きしないので何度実行しても安全。
 */
function backfill_materialList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getCallSheets_();
  const destSheet = ss.getSheetByName(SHEET_MATERIAL);

  if (!destSheet) {
    SpreadsheetApp.getUi().alert("資料送付リストシートが見つかりません: " + SHEET_MATERIAL);
    return;
  }

  // 転記済み電話番号のセット（重複防止）
  const existingTels = getMaterialTels_();
  const newRows = [];
  let skipped = 0;

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return;
    const numRows = lastRow - DATA_START_ROW + 1;
    const maxCol = Math.min(32, sheet.getLastColumn());
    if (maxCol < 1) return;

    const allData = sheet.getRange(DATA_START_ROW, 1, numRows, maxCol).getDisplayValues();

    allData.forEach(rowData => {
      // 各コール回をチェック
      CALL_ROUNDS.forEach(round => {
        if (round.status > maxCol) return;
        const status = String(rowData[round.status - 1] || "").trim();
        if (status !== "【資料メール】") return;

        // 「資料いらない」系コメントがあればスキップ
        if (hasNoMaterialComment_(rowData, maxCol)) return;

        const tel = normalizeTel_(rowData[3] || "");

        // 重複チェック（既存＋今回追加分）
        if (tel && (existingTels.has(tel) || newRows.some(r => normalizeTel_(r[5]) === tel))) {
          skipped++;
          return;
        }

        const comment = round.comment <= maxCol ? (rowData[round.comment - 1] || "") : "";
        newRows.push([
          "",                                                        // A: ランク（空）
          sheet.getName(),                                           // B: 元シート名
          rowData[0] || "",                                          // C: 商号
          rowData[1] || "",                                          // D: 代表者名
          rowData[2] || "",                                          // E: 住所
          rowData[3] || "",                                          // F: 電話番号
          extractEmail_(comment),                                    // G: メールアドレス
          "",                                                        // H: 発送方法（空）
          round.staff   <= maxCol ? rowData[round.staff   - 1] : "", // I: 担当
          round.date    <= maxCol ? rowData[round.date    - 1] : "", // J: 日付
          round.time    <= maxCol ? rowData[round.time    - 1] : "", // K: 時間
          comment,                                                   // L: コメント
        ]);
      });
    });
  });

  if (newRows.length > 0) {
    const startRow = destSheet.getLastRow() + 1;
    destSheet.getRange(startRow, 1, newRows.length, 12).setValues(newRows); // 12列（メールアドレス追加）
  }

  SpreadsheetApp.getUi().alert(
    `資料送付リスト 一括転記 完了\n追加: ${newRows.length} 件　重複スキップ: ${skipped} 件`
  );
}

// ============================================================
// 7. backfill_reapproachList() ── 再アプローチリストへの一括転記（一回きり）
// ============================================================

/**
 * 全コールシートをスキャンし「資料メール」行を再アプローチリスト（別ファイル）に転記。
 * 電話番号で重複チェック。既存データを上書きしないので何度実行しても安全。
 */
function backfill_reapproachList() {
  const ss = SpreadsheetApp.openById(SS_REAPPROACH_ID);
  const destSheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!destSheet) {
    SpreadsheetApp.getUi().alert("再アプローチリストシートが見つかりません: " + SHEET_REAPPROACH);
    return;
  }

  const sheets = getCallSheets_();
  const existingTels = getReapproachTels_();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  const newRows = [];
  let skipped = 0;

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return;
    const numRows = lastRow - DATA_START_ROW + 1;
    const maxCol = Math.min(32, sheet.getLastColumn());
    if (maxCol < 1) return;

    const allData = sheet.getRange(DATA_START_ROW, 1, numRows, maxCol).getDisplayValues();

    allData.forEach(rowData => {
      // いずれかのコール回が「資料メール」かチェック
      const hasMaterial = CALL_ROUNDS.some(r =>
        r.status <= maxCol && String(rowData[r.status - 1]).trim() === "【資料メール】"
      );
      if (!hasMaterial) return;

      // 「資料いらない」系コメントがあればスキップ
      if (hasNoMaterialComment_(rowData, maxCol)) return;

      const tel = normalizeTel_(rowData[3] || "");
      if (tel && (existingTels.has(tel) || newRows.some(r => normalizeTel_(r[5]) === tel))) {
        skipped++;
        return;
      }

      function get(col1) { return col1 <= maxCol ? (rowData[col1 - 1] || "") : ""; }

      const email = extractEmail_(get(9)) || extractEmail_(get(14)) ||
                    extractEmail_(get(19)) || extractEmail_(get(24));
      newRows.push([
        now, sheet.getName(),
        get(1), get(2), get(3), get(4),
        email,
        get(5),  get(6),  get(7),  get(8),  get(9),
        get(10), get(11), get(12), get(13), get(14),
        get(15), get(16), get(17), get(18), get(19),
        get(20), get(21), get(22), get(23), get(24),
      ]);
    });
  });

  if (newRows.length > 0) {
    const startRow = destSheet.getLastRow() + 1;
    destSheet.getRange(startRow, 1, newRows.length, 27).setValues(newRows); // 27列（メールアドレス追加）
  }

  SpreadsheetApp.getUi().alert(
    `再アプローチリスト 一括転記 完了\n追加: ${newRows.length} 件　重複スキップ: ${skipped} 件`
  );
}

// ============================================================
// 8. backfill_sourceSheet() ── 既存行の「元シート」をコールログから補填（一回きり）
// ============================================================

/**
 * 資料送付リストのB列（元シート）が空になっている既存行を対象に、
 * コールログを電話番号で照合して元シート名を書き込む。
 * B列追加後に一度だけ実行する。
 */
function backfill_sourceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = ss.getSheetByName(SHEET_MATERIAL);
  const logSheet  = ss.getSheetByName(SHEET_CALLLOG);

  if (!destSheet) {
    SpreadsheetApp.getUi().alert("資料送付リストが見つかりません");
    return;
  }
  if (!logSheet) {
    SpreadsheetApp.getUi().alert("コールログが見つかりません（④を先に実行してください）");
    return;
  }

  const destLast = destSheet.getLastRow();
  if (destLast < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("資料送付リストにデータがありません");
    return;
  }

  // コールログから「元シート名」を電話番号でマップ化
  // コールログ列: 1記録日時 2元シート 3行 4回数 5商号 6電話番号 ...
  const logLast = logSheet.getLastRow();
  const telToSheet = {};
  if (logLast >= 2) {
    const logData = logSheet.getRange(2, 1, logLast - 1, 6).getValues();
    logData.forEach(r => {
      const sheetName = String(r[1]).trim();
      const tel = normalizeTel_(r[5]);
      if (tel && sheetName && !telToSheet[tel]) {
        telToSheet[tel] = sheetName;
      }
    });
  }

  // 資料送付リストのB列（元シート）が空の行を補填
  // 列構成: A(1)ランク B(2)元シート C(3)商号 ... F(6)電話番号
  const destData = destSheet.getRange(DATA_START_ROW, 1, destLast - DATA_START_ROW + 1, 6).getValues();
  let filled = 0;

  destData.forEach((row, idx) => {
    const sourceCell = String(row[1]).trim(); // B列=元シート
    if (sourceCell !== "") return;            // すでに入力済みはスキップ

    const tel = normalizeTel_(row[5]);        // F列=電話番号
    if (!tel) return;

    const matched = telToSheet[tel];
    if (matched) {
      destSheet.getRange(DATA_START_ROW + idx, 2).setValue(matched); // B列に書き込み
      filled++;
    }
  });

  SpreadsheetApp.getUi().alert(`元シート補填 完了\n補填件数: ${filled} 件`);
}

// ============================================================
// 全コールシートに「担当者・メールアドレス」列を一括挿入（一回きり）
// ============================================================

/**
 * 全コールシートのコメント列（I/N/S/X）の右に
 * 「担当者」「メールアドレス」の2列を挿入する。
 * 右から順に挿入するので既存データはずれない。
 * ★ 1回だけ実行。すでに挿入済みの場合は「担当者」ヘッダーを検出してスキップ。
 */
function setup_insertCallColumns() {
  const sheets = getCallSheets_();
  // コメント列（元の列番号）。右から順に処理する
  const commentColsOriginal = [24, 19, 14, 9]; // X, S, N, I（右から）
  const headers = ["担当者", "メールアドレス"];
  let count = 0;

  sheets.forEach(sheet => {
    // 右から順にコメント列の右へ2列挿入
    commentColsOriginal.forEach(commentCol => {
      const lastCol = sheet.getLastColumn();

      // コメント列自体が存在しない場合はスキップ
      if (commentCol > lastCol) return;

      // すでに「担当者」列が挿入済みの場合はスキップ（再実行時の重複防止）
      if (commentCol + 1 <= lastCol &&
          String(sheet.getRange(1, commentCol + 1).getValue()).trim() === "担当者") return;

      sheet.insertColumnsAfter(commentCol, 2);
      sheet.getRange(1, commentCol + 1).setValue(headers[0]);
      sheet.getRange(1, commentCol + 2).setValue(headers[1]);
    });

    count++;
  });

  SpreadsheetApp.getUi().alert(
    `列挿入完了: ${count} シートに「担当者・メールアドレス」列を追加しました。\n` +
    `続けて ① プルダウン設定・② 条件付き書式設定 を再実行してください。`
  );
}

// ============================================================
// 資料送付リストから「資料いらない」行を削除するクリーンアップ
// ============================================================

/**
 * 資料送付リストのコメント列（L/P/T/X）に「資料いらない」「不要」等が含まれる行を削除する。
 * 削除前に件数を確認するアラートを表示する。
 */
function cleanup_noMaterialRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MATERIAL);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("資料送付リストが見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  const numRows = lastRow - DATA_START_ROW + 1;
  const maxCol = Math.min(sheet.getLastColumn(), 24);
  // 資料送付リストのコメント列: L(12) P(16) T(20) X(24)
  const COMMENT_COLS_MATERIAL = [12, 16, 20, 24];
  const pattern = /資料いらない|資料不要|いらない|不要/;

  const data = sheet.getRange(DATA_START_ROW, 1, numRows, maxCol).getDisplayValues();

  // 削除対象の行番号を後ろから順に収集（前から消すとずれるため）
  const deleteRows = [];
  data.forEach((row, idx) => {
    const hasNG = COMMENT_COLS_MATERIAL.some(col =>
      col <= maxCol && pattern.test(String(row[col - 1] || ""))
    );
    if (hasNG) deleteRows.push(DATA_START_ROW + idx);
  });

  if (deleteRows.length === 0) {
    SpreadsheetApp.getUi().alert("「資料いらない」に該当する行はありませんでした");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    `確認`,
    `「資料いらない」に該当する行が ${deleteRows.length} 件あります。削除しますか？`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // 後ろから削除（行番号のずれを防ぐ）
  deleteRows.reverse().forEach(rowNum => sheet.deleteRow(rowNum));

  SpreadsheetApp.getUi().alert(`削除完了: ${deleteRows.length} 件を削除しました`);
}

// ============================================================
// メールアドレス列のセットアップ
// ============================================================

/**
 * 資料送付リストにメールアドレス列（G列）を挿入してヘッダーを書き込む。
 * 既存のG列（発送方法）は右にずれてH列になる。
 * ★ 1回だけ実行。すでに列がある場合は再実行しない。
 */
function setup_emailColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MATERIAL);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("資料送付リストが見つかりません");
    return;
  }

  // G1のヘッダーがすでに「メールアドレス」なら二重挿入を防ぐ
  const existingHeader = sheet.getRange(1, 7).getValue();
  if (String(existingHeader).trim() === "メールアドレス") {
    SpreadsheetApp.getUi().alert("メールアドレス列はすでに存在します。スキップしました。");
    return;
  }

  // G列（7列目）の左に1列挿入
  sheet.insertColumnBefore(7);
  sheet.getRange(1, 7).setValue("メールアドレス");

  SpreadsheetApp.getUi().alert("G列に「メールアドレス」列を挿入しました。\n続けて「メールアドレス 補填」を実行してください。");
}

/**
 * 資料送付リストの既存行のコメント列（L/P/T/X）を読み、
 * G列（メールアドレス）が空の行にメールアドレスを抽出して書き込む。
 * 何度実行しても空欄のみ補填するので安全。
 */
function backfill_emailExtract() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MATERIAL);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("資料送付リストが見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("データがありません");
    return;
  }

  const numRows = lastRow - DATA_START_ROW + 1;
  const maxCol = sheet.getLastColumn();
  // 現在の列構成: G(7)メール H(8)発送方法 I(9)担当① J(10)日付① K(11)時間① L(12)コメント①
  //              M(13)担当② N(14)日付② O(15)時間② P(16)コメント②
  //              Q(17)担当③ R(18)日付③ S(19)時間③ T(20)コメント③
  //              U(21)担当④ V(22)日付④ W(23)時間④ X(24)コメント④
  const EMAIL_COL   = 7;  // G
  const COMMENT_COLS_MATERIAL = [12, 16, 20, 24]; // L, P, T, X

  const data = sheet.getRange(DATA_START_ROW, 1, numRows, Math.min(maxCol, 24)).getDisplayValues();
  let filled = 0;

  data.forEach((row, idx) => {
    const currentEmail = String(row[EMAIL_COL - 1] || "").trim();
    if (currentEmail !== "") return; // すでに入力済みはスキップ

    // 全コメント列からメールを探す
    let found = "";
    for (const commentCol of COMMENT_COLS_MATERIAL) {
      if (commentCol > maxCol) break;
      found = extractEmail_(row[commentCol - 1] || "");
      if (found) break;
    }

    if (found) {
      sheet.getRange(DATA_START_ROW + idx, EMAIL_COL).setValue(found);
      filled++;
    }
  });

  SpreadsheetApp.getUi().alert(`メールアドレス補填 完了\n補填件数: ${filled} 件`);
}

// ============================================================
// 再アプローチリストから「資料いらない」行を削除するクリーンアップ
// ============================================================

/**
 * 再アプローチリスト（別ファイル）のコメント列（L/Q/V/AA）に
 * 「資料いらない」「不要」等が含まれる行を削除する。
 */
function cleanup_reapproach_noMaterialRows() {
  const ss = SpreadsheetApp.openById(SS_REAPPROACH_ID);
  const sheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("再アプローチリストが見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  const numRows = lastRow - DATA_START_ROW + 1;
  const maxCol = Math.min(sheet.getLastColumn(), 27);
  // 再アプローチリストのコメント列: L(12) Q(17) V(22) AA(27)
  const COMMENT_COLS_REAPPROACH = [12, 17, 22, 27];
  const pattern = /資料いらない|資料不要|いらない|不要/;

  const data = sheet.getRange(DATA_START_ROW, 1, numRows, maxCol).getDisplayValues();
  const deleteRows = [];
  data.forEach((row, idx) => {
    const hasNG = COMMENT_COLS_REAPPROACH.some(col =>
      col <= maxCol && pattern.test(String(row[col - 1] || ""))
    );
    if (hasNG) deleteRows.push(DATA_START_ROW + idx);
  });

  if (deleteRows.length === 0) {
    SpreadsheetApp.getUi().alert("該当する行はありませんでした");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    `確認`,
    `再アプローチリストに「資料いらない」行が ${deleteRows.length} 件あります。削除しますか？`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  deleteRows.reverse().forEach(rowNum => sheet.deleteRow(rowNum));
  SpreadsheetApp.getUi().alert(`削除完了: ${deleteRows.length} 件を削除しました`);
}

// ============================================================
// 再アプローチリストのメールアドレス補填（一回きり）
// ============================================================

/**
 * 再アプローチリストの既存行のコメント列（L/Q/V/AA）を読み、
 * G列（メールアドレス）が空の行にメールアドレスを抽出して書き込む。
 */
function backfill_reapproach_emailExtract() {
  const ss = SpreadsheetApp.openById(SS_REAPPROACH_ID);
  const sheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("再アプローチリストが見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("データがありません");
    return;
  }

  const numRows = lastRow - DATA_START_ROW + 1;
  const maxCol = Math.min(27, sheet.getLastColumn());
  // 列構成: G(7)メール H(8)担当① ... L(12)コメント① Q(17)コメント② V(22)コメント③ AA(27)コメント④
  const EMAIL_COL = 7;
  const COMMENT_COLS_REAPPROACH = [12, 17, 22, 27]; // L, Q, V, AA

  const data = sheet.getRange(DATA_START_ROW, 1, numRows, maxCol).getDisplayValues();
  let filled = 0;

  data.forEach((row, idx) => {
    if (String(row[EMAIL_COL - 1] || "").trim() !== "") return; // 入力済みはスキップ

    let found = "";
    for (const col of COMMENT_COLS_REAPPROACH) {
      if (col > maxCol) break;
      found = extractEmail_(row[col - 1] || "");
      if (found) break;
    }

    if (found) {
      sheet.getRange(DATA_START_ROW + idx, EMAIL_COL).setValue(found);
      filled++;
    }
  });

  SpreadsheetApp.getUi().alert(`再アプローチリスト メールアドレス補填 完了\n補填件数: ${filled} 件`);
}

// ============================================================
// 再アプローチリスト（別ファイル）のプルダウン・条件付き書式設定
// ============================================================

/** 再アプローチリストのコール状況列（G列メールアドレス追加後: K/P/U/Z） */
const REAPPROACH_STATUS_COLS = [11, 16, 21, 26]; // K, P, U, Z

/**
 * 再アプローチリストのコール状況列にプルダウンを設定する
 */
function setup_reapproach_dropdowns() {
  const ss = SpreadsheetApp.openById(SS_REAPPROACH_ID);
  const sheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("再アプローチリストシートが見つかりません");
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DROPDOWN_VALUES, true)
    .setAllowInvalid(true)
    .build();

  const lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW);
  const numRows = lastRow - DATA_START_ROW + 1;
  const maxCol = sheet.getLastColumn();

  REAPPROACH_STATUS_COLS.forEach(col => {
    if (col > maxCol) return;
    sheet.getRange(DATA_START_ROW, col, numRows, 1).setDataValidation(rule);
  });

  SpreadsheetApp.getUi().alert("再アプローチリスト：プルダウン設定完了");
}

/**
 * 再アプローチリストに条件付き書式（色分け）を設定する
 * 優先順位: アポ獲得（赤） > 資料メール（水色） > NG系（グレー）
 */
function setup_reapproach_conditional() {
  const ss = SpreadsheetApp.openById(SS_REAPPROACH_ID);
  const sheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("再アプローチリストシートが見つかりません");
    return;
  }

  sheet.clearConditionalFormatRules();

  const lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW + 10);
  const numRows = lastRow - DATA_START_ROW + 1;
  const numCols = Math.min(26, sheet.getLastColumn()); // A〜Z列
  if (numCols < 1) return;

  const range = sheet.getRange(DATA_START_ROW, 1, numRows, numCols);

  // コール状況列: $J2, $O2, $T2, $Y2
  const formula = (keyword) =>
    `=REGEXMATCH(TEXTJOIN(" ",TRUE,$K${DATA_START_ROW},$P${DATA_START_ROW},$U${DATA_START_ROW},$Z${DATA_START_ROW}),"${keyword}")`;

  const ruleApo = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula("アポ"))
    .setBackground("#FF0000").setFontColor("#FFFFFF")
    .setRanges([range]).build();

  const ruleMaterial = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula("資料"))
    .setBackground("#00BCD4")
    .setRanges([range]).build();

  const ruleNG = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula("TEL禁|受付拒否|不通|番号相違|断り|186|代表"))
    .setBackground("#757575").setFontColor("#FFFFFF")
    .setRanges([range]).build();

  sheet.setConditionalFormatRules([ruleApo, ruleMaterial, ruleNG]);
  SpreadsheetApp.getUi().alert("再アプローチリスト：条件付き書式設定完了");
}

// ============================================================
// メニュー追加（任意）
// ============================================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する。
 * GAS エディタでこの関数をトリガー（onOpen）に設定すると、
 * スプレッドシートを開くたびにメニューが表示される。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 ToGツール")
    .addItem("0️⃣【最初に1回】担当者・メールアドレス列を挿入", "setup_insertCallColumns")
    .addSeparator()
    .addItem("① プルダウン設定",     "setup_dropdowns")
    .addItem("② 条件付き書式設定",   "setup_conditional")
    .addSeparator()
    .addItem("③【一回きり】コール状況 補填",       "backfill_callStatus")
    .addItem("④【一回きり】コールログ 収集",       "backfill_callLog")
    .addItem("⑤【一回きり】資料送付リスト 一括転記",             "backfill_materialList")
    .addItem("⑥【一回きり】再アプローチリスト 一括転記",         "backfill_reapproachList")
    .addItem("⑦【一回きり】元シート名 補填",                   "backfill_sourceSheet")
    .addItem("🗑️ 資料送付リスト：資料いらない行を削除",          "cleanup_noMaterialRows")
    .addItem("🗑️ 再アプローチリスト：資料いらない行を削除",      "cleanup_reapproach_noMaterialRows")
    .addSeparator()
    .addItem("📧【一回きり】メールアドレス列を挿入",             "setup_emailColumn")
    .addItem("📧【一回きり】メールアドレス補填",                 "backfill_emailExtract")
    .addSeparator()
    .addItem("📧【一回きり】再アプローチリスト：メールアドレス補填", "backfill_reapproach_emailExtract")
    .addSeparator()
    .addItem("🔁 再アプローチリスト：プルダウン設定",            "setup_reapproach_dropdowns")
    .addItem("🔁 再アプローチリスト：条件付き書式設定",          "setup_reapproach_conditional")
    .addToUi();
}

/*
=======================================================================
【実行順序 手順書】

■ 初回セットアップ（このスクリプトを貼り付けた直後に行う順番）

STEP 1: GAS エディタにこのスクリプトを貼り付けて保存
  - スプレッドシートのメニュー「拡張機能」→「Apps Script」を開く
  - 既存コードを全削除し、このスクリプトを貼り付けて保存（Ctrl+S / Cmd+S）

STEP 2: onOpen を手動実行してメニューを確認（任意）
  - GAS エディタ上部の関数選択から「onOpen」を選び「▶ 実行」
  - スプレッドシートに戻り「📋 ToGツール」メニューが出ていればOK

STEP 3: ① プルダウン設定 を実行
  - メニュー「📋 ToGツール」→「① プルダウン設定」
  - 全コールシートのH・M・R・W列にプルダウンが設定される
  - 完了アラートで設定シート数を確認する

STEP 4: ② 条件付き書式設定 を実行
  - メニュー「📋 ToGツール」→「② 条件付き書式設定」
  - アポ赤・資料水色・NG系グレーの色分けが全コールシートに適用される
  - 完了アラートで設定シート数を確認する

STEP 5: ③ コール状況 補填 を実行（既存データがある場合のみ・一回きり）
  - メニュー「📋 ToGツール」→「③【一回きり】コール状況 補填」
  - コメント列を読んでコール状況列が空のセルに自動推定値を書き込む
  - 完了アラートで設定件数を確認する
  - ※ 推定精度を確認し、誤りがあれば手動で修正してください

STEP 6: ④ コールログ 収集 を実行（既存データがある場合のみ・一回きり）
  - メニュー「📋 ToGツール」→「④【一回きり】コールログ 収集」
  - 全コールシートの既存データから担当・日付・時間入り行を収集し
    「コールログ」シートに一覧化する
  - 完了アラートで収集件数を確認する

■ 日常運用（以降は自動）

- onEdit トリガーを設定する（GAS エディタ → 「トリガー」→「トリガーを追加」）
  関数: onEdit / イベント: スプレッドシートから / 種別: 編集時
  ※ Simple trigger（function onEdit）は自動で動くが、権限が必要な操作
    （他シートへの書き込み等）には Installable trigger が必要。
    「トリガーを追加」から Installable trigger として登録することを推奨。

- コール状況列（H/M/R/W）に値を選択すると：
  →「資料メール」なら資料送付リストへ自動転記（重複はスキップ）
  → コールログに自動追記

- 担当・日付・時間を入力すると：
  → コールログに自動追記

■ 再設定が必要なとき

- シートを追加した場合: ①②を再実行（新シートにも適用される）
- プルダウン選択肢を変更したい場合: スクリプト冒頭の DROPDOWN_VALUES を編集して①を再実行
- 色を変えたい場合: setup_conditional() 内の setBackground() の値を変更して②を再実行

■ 注意事項

- EXCLUDE_SHEETS に含まれるシートは一切処理しない
- backfill_callStatus / backfill_callLog は「一回きり」実行が前提。
  再実行すると backfill_callLog はログが作り直しになる（既存ログは消える）。
  backfill_callStatus は「空セルのみ上書き」なので再実行しても安全。
- 資料送付リストシート名の変更時は SHEET_MATERIAL 定数を更新すること
- データ開始行（DATA_START_ROW）は 2 固定（1行目=ヘッダー）
=======================================================================
*/
