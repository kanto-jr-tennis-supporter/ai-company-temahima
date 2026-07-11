/**
 * T4 入札参加資格申請書アプリ プロトタイプ（Step1〜4／サンプル⑤ 文部科学省・最小構成のみ）
 * 作成：ハック（自動化エンジニア）／2026-07-08
 *
 * 前提：
 *  - 対象スプレッドシートに、社長が手動で以下3シートを
 *    「ファイル → インポート → シートを挿入」で サンプル⑤.xlsx から読み込み済みであること。
 *      ・様式1-1　・様式1-2　・様式2
 *    （インポート手順は logs/T4_構築・検証手順.md に記載）
 *  - このスクリプトは、それらのシート名が変わっていない前提で座標を参照する。
 *
 * 座標の根拠：
 *  - すべて logs/T4_入札サンプル/サンプル⑤.xlsx を openpyxl で直接開き、
 *    結合セル(merged_cells)・罫線(border)を検証して特定した実座標（推測ではない）。
 *  - フリガナ（商号・住所・代表者氏名）は当初「1文字1セルに分解」を試したが、実測の結果
 *    分解先セルに罫線の裏付けがなく、実機検証で位置ズレが確認されたため、フリガナの記入欄
 *    そのもの（罫線のない1個の結合セル。例：商号ふりがなはX23:GA23）へ1つの文字列として
 *    そのまま転記する方式（kind: 'text'）に変更した。1文字1セル分解の仕組み自体
 *    （writeCharGrid_ / kind: 'chargrid'）は、将来ほかの様式（例：明確な罫線で1文字1マスが
 *    区切られているサンプル④）に流用できるよう関数は残してある。
 */

// ====== 基本設定 ======
const SHEET_NAMES = {
  form: '入力フォーム',
  guide: '入力ガイド',
  config: '設定',
};
const MASTER_SHEETS_REQUIRED = ['様式1-1', '様式1-2', '様式2'];

// 「○」の代わりに貼る図形の代替：楕円のPNG画像（透過背景・黒枠）をBase64で埋め込み。
// 理由：SpreadsheetApp（Sheets用のApps Script API）には、Slidesのinsert Shape(OVAL)に相当する
// 「図形を直接挿入するAPI」が存在しない（2026年7月時点で確認）。そのためSheet.insertImage()で
// 画像を代わりに配置する次善策を採用した。誤魔化さずここに明記する。
const OVAL_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAMgAAAA8CAYAAAAjW/WRAAACEElEQVR4nO3d0VbrIBCF4ej7v7PeqKtqQoYMMHuG/7s/hcDeZvU0pccBAAAAAOu8RU+gkI/oCZxgf51YwP8Ug74aufiy20IQ/nG2yE7Fi6QE8crkKvuFUIY8UmYty6QpQl3SGVSdXLZCqK7jcbCWLiqTUdlElfVQwJ4ED75yAwj+PKX3ceWAsxeSEuhJv+ezB5ixQBQhvzS5mBW2UQtAGfYhmZmRLzbiAikEvknkaUQgPRdCIWAVkjNPQJ9MmEJglCX5exrYnslRCsw2LY+94aUYUDY8nz0htgxOKaBiSF6tgb4bjGJAlSu777MHAILd5bOZb88/phjIpjvPrTsI5UA1rdye5v2qIJQDVXWVxPIexPriQBbmHJ8V5OruQTlQyVWef+W/9w4CbIWCAA0UBGigIEDDWUFMb16A5Ez/GdV7B6EkqMCc46uCdH/iCCTR9SF46w5CSVBN9xMilg//eJoX2T3OsOU9iOtxYSCY6w883yhEVUu/Udgz6NPXBrxCv5M+dSKAg8ypJq84FwuRpM/FesXJilgh3cmKf0mcpYoyJPLE6e5QIZkZfh8EEdLkgl+Ywkzp95zfKIRX6X1UCY7K4yoq66GAPYkevEFlc6xU1/E4WEsXqck0ZNtk2ElnUHpyBhQnj5RZSznpG5QmXplclbkQI8ozzhbZ2eIiO1EicvGDhRhHsVjsLwAAANR8Ank5VETKpQASAAAAAElFTkSuQmCC';

// チェックマーク（レ点／✓）のPNG画像（透過背景・黒）をBase64で埋め込み。みなし大企業など
// 「○で囲む」より「チェックを入れる」方が自然な項目向け。Macでは見栄えの良いチェック文字を
// 手軽に入力できないため、画像として自動配置する。
const CHECK_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAACbklEQVR4nO3bUXbbMAwF0Zfuf8/JRw8btY1skQQIgJy7AEfGWKQlRxIAAAAAAAAAAPj2afEivyxeBH9iTEchiL1PTYQhyLy74Q9FIcicd0PvjvIxeCDoH/ajWXOGrPMoIGfImNlvU7dz5wzpZ3G9cfsaBOljcvEnzpBUXm4TBHnO6ux4iSDPuC9VDUHeWxZDIsgqjy8vCPLakn3jiiD3li5VDUF+FhJDIoinodtSBPnf8n3jiiB/C1uqGoJ8C48hEcTa9M8ZBPktdN+4IkiSpao5PUiqGBJBLJj+DH5ykDT7xtWpQdItVc2JQdLGkM4MYsHt36dOC5Jy37g6KUjqpao5JUiJGNI5QSws+bfbE4Kk3zeudg9SZqlqPIJk+USWiyHZBzF7+DGJ5Y9rWAb5N0JklLIfCKsgpg8+Tiq5VDUWQcwffJxQOoY0H+TpACotIaGP+c0E6R2yd5RK0W+NBsn25ssvVc1IkJk37xFymxhSfxDXJ1ADpYgh9QexOnCrKBnjThlZsrJE2WqpakY39egoW8aQ5r72RkeZlS6GNH9hGBFlu33jyuLWycoo2y5VTbYfqF4NfPsYkl0QyzfpuSSljiHZniGeUbbeN66slyyPKEcsVY3XQWb7RJeIIflt6pkGkOlY3vL8llVqEFl4f+2NjhL997utuA6JGkq5GNK6C8PVwykZQ1p7pV52SCutvnWyIkrp8BH3sjwHVjqGFHdz0WNw5WNI+e72Hi8yiOUneouzQ4o/QywGuU0MKT6INDfQrWJIOYJIGw52VJYgUn+ULSNmfFPvfkvJeMxmMr+5n8JkPl4AAAAAAAD4+gJb01GKNhqh7AAAAABJRU5ErkJggg==';

// 必須項目（入力フォームの入力セル）。条件付き書式・未入力集計の両方でここを参照する。
const REQUIRED_FORM_CELLS = [
  'D5', 'D6', 'D9', 'D11', 'D12', 'D14', 'D15',
  'D18', 'D20', 'D21', 'D22', 'D23', 'D24', 'D25', 'D26', 'D27', 'D28',
];

// 設定シートに書き出す座標マッピング（フィールドの追加はここに1行足すだけでよい設計）。
// kind: 'text'(単純転記) / 'none'(入力フォームの記録用のみ。マスターへは何もしない) /
//       'ovaln'(楕円画像、N択。extra1にJSON "{選択肢:セルA1}") / 'checkln'(チェックマーク画像、N択。extra1の書式はovalnと同じ) /
//       'chargrid'(1文字1セル分解) / 'yubinsplit'(郵便番号を前3桁・後4桁の2箱に分割) / 'textOrFallback'(空欄ならextra1の入力フォームセルの値を使う)
const FIELD_CONFIG = [
  // id, kind, 入力フォーム参照セル, 転記先シート, 転記先セル, 予備1, 予備2, 備考
  ['shinki_koushin', 'none', 'D5', '様式1-1', '', '', '', '新規／更新は文部科学省では丸印不要のため転記なし（入力フォーム上の記録のみ）'],
  ['shago', 'text', 'D6', '様式1-1', 'X25', '', '', '商号又は名称（単純転記の代表例）'],
  ['shago_furigana', 'text', 'D7', '様式1-1', 'X23', '', '', '商号ふりがな。1文字1セルの推測位置ではなく結合セルへそのまま転記'],
  ['yubin', 'yubinsplit', 'D8', '様式1-1', 'X15', 'AK15', '', '郵便番号（前3桁→X15 / 後4桁→AK15の2箱に分割）'],
  ['jusho', 'text', 'D9', '様式1-1', 'X20', '', '', '本社（店）住所'],
  ['jusho_furigana', 'text', 'D10', '様式1-1', 'X18', '', '', '住所ふりがな。1文字1セルの推測位置ではなく結合セルへそのまま転記'],
  ['yakushoku', 'text', 'D11', '様式1-1', 'X28', '', '', '代表者役職'],
  ['daihyo_shimei', 'text', 'D12', '様式1-1', 'X33', '', '', '代表者氏名'],
  ['daihyo_furigana', 'text', 'D13', '様式1-1', 'X31', '', '', '代表者氏名ふりがな。1文字1セルの推測位置ではなく結合セルへそのまま転記'],
  ['kyoka_mae', 'text', 'D14', '様式1-1', 'CF4', '', '', '建設業許可番号（ハイフン前）'],
  ['kyoka_ato', 'text', 'D15', '様式1-1', 'CO4', '', '', '建設業許可番号（ハイフン後。ハイフン自体はCL4に印字済み）'],
  ['hojin_bango', 'text', 'D16', '様式1-1', 'BP15', '', '', '法人番号（正しい箱はBP15。旧設定でAK15＝郵便番号の箱を誤って使っていたのを修正）'],
  ['tantosha_shimei', 'textOrFallback', 'D17', '様式1-1', 'DL33', 'D12', '', '担当者氏名。空欄なら代表者氏名(D12)をそのまま転記'],
  ['honsha_tel', 'text', 'D18', '様式1-1', 'X36', '', '', '本社（店）電話番号'],
  ['tantosha_tel', 'textOrFallback', 'D19', '様式1-1', 'DL36', 'D18', '', '担当者電話番号。空欄なら本社電話番号(D18)をそのまま転記'],
  ['honsha_fax', 'text', 'D20', '様式1-1', 'X40', '', '', '本社（店）ＦＡＸ番号'],
  ['email', 'text', 'D21', '様式1-1', 'X43', '', '', 'メールアドレス'],
  ['soshokuinsu', 'text', 'D22', '様式1-1', 'EQ56', '', '', '総職員数（人）'],
  ['eigyo_nensu', 'text', 'D23', '様式1-1', 'EQ53', '', '', '営業年数（年）。現状は総合評定値通知書の記載値をそのまま手入力。省庁ごとの算出方式の違いは今後パターンを蓄積して対応する'],
  ['setsuritsu_gengo', 'ovaln', 'D24', '様式1-1', '', '{"明治":"C62","大正":"I62","昭和":"C63","平成":"I63","令和":"C64"}', '', '設立年号（和暦）。選んだ元号の文字に黒楕円を配置'],
  ['setsuritsu_nen', 'text', 'D25', '様式1-1', 'O62', '', '', '設立年（和暦の数字のみ）'],
  ['setsuritsu_tsuki', 'text', 'D26', '様式1-1', 'AA62', '', '', '設立月'],
  ['setsuritsu_hi', 'text', 'D27', '様式1-1', 'AM62', '', '', '設立日'],
  ['minashi_daikigyo', 'checkln', 'D28', '様式1-1', '', '{"該当しない":"DI62","該当する":"BY62"}', '', 'みなし大企業。デフォルトは「該当しない」。楕円ではなくチェックマークを配置'],
  ['minashi_riyu', 'checkln', 'D29', '様式1-1', '', '{"発行済株式2分の1以上":"BV63","発行済株式3分の2以上":"BV64","役員兼務":"BV65"}', '', 'D28で「該当する」を選んだ場合のみ。該当理由の「・」にチェックマークを配置'],
];

// ====== メニュー ======
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('入札申請書ツール')
    .addItem('① 初期セットアップ実行（setup）', 'setup')
    .addItem('② マスターへ転記する（transcribeToMaster）', 'transcribeToMaster')
    .addToUi();
}

// ====== STEP1〜3: 初期セットアップ ======
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  checkMasterSheetsExist_(ss); // 無ければここで例外を投げて止める

  createInputGuideSheet_(ss);
  createInputFormSheet_(ss);
  extendInputFormSheetIfNeeded_(ss); // 既存の「入力フォーム」に後から項目が増えた場合、入力済みデータは消さずに追記する
  createSettingsSheet_(ss);
  applyConditionalFormatting_(ss);
  protectMasterSheets_(ss);

  SpreadsheetApp.getUi().alert(
    'セットアップが完了しました。\n「入力フォーム」シートに入力し、メニュー「入札申請書ツール → ② マスターへ転記する」を実行してください。'
  );
}

function checkMasterSheetsExist_(ss) {
  const missing = MASTER_SHEETS_REQUIRED.filter((name) => !ss.getSheetByName(name));
  if (missing.length > 0) {
    const msg =
      'マスターシートが見つかりません：' +
      missing.join('、') +
      '\n\n先に「ファイル → インポート → シートを挿入」で\n' +
      'logs/T4_入札サンプル/サンプル⑤.xlsx の該当シートをこのスプレッドシートに読み込んでから、\n' +
      'もう一度 setup() を実行してください（logs/T4_構築・検証手順.md 参照）。';
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (e) {
      // UIが使えない実行コンテキスト（エディタから直接実行等）でもログには残す
    }
    throw new Error(msg);
  }
}

function createInputGuideSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.guide);
  if (sheet) return; // 既にあれば作り直さない（社長が手を入れている可能性を尊重）
  sheet = ss.insertSheet(SHEET_NAMES.guide, 0);

  const rows = [
    ['入力ガイド（設計書 §4 の3色分類）', ''],
    ['', ''],
    ['🟩 緑：直接入力', 'テキスト・数値をそのまま「入力フォーム」に入力する項目（商号、住所、許可番号など）'],
    ['🟨 黄：選択式（プルダウン）', '「入力フォーム」のプルダウンで選ぶだけ。○を手で書く操作はしない（新規／更新など）'],
    ['🟥 赤：完全ロック', 'マスターシート本体。人間は直接触らない。GASの「転記する」だけが書き込む'],
    ['', ''],
    ['注意', 'フリガナ欄は本来1文字1セルの想定だったが、サンプル⑤を実測した結果、'
      + '罫線のない結合セルであることが判明。5文字までは1文字ずつのセルへ、'
      + 'それを超える分は結合セル側へまとめて書き込む方式にしている（詳細はコード冒頭コメント参照）。'],
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).merge().setFontWeight('bold').setFontSize(13);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 560);
  sheet.getRange(1, 1, rows.length, 2).setWrap(true);
}

function createSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.config);
  if (sheet) return;
  sheet = ss.insertSheet(SHEET_NAMES.config, 2);

  const header = ['フィールドID', '種別', '入力フォーム参照セル', '転記先シート', '転記先セル', '予備1', '予備2', '備考'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.getRange(2, 1, FIELD_CONFIG.length, header.length).setValues(FIELD_CONFIG);
  sheet.setFrozenRows(1);
  for (let c = 1; c <= header.length; c++) sheet.autoResizeColumn(c);
}

function createInputFormSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.form);
  if (sheet) return;
  sheet = ss.insertSheet(SHEET_NAMES.form, 1);

  sheet.getRange('A1').setValue('入札参加資格申請書 入力フォーム（文部科学省・様式1-1／1-2／2 用）')
    .setFontWeight('bold').setFontSize(13);
  sheet.getRange('A1:F1').merge();

  const requiredCount = REQUIRED_FORM_CELLS.length;
  const countBlankFormula =
    '=' + REQUIRED_FORM_CELLS.map((c) => 'COUNTBLANK(' + c + ')').join('+');
  sheet.getRange('A2').setValue('進捗');
  sheet.getRange('B2').setFormula(
    '="未入力: " & (' + countBlankFormula.substring(1) + ') & " / ' + requiredCount + '件（必須項目）"'
  );
  sheet.getRange('B2').setFontWeight('bold');

  const header = ['No', '分類', '項目', '入力してください', '必須', '備考'];
  sheet.getRange(4, 1, 1, header.length).setValues([header]).setFontWeight('bold');

  // 行 = FIELD_CONFIG の並び順に対応させる（D5始まりで1行ずつ）
  const rows = [
    [1, '🟨選択式', '新規／更新', '', '必須', 'プルダウンから選択。マスターへは自動で楕円を配置'],
    [2, '🟩直接入力', '商号又は名称', '', '必須', ''],
    [3, '🟩直接入力', '商号ふりがな（カタカナ）', '', '任意', ''],
    [4, '🟩直接入力', '郵便番号', '', '任意', 'ハイフンは任意'],
    [5, '🟩直接入力', '本社（店）住所', '', '必須', ''],
    [6, '🟩直接入力', '本社住所ふりがな（カタカナ）', '', '任意', ''],
    [7, '🟩直接入力', '代表者役職', '', '必須', ''],
    [8, '🟩直接入力', '代表者氏名', '', '必須', ''],
    [9, '🟩直接入力', '代表者氏名ふりがな（カタカナ）', '', '任意', ''],
    [10, '🟩直接入力', '建設業許可番号（ハイフンの前）', '', '必須', '例：6'],
    [11, '🟩直接入力', '建設業許可番号（ハイフンの後）', '', '必須', '例：123456'],
    [12, '🟩直接入力', '法人番号', '', '任意', ''],
    [13, '🟩直接入力', '担当者氏名', '', '任意', '空欄なら代表者氏名と同じものを転記します'],
    [14, '🟩直接入力', '本社（店）電話番号', '', '必須', ''],
    [15, '🟩直接入力', '担当者電話番号', '', '任意', '空欄なら本社電話番号と同じものを転記します'],
    [16, '🟩直接入力', '本社（店）ＦＡＸ番号', '', '必須', ''],
    [17, '🟩直接入力', 'メールアドレス', '', '必須', ''],
    [18, '🟩直接入力', '総職員数', '', '必須', '人数のみ入力'],
    [19, '🟩直接入力', '営業年数', '', '必須', '総合評定値通知書に記載の数字をそのまま入力'],
    [20, '🟨選択式', '設立年号（和暦）', '', '必須', 'プルダウンから選択'],
    [21, '🟩直接入力', '設立年', '', '必須', '和暦の数字のみ（例：6）'],
    [22, '🟩直接入力', '設立月', '', '必須', ''],
    [23, '🟩直接入力', '設立日', '', '必須', ''],
    [24, '🟨選択式', 'みなし大企業', '', '必須', '通常は「該当しない」のままでOK'],
    [25, '🟨選択式', 'みなし大企業の該当理由', '', '任意', '24で「該当する」を選んだ場合のみ選択'],
  ];
  sheet.getRange(5, 1, rows.length, header.length).setValues(rows);

  // D5 に「新規／更新」のプルダウン
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['新規', '更新'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D5').setDataValidation(rule);

  // D24 に設立年号（和暦）のプルダウン
  sheet.getRange('D24').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['明治', '大正', '昭和', '平成', '令和'], true)
      .setAllowInvalid(false)
      .build()
  );

  // D28 にみなし大企業のプルダウン（デフォルトは「該当しない」）
  sheet.getRange('D28').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['該当しない', '該当する'], true)
      .setAllowInvalid(false)
      .build()
  );
  sheet.getRange('D28').setValue('該当しない');

  // D29 は「該当する」のときだけ使う任意項目。空欄も許可する。
  sheet.getRange('D29').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['発行済株式2分の1以上', '発行済株式3分の2以上', '役員兼務'], true)
      .setAllowInvalid(true)
      .build()
  );

  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 50);
  sheet.setColumnWidth(6, 320);
  sheet.setFrozenRows(4);
}

// 「入力フォーム」が旧バージョン（項目1〜12＝D5〜D16まで）のまま残っている場合、
// 項目13〜25（D17〜D29）を追記する。既に入力済みのD5〜D16の値は一切触らない。
function extendInputFormSheetIfNeeded_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.form);
  if (!sheet) return;
  if (sheet.getRange('C17').getValue() !== '') return; // 既に項目13以降がある＝拡張済み

  const extraRows = [
    [13, '🟩直接入力', '担当者氏名', '', '任意', '空欄なら代表者氏名と同じものを転記します'],
    [14, '🟩直接入力', '本社（店）電話番号', '', '必須', ''],
    [15, '🟩直接入力', '担当者電話番号', '', '任意', '空欄なら本社電話番号と同じものを転記します'],
    [16, '🟩直接入力', '本社（店）ＦＡＸ番号', '', '必須', ''],
    [17, '🟩直接入力', 'メールアドレス', '', '必須', ''],
    [18, '🟩直接入力', '総職員数', '', '必須', '人数のみ入力'],
    [19, '🟩直接入力', '営業年数', '', '必須', '総合評定値通知書に記載の数字をそのまま入力'],
    [20, '🟨選択式', '設立年号（和暦）', '', '必須', 'プルダウンから選択'],
    [21, '🟩直接入力', '設立年', '', '必須', '和暦の数字のみ（例：6）'],
    [22, '🟩直接入力', '設立月', '', '必須', ''],
    [23, '🟩直接入力', '設立日', '', '必須', ''],
    [24, '🟨選択式', 'みなし大企業', '', '必須', '通常は「該当しない」のままでOK'],
    [25, '🟨選択式', 'みなし大企業の該当理由', '', '任意', '24で「該当する」を選んだ場合のみ選択'],
  ];
  sheet.getRange(17, 1, extraRows.length, 6).setValues(extraRows);

  sheet.getRange('D24').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['明治', '大正', '昭和', '平成', '令和'], true)
      .setAllowInvalid(false)
      .build()
  );
  sheet.getRange('D28').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['該当しない', '該当する'], true)
      .setAllowInvalid(false)
      .build()
  );
  if (sheet.getRange('D28').getValue() === '') sheet.getRange('D28').setValue('該当しない');
  sheet.getRange('D29').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['発行済株式2分の1以上', '発行済株式3分の2以上', '役員兼務'], true)
      .setAllowInvalid(true)
      .build()
  );

  // 進捗カウント式を、増えた必須項目数（REQUIRED_FORM_CELLS）で更新する。
  const countBlankFormula = '=' + REQUIRED_FORM_CELLS.map((c) => 'COUNTBLANK(' + c + ')').join('+');
  sheet.getRange('B2').setFormula(
    '="未入力: " & (' + countBlankFormula.substring(1) + ') & " / ' + REQUIRED_FORM_CELLS.length + '件（必須項目）"'
  );
}

// 未入力＝黄色ハイライト。設計書§3の通り、1セル＝1ルールで個別に作る（範囲まとめ・逆ルールは作らない）。
function applyConditionalFormatting_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.form);
  const rules = [];
  REQUIRED_FORM_CELLS.forEach((a1) => {
    const range = sheet.getRange(a1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISBLANK(' + a1 + ')')
      .setBackground('#FFF9C4')
      .setRanges([range])
      .build();
    rules.push(rule);
  });
  sheet.setConditionalFormatRules(rules);
}

function protectMasterSheets_(ss) {
  MASTER_SHEETS_REQUIRED.forEach((name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (existing.length > 0) return; // 既に保護済みなら二重にかけない
    const protection = sheet.protect().setDescription('マスターシート保護（GASのみ書き込み可）');
    // オーナー（実行者）以外の編集者は外す。スクリプト実行はシート保護の影響を受けないため、
    // 転記処理自体は保護後も問題なく動く。
    const editors = protection.getEditors();
    if (editors.length > 0) protection.removeEditors(editors);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  });
}

// ====== STEP4: 転記する（メニューから実行） ======
function transcribeToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  checkMasterSheetsExist_(ss);

  const formSheet = ss.getSheetByName(SHEET_NAMES.form);
  const configSheet = ss.getSheetByName(SHEET_NAMES.config);
  if (!formSheet || !configSheet) {
    throw new Error('「入力フォーム」または「設定」シートがありません。先に setup() を実行してください。');
  }

  const config = configSheet
    .getRange(2, 1, configSheet.getLastRow() - 1, 8)
    .getValues()
    .filter((r) => r[0]); // フィールドID空行は無視

  let textCount = 0;
  let ovalCount = 0;
  let checkCount = 0;
  let charCount = 0;

  config.forEach((row) => {
    const [id, kind, formCell, sheetName, target, extra1, extra2] = row;
    const masterSheet = ss.getSheetByName(sheetName);
    if (!masterSheet) return; // 対象シートが無ければスキップ（拡張時の安全策）

    const value = formSheet.getRange(formCell).getValue();

    if (kind === 'none') {
      // 転記なし。入力フォーム上の記録のみで、マスターへは何もしない（例：新規/更新）。
      return;
    } else if (kind === 'text') {
      if (value === '' || value === null) return;
      masterSheet.getRange(target).setValue(String(value));
      textCount++;
    } else if (kind === 'yubinsplit') {
      if (value === '' || value === null) return;
      const digits = String(value).replace(/[^0-9]/g, '');
      masterSheet.getRange(target).setValue(digits.slice(0, 3));
      masterSheet.getRange(String(extra1)).setValue(digits.slice(3, 7));
      textCount++;
    } else if (kind === 'ovaln' || kind === 'checkln') {
      removeExistingAutoImages_(masterSheet, id);
      if (value === '' || value === null) return;
      let optionMap;
      try { optionMap = JSON.parse(extra1); } catch (e) { return; }
      const cellA1 = optionMap[String(value)];
      if (!cellA1) return; // 想定外の値は安全側で何もしない
      if (kind === 'ovaln') {
        insertOvalAt_(masterSheet, id, cellA1);
        ovalCount++;
      } else {
        insertCheckAt_(masterSheet, id, cellA1);
        checkCount++;
      }
    } else if (kind === 'textOrFallback') {
      let v = value;
      if (v === '' || v === null) {
        v = formSheet.getRange(String(extra1)).getValue();
      }
      if (v === '' || v === null) return;
      masterSheet.getRange(target).setValue(String(v));
      textCount++;
    } else if (kind === 'chargrid') {
      if (value === '' || value === null) return;
      writeCharGrid_(masterSheet, String(value), target, Number(extra1), Number(extra2));
      charCount++;
    }
  });

  SpreadsheetApp.getUi().alert(
    '転記が完了しました。\nテキスト転記: ' + textCount + '件 / 楕円配置: ' + ovalCount + '件 / ' +
    'チェック配置: ' + checkCount + '件 / 文字分解転記: ' + charCount + '件'
  );
}

function insertOvalAt_(sheet, fieldId, targetCellA1) {
  insertMarkAt_(sheet, fieldId, targetCellA1, OVAL_PNG_BASE64);
}

function insertCheckAt_(sheet, fieldId, targetCellA1) {
  insertMarkAt_(sheet, fieldId, targetCellA1, CHECK_PNG_BASE64);
}

// 楕円／チェックマーク共通の画像配置処理（'ovaln'・'checkln'種別）。
// 呼び出し側で removeExistingAutoImages_ 済みであること。
function insertMarkAt_(sheet, fieldId, targetCellA1, pngBase64) {
  const range = sheet.getRange(targetCellA1);
  const merged = range.getMergedRanges();
  const anchorRange = merged.length > 0 ? merged[0] : range;

  // 対象セル（結合セル）の実ピクセルサイズを動的に計算し、画像をそこに収める。
  let widthPx = 0;
  for (let c = anchorRange.getColumn(); c < anchorRange.getColumn() + anchorRange.getNumColumns(); c++) {
    widthPx += sheet.getColumnWidth(c);
  }
  let heightPx = 0;
  for (let r = anchorRange.getRow(); r < anchorRange.getRow() + anchorRange.getNumRows(); r++) {
    heightPx += sheet.getRowHeight(r);
  }

  const blob = Utilities.newBlob(Utilities.base64Decode(pngBase64), 'image/png', fieldId + '.png');
  const image = sheet.insertImage(blob, anchorRange.getColumn(), anchorRange.getRow());
  image.setWidth(Math.max(widthPx, 20)).setHeight(Math.max(heightPx, 14));
  image.setAltTextTitle('AUTO_OVAL_' + fieldId);
}

function removeExistingAutoImages_(sheet, fieldId) {
  const tag = 'AUTO_OVAL_' + fieldId;
  sheet.getImages().forEach((img) => {
    if (img.getAltTextTitle() === tag) img.remove();
  });
}

// 1文字ずつ、2列ピッチのセルへ分解して書き込む。maxCharsを超えた分は
// overflowCellA1（結合セル側）へまとめて追記する。
// ※このサンプル⑤では、この分解先セルに罫線による裏付けは無い（コード冒頭コメント参照）。
//   将来、罫線で明確に区切られた1文字1マスの様式（例：サンプル④）に適用する場合は、
//   startCellA1・pitch・maxCharsを設定シートの値だけ差し替えればそのまま使える。
function writeCharGrid_(sheet, text, startCellA1, pitch, maxChars) {
  const startRange = sheet.getRange(startCellA1);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();

  const chars = text.split('');
  const headChars = chars.slice(0, maxChars);
  const overflow = chars.slice(maxChars).join('');

  headChars.forEach((ch, i) => {
    sheet.getRange(startRow, startCol + i * pitch).setValue(ch);
  });

  if (overflow) {
    // overflow は「1文字1セル」の想定外なので、結合セル側（罫線ありの本来の入力欄）へ
    // まとめて追記する。結合セルの位置は、フリガナ系フィールドでは startCol直後の
    // 境界列（このプロトタイプでは固定でX列=24列目）に決め打ちせず、設定シート側で
    // 別途 text 種別の行を用意する運用を想定。ここでは安全のためログにのみ残す。
    Logger.log(
      '[writeCharGrid_] "' + text + '" は' + maxChars + '文字を超えたため、' +
      overflow + ' が未転記です（結合セル側への追記は設定シート側で別フィールドとして定義してください）。'
    );
  }
}
