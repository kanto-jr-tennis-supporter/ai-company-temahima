/**
 * T5：部活動 出欠管理（2026年度・個人シート方式）
 * 作成：ハック（自動化エンジニア）
 *
 * 設計の考え方：
 *  - 生徒は「個人シート」に日付・状態・理由を書くだけ（鍵なし・スクリプト不要）。
 *  - 「出欠一覧」「出席率」「今日の一覧」は全部スプレッドシートの数式（QUERY/FILTER/INDIRECT等）で
 *    自動集計する。GASを使うのは最初の1回だけ（名簿から個人シートを作る作業）。
 *  - 「予定」シートに活動日を入れると、「出欠一覧」の日付・予定列に自動で反映される
 *    （入力順がバラバラでも日付順に並ぶようSORTしている）。
 *
 * 前提：
 *  - このスプレッドシートに「名簿」シートを作り、1行目を見出し（例：年組／番号／担任／氏名）、
 *    2行目から今年度の名簿を貼り付けておくこと（列の並びは A=年組 B=番号 C=担任 D=氏名 を想定）。
 *  - 「年組」列は「1年5組」のような文字列でOK（学年・クラスへの分解はこのスクリプト側の数式で行う）。
 */

const SHEET_NAMES = {
  roster: '名簿',
  schedule: '予定',
  overview: '出欠一覧',
  rate: '出席率',
  today: '今日の一覧',
  template: 'テンプレート',
};

const STATUS_OPTIONS = ['出席', '欠席', '遅刻', '早退', '遅早'];

// 「予定」の想定最大件数。年間の活動日数がこれを超える場合は、手順書の通り
// 「出欠一覧」の一番右の日付ペアの数式列を選択してそのまま右へコピーすれば拡張できる。
const MAX_DATE_PAIRS = 60;

// ====== メニュー ======
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('出欠管理')
    .addItem('① 初期セットアップ実行（setup）', 'setup')
    .addItem('② 個人シートを作り直す（新入部員が増えたとき）', 'addNewStudentSheets')
    .addToUi();
}

// ====== メイン ======
function setup() {
  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(SHEET_NAMES.roster);
  if (!roster || roster.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(
      '「名簿」シートが無いか、名簿が空です。\n' +
      '「名簿」シートを作り、1行目を見出し、2行目から今年度の名簿（年組・番号・担任・氏名）を貼ってから、' +
      'もう一度実行してください。'
    );
    return;
  }
  const studentCount = roster.getLastRow() - 1;

  createScheduleSheetIfNeeded_(ss);
  rebuildTodaySheet_(ss, studentCount);
  createTemplateSheetIfNeeded_(ss);
  rebuildOverviewSheet_(ss, studentCount);
  rebuildRateSheet_(ss, studentCount);
  createStudentSheets_(ss, roster, studentCount);

  SpreadsheetApp.getUi().alert(
    'セットアップ（再）実行が完了しました！\n\n' +
    '・「名簿」の人数分だけ個人シートができています（既存の個人シートの入力内容はそのままです）。\n' +
    '・「予定」シートに活動日と予定内容を入れると、「出欠一覧」に自動で反映されます。「休み」と書いた日は集計対象から外れます。\n' +
    '・生徒には、欠席・遅刻・早退・遅早のときだけ個人シートに記録してもらってください。何も記録しなければ「出席」として扱われます。\n' +
    '・「出欠一覧」「出席率」「今日の一覧」は毎回作り直されるので、ロジックを直したときは setup を再実行するだけで反映されます。'
  );
}

// 名簿に人数が増えたときだけ再実行する軽い版（個人シート追加のみ）
function addNewStudentSheets() {
  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(SHEET_NAMES.roster);
  if (!roster || roster.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('「名簿」シートが見つからないか、名簿が空です。');
    return;
  }
  const studentCount = roster.getLastRow() - 1;
  createStudentSheets_(ss, roster, studentCount);
  SpreadsheetApp.getUi().alert('個人シートを確認・追加しました（既存のシートはそのままです）。');
}

// ====== 各シート作成 ======

function createScheduleSheetIfNeeded_(ss) {
  if (ss.getSheetByName(SHEET_NAMES.schedule)) return;
  const sheet = ss.insertSheet(SHEET_NAMES.schedule, 0);
  sheet.getRange('A1:B1').setValues([['日付', '予定']]).setFontWeight('bold');
  sheet.getRange('A2').setValue('（ここに活動日と予定内容を追加していってください。順番はバラバラでもOK）');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 220);
  sheet.setFrozenRows(1);
}

// 「今日の一覧」は全部数式（ユーザーが直接データを書く場所ではない）ので、setup()を
// 再実行するたびに作り直して、常に最新のロジックにする（名簿・予定・個人シートは触らない）。
function rebuildTodaySheet_(ss, studentCount) {
  const old = ss.getSheetByName(SHEET_NAMES.today);
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(SHEET_NAMES.today, 1);

  sheet.getRange('A1:E1').setValues([STATUS_OPTIONS]).setFontWeight('bold');

  const lastRow = studentCount + 1;
  // H列=氏名、I列=今日の状態（各個人シートをTODAY()で引く）。表示上は補助列。
  // ・今日が「予定」で「休み」の日は、誰も表示しない（(休み)扱いにして5区分どれにも一致させない）
  // ・今日が活動日で、本人シートに記録が無ければ「出席」とみなす
  // ・今日がそもそも「予定」に無い日は、何も表示しない
  sheet.getRange('H1').setValue('（以下、補助列。消さないでください）');
  const namesFormulas = [];
  const statusFormulas = [];
  for (let i = 0; i < studentCount; i++) {
    const rosterRow = i + 2; // 名簿の実データ行
    const helperRow = i + 2; // 今日の一覧の補助列の行
    namesFormulas.push(['=IFERROR(名簿!D' + rosterRow + ',"")']);
    statusFormulas.push([
      '=IFERROR(IF(VLOOKUP(TODAY(),' + SHEET_NAMES.schedule + '!$A$2:$B,2,FALSE)="休み","(休み)",' +
      'IFERROR(VLOOKUP(TODAY(), INDIRECT("\'" & H' + helperRow + ' & "\'!A:C"), 2, FALSE), "出席")), "")',
    ]);
  }
  if (studentCount > 0) {
    sheet.getRange(2, 8, studentCount, 1).setFormulas(namesFormulas); // H2から
    sheet.getRange(2, 9, studentCount, 1).setFormulas(statusFormulas); // I2から
  }

  // A2〜E2 に各区分のFILTER（該当者が0人でも空欄になるだけでエラーにならないようIFERRORでくるむ）
  const lastHelperRow = Math.max(lastRow, 2);
  const filterFormulas = STATUS_OPTIONS.map((label) => {
    return '=IFERROR(FILTER($H$2:$H$' + lastHelperRow + ', $I$2:$I$' + lastHelperRow + '="' + label + '"), "")';
  });
  sheet.getRange(2, 1, 1, STATUS_OPTIONS.length).setFormulas([filterFormulas]);

  sheet.setColumnWidths(1, STATUS_OPTIONS.length, 110);
  sheet.setFrozenRows(1);
}

function createTemplateSheetIfNeeded_(ss) {
  if (ss.getSheetByName(SHEET_NAMES.template)) return;
  const sheet = ss.insertSheet(SHEET_NAMES.template, 2);
  sheet.getRange('A1:C1').setValues([['日付', '状態', '理由']]).setFontWeight('bold');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 220);
  sheet.setFrozenRows(1);

  // B列（状態）は5区分の固定リストから選ぶ（未入力＝出席とみなすので、
  // 「出席」は基本選ばなくてよい。欠席・遅刻・早退・遅早のときだけ選ぶ運用）。
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B2:B1000').setDataValidation(rule);
}

// 「出欠一覧」も全部数式なので、setup()のたびに作り直して最新ロジックにする。
function rebuildOverviewSheet_(ss, studentCount) {
  const old = ss.getSheetByName(SHEET_NAMES.overview);
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(SHEET_NAMES.overview, 3);

  sheet.getRange('A1:D1').setValues([['学年', 'クラス', '名前', '担任']]);
  sheet.getRange('A1:D2').setFontWeight('bold');

  // 日付・予定・出欠・理由のヘッダー（2列1セット×MAX_DATE_PAIRS）
  const headerRow1 = [];
  const headerRow2 = [];
  for (let i = 1; i <= MAX_DATE_PAIRS; i++) {
    headerRow1.push('=IFERROR(INDEX(SORT(' + SHEET_NAMES.schedule + '!$A$2:$B,1,TRUE),' + i + ',1),"")');
    headerRow1.push('=IFERROR(INDEX(SORT(' + SHEET_NAMES.schedule + '!$A$2:$B,1,TRUE),' + i + ',2),"")');
    headerRow2.push('出欠');
    headerRow2.push('理由');
  }
  sheet.getRange(1, 5, 1, headerRow1.length).setFormulas([headerRow1]); // E1から
  sheet.getRange(2, 5, 1, headerRow2.length).setValues([headerRow2]); // E2から

  // 学年・クラス・名前・担任（名簿から自動参照）＋ 各日付ペアの出欠・理由（本人シートをVLOOKUP）。
  // ルール：
  //  ・その日の「予定」が「休み」なら、出欠・理由とも空欄にする（誰も対象にしない）
  //  ・予定が無い（未使用の列）場合も空欄にする
  //  ・本人シートにその日の記録が無ければ「出席」とみなす（欠席・遅刻・早退・遅早が明示的に
  //    記録されている日だけ、その内容を表示する）
  for (let i = 0; i < studentCount; i++) {
    const row = i + 3; // 出欠一覧の行（3行目から）
    const rosterRow = i + 2; // 名簿の行

    sheet.getRange(row, 1).setFormula('=IFERROR(REGEXEXTRACT(名簿!A' + rosterRow + ',"^(\\d+)年"),"")');
    sheet.getRange(row, 2).setFormula('=IFERROR(REGEXEXTRACT(名簿!A' + rosterRow + ',"年(.+)組"),"")');
    sheet.getRange(row, 3).setFormula('=IFERROR(名簿!D' + rosterRow + ',"")');
    sheet.getRange(row, 4).setFormula('=IFERROR(名簿!C' + rosterRow + ',"")');

    const rowFormulas = [];
    for (let i2 = 1; i2 <= MAX_DATE_PAIRS; i2++) {
      const dateCol = columnToLetter_(5 + (i2 - 1) * 2);
      const planCol = columnToLetter_(6 + (i2 - 1) * 2);
      rowFormulas.push(
        '=IF(OR(' + planCol + '$1="休み",' + dateCol + '$1=""),"",' +
        'IFERROR(VLOOKUP(' + dateCol + '$1, INDIRECT("\'" & $C' + row + ' & "\'!A:C"), 2, FALSE), "出席"))'
      );
      rowFormulas.push(
        '=IF(OR(' + planCol + '$1="休み",' + dateCol + '$1=""),"",' +
        'IFERROR(VLOOKUP(' + dateCol + '$1, INDIRECT("\'" & $C' + row + ' & "\'!A:C"), 3, FALSE), ""))'
      );
    }
    sheet.getRange(row, 5, 1, rowFormulas.length).setFormulas([rowFormulas]);
  }

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(4);
}

// 「出席率」も全部数式なので、setup()のたびに作り直して最新ロジックにする。
function rebuildRateSheet_(ss, studentCount) {
  const old = ss.getSheetByName(SHEET_NAMES.rate);
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(SHEET_NAMES.rate, 4);
  sheet.getRange('A1:D1').setValues([['氏名', '出席率', '欠席数', '対象日数']]).setFontWeight('bold');

  // 対象日数＝「予定」のうち日付があって「休み」でない件数（全員共通なので1回だけ組み立てる）
  const targetDaysFormula =
    '=COUNTIFS(' + SHEET_NAMES.schedule + '!$A$2:$A$1000,"<>",' +
    SHEET_NAMES.schedule + '!$B$2:$B$1000,"<>休み")';

  for (let i = 0; i < studentCount; i++) {
    const row = i + 2;
    const rosterRow = i + 2;
    sheet.getRange(row, 1).setFormula('=IFERROR(名簿!D' + rosterRow + ',"")');
    // 欠席数：本人シートで明示的に「欠席」を選んだ日数のみ（遅刻・早退・遅早は出席扱い）
    sheet.getRange(row, 3).setFormula(
      '=IFERROR(COUNTIF(INDIRECT("\'" & A' + row + ' & "\'!B:B"),"欠席"), 0)'
    );
    sheet.getRange(row, 4).setFormula(targetDaysFormula);
    sheet.getRange(row, 2).setFormula(
      '=IFERROR((D' + row + '-C' + row + ')/D' + row + ',"")'
    );
    sheet.getRange(row, 2).setNumberFormat('0.0%');
  }
  sheet.setFrozenRows(1);
}

function createStudentSheets_(ss, roster, studentCount) {
  const template = ss.getSheetByName(SHEET_NAMES.template);
  const names = roster.getRange(2, 4, studentCount, 1).getValues().flat().filter((n) => n);
  names.forEach((name) => {
    if (ss.getSheetByName(name)) return; // 既にあればスキップ（生徒の入力を消さない）
    const newSheet = template.copyTo(ss);
    newSheet.setName(name);
  });
}

// ====== ユーティリティ ======
function columnToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
