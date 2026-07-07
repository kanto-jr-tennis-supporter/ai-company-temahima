# T9：部活動まるごと管理キット ― スプレッドシート構築仕様＋GASコード一式

- 作成日：2026-07-07
- 担当：⚙️ ハック（自動化エンジニア）
- 使い方：このファイルの手順どおりに、社長のGoogleアカウントで組み立てる。**GASコードはコピペ1回で全機能が動く**。タブ・見出しは「初期セットアップ」ボタンが自動生成するので手作業ほぼゼロ。

---

## 1. タブ構成（シート設計）

| タブ名 | 役割 | 主な列・構造 |
|---|---|---|
| **設定** | 部活名・顧問名などの基本情報（連絡文に差し込まれる） | A列=項目／B列=値（B2:部活動名、B3:年度、B4:顧問名、B5:出欠の記号、B6:最高学年） |
| **名簿** | 部員の一元管理。ここにだけ名前を書く | 状態（在籍/引退/卒業）・学年・組・番号・氏名・ふりがな・役職・保護者名・備考 |
| **出欠** | 部員×日付のマトリクス。出/欠/遅/早/見 をプルダウン入力 | A列=氏名（自動反映）／B列以降=日付（1行目） |
| **出欠集計** | ボタンで自動生成される集計結果 | 氏名・出席・欠席・遅刻・早退・見学・未入力・出席率 |
| **集金** | 集金項目×部員の○チェック表 | 1行目=項目名／2行目=金額／3行目=締切／5行目以降=氏名（自動反映）＋○ |
| **会計** | 部費の帳簿。残高は自動計算 | 日付・項目・収入・支出・残高（自動）・備考 |
| **大会・遠征** | 年間の大会情報台帳（連絡文の差し込み元） | 大会名・日付・会場・集合時刻・集合場所・解散予定・持ち物・参加費・弁当・備考 |
| **連絡文テンプレ** | 保護者向け定型文。差し込み用テンプレもここ | A列=テンプレ名／B列=本文 |
| **引き継ぎ** | 次の顧問への申し送り | 項目・内容（年間予定の慣例、業者連絡先、注意点など） |
| **出力** | GASが生成した文章（未回答一覧・催促文・連絡文）の置き場 | 自動生成。コピーしてLINEや配布文書へ |

### 出欠の記号ルール（現場仕様）

- `出`=出席／`欠`=欠席／`遅`=遅刻／`早`=早退／`見`=見学（怪我などで参加はしている扱い）
- **空欄=未回答（未入力）**。ここが「未回答者の自動抽出」の判定基準
- 出席率 =（出＋遅＋早）÷（出＋欠＋遅＋早）。見学と未入力は分母から除外（怪我人の率が不当に下がらないように）

---

## 2. GASコード（コピペで動く完全版）

以下を**丸ごと1ファイル**、Apps Scriptに貼り付ける。

```javascript
/**********************************************************************
 * 部活動まるごと管理キット v1.0
 * 現役教員の部活運営を1ファイルに：出欠・連絡・集金・会計・年度更新
 * 使い方：メニュー「🏸 部活キット」から。初回は「⓪ 初期セットアップ」
 **********************************************************************/

// ---- シート名の定義（タブ名を変えたい場合はここを直す） ----
const SH = {
  SETTEI:   '設定',
  MEIBO:    '名簿',
  SHUKKETSU:'出欠',
  SHUKEI:   '出欠集計',
  SHUKIN:   '集金',
  KAIKEI:   '会計',
  TAIKAI:   '大会・遠征',
  TEMPLATE: '連絡文テンプレ',
  HIKITSUGI:'引き継ぎ',
  OUTPUT:   '出力',
};

// ---- メニュー（ファイルを開くと自動で追加される） ----
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏸 部活キット')
    .addItem('⓪ 初期セットアップ（タブと見出しを自動作成）', 'setupSheets')
    .addSeparator()
    .addItem('① 名簿を出欠・集金シートに反映', 'syncMembers')
    .addItem('② 今日の日付列を出欠に追加', 'addTodayColumn')
    .addItem('③ 出欠を集計する', 'updateSummary')
    .addItem('④ 未入力（未回答）の部員を出す', 'listUnanswered')
    .addItem('⑤ 集金の未納チェック', 'checkPayments')
    .addItem('⑥ 大会・遠征の連絡文を作る', 'makeNotice')
    .addSeparator()
    .addItem('⑦ 年度更新（学年繰り上げ・卒業処理）', 'promoteGrades')
    .addToUi();
}

// ====================================================================
// ⓪ 初期セットアップ：タブ・見出し・入力規則を全自動生成
//    既にあるタブは壊さない（無いものだけ作る）
// ====================================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const created = [];

  // --- 設定 ---
  if (!ss.getSheetByName(SH.SETTEI)) {
    const s = ss.insertSheet(SH.SETTEI);
    s.getRange(1, 1, 6, 2).setValues([
      ['項目', '値'],
      ['部活動名', '○○部'],
      ['年度', '2026'],
      ['顧問名', '山元'],
      ['出欠の記号', '出=出席 欠=欠席 遅=遅刻 早=早退 見=見学（空欄=未回答）'],
      ['最高学年（中学=3、高校=3、小学=6）', 3],
    ]);
    s.getRange('A1:B1').setFontWeight('bold');
    s.setColumnWidth(1, 260); s.setColumnWidth(2, 420);
    created.push(SH.SETTEI);
  }

  // --- 名簿 ---
  if (!ss.getSheetByName(SH.MEIBO)) {
    const s = ss.insertSheet(SH.MEIBO);
    s.getRange(1, 1, 1, 9).setValues([
      ['状態', '学年', '組', '番号', '氏名', 'ふりがな', '役職', '保護者名', '備考'],
    ]).setFontWeight('bold');
    s.getRange(2, 1, 2, 9).setValues([
      ['記入例', 2, 'A', 12, '山田 太郎', 'やまだ たろう', '部長', '山田 花子', '←この行は見本。消してOK'],
      ['在籍', 1, 'B', 3, '佐藤 次郎', 'さとう じろう', '', '佐藤 恵子', '←「在籍」の行だけが反映される'],
    ]);
    // 状態のプルダウン
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['在籍', '引退', '卒業', '転出'], true).setAllowInvalid(true).build();
    s.getRange(2, 1, 300, 1).setDataValidation(rule);
    created.push(SH.MEIBO);
  }

  // --- 出欠 ---
  if (!ss.getSheetByName(SH.SHUKKETSU)) {
    const s = ss.insertSheet(SH.SHUKKETSU);
    s.getRange(1, 1).setValue('氏名').setFontWeight('bold');
    s.setFrozenRows(1); s.setFrozenColumns(1);
    // 出欠のプルダウン（B2から広めに設定しておく）
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['出', '欠', '遅', '早', '見'], true).setAllowInvalid(true).build();
    s.getRange(2, 2, 300, 60).setDataValidation(rule);
    created.push(SH.SHUKKETSU);
  }

  // --- 出欠集計 ---
  if (!ss.getSheetByName(SH.SHUKEI)) {
    const s = ss.insertSheet(SH.SHUKEI);
    s.getRange(1, 1, 1, 8).setValues([
      ['氏名', '出席', '欠席', '遅刻', '早退', '見学', '未入力', '出席率'],
    ]).setFontWeight('bold');
    created.push(SH.SHUKEI);
  }

  // --- 集金 ---
  if (!ss.getSheetByName(SH.SHUKIN)) {
    const s = ss.insertSheet(SH.SHUKIN);
    s.getRange(1, 1, 4, 2).setValues([
      ['集金項目 →', '部費（4月分）'],
      ['金額', 500],
      ['締切', ''],
      ['氏名↓（○=納入済み）', ''],
    ]);
    s.getRange('A1:A4').setFontWeight('bold');
    s.setFrozenRows(4); s.setFrozenColumns(1);
    created.push(SH.SHUKIN);
  }

  // --- 会計 ---
  if (!ss.getSheetByName(SH.KAIKEI)) {
    const s = ss.insertSheet(SH.KAIKEI);
    s.getRange(1, 1, 1, 6).setValues([
      ['日付', '項目', '収入', '支出', '残高（自動）', '備考'],
    ]).setFontWeight('bold');
    // 残高の自動計算（1本のARRAYFORMULAで全行に効く）
    s.getRange('E2').setFormula(
      '=ARRAYFORMULA(IF(B2:B="",,SUMIF(ROW(C2:C),"<="&ROW(C2:C),C2:C)-SUMIF(ROW(D2:D),"<="&ROW(D2:D),D2:D)))'
    );
    created.push(SH.KAIKEI);
  }

  // --- 大会・遠征 ---
  if (!ss.getSheetByName(SH.TAIKAI)) {
    const s = ss.insertSheet(SH.TAIKAI);
    s.getRange(1, 1, 1, 10).setValues([
      ['大会名', '日付', '会場', '集合時刻', '集合場所', '解散予定', '持ち物', '参加費', '弁当', '備考'],
    ]).setFontWeight('bold');
    created.push(SH.TAIKAI);
  }

  // --- 連絡文テンプレ ---
  if (!ss.getSheetByName(SH.TEMPLATE)) {
    const s = ss.insertSheet(SH.TEMPLATE);
    s.getRange(1, 1, 1, 2).setValues([['テンプレ名', '本文']]).setFontWeight('bold');
    s.getRange(2, 1, 5, 2).setValues([
      ['大会連絡（自動差し込み用）', defaultNoticeTemplate_()],
      ['欠席連絡への返信',
       'ご連絡ありがとうございます。承知いたしました。お大事になさってください。\n復帰の際は無理のない範囲で参加させますのでご安心ください。'],
      ['集金のお願い',
       '【集金のお願い】\n{部活動名}保護者の皆様\n\n○○の費用として、△△円を□月□日までに集めます。\nお釣りのないよう封筒に入れ、記名の上お子様に持たせてください。\nよろしくお願いいたします。（顧問 {顧問名}）'],
      ['悪天候などによる中止連絡',
       '【本日の活動中止のお知らせ】\n本日の{部活動名}の活動は、○○のため中止とします。\n生徒は通常どおり下校します。急な連絡となり申し訳ありません。（顧問 {顧問名}）'],
      ['保護者会・引率のお願い',
       '【ご協力のお願い】\n{部活動名}保護者の皆様\n\n○月○日の△△について、ご協力いただける方を募集しています。\n詳細は追ってご連絡しますので、まずはご都合をお知らせください。（顧問 {顧問名}）'],
    ]);
    s.setColumnWidth(1, 220); s.setColumnWidth(2, 600);
    s.getRange('B2:B10').setWrap(true);
    created.push(SH.TEMPLATE);
  }

  // --- 引き継ぎ ---
  if (!ss.getSheetByName(SH.HIKITSUGI)) {
    const s = ss.insertSheet(SH.HIKITSUGI);
    s.getRange(1, 1, 1, 2).setValues([['項目', '内容（次の顧問への申し送り）']]).setFontWeight('bold');
    s.getRange(2, 1, 7, 1).setValues([
      ['年間の大会・行事の流れ'], ['練習の曜日・場所の慣例'], ['部費の金額と集め方'],
      ['用具・備品と購入先（業者連絡先）'], ['外部コーチ・OBとの関係'], ['気をつけたいこと'], ['その他'],
    ]);
    s.setColumnWidth(1, 240); s.setColumnWidth(2, 600);
    created.push(SH.HIKITSUGI);
  }

  // --- 出力 ---
  if (!ss.getSheetByName(SH.OUTPUT)) {
    const s = ss.insertSheet(SH.OUTPUT);
    s.getRange(1, 1).setValue('ここにGASが作った文章（未回答一覧・催促文・連絡文）が追記されます');
    s.setColumnWidth(1, 640);
    created.push(SH.OUTPUT);
  }

  ui.alert(created.length
    ? '初期セットアップ完了！作成したタブ：' + created.join('、') +
      '\n\n次の手順：\n1.「名簿」に部員を入力（状態を「在籍」に）\n2. メニュー①で名簿を反映\n3. メニュー②で今日の日付列を追加'
    : 'すべてのタブが既にあります。セットアップ済みです。');
}

// ====================================================================
// 共通ヘルパー
// ====================================================================
function sheet_(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('シート「' + name + '」が見つかりません。メニュー「⓪ 初期セットアップ」を実行するか、タブ名を確認してください。');
  return s;
}

/** 名簿から「在籍」の部員だけを取り出す */
function activeMembers_() {
  const values = sheet_(SH.MEIBO).getDataRange().getValues();
  const members = [];
  for (let i = 1; i < values.length; i++) {
    const status = String(values[i][0]).trim();
    const name = String(values[i][4]).trim();
    if (status === '在籍' && name !== '') {
      members.push({ row: i + 1, grade: values[i][1], name: name });
    }
  }
  return members;
}

/** 日付を「7月10日（金）」形式の日本語にする */
function jaDate_(v) {
  if (!(v instanceof Date)) return String(v || '');
  const w = ['日', '月', '火', '水', '木', '金', '土'][v.getDay()];
  return Utilities.formatDate(v, Session.getScriptTimeZone(), 'M月d日') + '（' + w + '）';
}

/** 時刻セル（Date型）を「9:30」形式にする */
function jaTime_(v) {
  if (!(v instanceof Date)) return String(v || '');
  return Utilities.formatDate(v, Session.getScriptTimeZone(), 'H:mm');
}

/** 生成した文章を「出力」シートに追記して表示する */
function writeOutput_(title, text) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SH.OUTPUT);
  if (!s) { s = ss.insertSheet(SH.OUTPUT); s.setColumnWidth(1, 640); }
  const row = s.getLastRow() + 2;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d H:mm');
  s.getRange(row, 1).setValue('■ ' + title + '（' + now + ' 作成）').setFontWeight('bold');
  s.getRange(row + 1, 1).setValue(text).setWrap(true);
  s.activate();
  s.setActiveSelection(s.getRange(row + 1, 1));
}

// ====================================================================
// ① 名簿 → 出欠・集金シートへ反映
// ====================================================================
function syncMembers() {
  const ui = SpreadsheetApp.getUi();
  const members = activeMembers_();
  if (!members.length) {
    ui.alert('名簿に「在籍」の部員がいません。\n名簿シートのA列（状態）が「在籍」になっているか確認してください。');
    return;
  }
  const names = members.map(m => [m.name]);

  // 出欠：A2以降（古い名前は消してから貼り直す。出欠データの列はそのまま残る）
  const shu = sheet_(SH.SHUKKETSU);
  if (shu.getLastRow() > 1) shu.getRange(2, 1, shu.getLastRow() - 1, 1).clearContent();
  shu.getRange(2, 1, names.length, 1).setValues(names);

  // 集金：A5以降
  const kin = sheet_(SH.SHUKIN);
  if (kin.getLastRow() > 4) kin.getRange(5, 1, kin.getLastRow() - 4, 1).clearContent();
  kin.getRange(5, 1, names.length, 1).setValues(names);

  ui.alert('在籍 ' + names.length + ' 名を「出欠」「集金」シートに反映しました。');
}

// ====================================================================
// ② 出欠シートに今日の日付列を追加
// ====================================================================
function addTodayColumn() {
  const sh = sheet_(SH.SHUKKETSU);
  const col = Math.max(sh.getLastColumn() + 1, 2);
  sh.getRange(1, col).setValue(new Date()).setNumberFormat('m/d');
  // 新しい列にも出欠プルダウンを付ける
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['出', '欠', '遅', '早', '見'], true).setAllowInvalid(true).build();
  sh.getRange(2, col, 300, 1).setDataValidation(rule);
  sh.activate();
  SpreadsheetApp.getUi().alert('今日（' + jaDate_(new Date()) + '）の列を追加しました。出欠を入力してください。');
}

// ====================================================================
// ③ 出欠の自動集計
// ====================================================================
function updateSummary() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_(SH.SHUKKETSU);
  const data = sh.getDataRange().getValues();
  if (data.length < 2 || data[0].length < 2) {
    ui.alert('出欠シートに部員か日付列がありません。メニュー①②を先に実行してください。');
    return;
  }
  const header = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    if (!name) continue;
    const c = { 出: 0, 欠: 0, 遅: 0, 早: 0, 見: 0, 未: 0 };
    for (let j = 1; j < header.length; j++) {
      if (header[j] === '') continue; // 日付の入っていない列は数えない
      const v = String(data[i][j]).trim();
      if (v in c) c[v]++;
      else if (v === '') c.未++;
    }
    const bunbo = c.出 + c.欠 + c.遅 + c.早; // 見学・未入力は分母に入れない
    const rate = bunbo > 0 ? (c.出 + c.遅 + c.早) / bunbo : '';
    rows.push([name, c.出, c.欠, c.遅, c.早, c.見, c.未, rate]);
  }
  const out = sheet_(SH.SHUKEI);
  out.clearContents();
  out.getRange(1, 1, 1, 8).setValues([['氏名', '出席', '欠席', '遅刻', '早退', '見学', '未入力', '出席率']])
     .setFontWeight('bold');
  if (rows.length) {
    out.getRange(2, 1, rows.length, 8).setValues(rows);
    out.getRange(2, 8, rows.length, 1).setNumberFormat('0.0%');
  }
  out.activate();
  ui.alert('出欠集計が完了しました（' + rows.length + ' 名）。「出欠集計」シートをご覧ください。');
}

// ====================================================================
// ④ 未入力（未回答）の部員を抽出 → 一覧＋催促文を「出力」へ
// ====================================================================
function listUnanswered() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_(SH.SHUKKETSU);
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol < 2 || lastRow < 2) {
    ui.alert('出欠シートに日付列か部員がありません。メニュー①②を先に実行してください。');
    return;
  }
  const res = ui.prompt('未入力チェック',
    '日付を「7/10」の形で入力してください。\n（空欄のままOK ＝ いちばん右の列＝直近の日付）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const q = res.getResponseText().trim();

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const tz = Session.getScriptTimeZone();
  let col = lastCol;
  if (q) {
    col = -1;
    for (let j = 1; j < lastCol; j++) {
      const h = header[j];
      const label = (h instanceof Date) ? Utilities.formatDate(h, tz, 'M/d') : String(h).trim();
      if (label === q) { col = j + 1; break; }
    }
    if (col === -1) { ui.alert('「' + q + '」の日付列が見つかりませんでした。'); return; }
  }
  const hv = sh.getRange(1, col).getValue();
  const dateLabel = (hv instanceof Date) ? jaDate_(hv) : String(hv);

  const data = sh.getRange(2, 1, lastRow - 1, col).getValues();
  const missing = data
    .filter(r => String(r[0]).trim() !== '' && String(r[col - 1]).trim() === '')
    .map(r => String(r[0]).trim());

  if (!missing.length) {
    ui.alert(dateLabel + ' は全員入力済みです。お疲れさまでした！');
    return;
  }
  const text =
    '【出欠 未回答】' + dateLabel + '\n' +
    '未入力：' + missing.length + ' 名\n・' + missing.join('\n・') + '\n\n' +
    '――― そのまま送れる催促文 ―――\n' +
    dateLabel + 'の出欠がまだ入力されていません。本日中に入力をお願いします。（顧問より）';
  writeOutput_('未回答チェック ' + dateLabel, text);
  ui.alert(dateLabel + ' の未入力は ' + missing.length + ' 名です。\n「出力」シートに一覧と催促文を書き出しました。');
}

// ====================================================================
// ⑤ 集金の未納チェック → 項目ごとの未納者一覧を「出力」へ
// ====================================================================
function checkPayments() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_(SH.SHUKIN);
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol < 2 || lastRow < 5) {
    ui.alert('集金シートに項目か名簿がありません。\n1行目に項目名、メニュー①で名簿を反映してから実行してください。');
    return;
  }
  const meta = sh.getRange(1, 2, 3, lastCol - 1).getValues(); // 1行目:項目名 2行目:金額 3行目:締切
  const names = sh.getRange(5, 1, lastRow - 4, 1).getValues().map(r => String(r[0]).trim());
  const marks = sh.getRange(5, 2, lastRow - 4, lastCol - 1).getValues();

  const lines = [];
  let totalUnpaid = 0;
  for (let j = 0; j < lastCol - 1; j++) {
    const item = String(meta[0][j]).trim();
    if (!item) continue;
    const amount = meta[1][j];
    const due = meta[2][j];
    const unpaid = [];
    for (let i = 0; i < names.length; i++) {
      if (names[i] && String(marks[i][j]).trim() === '') unpaid.push(names[i]);
    }
    totalUnpaid += unpaid.length;
    const dueLabel = (due instanceof Date) ? jaDate_(due) : String(due || '未設定');
    lines.push(
      '▼ ' + item + '（' + (amount ? amount + '円' : '金額未設定') + '／締切：' + dueLabel + '）\n' +
      (unpaid.length
        ? '  未納 ' + unpaid.length + ' 名：' + unpaid.join('、')
        : '  全員納入済みです！')
    );
  }
  if (!lines.length) { ui.alert('集金シートの1行目（B列以降）に集金項目名を入力してください。'); return; }

  const text = lines.join('\n\n') + '\n\n――― そのまま送れる催促文 ―――\n' +
    '集金がまだの人は、封筒に記名の上、締切までに顧問へ提出してください。よろしくお願いします。';
  writeOutput_('集金チェック', text);
  ui.alert('集金チェックが完了しました（未納のべ ' + totalUnpaid + ' 名）。\n「出力」シートに一覧と催促文を書き出しました。');
}

// ====================================================================
// ⑥ 大会・遠征の連絡文を自動生成（テンプレに差し込み）
// ====================================================================
function makeNotice() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_(SH.TAIKAI);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    ui.alert('「大会・遠征」シートに予定を1行入力してから実行してください。');
    return;
  }
  const res = ui.prompt('連絡文の作成',
    '何行目の予定の連絡文を作りますか？（2 ＝ 1件目）\n（空欄のままOK ＝ いちばん下の行＝最新の予定）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const q = res.getResponseText().trim();
  const row = q ? parseInt(q, 10) : lastRow;
  if (!(row >= 2 && row <= lastRow)) { ui.alert('行番号が正しくありません（2〜' + lastRow + '）。'); return; }

  const v = sh.getRange(row, 1, 1, 10).getValues()[0];
  const [name, date, venue, meetTime, meetPlace, endTime, items, fee, bento, memo] = v;

  // 設定シートから部活名・顧問名
  const settei = sheet_(SH.SETTEI);
  const clubName = String(settei.getRange('B2').getValue() || '本部活動');
  const teacher  = String(settei.getRange('B4').getValue() || '顧問');

  // テンプレ（連絡文テンプレシートにあればそれを優先。無ければ内蔵テンプレ）
  const tpl = findTemplate_('大会連絡（自動差し込み用）') || defaultNoticeTemplate_();

  const feeLabel = (typeof fee === 'number') ? fee + '円' : String(fee || 'なし');
  const text = tpl
    .replace(/{部活動名}/g, clubName)
    .replace(/{顧問名}/g, teacher)
    .replace(/{大会名}/g, String(name || ''))
    .replace(/{日付}/g, jaDate_(date))
    .replace(/{会場}/g, String(venue || ''))
    .replace(/{集合時刻}/g, jaTime_(meetTime))
    .replace(/{集合場所}/g, String(meetPlace || ''))
    .replace(/{解散予定}/g, jaTime_(endTime))
    .replace(/{持ち物}/g, String(items || ''))
    .replace(/{参加費}/g, feeLabel)
    .replace(/{弁当}/g, String(bento || ''))
    .replace(/{備考}/g, memo ? '■その他：' + String(memo) : '');

  writeOutput_('連絡文：' + String(name || '大会'), text);
  ui.alert('「' + String(name || '大会') + '」の連絡文を作成しました。\n「出力」シートからコピーしてお使いください。');
}

/** 連絡文テンプレシートから名前でテンプレ本文を探す */
function findTemplate_(tplName) {
  const s = SpreadsheetApp.getActive().getSheetByName(SH.TEMPLATE);
  if (!s || s.getLastRow() < 2) return '';
  const data = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  for (const r of data) {
    if (String(r[0]).trim() === tplName && String(r[1]).trim() !== '') return String(r[1]);
  }
  return '';
}

/** 内蔵の大会連絡テンプレ */
function defaultNoticeTemplate_() {
  return [
    '【{大会名}のご連絡】',
    '{部活動名}保護者の皆様',
    '',
    'いつもお世話になっております。顧問の{顧問名}です。',
    '下記のとおり{大会名}に参加しますので、ご連絡いたします。',
    '',
    '■日時：{日付}',
    '■会場：{会場}',
    '■集合：{集合時刻} に {集合場所}',
    '■解散：{解散予定} ごろ（予定）',
    '■持ち物：{持ち物}',
    '■参加費：{参加費}',
    '■昼食：{弁当}',
    '{備考}',
    '',
    'ご不明な点があれば{顧問名}までご連絡ください。',
    'よろしくお願いいたします。',
  ].join('\n');
}

// ====================================================================
// ⑦ 年度更新：学年繰り上げ＋卒業処理
// ====================================================================
function promoteGrades() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('年度更新',
    '在籍部員の学年を1つ繰り上げます。最高学年は「卒業」に変わります。\n' +
    '※実行前に「ファイル → コピーを作成」で今年度の控えを残しておくのがおすすめです。\n\n実行しますか？',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const maxGrade = Number(sheet_(SH.SETTEI).getRange('B6').getValue()) || 3;
  const sh = sheet_(SH.MEIBO);
  const data = sh.getDataRange().getValues();
  let up = 0, grad = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== '在籍') continue;
    const g = Number(data[i][1]);
    if (!g) continue;
    if (g >= maxGrade) { sh.getRange(i + 1, 1).setValue('卒業'); grad++; }
    else { sh.getRange(i + 1, 2).setValue(g + 1); up++; }
  }
  ui.alert('年度更新が完了しました。進級 ' + up + ' 名／卒業 ' + grad + ' 名。\n' +
    'このあとメニュー「① 名簿を出欠・集金シートに反映」を実行して、名簿を貼り直してください。');
}
```

---

## 3. 設置手順（社長のGoogleアカウントで・1ステップずつ）

> 所要：約15分（シート自動生成5分＋動作確認10分）。スマホでも読めますが、**作業はPCで**。

### STEP 1：スプレッドシートを新規作成
1. https://sheets.new をブラウザで開く（Googleにログインした状態で）
2. 左上のファイル名を「部活動まるごと管理キット」に変更

### STEP 2：GASコードを貼る
1. メニュー「拡張機能」→「Apps Script」を開く
2. エディタに最初から入っている `function myFunction() {...}` を**全部消す**
3. 上の「2. GASコード」を**丸ごとコピーして貼り付け**
4. 上部の💾（プロジェクトを保存）を押す。プロジェクト名は「部活キット」などに
5. Apps Scriptのタブは閉じてよい

### STEP 3：メニューを出す
1. スプレッドシートのタブに戻り、**ページを再読み込み**（⌘R）
2. 数秒待つと、メニューバーに **「🏸 部活キット」** が現れる

### STEP 4：初回の承認（1回だけ）
1. 「🏸 部活キット」→「⓪ 初期セットアップ」をクリック
2. 「承認が必要です」→ 自分のGoogleアカウントを選択
3. 「このアプリは Google で確認されていません」と出たら：**「詳細」→「（プロジェクト名）に移動」→「許可」**
   ※自分で作った自分用のスクリプトなので、この警告は正常です
4. もう一度「⓪ 初期セットアップ」を実行 → **10個のタブと見出しが自動で揃う**

### STEP 5：動作確認（このシナリオを一周する）
1. 「名簿」に部員を3人ほど入力（**状態を「在籍」に**。記入例の行は消してOK）
2. メニュー「① 名簿を反映」→ 出欠・集金のA列に名前が入る
3. メニュー「② 今日の日付列を追加」→ 出欠に今日の列ができる
4. 出欠を2人だけ入力し、1人はわざと空欄のままにする
5. メニュー「③ 出欠を集計」→「出欠集計」タブに出席率が出る
6. メニュー「④ 未入力の部員を出す」（空欄でOK）→「出力」タブに未回答者と催促文が出る
7. 「集金」のB1に項目・B2に金額を入れ、1人だけ○ → メニュー「⑤ 未納チェック」→ 未納者一覧が出る
8. 「大会・遠征」に予定を1行入力 → メニュー「⑥ 連絡文を作る」→ 保護者向け連絡文が完成
9. 「設定」タブの部活動名・顧問名を自分のものに直す（連絡文に差し込まれます）

### STEP 6：商品として配る準備（販売時）
1. 完成したシートを「ファイル → コピーを作成」してマスター版を保管
2. 配布用のシートを「共有」→「リンクを知っている全員：**閲覧者**」に設定
3. 購入者への案内は「リンクを開く → ファイル → **コピーを作成**」。**GASはコピーに丸ごと付いてくる**ので、購入者側の設置作業はSTEP 4の承認だけ
4. 注意：配布用マスターには**実在の生徒名を絶対に残さない**（記入例はダミー名で）

---

## 4. 実装メモ（社長・引き継ぎ用）

- **個人情報の扱い**：このキットは外部送信を一切しない（GmailApp等は不使用）。データは顧問のGoogleドライブ内で完結。ここは販売ページでも安心材料として書く価値あり
- **既存の「部活動出欠確認シート」との合流**：既存シートの出欠データは、列構成（A列=氏名、1行目=日付）が合えばそのままコピペで移行可能
- 拡張候補（v2）：Googleフォーム連携（部員が自分で出欠回答→自動転記）／出席率の月別推移グラフ／LINE公式アカウントとの連携。**v1には入れない**（設置の簡単さが商品価値のため）
