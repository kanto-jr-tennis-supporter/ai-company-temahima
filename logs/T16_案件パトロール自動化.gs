/************************************************************************
 * 🔍 毎朝の案件パトロール自動化（クラウドワークス新着チェック）v1.0
 *
 * ── これは何？ ─────────────────────────────────────────────
 * T7で作った「毎朝ブックマーク8個をクリックする手動パトロール」を、
 * このスプレッドシート1枚で自動化する試みです。
 * 新しい空のスプレッドシートにこのファイルを貼り付けて「setup」を
 * 1回実行すると、シート一式ができます。
 *
 * ── ⚠️ 正直な注記（いちばん大事）─────────────────────────────
 * クラウドワークスは近年、ブラウザ以外からのアクセス（bot判定される
 * アクセス）を広くブロックしている可能性があります。作成中に検証した
 * ところ、Claude Codeの作業環境からは crowdworks.jp への接続が403で
 * 拒否され、さらに別経路（Anthropicのフェッチ基盤）で robots.txt を
 * 読みにいっても403でした。robots.txtは通常どのサイトでも誰でも読める
 * ファイルなので、これがブロックされるのは「特定のURLだけを狙い撃ち」
 * ではなく「入口（WAF）で機械的なアクセス全般を止めている」可能性を
 * 示す傍証です。
 *
 * Google Apps Script の UrlFetchApp は Google のサーバーから実行される
 * ため、上と同じ制限を受けない可能性はありますが、**保証はできません**。
 * 「動くはず」であって「確実に動く」ではないことを、社長には正直に
 * お伝えください。実際に動くかどうかは、初回実行（testFetch）で必ず
 * 確認してください。
 *
 * ブロックされていた場合の代替：
 * ・クラウドワークスの新着RSS配信は2018年12月に終了済み（公式アナウンス
 *   済み）。公開API等の代替手段も見当たりません。
 * ・そのため「取得」の自動化がダメなら、次善策は「判定」だけを自動化
 *   することです。このシートには手動貼付シートも同梱しています
 *   （T7の毎朝ブックマーク巡回はそのまま人がやり、見つけた案件の
 *   タイトルをこのシートに貼るだけで、即×判定・◎フラグ・報酬チェックを
 *   自動でやってくれます＝判断の自動化）。
 *
 * ── 初回にやること ───────────────────────────────────────
 * 1. メニュー「🔍案件パトロール」→「🛠 初期セットアップ」
 * 2. メニュー「🔍案件パトロール」→「🧪 接続テスト（1件だけ）」
 *    → ここでHTTPステータスが表示されます。200なら本番実行へ、
 *      403等ならブロックされています（「実行ログ」シートに詳細記録済み）。
 * 3. 問題なければ「⏰ 毎朝7時の自動実行をON」でトリガー登録。
 ************************************************************************/

/* ===== 基本設定 ===== */

var SHEET = {
  GUIDE:   'はじめに',
  LOG:     '案件ログ',
  CONFIG:  '設定',
  RUNLOG:  '実行ログ',
  PASTE:   '手動貼付判定',
  RAWHTML: '生HTML保存'
};

var COLOR = {
  CREAM:      '#fff8e1',
  RED_LIGHT:  '#f4cccc',
  YEL_LIGHT:  '#fff2cc',
  GREEN_LIGHT:'#d9ead3',
  HEADER:     '#efefef'
};

// T7の検索URL8本（★1〜3=毎日既定ON、☆4〜8=週次既定OFF）。
// 「設定」シートで各行のON/OFFをいつでも変更できます。
var DEFAULT_SOURCES = [
  { on: true,  label: '★1 GAS',              keyword: 'GAS',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=GAS&order=new' },
  { on: true,  label: '★2 スプレッドシート 自動化', keyword: 'スプレッドシート 自動化',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=%E3%82%B9%E3%83%97%E3%83%AC%E3%83%83%E3%83%89%E3%82%B7%E3%83%BC%E3%83%88%20%E8%87%AA%E5%8B%95%E5%8C%96' },
  { on: true,  label: '★3 Apps Script',      keyword: 'Apps Script',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=Apps%20Script&order=new' },
  { on: false, label: '☆4 自動転記',          keyword: '自動転記',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=%E8%87%AA%E5%8B%95%E8%BB%A2%E8%A8%98&order=new' },
  { on: false, label: '☆5 集計 効率化',       keyword: '集計 効率化',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=%E9%9B%86%E8%A8%88%20%E5%8A%B9%E7%8E%87%E5%8C%96' },
  { on: false, label: '☆6 スプレッドシート 関数', keyword: 'スプレッドシート 関数',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=%E3%82%B9%E3%83%97%E3%83%AC%E3%83%83%E3%83%89%E3%82%B7%E3%83%BC%E3%83%88%20%E9%96%A2%E6%95%B0' },
  { on: false, label: '☆7 Excel マクロ',      keyword: 'Excel マクロ',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=Excel%20%E3%83%9E%E3%82%AF%E3%83%AD&order=new' },
  { on: false, label: '☆8 業務 自動化',       keyword: '業務 自動化',
    url: 'https://crowdworks.jp/public/jobs/search?search%5Bkeywords%5D=%E6%A5%AD%E5%8B%99%20%E8%87%AA%E5%8B%95%E5%8C%96&order=new' }
];

// T7の30秒足切り基準（即×ワード／◎ワード）。「設定」シートで追加・削除可。
var DEFAULT_NG_WORDS = [
  '税理士', '社労士', '行政書士', '有資格者限定',
  '毎日対応', '日次報告', '即レス', '常駐', 'シフト', '平日日中', 'カスタマーサポート', 'CS', 'コールセンター',
  'データ入力', '文字起こし', '目視確認', 'チェック作業', 'モニター', 'アンケート',
  'フォーム送信', '営業代行 送信', 'ボタンを押すだけ', '1日5分で',
  'Web制作', 'LP', 'デザイン', '動画編集', 'Python環境構築', 'AWS', 'DB構築'
];

var DEFAULT_OK_WORDS = [
  'GAS', 'Apps Script', 'スプレッドシート', '自動化', '転記', '集計', '連携',
  'Googleフォーム', '請求書 自動', '管理シート'
];

var DAILY_TRIGGER_HOUR = 7; // 毎朝7時

/* =====================================================================
 * メニュー
 * =================================================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍案件パトロール')
    .addItem('🛠 初期セットアップ', 'setup')
    .addSeparator()
    .addItem('🧪 接続テスト（1件だけ・まずこれ）', 'testFetch')
    .addItem('🔎 今すぐ全件パトロール実行', 'patrol')
    .addSeparator()
    .addItem('⏰ 毎朝7時の自動実行をON', 'createDailyTrigger')
    .addItem('⏹ 自動実行をOFF', 'removeDailyTrigger')
    .addSeparator()
    .addItem('📋 手動貼付シートを判定する', 'judgePasted')
    .addItem('📖 使い方を開く（はじめに）', 'showGuide')
    .addToUi();
}

function showGuide() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET.GUIDE);
  if (sh) ss.setActiveSheet(sh);
}

/* =====================================================================
 * ① 初期セットアップ
 * =================================================================== */

function setup() {
  var ss = SpreadsheetApp.getActive();
  var made = [], skipped = [];

  var guide  = ensureSheet_(ss, SHEET.GUIDE, 0);
  var log    = ensureSheet_(ss, SHEET.LOG, 1);
  var config = ensureSheet_(ss, SHEET.CONFIG, 2);
  var runlog = ensureSheet_(ss, SHEET.RUNLOG, 3);
  var paste  = ensureSheet_(ss, SHEET.PASTE, 4);
  var raw    = ensureSheet_(ss, SHEET.RAWHTML, 5);

  if (config.created) { buildConfig_(config.sheet); made.push(SHEET.CONFIG); } else skipped.push(SHEET.CONFIG);
  if (log.created)    { buildLog_(log.sheet);       made.push(SHEET.LOG); }    else skipped.push(SHEET.LOG);
  if (runlog.created) { buildRunlog_(runlog.sheet);  made.push(SHEET.RUNLOG); } else skipped.push(SHEET.RUNLOG);
  if (paste.created)  { buildPaste_(paste.sheet);    made.push(SHEET.PASTE); }  else skipped.push(SHEET.PASTE);
  if (raw.created)    { buildRaw_(raw.sheet);        made.push(SHEET.RAWHTML); } else skipped.push(SHEET.RAWHTML);
  if (guide.created)  { buildGuide_(guide.sheet);    made.push(SHEET.GUIDE); }  else skipped.push(SHEET.GUIDE);

  removeDefaultSheet_(ss);
  if (/無題|Untitled/i.test(ss.getName())) ss.rename('案件パトロール自動化');

  if (guide.sheet) ss.setActiveSheet(guide.sheet);
  var msg = 'セットアップが完了しました！🔍\n\n';
  if (made.length)    msg += '✅ 作成：' + made.join('・') + '\n';
  if (skipped.length) msg += '⏭ すでにあるため触りませんでした：' + skipped.join('・') + '\n';
  msg += '\nまず「はじめに」をお読みのうえ、メニューの「🧪 接続テスト」を必ず実行してください。' +
         '\n（クラウドワークス側でブロックされていないか、ここで最初に確認します）';
  alert_(msg);
}

/* =====================================================================
 * ② 接続テスト：★1本だけ取得して、成功/失敗をその場で見せる
 * ---------------------------------------------------------------------
 * 本番実行の前に必ずこれで確認する。ブロックされていても実行ログ・
 * 生HTML保存シートに証拠が残るので、社長がそのままハックに転送できる。
 * =================================================================== */

function testFetch() {
  var ss = SpreadsheetApp.getActive();
  var target = DEFAULT_SOURCES[0];
  var result = fetchSearchPage_(target.url);
  writeRunlog_(ss, target.label, result);

  var msg;
  if (result.ok) {
    var jobs = parseJobs_(result.html);
    msg = '✅ 接続できました！（HTTPステータス ' + result.code + '）\n\n' +
          '「' + target.label + '」から ' + jobs.length + ' 件の案件を検出しました。\n';
    if (jobs.length === 0) {
      msg += '\n⚠️ ただし0件でした。接続はできてもページの中身が想定と違う（構造が変わっている）' +
             '可能性があります。「生HTML保存」シートに今回取得したHTMLを保存したので、ハックに' +
             '共有してください（解析ルールを直します）。';
      saveRawHtml_(ss, target.label, result.html);
    } else {
      msg += '\nこのまま「🔎 今すぐ全件パトロール実行」に進んで問題ありません。';
    }
  } else {
    msg = '❌ 接続できませんでした（' + (result.code || 'エラー') + '）。\n\n' +
          '理由：' + result.error + '\n\n' +
          'クラウドワークス側でbotアクセスとしてブロックされている可能性が高いです。' +
          '「実行ログ」シートに詳細を記録しました。\n' +
          '自動取得はここで一旦あきらめ、「📋 手動貼付シートを判定する」（T7の目視巡回＋自動判定の' +
          '合わせ技）に切り替えることをおすすめします。';
  }
  alert_(msg);
}

/* =====================================================================
 * ③ 本番パトロール（トリガーからも手動からも呼ばれる）
 * =================================================================== */

function patrol() {
  var ss = SpreadsheetApp.getActive();
  var config = ss.getSheetByName(SHEET.CONFIG);
  var log = ss.getSheetByName(SHEET.LOG);
  if (!config || !log) { alert_('先に「🛠 初期セットアップ」を実行してください。'); return; }

  var sources = readSources_(config);
  var ngWords = readWordList_(config, 'NG');
  var okWords = readWordList_(config, 'OK');
  var existingUrls = readExistingUrls_(log);

  var totalNew = 0, totalErr = 0;
  var reportLines = [];

  sources.forEach(function (src) {
    if (!src.on) return;
    var result = fetchSearchPage_(src.url);
    writeRunlog_(ss, src.label, result);

    if (!result.ok) {
      totalErr++;
      reportLines.push('❌ ' + src.label + '：取得失敗（' + (result.code || 'エラー') + '）');
      return;
    }

    var jobs = parseJobs_(result.html);
    if (jobs.length === 0) {
      saveRawHtml_(ss, src.label, result.html);
      reportLines.push('⚠️ ' + src.label + '：接続OKだが0件パース（生HTML保存済み・要確認）');
      return;
    }

    var newCount = 0;
    jobs.forEach(function (job) {
      if (existingUrls[job.url]) return; // 既出はスキップ（重複防止）
      var cls = classify_(job.title, job.pay, ngWords, okWords);
      appendLogRow_(log, src.keyword, job, cls, '自動取得');
      existingUrls[job.url] = true;
      newCount++;
    });
    totalNew += newCount;
    reportLines.push('✅ ' + src.label + '：新着' + newCount + '件（取得' + jobs.length + '件中）');
  });

  var activeNote = sources.filter(function (s) { return s.on; }).length + '本の検索をチェック';
  var msg = '🔍 パトロール完了（' + activeNote + '）\n\n' + reportLines.join('\n') +
    '\n\n新規追加：' + totalNew + '件／エラー：' + totalErr + '件\n' +
    '詳細は「案件ログ」「実行ログ」シートをご確認ください。';
  alert_(msg);
}

/* =====================================================================
 * ④ 取得（UrlFetchApp）。ブラウザに近いヘッダーを付け、例外を握りつぶさず
 *    ステータスコードとエラー内容を必ず持ち帰る。
 * =================================================================== */

function fetchSearchPage_(url) {
  var options = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
      'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8'
    }
  };
  try {
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      return { ok: true, code: code, html: res.getContentText(), error: '' };
    }
    return { ok: false, code: code, html: res.getContentText(), error: 'HTTPステータス ' + code + '（アクセス拒否の可能性）' };
  } catch (e) {
    return { ok: false, code: '', html: '', error: String(e) };
  }
}

/* =====================================================================
 * ⑤ 解析（あくまでベストエフォート・要検証）
 * ---------------------------------------------------------------------
 * ⚠️ この作業環境からは crowdworks.jp に一切接続できないため、実際の
 * ページHTML構造を見て検証することができていません。以下は「案件詳細
 * ページへのリンクは /public/jobs/数字 という形式」という既知の情報を
 * 手がかりにした、構造に依存しすぎない汎用的な抜き出し方です。
 * 初回実行で0件だったりおかしな値が入る場合は、「生HTML保存」シートの
 * 中身を見て、ここの正規表現をハックが調整します。
 * =================================================================== */

function parseJobs_(html) {
  var jobs = [];
  if (!html) return jobs;

  // 案件詳細へのリンク（/public/jobs/12345678）の出現位置をすべて拾う
  var linkPattern = /href="(\/public\/jobs\/(\d+))"/g;
  var seenIds = {};
  var matches = [];
  var m;
  while ((m = linkPattern.exec(html)) !== null) {
    var id = m[2];
    if (seenIds[id]) continue; // 同じ案件への複数リンク（サムネ・タイトル等）は最初の1回だけ使う
    seenIds[id] = true;
    matches.push({ id: id, path: m[1], index: m.index });
  }

  matches.forEach(function (mt, i) {
    var winStart = mt.index;
    var winEnd = (i + 1 < matches.length) ? matches[i + 1].index : Math.min(html.length, winStart + 4000);
    var block = html.substring(winStart, Math.min(winEnd, winStart + 4000));

    // タイトル：リンクのすぐ後ろのタグの中身をテキスト化して最初の意味のある文字列を採用
    var titleMatch = block.match(/>([^<>]{4,120})</);
    var title = titleMatch ? cleanText_(titleMatch[1]) : '（タイトル抽出失敗・要確認）';

    // 報酬：「円」を含む金額らしき文字列
    var payMatch = block.match(/([0-9,，]{2,9}\s*(?:円|~[0-9,，]{2,9}\s*円))/);
    var pay = payMatch ? payMatch[1].replace(/，/g, ',') : '';

    // 応募数：「応募」＋数字、または数字＋「人」等
    var appMatch = block.match(/応募[^0-9]{0,4}([0-9]{1,4})\s*(?:人|件)?/);
    var applicants = appMatch ? appMatch[1] : '';

    jobs.push({
      id: mt.id,
      url: 'https://crowdworks.jp' + mt.path,
      title: title,
      pay: pay,
      applicants: applicants
    });
  });

  return jobs;
}

function cleanText_(s) {
  return String(s).replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/\s+/g, ' ').trim();
}

/* =====================================================================
 * ⑥ 判定（T7の30秒足切り基準をそのまま関数化）
 * =================================================================== */

function classify_(title, pay, ngWords, okWords) {
  var t = title || '';

  for (var i = 0; i < ngWords.length; i++) {
    if (ngWords[i] && t.indexOf(ngWords[i]) !== -1) {
      return { status: '除外', reason: '即×ワード「' + ngWords[i] + '」' };
    }
  }

  // 時給1,500円未満の除外（タイトルに「時給」表記があり、金額が拾えた場合のみ判定）
  if (pay && t.indexOf('時給') !== -1) {
    var nums = pay.replace(/,/g, '').match(/[0-9]+/g);
    if (nums) {
      var minNum = Math.min.apply(null, nums.map(Number));
      if (minNum > 0 && minNum < 1500) {
        return { status: '除外', reason: '時給' + minNum + '円（1,500円未満）' };
      }
    }
  }

  var hitOk = [];
  for (var j = 0; j < okWords.length; j++) {
    if (okWords[j] && t.indexOf(okWords[j]) !== -1) hitOk.push(okWords[j]);
  }
  if (hitOk.length > 0) {
    return { status: '要チェック◎', reason: '◎ワード「' + hitOk.join('・') + '」' };
  }

  return { status: '通常', reason: '' };
}

/* =====================================================================
 * ⑦ 手動貼付シートの判定（自動取得がブロックされた場合の次善策）
 * ---------------------------------------------------------------------
 * 使い方：T7のとおり社長がブックマークを目で見て、気になった案件の
 * 「タイトル」と「URL」（分かれば報酬も）だけ「手動貼付判定」シートに
 * 貼り付ける → このメニューを実行 → 即×判定・◎フラグが自動で入る。
 * 「取得」は自動化できなくても「判断」だけは自動化できる、という設計。
 * =================================================================== */

function judgePasted() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET.PASTE);
  var config = ss.getSheetByName(SHEET.CONFIG);
  if (!sh || !config) { alert_('先に「🛠 初期セットアップ」を実行してください。'); return; }

  var ngWords = readWordList_(config, 'NG');
  var okWords = readWordList_(config, 'OK');

  var lastRow = sh.getLastRow();
  if (lastRow < 3) { alert_('貼り付けられたデータがありません。3行目からタイトル・URL・報酬（任意）を貼ってください。'); return; }

  var range = sh.getRange(3, 1, lastRow - 2, 5); // A題名 B URL C報酬 D判定 E理由
  var values = range.getValues();
  var count = 0;

  for (var i = 0; i < values.length; i++) {
    var title = String(values[i][0]).trim();
    if (title === '' || values[i][3] !== '') continue; // 空行・判定済みはスキップ
    var pay = String(values[i][2]).trim();
    var cls = classify_(title, pay, ngWords, okWords);
    values[i][3] = cls.status;
    values[i][4] = cls.reason;
    count++;
  }
  range.setValues(values);
  alert_('判定しました（' + count + '件）。「除外」以外の行を上から確認してください。');
}

/* =====================================================================
 * トリガー管理
 * =================================================================== */

function createDailyTrigger() {
  removeDailyTrigger(); // 二重登録防止
  ScriptApp.newTrigger('patrol')
    .timeBased()
    .everyDays(1)
    .atHour(DAILY_TRIGGER_HOUR)
    .create();
  alert_('⏰ 毎朝' + DAILY_TRIGGER_HOUR + '時ごろの自動実行をONにしました。\n' +
         '（Googleの時間主導トリガーは「ちょうど' + DAILY_TRIGGER_HOUR + '時」ではなく前後15分程度の幅があります）');
}

function removeDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'patrol') { ScriptApp.deleteTrigger(t); removed++; }
  });
  if (removed > 0) alert_('自動実行をOFFにしました（' + removed + '件のトリガーを削除）。');
}

/* =====================================================================
 * シート構築・補助関数
 * =================================================================== */

function ensureSheet_(ss, name, index) {
  var sh = ss.getSheetByName(name);
  if (sh) return { sheet: sh, created: false };
  return { sheet: ss.insertSheet(name, index), created: true };
}

function removeDefaultSheet_(ss) {
  ['シート1', 'Sheet1'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && ss.getSheets().length > 1 && sh.getLastRow() === 0 && sh.getLastColumn() === 0) {
      ss.deleteSheet(sh);
    }
  });
}

function alert_(msg) {
  try { SpreadsheetApp.getUi().alert('🔍案件パトロール', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }
}

function buildConfig_(sh) {
  sh.setTabColor('#999999');
  sh.getRange('A1').setValue('⚙️ 設定 ― 色つきセルだけ編集OK').setFontWeight('bold').setFontSize(12);

  // 検索ソース一覧
  sh.getRange('A3:D3').setValues([['実行する', 'ラベル', 'キーワード', '検索URL']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);
  var rows = DEFAULT_SOURCES.map(function (s) { return [s.on, s.label, s.keyword, s.url]; });
  sh.getRange(4, 1, rows.length, 4).setValues(rows);
  sh.getRange(4, 1, rows.length, 1).insertCheckboxes();
  sh.getRange(4, 1, rows.length, 4).setBackground(COLOR.CREAM);
  sh.getRange('E3').setValue('← ★1〜3は毎日、☆4〜8は週2〜3回が目安（T7の頻度）。チェックで有効/無効を切り替え')
    .setFontColor('#666666');

  var ngStart = 4 + rows.length + 2;
  sh.getRange(ngStart, 1).setValue('即×ワード（タイトルにあれば自動で「除外」）').setFontWeight('bold').setBackground(COLOR.HEADER);
  var ngRows = DEFAULT_NG_WORDS.map(function (w) { return [w]; });
  sh.getRange(ngStart + 1, 1, ngRows.length, 1).setValues(ngRows).setBackground(COLOR.RED_LIGHT);

  var okStart = ngStart + 1 + ngRows.length + 2;
  sh.getRange(okStart, 3).setValue('◎ワード（あれば「要チェック◎」フラグ）').setFontWeight('bold').setBackground(COLOR.HEADER);
  var okRows = DEFAULT_OK_WORDS.map(function (w) { return [w]; });
  sh.getRange(okStart + 1, 3, okRows.length, 1).setValues(okRows).setBackground(COLOR.GREEN_LIGHT);

  sh.setColumnWidth(1, 70).setColumnWidth(2, 220).setColumnWidth(3, 220).setColumnWidth(4, 560);
  sh.setFrozenRows(3);

  // 位置を後の読み取り関数と合わせるため、行番号をノートに記録
  sh.getRange('A1').setNote('sourcesStart=4,sourcesCount=' + rows.length +
    ';ngStart=' + (ngStart + 1) + ',ngCount=' + ngRows.length +
    ';okStart=' + (okStart + 1) + ',okCount=' + okRows.length);
}

function readSources_(config) {
  var note = config.getRange('A1').getNote();
  var meta = parseMeta_(note);
  var values = config.getRange(meta.sourcesStart, 1, meta.sourcesCount, 4).getValues();
  return values.map(function (r) {
    return { on: !!r[0], label: String(r[1]), keyword: String(r[2]), url: String(r[3]) };
  });
}

function readWordList_(config, kind) {
  var note = config.getRange('A1').getNote();
  var meta = parseMeta_(note);
  var col = (kind === 'NG') ? 1 : 3;
  var start = (kind === 'NG') ? meta.ngStart : meta.okStart;
  var count = (kind === 'NG') ? meta.ngCount : meta.okCount;
  var values = config.getRange(start, col, count, 1).getValues();
  return values.map(function (r) { return String(r[0]).trim(); }).filter(function (w) { return w !== ''; });
}

function parseMeta_(note) {
  var meta = {};
  (note || '').split(';').forEach(function (part) {
    part.split(',').forEach(function (kv) {
      var pair = kv.split('=');
      if (pair.length === 2) meta[pair[0].trim()] = Number(pair[1]);
    });
  });
  return meta;
}

function buildLog_(sh) {
  sh.setTabColor('#4285f4');
  sh.getRange('A1:J1').setValues([[
    '取得日時', '検索キーワード', '判定', '理由', 'タイトル', '報酬', '応募数', 'URL', '取込方法', 'メモ'
  ]]).setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 160);
  sh.setColumnWidth(5, 320);
  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 70);
  sh.setColumnWidth(8, 260);
  sh.setColumnWidth(9, 90);
  sh.setColumnWidth(10, 160);

  var dataRange = sh.getRange(2, 1, 2000, 10);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="要チェック◎"').setBackground(COLOR.GREEN_LIGHT).setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="除外"').setBackground('#eeeeee').setFontColor('#999999').setRanges([dataRange]).build()
  ]);
}

function buildRunlog_(sh) {
  sh.setTabColor('#ea4335');
  sh.getRange('A1:E1').setValues([['実行日時', '検索ラベル', 'HTTPステータス', '結果', '詳細（エラー内容など）']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 180); sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 90); sh.setColumnWidth(5, 500);
}

function buildPaste_(sh) {
  sh.setTabColor('#f4b400');
  sh.getRange('A1:E1').merge().setValue(
    '🟡 3行目から、気になった案件の「タイトル」「URL」（分かれば「報酬」）を貼り付けて、メニュー「📋 手動貼付シートを判定する」を実行'
  ).setBackground(COLOR.YEL_LIGHT).setFontWeight('bold').setWrap(true);
  sh.getRange('A2:E2').setValues([['タイトル', 'URL', '報酬（あれば）', '判定（自動入力）', '理由（自動入力）']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.getRange(3, 1, 200, 3).setBackground(COLOR.CREAM);
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 320); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 200);
}

function buildRaw_(sh) {
  sh.setTabColor('#666666');
  sh.getRange('A1:C1').setValues([['保存日時', '検索ラベル', '取得HTML（先頭部分・診断用）']])
    .setFontWeight('bold').setBackground(COLOR.HEADER);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 180); sh.setColumnWidth(3, 900);
}

function writeRunlog_(ss, label, result) {
  var sh = ss.getSheetByName(SHEET.RUNLOG);
  if (!sh) return;
  sh.appendRow([
    new Date(), label, result.code || '', result.ok ? 'OK' : 'NG', result.error || ''
  ]);
}

function saveRawHtml_(ss, label, html) {
  var sh = ss.getSheetByName(SHEET.RAWHTML);
  if (!sh) return;
  var snippet = String(html || '').substring(0, 45000); // セルの文字数上限(5万)対策
  sh.appendRow([new Date(), label, snippet]);
}

function readExistingUrls_(log) {
  var lastRow = log.getLastRow();
  var map = {};
  if (lastRow < 2) return map;
  var urls = log.getRange(2, 8, lastRow - 1, 1).getValues(); // H列=URL
  urls.forEach(function (r) { if (r[0]) map[String(r[0])] = true; });
  return map;
}

function appendLogRow_(log, keyword, job, cls, method) {
  log.appendRow([
    new Date(), keyword, cls.status, cls.reason, job.title, job.pay, job.applicants, job.url, method, ''
  ]);
}

/* --------------------------------------------------------------------
 * 「はじめに」シート
 * ------------------------------------------------------------------ */
function buildGuide_(sh) {
  var lines = [
    '🔍 毎朝の案件パトロール自動化（クラウドワークス新着チェック）',
    'T7「毎朝5分の手動パトロール」を自動化する試みです。まず必ず「🧪 接続テスト」から始めてください。',
    '',
    '■ ⚠️ 正直な注記（必ずお読みください）',
    'この仕組みを作った作業環境からは crowdworks.jp に一切接続できず（403で拒否）、実際のページを見て' +
      '検証することができませんでした。さらに別のネットワーク経路で robots.txt（本来は誰でも読める公開' +
      'ファイル）を読みにいっても同じく403だったため、クラウドワークスは機械的なアクセス全般を入口で' +
      'ブロックしている可能性があります。',
    'Google Apps Script の UrlFetchApp は Google のサーバーから実行されるため、上と同じ制限を受けない' +
      '可能性はありますが、確実に動くという保証はできません。「🧪 接続テスト」を必ず最初に実行し、' +
      'HTTPステータスが200で、かつ案件が実際にパースできているかを確認してください。',
    'なお、クラウドワークスの新着案件RSS配信は2018年12月に終了しており、公開APIも見当たりません。' +
      '自動取得がブロックされていた場合、代替の自動取得手段は現時点でありません。',
    '',
    '■ ブロックされていた場合の次善策：「判断」だけ自動化する',
    '自動取得ができなくても、「手動貼付判定」シートを使えば、T7のとおり社長がブックマークを目で見て' +
      '巡回したあと、気になった案件のタイトル・URLだけ貼り付け→メニュー実行、で即×判定・◎フラグを' +
      '自動で付けられます。「取得」は無理でも「判断」の自動化で、5分の作業を短縮できます。',
    '',
    '■ 初回セットアップ手順',
    '① メニュー「🔍案件パトロール」→「🛠 初期セットアップ」',
    '② メニュー「🧪 接続テスト（1件だけ・まずこれ）」を実行し、結果を確認する',
    '　・HTTPステータス200＋案件が検出できた → ③へ進んでOK',
    '　・403等でブロック／0件パース → 「実行ログ」「生HTML保存」シートの内容をハックに共有してください',
    '③ 問題なければ「⏰ 毎朝7時の自動実行をON」でトリガー登録（承認画面が出ます。「詳細」→「移動」で進めてください）',
    '',
    '■ 毎朝の確認（自動取得がうまくいっている場合）',
    '「案件ログ」シートの新着行（判定＝要チェック◎）を上から確認するだけ。',
    '除外行は薄いグレーになります（消してはいません＝あとで見返せる記録として残します）。',
    '有望案件を見つけたら、これまでどおりT9のチェックリストで精査依頼をどうぞ。',
    '',
    '■ 設定のカスタマイズ',
    '「設定」シートで、検索キーワードのON/OFF・即×ワード・◎ワードを自由に追加・編集できます。',
    '',
    '■ 既知の制約',
    '・報酬・応募数はページの表記ゆれにより取得できないことがあります（空欄でも判定は動きます）',
    '・タイトルの抽出は簡易ルールのため、まれに崩れることがあります（URLで実際のページを確認してください）',
    '・クラウドワークス側の仕様変更で、ある日から急にパースできなくなる可能性があります（生HTML保存シートが' +
      'その診断材料になります）'
  ];
  var values = lines.map(function (t) { return [t]; });
  sh.getRange(1, 1, values.length, 1).setValues(values).setWrap(true);
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].indexOf('■') === 0) sh.getRange(i + 1, 1).setFontWeight('bold').setFontSize(11).setBackground(COLOR.HEADER);
  }
  sh.setColumnWidth(1, 900);
  sh.setHiddenGridlines(true);
}
