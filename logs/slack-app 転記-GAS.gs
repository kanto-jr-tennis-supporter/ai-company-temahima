/**
 * ============================================================
 *  Slackアポ報告 → スプレッドシート C〜H列 自動転記
 *  テマヒマ・ラボ / add one 案件（担当：ハック）
 * ------------------------------------------------------------
 *  ● やること
 *    Slackの「外部とのつながり」系3会話に、アポインターさんから
 *    上がってくる「アポ報告メッセージ」を、指定スプレッドシートの
 *    C〜H列（A・B列は社長管理用なので触らない）へ自動で追記する。
 *
 *  ● 動き方（本命ルート＝常時自動）
 *    Slack Events API → このGASの doPost(Webアプリ) が受け取る
 *      → 報告文をパース（複数ブロック対応）→ シートのC〜Hへ append
 *
 *  ● 手動ルート（Slack設定ゼロでもすぐ使える縮小版）
 *    スプレッドシートのメニュー「▶ アポ転記」→「報告文を貼り付けて転記」
 *      → 出てきた枠に報告メッセージを丸ごと貼り付け → OK で転記
 *
 *  ● 設定（Script Properties / プロジェクトの設定 → スクリプト プロパティ）
 *    SHEET_ID           転記先スプレッドシートのID（未設定なら下の既定を使用）
 *    TARGET_GID         対象タブのgid（未設定なら下の既定を使用）
 *    ALLOWED_CHANNELS   対象Slack会話IDをカンマ区切り（空なら全会話を許可）
 *    SLACK_VERIFY_TOKEN （任意）WebアプリURLの ?token= と照合する合言葉
 *    ERROR_MAIL         （任意）エラー時に通知するメールアドレス
 * ============================================================
 */

// ===== 既定値（Script Propertiesが未設定のときに使う） =====
var DEFAULT_SHEET_ID = '19sk5tKvAoC8bEBHAVVbL5AdV_sD6_9d_crRgCd68tIg';
var DEFAULT_TARGET_GID = 1534991540;

// C〜H の列番号（A=1, B=2, C=3 ... H=8）
var COL_START = 3; // C列
var COL_COUNT = 6; // C〜H の6列

// ------------------------------------------------------------
// 設定の読み出し
// ------------------------------------------------------------
function prop_(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === '') ? fallback : v;
}

function getConfig_() {
  return {
    sheetId: prop_('SHEET_ID', DEFAULT_SHEET_ID),
    targetGid: Number(prop_('TARGET_GID', DEFAULT_TARGET_GID)),
    allowedChannels: String(prop_('ALLOWED_CHANNELS', ''))
      .split(',').map(function (s) { return s.trim(); }).filter(String),
    verifyToken: prop_('SLACK_VERIFY_TOKEN', ''),
    errorMail: prop_('ERROR_MAIL', '')
  };
}

// ------------------------------------------------------------
// Slack Webアプリ受け口
// ------------------------------------------------------------
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
    var body = raw ? JSON.parse(raw) : {};

    // 1) SlackのURL検証（Event Subscriptions登録時の初回だけ来る）
    if (body.type === 'url_verification') {
      return ContentService.createTextOutput(body.challenge)
        .setMimeType(ContentService.MimeType.TEXT);
    }

    var cfg = getConfig_();

    // 2) 簡易認証（任意）：URLの ?token=... が合言葉と一致しなければ黙って無視
    //    ※ GASのdoPostはHTTPヘッダを読めないため、Slackの正式なHMAC署名検証は
    //      できない。代わりにこの合言葉方式で「なりすまし投稿」を実用上ふせぐ。
    if (cfg.verifyToken) {
      var token = (e && e.parameter && e.parameter.token) ? e.parameter.token : '';
      if (token !== cfg.verifyToken) return okJson_();
    }

    // 3) イベント本体
    if (body.type === 'event_callback' && body.event) {
      handleEvent_(body, cfg);
    }
    return okJson_();
  } catch (err) {
    notifyError_(err);
    // Slackへは常に200を返す（エラー返却するとSlackが何度もリトライして二重転記の原因になる）
    return okJson_();
  }
}

function okJson_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleEvent_(body, cfg) {
  var ev = body.event;
  if (!ev || ev.type !== 'message') return;

  // 通常の人間の投稿だけを対象に。編集・削除・ボット投稿・スレッド通知等は無視
  if (ev.subtype) return;      // subtypeが付くもの（編集/削除/join等）は対象外
  if (ev.bot_id) return;       // ボット投稿は対象外

  // 対象会話フィルタ（ALLOWED_CHANNELSが空なら全許可）
  if (cfg.allowedChannels.length && cfg.allowedChannels.indexOf(ev.channel) === -1) return;

  // 二重転記ふせぎ：同じイベントは10分間キャッシュで弾く（Slackのリトライ対策）
  var cache = CacheService.getScriptCache();
  var key = 'ev_' + (body.event_id || ev.channel + '_' + ev.ts);
  if (cache.get(key)) return;
  cache.put(key, '1', 600);

  var rows = parseReport_(ev.text || '');
  if (!rows.length) return; // アポ報告の構造ブロックが無い普通の会話は転記しない

  appendRows_(rows, cfg);
}

// ------------------------------------------------------------
// 報告文パーサ（複数ブロック対応・全角半角ゆれ・tel/mailtoリンク剥がし）
// ------------------------------------------------------------

/**
 * 1メッセージ → 行の配列。「⭐アポ獲得」「［送付先］」ブロックごとに1行。
 * 返り値: [{ C,D,E,F,G,H }, ...]
 */
function parseReport_(text) {
  if (!text) return [];
  text = unwrapSlackLinks_(text);
  var lines = text.replace(/\r\n?/g, '\n').split('\n');

  var blocks = [];
  var cur = null;
  for (var i = 0; i < lines.length; i++) {
    if (isHeader_(lines[i])) { cur = []; blocks.push(cur); continue; }
    if (cur) cur.push(lines[i]);
  }
  var rows = [];
  for (var b = 0; b < blocks.length; b++) {
    var r = parseBlock_(blocks[b]);
    if (r.C || r.D || r.E || r.F || r.G || r.H) rows.push(r);
  }
  return rows;
}

/**
 * ブロックの見出し行か？（「⭐アポ獲得」「⭐️アポ獲得」「★ アポ獲得」「［送付先］」「[送付先]」など揺れを許容）
 * 記号・括弧・空白・コロンを取り除いた芯が「アポ獲得」か「送付先」ならヘッダとみなす。
 */
function isHeader_(line) {
  var t = String(line).replace(/[\s　［］\[\]（）()★☆⭐️:：]/g, '');
  return t === 'アポ獲得' || t === '送付先';
}

/**
 * 1ブロック（見出しの次行〜次見出し手前）を1行の {C..H} に変換。
 */
function parseBlock_(blockLines) {
  var row = { C: '', D: '', E: '', F: '', G: '', H: '' };
  // ラベル → 転記先列。keysのどれかに前方一致したら採用（最初に埋まった値を優先）
  var LABELS = [
    { keys: ['会社名', '商号', '社名'], col: 'C' },   // C 商号
    { keys: ['担当者', '代表者名', '代表者', '代表'], col: 'D' }, // D 代表者名
    { keys: ['住所', '所在地'], col: 'E' },            // E 住所
    { keys: ['電話番号', '電話', 'TEL', 'Tel'], col: 'F' }, // F 電話番号
    { keys: ['アドレス', 'メールアドレス', 'メール', 'Email', 'E-mail', 'mail'], col: 'G' } // G メール
    // 「アポ日時：」など、上記に無いラベルは転記しない
  ];

  for (var i = 0; i < blockLines.length; i++) {
    var line = String(blockLines[i]).trim();
    if (!line) continue;

    // ※以降の一文 → H（アポ時メモ）。複数の※があれば連結
    if (line.charAt(0) === '※' /* ※ */) {
      var memo = line.replace(/^※\s*/, '').trim();
      row.H = row.H ? (row.H + ' ' + memo) : memo;
      continue;
    }

    // 「ラベル：値」を分解（全角：・半角:どちらも可）
    var m = line.match(/^([^：:]+)[：:]\s*(.*)$/);
    if (!m) continue;
    var label = m[1].replace(/[\s　]/g, '');
    var value = cleanValue_(m[2]);

    for (var j = 0; j < LABELS.length; j++) {
      var def = LABELS[j];
      var hit = def.keys.some(function (k) { return label === k || label.indexOf(k) === 0; });
      if (hit) {
        if (!row[def.col]) row[def.col] = value;
        break;
      }
    }
  }
  return row;
}

/** Slackのリンク記法を「表示テキスト」だけに戻す。<tel:..|042..> → 042.. / <mailto:a@b> → a@b */
function unwrapSlackLinks_(text) {
  return String(text)
    .replace(/<(?:tel|mailto|https?):[^|>]*\|([^>]*)>/gi, '$1') // 表示付き
    .replace(/<(?:tel|mailto):([^|>]*)>/gi, '$1')              // 表示なし tel/mailto
    .replace(/<(https?:[^|>]*)>/gi, '$1');                     // 表示なし URL
}

/** 値の掃除：全角スペース→半角、前後トリム、末尾コロン等の除去 */
function cleanValue_(v) {
  return String(v || '').replace(/　/g, ' ').trim();
}

// ------------------------------------------------------------
// シートへの追記（C〜Hのみ。A・B列には一切触れない）
// ------------------------------------------------------------
function getTargetSheet_(cfg) {
  var ss = SpreadsheetApp.openById(cfg.sheetId);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === cfg.targetGid) return sheets[i];
  }
  throw new Error('gid=' + cfg.targetGid + ' のタブが見つかりません（SHEET_ID/TARGET_GIDを確認）');
}

function appendRows_(rows, cfg) {
  cfg = cfg || getConfig_();
  var sh = getTargetSheet_(cfg);

  // ロック：同時実行での行かぶりを防ぐ
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var startRow = sh.getLastRow() + 1; // 既存最終行の次から
    var values = rows.map(function (r) { return [r.C, r.D, r.E, r.F, r.G, r.H]; });
    sh.getRange(startRow, COL_START, values.length, COL_COUNT).setValues(values);
    return { startRow: startRow, count: values.length };
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// 手動ルート：スプレッドシートのメニューから貼り付け転記（Slack設定ゼロで使える縮小版）
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('▶ アポ転記')
    .addItem('報告文を貼り付けて転記', 'manualPaste_')
    .addItem('サンプルで動作テスト', 'runSelfTest_')
    .addToUi();
}

function manualPaste_() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt(
    'アポ報告メッセージを貼り付け',
    'Slackの報告メッセージを丸ごと貼り付けて［OK］を押してください（⭐アポ獲得／［送付先］ブロックを自動で拾います）。',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var rows = parseReport_(res.getResponseText());
  if (!rows.length) {
    ui.alert('転記できるブロックが見つかりませんでした。「⭐アポ獲得」「［送付先］」の見出しが含まれているか確認してください。');
    return;
  }
  var r = appendRows_(rows);
  ui.alert(rows.length + '件を ' + r.startRow + ' 行目から C〜H に転記しました。');
}

// ------------------------------------------------------------
// 動作テスト（スプレッドシートに書かず、ログで抽出結果を確認）
// ------------------------------------------------------------
function runSelfTest_() {
  var sample =
    '午後は1件、来週のアポが取れましたのでご確認お願いします。\n' +
    '\n' +
    'toG（東京都)\n' +
    '60min（12:30ー13:30）\n' +
    'コール：30件 / 主権：3件 / 断り：6件 / アポ：1件 / 資料送付：1件\n' +
    '\n' +
    '⭐アポ獲得\n' +
    '会社名：（株）未来想造プログレス\n' +
    '住所：東京都武蔵村山市大南２－５３－４\n' +
    '電話番号：<tel:+81428435862|042-843-5862>\n' +
    '担当者：小林代表\n' +
    'アポ日時：9/29の14時ズーム\n' +
    'アドレス：<mailto:riko@m-progress.net|riko@m-progress.net>\n' +
    '※公共は東京都で下請け経験あり。業務エリアは東京都内。5代くらいの女性代表者様。\n' +
    '\n' +
    '［送付先］\n' +
    '会社名：（株）佐藤電気\n' +
    '住所：東京都武蔵村山市伊奈平１－８－１９\n' +
    '電話番号：09053256915\n' +
    '担当者：代表宛\n' +
    'アドレス：<mailto:uzd735310@nifty.com>\n' +
    '※メールでの送付。興味はあるが基本現場なので19時くらいしか難しい。';

  var rows = parseReport_(sample);
  Logger.log('抽出件数: ' + rows.length);
  Logger.log(JSON.stringify(rows, null, 2));
  // 期待：2件。1件目=（株）未来想造プログレス / 小林代表 / …、2件目=（株）佐藤電気 / 代表宛 / …
  return rows;
}

// ------------------------------------------------------------
// エラー通知
// ------------------------------------------------------------
function notifyError_(err) {
  try {
    var mail = getConfig_().errorMail;
    if (!mail) return;
    MailApp.sendEmail(mail, '[アポ転記GAS] エラー', String(err && err.stack ? err.stack : err));
  } catch (e) { /* 通知失敗は握りつぶす */ }
}
