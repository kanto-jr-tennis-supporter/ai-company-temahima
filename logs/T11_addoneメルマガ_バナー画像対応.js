/****************************************************
 * 再アプローチリスト直接参照メルマガ送信 完全版（バナー画像・フッター画像・宛名修正対応）
 * add one様向け／テマヒマ・ラボ ハック作成 / 2026-08-05
 *
 * 【変更点（元スクリプトからの差分）】
 * ・メルマガ用シートに H2：バナー画像URL（本文の上）を追加
 * ・メルマガ用シートに I2：フッター画像URL（本文の下）を追加
 *   （どちらもGoogleドライブの共有リンクをそのまま貼ってOK。中でファイルIDを抽出して変換します）
 * ・本文をプレーンテキストではなく HTML メールとして送信するように変更
 *   → 本文の一番上にバナー画像、一番下にフッター画像が表示される
 * ・plainBody（画像非表示の環境向けの文字だけ版）は今まで通り自動生成されるので、
 *   万一画像が読み込めない受信環境でも本文は読める
 * ・宛名の「様」が二重に付く不具合を修正（addSama_ が末尾の「様」を必ず1個に正規化）
 * ・{{代表者名}} は「代表者名 様」＋改行＋「ご担当者様」の2行に自動展開されるように変更
 *   （「ご担当者様」はスプレッドシートには持たず、スクリプト側の固定文言として追加）
 * ・連続した「様」（例：「ご担当者様様」）を自動で1個に畳み込む保険を追加
 *   （collapseDuplicateHonorific_）。ただし根本原因はA2側の手入力なので、下記も確認すること
 * ・本文中の装飾記法を追加（applyHighlightMarkup_）：
 *     ****文章**** → 太字のみ（色もサイズも変えない）
 *     ***文章***  → オレンジ・太字・大きめ（色もサイズも両方変えたいとき）
 *     **文章**    → オレンジ・太字（サイズは本文と同じ・色だけ変えたいとき）
 *     *文章*     → 色はそのまま（黒字）・太字・大きめ（サイズだけ大きくしたいとき）
 *   文字サイズは applyHighlightMarkup_ 内の font-size（例：1.6em）を変えるだけで調整可
 * ・本文中に [[ボタンの文字|URL]] と書くと、オレンジ背景の目立つボタンリンクに変換される
 *   （applyButtonMarkup_）。例：[[お申し込みはこちら|https://forms.gle/xxxxxxxx]]
 *   揃え位置は左揃え（デフォルト）。applyButtonMarkup_ 内の text-align を変えれば調整可
 *
 * 【事前準備・注意点】
 * ・バナー・フッターどちらの画像ファイルも、Googleドライブで
 *   「リンクを知っている全員が閲覧可」に共有設定してください。
 *   （限定公開のままだと相手のメールで画像が表示されません＝赤い×アイコンになります）
 * ・メルマガ用シートA2の本文で、{{代表者名}} の前後に「様」や「ご担当者様」を手入力していないか
 *   必ず確認してください。プレースホルダーだけを単独で置き、前後には何も書かないのが正解です
 *   （敬称も「ご担当者様」も、このスクリプトが自動で付けます）
 *
 * 再アプローチリスト：
 * C列：商号
 * D列：代表者名
 * G列：メールアドレス
 * K/P/U/Z列：「アポ」を含む場合は除外
 *
 * メルマガ用シート：
 * A2：本文
 * B2：添付ファイルURL
 * C2：件名
 * D2：テストモード チェックボックス
 * E2：テスト送信先
 * F2：テスト用 商号
 * G2：テスト用 代表者名
 * H2：バナー画像URL（本文の上・Googleドライブの共有リンクを貼る）
 * I2：フッター画像URL（本文の下・Googleドライブの共有リンクを貼る）
 ****************************************************/

const SHEET_REAPPROACH = "再アプローチリスト";
const SHEET_MAIL = "メルマガ用シート";
const SHEET_SEND_LOG = "メルマガ送信ログ";
const SHEET_FAILED_LOG = "メルマガ送信失敗ログ";

/**
 * メニュー作成
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📩 メルマガ")
    .addItem("① 送信対象件数を確認", "checkMailTargets")
    .addItem("② テスト送信", "sendTestMail")
    .addItem("③ 本送信", "sendMainMail")
    .addItem("④ アポ除外チェック", "checkExcludedApoRows")
    .addToUi();
}

/**
 * メルマガ用シート設定取得
 */
function getMailSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAIL);
  if (!sheet) throw new Error("メルマガ用シートが見つかりません");

  return {
    body: String(sheet.getRange("A2").getValue() || ""),
    attachmentUrl: String(sheet.getRange("B2").getValue() || ""),
    subject: String(sheet.getRange("C2").getValue() || ""),
    testMode: sheet.getRange("D2").getValue() === true,
    testTo: String(sheet.getRange("E2").getValue() || "").trim(),
    testCompany: String(sheet.getRange("F2").getValue() || "").trim(),
    testRepresentative: String(sheet.getRange("G2").getValue() || "").trim(),
    bannerImageUrl: String(sheet.getRange("H2").getValue() || "").trim(),
    footerImageUrl: String(sheet.getRange("I2").getValue() || "").trim(),
    sheet,
  };
}

/**
 * 送信対象取得
 */
function getMailTargets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!sheet) throw new Error("再アプローチリストが見つかりません");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 26).getDisplayValues();
  const targets = [];

  data.forEach((row, i) => {
    const rowNumber = i + 2;
    const company = String(row[2] || "").trim();        // C
    const representative = String(row[3] || "").trim(); // D
    const email = String(row[6] || "").trim();          // G

    if (!email) return;

    const statusText = [
      row[10], // K
      row[15], // P
      row[20], // U
      row[25], // Z
    ].join(" ");

    if (statusText.includes("アポ")) return;

    targets.push({
      rowNumber,
      company,
      representative,
      email,
    });
  });

  return targets;
}

/**
 * 代表者名に「 様」を付ける
 * 末尾に「様」が（スペース有無や複数回に関わらず）既に付いていたら、
 * 一旦すべて取り除いてから必ず1回だけ付け直す（二重敬称の防止）
 */
function addSama_(name) {
  let n = String(name || "").trim();
  if (!n) return "";
  while (/[\s　]*様$/.test(n)) {
    n = n.replace(/[\s　]*様$/, "").trim();
  }
  if (!n) return "";
  return n + " 様";
}

/**
 * 「様」が連続している箇所（間に空白・改行があってもOK）を1個に畳み込む。
 * 例：「ご担当者様様」「ご担当者様 様」「山田太郎 様\n様」→ すべて「様」1個に。
 * スプレッドシート側に手入力の「様」が残っていた場合の保険。
 */
function collapseDuplicateHonorific_(text) {
  return String(text || "").replace(/様(?:[\s　]*様)+/g, "様");
}

/**
 * 差し込み
 * {{代表者名}} は「代表者名 様」＋改行＋「ご担当者様」の2行に自動展開される。
 * 「ご担当者様」はスプレッドシート側には持たず、ここで固定文言として付与している。
 * ※ メルマガ用シートA2の本文側に、手入力の「様」や「ご担当者様」が別途残っていると
 *   二重表示の原因になるので、{{代表者名}} の前後には何も手入力しないこと。
 *   （念のため、連続した「様」は collapseDuplicateHonorific_ で自動的に1個へ畳み込む）
 */
function applyTemplate_(text, target) {
  const representativeBlock = addSama_(target.representative) + "\nご担当者様";

  const merged = String(text || "")
    .replaceAll("{{商号}}", target.company || "")
    .replaceAll("{{ 商号 }}", target.company || "")
    .replaceAll("{{会社名}}", target.company || "")
    .replaceAll("{{ 会社名 }}", target.company || "")
    .replaceAll("{{代表者名}}", representativeBlock)
    .replaceAll("{{ 代表者名 }}", representativeBlock);

  return collapseDuplicateHonorific_(merged);
}

/**
 * Googleドライブの共有リンク／直リンク／IDそのもの、どの形式でもファイルIDを取り出す
 * 対応例：
 *   https://drive.google.com/file/d/XXXX/view?usp=sharing
 *   https://drive.google.com/uc?export=view&id=XXXX
 *   XXXX（IDだけ）
 */
function extractDriveFileId_(url) {
  if (!url) return "";
  const text = String(url).trim();

  let m = text.match(/\/d\/([-\w]{20,})/);
  if (m) return m[1];

  m = text.match(/[?&]id=([-\w]{20,})/);
  if (m) return m[1];

  m = text.match(/^([-\w]{20,})$/);
  if (m) return m[1];

  return "";
}

/**
 * バナー画像URL（ドライブ共有リンク）を、メール本文に埋め込める直リンクに変換
 * ※ 変換前に、対象ファイルを「リンクを知っている全員が閲覧可」に設定しておくこと
 */
function buildBannerImageUrl_(driveUrl) {
  const id = extractDriveFileId_(driveUrl);
  if (!id) return "";
  return "https://drive.google.com/uc?export=view&id=" + id;
}

/**
 * HTML特殊文字のエスケープ（本文中の &, <, > 等がタグ崩れしないように）
 */
function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * 本文中の記号で囲まれた部分を装飾する。
 * ・****文章****（アスタリスク4つ）→ 太字のみ（色もサイズも変えず、太字だけにしたいとき）
 * ・***文章***（アスタリスク3つ）  → オレンジ・太字・大きめ（色もサイズも両方変えたいとき）
 * ・**文章**（アスタリスク2つ）    → オレンジ・太字（色だけ変えたいとき。文字サイズは本文と同じ）
 * ・*文章*（アスタリスク1つ）      → 色はそのまま（黒字）・太字・大きめ（サイズだけ大きくしたいとき）
 * 例：****建設業許可番号 第12345号****（太字のみ）
 *     ***最低利用6ヶ月から始められるキャンペーンを実施いたします。***（オレンジ・大きめ両方）
 *     **建設会社様向け**（オレンジのみ）　　⭐*建設会社様向け*⭐（黒字のまま大きく）
 * ※ **** → *** → ** → * の順に処理するので、混ぜて使っても崩れない
 *   （逆順で処理すると内側の * が外側の記号に誤って食われてしまうため、必ずこの順番で処理する）
 * ※ 文字サイズは font-size の倍率（例：1.6em）を変えるだけで簡単に調整できる
 */
function applyHighlightMarkup_(escapedText) {
  let result = escapedText.replace(
    /\*\*\*\*([\s\S]+?)\*\*\*\*/g,
    '<span style="font-weight:bold;">$1</span>'
  );

  result = result.replace(
    /\*\*\*([\s\S]+?)\*\*\*/g,
    '<span style="color:#e2551c;font-weight:bold;font-size:1.6em;line-height:1.4;">$1</span>'
  );

  result = result.replace(
    /\*\*([\s\S]+?)\*\*/g,
    '<span style="color:#e2551c;font-weight:bold;">$1</span>'
  );

  result = result.replace(
    /\*([\s\S]+?)\*/g,
    '<span style="font-weight:bold;font-size:1.6em;line-height:1.4;">$1</span>'
  );

  return result;
}

/**
 * 本文中の [[ボタンの文字|URL]] を、オレンジ背景の目立つボタンリンクに変換する。
 * 例：[[お申し込みはこちら|https://forms.gle/xxxxxxxx]]
 * 前後に余白が入ったブロックとして表示される（フォームへの誘導などに）。
 * 揃え位置は下の text-align（現在は "left"）を "center" や "right" に変えれば調整できる。
 * ※ Outlookデスクトップ版など一部メーラーでは角丸（border-radius）が四角のまま
 *   表示されることがあるが、ボタンとしての見た目・クリックは問題なく機能する。
 */
function applyButtonMarkup_(escapedText) {
  return escapedText.replace(
    /\[\[([^|\]]+)\|([^\]]+)\]\]/g,
    (match, label, url) => {
      const safeUrl = url.trim();
      const safeLabel = label.trim();
      return (
        `<div style="text-align:left;margin:24px 0;">` +
        `<a href="${safeUrl}" target="_blank" style="display:inline-block;background-color:#e2551c;` +
        `color:#ffffff;font-weight:bold;font-size:16px;padding:14px 36px;border-radius:6px;` +
        `text-decoration:none;">${safeLabel}</a>` +
        `</div>`
      );
    }
  );
}

/**
 * プレーン本文＋バナー画像URL＋フッター画像URLから、HTMLメール本文を組み立てる
 */
function buildHtmlBody_(plainBody, bannerImageUrl, footerImageUrl) {
  const bannerHtml = bannerImageUrl
    ? `<img src="${bannerImageUrl}" alt="バナー" style="max-width:600px;width:100%;height:auto;display:block;margin:0 0 16px 0;border:0;">`
    : "";

  const footerHtml = footerImageUrl
    ? `<img src="${footerImageUrl}" alt="フッター" style="max-width:600px;width:100%;height:auto;display:block;margin:16px 0 0 0;border:0;">`
    : "";

  let textHtml = escapeHtml_(plainBody);
  textHtml = applyButtonMarkup_(textHtml);
  textHtml = applyHighlightMarkup_(textHtml);
  textHtml = textHtml.replace(/\n/g, "<br>");

  return (
    `<div style="font-family:'Hiragino Kaku Gothic ProN','Meiryo',sans-serif;` +
    `font-size:14px;line-height:1.8;color:#333333;max-width:600px;">` +
    `${bannerHtml}${textHtml}${footerHtml}</div>`
  );
}

/**
 * 添付ファイル取得
 */
function getAttachments_(urlText) {
  if (!urlText) return [];

  const text = String(urlText);
  const ids = text.match(/[-\w]{25,}/g);

  if (!ids || ids.length === 0) return [];

  const files = [];

  ids.forEach(id => {
    try {
      const file = DriveApp.getFileById(id);
      files.push(file.getBlob());
    } catch (e) {
      Logger.log("添付ファイル取得失敗: " + id + " / " + e.message);
    }
  });

  return files;
}

/**
 * 送信元アカウント取得
 */
function getSenderEmail_() {
  return (
    Session.getActiveUser().getEmail() ||
    Session.getEffectiveUser().getEmail() ||
    "取得できませんでした"
  );
}

/**
 * 送信対象件数確認
 */
function checkMailTargets() {
  const ui = SpreadsheetApp.getUi();

  try {
    const targets = getMailTargets_();

    if (targets.length === 0) {
      ui.alert("送信対象件数：0件");
      return;
    }

    const sample = targets[0];

    ui.alert(
      "送信対象確認",
      `送信対象件数：${targets.length}件\n\n` +
      `【1件目サンプル】\n` +
      `商号：${sample.company}\n` +
      `代表者名：${addSama_(sample.representative)}\n` +
      `メール：${sample.email}`,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert("エラー：" + e.message);
  }
}

/**
 * テスト送信
 */
function sendTestMail() {
  const ui = SpreadsheetApp.getUi();

  try {
    const settings = getMailSettings_();

    if (!settings.testMode) {
      ui.alert("D2のテストモードにチェックが入っていません。");
      return;
    }

    if (!settings.testTo) {
      ui.alert("E2にテスト送信先を入力してください。");
      return;
    }

    if (!settings.subject) {
      ui.alert("C2に件名を入力してください。");
      return;
    }

    if (!settings.body) {
      ui.alert("A2に本文を入力してください。");
      return;
    }

    const senderEmail = getSenderEmail_();

    const testTarget = {
      company: settings.testCompany,
      representative: settings.testRepresentative,
      email: settings.testTo,
    };

    const subject = applyTemplate_(settings.subject, testTarget);
    const plainBody = applyTemplate_(settings.body, testTarget);
    const bannerUrl = buildBannerImageUrl_(settings.bannerImageUrl);
    const footerUrl = buildBannerImageUrl_(settings.footerImageUrl);
    const htmlBody = buildHtmlBody_(plainBody, bannerUrl, footerUrl);
    const attachments = getAttachments_(settings.attachmentUrl);

    const confirm = ui.alert(
      "テスト送信確認",
      `現在ログインしているアカウント：\n${senderEmail}\n\n` +
      `このアカウントから送信されます。\n\n` +
      `送信先：${settings.testTo}\n` +
      `商号：${testTarget.company}\n` +
      `代表者名：${addSama_(testTarget.representative)}\n` +
      `件名：${subject}\n` +
      `バナー画像：${bannerUrl ? "あり" : "なし（H2が未設定 or URLを認識できません）"}\n` +
      `フッター画像：${footerUrl ? "あり" : "なし（I2が未設定 or URLを認識できません）"}\n` +
      `添付ファイル数：${attachments.length}件\n\n` +
      `この内容でテスト送信しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    GmailApp.sendEmail(settings.testTo, subject, plainBody, {
      htmlBody: htmlBody,
      attachments,
    });

    ui.alert("テスト送信が完了しました。実際の見え方（画像が出るか）を必ず自分の受信箱で確認してください。");

  } catch (e) {
    ui.alert("エラー：" + e.message);
  }
}

/**
 * 本送信
 */
function sendMainMail() {
  const ui = SpreadsheetApp.getUi();

  try {
    const settings = getMailSettings_();

    if (settings.testMode) {
      ui.alert("D2がテストモードのままです。本送信する場合はチェックを外してください。");
      return;
    }

    if (!settings.subject) {
      ui.alert("C2に件名を入力してください。");
      settings.sheet.getRange("D2").setValue(true);
      return;
    }

    if (!settings.body) {
      ui.alert("A2に本文を入力してください。");
      settings.sheet.getRange("D2").setValue(true);
      return;
    }

    const senderEmail = getSenderEmail_();
    const targets = getMailTargets_();

    if (targets.length === 0) {
      ui.alert("送信対象が0件です。");
      settings.sheet.getRange("D2").setValue(true);
      return;
    }

    const attachments = getAttachments_(settings.attachmentUrl);
    const bannerUrl = buildBannerImageUrl_(settings.bannerImageUrl);
    const footerUrl = buildBannerImageUrl_(settings.footerImageUrl);
    const sample = targets[0];

    const sampleSubject = applyTemplate_(settings.subject, sample);
    const sampleBody = applyTemplate_(settings.body, sample);

    const confirm = ui.alert(
      "本送信 最終確認",
      `現在ログインしているアカウント：\n${senderEmail}\n\n` +
      `このアカウントから送信されます。\n\n` +
      `送信対象件数：${targets.length}件\n` +
      `バナー画像：${bannerUrl ? "あり" : "なし（H2が未設定 or URLを認識できません）"}\n` +
      `フッター画像：${footerUrl ? "あり" : "なし（I2が未設定 or URLを認識できません）"}\n` +
      `添付ファイル数：${attachments.length}件\n\n` +
      `【1件目サンプル】\n` +
      `商号：${sample.company}\n` +
      `代表者名：${addSama_(sample.representative)}\n` +
      `送信先：${sample.email}\n` +
      `件名：${sampleSubject}\n\n` +
      `本文冒頭：\n${sampleBody.substring(0, 300)}\n\n` +
      `この内容で本送信しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) {
      settings.sheet.getRange("D2").setValue(true);
      return;
    }

    const duplicateConfirm = checkSameSubjectWarning_(settings.subject);
    if (!duplicateConfirm) {
      settings.sheet.getRange("D2").setValue(true);
      return;
    }

    let sent = 0;
    let failed = 0;
    const failedRows = [];

    targets.forEach(target => {
      try {
        const subject = applyTemplate_(settings.subject, target);
        const plainBody = applyTemplate_(settings.body, target);
        const htmlBody = buildHtmlBody_(plainBody, bannerUrl, footerUrl);

        GmailApp.sendEmail(target.email, subject, plainBody, {
          htmlBody: htmlBody,
          attachments,
        });

        sent++;

      } catch (e) {
        failed++;
        failedRows.push([
          new Date(),
          target.rowNumber,
          target.company,
          addSama_(target.representative),
          target.email,
          e.message,
        ]);
      }
    });

    writeSendLog_({
      senderEmail,
      subject: settings.subject,
      targetCount: targets.length,
      sent,
      failed,
      attachmentCount: attachments.length,
    });

    if (failedRows.length > 0) {
      writeFailedLog_(failedRows);
    }

    settings.sheet.getRange("D2").setValue(true);

    ui.alert(
      `本送信完了\n\n` +
      `送信成功：${sent}件\n` +
      `送信失敗：${failed}件\n\n` +
      `D2を自動でテストモードに戻しました。`
    );

  } catch (e) {
    ui.alert("エラー：" + e.message);
  }
}

/**
 * 前回と同じ件名チェック
 */
function checkSameSubjectWarning_(subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SEND_LOG);
  const ui = SpreadsheetApp.getUi();

  if (!sheet || sheet.getLastRow() < 2) return true;

  const lastRow = sheet.getLastRow();
  const lastSubject = String(sheet.getRange(lastRow, 3).getValue() || "");

  if (lastSubject !== subject) return true;

  const confirm = ui.alert(
    "件名重複確認",
    `前回送信した件名と同じです。\n\n件名：${subject}\n\n本当に送信しますか？`,
    ui.ButtonSet.YES_NO
  );

  return confirm === ui.Button.YES;
}

/**
 * 送信ログ記録
 */
function writeSendLog_(log) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SEND_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SEND_LOG);
    sheet.appendRow([
      "送信日時",
      "送信元アカウント",
      "件名",
      "対象件数",
      "送信成功",
      "送信失敗",
      "添付数",
    ]);
  }

  sheet.appendRow([
    new Date(),
    log.senderEmail,
    log.subject,
    log.targetCount,
    log.sent,
    log.failed,
    log.attachmentCount,
  ]);
}

/**
 * 失敗ログ記録
 */
function writeFailedLog_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_FAILED_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FAILED_LOG);
    sheet.appendRow([
      "日時",
      "元行",
      "商号",
      "代表者名",
      "メールアドレス",
      "エラー内容",
    ]);
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
}

function checkExcludedApoRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_REAPPROACH);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("再アプローチリストが見つかりません");
    return;
  }

  const checkSheetName = "メルマガ除外確認";
  let checkSheet = ss.getSheetByName(checkSheetName);
  if (!checkSheet) {
    checkSheet = ss.insertSheet(checkSheetName);
  } else {
    checkSheet.clearContents();
  }

  checkSheet.appendRow([
    "判定",
    "元行",
    "商号",
    "代表者名",
    "メールアドレス",
    "K列",
    "P列",
    "U列",
    "Z列"
  ]);

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("再アプローチリストにデータがありません");
    return;
  }

  const data = sourceSheet.getRange(2, 1, lastRow - 1, 26).getDisplayValues();

  let sendCount = 0;
  let apoExcludedCount = 0;
  let noEmailCount = 0;

  const output = [];

  data.forEach((row, i) => {
    const rowNumber = i + 2;
    const company = String(row[2] || "").trim();        // C
    const representative = String(row[3] || "").trim(); // D
    const email = String(row[6] || "").trim();          // G

    const k = String(row[10] || "").trim();
    const p = String(row[15] || "").trim();
    const u = String(row[20] || "").trim();
    const z = String(row[25] || "").trim();

    const statusText = [k, p, u, z].join(" ");
    const hasApo = statusText.includes("アポ");

    if (!email) {
      noEmailCount++;
      return;
    }

    if (hasApo) {
      apoExcludedCount++;
      output.push([
        "除外：アポあり",
        rowNumber,
        company,
        representative,
        email,
        k,
        p,
        u,
        z
      ]);
      return;
    }

    sendCount++;
  });

  if (output.length > 0) {
    checkSheet.getRange(2, 1, output.length, 9).setValues(output);
  }

  checkSheet.autoResizeColumns(1, 9);

  SpreadsheetApp.getUi().alert(
    `アポ除外チェック完了\n\n` +
    `送信対象：${sendCount}件\n` +
    `アポ除外：${apoExcludedCount}件\n` +
    `メールなし除外：${noEmailCount}件\n\n` +
    `「メルマガ除外確認」シートにアポ除外分を出力しました。`
  );
}
