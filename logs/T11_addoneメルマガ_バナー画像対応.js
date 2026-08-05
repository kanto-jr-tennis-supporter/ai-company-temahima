/****************************************************
 * 再アプローチリスト直接参照メルマガ送信 完全版（バナー画像対応）
 * add one様向け／テマヒマ・ラボ ハック作成 / 2026-08-05
 *
 * 【変更点（元スクリプトからの差分）】
 * ・メルマガ用シートに H2：バナー画像URL を追加
 *   （Googleドライブの共有リンクをそのまま貼ってOK。中でファイルIDを抽出して変換します）
 * ・本文をプレーンテキストではなく HTML メールとして送信するように変更
 *   → バナー画像がメール本文の一番上に表示される
 * ・plainBody（画像非表示の環境向けの文字だけ版）は今まで通り自動生成されるので、
 *   万一画像が読み込めない受信環境でも本文は読める
 *
 * 【事前準備・唯一の注意点】
 * バナー画像のGoogleドライブファイルを「リンクを知っている全員が閲覧可」に共有設定してください。
 * （限定公開のままだと相手のメールで画像が表示されません＝赤い×アイコンになります）
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
 * H2：バナー画像URL（新規追加・Googleドライブの共有リンクを貼る）
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
 */
function addSama_(name) {
  const n = String(name || "").trim();
  if (!n) return "";
  if (n.endsWith("様")) return n;
  return n + " 様";
}

/**
 * 差し込み
 */
function applyTemplate_(text, target) {
  return String(text || "")
    .replaceAll("{{商号}}", target.company || "")
    .replaceAll("{{ 商号 }}", target.company || "")
    .replaceAll("{{会社名}}", target.company || "")
    .replaceAll("{{ 会社名 }}", target.company || "")
    .replaceAll("{{代表者名}}", addSama_(target.representative))
    .replaceAll("{{ 代表者名 }}", addSama_(target.representative));
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
 * プレーン本文＋バナー画像URLから、HTMLメール本文を組み立てる
 */
function buildHtmlBody_(plainBody, bannerImageUrl) {
  const bannerHtml = bannerImageUrl
    ? `<img src="${bannerImageUrl}" alt="バナー" style="max-width:600px;width:100%;height:auto;display:block;margin:0 0 16px 0;border:0;">`
    : "";

  const textHtml = escapeHtml_(plainBody).replace(/\n/g, "<br>");

  return (
    `<div style="font-family:'Hiragino Kaku Gothic ProN','Meiryo',sans-serif;` +
    `font-size:14px;line-height:1.8;color:#333333;max-width:600px;">` +
    `${bannerHtml}${textHtml}</div>`
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
    const htmlBody = buildHtmlBody_(plainBody, bannerUrl);
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
    const sample = targets[0];

    const sampleSubject = applyTemplate_(settings.subject, sample);
    const sampleBody = applyTemplate_(settings.body, sample);

    const confirm = ui.alert(
      "本送信 最終確認",
      `現在ログインしているアカウント：\n${senderEmail}\n\n` +
      `このアカウントから送信されます。\n\n` +
      `送信対象件数：${targets.length}件\n` +
      `バナー画像：${bannerUrl ? "あり" : "なし（H2が未設定 or URLを認識できません）"}\n` +
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
        const htmlBody = buildHtmlBody_(plainBody, bannerUrl);

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
