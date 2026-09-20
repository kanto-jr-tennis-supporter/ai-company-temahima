// Slack秘書：未対応タスク（＝まだ自分が返していない/片づけていないやり取り）の洗い出し
//
// 使い方（Claude Code / 秘書アイ が案内します）:
//   1. Slackアプリで search:read のユーザートークン(xoxp-…)を取得し、
//      server/.env に SLACK_USER_TOKEN=xoxp-... と貼る（詳しくは server/SETUP.md）
//   2. `cd server && node scripts/slack-followups.mjs` を実行
//   3. logs/slack-未対応タスク.md にレポートが保存され、画面にも出ます
//
// 設定（すべて環境変数）:
//   SLACK_USER_TOKEN    … 必須。search:read を持つユーザートークン(xoxp-)。社長本人のトークン＝"me"。
//   SLACK_TARGET_USER   … 追いかけたい相手のユーザーID（既定 U07G2BM07LZ）
//   SLACK_LOOKBACK_DAYS … 何日さかのぼるか（既定 60）
import "dotenv/config";
import { fileURLToPath } from "node:url";
import { dirname, join, resolve } from "node:path";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import {
  authTest, searchMessages, conversationsReplies, usersInfo,
} from "../lib/slack.mjs";

// 「対応済み」とみなす絵文字リアクション（自分が付けていたら＝片づけ済み）
const DONE_EMOJIS = new Set([
  "white_check_mark", "heavy_check_mark", "+1", "thumbsup", "ok_hand", "pray", "done",
]);

const MAX_PAGES = 10;     // 検索の最大ページ数（保険）
const MAX_MATCHES = 200;  // 解析する最大件数（保険）

// lookback 日数 → after:YYYY-MM-DD 用の日付文字列
function afterDateStr(days) {
  const d = new Date(Date.now() - days * 86400000);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${dd}`;
}

// Slackの ts（"1699999999.000100"）→ JSTの読みやすい日時
function tsToJst(ts) {
  const ms = Math.floor(parseFloat(ts) * 1000);
  if (!ms) return "";
  return new Date(ms).toLocaleString("ja-JP", {
    timeZone: "Asia/Tokyo",
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit",
  });
}

function snippet(text, max = 140) {
  const t = String(text || "").replace(/\s+/g, " ").trim();
  return t.length > max ? t.slice(0, max) + "…" : t;
}

// このメッセージに「自分の対応済みリアクション」が付いているか
function doneReactionByMe(msg, me) {
  const reactions = (msg && msg.reactions) || [];
  for (const r of reactions) {
    if (DONE_EMOJIS.has(r.name) && Array.isArray(r.users) && r.users.includes(me)) return true;
  }
  return false;
}

// 検索クエリで matches を全ページ集める（保険つき）
async function searchAll(token, query) {
  const all = [];
  let page = 1;
  let pages = 1;
  do {
    const { matches, paging } = await searchMessages(token, query, { count: 100, page });
    all.push(...matches);
    pages = paging.pages || 1;
    page += 1;
  } while (page <= pages && page <= MAX_PAGES && all.length < MAX_MATCHES);
  return all.slice(0, MAX_MATCHES);
}

// ── 中核ロジック：未対応の項目を集めて返す（ファイルには書かない）──
// サーバーの読み取りエンドポイントからも再利用できるよう、副作用なしで返す。
export async function collectFollowups({ token, target, lookbackDays, onLog } = {}) {
  const log = (m) => { if (typeof onLog === "function") onLog(m); };
  const afterDate = afterDateStr(lookbackDays);
  const warnings = [];

  // 1) 自分（"me"）が誰かを確定する
  const me = await authTest(token);
  log(`自分（社長）のSlackユーザーID: ${me.userId}（${me.user}）`);
  const targetName = await usersInfo(token, target);
  log(`追いかける相手: ${targetName}（${target}）／${afterDate} 以降を対象`);

  // 2) 相手からのメッセージを検索（本人発言＋自分へのメンション）。重複は除外。
  const queries = [
    `from:<@${target}> after:${afterDate}`,          // 相手が発したメッセージ
    `from:<@${target}> <@${me.userId}> after:${afterDate}`, // うち自分にメンションしているもの
  ];
  const byKey = new Map();
  for (const q of queries) {
    const matches = await searchAll(token, q);
    for (const m of matches) {
      const chId = m.channel?.id || "";
      const key = chId + ":" + m.ts;
      if (!byKey.has(key)) byKey.set(key, m);
    }
  }
  log(`検索でヒットした相手のメッセージ: ${byKey.size} 件（重複除外後）`);

  // 3) 1件ずつ、スレッドを見て「未対応か」を判定する
  const items = [];
  let repliesBlocked = false; // conversations.replies がスコープ不足などで使えない
  for (const m of byKey.values()) {
    const chId = m.channel?.id || "";
    const chName = m.channel?.name || m.channel?.name_normalized || "";
    const isDM = m.channel?.is_im || (!chName && (m.channel?.is_private === undefined));
    const convLabel = chName ? `#${chName}` : (isDM ? `DM（${targetName}）` : (chId || "(不明な会話)"));
    const threadRoot = m.thread_ts || m.ts;

    let thread = null;
    if (!repliesBlocked) {
      try {
        thread = await conversationsReplies(token, chId, threadRoot);
      } catch (e) {
        // スコープ不足など：以降はスレッド取得を諦め、単発として扱う
        if (/missing_scope|not_allowed_token_type|not_in_channel|channel_not_found/.test(e.message)) {
          repliesBlocked = true;
          warnings.push(
            "conversations.replies が使えないため、スレッドの中身は解析できませんでした" +
            "（単発メッセージとして判定）。相手からの発言に返信済みかどうかは目視で確認してください。"
          );
        } else {
          throw e;
        }
      }
    }

    // このメッセージ本体（リアクション確認用）：スレッド内の同 ts、無ければ検索結果そのもの
    const selfMsg = (thread && thread.find((x) => x.ts === m.ts)) || m;
    if (doneReactionByMe(selfMsg, me.userId)) continue; // ✅ 自分が対応済みリアクション → 除外

    let reason;
    if (thread && thread.length > 1) {
      const last = thread[thread.length - 1];
      if (last.user === me.userId) continue; // 自分が最後に返信済み → 対応済みとして除外
      const lastName = await usersInfo(token, last.user);
      reason = `スレッドの最後が相手側（${lastName}）の発言で、自分がまだ返していません`;
    } else {
      // スレッド化していない or スレッドを読めない → 単発メッセージ扱い
      reason = "相手からの単発メッセージに、自分がまだ返信していません";
    }

    const authorName = await usersInfo(token, m.user || target);
    items.push({
      convLabel,
      channelId: chId,
      ts: m.ts,
      timeJST: tsToJst(m.ts),
      author: authorName,
      snippet: snippet(m.text),
      permalink: m.permalink || "",
      reason,
    });
  }

  // 新しい順に並べる
  items.sort((a, b) => parseFloat(b.ts) - parseFloat(a.ts));
  return { me: me.userId, meName: me.user, target, targetName, afterDate, lookbackDays, items, warnings };
}

// ── レポート（Markdown）を組み立てる ──
export function buildReport(data) {
  const { targetName, target, afterDate, lookbackDays, items, warnings = [] } = data;
  const lines = [];
  lines.push(`# Slack未対応タスク：${targetName}（${target}）`);
  lines.push("");
  lines.push(`- 対象期間：${afterDate} 以降（約${lookbackDays}日）`);
  lines.push(`- 未対応の件数：**${items.length}件**`);
  lines.push(`- 作成：${new Date().toLocaleString("ja-JP", { timeZone: "Asia/Tokyo" })}`);
  lines.push("");

  if (warnings.length) {
    lines.push("> ⚠️ " + warnings.join(" / "));
    lines.push("");
  }

  if (items.length === 0) {
    lines.push("🎉 未対応のやり取りは見つかりませんでした。すべて返信済み／片づけ済みです。");
    lines.push("");
    return lines.join("\n");
  }

  // 会話（チャンネル/DM）ごとにグループ化
  const groups = new Map();
  for (const it of items) {
    if (!groups.has(it.convLabel)) groups.set(it.convLabel, []);
    groups.get(it.convLabel).push(it);
  }

  for (const [conv, list] of groups) {
    lines.push(`## ${conv}（${list.length}件）`);
    lines.push("");
    for (const it of list) {
      const head = it.permalink
        ? `**[${it.timeJST}](${it.permalink})**`
        : `**${it.timeJST}**`;
      lines.push(`- ${head} ／ ${it.author}`);
      lines.push(`  - 発言：${it.snippet || "(本文なし)"}`);
      lines.push(`  - なぜ未対応か：${it.reason}`);
    }
    lines.push("");
  }
  return lines.join("\n");
}

// ── CLI 本体 ──
async function main() {
  const token = process.env.SLACK_USER_TOKEN;
  const target = process.env.SLACK_TARGET_USER || "U07G2BM07LZ";
  const lookbackDays = Number(process.env.SLACK_LOOKBACK_DAYS || 60) || 60;

  if (!token) {
    console.error("秘書アイです。Slack連携のトークン（SLACK_USER_TOKEN）がまだ設定されていません。");
    console.error("Slackアプリで search:read スコープのユーザートークン（xoxp-…）を取得し、");
    console.error("server/.env に  SLACK_USER_TOKEN=xoxp-...  として貼り付けてください。");
    console.error("手順は server/SETUP.md の「Slack秘書（未対応タスクの洗い出し）」にあります。");
    process.exit(1);
  }

  let data;
  try {
    data = await collectFollowups({ token, target, lookbackDays, onLog: (m) => console.error("… " + m) });
  } catch (e) {
    // よくあるトークン権限エラーはヒントを添える
    if (/missing_scope|not_allowed_token_type/.test(e.message)) {
      console.error("⚠️ Slackから権限エラーが返りました：" + e.message);
      console.error("→ このトークンには search:read スコープが必要で、かつ『ユーザートークン(xoxp-)』である必要があります。");
      console.error("  ボットトークン(xoxb-)では search.messages は使えません。SETUP.md を確認してください。");
    } else {
      console.error("⚠️ Slackの処理でつまずきました：" + e.message);
    }
    process.exit(1);
  }

  const md = buildReport(data);

  // logs/ に保存（キット直下・git管理外。無ければ作る）
  const KIT_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..", "..");
  const outDir = join(KIT_ROOT, "logs");
  if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });
  const outPath = join(outDir, "slack-未対応タスク.md");
  writeFileSync(outPath, md, "utf-8");

  console.log(md);
  console.log("");
  console.log(`✅ レポートを保存しました → ${outPath}`);
}

// import されたとき（サーバーからの再利用など）は main を実行しない
const isMain = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) main();
