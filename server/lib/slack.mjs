// Slack 連携（Web API の小さなクライアント）
// 「未対応タスクの洗い出し」に必要な範囲だけを、依存ゼロ（組み込み fetch）で実装する。
// - トークンは必ず引数で受け取る（ハードコードしない）。search:read を持つユーザートークン(xoxp-)を想定。
// - Slack の ok:false エラーと 429（レート制限。Retry-After を尊重）に対応する。

const MAX_RETRY = 4; // 429（レート制限）時のリトライ上限

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// Slack Web API 共通の呼び出し（POST form-urlencoded・トークンは Authorization ヘッダ）。
// ok:false のときは Slack の error 文字列を含む Error を投げる。
async function slackApi(token, method, params = {}) {
  if (!token) throw new Error("Slackトークンがありません");
  const url = "https://slack.com/api/" + method;
  const body = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null) continue;
    body.set(k, String(v));
  }
  let attempt = 0;
  while (true) {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded; charset=utf-8",
        Authorization: "Bearer " + token,
      },
      body,
    });
    // レート制限：Retry-After（秒）ぶん待って数回までリトライ
    if (res.status === 429) {
      if (attempt++ >= MAX_RETRY) {
        throw new Error(`Slack ${method}: レート制限が続いています（HTTP 429）`);
      }
      const wait = Number(res.headers.get("retry-after") || "1") || 1;
      await sleep(wait * 1000);
      continue;
    }
    let data;
    try {
      data = await res.json();
    } catch {
      throw new Error(`Slack ${method}: 応答を解析できませんでした（HTTP ${res.status}）`);
    }
    if (!data.ok) {
      throw new Error(`Slack ${method}: ${data.error || "unknown_error"}`);
    }
    return data;
  }
}

// 認証テスト：このトークンの持ち主（＝社長本人＝"me"）の情報を返す。
export async function authTest(token) {
  const d = await slackApi(token, "auth.test");
  return { userId: d.user_id, user: d.user, team: d.team, teamId: d.team_id, url: d.url };
}

// メッセージ検索（search.messages）。1ページぶんの matches とページ情報を返す。
// query 例: `from:<@U07G2BM07LZ> after:2025-07-20`
export async function searchMessages(token, query, { count = 100, page = 1 } = {}) {
  const d = await slackApi(token, "search.messages", {
    query,
    count,
    page,
    sort: "timestamp",
    sort_dir: "desc",
  });
  const m = d.messages || {};
  return {
    matches: m.matches || [],
    paging: m.paging || { page: 1, pages: 1, total: (m.matches || []).length },
  };
}

// スレッドの全返信を取得（conversations.replies）。古い→新しい順の配列を返す。
export async function conversationsReplies(token, channel, ts) {
  const d = await slackApi(token, "conversations.replies", { channel, ts, limit: 200 });
  return d.messages || [];
}

// ── ユーザーID → 表示名 の解決（メモリキャッシュ付き）──
const _userCache = new Map();

export async function usersInfo(token, userId) {
  if (!userId) return "";
  if (_userCache.has(userId)) return _userCache.get(userId);
  try {
    const d = await slackApi(token, "users.info", { user: userId });
    const p = d.user?.profile || {};
    const name = p.display_name || p.real_name || d.user?.real_name || d.user?.name || userId;
    _userCache.set(userId, name);
    return name;
  } catch {
    // 解決できなくても止めない（IDのまま返す）
    _userCache.set(userId, userId);
    return userId;
  }
}

// テスト・再実行用にキャッシュを空にする
export function clearUserCache() {
  _userCache.clear();
}
