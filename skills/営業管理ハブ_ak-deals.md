# ⚙️ スキル：営業管理ハブ（ak-deals）— 担当：ハック

7媒体（Wantedly / 複業クラウド等）を横断して、**見込み案件のメッセージ取得→返信ドラフト生成→
日程登録までのパイプラインをスプレッドシートで一元管理**するハブ。

- **場所**: `tools/ak-deals`（キット内。実体プロジェクトへのリンク。無い場合は導入オプション）
- **詳細な使い方**: `tools/ak-deals/CLAUDE.md` を読むこと（スキル発動条件・鉄則が全部書いてある）

## 主要コマンド（`tools/ak-deals` で実行）

```bash
cd tools/ak-deals
node daily-run.mjs                        # 全媒体からメッセージ取得（毎日の巡回）
node fetch.mjs --media=Wantedly           # 1媒体だけ取得
node list-drafts-pending.mjs --json       # 返信ドラフト未生成の案件一覧
node show-deal.mjs --deal=3               # 案件3のドラフト材料を表示
node set-draft.mjs --deal=3 --file=tmp/draft-3.txt   # ドラフトをスプシに書き込み
node register-event.mjs --deal=3 --yes    # 面談日程をカレンダー登録
node tools/check-setup.mjs                # セットアップ診断（不調時はまずこれ）
```

## 進め方（ハックの仕事の型）

1. 「巡回して」と言われたら `daily-run.mjs` → 新着・未対応・期限切れを要約して報告。
2. 未対応で最も古いものを最優先として提示（選択肢を並べず、次の一手を決めて出す）。
3. 返信ドラフトは生成してスプシに書くまで。**送信は社長の決裁**（秘書アイがタスクを🖐確認待ち＝`status:"review"` にする）。
4. 認証Cookieが古い等の障害は `check-setup.mjs` で特定し、解消手順を案内。

## 前提（動かない時はここを確認）

- `.env`・Google OAuth・各媒体の認証情報（`node save-creds.mjs --status` で確認）が必要。
- 媒体認証が一部切れていても、生きている媒体だけで巡回は続行できる。

## 納品ルール

- 巡回結果のサマリ（新着N件・要対応N件・最優先1件）を報告。管理スプシURLを報告に含める。
