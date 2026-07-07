// 🔄 会社の「いまの状態」。Claude Code がここを書き換えると、
//    オフィス画面（看板・掲示板・タスクリスト・社員の動き）に反映されます。
//    ※ 手で編集してもOK。チャット＝実働 / オフィス＝俯瞰 をつなぐ橋渡しファイルです。
//
// ★ v2：すべての仕事は「タスク」を中心に回る。
//    社員の稼働・確認待ち（旧 approvals）・成果物は、ぜんぶ tasks の中に入れる。
//
// オフィスでの見え方:
//   setup     → 壁の「開業準備」看板（＋未連携の壁アプリに🚧）
//   business  → 上壁の売上ボード🎯（目標への進捗バー。クリックで営業ファネル）
//   tasks     → ワークエリアのホワイトボード（クリックでAppleメモ風タスクリスト）
//               ・status:"review" のタスク = 社長の確認待ち → 社長室の決裁トレイ📥にも積まれる
//               ・タスクの deliverables = 資料室の納品BOX📦に集約される
//   proposals → ワークエリアの提案箱💡（社員からの提案がたまる。採用したら tasks に起票）
//   activity  → 上壁の電光掲示板（流れるテキスト・クリックで履歴）
//   employees → アバターのバッジ💻🗣（taskId のタスク名が頭上に出る）・机の進捗バー・机の📄
//   links     → 資料室の資料棚🔗
//
// スキーマ:
//   updatedAt : 更新時刻（ISO文字列）。更新するたびに必ず新しくする。
//   setup     : 初回セットアップの進捗 { completed, steps:[{ key, label, done }] }
//   business  : 売上とゴール { goalLabel, goalAmount, current, pipeline: [{ label, count }] }
//   tasks     : ★仕事の中心。1仕事 = 1タスク。 [{
//                 id,                       ← "T1" 形式の通し番号。一度振ったら変えない
//                 title,                    ← 何をするか（社長が読んで分かる言葉で）
//                 owner,                    ← 担当社員名（"リサ" 等）。未定なら ""
//                 status,                   ← "todo"（未着手）| "doing"（進行中）|
//                                             "review"（社長の確認待ち）| "done"（完了）
//                 progress,                 ← 0〜100。doing のとき進捗リングに出る
//                 hint,                     ← review のとき「社長は何をすればいいか」を一言で（任意）
//                 cmd,                      ← このタスクに対する指示文（コピー用）。
//                                             例 "T3の途中経過を見せて" / review時 "T5の1通目OK。量産して"
//                 log: [{ time, text }],    ← 経過ログ（新しいものを下に。最大8件）
//                 deliverables: [{          ← このタスクの成果物（完成したら積む）
//                   title, type, at, body, url, path, app   ← §9.3 参照
//                 }]
//               }]
//   proposals : 社員からの提案 [{ from, title, detail, at }]（採用されたら tasks に起票して消す）
//   employees : 社員の稼働 [{ name, status: "idle"|"working"|"meeting",
//                            taskId }]      ← いま取り組んでいるタスクのid（idle なら ""）
//   links     : 会社の資料・リンク集 [{ title, url, type }]
//   activity  : 最近の動き [{ time, who, text }]（新しいものを上に。電光掲示板に流れる）
//   command   : 一度だけ再生する演出 { id, type: "inauguration"|"meeting"|"founding" } or null

window.AI_STATE = {
  updatedAt: "2026-07-07T16:00:00+09:00",

  setup: {
    completed: true,
    steps: [
      { key: "design", label: "会社を設計する（あなたの情報を聞く）", done: true },
      { key: "launch", label: "会社を立ち上げる（設立・就任式）", done: true },
      { key: "office", label: "オフィスを開く（この画面）", done: true },
      { key: "gmail", label: "メール連携（任意）", done: false },
      { key: "calendar", label: "カレンダー連携（任意）", done: false },
      { key: "line", label: "LINE連携：外出先から指示（任意）", done: false },
      { key: "skills", label: "社員の専用ツールをつなぐ（任意・導入オプション）", done: false },
    ],
  },

  business: {
    goalLabel: "7月の売上目標 6.5万円（→安定月16.5万円へ）",
    goalAmount: 65000,
    current: 0,
    pipeline: [
      { label: "応募", count: 0 },
      { label: "返信", count: 0 },
      { label: "受注", count: 0 },
    ],
  },

  tasks: [
    {
      id: "T9",
      title: "クラウドワークス6案件の精査（社長にできるか判定）",
      owner: "リサ",
      status: "done",
      progress: 100,
      hint: "",
      cmd: "",
      log: [
        { time: "15:00", text: "社長指定の6案件URLを1件ずつ精査開始" },
        { time: "15:05", text: "この環境からクラウドワークスへの接続がブロックされ本文を取得できず。判定保留。社長に本文の貼り付けを依頼中" },
        { time: "15:30", text: "社長から6案件の本文を受領。リサが◎○×の判定を再開" },
        { time: "15:45", text: "判定完了。応募推奨は①データ集計効率化のみ（○条件付きGO）。③⑤は条件不一致、②④は時間の切り売り、⑥はスパム加担リスクで×" }
      ],
      deliverables: [{ title: "T9_クラウドワークス6件精査.md（判定完了版）", type: "ドキュメント", at: "7/7", path: "logs/T9_クラウドワークス6件精査.md", app: "Visual Studio Code" }]
    },
    {
      id: "T10",
      title: "精査でGOになった案件の応募文作成",
      owner: "コトハ",
      status: "review",
      progress: 90,
      hint: "応募文を確認してOKか修正指示を出すだけ。OKなら応募は社長の手で（今日〜明日推奨）",
      cmd: "T10の応募文OK。このまま応募する",
      log: [
        { time: "15:00", text: "T9（リサの精査）の結果待ち。GO案件が決まり次第着手" },
        { time: "15:45", text: "GO案件は①データ集計効率化の1件。コトハが応募文の執筆を開始" },
        { time: "16:00", text: "応募文完成。テレアポシート実績を冒頭に、稼働条件を正直に明示。社長の承認待ち" }
      ],
      deliverables: [{ title: "T10_応募文_データ集計効率化.md（応募前チェックメモ付き）", type: "ドキュメント", at: "7/7", path: "logs/T10_応募文_データ集計効率化.md", app: "Visual Studio Code" }]
    },
    {
      id: "T11",
      title: "部活動まるごと管理キット：商品設計書v1（機能・構成・価格）",
      owner: "サトル",
      status: "review",
      progress: 100,
      hint: "論点3つに答えるだけ（会計の版切り分け・出欠区分・販売名義）",
      cmd: "T11の論点1は賛成。区分は◯◯を使ってる。名義は◯◯で",
      log: [
        { time: "15:15", text: "社長が第1候補「部活動まるごと管理キット」を採択。T8のリサーチをもとに商品設計を開始" },
        { time: "15:30", text: "設計書v1完成。基本版¥1,980／完全版¥3,980／カスタム¥5,000〜の3段構成。実装可能な粒度で機能要件を定義。社長の論点3つを提示" }
      ],
      deliverables: [{ title: "T11_部活動まるごと管理キット_商品設計書v1.md", type: "ドキュメント", at: "7/7", path: "logs/T11_部活動まるごと管理キット_商品設計書v1.md", app: "Visual Studio Code" }]
    },
    {
      id: "T8",
      title: "教員向けテンプレ商品の市場リサーチ（何が売れるか・社長に作れるか）",
      owner: "リサ",
      status: "done",
      progress: 100,
      hint: "",
      cmd: "",
      log: [
        { time: "今日", text: "教員の困りごと×売れているテンプレ×社長のスキルで作れるか、の3点調査を開始" },
        { time: "今日", text: "調査完了。noteで教員テンプレが売れている実例を確認、ココナラの教員向けGASは空白地帯。トップ3を提案" },
        { time: "15:15", text: "社長が第1位「部活動まるごと管理キット」を採択。商品化はT11へ" }
      ],
      deliverables: [{ title: "T8_教員向けテンプレ市場リサーチ.md", type: "ドキュメント", at: "7/7", path: "logs/T8_教員向けテンプレ市場リサーチ.md", app: "Visual Studio Code" }]
    },
    {
      id: "T6",
      title: "クラウドワークス提案文の改善（コトハ）",
      owner: "コトハ",
      status: "doing",
      progress: 10,
      hint: "",
      cmd: "T6の提案文を見せて",
      log: [{ time: "今日", text: "現行2パターンを分析中。社長の実績（GAS業務委託）を武器に書き直す" }],
      deliverables: []
    },
    {
      id: "T7",
      title: "クラウドワークス案件リサーチ（リサ）",
      owner: "リサ",
      status: "doing",
      progress: 10,
      hint: "",
      cmd: "T7の案件一覧を見せて",
      log: [{ time: "今日", text: "スプレッドシート・GAS・自動転記系を中心に探索中" }],
      deliverables: []
    },
    {
      id: "T1",
      title: "テレアポシート①：「資料」企業を資料送付一覧に転記（同シート内）",
      owner: "ハック",
      status: "done",
      progress: 100,
      hint: "",
      cmd: "",
      log: [{ time: "7/2", text: "GASスクリプト完成。プルダウン・色付け・自動転記すべて完了" }],
      deliverables: [{ title: "GAS_ToGList_v3.js", type: "コード", at: "7/2", path: "logs/GAS_ToGList_v3.js", app: "Visual Studio Code" }]
    },
    {
      id: "T2",
      title: "テレアポシート②：「資料」企業を別ファイルに転記",
      owner: "ハック",
      status: "done",
      progress: 100,
      hint: "",
      cmd: "",
      log: [{ time: "7/2", text: "再アプローチリストへの自動転記・一括転記・プルダウン・色付け完了" }],
      deliverables: []
    },
    {
      id: "T3",
      title: "テレアポシート③：コール数集計の修正",
      owner: "ハック",
      status: "doing",
      progress: 10,
      hint: "",
      cmd: "T3の状況を教えて",
      log: [{ time: "7/2", text: "着手。シートの構成確認中" }],
      deliverables: []
    },
    {
      id: "T4",
      title: "入札参加資格申請書＋メール作成アプリの完成",
      owner: "ハック",
      status: "todo",
      progress: 0,
      hint: "",
      cmd: "T4を始めて",
      log: [],
      deliverables: []
    },
    {
      id: "T5",
      title: "部活動出欠確認シート：今年のメンバーに更新・修正",
      owner: "ハック",
      status: "todo",
      progress: 0,
      hint: "",
      cmd: "T5を始めて",
      log: [
        { time: "今日", text: "着手方針が決定。社長の準備（今年のメンバー情報など）が整い次第、開始" },
        { time: "15:05", text: "社長から今年の女子メンバー20名（1年8名・2年12名）を受領。logs/T5_2026年度メンバー_女子.csv に保存。男子リスト待ち" }
      ],
      deliverables: []
    },
  ],
  proposals: [
    { from: "リサ", title: "毎朝の自動案件パトロール", detail: "GAS・スプシ自動化の新着案件を毎朝リサーチして一覧化。応募の弾を切らさないのが今月6.5万の生命線です。", at: "12:00" },
    { from: "コトハ", title: "既存クライアントに月額サポートを提案", detail: "テレアポシートの納品先に「月額保守・改善サポート」を提案。単発で終わらせないのが月16.5万への一番の近道です。", at: "12:00" },
    { from: "サトル", title: "納品後の定番メニュー「継続プラン」を設計", detail: "受注のたびに月額プランを必ず添える型を作る。3社×3万円で毎月9万円の土台になります。", at: "12:00" },
  ],

  employees: [
    { name: "リサ", status: "idle", taskId: "" },
    { name: "コトハ", status: "idle", taskId: "" },
    { name: "サトル", status: "idle", taskId: "" },
    { name: "ハック", status: "working", taskId: "T3" },
  ],

  links: [
    { title: "テマヒマGitHubリポジトリ（スマホ連携の本体）", url: "https://github.com/kanto-jr-tennis-supporter/ai-company-temahima", type: "リンク" },
    { title: "テレアポシート（コールログ・転記元）", url: "https://docs.google.com/spreadsheets/d/1oQpq8VukdjWfcKqv_YCZTlLUQHPPkqRTcvyqfAny3YU/edit", type: "スプレッドシート" },
    { title: "テレアポシート（コール数集計）", url: "https://docs.google.com/spreadsheets/d/1NqLmQez1G--M8A6vDCnXLnwy-oBy1wutr9NC84ZO1Gs/edit?gid=1036168354", type: "スプレッドシート" },
    { title: "入札申請書＋メール作成アプリ", url: "https://script.google.com/macros/s/AKfycbycvMzYrtTti87O_FQ7th8zsT-KSfa5z4a3g8rlumgdsZyuaLJ5Ic5043IH6yN8oC3J5A/exec", type: "アプリ" },
    { title: "部活動出欠確認シート", url: "https://docs.google.com/spreadsheets/d/19439G8kmy-HtxeLAM8ooHoIWA8v5kIxBN7nv1RJlnrA/edit?usp=drive_link", type: "スプレッドシート" },
  ],

  activity: [
    { time: "16:00", who: "コトハ", text: "T10：応募文が完成！決裁トレイに上げました。社長のOK待ちです📥" },
    { time: "15:45", who: "リサ", text: "T9：判定完了！応募推奨は①データ集計効率化の1件。⑥は危険案件として見送り推奨🔍" },
    { time: "15:45", who: "コトハ", text: "T10：①の応募文を書き始めました。テレアポシート実績を武器にします✍️" },
    { time: "15:30", who: "サトル", text: "T11：商品設計書v1が完成！社長の論点3つの回答待ちです🧭" },
    { time: "15:30", who: "リサ", text: "T9：社長から6案件の本文を受領。判定を再開しました🔍" },
    { time: "15:15", who: "アイ", text: "T8決裁！最初の商品は「部活動まるごと管理キット」に決定🎉 サトルが商品設計に着手（T11）" },
    { time: "15:05", who: "リサ", text: "T9：クラウドワークスに接続できず判定保留。社長に案件本文の貼り付けをお願いしています🙏" },
    { time: "15:05", who: "アイ", text: "T5：今年の女子メンバー20名を受領しました。男子リストが揃えば着手できます📝" },
    { time: "15:00", who: "リサ", text: "T9：社長指定のクラウドワークス6案件の精査を開始しました🔍" },
    { time: "今日", who: "アイ", text: "スマホ連携の土台が完成！会社一式をGitHub（Private）にアップロードしました📱" },
    { time: "今日", who: "リサ", text: "T8：リサーチ完了！ココナラの教員向けGASは空白地帯。商品候補トップ3を提案しました📊" },
    { time: "今日", who: "リサ", text: "T8：教員向けテンプレ商品の市場リサーチを開始しました🔍" },
    { time: "今日", who: "アイ", text: "売上目標を更新：7月6.5万円→安定月16.5万円。作戦を提案箱に投函しました💡" },
    { time: "今日", who: "アイ", text: "T5（部活動出欠確認シート更新）から着手することが決定。社長の準備を待機中" },
    { time: "今日", who: "アイ", text: "5つのタスクを起票。ハックが T1〜T5 に着手します" },
    { time: "14:35", who: "アイ", text: "「テマヒマ・ラボ」設立！山元光樹社長が代表取締役CEOに就任しました🎉" },
  ],

  command: { id: "inauguration-2026-06-24", type: "inauguration" },
};
