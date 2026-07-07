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
  updatedAt: "2026-01-01T00:00:00+09:00",

  setup: {
    completed: false,
    steps: [
      { key: "design", label: "会社を設計する（あなたの情報を聞く）", done: false },
      { key: "launch", label: "会社を立ち上げる（設立・就任式）", done: false },
      { key: "office", label: "オフィスを開く（この画面）", done: false },
      { key: "gmail", label: "メール連携（任意）", done: false },
      { key: "calendar", label: "カレンダー連携（任意）", done: false },
      { key: "line", label: "LINE連携：外出先から指示（任意）", done: false },
      { key: "skills", label: "社員の専用ツールをつなぐ（任意・導入オプション）", done: false },
    ],
  },

  business: {
    goalLabel: "売上目標（セットアップで決めます）",
    goalAmount: 0,
    current: 0,
    pipeline: [
      { label: "応募", count: 0 },
      { label: "返信", count: 0 },
      { label: "受注", count: 0 },
    ],
  },

  tasks: [],
  proposals: [],

  employees: [
    { name: "リサ", status: "idle", taskId: "" },
    { name: "コトハ", status: "idle", taskId: "" },
    { name: "サトル", status: "idle", taskId: "" },
    { name: "ハック", status: "idle", taskId: "" },
  ],

  links: [],

  activity: [
    { time: "09:00", who: "アイ", text: "キットへようこそ！「セットアップを始めて」で、あなたの会社を立ち上げます" },
  ],

  command: null,
};
