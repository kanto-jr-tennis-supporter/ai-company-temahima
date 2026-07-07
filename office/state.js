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
  updatedAt: "2026-07-07T14:00:00+09:00",

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
      id: "T8",
      title: "教員向けテンプレ商品の市場リサーチ（何が売れるか・社長に作れるか）",
      owner: "リサ",
      status: "review",
      progress: 100,
      hint: "トップ3から「最初に作る商品」を選ぶだけ",
      cmd: "T8を読んだ。1位の部活動管理キットでいく",
      log: [
        { time: "今日", text: "教員の困りごと×売れているテンプレ×社長のスキルで作れるか、の3点調査を開始" },
        { time: "今日", text: "調査完了。noteで教員テンプレが売れている実例を確認、ココナラの教員向けGASは空白地帯。トップ3を提案" }
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
      log: [{ time: "今日", text: "着手方針が決定。社長の準備（今年のメンバー情報など）が整い次第、開始" }],
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
