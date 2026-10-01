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
  updatedAt: "2026-10-01T11:00:00+09:00",

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
      title: "商品化第1弾：部活動まるごと管理キット（設計・GAS・note販売ページ）",
      owner: "ハック",
      status: "review",
      progress: 100,
      hint: "3点セット完成。商品設計と価格（¥1,980/¥2,980の2段構え）を確認してGOなら組み立てへ",
      cmd: "T9の商品設計を見せて",
      log: [
        { time: "今日", text: "社長がT8の1位案を採用。商品設計・GASコード・販売ページの3点セットを作成開始" },
        { time: "今日", text: "3点セット完成：商品設計書・GAS一式（設置手順つき）・note販売ページ下書き。価格はライト¥1,980/完全版¥2,980" }
      ],
      deliverables: [
        { title: "T9_部活キット_商品設計.md", type: "ドキュメント", at: "7/7", path: "logs/T9_部活キット_商品設計.md", app: "Google Chrome" },
        { title: "T9_部活キット_GAS一式.md（設置手順つき）", type: "コード", at: "7/7", path: "logs/T9_部活キット_GAS一式.md", app: "Visual Studio Code" },
        { title: "T9_部活キット_note販売ページ.md", type: "ドキュメント", at: "7/7", path: "logs/T9_部活キット_note販売ページ.md", app: "Google Chrome" }
      ]
    },
    {
      id: "T10",
      title: "SNS立ち上げ準備：アカウント設計＋初週投稿バッチ（顔出しなし）",
      owner: "コトハ",
      status: "review",
      progress: 100,
      hint: "スマホで番号を返すだけ：名前5案・プロフ3案・アイコン3案・初週7投稿から選ぶ",
      cmd: "T10のSNSキットを見せて",
      log: [
        { time: "今日", text: "X中心のアカウント設計（名前・アイコン案・プロフィール）と初週投稿の作成開始。今晩スマホで承認できる形に" },
        { time: "今日", text: "完成：X一択の結論＋名前5案・プロフ3案・アイコン3案・固定ポスト・初週7投稿・運用ルール。全項目に番号を振り、スマホで返信するだけで確定できる形式" }
      ],
      deliverables: [
        { title: "T10_SNS立ち上げキット.md", type: "ドキュメント", at: "7/7", path: "logs/T10_SNS立ち上げキット.md", app: "Google Chrome" }
      ]
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
        { time: "今日", text: "社長決裁：1位「部活動まるごと管理キット」の商品化が決定 → T9へ" }
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
      id: "T11",
      title: "ときわ電設商会様：厚生労働省 入札参加資格申請書＋メール文面作成（遅延2社の1社目）",
      owner: "ハック",
      status: "review",
      progress: 90,
      hint: "申請書と郵送メール文面が完成。メールアドレス欄の空欄確認と、手書き○3箇所（新規／平成／電気工種）だけ社長にお願いしたい",
      cmd: "T11の成果物を見せて",
      log: [
        { time: "今日", text: "過去4省庁分（国交省・防衛省・農水省・文科省）の申請書を精査し、会社基本情報を確定。営業年数／住所ふりがな／営業所一覧の書き方が省庁ごとに違うことを発見" },
        { time: "今日", text: "厚労省の要項・テンプレを読み込み。フリガナは都道府県から記載、営業年数は建設業許可取得日からの計算、営業所一覧は本店を含め全記載、と判明" },
        { time: "今日", text: "社長回答（設立日基準・メール確認中・新規）とP点(611)を反映し、申請書（様式1・2枚目、様式2）とメール文面を完成。メールアドレス欄と手書き○3箇所は社長確認待ち" }
      ],
      deliverables: [
        { title: "【ときわ電設商会】様厚生労働省入札参加資格申請書（要確認）.xlsx", type: "ファイル", at: "今日", path: "../ときわ電設商会様/【ときわ電設商会】様厚生労働省入札参加資格申請書（要確認）.xlsx", app: "Numbers" },
        { title: "【ときわ電設商会】様厚生労働省_メール文面（要確認）.md", type: "ドキュメント", at: "今日", path: "../ときわ電設商会様/【ときわ電設商会】様厚生労働省_メール文面（要確認）.md", app: "Visual Studio Code" }
      ]
    },
    {
      id: "T12",
      title: "ときわ電設商会様：法務省 入札参加資格申請書＋メール文面作成",
      owner: "ハック",
      status: "review",
      progress: 95,
      hint: "数値は全て反映完了。あとは工事分割内訳表（社長対応予定）とメール確認だけ",
      cmd: "T12の成果物を見せて",
      log: [
        { time: "今日", text: "法務省の要項・様式（マクロ付きExcel）・郵送/メール送付先をWeb検索とダウンロードで自力発見。ときわ電設商会のプロフィールを流用し会社情報・営業年数(3)・資本金額(1000)・評点Y(1046)・希望地域(東京=A,他=B)を反映" },
        { time: "今日", text: "社長から総合評定通知書の残り数値を受領し反映完了（自己資本額49,014／利益額10,937／評点X1=626・X2=647・Z=515・W=131）。工事分割内訳表はマクロ依存のため社長対応、PDF出力ボタンはクライアント様対応の運用に確定" }
      ],
      deliverables: [
        { title: "【ときわ電設商会】様法務省入札参加資格申請書（要確認）.xlsm", type: "ファイル", at: "今日", path: "../ときわ電設商会様/【ときわ電設商会】様法務省入札参加資格申請書（要確認）.xlsm", app: "Numbers" },
        { title: "【ときわ電設商会】様法務省_メール文面（要確認）.md", type: "ドキュメント", at: "今日", path: "../ときわ電設商会様/【ときわ電設商会】様法務省_メール文面（要確認）.md", app: "Visual Studio Code" },
        { title: "入札申請_台帳.md（省庁別ルールの蓄積）", type: "ドキュメント", at: "今日", path: "logs/入札申請_台帳.md", app: "Visual Studio Code" },
        { title: "総合評定通知書_読み取りシート.md（今後の情報収集テンプレ）", type: "ドキュメント", at: "今日", path: "logs/総合評定通知書_読み取りシート.md", app: "Visual Studio Code" }
      ]
    },
    {
      id: "T13",
      title: "関東設備工業様：国立印刷局 入札参加資格申請書作成",
      owner: "ハック",
      status: "review",
      progress: 95,
      hint: "経審PDFのおかげで数値は全部確定・反映済み。設立年月日の年月日欄と、みなし大企業のチェック位置だけ目視確認をお願いしたい",
      cmd: "T13の成果物を見せて",
      log: [
        { time: "今日", text: "経審PDF（写真でなく原本）と国交省確定済みファイルから会社プロフィールを完全取得。管工事・水道施設工事の2業種で申請（電気ではない）" },
        { time: "今日", text: "国立印刷局の要領・Excel様式をWebから発見・ダウンロードし、全項目を反映して完成。郵送のみ・返信用封筒必須という点も要領で確認済み" }
      ],
      deliverables: [
        { title: "【関東設備工業】様国立印刷局入札参加資格申請書（要確認）.xlsx", type: "ファイル", at: "今日", path: "../関東設備工業様/【関東設備工業】様国立印刷局入札参加資格申請書（要確認）.xlsx", app: "Numbers" }
      ]
    },
    {
      id: "T14",
      title: "関東設備工業様：経済産業省 入札参加資格申請書作成",
      owner: "ハック",
      status: "review",
      progress: 80,
      hint: "内容は完成。公式様式が編集困難な旧式Wordだったため新規docxで作成。工事経歴書の代用可否だけ要項で最終確認をお願いしたい",
      cmd: "T14の成果物を見せて",
      log: [
        { time: "今日", text: "提出窓口は関東経済産業局に確定。公式様式(.doc)が罫線のない旧式ファイルで機械的な書き込みが困難と判明（このPCにLibreOffice/Node.js無く変換・見た目確認もできず）" },
        { time: "今日", text: "内容を正確に転記した新規docxを作成して代用（様式1-1・1-2・3相当）。印刷前に公式様式との見比べを推奨する注記を追加" }
      ],
      deliverables: [
        { title: "【関東設備工業】様経済産業省入札参加資格申請書（要確認）.docx", type: "ファイル", at: "今日", path: "../関東設備工業様/【関東設備工業】様経済産業省入札参加資格申請書（要確認）.docx", app: "Visual Studio Code" }
      ]
    },
    {
      id: "T15",
      title: "資料請求／送付の自動通知（アド・クリックス様スプシ → LINE・メール・Slack）",
      owner: "ハック",
      status: "done",
      progress: 100,
      hint: "",
      cmd: "T15の仕組みを見せて",
      log: [
        { time: "今日", text: "設計確定（Q1=資料請求/送付の両方・Q2=LINE未設定・Q3=出し分け・Q4=アド社のみ）。GAS本体と手順書を作成" },
        { time: "今日", text: "メール・Slackは設定即稼働／LINEはMessaging API＋グループID取得後に有効化する2段構え" },
        { time: "今日", text: "衝突チェック：既存アポ転記GASの notifyError_ と名前かぶりを発見 → 全名前を RR_CONFIG / rr〜 / setupTriggerRR / testNotifyRR に改名して解消" },
        { time: "今日", text: "手順書STEP1を『既存コードを消さず＋で新規ファイル追加』に修正。残りはスプシへの貼り付け〜テスト送信（画面操作）" },
        { time: "今日", text: "既存2本（架電ログ/資料送付転記）と通知を1ファイルに統合。通知の見出し不一致（コール状況①〜⑤・TEL）を修正し、モックで動作確認" },
        { time: "今日", text: "社長の並び替えに合わせ架電ログ③〜⑤の連絡先/メール列を修正。資料請求も資料送付リストへ転記するよう追加（モックで確認）" },
        { time: "今日", text: "キャッチ判定に【資料請求】を追加（社長決定）" },
        { time: "今日", text: "①完了：Apps Scriptを完成版1本（1039行）に入れ替え。次は②送り先の記入" },
        { time: "今日", text: "②メール宛先・③Slack Webhook（アプリ Shiryo Notify）を社長が設定完了。次は④自動通知ON→⑤テスト" },
        { time: "今日", text: "④自動通知ON・⑤テスト送信成功（メール＋Slack #PJ-アドクリックス）。残りは実シートでの最終確認" },
        { time: "今日", text: "実シートで最終確認OK（通知・資料請求の転記・二重通知防止の3点）。メール＋Slackで稼働開始。LINEはT16へ" }
      ],
      deliverables: [
        { title: "アドクリックス_GAS完成版.gs（1ファイル統合版）", type: "コード", at: "今日", path: "logs/アドクリックス_GAS/アドクリックス_GAS完成版.gs", app: "Visual Studio Code" },
        { title: "資料請求通知_セットアップ手順.md", type: "ドキュメント", at: "今日", path: "logs/資料請求通知_セットアップ手順.md", app: "Google Chrome" },
        { title: "T15_引き継ぎメモ.md", type: "ドキュメント", at: "今日", path: "logs/T15_引き継ぎメモ.md", app: "Google Chrome" }
      ]
    },
    {
      id: "T16",
      title: "資料請求通知のLINE対応（クライアント様グループへ企業名・住所・TELを自動報告）",
      owner: "ハック",
      status: "doing",
      progress: 5,
      hint: "",
      cmd: "T16の続きを案内して",
      log: [
        { time: "今日", text: "T15から切り出し。コードは対応済み（NOTIFY_CFG.LINE を埋めて enabled:true にするだけ）。まず add one 公式LINEの現状（Lステップ等の連携有無）を確認" }
      ],
      deliverables: [
        { title: "資料請求通知_セットアップ手順.md（STEP6がLINE）", type: "ドキュメント", at: "今日", path: "logs/資料請求通知_セットアップ手順.md", app: "Google Chrome" }
      ]
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
    { from: "サトル", title: "納品後の定番メニュー「継続プラン」を設計", detail: "受注のたびに月額プランを必ず添える型を作る。3社×3万円で毎月9万円の土台になります。", at: "12:00" },
  ],

  employees: [
    { name: "リサ", status: "idle", taskId: "" },
    { name: "コトハ", status: "idle", taskId: "" },
    { name: "サトル", status: "idle", taskId: "" },
    { name: "ハック", status: "working", taskId: "T16" },
  ],

  links: [
    { title: "テマヒマGitHubリポジトリ（スマホ連携の本体）", url: "https://github.com/kanto-jr-tennis-supporter/ai-company-temahima", type: "リンク" },
    { title: "テレアポシート（コールログ・転記元）", url: "https://docs.google.com/spreadsheets/d/1oQpq8VukdjWfcKqv_YCZTlLUQHPPkqRTcvyqfAny3YU/edit", type: "スプレッドシート" },
    { title: "テレアポシート（コール数集計）", url: "https://docs.google.com/spreadsheets/d/1NqLmQez1G--M8A6vDCnXLnwy-oBy1wutr9NC84ZO1Gs/edit?gid=1036168354", type: "スプレッドシート" },
    { title: "入札申請書＋メール作成アプリ", url: "https://script.google.com/macros/s/AKfycbycvMzYrtTti87O_FQ7th8zsT-KSfa5z4a3g8rlumgdsZyuaLJ5Ic5043IH6yN8oC3J5A/exec", type: "アプリ" },
    { title: "部活動出欠確認シート", url: "https://docs.google.com/spreadsheets/d/19439G8kmy-HtxeLAM8ooHoIWA8v5kIxBN7nv1RJlnrA/edit?usp=drive_link", type: "スプレッドシート" },
  ],

  activity: [
    { time: "11:00", who: "ハック", text: "T15完了：資料請求・資料送付の自動通知がメール＋Slackで稼働開始。T16（LINE対応）に着手" },
    { time: "10:50", who: "ハック", text: "T15：資料請求通知のテスト送信に成功！メール・Slackに届きました" },
    { time: "10:45", who: "ハック", text: "T15：架電ログ③〜⑤の列修正＋資料請求も転記に対応。入れ替え待ちです" },
    { time: "10:15", who: "ハック", text: "T15：既存GAS2本＋通知を1ファイルの完成版に統合しました" },
    { time: "09:30", who: "ハック", text: "T15：通知GASの名前かぶり（notifyError_）を解消。貼るだけの状態です" },
    { time: "今日", who: "コトハ", text: "T10：SNS立ち上げキット完成！名前・プロフ・アイコン・初週7投稿を、スマホで番号を返すだけで確定できる形に✍️" },
    { time: "今日", who: "ハック", text: "T9：部活キット3点セット完成！商品設計・GAS一式・note販売ページ。価格は¥1,980/¥2,980の2段構え💻" },
    { time: "今日", who: "アイ", text: "T9（部活キット商品化）とT10（SNS立ち上げ）が同時始動。ハックとコトハが並列で作業中💻" },
    { time: "今日", who: "アイ", text: "スマホ連携が開通！社長がスマホからテマヒマに接続成功📱🎉" },
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
