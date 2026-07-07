/* oVice風バーチャルオフィス（Phaser 3 / サーバー不要・外部素材ゼロ） */
(function () {
  // 会社データは再構築（プリセット切替・ライブ設立）のたびに読み直す
  let C = window.COMPANY || { name: "あなたのAI会社", ceo: { name: "あなた", emoji: "👑" }, employees: [] };

  const W = 1080, H = 680;

  const SLOTS = [
    { x: 200, y: 235 }, { x: 400, y: 235 },
    { x: 200, y: 430 }, { x: 400, y: 430 },
    { x: 600, y: 430 }, { x: 600, y: 235 },
  ];
  const GATHER = { x: 400, y: 520 };                            // 「集合」演出の整列位置（ホワイトボード前）

  // 情報が宿る「モノ」の配置
  const TICKER = { x1: 36, y1: 24, x2: 500, y2: 48 };          // 電光掲示板（上壁）
  const LOGO = { x: 36, y: 56, w: 250, h: 84 };                 // 社名ロゴ壁
  const SIGN = { x: 161, y: 162 };                              // 開業準備の看板
  const APPS = [
    { e: "📧", x: 560, y: 40, label: "メール", type: "mail", lampKeys: ["gmail"] },
    { e: "📅", x: 630, y: 40, label: "カレンダー", type: "calendar", lampKeys: ["calendar"] },
    { e: "💬", x: 700, y: 40, label: "LINE", type: "line", lampKeys: ["line"] },
  ];
  const TRAY = { x: 770, y: 200 };                              // 社長室の決裁トレイ
  const GUIDE = { x: 140, y: 575 };                             // 秘書アイ（ガイド役）の定位置
  const BIZ = { x1: 730, y1: 23, x2: 968, y2: 49 };             // 売上ボード（上壁・お金の見える化）
  const BOARD = { x: 400, y: 578 };                             // タスクのホワイトボード（ワークエリア内）
  const PROPOSAL = { x: 565, y: 583 };                          // 提案箱（社員からの提案）
  // 📚 資料室（ストック情報の部屋）：資料棚＝links ／ 納品BOX＝deliverables
  const SHELF = { x: 810, y: 480 };                             // 資料棚（ワイド本棚）
  const CABINET = { x: 975, y: 470 };                           // 納品キャビネット

  // 識別カラー（リング）
  const RING = ["#23b3a3", "#e87fb0", "#5aa95b", "#e0913a", "#5b8def", "#c0a93f"];

  // 社員ごとの見た目プリセット
  const SKINS = ["#f3c9a0", "#e8b083", "#d99a6c", "#f6d3b2"];
  const LOOKS = [
    { skin: SKINS[0], hair: "#6a4326", style: "bob", glasses: false },     // リサ
    { skin: SKINS[3], hair: "#caa24a", style: "ponytail", glasses: false }, // コトハ
    { skin: SKINS[2], hair: "#2b2b30", style: "short", glasses: true },     // サトル
    { skin: SKINS[1], hair: "#7a4ad0", style: "spiky", glasses: false },    // ハック
    { skin: SKINS[0], hair: "#3a3a40", style: "bun", glasses: false },
    { skin: SKINS[1], hair: "#9a3b2a", style: "short", glasses: true },
  ];
  const CEO_LOOK = { skin: SKINS[0], hair: "#33281f", style: "short", glasses: false, shirt: "#3b3f78", crown: true };

  // ---- 人物アバター（丸フレーム＋フラット人物像）----
  function drawCrown(ctx, cx, topY) {
    ctx.fillStyle = "#ffce4a";
    ctx.beginPath();
    ctx.moveTo(cx - 16, topY + 12);
    ctx.lineTo(cx - 16, topY + 2); ctx.lineTo(cx - 8, topY + 8);
    ctx.lineTo(cx, topY - 4); ctx.lineTo(cx + 8, topY + 8);
    ctx.lineTo(cx + 16, topY + 2); ctx.lineTo(cx + 16, topY + 12);
    ctx.closePath(); ctx.fill();
    ctx.strokeStyle = "#e0a82e"; ctx.lineWidth = 1.5; ctx.stroke();
    ctx.fillStyle = "#ff5e7a"; ctx.beginPath(); ctx.arc(cx, topY + 6, 2.2, 0, 7); ctx.fill();
  }

  function drawHair(ctx, style, color, cx, cy, r) {
    ctx.fillStyle = color;
    // ベースのトップドーム
    ctx.beginPath(); ctx.arc(cx, cy - 2, r + 2, Math.PI * 1.02, Math.PI * 1.98); ctx.lineTo(cx + r, cy - 2); ctx.closePath(); ctx.fill();
    ctx.beginPath(); ctx.arc(cx, cy - 4, r + 1, Math.PI, 0); ctx.fill();
    if (style === "bob") {
      ctx.fillRect(cx - r - 1, cy - 6, 8, r + 6); ctx.fillRect(cx + r - 7, cy - 6, 8, r + 6);
    } else if (style === "ponytail") {
      ctx.fillRect(cx - r - 1, cy - 6, 7, r); ctx.fillRect(cx + r - 6, cy - 6, 7, r);
      ctx.beginPath(); ctx.ellipse(cx + r + 4, cy + 6, 7, 14, -0.3, 0, 7); ctx.fill();
    } else if (style === "bun") {
      ctx.fillRect(cx - r - 1, cy - 6, 6, r - 2); ctx.fillRect(cx + r - 5, cy - 6, 6, r - 2);
      ctx.beginPath(); ctx.arc(cx, cy - r - 6, 8, 0, 7); ctx.fill();
    } else if (style === "spiky") {
      for (let i = -2; i <= 2; i++) {
        ctx.beginPath();
        ctx.moveTo(cx + i * 9 - 6, cy - r + 2); ctx.lineTo(cx + i * 9, cy - r - 9); ctx.lineTo(cx + i * 9 + 6, cy - r + 2);
        ctx.closePath(); ctx.fill();
      }
    } else { // short
      ctx.fillRect(cx - r - 1, cy - 6, 6, 12); ctx.fillRect(cx + r - 5, cy - 6, 6, 12);
    }
  }

  function makeAvatar(scene, key, look, ring) {
    if (scene.textures.exists(key)) return;
    const SZ = 124, cx = SZ / 2, R = 56;
    const headCy = 52, headR = 25;
    const shirt = look.shirt || ring;
    const cv = document.createElement("canvas"); cv.width = cv.height = SZ;
    const ctx = cv.getContext("2d");

    // 影
    ctx.fillStyle = "rgba(40,35,20,0.18)";
    ctx.beginPath(); ctx.ellipse(cx, SZ - 12, 30, 8, 0, 0, Math.PI * 2); ctx.fill();

    // 円内クリップ
    ctx.save();
    ctx.beginPath(); ctx.arc(cx, cx, R, 0, Math.PI * 2); ctx.clip();
    // 背景
    ctx.fillStyle = "#f2efe9"; ctx.fillRect(0, 0, SZ, SZ);
    ctx.fillStyle = "#e9e4da"; ctx.beginPath(); ctx.arc(cx, cx, R, 0, Math.PI * 2); ctx.fill();
    // 肩・服
    ctx.fillStyle = shirt;
    ctx.beginPath(); ctx.ellipse(cx, SZ + 2, 46, 34, 0, 0, Math.PI * 2); ctx.fill();
    // 襟
    ctx.fillStyle = "rgba(255,255,255,0.25)";
    ctx.beginPath(); ctx.moveTo(cx - 14, 84); ctx.lineTo(cx, 98); ctx.lineTo(cx + 14, 84); ctx.lineTo(cx, 90); ctx.closePath(); ctx.fill();
    // 首
    ctx.fillStyle = look.skin; ctx.fillRect(cx - 9, headCy + 16, 18, 16);
    // 頭
    ctx.fillStyle = look.skin; ctx.beginPath(); ctx.arc(cx, headCy, headR, 0, Math.PI * 2); ctx.fill();
    // 耳
    ctx.beginPath(); ctx.arc(cx - headR + 2, headCy + 2, 5, 0, 7); ctx.arc(cx + headR - 2, headCy + 2, 5, 0, 7); ctx.fill();
    // 髪
    drawHair(ctx, look.style, look.hair, cx, headCy, headR);
    // 眉
    ctx.strokeStyle = look.hair; ctx.lineWidth = 2; ctx.lineCap = "round";
    ctx.beginPath(); ctx.moveTo(cx - 13, headCy - 5); ctx.lineTo(cx - 5, headCy - 6);
    ctx.moveTo(cx + 5, headCy - 6); ctx.lineTo(cx + 13, headCy - 5); ctx.stroke();
    // 目
    ctx.fillStyle = "#3a3530";
    ctx.beginPath(); ctx.ellipse(cx - 9, headCy + 1, 2.6, 3.6, 0, 0, 7); ctx.ellipse(cx + 9, headCy + 1, 2.6, 3.6, 0, 0, 7); ctx.fill();
    // ほお
    ctx.fillStyle = "rgba(255,140,140,0.20)";
    ctx.beginPath(); ctx.arc(cx - 13, headCy + 8, 4.5, 0, 7); ctx.arc(cx + 13, headCy + 8, 4.5, 0, 7); ctx.fill();
    // 口
    ctx.strokeStyle = "#bf6e5c"; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.arc(cx, headCy + 9, 6, 0.18 * Math.PI, 0.82 * Math.PI); ctx.stroke();
    // メガネ
    if (look.glasses) {
      ctx.strokeStyle = "#3a3a42"; ctx.lineWidth = 2;
      ctx.beginPath(); ctx.arc(cx - 9, headCy + 1, 7, 0, 7); ctx.arc(cx + 9, headCy + 1, 7, 0, 7); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(cx - 2, headCy + 1); ctx.lineTo(cx + 2, headCy + 1); ctx.stroke();
    }
    ctx.restore();

    // 王冠
    if (look.crown) drawCrown(ctx, cx, headCy - headR - 6);

    // フレーム
    ctx.lineWidth = 6; ctx.strokeStyle = "#ffffff";
    ctx.beginPath(); ctx.arc(cx, cx, R, 0, Math.PI * 2); ctx.stroke();
    ctx.lineWidth = 3; ctx.strokeStyle = ring;
    ctx.beginPath(); ctx.arc(cx, cx, R - 4, 0, Math.PI * 2); ctx.stroke();

    scene.textures.addCanvas(key, cv);
  }

  class Office extends Phaser.Scene {
    constructor() { super("office"); }

    create() {
      C = window.COMPANY || { name: "あなたのAI会社", ceo: { name: "あなた", emoji: "👑" }, employees: [] };
      window.__officeScene = this;
      this.buildFloor();

      // ---- 壁アプリ（連携ランプ＋🚧つき）----
      const emojiTex = (key, emoji, size) => {
        if (this.textures.exists(key)) return;
        const s = size || 72, cv = document.createElement("canvas"); cv.width = cv.height = s;
        const ctx = cv.getContext("2d");
        ctx.font = Math.floor(s * 0.72) + "px serif"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
        ctx.fillText(emoji, s / 2, s / 2 + s * 0.04); this.textures.addCanvas(key, cv);
      };
      const gA = this.add.graphics().setDepth(3);
      this.appLamps = {}; this.appCones = {};
      APPS.forEach((a, i) => {
        const k = "app" + i; emojiTex(k, a.e, 72);
        gA.fillStyle(0xffffff, 1); gA.fillRoundedRect(a.x - 22, a.y - 18, 44, 44, 10);
        gA.lineStyle(2, 0xe2dccd, 1); gA.strokeRoundedRect(a.x - 22, a.y - 18, 44, 44, 10);
        this.add.image(a.x, a.y + 4, k).setDisplaySize(34, 34).setDepth(4);
        this.add.text(a.x, a.y + 30, a.label, { fontSize: "10px", color: "#8a8470" }).setOrigin(0.5).setDepth(4);
        // 連携ランプ（🟢=実データ／⚪=サンプル）
        const lamp = this.add.circle(a.x + 17, a.y - 13, 4.5, 0xc7cbd8, 1).setDepth(6);
        lamp.setStrokeStyle(1.5, 0xffffff, 1);
        a.lampKeys.forEach((key) => { this.appLamps[key] = lamp; });
        // セットアップ未了の🚧
        const cone = this.add.text(a.x - 19, a.y + 14, "🚧", { fontSize: "13px" }).setOrigin(0.5).setDepth(6).setVisible(false);
        a.lampKeys.forEach((key) => { this.appCones[key] = cone; });
        const z = this.add.zone(a.x, a.y + 4, 56, 70).setInteractive({ useHandCursor: true }).setDepth(50);
        z.on("pointerdown", () => { if (!window.__talkOpen) window.openApp(a.type); });
      });

      // ---- 電光掲示板（活動ログのティッカー）----
      this.add.graphics().setDepth(3)
        .fillStyle(0x23262f, 1).fillRoundedRect(TICKER.x1, TICKER.y1, TICKER.x2 - TICKER.x1, TICKER.y2 - TICKER.y1, 7)
        .lineStyle(2, 0x3a3f4e, 1).strokeRoundedRect(TICKER.x1, TICKER.y1, TICKER.x2 - TICKER.x1, TICKER.y2 - TICKER.y1, 7);
      this.tickerText = this.add.text(TICKER.x2, (TICKER.y1 + TICKER.y2) / 2, "", {
        fontSize: "13px", color: "#ffd166", fontStyle: "bold",
      }).setOrigin(0, 0.5).setDepth(4);
      const tickerMask = this.make.graphics({ add: false });
      tickerMask.fillRect(TICKER.x1 + 6, TICKER.y1, TICKER.x2 - TICKER.x1 - 12, TICKER.y2 - TICKER.y1);
      this.tickerText.setMask(tickerMask.createGeometryMask());
      const tz = this.add.zone((TICKER.x1 + TICKER.x2) / 2, (TICKER.y1 + TICKER.y2) / 2, TICKER.x2 - TICKER.x1, 30)
        .setInteractive({ useHandCursor: true }).setDepth(50);
      tz.on("pointerdown", () => { if (!window.__talkOpen && window.openActivity) window.openActivity(); });

      // ---- 売上ボード（上壁・目標への進捗）----
      this.add.graphics().setDepth(3)
        .fillStyle(0x1e3a2f, 1).fillRoundedRect(BIZ.x1, BIZ.y1, BIZ.x2 - BIZ.x1, BIZ.y2 - BIZ.y1, 7)
        .lineStyle(2, 0x2e5546, 1).strokeRoundedRect(BIZ.x1, BIZ.y1, BIZ.x2 - BIZ.x1, BIZ.y2 - BIZ.y1, 7);
      this.bizText = this.add.text(BIZ.x1 + 10, (BIZ.y1 + BIZ.y2) / 2, "🎯 売上ボード", {
        fontSize: "12px", color: "#d9f7e8", fontStyle: "bold",
      }).setOrigin(0, 0.5).setDepth(4);
      this.bizG = this.add.graphics().setDepth(4);
      const bizZone = this.add.zone((BIZ.x1 + BIZ.x2) / 2, (BIZ.y1 + BIZ.y2) / 2, BIZ.x2 - BIZ.x1, 30)
        .setInteractive({ useHandCursor: true }).setDepth(50);
      bizZone.on("pointerdown", () => { if (!window.__talkOpen && window.openBusiness) window.openBusiness(); });

      // ---- 壁の社名ロゴ＋開業準備の看板 ----
      const gL = this.add.graphics().setDepth(3);
      gL.fillStyle(0x2c3140, 1); gL.fillRoundedRect(LOGO.x, LOGO.y, LOGO.w, LOGO.h, 10);
      gL.lineStyle(2, 0x454c60, 1); gL.strokeRoundedRect(LOGO.x, LOGO.y, LOGO.w, LOGO.h, 10);
      const cName = (C.name || "あなたのAI会社");
      this.add.text(LOGO.x + LOGO.w / 2, LOGO.y + 30, "🏢 " + cName, {
        fontSize: cName.length > 10 ? "16px" : "19px", color: "#ffffff", fontStyle: "bold",
        wordWrap: { width: LOGO.w - 20 }, align: "center",
      }).setOrigin(0.5).setDepth(4);
      this.add.text(LOGO.x + LOGO.w / 2, LOGO.y + 61, (C.slogan || ""), {
        fontSize: "11px", color: "#b9c0d4", wordWrap: { width: LOGO.w - 20 }, align: "center",
      }).setOrigin(0.5).setDepth(4);
      this.signText = this.add.text(SIGN.x, SIGN.y, "", {
        fontSize: "11px", color: "#fff", fontStyle: "bold",
        backgroundColor: "#e0913a", padding: { x: 10, y: 4 },
      }).setOrigin(0.5).setDepth(6).setVisible(false);
      const sz = this.add.zone(SIGN.x, SIGN.y, 150, 26).setInteractive({ useHandCursor: true }).setDepth(50);
      sz.on("pointerdown", () => { if (!window.__talkOpen && window.openSetupStatus) window.openSetupStatus(); });

      // ---- アバター ----
      makeAvatar(this, "player", CEO_LOOK, "#ffd166");
      const emps = (C.employees || []).slice(0, SLOTS.length);
      emps.forEach((emp, i) => makeAvatar(this, "emp" + i, LOOKS[i % LOOKS.length], RING[i % RING.length]));

      this.stations = [];
      emps.forEach((emp, i) => {
        const s = SLOTS[i];
        const home = { x: s.x, y: s.y };
        const npc = this.add.image(home.x, home.y, "emp" + i).setDisplaySize(58, 58).setDepth(10);
        this.namePill(s.x, s.y + 36, emp.name || "社員", "#3a3630");
        this.add.text(s.x, s.y + 52, emp.role || "", { fontSize: "10px", color: "#9b927f" }).setOrigin(0.5).setDepth(10);

        const say = this.add.text(s.x, s.y - 46, "", {
          fontSize: "12px", color: "#3a3630", fontStyle: "bold",
          backgroundColor: "#ffffff", padding: { x: 9, y: 5 }, align: "center",
          wordWrap: { width: 150 },
        }).setOrigin(0.5).setDepth(13).setVisible(false);

        // 「今なにをしているか」の常設ラベル（💻 ◯◯を作成中…）。クリックでパネル。
        const badge = this.add.text(s.x, s.y - 38, "", {
          fontSize: "10.5px", color: "#3a3630", fontStyle: "bold",
          backgroundColor: "#fff8e0", padding: { x: 7, y: 3 },
        }).setOrigin(0.5).setDepth(14).setVisible(false);

        // 机の上の「成果物スタック」（📄 N）。クリックでその社員のパネルを開く。
        const deskBadge = this.add.text(s.x + 44, s.y + 30, "", {
          fontSize: "11px", color: "#fff", fontStyle: "bold",
          backgroundColor: "#e0913a", padding: { x: 6, y: 2 },
        }).setOrigin(0.5).setDepth(16).setVisible(false);
        const dbz = this.add.zone(s.x + 44, s.y + 30, 46, 26).setInteractive({ useHandCursor: true }).setDepth(51);
        dbz.on("pointerdown", () => { if (!window.__talkOpen) window.openTalk(emp); });

        // 机のミニ進捗バー（作業中だけ表示）
        const progressG = this.add.graphics().setDepth(16);

        const z = this.add.zone(s.x, s.y, 64, 64).setInteractive({ useHandCursor: true }).setDepth(50);
        z.on("pointerdown", () => { if (!window.__talkOpen) window.openTalk(emp); });

        const st = { emp, x: s.x, y: s.y, home, npc, say, badge, deskBadge, progressG, idx: i };
        this.stations.push(st);
        this.scheduleWander(st);
      });

      // ---- 決裁トレイ（承認待ち書類）----
      this.trayPapers = this.add.graphics().setDepth(6);
      this.trayBadge = this.add.text(TRAY.x + 24, TRAY.y - 20, "", {
        fontSize: "11px", color: "#fff", fontStyle: "bold",
        backgroundColor: "#e05a5a", padding: { x: 6, y: 2 },
      }).setOrigin(0.5).setDepth(16).setVisible(false);
      const trz = this.add.zone(TRAY.x, TRAY.y + 2, 70, 56).setInteractive({ useHandCursor: true }).setDepth(51);
      trz.on("pointerdown", () => { if (!window.__talkOpen && window.openApprovals) window.openApprovals(); });

      // ---- ホワイトボード（タスクの付箋）----
      this.boardG = this.add.graphics().setDepth(6);
      this.boardCount = this.add.text(BOARD.x, BOARD.y + 40, "", {
        fontSize: "10.5px", color: "#5a6678", fontStyle: "bold",
      }).setOrigin(0.5).setDepth(6);
      const bz = this.add.zone(BOARD.x, BOARD.y - 2, 180, 128).setInteractive({ useHandCursor: true }).setDepth(51);
      bz.on("pointerdown", () => { if (!window.__talkOpen && window.openTasks) window.openTasks(); });

      // ---- 提案箱（社員からの「勝手に提案」）----
      this.proposalBadge = this.add.text(PROPOSAL.x + 24, PROPOSAL.y - 28, "", {
        fontSize: "11px", color: "#fff", fontStyle: "bold",
        backgroundColor: "#c9a227", padding: { x: 6, y: 2 },
      }).setOrigin(0.5).setDepth(16).setVisible(false);
      const ppz = this.add.zone(PROPOSAL.x, PROPOSAL.y + 4, 70, 76).setInteractive({ useHandCursor: true }).setDepth(51);
      ppz.on("pointerdown", () => { if (!window.__talkOpen && window.openProposals) window.openProposals(); });

      // ---- 納品キャビネット（全社の成果物）----
      this.cabinetBadge = this.add.text(CABINET.x + 38, CABINET.y - 52, "", {
        fontSize: "11px", color: "#fff", fontStyle: "bold",
        backgroundColor: "#e0913a", padding: { x: 6, y: 2 },
      }).setOrigin(0.5).setDepth(16).setVisible(false);
      const cz = this.add.zone(CABINET.x, CABINET.y + 4, 100, 130).setInteractive({ useHandCursor: true }).setDepth(51);
      cz.on("pointerdown", () => { if (!window.__talkOpen && window.openCabinet) window.openCabinet(); });

      // ---- 空席：社員を雇う ----
      if (emps.length < SLOTS.length) {
        const s = SLOTS[emps.length];
        const gh = this.add.graphics().setDepth(6);
        gh.lineStyle(2, 0xc7bfac, 1); gh.strokeRoundedRect(s.x - 30, s.y - 30, 60, 60, 30);
        this.add.text(s.x, s.y - 2, "＋", { fontSize: "30px", color: "#b3a98f", fontStyle: "bold" }).setOrigin(0.5).setDepth(6);
        this.add.text(s.x, s.y + 30, "社員を雇う", { fontSize: "12px", color: "#9b927f" }).setOrigin(0.5).setDepth(6);
        const hz = this.add.zone(s.x, s.y, 80, 90).setInteractive({ useHandCursor: true }).setDepth(50);
        hz.on("pointerdown", () => { if (!window.__talkOpen && window.openHire) window.openHire(); });
      }

      // ---- 秘書アイ（ガイド役・常駐）----
      makeAvatar(this, "guide", { skin: SKINS[3], hair: "#4a3b8f", style: "bun", glasses: true }, "#8a7ff0");
      this.add.image(GUIDE.x, GUIDE.y, "guide").setDisplaySize(56, 56).setDepth(10);
      this.namePill(GUIDE.x, GUIDE.y + 34, ((C.secretary && C.secretary.name) || "アイ") + "（秘書）", "#5b3df5");
      this.add.text(GUIDE.x, GUIDE.y + 50, "クリックで案内", { fontSize: "9.5px", color: "#9b927f" }).setOrigin(0.5).setDepth(10);
      // 吹き出し：常に「次の一手」を言っている
      this.guideBubble = this.add.text(GUIDE.x, GUIDE.y - 44, "", {
        fontSize: "11px", color: "#4a32c0", fontStyle: "bold",
        backgroundColor: "#f4f2fe", padding: { x: 8, y: 4 }, align: "center",
        wordWrap: { width: 170 },
      }).setOrigin(0.5).setDepth(14);
      const gz = this.add.zone(GUIDE.x, GUIDE.y - 6, 80, 110).setInteractive({ useHandCursor: true }).setDepth(51);
      gz.on("pointerdown", () => { if (!window.__talkOpen && window.openGuide) window.openGuide(); });

      // ---- 社長（デスクに着席・操作キャラではない）----
      this.add.image(885, 128, "player").setDisplaySize(52, 52).setDepth(5);
      const ceoPill = this.namePill(885, 96, (C.ceo && C.ceo.name) || "あなた", "#5b3df5", "#ffffff");
      ceoPill.setDepth(6);

      this.meeting = false;

      this.time.addEvent({
        delay: 3800, loop: true, callback: () => {
          if (this.meeting || window.__talkOpen || !this.stations.length) return;
          // 作業中の社員は雑談しない（仕事ラベルを邪魔しない）
          const free = this.stations.filter((s) => !s.live || s.live.status === "idle");
          if (!free.length) return;
          const st = Phaser.Utils.Array.GetRandom(free);
          if (st.emp.catch) this.showSay(st, st.emp.catch, 2600);
        },
      });

      window.__startMeeting = () => this.startMeeting();

      // 直近の状態があれば即反映（state.js は game より先に読み込まれる）
      if (window.__lastState) this.applyState(window.__lastState);
      if (window.__lastHealth) this.applyHealth(window.__lastHealth);
    }

    // ====== 背景（オフィスの描き込み）======
    buildFloor() {
      // 木目床テクスチャ
      if (!this.textures.exists("wood")) {
        const T = 120, cv = document.createElement("canvas"); cv.width = cv.height = T;
        const ctx = cv.getContext("2d");
        ctx.fillStyle = "#dcc09a"; ctx.fillRect(0, 0, T, T);
        for (let y = 0; y < T; y += 24) {
          const shade = 0.5 + Math.random() * 0.3;
          ctx.fillStyle = `rgba(196,160,116,${0.25 * shade})`; ctx.fillRect(0, y, T, 24);
          ctx.strokeStyle = "rgba(150,116,78,0.5)"; ctx.lineWidth = 1;
          ctx.beginPath(); ctx.moveTo(0, y + 0.5); ctx.lineTo(T, y + 0.5); ctx.stroke();
          for (let i = 0; i < 5; i++) {
            ctx.strokeStyle = `rgba(150,116,78,${0.12 + Math.random() * 0.12})`;
            const yy = y + 4 + Math.random() * 16, xx = Math.random() * T;
            ctx.beginPath(); ctx.moveTo(xx, yy); ctx.lineTo(xx + 10 + Math.random() * 30, yy); ctx.stroke();
          }
        }
        this.textures.addCanvas("wood", cv);
      }
      this.add.tileSprite(W / 2, H / 2, W - 24, H - 24, "wood").setDepth(0);

      const g = this.add.graphics().setDepth(2);
      // 外壁
      g.lineStyle(10, 0xefe7d6, 1); g.strokeRoundedRect(12, 12, W - 24, H - 24, 20);
      g.lineStyle(2, 0xcfc4ac, 1); g.strokeRoundedRect(12, 12, W - 24, H - 24, 20);

      const shadow = (x, y, w, h, r) => { g.fillStyle(0x3a2f1c, 0.12); g.fillRoundedRect(x + 2, y + 5, w, h, r); };
      const panel = (x, y, w, h, r, fill, stroke) => { shadow(x, y, w, h, r); g.fillStyle(fill, 1); g.fillRoundedRect(x, y, w, h, r); if (stroke) { g.lineStyle(2, stroke, 1); g.strokeRoundedRect(x, y, w, h, r); } };
      const rug = (cx, cy, w, h, border, fill) => { g.fillStyle(border, 1); g.fillRoundedRect(cx - w / 2, cy - h / 2, w, h, 18); g.fillStyle(fill, 1); g.fillRoundedRect(cx - w / 2 + 7, cy - h / 2 + 7, w - 14, h - 14, 12); };

      // ラグ（ゾーン）
      rug(375, 390, 570, 500, 0xd8d0bf, 0xefe9dc);      // ワークエリア（フロー：人とタスク）
      rug(885, 185, 330, 270, 0xc9bbe6, 0xece6f8);      // 社長室（決裁）
      rug(885, 500, 330, 300, 0xd9cba8, 0xf2ead6);      // 資料室（ストック：成果物と資料）

      // ゾーン名
      this.add.text(735, 60, "👑 社長室", { fontSize: "13px", color: "#7c6bf5", fontStyle: "bold" }).setDepth(2);
      this.add.text(735, 362, "📚 資料室", { fontSize: "13px", color: "#a9842f", fontStyle: "bold" }).setDepth(2);

      // 壁掛け時計
      g.fillStyle(0xffffff, 1); g.fillCircle(1000, 36, 15);
      g.lineStyle(2, 0xcfc4ac, 1); g.strokeCircle(1000, 36, 15);
      g.lineStyle(2, 0x3a3630, 1); g.beginPath(); g.moveTo(1000, 36); g.lineTo(1000, 27); g.moveTo(1000, 36); g.lineTo(1007, 38); g.strokePath();

      // デスク（社員席ぶん）
      const emps = (C.employees || []);
      const deskCount = Math.min(emps.length, SLOTS.length);
      for (let i = 0; i < deskCount; i++) this.drawDesk(g, panel, SLOTS[i].x, SLOTS[i].y + 40);

      // 社長デスク＋チェア＋名札
      this.drawChair(g, 885, 135);
      this.drawExecDesk(g, panel, 885, 185);
      // 決裁トレイ（社長室・承認待ち書類が積まれる）
      panel(TRAY.x - 26, TRAY.y - 12, 52, 26, 4, 0xd8c3a4, 0xbfa985);
      this.add.text(TRAY.x, TRAY.y + 24, "📥 決裁", { fontSize: "10px", color: "#7c6bf5", fontStyle: "bold" }).setOrigin(0.5).setDepth(6);
      // 名札
      panel(845, 248, 60, 16, 4, 0xffffff, 0xd9cdf0);
      this.add.text(875, 256, "社長", { fontSize: "10px", color: "#7c6bf5", fontStyle: "bold" }).setOrigin(0.5).setDepth(2);

      // ホワイトボード（タスクかんばん・ワークエリア内）
      panel(BOARD.x - 85, BOARD.y - 60, 170, 110, 8, 0xffffff, 0xb9c4d6);
      g.lineStyle(3, 0x9aa6b8, 1);
      g.beginPath(); g.moveTo(BOARD.x - 62, BOARD.y + 50); g.lineTo(BOARD.x - 74, BOARD.y + 72); g.strokePath();
      g.beginPath(); g.moveTo(BOARD.x + 62, BOARD.y + 50); g.lineTo(BOARD.x + 74, BOARD.y + 72); g.strokePath();
      this.add.text(BOARD.x, BOARD.y - 46, "📋 タスク", { fontSize: "13px", color: "#4d7fc4", fontStyle: "bold" }).setOrigin(0.5).setDepth(6);
      // かんばんの列ヘッダー（未着手・進行中・🖐確認待ち・済）
      [["未", BOARD.x - 57, "#c98a2e"], ["中", BOARD.x - 19, "#3f74d6"], ["🖐", BOARD.x + 19, "#d9930d"], ["済", BOARD.x + 57, "#4a9d5b"]].forEach((c) => {
        this.add.text(c[1], BOARD.y - 28, c[0], { fontSize: "10px", color: c[2], fontStyle: "bold" }).setOrigin(0.5).setDepth(6);
      });

      // 提案箱（社員からの提案がたまる・ワークエリア内）
      panel(PROPOSAL.x - 26, PROPOSAL.y - 18, 52, 40, 6, 0xc9a227, 0xa8861d);
      g.fillStyle(0x8a6d12, 1); g.fillRoundedRect(PROPOSAL.x - 15, PROPOSAL.y - 11, 30, 5, 2); // 投入口
      this.add.text(PROPOSAL.x, PROPOSAL.y + 6, "💡", { fontSize: "16px" }).setOrigin(0.5).setDepth(6);
      this.add.text(PROPOSAL.x, PROPOSAL.y + 36, "提案箱", { fontSize: "10px", color: "#7c6bf5", fontStyle: "bold" }).setOrigin(0.5).setDepth(6);

      // ===== 📚 資料室（ストック情報の部屋）=====
      // 資料棚（ワイド本棚＝links。クリックで資料パネル）
      const sx = SHELF.x - 75, sy = SHELF.y - 95;            // 棚の左上（150×190）
      panel(sx, sy, 150, 190, 6, 0x8a5a36, 0x6e4527);
      g.lineStyle(2, 0x6e4527, 1);
      g.beginPath(); g.moveTo(SHELF.x, sy + 6); g.lineTo(SHELF.x, sy + 184); g.strokePath(); // 中央仕切り
      for (let r = 0; r < 5; r++) {
        const ry = sy + 25 + r * 38;
        g.beginPath(); g.moveTo(sx + 4, ry); g.lineTo(sx + 146, ry); g.strokePath();
        const cols = [0x6ea8fe, 0xe87fb0, 0x5aa95b, 0xe0913a, 0xc0a93f, 0x8a7ff0];
        for (let b = 0; b < 14; b++) {
          if (b === 6 || b === 7) continue; // 仕切りまわりを空ける
          g.fillStyle(cols[(r + b) % cols.length], 1); g.fillRect(sx + 8 + b * 10, ry - 17, 7, 15);
        }
      }
      this.add.text(SHELF.x, sy + 202, "🔗 資料棚（リンク・資料）", { fontSize: "10px", color: "#7c6bf5", fontStyle: "bold" }).setOrigin(0.5).setDepth(6);
      const shelfZone = this.add.zone(SHELF.x, SHELF.y + 8, 160, 220).setInteractive({ useHandCursor: true }).setDepth(50);
      shelfZone.on("pointerdown", () => { if (!window.__talkOpen && window.openLinks) window.openLinks(); });

      // 納品キャビネット（＝全社の成果物。クリックで一覧）
      panel(CABINET.x - 45, CABINET.y - 55, 90, 110, 6, 0x8a6a4a, 0x6e5238);
      for (let r = 0; r < 4; r++) {
        g.fillStyle(0xa07d58, 1); g.fillRoundedRect(CABINET.x - 38, CABINET.y - 48 + r * 26, 76, 22, 4);
        g.fillStyle(0x6e5238, 1); g.fillRoundedRect(CABINET.x - 9, CABINET.y - 40 + r * 26, 18, 5, 2);
      }
      this.add.text(CABINET.x, CABINET.y + 68, "📦 納品BOX", { fontSize: "10px", color: "#7c6bf5", fontStyle: "bold" }).setOrigin(0.5).setDepth(6);

      // 資料室の閲覧テーブル（飾り）
      panel(885, 600, 110, 42, 12, 0xd8c3a4, 0xbfa985);
      g.fillStyle(0xffffff, 1); g.fillRoundedRect(900, 610, 34, 22, 3); // 開いた本
      g.lineStyle(1, 0xc9bda4, 1); g.beginPath(); g.moveTo(917, 612); g.lineTo(917, 630); g.strokePath();

      // 観葉植物
      const plant = (x, y, s) => {
        g.fillStyle(0x3a2f1c, 0.12); g.fillEllipse(x, y + 18, 26 * s, 8 * s);
        g.fillStyle(0xc88f57, 1); g.fillRoundedRect(x - 12 * s, y + 6 * s, 24 * s, 16 * s, 4);
        g.fillStyle(0x57b15a, 1); g.fillCircle(x, y - 6 * s, 15 * s); g.fillCircle(x - 11 * s, y + 2 * s, 10 * s); g.fillCircle(x + 11 * s, y + 2 * s, 10 * s);
        g.fillStyle(0x6cc56f, 1); g.fillCircle(x - 4 * s, y - 12 * s, 8 * s); g.fillCircle(x + 6 * s, y - 8 * s, 8 * s);
      };
      plant(50, 620, 1.0); plant(690, 620, 0.9); plant(46, 330, 0.9); plant(1030, 620, 0.9);
    }

    drawDesk(g, panel, cx, cy) {
      panel(cx - 62, cy - 28, 124, 56, 10, 0x9c6b3e, 0x7e5430);
      panel(cx - 60, cy - 26, 120, 24, 6, 0xb98a57, 0x8a5f35); // 天板手前
      // モニター
      g.fillStyle(0x2b3242, 1); g.fillRoundedRect(cx - 24, cy - 22, 48, 28, 4);
      g.fillStyle(0x6fc4d8, 1); g.fillRoundedRect(cx - 21, cy - 19, 42, 22, 3);
      g.fillStyle(0x9fe0ee, 0.5); g.fillRoundedRect(cx - 21, cy - 19, 18, 22, 3);
      g.fillStyle(0x2b3242, 1); g.fillRect(cx - 3, cy + 6, 6, 5);
      // キーボード・マウス
      g.fillStyle(0xece7db, 1); g.fillRoundedRect(cx - 26, cy + 12, 40, 11, 3);
      g.fillStyle(0xece7db, 1); g.fillCircle(cx + 24, cy + 17, 4);
      // マグ
      g.fillStyle(0xe06a5a, 1); g.fillRoundedRect(cx + 34, cy - 14, 12, 12, 3);
    }

    drawChair(g, cx, cy) {
      g.fillStyle(0x3a2f1c, 0.12); g.fillEllipse(cx, cy + 16, 40, 12);
      g.fillStyle(0x44495a, 1); g.fillRoundedRect(cx - 20, cy - 18, 40, 30, 10);
      g.fillStyle(0x565d73, 1); g.fillRoundedRect(cx - 16, cy - 8, 32, 22, 8);
    }

    drawExecDesk(g, panel, cx, cy) {
      panel(cx - 70, cy - 30, 140, 60, 12, 0x6e4a8a, 0x533569);
      panel(cx - 68, cy - 28, 136, 26, 8, 0x845aa3, 0x6a4787);
      g.fillStyle(0x2b3242, 1); g.fillRoundedRect(cx - 26, cy - 24, 52, 30, 4);
      g.fillStyle(0x8a7ff0, 1); g.fillRoundedRect(cx - 23, cy - 21, 46, 24, 3);
      g.fillStyle(0xb7aef8, 0.5); g.fillRoundedRect(cx - 23, cy - 21, 20, 24, 3);
      g.fillStyle(0xece7db, 1); g.fillRoundedRect(cx - 24, cy + 10, 40, 11, 3);
    }

    namePill(x, y, text, color, txtColor) {
      return this.add.text(x, y, text, {
        fontSize: "12px", color: txtColor || "#ffffff", fontStyle: "bold",
        backgroundColor: color, padding: { x: 9, y: 3 },
      }).setOrigin(0.5).setDepth(11);
    }

    scheduleWander(st) {
      st.wanderEv = this.time.delayedCall(Phaser.Math.Between(2600, 6000), () => { this.wander(st); this.scheduleWander(st); });
    }
    wander(st) {
      if (this.meeting || window.__talkOpen) return;
      const tx = st.home.x + Phaser.Math.Between(-24, 24);
      const ty = st.home.y + Phaser.Math.Between(-16, 12);
      this.tweens.add({ targets: st.npc, x: tx, y: ty, duration: Phaser.Math.Between(900, 1500), ease: "Sine.inOut" });
    }

    showSay(st, text, ms) {
      const dy = st.badge && st.badge.visible ? -60 : -46;
      st.say.setText(text).setVisible(true).setPosition(st.npc.x, st.npc.y + dy);
      this.time.delayedCall(ms, () => st.say.setVisible(false));
    }

    // チャット連動：社員の稼働状況をオフィスに反映（💻=稼働中 / 🗣=会議中 / 📄=成果物 / 進捗バー）
    // v2：今のタスク名・進捗・成果物は tasks（taskId / owner）から引く
    applyEmployees(list, tasks) {
      if (!this.stations) return;
      tasks = tasks || (window.AI_STATE && window.AI_STATE.tasks) || [];
      const byName = {};
      (list || []).forEach((e) => { if (e && e.name) byName[e.name] = e; });
      const byId = {};
      const delsByOwner = {};
      tasks.forEach((t) => {
        if (!t) return;
        if (t.id) byId[t.id] = t;
        if (t.owner) delsByOwner[t.owner] = (delsByOwner[t.owner] || 0) + ((t.deliverables || []).length);
      });
      const trunc = (s, n) => { s = String(s || ""); return s.length > n ? s.slice(0, n) + "…" : s; };
      this.stations.forEach((st) => {
        const e = byName[st.emp.name];
        st.live = e || null;
        const tk = e && e.taskId ? byId[e.taskId] : null;
        const taskLabel = tk ? tk.title : (e && e.task) || "";
        if (st.badge) {
          if (e && e.status === "working") {
            st.badge.setText("💻 " + (taskLabel ? trunc(taskLabel, 13) : "作業中")).setVisible(true);
          } else if (e && e.status === "meeting") {
            st.badge.setText("🗣 会議中").setVisible(true);
          } else st.badge.setVisible(false);
        }
        if (st.deskBadge) {
          const n = (delsByOwner[st.emp.name] || 0) + ((e && e.deliverables && e.deliverables.length) || 0);
          if (n > 0) st.deskBadge.setText("📄 " + n).setVisible(true);
          else st.deskBadge.setVisible(false);
        }
        if (st.progressG) {
          st.progressG.clear();
          const p = tk ? Number(tk.progress || 0) : (e ? Number(e.progress || 0) : 0);
          const show = e && (e.status === "working" || (p > 0 && p < 100));
          if (show) {
            const bx = st.x - 24, by = st.y + 60, bw = 48;
            st.progressG.fillStyle(0xd8d0bf, 1).fillRoundedRect(bx, by, bw, 5, 2);
            if (p > 0) st.progressG.fillStyle(0x23b3a3, 1).fillRoundedRect(bx, by, Math.max(4, bw * Math.min(p, 100) / 100), 5, 2);
          }
        }
      });
    }

    // チャット連動：state.js の全情報をオフィスの「モノ」に反映
    applyState(s) {
      if (!s) return;
      this.applyEmployees(s.employees || [], s.tasks || []);

      // 電光掲示板（活動ログ）
      if (this.tickerText) {
        const acts = (s.activity || []).slice(0, 6);
        const line = acts.length
          ? acts.map((a) => [a.time, a.who ? a.who + "：" : "", a.text].filter(Boolean).join(" ")).join("　◆　")
          : "ようこそ！ここに会社の動きが流れます";
        if (this._tickerLine !== line) {
          this._tickerLine = line;
          this.tickerText.setText(line);
          this.tickerText.x = TICKER.x2;
        }
      }

      // 壁の看板（セットアップ進捗）＋未連携アプリの🚧
      const setup = s.setup || {};
      const steps = setup.steps || [];
      const done = steps.filter((x) => x.done).length;
      if (this.signText) {
        if (!steps.length) this.signText.setVisible(false);
        else if (setup.completed && done >= steps.length) {
          this.signText.setText("✅ OPEN 営業中").setBackgroundColor("#2e9e6b").setVisible(true);
        } else {
          this.signText.setText("🚧 開業準備 " + done + "/" + steps.length).setBackgroundColor("#e0913a").setVisible(true);
        }
      }
      if (this.appCones) {
        const stepDone = {};
        steps.forEach((x) => { stepDone[x.key] = !!x.done; });
        Object.keys(this.appCones).forEach((k) => {
          const has = steps.some((x) => x.key === k);
          this.appCones[k].setVisible(has && !stepDone[k]);
        });
      }

      // 売上ボード（目標への進捗バー）
      if (this.bizText) {
        const b = s.business;
        const fmt = (v) => { v = Number(v) || 0; return v >= 10000 ? (Math.round(v / 1000) / 10) + "万円" : v.toLocaleString() + "円"; };
        this.bizG.clear();
        if (!b || !Number(b.goalAmount)) {
          this.bizText.setText("🎯 売上ボード");
        } else {
          const cur = Number(b.current) || 0, goal = Number(b.goalAmount);
          this.bizText.setText("🎯 " + fmt(cur) + " / " + fmt(goal));
          const bx = BIZ.x2 - 96, by = (BIZ.y1 + BIZ.y2) / 2 - 4, bw = 86;
          this.bizG.fillStyle(0x2e5546, 1).fillRoundedRect(bx, by, bw, 8, 4);
          const r = Math.min(cur / goal, 1);
          if (r > 0) this.bizG.fillStyle(0x3ddc97, 1).fillRoundedRect(bx, by, Math.max(6, bw * r), 8, 4);
        }
      }

      // 秘書アイの吹き出し（次の一手）
      if (this.guideBubble && window.computeNextAction) {
        const na = window.computeNextAction(s);
        const lbl = na.label.length > 22 ? na.label.slice(0, 22) + "…" : na.label;
        this.guideBubble.setText("💡 次は：" + lbl);
      }

      // 提案箱（社員からの提案数）
      if (this.proposalBadge) {
        const np = (s.proposals || []).length;
        this.proposalBadge.setText("💡 " + np).setVisible(np > 0);
      }

      // 決裁トレイ（v2：確認待ち = status:"review" のタスク。旧 approvals も保険で数える）
      if (this.trayPapers) {
        const apv = (s.tasks || []).filter((t) => t && t.status === "review").length +
          (s.approvals || []).length;
        this.trayPapers.clear();
        for (let i = 0; i < Math.min(apv, 3); i++) {
          const px = TRAY.x - 17 + i * 2, py = TRAY.y - 6 - i * 5;
          this.trayPapers.fillStyle(0xffffff, 1).fillRoundedRect(px, py, 34, 11, 2);
          this.trayPapers.lineStyle(1, 0xc9bda4, 1).strokeRoundedRect(px, py, 34, 11, 2);
          this.trayPapers.lineStyle(1, 0xd8d2c2, 1);
          this.trayPapers.beginPath();
          this.trayPapers.moveTo(px + 4, py + 4); this.trayPapers.lineTo(px + 26, py + 4);
          this.trayPapers.moveTo(px + 4, py + 7); this.trayPapers.lineTo(px + 20, py + 7);
          this.trayPapers.strokePath();
        }
        if (this.trayBadge) this.trayBadge.setText(String(apv)).setVisible(apv > 0);
      }

      // ホワイトボード（タスクの付箋かんばん：未・中・🖐確認待ち・済）
      if (this.boardG) {
        const tasks = s.tasks || [];
        const cols = [
          { key: "todo", x: BOARD.x - 57, color: 0xffd166 },
          { key: "doing", x: BOARD.x - 19, color: 0x9ecbff },
          { key: "review", x: BOARD.x + 19, color: 0xffb84d },
          { key: "done", x: BOARD.x + 57, color: 0xa9dca9 },
        ];
        this.boardG.clear();
        const counts = { todo: 0, doing: 0, review: 0, done: 0 };
        tasks.forEach((t) => { if (counts[t.status] != null) counts[t.status]++; });
        cols.forEach((c) => {
          const n = Math.min(counts[c.key], 3);
          for (let i = 0; i < n; i++) {
            this.boardG.fillStyle(c.color, 1).fillRoundedRect(c.x - 14, BOARD.y - 19 + i * 18, 28, 15, 3);
            this.boardG.lineStyle(1, 0x8896aa, 0.5).strokeRoundedRect(c.x - 14, BOARD.y - 19 + i * 18, 28, 15, 3);
          }
        });
        if (this.boardCount) {
          this.boardCount.setText(tasks.length
            ? "未" + counts.todo + " ・ 中" + counts.doing + " ・ 🖐" + counts.review + " ・ 済" + counts.done
            : "（タスクなし）");
        }
      }

      // 納品キャビネット（v2：tasks の成果物数。旧 employees 形式も保険で数える）
      if (this.cabinetBadge) {
        const total = (s.tasks || []).reduce((n, t) => n + ((t.deliverables || []).length), 0) +
          (s.employees || []).reduce((n, e) => n + ((e.deliverables || []).length), 0);
        this.cabinetBadge.setText("📄 " + total).setVisible(total > 0);
      }
    }

    // サーバーの /health から連携状態をランプに反映
    applyHealth(h) {
      if (!this.appLamps) return;
      const set = (k, on) => { const l = this.appLamps[k]; if (l) l.setFillStyle(on ? 0x35c46a : 0xc7cbd8, 1); };
      set("gmail", !!(h && h.google));
      set("calendar", !!(h && h.google));
      set("line", !!(h && h.line));
    }

    startMeeting() {
      if (this.meeting || !this.stations.length) return;
      this.meeting = true;
      const n = this.stations.length, spacing = 74, startX = GATHER.x - (n - 1) * spacing / 2;
      this.stations.forEach((st, i) => {
        st.say.setVisible(false);
        this.tweens.killTweensOf(st.npc);
        this.tweens.add({ targets: st.npc, x: startX + i * spacing, y: GATHER.y, duration: 750, ease: "Sine.inOut" });
      });
      let t = 950;
      this.stations.forEach((st) => { this.time.delayedCall(t, () => this.showSay(st, st.emp.catch || "がんばります！", 1500)); t += 1500; });
      this.time.delayedCall(t + 400, () => {
        this.stations.forEach((st) => this.tweens.add({ targets: st.npc, x: st.home.x, y: st.home.y, duration: 750, ease: "Sine.inOut" }));
        this.time.delayedCall(950, () => { this.meeting = false; });
      });
    }

    update() {
      // 電光掲示板のスクロール
      if (this.tickerText && this.tickerText.text) {
        this.tickerText.x -= 0.7;
        if (this.tickerText.x < TICKER.x1 - this.tickerText.width) this.tickerText.x = TICKER.x2;
      }
      // 「今なにをしているか」ラベルはアバターの動きに追従
      for (const st of this.stations) {
        if (st.badge && st.badge.visible) st.badge.setPosition(st.npc.x, st.npc.y - 38);
      }
    }
  }

  const game = new Phaser.Game({
    type: Phaser.AUTO,
    parent: "game",
    width: W, height: H,
    backgroundColor: "#eae4d8",
    scale: { mode: Phaser.Scale.FIT, autoCenter: Phaser.Scale.CENTER_BOTH },
    scene: [Office],
  });

  // ライブ設立／プリセット切替：window.COMPANY を更新してオフィスを作り直す
  window.AIOffice = {
    rebuild: function (company) {
      if (company) window.COMPANY = company;
      const scene = window.__officeScene;
      if (scene && scene.scene) scene.scene.restart();
    },
    // チャット連動：社員の稼働状況をオフィスに反映
    applyEmployees: function (list) {
      const scene = window.__officeScene;
      if (scene && scene.applyEmployees) scene.applyEmployees(list);
    },
    // チャット連動：state.js の全情報をオフィスに反映
    applyState: function (s) {
      window.__lastState = s;
      const scene = window.__officeScene;
      if (scene && scene.applyState) scene.applyState(s);
    },
    // サーバー連携状態（/health）をランプに反映
    applyHealth: function (h) {
      window.__lastHealth = h;
      const scene = window.__officeScene;
      if (scene && scene.applyHealth) scene.applyHealth(h);
    },
  };
})();
