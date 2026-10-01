/*************************************************
 * 05_資料送付リスト 自動転記（アド・クリックス様用）
 * リスト（寺社仏閣）・リスト（社福）の「コール状況①〜⑤」に
 * 「資料送付」が入った行を「資料送付リスト」へ転記する。
 *  - 自動：onEdit（入力した瞬間に転記）
 *  - 一括：メニュー「コールログ」→「資料送付リストへ一括転記」
 *  - 同じ企業（企業名＋TEL）は重複させず、最新内容で上書き
 *  - E列＝資料送付になった日付、F列＝「資料送付」、G列＝転記元シート名
 *  - 資料送付リストのA〜D列には書き込まない
 *************************************************/

const SS_CFG = {
  sources: [
    { name: 'リスト（寺社仏閣）', category: '寺社仏閣' },
    { name: 'リスト（社福）',     category: '社福' }
  ],
  dest: '資料送付リスト',
  keyword: '資料送付',
  headerRow: 1,
  fixedHeaders: ['企業名', '住所', '担当者名', '役職', '部署', 'TEL', 'Mail'],
  blockStartHeader: '担当①',   // ここから右（①〜⑤の架電履歴）はまとめて転記
  statusPrefix: 'コール状況',
  destCol: {                   // 資料送付リストの列番号（1始まり）
    sofuDate: 5,               // E：資料送付になった日付
    label: 6,                  // F：「資料送付」
    sourceSheet: 7             // G：転記元シート名
  }
};

/********* 自動転記（シンプルトリガー） *********/
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const src = SS_CFG.sources.find(s => s.name === sh.getName());
    if (!src) return;

    const startRow = Math.max(e.range.getRow(), SS_CFG.headerRow + 1);
    const endRow = e.range.getLastRow();
    if (endRow < startRow) return;

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(10000)) return;
    try {
      const items = collectItems_(sh, src, startRow, endRow - startRow + 1);
      if (items.length === 0) return;
      const res = upsertToDest_(items);
      if (res.added > 0) {
        sh.getParent().toast(res.added + '件を資料送付リストに転記しました', '資料送付', 3);
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error('資料送付 自動転記エラー: ' + err);
  }
}

/********* 一括転記（既存分もまとめて） *********/
function syncShiryoSofuAll() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let items = [];
    SS_CFG.sources.forEach(src => {
      const sh = ss.getSheetByName(src.name);
      if (!sh) throw new Error('シートが見つかりません：' + src.name);
      const lastRow = sh.getLastRow();
      if (lastRow <= SS_CFG.headerRow) return;
      items = items.concat(collectItems_(sh, src, SS_CFG.headerRow + 1, lastRow - SS_CFG.headerRow));
    });
    const res = upsertToDest_(items);
    ui.alert('資料送付リストへ転記しました\n\n新規追加：' + res.added + '件\n更新：' + res.updated + '件');
  } catch (err) {
    ui.alert('エラーが発生しました：\n' + err);
  }
}

/********* 元シートから「資料送付」の行を集める *********/
function collectItems_(sh, src, startRow, numRows) {
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(SS_CFG.headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const fixedIdx = SS_CFG.fixedHeaders.map(h => header.indexOf(h)); // 最初に出てくる列を採用
  const blockStart = header.indexOf(SS_CFG.blockStartHeader);
  const statusIdx = [];
  header.forEach((h, i) => { if (h.indexOf(SS_CFG.statusPrefix) === 0) statusIdx.push(i); });
  if (blockStart < 0 || statusIdx.length === 0) throw new Error(src.name + ' のヘッダーが想定と違います');

  const values = sh.getRange(startRow, 1, numRows, lastCol).getValues();
  const items = [];
  values.forEach(row => {
    const hit = statusIdx.some(i => String(row[i]).indexOf(SS_CFG.keyword) !== -1);
    if (!hit) return;

    const fixed = {};
    SS_CFG.fixedHeaders.forEach((h, j) => { fixed[h] = fixedIdx[j] >= 0 ? row[fixedIdx[j]] : ''; });
    if (String(fixed['企業名']).trim() === '') return;

    // 「資料送付」になった日付（複数回ある場合は一番右＝最新の回）
    let sofuDate = '';
    statusIdx.forEach(i => {
      if (String(row[i]).indexOf(SS_CFG.keyword) !== -1) {
        sofuDate = row[i - 2]; // 担当｜日付｜時間｜コール状況 の並び
      }
    });

    items.push({
      key: makeKey_(fixed['企業名'], fixed['TEL']),
      fixed: fixed,
      block: row.slice(blockStart),
      sofuDate: sofuDate,
      sourceSheet: src.name
    });
  });
  return items;
}

/********* 資料送付リストへ追加 or 上書き *********/
function upsertToDest_(items) {
  const result = { added: 0, updated: 0 };
  if (items.length === 0) return result;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dest = ss.getSheetByName(SS_CFG.dest);
  if (!dest) throw new Error('シートが見つかりません：' + SS_CFG.dest);

  const lastCol = dest.getLastColumn();
  const header = dest.getRange(SS_CFG.headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const fixedIdx = SS_CFG.fixedHeaders.map(h => header.indexOf(h));
  const companyIdx = header.indexOf('企業名');
  const telIdx = header.indexOf('TEL');
  const blockStart = header.indexOf(SS_CFG.blockStartHeader);
  if (companyIdx < 0 || blockStart < 0) throw new Error(SS_CFG.dest + ' のヘッダーが想定と違います');

  const C = SS_CFG.destCol;
  const writeFrom = C.sofuDate;              // E列から右だけ書き込む（A〜Dは触らない）
  const width = lastCol - writeFrom + 1;

  // 既存データのキー一覧
  const lastRow = dest.getLastRow();
  const keyToRow = {};
  let lastFilled = SS_CFG.headerRow;
  if (lastRow > SS_CFG.headerRow) {
    const data = dest.getRange(SS_CFG.headerRow + 1, 1, lastRow - SS_CFG.headerRow, lastCol).getValues();
    data.forEach((r, i) => {
      const company = String(r[companyIdx]).trim();
      if (company === '') return;
      const rowNum = SS_CFG.headerRow + 1 + i;
      keyToRow[makeKey_(company, telIdx >= 0 ? r[telIdx] : '')] = rowNum;
      lastFilled = rowNum;
    });
  }

  const newRows = [];
  const newKeys = {};

  items.forEach(it => {
    // E列以降の1行分をつくる
    const out = new Array(width).fill('');
    const put = (col1, v) => { const k = col1 - writeFrom; if (k >= 0 && k < width) out[k] = v; };
    put(C.sofuDate, it.sofuDate);
    put(C.label, SS_CFG.keyword);
    put(C.sourceSheet, it.sourceSheet);
    SS_CFG.fixedHeaders.forEach((h, j) => { if (fixedIdx[j] >= 0) put(fixedIdx[j] + 1, it.fixed[h]); });
    const blockLen = Math.min(it.block.length, lastCol - blockStart);
    for (let b = 0; b < blockLen; b++) put(blockStart + 1 + b, it.block[b]);

    if (keyToRow[it.key]) {
      dest.getRange(keyToRow[it.key], writeFrom, 1, width).setValues([out]);
      result.updated++;
    } else if (newKeys[it.key] !== undefined) {
      newRows[newKeys[it.key]].out = out;     // 同じ企業が元リストに2回ある場合
    } else {
      newKeys[it.key] = newRows.length;
      newRows.push({ out: out });
    }
  });

  if (newRows.length > 0) {
    const start = lastFilled + 1;
    // E列以降をまとめて書き込み
    const block = newRows.map(n => n.out);
    dest.getRange(start, writeFrom, block.length, block[0].length).setValues(block);

    // B列のチェックボックスが無い行には追加（メルマガ転記用）
    const bRange = dest.getRange(start, 2, block.length, 1);
    const dv = bRange.getDataValidations();
    dv.forEach((r, i) => { if (!r[0]) dest.getRange(start + i, 2).insertCheckboxes(); });

    result.added = newRows.length;
  }
  return result;
}

/********* 重複判定キー（企業名＋TEL数字のみ） *********/
function makeKey_(company, tel) {
  const c = String(company || '').replace(/[\s　]/g, '');
  const t = String(tel || '').replace(/[^0-9０-９]/g, '');
  return c + '|' + t;
}
