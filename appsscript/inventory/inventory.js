/***** 盤點小幫手 — 後端 (Google Apps Script)  v3 *****
 * v3 變更（配合前端穩定性強化）：
 *  1. 每筆請求帶唯一 uid，後端用 CacheService 去重（保存 6 小時）
 *     → 前端網路逾時自動重試時，若其實已寫入成功，不會重複 +1
 *  2. 其餘邏輯與 v2 相同
 *
 * v2 功能：
 *  1. 營業日換日點為早上 8:00 → 例如分頁「0623」代表 6/23 08:00 ~ 6/24 07:59（補零四碼 MMdd）
 *  2. 寫入後直接回傳，不再讀「總計」、不 flush，加快每筆速度
 *  3. 加 LockService，多支手機同時寫也不會互相蓋掉（+1 不漏算）
 *
 * 部署：擴充功能 → Apps Script → 貼上 → 部署 → 管理部署 →
 *      編輯(鉛筆) → 版本選「新版本」→ 部署。每次改完都要走這步。
 ******************************************************/

// 若指令碼不是綁在試算表上，填入試算表 ID；綁定就留空
const SHEET_ID = "";

// 時區（台灣固定 +8，無日光節約）
const TZ = "Asia/Taipei";

// 換日點：早上 8 點（要改成別的時間就改這個數字）
const DAY_START_HOUR = 8;

// 各區段辨識特徵：標題列關鍵字 + 一個只屬於該區的欄位
const BLOCKS = {
  "假裝清醒": { header: "品名",          sig: "橘桶" },
  "烘焙雲":   { header: "品名",          sig: "POS冰箱" },
  "酉鬼":     { header: "品名",          sig: "酉鬼內未冰" },
  "備品":     { header: "品名（Key個數）", sig: "小倉(二店)" },
};

function getSS() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// 依「早上8點換日」算出今天的分頁名稱（補零四碼 MMdd）
function businessDayTab() {
  const shifted = new Date(Date.now() - DAY_START_HOUR * 3600 * 1000);
  return Utilities.formatDate(shifted, TZ, "MMdd");  // 補零四碼：0623、0101、0121、1201
}

// 去掉開頭「12.」與結尾括號規格，方便比對品名
function norm(s) {
  return String(s == null ? "" : s)
    .replace(/^\s*\d+\.\s*/, "")
    .replace(/[（(][^）)]*[）)]\s*$/, "")
    .trim();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // uid 狀態查詢：前端重試前先問「這筆處理過了嗎」，避免重複寫入
  //   回傳 status: "done"(已完成,附現值) / "processing"(處理中) / "unknown"(沒看過)
  if (e && e.parameter && e.parameter.checkUid) {
    const cache = CacheService.getScriptCache();
    const rec = cache.get("uid_" + e.parameter.checkUid);
    if (!rec) return json({ status: "unknown" });
    if (rec === "PROCESSING") return json({ status: "processing" });
    // rec 是 JSON 字串 {written, name}
    try {
      const o = JSON.parse(rec);
      return json({ status: "done", written: o.w, name: o.n });
    } catch (x) {
      return json({ status: "done" });
    }
  }

  // reminder 即時讀取：給區段(block)+位置(col)+一批品名(names，用 | 分隔)，
  //   回傳每個品名在當天分頁、該位置那一欄的「原始儲存格值」（空白判定交給前端，
  //   因為雙格式「兩側都要有數字」這種判斷屬於前端的品項設定，後端只單純回報原始值）。
  if (e && e.parameter && e.parameter.remind) {
    const block = e.parameter.block || "";
    const col = e.parameter.col || "";
    const names = (e.parameter.names || "").split("|").filter(Boolean);

    const tab = businessDayTab();
    const sheet = getSS().getSheetByName(tab);
    if (!sheet) return json({ success: false, error: "找不到工作表「" + tab + "」" });

    const blk = BLOCKS[block];
    if (!blk) return json({ success: false, error: "未知區段：" + block });

    const data = sheet.getDataRange().getValues();
    let headerRow = -1;
    for (let r = 0; r < data.length; r++) {
      const a = String(data[r][0] || "");
      const hitHeader = (block === "備品") ? a.indexOf("品名（Key") === 0 : a === blk.header;
      if (hitHeader && data[r].indexOf(blk.sig) !== -1) { headerRow = r; break; }
    }
    if (headerRow === -1) return json({ success: false, error: "找不到「" + block + "」區段標題" });

    const colIdx = data[headerRow].indexOf(col);
    if (colIdx === -1) return json({ success: false, error: "找不到欄位「" + col + "」" });

    const values = {};
    names.forEach(function (name) {
      const target = norm(name);
      let val = "";
      for (let r = headerRow + 1; r < data.length && r < headerRow + 80; r++) {
        const a = String(data[r][0] || "");
        if (a === "品名" || a.indexOf("品名（Key") === 0) break;
        if (norm(a) === target) { val = data[r][colIdx]; break; }
      }
      values[name] = (val === undefined || val === null) ? "" : val;
    });

    return json({ success: true, tab: tab, values: values });
  }

  return json({ ok: true, msg: "盤點後端運作中", today: businessDayTab() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);  // 多支手機不會同時改同一格

    const p = JSON.parse(e.postData.contents);
    const ss = getSS();

    // 0) uid 去重（提前版）：拿到鎖後、動手寫之前就先檢查並標記。
    //    - 已完成(JSON) → 直接回成功並附上當時寫入的現值，絕不重寫
    //    - 處理中(PROCESSING) → 有另一個相同請求正在跑，直接回成功避免雙寫
    //    - 沒看過 → 立刻標記 PROCESSING，佔住位子，接著才真正寫入
    const cache = CacheService.getScriptCache();
    const uid = String(p.uid || "");
    if (uid) {
      const rec = cache.get("uid_" + uid);
      if (rec) {
        if (rec === "PROCESSING") return json({ success: true, dup: true, name: p.name });
        try { const o = JSON.parse(rec); return json({ success: true, dup: true, name: p.name, written: o.w }); }
        catch (x) { return json({ success: true, dup: true, name: p.name }); }
      }
      cache.put("uid_" + uid, "PROCESSING", 21600);   // 佔位，擋住重試的第二筆
    }

    // 業務失敗時用這個回傳：先清掉 PROCESSING 佔位，讓這筆之後能重試（不會被誤當成處理中卡死）
    function fail(msg) {
      if (uid) cache.remove("uid_" + uid);
      return json({ success: false, error: msg });
    }

    // 1) 今天的分頁（後端依 8 點規則決定，忽略前端送來的 date）
    const tab = businessDayTab();
    const sheet = ss.getSheetByName(tab);
    if (!sheet) return fail("找不到工作表「" + tab + "」（請先建立今天的分頁）");

    // 2) 一次讀取整個分頁
    const data = sheet.getDataRange().getValues();
    const blk = BLOCKS[p.block];
    if (!blk) return fail("未知區段：" + p.block);

    let headerRow = -1;
    for (let r = 0; r < data.length; r++) {
      const a = String(data[r][0] || "");
      const hitHeader = (p.block === "備品") ? a.indexOf("品名（Key") === 0 : a === blk.header;
      if (hitHeader && data[r].indexOf(blk.sig) !== -1) { headerRow = r; break; }
    }
    if (headerRow === -1) return fail("找不到「" + p.block + "」區段標題");

    const col = data[headerRow].indexOf(p.writeCol);
    if (col === -1) return fail("找不到欄位「" + p.writeCol + "」");

    const target = norm(p.name);
    let itemRow = -1;
    for (let r = headerRow + 1; r < data.length && r < headerRow + 80; r++) {
      const a = String(data[r][0] || "");
      if (a === "品名" || a.indexOf("品名（Key") === 0) break;
      if (norm(a) === target) { itemRow = r; break; }
    }
    if (itemRow === -1) return fail("在「" + p.block + "」找不到品項「" + p.name + "」");

    // set=覆蓋；add=用剛剛讀到的現值累加（不必再讀一次）
    let newVal;
    if (p.dual) {
      // ===== 雙格式（大/小、假清/烘焙）：同一格填「x/y」，只累加指定那一側 =====
      // 規則：兩側預設 n；本次對 p.dual 那一側 += p.value；另一側保留原值（原本是空則維持 n）
      var slots = p.dualSlots || [];
      var idx = slots.indexOf(p.dual);
      if (idx === -1) return fail("未知的規格側別：" + p.dual);

      var raw = data[itemRow][col];
      var parts = String(raw == null ? "" : raw).split("/");
      // 補足成兩格；空字串或非數字都當成尚未填（用 n 呈現）
      var vals = [];
      for (var i = 0; i < slots.length; i++) {
        var cell = (parts[i] == null ? "" : String(parts[i]).trim());
        vals.push((cell === "" || cell.toLowerCase() === "n" || isNaN(Number(cell))) ? null : Number(cell));
      }
      var base = (vals[idx] == null) ? 0 : vals[idx];
      vals[idx] = base + Number(p.value);
      newVal = vals.map(function (v) { return v == null ? "n" : v; }).join("/");
      var cell = sheet.getRange(itemRow + 1, col + 1);
      cell.setNumberFormat("@");          // 強制純文字，避免 6/3 被 Sheet 誤判成日期或分數
      cell.setValue(newVal);
      if (uid) cache.put("uid_" + uid, JSON.stringify({ w: newVal, n: p.name }), 21600);
      return json({ success: true, name: p.name, written: newVal, date: tab });
    }
    if (p.op === "set") {
      newVal = Number(p.value);
    } else {
      const raw = data[itemRow][col];
      const cur = (raw === "" || raw == null || isNaN(Number(raw))) ? 0 : Number(raw);
      newVal = cur + Number(p.value);
    }

    sheet.getRange(itemRow + 1, col + 1).setValue(newVal);  // 寫完直接回，不 flush、不讀總計

    // 寫入成功 → 記錄已完成 + 現值（6 小時內同 uid 重送/查詢都會拿到這個結果，不再寫一次）
    if (uid) cache.put("uid_" + uid, JSON.stringify({ w: newVal, n: p.name }), 21600);

    return json({ success: true, name: p.name, written: newVal, date: tab });

  } catch (err) {
    // 例外時也要清掉可能殘留的 PROCESSING 佔位，否則那筆會卡住無法重試
    try {
      const pp = JSON.parse(e.postData.contents);
      if (pp && pp.uid) CacheService.getScriptCache().remove("uid_" + String(pp.uid));
    } catch (x) {}
    return json({ success: false, error: "後端錯誤：" + err.message });
  } finally {
    lock.releaseLock();
  }
}


/***** 每天自動建立盤點表 *****
 * 設定方式：Apps Script 左側「觸發條件(時鐘圖示)」→ 新增觸發條件 →
 *   函式：createDailySheet／事件來源：時間驅動／類型：日計時器／時間：凌晨 2~3 點
 *
 * 邏輯：凌晨觸發時，建立「即將到來的那個營業日」分頁（= 今天營業日 + 1）。
 *   例：7/7 03:00 觸發 → 今天營業日是 0706 → 建好 0707（給 7/7 08:00 換日後使用）。
 *   這樣 08:00 換日時分頁已存在，不會有空窗；此時輸入仍寫進舊分頁(0706)不受影響。
 ************************************************/

const TEMPLATE_NAME = "範本(含進銷存)";  // 要複製的範本分頁名稱

// 即將到來的營業日分頁名稱（今天營業日 + 1 天）
function upcomingDayTab() {
  const shifted = new Date(Date.now() - DAY_START_HOUR * 3600 * 1000 + 24 * 3600 * 1000);
  return Utilities.formatDate(shifted, TZ, "MMdd");
}

function createDailySheet() {
  const ss = getSS();
  const tab = upcomingDayTab();            // 即將到來的營業日，例如 7/7 凌晨執行 → 0707
  if (ss.getSheetByName(tab)) return;      // 已存在就不重複建

  const tpl = ss.getSheetByName(TEMPLATE_NAME);
  if (!tpl) throw new Error("找不到範本分頁：" + TEMPLATE_NAME);

  const sheet = tpl.copyTo(ss).setName(tab);   // 連同公式整張複製，分頁名稱改成 MMdd

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(3);                    // 移到前面，好找

  // ---- 把前一天(今天營業日，此刻收班已完成)的「威杯/Highball杯/Shot杯」
  //      在「假裝現場、烘焙現場、小倉」的收班盤，複製到新分頁的開班盤 ----
  // 失敗不影響建表本身（try/catch 包起來），頂多這部分沒帶到、之後人工補
  try {
    const prevTab = businessDayTab();          // 凌晨執行時，這是即將變成「昨天」、收班已完成的那一天
    const prevSheet = ss.getSheetByName(prevTab);
    if (prevSheet) {
      copyClosingToOpening(prevSheet, sheet);
      copySmallPlaClosingToOpening(prevSheet, sheet);
      copyBeerTotalToOpening(prevSheet, sheet);
    }
  } catch (err) {
    // 靜默略過，不讓進銷存複製失敗擋住每日建表
  }
}

// 進銷存區塊共用的正規化比對（大小寫、全形/半形空白皆容錯）
function provNorm(s) {
  return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, "").trim();
}

// 小 pla 杯區塊：表頭本身就是 A欄="小pla杯"（無「盤點分類」欄、無合併儲存格、無「補到現場」列）
//   A=小pla杯 B=假裝現場 C=烘焙現場 D=現場總計(公式) E=一倉 F=總計(公式)
//   下面兩列：A="開班盤"/"收班盤"
const SMALLPLA_TITLE = "小pla杯";
const SMALLPLA_COLS = [1, 2, 4];   // 0-based：B 假裝現場、C 烘焙現場、E 一倉（D、F 是公式，不複製）

function findSimpleHeader(data, title) {
  for (let r = 0; r < data.length; r++) {
    if (provNorm(data[r][0]) === provNorm(title)) return r;
  }
  return -1;
}

function findSimpleRow(data, headerRow, wantLabel) {
  const limit = Math.min(data.length, headerRow + 6);
  for (let r = headerRow + 1; r < limit; r++) {
    if (provNorm(data[r][0]) === provNorm(wantLabel)) return r;
  }
  return -1;
}

function copySmallPlaClosingToOpening(prevSheet, newSheet) {
  const prevData = prevSheet.getDataRange().getValues();
  const newData = newSheet.getDataRange().getValues();

  const hPrev = findSimpleHeader(prevData, SMALLPLA_TITLE);
  const hNew = findSimpleHeader(newData, SMALLPLA_TITLE);
  if (hPrev === -1 || hNew === -1) return;

  const srcRow = findSimpleRow(prevData, hPrev, "收班盤");
  const dstRow = findSimpleRow(newData, hNew, "開班盤");
  if (srcRow === -1 || dstRow === -1) return;

  SMALLPLA_COLS.forEach(function (c) {
    const v = prevData[srcRow][c];
    if (v === "" || v === null || v === undefined) return;   // 來源空白 → 開班盤留空
    newSheet.getRange(dstRow + 1, c + 1).setValue(v);
  });
}

// ---- 啤酒總量複製：把前一天酉鬼區塊「總計」欄（秤重換算後的即時庫存數，公式算出的值）
//      複製到新分頁對應品項的「營業前庫存」欄 ----
// 範圍：YG001~YG006（酉鬼區塊裡除了「空桶」以外的品項）。「空桶」沒有「營業前庫存」欄，略過不複製。
// 「總計」是公式（=SUM(位置欄)），但 getValues() 拿到的是算好的數值不是公式本身，所以直接寫數值即可，
// 不會有複製公式導致相對位置跑掉的問題。
const BEER_EXCLUDE_NAME = "空桶";

// 找酉鬼區塊表頭列（沿用 BLOCKS["酉鬼"] 的辨識特徵：header="品名" 且該列含 sig="酉鬼內未冰"）
function findBeerHeader(data) {
  const blk = BLOCKS["酉鬼"];
  for (let r = 0; r < data.length; r++) {
    const a = String(data[r][0] || "");
    if (a === blk.header && data[r].indexOf(blk.sig) !== -1) return r;
  }
  return -1;
}

// 在表頭之後找指定品名所在列（品名已用 norm() 正規化去掉開頭編號），遇到下一個「品名」表頭列就停止
function findBeerRow(data, headerRow, wantName) {
  const limit = Math.min(data.length, headerRow + 20);
  for (let r = headerRow + 1; r < limit; r++) {
    const a = String(data[r][0] || "");
    if (a === "品名") break;
    if (norm(a) === wantName) return r;
  }
  return -1;
}

function copyBeerTotalToOpening(prevSheet, newSheet) {
  const prevData = prevSheet.getDataRange().getValues();
  const newData = newSheet.getDataRange().getValues();

  const hPrev = findBeerHeader(prevData);
  const hNew = findBeerHeader(newData);
  if (hPrev === -1 || hNew === -1) return;

  const totalCol = prevData[hPrev].indexOf("總計");
  const openCol = newData[hNew].indexOf("營業前庫存");
  if (totalCol === -1 || openCol === -1) return;

  const limit = Math.min(prevData.length, hPrev + 20);
  for (let r = hPrev + 1; r < limit; r++) {
    const rawName = String(prevData[r][0] || "");
    if (rawName === "品名") break;                          // 下一個區塊表頭，結束
    const name = norm(rawName);
    if (!name || name === BEER_EXCLUDE_NAME) continue;      // 空白列或「空桶」不複製

    const dstRow = findBeerRow(newData, hNew, name);
    if (dstRow === -1) continue;

    const v = prevData[r][totalCol];
    if (v === "" || v === null || v === undefined) continue;
    const n = Number(v);
    newSheet.getRange(dstRow + 1, openCol + 1).setValue(isNaN(n) ? v : Math.round(n * 100) / 100);
  }
}

// 進銷存區塊：品名（合併儲存格）/ 盤點分類（開班盤·補到現場·收班盤）/ 假裝現場 / 烘焙現場 / 現場總計 / 小倉 / 全部總計 …
const PROVISION_ITEMS = ["威杯", "highball杯", "shot杯"];   // 大小寫、全形/半形空白皆容錯比對
const PROVISION_COLS = [2, 3, 5];                          // 0-based：C 假裝現場、D 烘焙現場、F 小倉

// 找「品名／盤點分類」表頭列（0-based row index）
function findProvHeader(data) {
  for (let r = 0; r < data.length; r++) {
    if (provNorm(data[r][0]) === "品名" && provNorm(data[r][1]) === "盤點分類") return r;
  }
  return -1;
}

// 在表頭之後找指定品項＋指定列標籤（開班盤/收班盤）所在列（0-based）。
// 品名是合併儲存格，只有起始列有值，往下延用「目前品名」直到遇到下一個非空 A 欄。
function findProvRow(data, headerRow, wantName, wantLabel) {
  let currentName = "";
  const limit = Math.min(data.length, headerRow + 40);
  for (let r = headerRow + 1; r < limit; r++) {
    const a = String(data[r][0] || "").trim();
    if (a !== "") currentName = a;
    if (provNorm(currentName) === provNorm(wantName) && provNorm(data[r][1]) === provNorm(wantLabel)) return r;
  }
  return -1;
}

// 把 prevSheet 的收班盤（三品項 × 三位置）複製到 newSheet 的開班盤。
// 來源格是空白就跳過、不寫（開班盤維持範本原本的空白，不會寫入 0）。
function copyClosingToOpening(prevSheet, newSheet) {
  const prevData = prevSheet.getDataRange().getValues();
  const newData = newSheet.getDataRange().getValues();

  const hPrev = findProvHeader(prevData);
  const hNew = findProvHeader(newData);
  if (hPrev === -1 || hNew === -1) return;   // 找不到表頭就整段跳過

  PROVISION_ITEMS.forEach(function (name) {
    const srcRow = findProvRow(prevData, hPrev, name, "收班盤");
    const dstRow = findProvRow(newData, hNew, name, "開班盤");
    if (srcRow === -1 || dstRow === -1) return;   // 找不到就跳過這個品項，不報錯

    PROVISION_COLS.forEach(function (c) {
      const v = prevData[srcRow][c];
      if (v === "" || v === null || v === undefined) return;   // 來源空白 → 開班盤留空
      newSheet.getRange(dstRow + 1, c + 1).setValue(v);
    });
  });
}