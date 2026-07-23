/***** 易開罐盤點表 — 後端 (Google Apps Script)  v1 *****
 *
 * 這是「易開罐盤點表」專用的獨立後端，跟日盤表的後端各自獨立、互不影響。
 * 部署在「易開罐」那個 Google 試算表上（跟日盤表不是同一個檔案）。
 *
 * 功能：
 *  1. 手機掃到 CA001~CA004 的條碼 → 寫進易開罐盤點表對應位置/品項/時段的格子
 *  2. 時段判斷（台灣時間）：
 *       - 換日點 08:00：例如分頁「0707」代表 7/7 08:00 ~ 7/8 07:59
 *       - 盤點時段：08:00~22:00 寫「開場」列；22:00~隔天08:00 寫「收班盤點」列
 *  3. 自動找格子：靠「位置區塊標題 + 品名字串」定位，2 品項或 4 品項的範本都能自動適應
 *  4. LockService 防並發、CacheService 用 uid 去重（重試不會重複累加）
 *  5. 每天 08:00 後自動複製範本建立當日分頁（分頁名 MMdd）
 *
 * 部署：擴充功能 → Apps Script → 貼上 → 部署 → 新增部署作業 → 類型「網頁應用程式」
 *      → 執行身分：我／存取權：所有人 → 部署，複製網址填到 index.html 的 CAN_SCRIPT_URL。
 *      之後每次改都要「管理部署 → 編輯 → 新版本 → 部署」。
 ******************************************************/

/* ============================================================
 *  ★★★ 需要你填的兩個地方 ★★★
 * ============================================================ */

// 【第 34 行】易開罐試算表的 ID（網址 /d/ 後面那段）。
//   若這支 Apps Script 是「綁定」在易開罐試算表上（從該表的擴充功能開啟），可留空 ""。
//   若是獨立指令碼，必須填 ID。
const SHEET_ID = "";

// 【第 39 行】易開罐的「範本」分頁名稱（自動建立每日分頁時要複製的那一頁）。
//   請填你易開罐試算表裡範本頁的實際名稱。
const TEMPLATE_NAME = "「範本」";

/* ============================================================
 *  以下一般不需要改
 * ============================================================ */

const TZ = "Asia/Taipei";
const DAY_START_HOUR = 8;    // 換日點：早上 8 點
const SHIFT_SWITCH_HOUR = 22; // 盤點時段分界：晚上 10 點（之後算收班）

// 品項代碼 → 試算表上的品名字串（比對時會正規化，斜線/空白差異不影響）
const CAN_NAMES = {
  "CA001": "黑 / 洛神",
  "CA002": "綠 / 桂花",
  "CA003": "藍／白桃",
  "CA004": "粉／荔枝"
};

function getSS() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// 正規化：全形斜線→半形、去空白，讓「藍／白桃」和「藍 / 白桃」視為相同
function norm(s) {
  return String(s == null ? "" : s).replace(/／/g, "/").replace(/\s+/g, "").trim();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 依「早上8點換日」算出今天的分頁名稱（補零四碼 MMdd）
function businessDayTab() {
  const shifted = new Date(Date.now() - DAY_START_HOUR * 3600 * 1000);
  return Utilities.formatDate(shifted, TZ, "MMdd");
}

// 判斷現在該寫「開場」還是「收班盤點」
//  08:00~21:59 → 開場；22:00~隔天07:59 → 收班盤點
function currentRowLabel() {
  const h = Number(Utilities.formatDate(new Date(), TZ, "H"));
  return (h >= DAY_START_HOUR && h < SHIFT_SWITCH_HOUR) ? "開場" : "收班盤點";
}

function doGet(e) {
  // uid 狀態查詢：前端重試前先問「這筆處理過了嗎」
  if (e && e.parameter && e.parameter.checkUid) {
    const cache = CacheService.getScriptCache();
    const rec = cache.get("uid_" + e.parameter.checkUid);
    if (!rec) return json({ status: "unknown" });
    if (rec === "PROCESSING") return json({ status: "processing" });
    try { const o = JSON.parse(rec); return json({ status: "done", written: o.w, name: o.n }); }
    catch (x) { return json({ status: "done" }); }
  }
  return json({ ok: true, msg: "易開罐後端運作中", today: businessDayTab(), row: currentRowLabel() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);

    const p = JSON.parse(e.postData.contents);
    const ss = getSS();

    // uid 去重（提前版）：拿到鎖後、動手寫之前就檢查並標記，防止逾時重試造成雙寫
    const cache = CacheService.getScriptCache();
    const uid = String(p.uid || "");
    if (uid) {
      const rec = cache.get("uid_" + uid);
      if (rec) {
        if (rec === "PROCESSING") return json({ success: true, dup: true, name: p.name });
        try { const o = JSON.parse(rec); return json({ success: true, dup: true, name: p.name, written: o.w }); }
        catch (x) { return json({ success: true, dup: true, name: p.name }); }
      }
      cache.put("uid_" + uid, "PROCESSING", 21600);
    }
    // 業務失敗時先清掉 PROCESSING 佔位，讓這筆之後能重試
    function fail(msg) {
      if (uid) cache.remove("uid_" + uid);
      return json({ success: false, error: msg });
    }

    // 今天的分頁
    const tab = businessDayTab();
    const sheet = ss.getSheetByName(tab);
    if (!sheet) return fail("找不到工作表「" + tab + "」（今天的分頁尚未建立）");

    // 品名（後端自己依代碼決定，不信任前端傳來的字串）
    const wantName = CAN_NAMES[p.sku];
    if (!wantName) return fail("未知的易開罐品項代碼：" + p.sku);

    // 時段（後端用台灣時間判斷，前端不參與）
    const rowLabel = currentRowLabel();

    // 讀整張表定位
    const data = sheet.getDataRange().getValues();
    const loc = locateCell(data, p.writeCol, wantName, rowLabel);
    if (loc.error) return fail(loc.error);

    // 累加（易開罐一律累加；缺格當 0）
    const raw = data[loc.r][loc.c];
    const cur = (raw === "" || raw == null || isNaN(Number(raw))) ? 0 : Number(raw);
    const newVal = cur + Number(p.value);
    sheet.getRange(loc.r + 1, loc.c + 1).setValue(newVal);

    if (uid) cache.put("uid_" + uid, JSON.stringify({ w: newVal, n: p.name }), 21600);
    return json({ success: true, name: p.name + "（" + rowLabel + "）", written: newVal, date: tab });

  } catch (err) {
    try {
      const pp = JSON.parse(e.postData.contents);
      if (pp && pp.uid) CacheService.getScriptCache().remove("uid_" + String(pp.uid));
    } catch (x) {}
    return json({ success: false, error: "後端錯誤：" + err.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 定位：先找位置區塊標題（整表唯一），標題下一列是「Can 品項」品名列，
 * 在品名列找到目標品名的欄，再往下找「開場」或「收班盤點」列。
 * 這樣不論 2 品項或 4 品項、不論左右半邊，都能自動找到正確格子。
 * @return {r, c} 皆為 0-based；或 {error}
 */
function locateCell(data, title, wantName, rowLabel) {
  const wantTitle = norm(title);
  const isIce = wantTitle.indexOf("儲冰槽") === 0;   // 儲冰槽標題含括號，用開頭比對
  const nName = norm(wantName);
  const nCanHdr = norm("Can 品項");
  const nRowLabel = norm(rowLabel);

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cell = norm(data[r][c]);
      const titleHit = isIce ? (cell.indexOf("儲冰槽") === 0) : (cell === wantTitle);
      if (!titleHit) continue;

      // 標題下一列必須是 Can 品項（確認這是位置區塊，不是剛好同名的其他格）
      const nameRow = r + 1;
      if (nameRow >= data.length || norm(data[nameRow][c]) !== nCanHdr) continue;

      // 在品名列往右找目標品名
      let col = -1;
      for (let cc = c + 1; cc < data[nameRow].length; cc++) {
        if (norm(data[nameRow][cc]) === nName) { col = cc; break; }
      }
      if (col === -1) return { error: "在「" + title + "」找不到品項「" + wantName + "」（可能範本品項數不符）" };

      // 往下找 開場 / 收班盤點 列
      for (let rr = nameRow + 1; rr < Math.min(nameRow + 4, data.length); rr++) {
        if (norm(data[rr][c]).indexOf(nRowLabel) === 0) return { r: rr, c: col };
      }
      return { error: "在「" + title + "」找不到「" + rowLabel + "」列" };
    }
  }
  return { error: "找不到位置「" + title + "」" };
}


/***** 每天自動建立易開罐盤點表 *****
 * 設定：Apps Script 左側「觸發條件(時鐘)」→ 新增 →
 *   函式：createDailySheet／時間驅動／日計時器／凌晨 2~3 點
 *
 * 邏輯：凌晨觸發時建立「即將到來的營業日」分頁（= 今天營業日 + 1）。
 *   例：7/7 03:00 觸發 → 建好 0707（給 7/7 08:00 換日後使用）。
 ************************************************/

// 即將到來的營業日分頁名稱（今天營業日 + 1 天）
function upcomingDayTab() {
  const shifted = new Date(Date.now() - DAY_START_HOUR * 3600 * 1000 + 24 * 3600 * 1000);
  return Utilities.formatDate(shifted, TZ, "MMdd");
}

function createDailySheet() {
  const ss = getSS();
  const tab = upcomingDayTab();
  if (ss.getSheetByName(tab)) return;

  const tpl = ss.getSheetByName(TEMPLATE_NAME);
  if (!tpl) throw new Error("找不到範本分頁：" + TEMPLATE_NAME);

  const sheet = tpl.copyTo(ss).setName(tab);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(3);   // 移到最前面
}