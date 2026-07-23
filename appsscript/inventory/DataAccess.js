/**
 * ============================================================
 *  資料存取層：定位日期分頁、解析每個區塊的品項與數值
 * ============================================================
 */

/** 取得來源(日盤表)試算表物件。 */
function ccGetSourceSS_() {
  return SpreadsheetApp.openById(CC_CONFIG.SOURCE_SPREADSHEET_ID);
}

/** 取得輸出(比較結果)試算表物件。 */
function ccGetOutputSS_() {
  return SpreadsheetApp.openById(CC_CONFIG.OUTPUT_SPREADSHEET_ID);
}

/**
 * 一次讀出來源所有分頁名稱，後續定位都用它，避免重複呼叫 API。
 * 回傳 Set 方便 O(1) 查找。
 */
function ccListSourceTabNames_(ss) {
  ss = ss || ccGetSourceSS_();
  return new Set(ss.getSheets().map(function (s) { return s.getName(); }));
}

/**
 * 由日期物件產生「候選分頁名稱」清單（因為命名歷史不一致）。
 * 優先順序：4碼 MMDD (新格式) → 不補零 M+D (舊格式) → 混合格式。
 * 例：2025/07/02 → ['0702','72','072','0702']；2025/06/25 → ['0625','625']
 */
function ccCandidateTabNames_(date) {
  const tz = CC_CONFIG.TIMEZONE;
  const m = Number(Utilities.formatDate(date, tz, 'M'));   // 1..12
  const d = Number(Utilities.formatDate(date, tz, 'd'));   // 1..31
  const p2 = function (n) { return (n < 10 ? '0' : '') + n; };
  const set = [];
  const push = function (v) { if (set.indexOf(v) === -1) set.push(v); };
  push(p2(m) + p2(d));     // 0702  ← 新格式，最優先
  push('' + m + d);        // 72 / 625 / 611  ← 舊格式
  push(p2(m) + '' + d);    // 072
  push('' + m + p2(d));    // 702
  return set;
}

/**
 * 找出某一天對應的分頁名稱（存在才回傳），找不到回傳 null。
 * @return {string|null}
 */
function ccResolveTabName_(date, tabSet) {
  const cands = ccCandidateTabNames_(date);
  for (var i = 0; i < cands.length; i++) {
    if (tabSet.has(cands[i])) return cands[i];
  }
  return null;
}

/** 把 Sheet 的整個資料範圍讀成二維陣列（含快取，一次 run 內不重讀）。 */
const __ccGridCache = {};
function ccReadGrid_(ss, tabName) {
  const cacheKey = tabName;
  if (__ccGridCache[cacheKey]) return __ccGridCache[cacheKey];
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return null;
  const grid = sheet.getDataRange().getValues(); // 0-indexed [row][col]
  __ccGridCache[cacheKey] = grid;
  return grid;
}

/** 清掉快取（每次完整 run 開始時呼叫）。 */
function ccClearGridCache_() {
  for (var k in __ccGridCache) delete __ccGridCache[k];
}

function ccIsNum_(v) { return typeof v === 'number' && isFinite(v); }

function ccNormName_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * 把儲存格的原始值解析成數字，支援兩種特殊格式（其餘情況等同一般數字判斷）：
 *   'n/m' （斜線，例：一店/二店分別的量）→ 回傳 n + m（消耗看合計）
 *   'n(m)'（括號，例：331(165.5)，m 只是方便計算用的半值）→ 回傳 n，m 忽略
 * 純數字或無法辨識的字串一律走原本邏輯（不是數字就回傳 null）。
 */
function ccParseNum_(raw) {
  if (typeof raw === 'number' && isFinite(raw)) return raw;
  if (typeof raw !== 'string') return null;
  var s = raw.trim();
  if (s === '') return null;

  var slash = s.match(/^(-?\d+(?:\.\d+)?)\s*\/\s*(-?\d+(?:\.\d+)?)$/);
  if (slash) return Number(slash[1]) + Number(slash[2]);

  var paren = s.match(/^(-?\d+(?:\.\d+)?)\s*[\(（]\s*-?\d+(?:\.\d+)?\s*[\)）]$/);
  if (paren) return Number(paren[1]);

  var n = Number(s);
  return isFinite(n) ? n : null;
}

/**
 * 解析一整張日盤分頁，回傳：
 *   { JZ: {items:{normName:{name, value, total, mild, peak}}, order:[...]}, YG:{...}, ... }
 * DIRECT 區塊 value = 該區「valueHeader」欄的數值；
 * CROSS_DAY_DIFF 區塊 value = 該區「總計」欄的數值（消耗留待 Compare 層用跨日差算）。
 * 若 sec 設定了 totalHeader / stockPrepMildTitle / stockPrepPeakTitle（目前只有 JZ、HB），
 * 每個品項還會多帶：
 *   total : 該列「總計」欄數值（目前庫存）
 *   mild  : [一二, 三四日, 五六] 淡季最低備量（找不到就是 null）
 *   peak  : [一二, 三四日, 五六] 旺季最低備量（找不到就是 null）
 */
function ccParseDailySheet_(grid) {
  const result = {};
  if (!grid || !grid.length) return result;

  const titleColIdx = CC_CONFIG.SECTION_TITLE_COL - 1;
  const nrows = grid.length;

  // 先找出每個區塊標題所在的列
  const titleToRow = {};
  const titleSet = {};
  CC_CONFIG.SECTIONS.forEach(function (sec) { titleSet[sec.title] = true; });
  for (var r = 0; r < nrows; r++) {
    var cell = grid[r][titleColIdx];
    if (cell == null) continue;
    var t = String(cell).trim();
    if (!titleSet[t]) continue;
    // 真正的區塊標題那列，A 欄一定是空的；
    // 這樣可排除「備品表頭列」裡剛好也叫『酉鬼』的庫位欄名（該列 A 欄是品名，非空）。
    var aBlank = (grid[r][0] == null || String(grid[r][0]).trim() === '');
    if (!aBlank) continue;
    if (!(t in titleToRow)) titleToRow[t] = r; // 取第一個出現的列
  }

  CC_CONFIG.SECTIONS.forEach(function (sec) {
    const secResult = { items: {}, order: [], mode: sec.mode };
    result[sec.key] = secResult;

    const titleRow = titleToRow[sec.title];
    if (titleRow == null) return; // 這張表沒有這個區塊，跳過

    // 找表頭列：標題之後第一列，A欄以「品名」開頭
    var headerRow = -1;
    for (var r = titleRow; r < Math.min(titleRow + 5, nrows); r++) {
      var a = grid[r][0];
      if (a != null && String(a).indexOf(CC_CONFIG.NAME_HEADER_PREFIX) === 0) { headerRow = r; break; }
    }
    if (headerRow === -1) return;

    // 在表頭列裡找「品名欄」「數值欄」「總計欄」
    var nameCol = -1, valueCol = -1, totalCol = -1;
    for (var c = 0; c < grid[headerRow].length; c++) {
      var h = grid[headerRow][c];
      if (h == null) continue;
      var ht = String(h).trim();
      if (nameCol === -1 && ht.indexOf(CC_CONFIG.NAME_HEADER_PREFIX) === 0) nameCol = c;
      if (valueCol === -1 && ht === sec.valueHeader) valueCol = c;
      if (sec.totalHeader && totalCol === -1 && ht === sec.totalHeader) totalCol = c;
    }
    if (nameCol === -1 || valueCol === -1) return;

    // 在標題列（titleRow）裡找「最低備量(淡季)」「最低備量(旺季)」起始欄（後面接三欄：一二/三四日/五六）
    var mildStart = -1, peakStart = -1;
    if (sec.stockPrepMildTitle || sec.stockPrepPeakTitle) {
      for (var c2 = 0; c2 < grid[titleRow].length; c2++) {
        var th = grid[titleRow][c2];
        if (th == null) continue;
        var tht = String(th).trim();
        if (sec.stockPrepMildTitle && tht === sec.stockPrepMildTitle) mildStart = c2;
        if (sec.stockPrepPeakTitle && tht === sec.stockPrepPeakTitle) peakStart = c2;
      }
    }

    // 從表頭下一列開始，讀到：A欄空白 或 撞到下一個區塊標題
    for (var r = headerRow + 1; r < nrows; r++) {
      var nm = grid[r][nameCol];
      var titleCell = grid[r][titleColIdx];
      if (titleCell != null && titleSet[String(titleCell).trim()]) break; // 下一區塊
      if (nm == null || String(nm).trim() === '') break;                  // 區塊結束
      var key = ccNormName_(nm);
      if (!key) continue;
      var val = grid[r][valueCol];

      var item = { name: String(nm).trim(), value: ccParseNum_(val) };

      if (totalCol !== -1) {
        var tv = grid[r][totalCol];
        item.total = ccParseNum_(tv);
      }
      if (mildStart !== -1) {
        item.mild = [grid[r][mildStart], grid[r][mildStart + 1], grid[r][mildStart + 2]]
          .map(function (v) { return ccParseNum_(v); });
      }
      if (peakStart !== -1) {
        item.peak = [grid[r][peakStart], grid[r][peakStart + 1], grid[r][peakStart + 2]]
          .map(function (v) { return ccParseNum_(v); });
      }

      secResult.items[key] = item;
      secResult.order.push(key);
    }
  });

  return result;
}