/**
 * ============================================================
 *  比較引擎
 * ------------------------------------------------------------
 *  1) 今日消耗 vs 每個比較欄背後那組週數的「可用樣本平均」，各欄各自評級
 *     （-1w 欄只有 1 週，等於直接比那天的值；-3w 欄是 -1w/-2w/-3w 三週平均）
 *  2) JZ / HB 的「今日備量」：隔天(依星期分組)的最低備量 − 目前總計，
 *     四捨五入取整、負值顯示 0（淡季/旺季各算一次）
 * ============================================================
 */

/**
 * 針對某個目標日期，計算完整比較結果。
 * @param {Date} targetDate  目標日（通常是今天）
 * @return {{date:Date, rows:Array, missingTabs:Array<string>}}
 *   rows[i] = {
 *     section, name, today,
 *     cols: [ {header, value, grade, pct, samples}, ... ],  // 對應 CC_CONFIG.COMPARE_COLUMNS 順序
 *     prepMild: number|null, prepPeak: number|null
 *   }
 */
function ccBuildComparison_(targetDate) {
  ccClearGridCache_();
  const ss = ccGetSourceSS_();
  const tabSet = ccListSourceTabNames_(ss);

  function parsedOf(date) {
    const name = ccResolveTabName_(date, tabSet);
    if (!name) return null;
    const grid = ccReadGrid_(ss, name);
    return grid ? ccParseDailySheet_(grid) : null;
  }

  const parsedCache = {};
  function parsedCached(date) {
    const k = +date;
    if (!(k in parsedCache)) parsedCache[k] = parsedOf(date);
    return parsedCache[k];
  }

  /**
   * 取得某區塊某品項在某日的「消耗值」。
   *   DIRECT         → 直接讀 value（負值視 DROP_NEGATIVE_DIRECT 設定剔除）
   *   CROSS_DAY_DIFF → 前一日總計 − 當日總計（負值=補貨；歷史樣本仍剔除，
   *                    但傳 opts.keepNegative=true 時原樣回傳負值，供「今天」顯示用）
   */
  function consumptionOf(date, secKey, itemKey, opts) {
    const parsed = parsedCached(date);
    if (!parsed || !parsed[secKey]) return null;
    const item = parsed[secKey].items[itemKey];
    if (!item) return null;

    if (parsed[secKey].mode !== 'CROSS_DAY_DIFF') {
      if (!ccIsNum_(item.value)) return null;
      if (CC_CONFIG.DROP_NEGATIVE_DIRECT && item.value < 0) return null;
      return item.value;
    }
    const prevParsed = parsedCached(ccAddDays_(date, -1));
    if (!prevParsed || !prevParsed[secKey]) return null;
    const prevItem = prevParsed[secKey].items[itemKey];
    if (!prevItem || !ccIsNum_(prevItem.value) || !ccIsNum_(item.value)) return null;
    const used = prevItem.value - item.value;
    if (used < 0 && !(opts && opts.keepNegative) && CC_CONFIG.BP_NEGATIVE_AS_RESTOCK) return null;
    return used;
  }

  const todayParsed = parsedCached(targetDate);
  if (!todayParsed) {
    throw new Error('找不到今天(' + ccFmtDate_(targetDate) + ')的日盤分頁，無法比較。');
  }

  // 收集這次會用到的所有週數（去重），每個週數對應的分頁日期／是否存在，只算一次
  const allWeeks = [];
  CC_CONFIG.COMPARE_COLUMNS.forEach(function (col) {
    col.weeks.forEach(function (w) { if (allWeeks.indexOf(w) === -1) allWeeks.push(w); });
  });
  const weekDates = {};   // { 1: Date|null, 2: Date|null, 3: Date|null, ... }
  const missingTabs = [];
  allWeeks.forEach(function (w) {
    const pd = ccAddDays_(targetDate, -7 * w);
    const tab = ccResolveTabName_(pd, tabSet);
    weekDates[w] = tab ? pd : null;
    if (!tab) missingTabs.push('-' + w + 'w(' + ccFmtDate_(pd) + ')');
  });

  // 隔天所屬的星期分組（決定用最低備量的哪一欄）
  const tomorrowDow = ccAddDays_(targetDate, 1).getDay(); // 0=日...6=六
  const groupIdx = CC_CONFIG.STOCK_PREP_GROUP_BY_DOW[tomorrowDow];

  const rows = [];
  CC_CONFIG.SECTIONS.forEach(function (sec) {
    const todaySec = todayParsed[sec.key];
    if (!todaySec) return;

    todaySec.order.forEach(function (itemKey) {
      const itemToday = todaySec.items[itemKey];
      const today = consumptionOf(targetDate, sec.key, itemKey, { keepNegative: true });

      const cols = CC_CONFIG.COMPARE_COLUMNS.map(function (colDef) {
        const samples = [];
        colDef.weeks.forEach(function (w) {
          const pd = weekDates[w];
          if (!pd) return;
          const v = consumptionOf(pd, sec.key, itemKey);
          if (ccIsNum_(v)) samples.push(v);
        });
        const value = samples.length
          ? samples.reduce(function (a, b) { return a + b; }, 0) / samples.length
          : null;
        const g = ccGradeOfSingle_(today, value);
        return { header: colDef.header, value: value, grade: g.grade, pct: g.pct, samples: samples.length };
      });

      // 備品當日消耗算出負值（通常是補貨）→ 不套用一般評級，整列直接標「變多了!」
      if (sec.mode === 'CROSS_DAY_DIFF' && ccIsNum_(today) && today < 0) {
        cols.forEach(function (c) { c.grade = CC_CONFIG.GRADE_LABELS.INCREASED; c.pct = null; });
      }

      // 今日備量（只有設定了 stockPrepMildTitle/PeakTitle 的區塊才算，即 JZ、HB）
      var prepMild = null, prepPeak = null;
      if ((sec.stockPrepMildTitle || sec.stockPrepPeakTitle) && ccIsNum_(itemToday.total)) {
        if (itemToday.mild && ccIsNum_(itemToday.mild[groupIdx])) {
          prepMild = ccClampRound0_(itemToday.mild[groupIdx] - itemToday.total);
        }
        if (itemToday.peak && ccIsNum_(itemToday.peak[groupIdx])) {
          prepPeak = ccClampRound0_(itemToday.peak[groupIdx] - itemToday.total);
        }
      }

      rows.push({
        section: sec.key,
        name: itemToday.name,
        today: today,
        cols: cols,
        prepMild: prepMild,
        prepPeak: prepPeak,
      });
    });
  });

  return { date: targetDate, rows: rows, missingTabs: missingTabs };
}

/**
 * 今日值 vs 單一基準值（可能是單週值，也可能是多週平均） 的評級。
 * @param {?number} today
 * @param {?number} baseline
 * @return {{grade:string, pct:?number}}
 */
function ccGradeOfSingle_(today, baseline) {
  const L = CC_CONFIG.GRADE_LABELS;
  if (!ccIsNum_(today)) return { grade: L.MISSING, pct: null };
  if (!ccIsNum_(baseline)) return { grade: L.NODATA, pct: null };

  if (baseline === 0) {
    return today === 0 ? { grade: L.NORMAL, pct: 0 } : { grade: L.ERROR, pct: null };
  }
  const pct = (today - baseline) / baseline;
  const a = Math.abs(pct);
  var grade;
  if (a <= CC_CONFIG.GRADE_NORMAL_PCT) grade = L.NORMAL;
  else if (a <= CC_CONFIG.GRADE_CAUTION_PCT) grade = L.CAUTION;
  else grade = L.ERROR;
  return { grade: grade, pct: pct };
}

/** 二捨三入取整（只要有小數餘數就無條件進位）；負值顯示 0。 */
function ccClampRound0_(n) {
  const EPS = 1e-9; // 容錯：避免浮點數誤差（例如 3.0000000001）被誤判成有小數而多進位
  const nearestInt = Math.round(n);
  const r = Math.abs(n - nearestInt) < EPS ? nearestInt : Math.ceil(n);
  return r < 0 ? 0 : r;
}