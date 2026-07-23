/**
 * ============================================================
 *  共用小工具
 * ============================================================
 */

/** 日期加減天數，回傳新的 Date（不改動原物件）。 */
function ccAddDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

/**
 * 依「營業日換日點」(CC_CONFIG.DAY_START_HOUR) 算出「現在算哪個營業日」。
 * 例：換日點是早上 8 點，現在是凌晨 0:30 → 回傳的日期會落在「昨天」，
 * 跟你掃碼盤點後端 businessDayTab() 的邏輯完全對齊。
 * @param {Date} [date]  基準時間，預設現在
 * @return {Date}
 */
function ccBusinessDate_(date) {
  const base = date || new Date();
  return new Date(base.getTime() - CC_CONFIG.DAY_START_HOUR * 3600 * 1000);
}

/** 給人看的日期字串，例：2025/07/09。 */
function ccFmtDate_(date) {
  return Utilities.formatDate(date, CC_CONFIG.TIMEZONE, 'yyyy/MM/dd');
}

/** 四捨五入到小數第二位。 */
function ccRound2_(n) {
  return Math.round(n * 100) / 100;
}

/** 由分頁名稱反推日期（僅供 ALL_FILLED / 手動指定用）。
 *  支援 4碼 MMDD 與不補零 M+D 兩種；跨年以「今天的年份」為準，
 *  若推出來的日期比今天晚超過 300 天，視為去年。
 */
function ccDateFromTabName_(tabName) {
  const clean = String(tabName).replace(/（\d+）/g, '').trim();
  if (!/^\d{2,4}$/.test(clean)) return null;

  const now = new Date();
  const year = Number(Utilities.formatDate(now, CC_CONFIG.TIMEZONE, 'yyyy'));
  let m, d;

  if (clean.length === 4) {            // MMDD
    m = Number(clean.slice(0, 2)); d = Number(clean.slice(2));
  } else if (clean.length === 3) {     // M DD  (e.g. 625, 611)  → 舊格式，月為單碼
    m = Number(clean.slice(0, 1)); d = Number(clean.slice(1));
    if (d > 31) { m = Number(clean.slice(0, 2)); d = Number(clean.slice(2)); }
  } else {                              // 2碼 MD (e.g. 76 = 7/6)
    m = Number(clean.slice(0, 1)); d = Number(clean.slice(1));
  }
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;

  let date = new Date(year, m - 1, d);
  // 跨年修正：若算出來遠在未來，往前一年
  if ((date.getTime() - now.getTime()) > 300 * 86400000) {
    date = new Date(year - 1, m - 1, d);
  }
  return date;
}
