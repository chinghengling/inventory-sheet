/**
 * ============================================================
 *  盤點消耗比較系統  ─  設定檔
 * ============================================================
 *  這是唯一你「一定要改」的檔案。把兩個試算表 ID 填進去，
 *  其餘參數維持預設即可運作，需要微調再動。
 * ------------------------------------------------------------
 */
const CC_CONFIG = {

  /* ── 必填：兩個試算表的 ID ──────────────────────────────── */
  // 日盤表（來源，就是你上傳的那張多分頁的表）
  SOURCE_SPREADSHEET_ID: '1S53wjfr0k6zOYbXtd6J-88EA0V5v-kdWcaK67oplPZE',
  // 比較結果要寫進去的「新試算表」
  OUTPUT_SPREADSHEET_ID: '1i1z3GMZpNEmZ31QOm84FKf7NfOEI-harCgLYkOyaY3U',

  TIMEZONE: 'Asia/Taipei',

  // 營業日換日點：早上 8 點前執行，「今天」視為前一個營業日（跟你掃碼盤點後端的 DAY_START_HOUR 對齊）
  DAY_START_HOUR: 8,

  /* ── 比較邏輯：每一欄背後對應「一組週數」，取這組週數裡可取得的樣本平均 ──
   * -1w 欄只有 1 週（等於直接顯示那天的值）；
   * -3w 欄是 -1w、-2w、-3w 三週的平均（哪幾週有資料就平均哪幾週，不是硬性要求三週都在）。
   * 要調整成別的組合，直接改這個陣列就好，欄位標題、評級門檻都會自動套用。
   */
  COMPARE_COLUMNS: [
    { header: '-1w 比較', weeks: [1] },
    { header: '-3w 比較', weeks: [1, 2, 3] },
  ],

  // 備品區用跨日差計算消耗；負值(=補貨日)是否視為無效、不拿來比較
  BP_NEGATIVE_AS_RESTOCK: true,

  // 三大區的「當日消耗」若為負，多半是那天「營業前庫存」沒填造成的無效值，
  // 預設剔除、不納入比較（強烈建議保持 true）
  DROP_NEGATIVE_DIRECT: true,

  /* ── 評級門檻：今日 vs 該欄基準值（單週值或多週平均）的變化% ─────── */
  GRADE_NORMAL_PCT: 0.10,   // |變化%| ≤ 10% → 正常
  GRADE_CAUTION_PCT: 0.25,  // |變化%| ≤ 25% → 注意；超過 → 異常
  GRADE_LABELS: {
    NORMAL: '正常',
    CAUTION: '注意',
    ERROR: '異常',
    NODATA: '資料不足',   // 這一欄對應的幾週全部都沒有歷史值可比
    MISSING: '本日缺值',  // 今天這個品項本身算不出消耗
    INCREASED: '變多了!', // 備品當日消耗算出負值（通常是補貨），直接標記，不套用一般評級
  },

  /* ── 今日備量（只有 JZ、HB 有；隔天備量 − 目前總計，四捨五入取整、負值顯示0）── */
  // 星期分組：日盤表右側「最低備量」的三個欄位分別對應這三組
  //   一二 = 週一、週二     三四日 = 週三、週四、週日     五六 = 週五、週六
  // 用「隔天」的星期幾決定要用哪一組。0=週日,1=週一,...6=週六（對應 JS Date.getDay()）
  STOCK_PREP_GROUP_BY_DOW: { 0: 1, 1: 0, 2: 0, 3: 1, 4: 1, 5: 2, 6: 2 },
  STOCK_PREP_GROUP_LABELS: ['一二', '三四日', '五六'], // 僅供註解對照，程式不依賴文字比對

  /* ── 觸發完成的方式 ─────────────────────────────────────── */
  // 'CELL_INPUT' ：偵測指定儲存格（COMPLETION_CELL）被輸入內容就跑（目前使用中）
  // 'CHECKBOX'   ：每日分頁放一個「盤點完成」勾選框，打勾就跑
  // 'ALL_FILLED' ：偵測三大區所有品項的「當日消耗」都算得出數字就跑
  COMPLETION_MODE: 'CELL_INPUT',
  COMPLETION_LABEL_CELL: 'K1', // 僅 CHECKBOX 模式會用到（勾選框旁邊的說明文字位置）
  COMPLETION_CELL: 'B5',       // CELL_INPUT 模式：要偵測輸入的儲存格；CHECKBOX 模式：勾選框位置

  /* ── 輸出分頁命名 ───────────────────────────────────────── */
  // 用來源當天的日期產生，例：'MMDD' → 0709。與你日盤表、範本檔一致。
  OUTPUT_TAB_PATTERN: 'MMDD',

  /* ── 區塊定義（順序＝報告裡 zone 出現的順序）────────────────── */
  // key        ：zone 代碼，報告裡會顯示成「JZ zone」這種列
  // title      ：日盤表裡的區塊標題（在 B 欄，用來定位整個區塊）
  // valueHeader：這個區塊「要拿來比較的那一欄」的表頭文字
  // mode       ：DIRECT = 直接讀該欄；CROSS_DAY_DIFF = 用(昨日-今日)算消耗
  // totalHeader / stockPrepMildTitle / stockPrepPeakTitle：
  //   只有 JZ、HB 有「最低備量」欄位，才需要填這三個；YG、BP 不填。
  SECTIONS: [
    {
      key: 'JZ', title: '假裝清醒', valueHeader: '當日消耗', mode: 'DIRECT',
      totalHeader: '總計', stockPrepMildTitle: '最低備量(淡季)', stockPrepPeakTitle: '最低備量(旺季)',
    },
    { key: 'YG', title: '酉鬼',     valueHeader: '當日消耗', mode: 'DIRECT' },
    {
      key: 'HB', title: '烘焙雲',   valueHeader: '當日消耗', mode: 'DIRECT',
      totalHeader: '總計', stockPrepMildTitle: '最低備量(淡季)', stockPrepPeakTitle: '最低備量(旺季)',
    },
    { key: 'BP', title: '備品',     valueHeader: '總計',     mode: 'CROSS_DAY_DIFF' },
  ],

  // 只在 B 欄（第 2 欄）尋找區塊標題，避免誤抓到表頭裡同名的欄位（如「酉鬼」也是備品表頭的庫位名）
  SECTION_TITLE_COL: 2,
  // 品名表頭（用 startsWith 比對，因為備品是「品名（Key個數）」）
  NAME_HEADER_PREFIX: '品名',

  /* ── 只有分頁名稱像日期的才被視為「有效日盤分頁」───────────── */
  // 0709 / 72 / 625 / 611 … 會被當成盤點分頁；範本、工作表83 這類自動略過。
  DATE_TAB_REGEX: /^\d{2,4}(（\d+）)?$/,

  /* ── 報告樣式（照你提供的「當日消耗報告.xlsx」樣板抄）───────── */
  FONT: 'Arial',
  NAME_COL_WIDTH: 190,   // A 欄（品名）欄寬
  DATA_COL_WIDTH: 100,   // B/C/D 數值欄欄寬
  GRADE_COL_WIDTH: 100,  // E 評級欄欄寬

  COL_B_FILL: 'FFF2CC',  // 當日消耗欄，整欄淡黃底（樣板既有設計）

  // 各 zone 標題列底色（顏色取自樣板原檔）
  ZONE_COLORS: {
    JZ: 'FF9900',
    YG: 'FFE599',
    HB: '93C47D',
    BP: '9FC5E8',
  },

  // 評級欄底色（樣板沒有，這是我加的辨識用色，可以關掉，見 Output.gs 的 APPLY_GRADE_COLOR）
  GRADE_COLORS: {
    '正常': 'D9EAD3',
    '注意': 'FCE5CD',
    '異常': 'F4CCCC',
    '資料不足': 'EFEFEF',
    '本日缺值': 'EFEFEF',
    '變多了!': 'C9DAF8',
  },
  APPLY_GRADE_COLOR: true,
};